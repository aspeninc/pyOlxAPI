"""
Purpose: Generate list of backups relay groups on transmission lines
        using following criteria:
        - Backup must be located on transmission line behind the primary
        - Close-in with end open fault at the primary group must be
                seen as forward fault (MTA = 30) at backup location
         Output RAT file with coordination pairs which can be imported to OLR
         file using Relay | Import command
"""
from __future__ import print_function
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "0.1.1"
__email__     = "support@aspeninc.com"
__status__    = "In development"

# IMPORT -----------------------------------------------------------------------
import sys,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
sys.path.insert(0, PATH_LIB)
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
# IMPORT -----------------------------------------------------------------------

#add more input cmd here
PARSER_INPUTS.usage =\
"\n\tGenerate list of backups relay groups on transmission lines\
\n\tusing following criteria:\
\n\t - Backup must be located on transmission line behind the primary\
\n\t - Close-in with end open fault at the primary group must be\
\n\t        seen as forward fault (MTA = 30) at backup location\
\n\t Output RAT file with coordination pairs which can be imported to OLR\
\n\t file using Relay | Import command"

PARSER_INPUTS.add_argument('-kv',metavar='', help = 'study KV level',default = 0  , type=float)
#
PARSER_INPUTS.add_argument('-na1',metavar='', help = 'bus1 Name'    ,default = "" , type=str)
PARSER_INPUTS.add_argument('-kv1',metavar='', help = 'bus1 KV'      ,default = 0  , type=float)
PARSER_INPUTS.add_argument('-na2',metavar='', help = 'bus2 Name'    ,default = "" , type=str)
PARSER_INPUTS.add_argument('-kv2',metavar='', help = 'bus2 KV'      ,default = 0  , type=float)
#
PARSER_INPUTS.add_argument('-nu1',metavar='', help = 'bus1 Number'  ,default = -1  , type=int)
PARSER_INPUTS.add_argument('-nu2',metavar='', help = 'bus2 Number'  ,default = -1  , type=int)
PARSER_INPUTS.add_argument('-cid',metavar='', help = 'circuit ID'   ,default = ""  , type=str)
#
PARSER_INPUTS.add_argument('-nbr',metavar='', help = 'name Branch'  ,default = ""  , type=str)
#
PARSER_INPUTS.add_argument('-gui',metavar='', help = '0/1 GUI option to select branch (default=0)',default = 0, type=int)
PARSER_INPUTS.add_argument('-fo' ,metavar='', help = 'txt report file',default = "", type=str)
PARSER_INPUTS.add_argument('-fr' ,metavar='', help = 'RAT file (imported to OLR file using Relay|Import command)',default = "", type=str)
ARGVS = PARSER_INPUTS.parse_args()

#
def getpriback_kV(kV):
    relG = OlxAPILib.getEquipmentHandle(TC_RLYGROUP)
    bra  = OlxAPILib.getEquipmentData(relG,RG_nBranchHnd)
    bus1 = OlxAPILib.getEquipmentData(bra,BR_nBus1Hnd)
    inSer = OlxAPILib.getEquipmentData(bra,BR_nInService)
    kVa  = OlxAPILib.getEquipmentData(bus1,BUS_dKVnominal)
    typa = OlxAPILib.getEquipmentData(bra,BR_nType)
    #
    brs,bf_a,be_a,pe_a,fullB = [],[],[],[],[]
    for i in range(len(relG)):
        if abs(kVa[i]-kV)<=0.1 and typa[i]==TC_LINE and inSer[i]==1:
            bf1,be1,pe1,fullB1 = getpriback1(bra[i],relG[i])
            bf_a.append(bf1)
            be_a.append(be1)
            pe_a.append(pe1)
            brs.append(relG[i])
            fullB.append(fullB1)
    #
    return createReport_RAT(brs,bf_a,be_a,pe_a,fullB)
#
def getpriback1(br1,rg1):
    # get pairs exist
    back_exist = OlxAPILib.get1EquipmentData_array(rg1,RG_nBackupHnd)
    pri_exist  = OlxAPILib.get1EquipmentData_array(rg1,RG_nPrimaryHnd)
    back_found = []
    #
    # bus 1
    bus1 = OlxAPILib.getEquipmentData([br1],BR_nBus1Hnd)[0]
    bra =  OlxAPILib.getBusEquipmentData([bus1],TC_BRANCH)[0]
    typa = OlxAPILib.getEquipmentData(bra,BR_nType)
    #
    rgOppo = [] # relay group opposite
    fullBackup = True
    for i in range(len(bra)):
        if bra[i]!=br1 and typa[i]==TC_LINE:
            bres = OlxAPILib.getOppositeBranch(bra[i],[TC_LINE])
            for bri in bres:
                rgi = OlxAPILib.get1EquipmentData_try(bri,BR_nRlyGrp1Hnd)
                if rgi>0:
                    rgOppo.append(rgi)
                else:
                    fullBackup = False
            #
            if len(bres)==0:
                fullBackup = False
    #
    rgOppo = list(set(rgOppo))
    # run fault
    fltConn   = [1,0,0,0] # 3LG
    fltOpt    = [0]*15
    fltOpt[2] = 1 # Bus or Close-in with end opened
    OlxAPILib.doFault(rg1,fltConn=fltConn,fltOpt=fltOpt,outageOpt=[],outageLst=[],fltR=0,fltX=0,clearPrev=1)
    OlxAPILib.pick1stFault()
    # MTA method
    MTA = 30
    for rgi in rgOppo:
        # Check if fault is in forward direction
        mi1,ai1 = OlxAPILib.getSCCurrent_p(rgi)
        mv1,av1 = OlxAPILib.getSCVoltage_p(rgi)
        if mv1>=1 and mi1>=10:
            difa = av1-ai1
            #
            if difa >= MTA-90 and  difa <= MTA+90:
                back_found.append(rgi)
    #
    return back_found,back_exist,pri_exist,fullBackup

def createReport_RAT(brs,bf_a,be_a,pe_a,fullB):
    sr = ""
    srat = "[ASPEN RELAY DATA]"
    srat += "\ndelimiter='"
    srat += "\napp= 'ASPEN OneLiner and Power Flow'"
    srat += "\nver= 2014 'A' 14 6"
    srat += "\n"
    srat += "\nCOORDINATION PAIRS"
    #
    nFull,nPartial,nNoback = 0, 0, 0
    nNewp, nExistP, nMissed = 0, 0 ,0
    #
    for i in range(len(brs)):
        sr+= "\nRelaygroup: " + OlxAPILib.fullBranchName(brs[i])
        if len(bf_a[i])<=0:
            sr+= "\nNo backup (##)"
            nNoback +=1
        else:
            if fullB[i]:
                sr+= "\nBackups: "
                nFull += 1
            else:
                sr+= "\nBackups (#):"
                nPartial += 1
            #
            for bf1 in bf_a[i]:
                sr+= "\n\t" + OlxAPILib.fullBranchName(bf1)
                 #
                if bf1 in be_a[i]:
                    sr+= "(!)"
                else: # New pairs found by script
                    nNewp += 1
                    sr+= "(*)"
                    #
                    srat += "\n" + getStrBrForPair(brs[i])
                    srat += "\n" + getStrBrForPair(bf1)
            # missed by script
            for be1 in be_a[i]:
                if be1 not in bf_a[i]:
                    sr+= "\n\t" + OlxAPILib.fullBranchName(be1) +"(!!)"
                    nMissed += 1
            #
            nExistP += len(be_a[i])
        #
        sr+= "\n"
    #
    sr+= "\nTotal groups processed = " + str(len(brs))
    sr+= "\nGroups with full backup = " + str(nFull)
    sr+= "\n(#)Groups with patial backup = " + str(nPartial)
    sr+= "\n(##)Groups with no backup = " + str(nNoback)
    sr+= "\n(*)New pairs found by script = " + str(nNewp)
    sr+= "\n(!)Existing pairs = " + str(nExistP)
    sr+= "\n(!!)Existing pairs missed by script = " + str(nMissed)
    #
    OlxAPILib.saveString2File(ARGVS.fo,sr)
    OlxAPILib.saveString2File(ARGVS.fr,srat)
    #
    return sr,srat

def getStrBrForPair(rg1):
    br1 = OlxAPILib.getEquipmentData([rg1],RG_nBranchHnd)[0]
    bus = OlxAPILib.getBusByBranch(br1)
    sres = ""
    for i in range(2):
        b1 = bus[i]
        bnu = OlxAPILib.getEquipmentData([b1],BUS_nNumber)[0]
        bna = OlxAPILib.getEquipmentData([b1],BUS_sName)[0]
        bkv = OlxAPILib.getEquipmentData([b1],BUS_dKVnominal)[0]
        sres += " " +str(bnu) + "; '" + bna + "'; " +str(round(bkv,3)) + "; "
    #
    sID = OlxAPILib.getDataByBranch(br1,"sID")
    sres += "'" + sID + "';"
    #
    typ = OlxAPILib.getEquipmentData([br1],BR_nType)[0]
    sres += str(OlxAPILib.dictCode_BR_RAT[typ])
    #
    return sres
#
def unit_test():
    sres = "\nUNIT TEST: " + PY_FILE
    ARGVS.fi = "MAKEPRIBACK"
    ARGVS.kv = 132
    ARGVS.fo, ARGVS.fr = "ut.txt","ut.RAT"
    sres += run()
    #
    OlxAPILib.unit_test_compare(PATH_FILE,PY_FILE,sres)
    #
    OlxAPILib.deleteFile(ARGVS.fo)
    OlxAPILib.deleteFile(ARGVS.fr)

#
def quality_test():
    print("\nQUALITY TEST: " + PY_FILE)
    ARGVS.fi = "MAKEPRIBACK"
    ARGVS.na1, ARGVS.na2 = "BUS1", "BUS7"
    ARGVS.fo, ARGVS.fr = "qt.txt","qt.RAT"
    run()
#
def run():
    OlxAPILib.open_olrFile(ARGVS.fi,readonly=1) # default "SAMPLE30.OLR"
    #
    ARGVS.fo  = OlxAPILib.updateNameFile(PATH_FILE,ARGVS.fi,".OLR", ARGVS.fo,".txt")
    ARGVS.fr  = OlxAPILib.updateNameFile(PATH_FILE,ARGVS.fi,".OLR", ARGVS.fr,".RAT")
    #
    sres  = "\nRun : " + PY_FILE
    sres += "\nOLR file : " + os.path.basename(ARGVS.fi)
    srat = ""
    if ARGVS.kv>0: # study KV level
        sres += "\nStudy KV level: "+str(ARGVS.kv)
        sr,srat = getpriback_kV(ARGVS.kv)
        sres += sr
    else: # branch
        bs = OlxAPILib.BranchSearch(gui=ARGVS.gui)
        #
        bs.setBusNum1(ARGVS.nu1)
        bs.setBusNameKV1(ARGVS.na1 , ARGVS.kv1)
        bs.setBusNum2(ARGVS.nu2)
        bs.setBusNameKV2(ARGVS.na2 , ARGVS.kv2)
        bs.setCktID(ARGVS.cid)
        bs.setNameBranch(ARGVS.nbr)
        #
        br1 = bs.runSearch()
        if br1>0:
            rg1 = OlxAPILib.get1EquipmentData_try(br1,BR_nRlyGrp1Hnd)
            if rg1<=0:
                sTitle = "NOT found Relay Group"
                sMain = "on branch: " + OlxAPILib.fullBranchName(br1)
                OlxAPILib.gui_Error(sTitle,sMain)
                #
                sres += "\n"+ sTitle + " " + sMain
            else:
                sres += "\nSearch backup of Relay Group on branch: " + OlxAPILib.fullBranchName(br1)
                bf1,be1,pe1,fullB1 = getpriback1(br1,rg1)
                sr,srat = createReport_RAT([rg1],[bf1],[be1],[pe1],[fullB1])
                sres += sr
        else:
            if ARGVS.gui==0:
                bs.setGui(1)
                br1 = bs.runSearch()
            sres += "\nBranch search: No branch found!"
    #
    print(sres)
    #
    return sres

if __name__ == '__main__':
    if (ARGVS.ut ==1):
        unit_test()
    elif (ARGVS.ut ==2):
        quality_test()
    else:
        run()







