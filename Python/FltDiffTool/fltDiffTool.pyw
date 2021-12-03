"""
Purpose:
     Read input file, simulate faults and generate report with relay quantities.
     The generated report can also be use as input to generate a deviation report
     that shows the differences in the relay quantities between runs.

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

# IMPORT -----------------------------------------------------------------------
import sys,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
from AppUtils import *
#
chekPythonVersion(PY_FILE)
#
import xml.etree.ElementTree as ET
import tkinter as tk
import tkinter.filedialog as tkf
import tkinter.messagebox as tkm
from tkinter import ttk

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = iniInput(usage=
        "\n\tRead input file, simulate faults and generate report with relay quantities.\
        \n\tThe generated report can also be use as input to generate a deviation report\
        \n\tthat shows the differences in the relay quantities between runs.")
#
PARSER_INPUTS.add_argument('-ft' , help = ' (str) Input template file (CSV/Excel format)',default = "", metavar='')
PARSER_INPUTS.add_argument('-bo' , help = ' (int) Branch Outages review option ([yes,no(default)])',default = '', metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) File (CSV/Excel format) to save calculation results ',default = '', metavar='')
PARSER_INPUTS.add_argument('-demo',help = ' (int) demo [0-ignore, 1-with Branch Outages review,2-without Branch Outages review,\
                                                                  3-comparison with previous run]', default = 0,type=int,metavar='')
ARGVS = parseInput(PARSER_INPUTS)

#
def getFdes(s1):
    sa = str(s1).splitlines()
    fdes = sa[0]
    idx = str(fdes).index(".")
    if idx>0:
        fdes = fdes[idx+1:]
    #
    for i in range(1,len(sa)-2):
        fdes += "\n"+ deleteSpace1(sa[i])
    return fdes
#
def run1SIMULATEFAULT(opt,fltstr):
    #
    ao = str(opt).split()
    #
    data = ET.Element('SIMULATEFAULT')
    data1 = ET.SubElement(data, 'FAULT')
    data1.set('FLTDESC',"")
    sfltop = ['PREFV','VPF','GENZTYPE','IGNORE','GENILIMIT','VCCS','MOV','MOVITERF','MINZMUPAIR','MVASTYLE']
    for i in range(len(sfltop)):
        data1.set(sfltop[i],ao[i])
    #
    data1.set('FLTSPCVS',fltstr)
    sInput = ET.tostring(data)
    if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(sInput):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

def compare(va1,va0,valC):
    a1 = []
    k = 0
    for j in range(max(len(va1),len(va0))):
        try:# float
            v1 = float(va1[j])
            v0 = float(va0[j])
            av1 = abs(v1)
            av0 = abs(v0)
            try:
                c = abs(v1-v0)/min(av1,av0)*100
            except:
                try:
                    c = abs(v1-v0)/max(av1,av0)*100
                except:
                    c = 0
            #
            if c>=valC:
                k+=1
            #
            a1.extend([v1,v0,round(c,1)])
        except:
            # String
            try:
                v1 = str(va1[j])
                v0 = str(va0[j])
                if v1==v0:
                    a1.extend([v1,v0,"0"])
                else:
                    a1.extend([v1,v0,"!"])
                    k+=1
            except:
                #
                try:
                    a1.append(va1[j])
                except:
                    a1.append("")
                #
                try:
                    a1.append(va0[j])
                except:
                    a1.append("")
                #
                a1.append("")
    return a1,k

#
class FltDiffTool:
    def __init__(self,ft):
        #
        self.t0 = time.time()
        self.ft = os.path.abspath(ft)
        #
        self.ck = checkFileType(fi=self.ft,ext=[".CSV",".XLSX",".XLS",".XLSM"],err=True,sTitle0="Input template file")
        #
        self.ws = ToolCSVExcell()
        self.ws.readFile(fi=self.ft,delim=',')
        #
        self.bo = ARGVS.bo
        #
        self.ws.selectSheet(nameSheet = "config")
        #
        self.bs = OlxAPILib.BusSearch(gui=0)
        self.brs = OlxAPILib.BranchSearch(gui=0)
        #
        self.relayQR_ak_0  = []
        self.relayQR_key_0 = []
        self.relayQR_VAL_0 = {}
        #
        self.relayQR_ak_1  = []
        self.relayQR_key_1 = []
        self.relayQR_VAL_1 = {}
        #
        self.sBranchOutage = []
        self.sFaultSpec = []
        self.sOLR = []
        self.sOLR0 = ["",""]
        self.sDate0 = ["",""]
        #
        self.outageLst = []
        self.sBranchOutage = [["BRANCH OUTAGES"]]
        self.sFaultSpec = [["FAULTS DESCRIPTION","DESC","[PREFV VPF GENZTYPE IGNORE GENILIMIT VCCS MOV MOVITERF MINZMUPAIR MVASTYLE]","FLTSPCVS"]]
        self.rowFaultSpec = []
    #
    def run2BranchOutage(self):
        needRun = False
        if len(self.outageLst)==0 and len(self.outageBus)>0: #not exit yet branch outage
            needRun = True
        #
        if needRun:
            for i in range(len(self.outageBus)):
                bra = OlxAPILib.getOutageList(self.outageBus[i], self.outageTier[i])
                self.outageLst.extend(bra)

        #removeDuplicateBranchOutage
        ehnda = []
        if self.br0>0:
            e0 = OlxAPILib.getEquipmentData([self.br0],BR_nHandle)[0]
            ehnda.append(e0)
        #
        res = []
        ea =  OlxAPILib.getEquipmentData(self.outageLst,BR_nHandle)
        for i in range(len(self.outageLst)):
            if ea[i] not in ehnda:
                res.append(self.outageLst[i])
                ehnda.append(ea[i])
        #
        self.outageLst = []
        self.outageLst.extend(res)
        #
        if needRun:
            for i in range(len(self.outageLst)):
                br1 = self.outageLst[i]
                s1 = [i+1]
                busa = OlxAPILib.getBusByBranch(br1)
                #
                sb = OlxAPILib.getEquipmentData(busa,BUS_sName)
                kv = OlxAPILib.getEquipmentData(busa,BUS_dKVnominal)
                #
                id1 = OlxAPILib.getDataByBranch(br1,"sID")
                bc1 = OlxAPILib.getDataByBranch(br1,"sC")
                #
                for j in range(2):
                    s1.append(sb[j])
                    s1.append(round(kv[j],5))
                s1.append(id1)
                s1.append(bc1)
                #
                self.sBranchOutage.append(s1)
        #
        return
    #
    def getFltOpt(self):
        self.fltOpt = []
        vo1 = 0 if len(self.outageLst)==0 else 1
        if self.br0<=0: # "Study Bus"
            fltOpt1 = [0]*15
            fltOpt1[0],fltOpt1[1] = 1,vo1
            self.fltOpt.append(fltOpt1)
            return
        #
        for i in range(len(self.faultLocation)):
            v1 = self.faultLocation[i]
            flag = self.faultLocationFlag[i]
            vo = 0 if len(self.outageLst)==0 else v1
            fltOpt1 = [0]*15
            if flag==0 : # String
                if v1=="CLOSEIN":
                    fltOpt1[0],fltOpt1[1] = 1,vo1
                elif v1=="CLOSEIN_EOP":
                    fltOpt1[2],fltOpt1[3] = 1,vo1
                elif v1=="FARBUS":
                    fltOpt1[4],fltOpt1[5] = 1,vo1
                elif v1=="LINEEND":
                    fltOpt1[6],fltOpt1[7] = 1,vo1
                else:
                    raise Exception("error faultLocation")
                #
            elif flag==1: # 8,9
                fltOpt1[8],fltOpt1[9] = v1,vo
            #
            elif flag==2: #10,11
                fltOpt1[10],fltOpt1[11] = v1,vo
            else:
                raise Exception("error faultLocationFlag")
            #
            self.fltOpt.append(fltOpt1)
    #
    def runSIMULATEFAULT(self):
        #
        for i in range(1,len(self.sFaultSpec)):
            s1 = self.sFaultSpec[i]
            k = int(s1[0])
            fdes0 = s1[1]
            #
            try:
                run1SIMULATEFAULT(opt=s1[2],fltstr=s1[3])
                fdes = self.getResult1(k)
            except:
                self.errData("Error Fault Description",row=self.rowFaultSpec[i],column=1,nc=4,sheet='config')
    #
    def runDoFault(self):
        #
        bhnd = self.bus0
        if self.br0>0:
            bhnd = self.br0
        #
        k = 0
        for rx in self.rx:
            for fltConn in self.fltConn:
                for fltOpt in self.fltOpt:
                    OlxAPILib.doFault(bhnd,fltConn,fltOpt,self.outageOpt,self.outageLst, rx[0], rx[1], clearPrev=1)
                    k = self.getResult(k)
    #
    def getResult1(self,k): # get result for simulation fault
        try:
            OlxAPILib.pick1stFault()
        except:
            return ""
        #
        s1 = OlxAPI.FaultDescriptionEx(0,2)
        fdes = getFdes(s1)
        self.get1UI_RQR(k)
        return fdes
    #
    def getResult(self,k):
        try:
            OlxAPILib.pick1stFault()
        except:
            return k
        #
        k1 = k+1
        while True:
            s1 = OlxAPI.FaultDescriptionEx(0,2)
            fdes = getFdes(s1)
            sa = str(s1).splitlines()
            fspec = sa[len(sa)-2]
            fopt = sa[len(sa)-1]
            #
            op1 = [k1,fdes,fopt,fspec]
            self.sFaultSpec.append(op1)
            #
            self.get1UI_RQR(k1)
            #
            if not OlxAPILib.pickNextFault():
                break
            k1+=1
        return k1
    #
    def get1UI_RQR(self,k1):
        nRound = 1
        for i in range(len(self.relayGroup)):
            rg1 = self.relayGroup[i]
            ra1 = self.relay[i]
            br1 = rg1[0]
            key1 = str(k1)+"_"+rg1[2] + "_"
            a1 = [k1,rg1[2],""]
            #
            resBr1 = ["",""]
            magi,angi = OlxAPILib.getSCCurrent(br1,style=4)
            for i in range(3):
                resBr1.append(round(magi[i],nRound))
                resBr1.append(round(angi[i],nRound))
            magv,angv = OlxAPILib.getSCVoltage(br1,style=4)
            for i in range(3):
                resBr1.append(round(magv[i],nRound))
                resBr1.append(round(angv[i],nRound))
            #
            if len(ra1)==0:
                self.relayQR_ak_1.append(a1)
                self.relayQR_key_1.append(key1)
                self.relayQR_VAL_1[key1] = resBr1
            #
            for r1 in ra1:
                v1 = resBr1[:]
                #
                key2 = key1 + r1[1]
                a1 = [k1,rg1[2],r1[1]]
                self.relayQR_ak_1.append(a1)
                self.relayQR_key_1.append(key2)
                #
                t1,roc = OlxAPILib.getRelayTime(hndRelay= r1[0], mult=1.0, consider_signalonly=1)
                v1[0] = roc # TO ADD HERE
                v1[1]= round(t1,3)
                #
                self.relayQR_VAL_1[key2] = v1
    #
    def getBusHnd(self,name,kv):
        self.bs.setBusNameKV(str(name),float(kv))
        return self.bs.runSearch()
    #
    def errData(self,s1,row,column,nc,sheet):
        sTitle = "ERROR Input Template File "
        sMain = s1
        sMain += "\n\nFile: " + os.path.basename(self.ft)
        if self.ws.isExcel:
            self.ws.selectSheet(sheet)
            sMain += "\nSheet : "+sheet

        sMain += "\nrow = "+str(row)
        if nc==1:
            sMain += ", column = "+str(column)
        else:
            sMain += ", column = ["+str(column)
            for i in range(nc-1):
                sMain += "," + str(column+i+1)
            sMain += "]"
        #
        vala = self.ws.getValRowf(row=row,columnFr=column,leng=nc)
        sMain += "\n" + array2String(vala,",")
        #
        raise Exception("\n"+sTitle+"\n"+sMain)
        #
    def warningData(self,s1,row,column,nc):
        sTitle = "Warning input data"
        sMain = s1
        vala = self.ws.getValRowf(row=row,columnFr=column,leng=nc)
        sMain += "\n" + array2String(vala,",")
        sMain+="\nFile : " + os.path.basename(self.ft)
        if self.ws.isExcel:
            sMain += "\nSheet : "+self.ws.currentSheet

        sMain += "\nrow = "+str(row)
        if nc==1:
            sMain += ", column = "+str(column)
        else:
            sMain += ", column = ["+str(column)
            for i in range(nc-1):
                sMain += "," + str(column+i+1)
            sMain += "]"
        gui_info(sTitle+"(OK continue)",sMain)
        return "\n"+sTitle+": "+sMain
    #
    def warningLog(self,s1,row,column,nc):
        sTitle = "Warning "
        sMain = s1
        vala = self.ws.getValRowf(row=row,columnFr=column,leng=nc)
        sMain += array2String(vala,",")
        #
        sMain += "\n\trow = "+str(row)
        if nc==1:
            sMain += ", column = "+str(column)
        else:
            sMain += ", column = ["+str(column)
            for i in range(nc-1):
                sMain += "," + str(column+i+1)
            sMain += "]"
        #
        if self.ws.currentSheet != None:
            sMain += ", Sheet : "+self.ws.currentSheet
        #
        print(sTitle+": "+sMain)
    #
    def errSheet(self,sheetName):
        sTitle = "Sheet not found"
        sMain  = "\nFile : " + os.path.basename(self.ft)
        sMain += "\nSheet : " + sheetName + " not found"
        gui_rrror(sTitle,sMain)
        raise Exception("\n"+sTitle+"\n"+sMain)

    #
    def getBranchHnd(self,arg):
        self.brs.setBusNameKV1(str(arg[0]),float(arg[1]))
        self.brs.setBusNameKV2(str(arg[2]),float(arg[3]))
        self.brs.setCktID(str(arg[4]))
        return self.brs.runSearch()

    # OLR  ---------------------------------------------------------------------
    def getOLR(self,kr):
        vala = self.ws.getValRowf(row=kr,columnFr=1,leng=2)
        if vala[0] !="OLR FILE":
            self.errData("Value must be 'OLR FILE'",kr,1,1,sheet='config')
        #
        self.OLR = str(vala[1])
        self.OLR = os.path.abspath(self.OLR)
        #
        print("Input template file: " + self.ft)
        try:
            OlxAPILib.open_olrFile(self.OLR,ARGVS.olxpath)
        except Exception as e:
            s1 = str(e)
            if s1.startswith('Error: Another OlxAPI'):
                raise Exception('\n'+s1)
            else:
                self.errData(s1,kr,1,2,sheet='config')
        #
        self.sOLR = ["OLR FILE",self.OLR]
        #
        return kr+1

    # Study---------------------------------------------------------------------
    def getStudy(self,kr):
        vala = self.ws.getValRow(row=kr,columnFr=1)
        vala[0]= deleteSpace1(vala[0])
        self.sstudy = vala
        self.bus0, self.br0 = 0, 0
        #
        if vala[0] not in ["Study Branch","Study Bus"]:
            self.errData("Value must be 'Study Branch' or 'Study Bus'",kr,1,1,sheet='config')
        #
        try:
            self.bus0 = self.getBusHnd(vala[1],vala[2])
        except:
            pass
        #
        if self.bus0<=0:
            self.errData("Bus Not Found",row=kr,column=2,nc=2,sheet='config')
        #
        if vala[0]=="Study Branch":
            try:
                self.br0 = self.getBranchHnd(vala[1:6])
            except:
                pass
            #
            if self.br0<=0:
                self.errData("Branch Not Found",row=kr,column=2,nc=5,sheet='config')
        #
        return kr+1

    # Fault Location------------------------------------------------------------
    def getFaultLocation(self,kr):
        vala = self.ws.getValRow(row=kr,columnFr=1)
        self.sFaultLocation = vala
        #
        if vala[0] != "Fault Location":
            self.errData("Value must be 'Fault Location'",row=kr,column=1,nc=1,sheet='config')
        #
        if len(vala)==1:
            self.errData("Error None 'Fault Location'",row=kr,column=2,nc=1,sheet='config')
        #
        self.faultLocation = []
        self.faultLocationFlag = [] # 0: string, 1: INT, 2 INT_EOP
        #
        for k in range(1,len(vala)):
            v1 = vala[k]
            if v1 in ["CLOSEIN","CLOSEIN_EOP","FARBUS","LINEEND"]:
                (self.faultLocation).append(v1)
                (self.faultLocationFlag).append(0)
            elif v1.startswith("INT_"):
                sa = v1.split("_")
                km = len(sa)
                if sa[len(sa)-1] =="EOP":
                    km = len(sa)-1
                for i in range(1, km):
                    vf1 = -1
                    try:
                        vf1 = float(sa[i])
                    except:
                        pass
                    if vf1<0:
                        self.errData("Error value",kr,k+2,1,sheet='config')
                    (self.faultLocation).append(vf1)
                    if km == len(sa): # no EOP
                        (self.faultLocationFlag).append(1)
                    else: # EOP
                        (self.faultLocationFlag).append(2)
        #
        return kr+1

    #Row=4 Fault Type ----------------------------------------------------------
    def getFaultType(self,kr):
        vala = self.ws.getValRow(row=kr,columnFr=1)
        self.sFaultType = vala
        self.fltConn = []
        #
        if vala[0] != "Fault Type":
            self.errData("Value must be 'Fault Type'",kr,1,1,sheet='config')
        #
        for k in range(1,len(vala)):
            v1 = vala[k]
            try:
                fltConn1 = OlxAPILib.dictFltConn[v1]
                self.fltConn.append(fltConn1)
            except:
                self.errData("Value must be [3PH, 2LG_BC, 2LG_CA, 2LG_AB, 1LG_A, 1LG_B, 1LG_C, LL_BC, LL_CA, LL_AB]",kr,k+1,1,sheet='config')
        #
        return kr+1

    #Row=5 Outage---------------------------------------------------------------
    def getOutage(self,kr):
        vala = self.ws.getValRow(row=kr,columnFr=1)
        self.sOutage = vala
        self.outageBus = []
        self.outageTier = []
        #
        if vala[0] != "Outage (bus;tier)": # check
            self.errData("Value must be 'Outage (bus;tier)'",kr,1,1,sheet='config')
        #
        if (len(vala)-1)%3 !=0:
            self.errData("Error format Outage (bus;tier)",kr,2,len(vala)-1,sheet='config')
        #
        k = 1
        while k+2<len(vala):
            try:
                b1 = self.getBusHnd(vala[k],vala[k+1])
                t1 = float(vala[k+2])
            except:
                b1, t1 = 0, 0.1 #error
            #
            if b1<=0 or abs(int(t1)-t1)>1e-5:
                self.errData("Bus Not Found or Error tier",row=kr,column=k,nc=3,sheet='config')
            #
            k+=3
            self.outageBus.append(b1)
            self.outageTier.append(int(t1))
        #
        return kr+1

    #Row=6 Outage Type----------------------------------------------------------
    def getOutageType(self,kr):
        vala = self.ws.getValRow(row=kr,columnFr=1)
        self.sOutageType = vala
        #
        if vala[0] != "Outage Type": # check
            self.errData("Value must be 'Outage Type'",kr,1,1,sheet='config')
        #
        self.outageOpt = [0]*4
        for k in range(1,len(vala)):
            v1 = vala[k]
            if v1=="Single":
                self.outageOpt[0]=1
            elif v1=="Double":
                self.outageOpt[1]=1
            else:
                self.errData("Value must be [Single,Double]",kr,k+1,1,sheet='config')
        #
        return kr+1

    #Row=7 FaultZ----------------------------------------------------------
    def getFaultZ(self,kr):
        vala = self.ws.getValRowf(row=kr,columnFr=1,leng=6)
        self.sFaultZ = vala
        #
        if vala[0] != "Fault Z(R;X from/to/steps Ohm)": # check
            self.errData("Value must be 'Fault Z(R;X from/to/steps Ohm)'",kr,1,1,sheet='config')
        #
        z = []
        for k in range(1,len(vala)):
            try:
                t1 = float(vala[k])
            except:
                t1 = -1
            #
            if t1<0 or (k==6 and abs(t1-int(t1))>1e-5):
                self.errData("Error value",kr,k+1,1,sheet='config')
            #
            z.append(t1)
        # rx -------------------------------------------------------------------
        self.rx= [[z[0],z[1]]]
        if ((z[2]-z[0])>1e-3 or (z[3]-z[1])>1e-3) and z[4]>1:
            for i in range(int(z[4]-1)):
                ri = z[0] + (i+1.0) * (z[2]-z[0])/(z[4]-1.0)
                xi = z[1] + (i+1.0) * (z[3]-z[1])/(z[4]-1.0)
                self.rx.append([ri,xi])
        #
        return kr+1

    #Row=8 Relay Group----------------------------------------------------------
    def getRelay(self,kr):
        self.relay = []
        self.relayGroup = []
        self.sRelay = []
        kr1 = kr
        while True:
            vala = self.ws.getValRow(row=kr1,columnFr=1)
            if vala[0]=="":
                break
            if str(vala[0]).startswith("#"):
                kr1+=1
                self.sRelay.append(vala)
            else:
                #
                if vala[0] != "Relay Group":
                    self.errData("Value must be 'Relay Group'",row=kr1,column=1,nc=1,sheet='config')
                #
                self.sRelay.append(vala)
                try:
                    br1 = self.getBranchHnd(vala[1:6])
                    rg1 = OlxAPILib.get1EquipmentData_try(br1,BR_nRlyGrp1Hnd)
                    na1 = OlxAPILib.fullBranchName(br1)
                    if rg1>0:
                        self.relayGroup.append([br1,rg1,na1])
                except:
                    rg1 = 0
                #
                if rg1<=0:
                    self.warningLog("Relay Group Not Found: ",row=kr1,column=2,nc=5)
                    while True:
                        kr1+=1
                        vala = self.ws.getValRow(row=kr1,columnFr=1)
                        if vala[0]==""or vala[0] == "Relay Group":
                            break
                        self.sRelay.append(vala)
                else:
                    kr1 +=1
                    vala = self.ws.getValRow(row=kr1,columnFr=1)
                    if vala[0]=="":
                        break
                    if vala[0] == "Relay Group" or str(vala[0]).startswith("#"):
                        self.relay.append([])
                        pass
                    elif vala[0] == "Relay":
                        relay1 = []
                        ra = OlxAPILib.getAllRelay(rg1)
                        rlyID = OlxAPILib.getRelayID(ra)
                        #
                        for i in range(1,len(vala)):
                            test = True
                            for j in range(len(ra)):
                                if rlyID[j]==vala[i]:
                                    relay1.append([ra[j],rlyID[j]])
                                    test = False
                                    break
                            if test:
                                 self.warningLog("Relay id incorect: ",kr1,i+1,1)
                        #
                        self.relay.append(relay1)
                        self.sRelay.append(vala)
                        kr1 += 1
                    else:
                        self.errData("Value must be 'Relay/Relay Group'",kr1,1,1,sheet='config')
            #
        return kr1+1
    #
    def getNextCellNoNone(self,kr):
        kr1 = 0
        while True:
            vala = self.ws.getValRowf(row=kr+kr1,columnFr=1,leng=1)
            if vala[0]!="":
                break
            if kr1>=30:
                return -1
            kr1+=1
        return kr+kr1
    #
    def getBranchOutage(self,kr):
        kr1 = self.getNextCellNoNone(kr)
        if kr1<0:
            return -1
        vala = self.ws.getValRow(row=kr1,columnFr=1)
        #
        self.sBranchOutage = []
        if vala[0] != "BRANCH OUTAGES":
            self.errData("Value must be 'BRANCH OUTAGES'",kr,1,1,sheet='config')
        self.sBranchOutage.append(vala)
        #
        while True:
            kr1+=1
            vala = self.ws.getValRow(row=kr1,columnFr=1)
            if vala[0]=="":
                break
            #
            self.sBranchOutage.append(vala)
            if not str(vala[0]).startswith("#"):
                try:
                    br1 = self.getBranchHnd(vala[1:6])
                except:
                    br1 = 0
                #
                if br1<=0:
                    self.warningLog("Branch Outage Not Found : ",row=kr1,column=2,nc=5)
                else:
                    self.outageLst.append(br1)
        #
        return kr1+1
    #
    def getFaultDescription(self,kr):
        kr1 = self.getNextCellNoNone(kr)
        if kr1<0:
            return -1
        vala = self.ws.getValRow(row=kr1,columnFr=1)
        #
        self.rowFaultSpec = [kr1]
        self.sFaultSpec = []
        #
        self.sFaultSpec.append(vala)
        if vala[0] != "FAULTS DESCRIPTION":
            self.errData("Value must be 'FAULTS DESCRIPTION'",kr,1,1,sheet='config')
        #
        fNo = []
        while True:
            kr1+=1
            vala = self.ws.getValRow(row=kr1,columnFr=1)
            if vala[0]=="":
                break
            #
            if not str(vala[0]).startswith("#"):
                self.sFaultSpec.append(vala)
                self.rowFaultSpec.append(kr1)
                try:
                    n1 = int(vala[0])
                except:
                    n1 = -1
                #
                if n1>0 and n1 not in fNo:
                    fNo.append(n1)
                else:
                    self.errData("Error value 'Fault No' Integer>0 and No Duplicate ",kr1,1,1,sheet='config')
        #
        return kr1+1
    #
    def getRelayQR_0(self,kr):
        kr1 = self.getNextCellNoNone(kr)
        if kr1<0:
            return -1
        vala = self.ws.getValRow(row=kr1,columnFr=1)
        #
        if vala[0] != "RELAY QUANTITIES REPORT":
            self.errData("Value must be 'RELAY QUANTITIES REPORT'",kr1,1,1,sheet='result')
        #
        kr1 += 1
        vala = self.ws.getValRowf(row=kr1,columnFr=1,leng=2)
        if vala[0]!="OLR FILE":
            self.errData("Value must be 'OLR FILE'",kr1,1,1,sheet='result')
        self.sOLR0 = vala
        #
        kr1 += 1
        vala = self.ws.getValRowf(row=kr1,columnFr=1,leng=2)
        if vala[0]!="DATE":
            self.errData("Value must be 'DATE'",kr1,1,1,sheet='result')
        self.sDate0 = vala
        #
        kr1 += 1
        vala = self.ws.getValRow(row=kr1,columnFr=1)
        if vala[0] != "Fault No":
            self.errData("Value must be 'Fault No'",kr1,1,1,sheet='result')

        #
        while True:
            kr1+=1
            vala = self.ws.getValRowf(row=kr1,columnFr=1,leng=17)
            if vala[0]=="":
                break
            #
            try:
                self.relayQR_ak_0.append([int(vala[0]),vala[1],vala[2]])
            except:
                self.errData("Error: Not a Integer number",kr1,1,1,sheet='result')
            #
            key1 = str(vala[0])+"_"+str(vala[1])+"_"+str(vala[2])
            self.relayQR_key_0.append(key1)
            #
            v1 = [ vala[3] ]
            for i in range(4,len(vala)):
                try:
                    v1.append(float(vala[i]))
                except:
                    if vala[i]=="":
                        v1.append(vala[i])
                    else:
                        self.errData("Error: Not a number ("+vala[i]+")",kr1,i+1,1,sheet='result')
            self.relayQR_VAL_0[key1] = v1
        #
        return kr1+1
    #
    def getData(self):
        kr = self.getOLR(kr=2)
        kr = self.getStudy(kr)
        kr = self.getFaultLocation(kr)
        kr = self.getFaultType(kr)
        kr = self.getOutage(kr)
        kr = self.getOutageType(kr)
        kr = self.getFaultZ(kr)
        kr = self.getRelay(kr)
        #
        kr = self.getBranchOutage(kr)
        if kr<0:
            return
        #
        kr = self.getFaultDescription(kr)
        if kr<0:
            return
        #
        kr = self.getRelayQR_0(kr)
        if kr<0 and self.ck>0: # excel
            if "result" in self.ws.allSheet:
                self.ws.selectSheet("result")
                self.getRelayQR_0(kr=1)
        return
    #
    def getVal2Save(self):
        valC = 0.1 # val% warning if dif
        ac = [] # config
        ac.append(["STUDY CONFIGURATION"])
        ac.append(self.sOLR)
        ac.append(self.sstudy)
        ac.append(self.sFaultLocation)
        ac.append(self.sFaultType)
        ac.append(self.sOutage)
        ac.append(self.sOutageType)
        ac.append(self.sFaultZ)
        ac.extend(self.sRelay)
        ac.append([])
        ac.extend(self.sBranchOutage)
        ac.append([])
        if self.bo!='yes':
            ac.extend(self.sFaultSpec)
        #
        ac.append([])

        ar1 = [] # result ------------------------------------------------------
        #-----------------------------------------------------------------------
        vf = ["Fault No","Relay group","Relay ID","Relay Op. Qty (amps or ohm)","RelayTime (s)",\
              "IAmag","IAang","IBmag","IBang","Icmag","ICang","Vamag","VAang","VBmag","VBang","VCmag","VCang"]
        #
        ar1.append(["RELAY QUANTITIES REPORT"])
        ar1.append(self.sOLR)
        self.da = getStrNow()
        self.sDate1 = ["DATE",self.da]
        ar1.append(self.sDate1)
        #
        ar1.append(vf)
        for i in range(len(self.relayQR_ak_1)):
            a1 = self.relayQR_ak_1[i][:]
            key1 = self.relayQR_key_1[i]
            a1.extend(self.relayQR_VAL_1[key1])
            ar1.append(a1)
        ar1.append([])

        ar2 = [] # deviation ---------------------------------------------------
        ar2.append(["RELAY QUANTITIES DEVIATION REPORT"])
        #
        ar2.append(["THIS OLR FILE",self.sOLR[1]])
        ar2.append(["DATE",self.sDate1[1]])
        #
        ar2.append(["BASE OLR FILE",self.sOLR0[1]])
        ar2.append(["BASE DATE",self.sDate0[1]])
        ar2.append([])
        #
        v1 = vf[:3]
        for i in range(3,len(vf)):
            v1.extend([vf[i],"",""])
        ar2.append(v1)
        v1 = ["","",""]
        for i in range(14):
            v1.extend(["This run","Baseline","% deviation"])
        ar2.append(v1)
        #
        kdif = 0
        #
        for i in range(len(self.relayQR_ak_1)):
            a1 = self.relayQR_ak_1[i][:]
            key1 = self.relayQR_key_1[i]
            va1 = self.relayQR_VAL_1[key1]
            try:
                va0 = self.relayQR_VAL_0[key1]
            except:
                va0 = []
            #
            c1,k = compare(va1,va0,valC)
            a1.extend(c1)
            kdif += k
            #
            ar2.append(a1)

        # exist in baseline but not found in this run
        for i in range(len(self.relayQR_ak_0)):
            a1 = self.relayQR_ak_0[i][:]
            key1 = self.relayQR_key_0[i]
            va0 = self.relayQR_VAL_0[key1]
            try:
                va1 = self.relayQR_VAL_1[key1]
            except:
                va1 = []
            if len(va1)==0:
                c1,k = compare(va1,va0,valC)
                a1.extend(c1)
                kdif += k
                #
                ar2.append(a1)
        #
        ar2.append([])
        ar2.append(["number diff",str(kdif)])
        #
        return ac,ar1,ar2
    #
    def saveReport(self):
        ac,ar1,ar2 = self.getVal2Save()
        #-----------------------------------------------------------------------
        ext = (os.path.splitext(ARGVS.fo)[1]).upper()
        if ext=='':
            ext='.XLSX'
        #
        self.resFile = []
        sp = "Result file had been saved in:"
        if ext=='.CSV':
            fo = get_file_out(fo=ARGVS.fo, fi=self.ft , subf='res' , ad='_res', ext='.CSV')
            base = os.path.splitext(fo)[0]
            #
            fe = base +"_config.csv"
            save2CSV(fe,ac,",")
            self.resFile.append(fe)
            sp += "\n  "+fe
            #
            if self.bo!='yes': # NO branch outage review
                fe = base +"_result.csv"
                save2CSV(fe,ar1,",")
                self.resFile.append(fe)
                sp += "\n  "+fe
                #
                fe = base +"_deviation.csv"
                save2CSV(fe,ar2,",")
                self.resFile.append(fe)
                sp +="\n  "+fe
        else:
            fo = get_file_out(fo=ARGVS.fo, fi=self.ft , subf='res' , ad='_res', ext=ext)
            if self.bo!='yes':# NO branch outage review
                self.ws.save2Excel(fo ,[ac,ar1,ar2],["config","result","deviation"])
            else:
                self.ws.save2Excel(fo ,[ac],["config"])
            #
            self.resFile.append(fo)
            sp+="\n  "+fo
        #
        print("runtime: %.1fs" %(time.time()- self.t0))
        print(sp)
    #
    def openExcelResult(self):
        for se in self.resFile:
            if ARGVS.ut==0:
                launch_excel(se)
    #
    def run(self):
        self.getData()
        #
        if len(self.sFaultSpec)>1:
            self.runSIMULATEFAULT()
        else:
            self.run2BranchOutage()
            if self.bo!='yes':
                self.getFltOpt()
                self.runDoFault()
        #
        self.saveReport()
        #
        self.openExcelResult()
        return 1
#
class MainGUI(tk.Frame):
    def __init__(self,master):
        tk.Frame.__init__(self, master=master)
        self.sw = self.master.winfo_screenwidth()
        self.sh = self.master.winfo_screenheight()
        w,h = 900,500 #
        self.master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(self.sw/2-w/2),int(self.sh/2-h/2)))
        self.master.resizable(0,0)# fixed size
        self.master.wm_title("FAULT DIFF TOOL")
        setIco_1Liner(self.master)
        remove_button(self.master)
        # registry
        self.reg = WIN_REGISTRY(path = "SOFTWARE\ASPEN\OneLiner\PyManager\FltDiffTool",keyUser="",nmax =1)
        self.initGUI()
   #
    def write(self, txt):
        self.text1.insert(tk.INSERT,txt)
        #
    def flush(self):
        pass
    #
    def close_buttons(self):
        self.master.destroy()
    #
    def clearConsol(self):
        self.text1.delete(1.0,tk.END)
    #
    def editInput(self):
        launch_excel(self.ipf_v.get())
    #
    def editOutput(self):
        launch_excel(self.opf_v.get())
    #
    def initGUI(self):
        sys.stdout = self
        #
        fileFrame = tk.LabelFrame(self.master, text = "Files")
        ipf = tk.Label(fileFrame, text="Input File : ")
        ipf.grid(row=0, column=0, sticky='E', padx=5, pady=5)

        self.ipf_v = tk.StringVar()
        if ARGVS.ft!="":
            self.ipf_v.set(os.path.abspath(ARGVS.ft))
        else:
            try:
                ft1 = self.reg.getAllValue()[0]
                if os.path.isfile(ft1):
                    self.ipf_v.set(ft1)
            except:
                pass
        ipfTxt = tk.Entry(fileFrame,width= 103,textvariable=self.ipf_v)
        ipfTxt.grid(row=0, column=1, sticky="W",padx=5, pady=5)
        #
        ipf_b1 = tk.Button(fileFrame, text="...",width= 6,relief= tk.GROOVE,command=self.selectInputFile)
        ipf_b1.grid(row=0, column=2, sticky='W', padx=5, pady=5)
        ipf_b2 = tk.Button(fileFrame, text="Edit in Excel",width= 10,relief= tk.GROOVE,command=self.editInput)
        ipf_b2.grid(row=0, column=3, sticky='W', padx=5, pady=5)
        #
        csFrame = tk.LabelFrame(self.master, relief= tk.GROOVE, bd=0)
        #
        self.text1 = tk.Text(csFrame,wrap = tk.NONE,width=108,height=20)#,
        # yScroll
        yscroll = tk.Scrollbar(csFrame, orient=tk.VERTICAL, command=self.text1.yview)
        self.text1['yscroll'] = yscroll.set
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        #
        # xScroll
        xscroll = tk.Scrollbar(csFrame, orient=tk.HORIZONTAL, command=self.text1.xview)
        self.text1['xscroll'] = xscroll.set
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        #
        self.text1.pack(fill=tk.BOTH, expand=tk.Y)

        #
        btFrame = tk.Frame(self.master, relief= tk.GROOVE, bd=0)
        #
        self.run_b = tk.Button(btFrame, text="Launch",relief= tk.GROOVE,width= 10, command=self.run1)
        self.run_b.grid(row=0,column=0, padx=0, pady=5)
        #
        close_b = tk.Button(btFrame, text="Exit",width= 10,relief= tk.GROOVE, command=self.close_buttons)
        close_b.grid(row=0,column=1, padx=35, pady=5)
        #
        clearCs_b = tk.Button(btFrame, text="Clear console",width= 10,relief= tk.GROOVE, command=self.clearConsol)
        clearCs_b.grid(row=0,column=2, padx=0, pady=5)

        #
        self.pgFrame = tk.Frame(self.master, relief= tk.GROOVE, bd=0)
        self.var1 = tk.StringVar()
        self.percent =  ttk.Label(self.pgFrame , textvariable=self.var1,width= 18) #
        self.percent.grid(row=0,column=0, padx=10, pady=5)
        #
        self.progress = ttk.Progressbar(self.pgFrame,orient='horizontal',length=250,maximum=100, mode='determinate')
        self.progress.grid(row=0,column=1, padx=10, pady=5)
        #
        self.stop_b = tk.Button(self.pgFrame, text="Stop",relief= tk.GROOVE,width= 10, command=self.stop_progressbar)
        self.stop_b.grid(row=0,column=2, padx=10, pady=5)
        self.stop_b['state']='disabled'
        #
        fileFrame.grid(row=0, sticky='W', padx=10, pady=5, ipadx=5, ipady=5)
        csFrame.grid(row=1, column=0, padx=10, pady=5)
        btFrame.grid(row=3, column=0, padx=10, pady=2)
        #
        sys.stdout = self
    #
    def selectInputFile(self):
        # Excell/CSV file
        v1 = tkf.askopenfilename(filetypes=[('Excel/csv file','*.xlsx *.xls .*csv'), ("All Files", "*.*")],title='Select Input file')
        if v1!='':
            self.ipf_v.set(v1)
            ARGVS.ft = v1
            self.reg.appendValue(v1)
     #
    def selectOutputFile(self):
        v1 = tkf.asksaveasfile(defaultextension=".xlsx", filetypes=(("Excel file", "*.xlsx"),("All Files", "*.*") ))
        try:
            self.opf_v.set(v1.name)
        except:
            pass
    #
    def simulate(self):
        self.pgFrame.grid(row=3,column=0, padx=0, pady=0)
        self.stop_b['state']='disabled'
        self.var1.set("Reading data")
        try:
            dt = FltDiffTool(ARGVS.ft)
            dt.getData()
            self.stop_b['state']='active'
            self.var1.set("Running")
            self.progress['value'] = 0
            v1 = 0
            #
            if len(dt.sFaultSpec)>1:
                pMax = len(dt.sFaultSpec)-1
                sMax = "/"+str(pMax)
                #
                for i in range(1,len(dt.sFaultSpec)):
                    s1 = dt.sFaultSpec[i]
                    k = int(s1[0])
                    fdes0 = s1[1]
                    #
                    try:
                        run1SIMULATEFAULT(opt=s1[2],fltstr=s1[3])
                        fdes = dt.getResult1(k)
                    except:
                        dt.errData("Error Fault Description",row=dt.rowFaultSpec[i],column=1,nc=4,sheet='config')
                    #
                    v1 = round(i/pMax*100)
                    self.progress['value'] = v1
                    self.var1.set("Running: " +str(i) + sMax)
            else:
                dt.run2BranchOutage()
                #
                if dt.bo=='':
                    dt.bo = tkm.askquestion("", "Do you want to review BRANCH OUTAGES?")
                #
                if dt.bo!='yes': # no review branch outages
                    dt.getFltOpt()
                    #
                    pMax = max(len(dt.rx)* len(dt.fltConn)* len(dt.fltOpt),1)
                    sMax = "/"+str(pMax)
                    bhnd = dt.bus0
                    if dt.br0>0:
                        bhnd = dt.br0
                    #
                    k = 0
                    i = 0 # progress
                    for rx in dt.rx:
                        for fltConn in dt.fltConn:
                            for fltOpt in dt.fltOpt:
                                i+=1
                                #
                                OlxAPILib.doFault(bhnd,fltConn,fltOpt,dt.outageOpt,dt.outageLst, rx[0], rx[1], clearPrev=1)
                                k = dt.getResult(k)
                                #
                                v1 = round(i/pMax*100)
                                self.progress['value'] = v1
                                self.var1.set("Running: " +str(i) + sMax)
            #
            if v1==100 or dt.bo=='yes':
                self.var1.set("Saving...")
                self.stop_b['state']='disabled'
                dt.saveReport()
                #
                self.var1.set("Openning Excel...")
                dt.openExcelResult()
        except Exception as err:
            print(err)
        #
        self.finish()
    #
    def finish(self):
        self.run_b['state']='active'
        self.stop_b['state']='disabled'
        self.progress.stop()
        self.pgFrame.grid_forget()
    #
    def stop_progressbar(self):
        self.finish()
        self.t.killed = True
    #
    def run1(self):
        ARGVS.ft = self.ipf_v.get()
        if os.path.isfile(ARGVS.ft):
            self.reg.appendValue(ARGVS.ft)
        else:
            print('Input file not found:\n%s'%ARGVS.ft)
            return
        #
        try:
            self.progress.start()
            self.t = TraceThread(target=self.simulate)
            self.t.start()
            #
        except Exception as e:
            print(str(e))

        self.run_b['state']='disabled'
        self.stop_b['state']='active'
        #
#
def run_demo():
    if ARGVS.demo in [1,2,3]:
        if ARGVS.demo==1:
            sMain = "FltDiffTool demo (with Branch Outages review)"
            ARGVS.ft = os.path.join( PATH_FILE,"fltDiff_InputFile1.xlsx")
            ARGVS.bo = 'yes'
        elif ARGVS.demo==2:
            sMain = "FltDiffTool demo (without Branch Outages review)"
            ARGVS.ft = os.path.join( PATH_FILE,"fltDiff_InputFile1.xlsx")
            ARGVS.bo = 'no'
        else :
            sMain = "FltDiffTool demo (comparison with previous run)"
            ARGVS.ft = os.path.join( PATH_FILE,"fltDiff_InputFile2.xlsx")
        #
        ARGVS.fo = get_file_out(fo='' , fi=ARGVS.ft , subf='res' , ad='_demo', ext='.xlsx')
        #
        choice = ask_run_demo(PY_FILE,ARGVS.ut,'',"Input template file: " + ARGVS.ft)
        if choice=="yes":
            dt = FltDiffTool(ARGVS.ft)
            return dt.run()
    else:
        demo_notFound(PY_FILE,ARGVS.demo,[1,2,3])
#
def main():
    if ARGVS.demo>0:
        return run_demo()
    #
    root = tk.Tk()
    feedback = MainGUI(root)
    root.mainloop()
#
if __name__ == '__main__':
    # ARGVS.demo = 1
    try:
        main()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)
