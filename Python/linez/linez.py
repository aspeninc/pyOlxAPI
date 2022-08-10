"""
Purpose: Report total line impedance and length
        All line from a bus (picked up)
        All taps are ignored. Close switches,Serie capacitor/reactor bypass are included

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

# IMPORT -----------------------------------------------------------------------
import sys,os,time
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPILib
from OlxAPIConst import *
import AppUtils
# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage=
        "\n\tReport total line impedance and length\
        \n\tAll lines from a bus defined by (name,kV) or bus Number\
        \n\tAll taps are ignored. Close switches, Serie capacitor/reactor bypass are included")
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-pk' , help = '*(str) Selected Bus in the 1-line diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name of the report file ',default = '',type=str,metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)
#
def linez_fromBus(bhnd):
    """
     Report total line impedance and length
        All line from a bus
        All taps are ignored. Close switches are included
        serie capacitor/reactor bypass are ignored

        Args :
            bhnd = bus handle

        Returns:
            nC : number of lines
            sres: string result

        Raises:
            OlrxAPIException
    """
    typeConsi = [TC_LINE,TC_SWITCH,TC_SCAP]
    allBr = OlxAPILib.getBusEquipmentData([bhnd],TC_BRANCH)[0]
    sres = ""
    sres += str("No").ljust(5)+str(",Type   ").ljust(20)
    sres += "," + str("Bus1").ljust(30)+","+str("Bus2").ljust(30)+",ID  "
    sres += "," + "Z1".ljust(35)
    sres += "," + "Z0".ljust(35)
    sres += ",Length"
    #
    kd = 0
    for hndBr in allBr:
        if OlxAPILib.branchIsInType(hndBr,typeConsi):
            bra = OlxAPILib.lineComponents(hndBr,typeConsi)
            for bra1 in bra:
                kd +=1
                s1 = getsImpedanceLine(kd,bra1)
                sres +=s1
    return kd,sres
#
def linez_fromBus_str(bhnd,kd,sres):
    #
    s1 = "\n"+ str(kd)+" Line impedances from bus: " + OlxAPILib.fullBusName(bhnd) +"\n"
    s1+=sres
    return s1
#
def getsImpedanceLine(kd,bra): #[TC_LINE,TC_SWITCH,TC_SCAP]
    sres = ""
    typa =  OlxAPILib.getEquipmentData(bra,BR_nType)
    #
    b1 = OlxAPILib.getEquipmentData([bra[0]],BR_nBus1Hnd)[0]
    b2 = OlxAPILib.getEquipmentData([bra[len(bra)-1]],BR_nBus1Hnd)[0]
    #
    kV = OlxAPILib.getEquipmentData([b1],BUS_dKVnominal)[0]
    baseMVA = OlxAPILib.getParaSys(cod=SY_dBaseMVA)

    #
    if len(bra)==2: # simple line
        br1 = bra[0]
        typ1 = typa[0]
        equi_hnd = OlxAPILib.getEquipmentData([br1],BR_nHandle)[0]
        #
        if typ1== TC_LINE:
            r1 = OlxAPILib.getEquipmentData([equi_hnd],LN_dR)[0]
            x1 = OlxAPILib.getEquipmentData([equi_hnd],LN_dX)[0]
            r0 = OlxAPILib.getEquipmentData([equi_hnd],LN_dR0)[0]
            x0 = OlxAPILib.getEquipmentData([equi_hnd],LN_dX0)[0]
            length = OlxAPILib.getEquipmentData([equi_hnd],LN_dLength)[0]
            #
            idBr = OlxAPILib.getEquipmentData([equi_hnd],LN_sID)[0] # id line
            len_unit = OlxAPILib.getEquipmentData([equi_hnd], LN_sLengthUnit)[0]
            sres += stringImpedance(kd,",Line",b1,b2,idBr,r1,x1,r0,x0,kV,length,baseMVA) + len_unit
        elif typ1 == TC_SWITCH:
            idBr = OlxAPILib.getEquipmentData([equi_hnd],SW_sID)[0] # id switch
            #
            sres +="\n" + str(kd).ljust(5) + str(",Switch").ljust(16) + ","
            sres += str(OlxAPILib.fullBusName(b1)).ljust(30) + ","
            sres += str(OlxAPILib.fullBusName(b2)).ljust(30) + ","
            sres += str(idBr).ljust(4) + ","
        elif typ1 == TC_SCAP:
            idBr = OlxAPILib.getEquipmentData([equi_hnd],SC_sID)[0] # id switch
            st = OlxAPILib.getEquipmentData([equi_hnd], SC_nInService)[0]
            if st==3:#bypass
                sres +="\n" + str(kd).ljust(5) + str(",SerieCap-bypass").ljust(20) + ","
                sres += str(OlxAPILib.fullBusName(b1)).ljust(30) + ","
                sres += str(OlxAPILib.fullBusName(b2)).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
            else:
                r1 = OlxAPILib.getEquipmentData([equi_hnd], SC_dR)[0]
                x1 = OlxAPILib.getEquipmentData([equi_hnd], SC_dX)[0]
                r0 = OlxAPILib.getEquipmentData([equi_hnd], SC_dR0)[0]
                x0 = OlxAPILib.getEquipmentData([equi_hnd], SC_dX0)[0]
                sres += stringImpedance(kd,",SerieCap",b1,b2,idBr,r1,x1,r0,x0,kV,0,baseMVA)
    else: #Line with Tap bus
        r1a,x1a,r0a,x0a,lengtha = 0,0,0,0,0
        len_unit1 = ""
        hndLna = []
        for i in range(len(bra)-1):
            br1 = bra[i]
            typ1 = typa[i]
            b1i = OlxAPILib.getEquipmentData([br1],BR_nBus1Hnd)[0]
            b2i = OlxAPILib.getEquipmentData([br1],BR_nBus2Hnd)[0]
            #
            equi_hnd = OlxAPILib.getEquipmentData([br1],BR_nHandle)[0]
            if typ1== TC_LINE:
                hndLna.append(equi_hnd)
                idBr = OlxAPILib.getEquipmentData([equi_hnd],LN_sID)[0] # id line
                len_unit = OlxAPILib.getEquipmentData([equi_hnd], LN_sLengthUnit)[0]
                #
                r1 = OlxAPILib.getEquipmentData([equi_hnd],LN_dR)[0]
                x1 = OlxAPILib.getEquipmentData([equi_hnd],LN_dX)[0]
                r0 = OlxAPILib.getEquipmentData([equi_hnd],LN_dR0)[0]
                x0 = OlxAPILib.getEquipmentData([equi_hnd],LN_dX0)[0]
                length = OlxAPILib.getEquipmentData([equi_hnd],LN_dLength)[0]
                #
                r1a,x1a,r0a,x0a = r1a+r1,x1a+x1,r0a+r0,x0a+x0
                #
                len_unit1,valC = AppUtils.convert_LengthUnit(len_unit1,len_unit)
                lengtha += length *valC
                #
                sres += stringImpedance("",",  Line_seg",b1i,b2i,idBr,r1,x1,r0,x0,kV,length,baseMVA) + len_unit
            elif typ1 == TC_SWITCH:
                idBr = OlxAPILib.getEquipmentData([equi_hnd],SW_sID)[0] # id switch
                #
                sres +="\n" + str("").ljust(5) + str(",  Switch_seg").ljust(20) + ","
                sres += str(OlxAPILib.fullBusName(b1i)).ljust(30) + ","
                sres += str(OlxAPILib.fullBusName(b2i)).ljust(30) + ","
                sres += str(idBr).ljust(4) + ","
            elif typ1 == TC_SCAP:
                idBr = OlxAPILib.getEquipmentData([equi_hnd], SC_sID)[0] # id switch
                st   = OlxAPILib.getEquipmentData([equi_hnd], SC_nInService)[0]
                if st==3:# bypass
                    sres +="\n" + str("").ljust(5) + str(",SerieCap_seg-bypass").ljust(20) + ","
                    sres += str(OlxAPILib.fullBusName(b1)).ljust(30) + ","
                    sres += str(OlxAPILib.fullBusName(b2)).ljust(30) + ","
                    sres += str(idBr).ljust(4) + ","
                else:
                    r1 = OlxAPILib.getEquipmentData([equi_hnd], SC_dR)[0]
                    x1 = OlxAPILib.getEquipmentData([equi_hnd], SC_dX)[0]
                    r0 = OlxAPILib.getEquipmentData([equi_hnd], SC_dR0)[0]
                    x0 = OlxAPILib.getEquipmentData([equi_hnd], SC_dX0)[0]
                    r1a,x1a,r0a,x0a = r1a+r1,x1a+x1,r0a+r0,x0a+x0
                    sres += stringImpedance("",",  SerieCap_seg",b1i,b2i,idBr,r1,x1,r0,x0,kV,0,baseMVA)
        #summary
        if OlxAPILib.isMainLine(hndLna):
            sres += stringImpedance(kd,",Line_sum",b1,b2,"",r1a,x1a,r0a,x0a,kV,lengtha,baseMVA) + len_unit1
        else:
            sres +="\n" + str(kd).ljust(5) + str(",Tap Line").ljust(20) + ","
            sres += str(OlxAPILib.fullBusName(b1)).ljust(30) + ","
            sres += str(OlxAPILib.fullBusName(b2)).ljust(30) + ","
    return sres +"\n"

def stringImpedance(kd,s1,b1,b2,idBr,r1,x1,r0,x0,kV,length,baseMVA):
    sres ="\n" + str(kd).ljust(5) + s1.ljust(20) + ","
    sres += str(OlxAPILib.fullBusName(b1)).ljust(30) + ","
    sres += str(OlxAPILib.fullBusName(b2)).ljust(30) + ","
    sres += str(idBr).ljust(4) + ","
    sres += AppUtils.printImpedance(r1,x1,kV,baseMVA).ljust(35) + ","
    sres += AppUtils.printImpedance(r0,x0,kV,baseMVA).ljust(35) + ","
    sres +=  "{0:.2f}".format(length)
    return sres

def unit_test():
    ARGVS.fi = os.path.join (PATH_FILE, "LINEZ.OLR")
    ARGVS.pk = ["[BUS] 28 'ARIZONA' 132 kV", "[BUS] 5 'FIELDALE' 132 kV"]
    #
    sres ='OLR file: ' + os.path.basename(ARGVS.fi) +'\n'
    run()
    sres += AppUtils.read_File_text_2(ARGVS.fo,4)
    #
    AppUtils.deleteFile(ARGVS.fo)
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)

#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.txt')
        ARGVS.pk = ["[BUS] 6 'NEVADA' 132 kV", "[BUS] 14 'MONTANA' 33 kV"]
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Bus: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def run():
    """
    run linez:
    """
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Bus',PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    sres = 'App : ' + PY_FILE
    sres +='\nUser: ' + os.getlogin()
    sres +='\nDate: ' + time.asctime()
    sres +='\nOLR file: ' + ARGVS.fi
    print('Selected Bus:' + str(ARGVS.pk))
    #
    for i in range(len(ARGVS.pk)):
        sres+= '\n\nSelected Bus:' + str(ARGVS.pk[i])
        # get handle of object picked up
        bhnd = OlxAPILib.FindObj1LPF_check(ARGVS.pk[i],TC_BUS)
        #
        nTap1 = OlxAPILib.getEquipmentData([bhnd],BUS_nTapBus)[0]
        if nTap1>0:
            sMain = "\nTap bus found: " +ARGVS.pk[i]
            sMain += "\nUnable to continue."
            AppUtils.gui_info('INFO',sMain)
            return
        #
        kd,sr = linez_fromBus(bhnd)
        sres += linez_fromBus_str(bhnd,kd,sr)
    #
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='res', ad='_res' , ext='.txt')
    #
    AppUtils.saveString2File(ARGVS.fo,sres)
    print("\nReport file had been saved in:\n" +ARGVS.fo )
    if ARGVS.ut==0 :
        AppUtils.launch_notepad(ARGVS.fo)
    return 1
#
def main():
    if ARGVS.ut>0:
        return unit_test()
    if ARGVS.demo>0:
        return run_demo()
    run()
#
if __name__ == '__main__':
    # ARGVS.ut = 1
    # ARGVS.demo = 1
    try:
        main()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)




