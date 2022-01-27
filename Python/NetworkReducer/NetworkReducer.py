"""
Purpose:
     Reduce a network to a smaller boundary equivalent
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.2"

# IMPORT -----------------------------------------------------------------------
import sys,os,time
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
import AppUtils

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage="\n\tReduce a network to a smaller boundary equivalent.")
#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR input file path' , default = "", metavar='')
PARSER_INPUTS.add_argument('-pk' , help = '*(str) Selected Objects in the 1-line diagram', default = [],nargs='+',metavar='')
#
PARSER_INPUTS.add_argument('-fb' , help = ' (str) Report bus file', default = "", metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) OLR ouput file path (only when OLR reduced file is desired)', default = "", metavar='')
#
PARSER_INPUTS.add_argument('-ar' , help = ' (int) area number (default=0)',default = [] , nargs='+', metavar='')
PARSER_INPUTS.add_argument('-zo' , help = ' (int) zone number (default=0)',default = [] , nargs='+', metavar='')
PARSER_INPUTS.add_argument('-ti' , help = ' (int) max tier (default=1)',default = 1 , type=int, metavar='')
#
PARSER_INPUTS.add_argument('-opt1', help = ' (int) Per-unit elimination threshold (default=99)', default = 99,type=int, metavar='')
PARSER_INPUTS.add_argument('-opt2', help = ' (int) Keep existing equipment at retained buses [0- reset,1- set (default)]', default = 1,type=int, metavar='')
PARSER_INPUTS.add_argument('-opt3', help = ' (int) Keep all existing annotations [0- reset,1- set (default)]', default = 1,type=int, metavar='')
#
PARSER_INPUTS.add_argument('-gui' , help = ' (int) $ASPEN internal parameter$ GUI [0-No GUI, 1-GUI (default)]', default = 1,type=int, metavar='')
#
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

#
class NetworkReducer:
    def __init__(self):
        print(ARGVS.fi)
        OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)

        #
        self.equipA = set()
        self.busA = set()
        #
        self.brs = OlxAPILib.BranchSearch(gui=0)
    #
    def saveReportBusFile(self):
        ARGVS.fb = AppUtils.get_file_out(fo=ARGVS.fb , fi=ARGVS.fi , subf='' , ad='_bus' , ext='.txt')
        #
        sr = []
        sr.append('App : ' + PY_FILE)
        sr.append('User: ' + os.getlogin())
        sr.append('Date: ' + time.asctime())
        sr.append('\nBus report:')

        k = 0
        for b1 in self.busA:
            k+=1
            sb = OlxAPILib.fullBusName(b1)
            sr.append(str(k).ljust(5)+sb)
        if k==0:
            sr.append('\tNo bus found')
        #
        AppUtils.saveArString2File(ARGVS.fb,sr)
        print('\nReport bus file had been saved in:\n"%s"'%ARGVS.fb)
        #
        if ARGVS.gui>0 and ARGVS.ut==0:
            AppUtils.launch_notepad(ARGVS.fb)
    #
    def runBoundaryEquivalent(self):
        #
        ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='res' , ad='_res' , ext='.OLR')
        #
        opt = [ARGVS.opt1, ARGVS.opt2, ARGVS.opt3]
        OlxAPILib.BoundaryEquivalent(ARGVS.fo, list(self.busA),opt)
        #
        OlxAPI.CloseDataFile()
        if not os.path.isfile(ARGVS.fo):
            raise Exception(" Run BoundaryEquivalent")
        #
        if ARGVS.gui>0 and ARGVS.ut==0:
            AppUtils.launch_OneLiner(ARGVS.fo)
        print('\nOLR result file had been saved in:\n"%s"'%ARGVS.fo)
    #
    def get_equipByBus(self,bus0):
        res = set()
        for i in range(len(bus0)):
            for j in range(len(bus0)):
                if i!=j:
                    self.brs.setBusHnd1(bus0[i])
                    self.brs.setBusHnd2(bus0[j])
                    br1 = self.brs.runSearch()
                    #
                    if br1>0:
                        e1 = OlxAPILib.getEquipmentData([br1],BR_nHandle)[0]
                        res.add(e1)
        return res
    #
    def getBusPK(self):
        self.busPK = set()
        self.equipPK = set()
        bus0 = []
        #
        TC1 = [TC_LOAD,TC_LOADUNIT,TC_SHUNT,TC_SHUNTUNIT,TC_GEN,TC_GENUNIT]
        TC2 = [TC_LINE,TC_SCAP,TC_PS,TC_SWITCH,TC_XFMR,TC_XFMR3]
        for i in range(len(ARGVS.pk)):
            s1 = str(ARGVS.pk[i]).strip()
            hnd,tc = OlxAPILib.FindObj1LPF(s1)
            if hnd<=0:
                raise Exception("\nObject not found:"+str(ARGVS.pk))
            if tc== TC_BUS:
                bus0.append(hnd)
                self.busPK.add(hnd)
            elif tc in TC1:
                self.busPK.add(hnd)
            elif tc in TC2:
                bres = OlxAPILib.getBusByEquipment(hnd,tc)
                self.busPK.update(bres)
                self.equipPK.add(hnd)
            elif tc== TC_RLYGROUP:
                # select one option here
##                self.getBusPK_RLGlikeBranch(hnd) # line branch
                self.getBusPK_RLG(hnd) # line terminal
            else:
                raise Exception("not supported yet:"+str(ARGVS.pk))
        #
        self.equipPK.update(self.get_equipByBus(bus0))
    #
    def getBusPK_RLGlikeBranch(self,hnd):
         # like
        bra  = OlxAPILib.getEquipmentData([hnd],RG_nBranchHnd)
        for br1 in bra:
            bres = OlxAPILib.getBusByBranch(br1)
            self.busPK.update(bres)
            #
            e1 = OlxAPILib.getEquipmentData([br1],BR_nHandle)[0]
            self.equipPK.add(e1)
    #
    def getBusPK_RLG(self,hnd):
        br1 = OlxAPILib.getEquipmentData([hnd],RG_nBranchHnd)[0]
##        br_res   [] list of branch terminal
##        bus_res  [] list of terminal bus
##        bus_resa [] list of all bus
##        equip    [] list of all equipement
        br_res,bus_res,bus_resa,equip = OlxAPILib.getRemoteTerminals(hndBr=br1,typeConsi=[])
        self.busPK.update(bus_resa)
        self.equipPK.update(equip)

    #
    def get1(self,b1):
        busTerminal = set()
        abr = OlxAPILib.getBusEquipmentData([b1],TC_BRANCH)[0]
        for br1 in abr:
            inSer1 = OlxAPILib.getEquipmentData([br1],BR_nInService)[0]
            if inSer1==1:
                e1 = OlxAPILib.getEquipmentData([br1],BR_nHandle)[0]
                if e1 not in self.equipA:
                    br_res,bus_res,bus_resa,ea1 = OlxAPILib.getRemoteTerminals(br1,typeConsi=[])
                    self.equipA.update(ea1)
                    self.busA.update(bus_resa)
                    busTerminal.update(bus_res)
        return busTerminal
    #
    def runTier(self,tier,bus0):
        busThis = bus0[:]
        ti = 0
        while True:
            busTerminal = set()
            if ti>=tier:
                break
            ti +=1
            #
            for b1 in busThis:
                bt1 = self.get1(b1)
                busTerminal.update(bt1)
            #
            busThis = set()
            busThis.update(busTerminal)
        return busTerminal
    #
    def runTier_0(self,tier):
        self.getBusPK()
        self.equipA.update(self.equipPK)
        self.busA.update(self.busPK)
        self.runTier(ARGVS.ti,list(self.busA))
    #
    def run_area_zone(self,ar,zo):
        # get all bus with area = area, zone = zone
        bus_az = set()
        equiCode1 = [TC_LINE, TC_SCAP, TC_PS, TC_SWITCH, TC_XFMR, TC_XFMR3] #TC_LINE,
        #
        for ec1 in equiCode1:
            eqa = OlxAPILib.getEquipmentHandle(ec1)
            for ehnd1 in eqa:
                bus = OlxAPILib.getBusByEquipment(ehnd1,ec1)
                #
                if len(zo)==0:
                    ar1 = OlxAPILib.getEquipmentData(bus,BUS_nArea)
                    for i in range(len(bus)):
                        if ar1[i] in ar:
                            bus_az.add(bus[i])
                            if ec1!=TC_LINE:
                                for j in range(len(bus)):
                                    if i!=j:
                                        bus_az.add(bus[j])
                                break # for i
                #
                elif len(ar)==0:
                    zo1 = OlxAPILib.getEquipmentData(bus,BUS_nZone)
                    for i in range(len(bus)):
                        if zo1[i] in zo:
                            bus_az.add(bus[i])
                            if ec1!=TC_LINE:
                                for j in range(len(bus)):
                                    if i!=j:
                                        bus_az.add(bus[j])
                                break # for i
                #
                else:
                    ar1 = OlxAPILib.getEquipmentData(bus,BUS_nArea)
                    zo1 = OlxAPILib.getEquipmentData(bus,BUS_nZone)
                    for i in range(len(bus)):
                        if (ar1[i] in ar) and (zo1[i] in zo):
                            bus_az.add(bus[i])
                            if ec1!=TC_LINE:
                                for j in range(len(bus)):
                                    if i!=j:
                                        bus_az.add(bus[j])
                                break # for i
        #
        return bus_az
#
def checkInput():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return False
    #
    if not (ARGVS.pk or ARGVS.ar or ARGVS.zo):
        s = "\nSelected Objects,Area number,Zone number = Empty.\nUnable to continue.\n"
        return AppUtils.FinishCheck(PY_FILE,s,PARSER_INPUTS)
    #
    try:
        ARGVS.ar = list(map(int, ARGVS.ar))
    except:
        raise Exception('Area Number is not an INTEGER')
    try:
        ARGVS.zo = list(map(int, ARGVS.zo))
    except:
        raise Exception('Zone Number is not an INTEGER')
    return True

def run():
    print('Selected Objects: ',str(ARGVS.pk))
    print('Area Number     : ',str(ARGVS.ar))
    print('Zone Number     : ',str(ARGVS.zo))
    #
    if not checkInput():
        return
    #
    nr = NetworkReducer()
    #
    if len(ARGVS.ar)>0 or len(ARGVS.zo)>0 :
        bus_az = nr.run_area_zone(ARGVS.ar,ARGVS.zo)
        nr.busA.update(bus_az)
        nr.runTier(ARGVS.ti,list(nr.busA))
    else:# just tier with pikup objects
        nr.runTier_0(ARGVS.ti)
    #END
    nr.saveReportBusFile()
    if ARGVS.fo!='' and len(nr.busA)>0:
        nr.runBoundaryEquivalent()
    return 1
#
def run_demo():
    if ARGVS.demo ==1:
        ARGVS.fi = AppUtils.getASPENFile('','SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.OLR')
        ARGVS.pk = ["[XFORMER3] 6 'NEVADA' 132 kV-10 'NEVADA' 33 kV-'DOT BUS' 13.8 kV 1"]
        ARGVS.ti = 1
        ARGVS.ar = []
        ARGVS.zo = []
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Object: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def main():
    if ARGVS.demo>0:
        return run_demo()
    run()
#
if __name__ == '__main__':
    # ARGVS.demo = 1
    try:
        main()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)







