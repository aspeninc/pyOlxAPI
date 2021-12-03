"""
Purpose: List the relay group(s) on the remoted terminal(s) of the selected line or transformer

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
PARSER_INPUTS = AppUtils.iniInput(usage="\n\tList the Relay Group(s) on the remoted terminal(s)\n\tof the selected line or transformer")
PARSER_INPUTS.add_argument('-fi' , metavar='', help = '*(str) OLR input file', default = "")
PARSER_INPUTS.add_argument('-pk' , metavar='', help = '*(str) Selected Relay Group in the 1-line diagram', default = [],nargs='+')
PARSER_INPUTS.add_argument('-fo' , metavar='', help = ' (str) Path name of the report file',default = "")
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Relay Group',PY_FILE,PARSER_INPUTS):
        return
    #

    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    bhnd,tc = OlxAPILib.FindObj1LPF(ARGVS.pk[0])
    if tc!= TC_RLYGROUP:
        s  = "\nNo Relay Group is selected: "+str(ARGVS.pk)+ "\nUnable to continue."
        return AppUtils.FinishCheck(PY_FILE,s,PARSER_INPUTS)
    #
    br1 = OlxAPILib.getEquipmentData([bhnd],RG_nBranchHnd)[0]
    b1 = OlxAPILib.getEquipmentData([br1],BR_nBus1Hnd)[0]
    nTap1 = OlxAPILib.getEquipmentData([b1],BUS_nTapBus)[0]
    inSer1 = OlxAPILib.getEquipmentData([br1],BR_nInService)[0]
    sLocal = "Local Relay Group: " + OlxAPILib.fullBranchName(br1)
    if nTap1>0:
        s = '\n'+ sLocal +"\n\tis located at a TAP bus. Unable to continue."
        return AppUtils.FinishCheck(PY_FILE,s,None)
    #
    if inSer1!=1:
        s = '\n'+sLocal +"\n\tis located on out-of-service branch. Unable to continue."
        return AppUtils.FinishCheck(PY_FILE,s,None)
    #
    kr = 0
    sres = ""
    bres = OlxAPILib.getOppositeBranch(hndBr=br1,typeConsi=[]) # consider all type [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]
    #
    for bri in bres:
        rg1 = -1
        try:
            rg1 = OlxAPILib.getEquipmentData([bri],BR_nRlyGrp1Hnd)[0]
        except:
            pass
        #
        if rg1>0:
            kr +=1
            sres +="\n\t" + OlxAPILib.fullBranchName(bri)
    #
    s0 = 'App: ' + PY_FILE
    s0 +='\nUser: ' + os.getlogin()
    s0 +='\nDate: ' + time.asctime()
    s0 +='\nOLR file: ' + ARGVS.fi
    s0+= '\n\nSelected Relay Group:' + str(ARGVS.pk[0])

    if kr>1:
        s0 +="\nRemote groups (" + str(kr) + "):"
    else:
        s0 +="\nRemote group (" + str(kr) + "):"
    #
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='res', ad='_res' , ext='.txt')
    #
    AppUtils.saveString2File(ARGVS.fo,s0+sres)
    print("\nReport file had been saved in:\n" +ARGVS.fo )
    if ARGVS.ut==0 :
        AppUtils.launch_notepadpp(ARGVS.fo)
    return 1
#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.pk = ["[RELAYGROUP] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"]
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.txt')
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Relay Group: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
    return 0
#
def unit_test():
    ARGVS.fi = os.path.join (PATH_FILE, "LINETERM.OLR")
    ARGVS.pk = ["[RELAYGROUP] 5 'FIELDALE' 132 kV-'BUS3' 132 kV 1 L"]
    sres ='OLR file: ' + os.path.basename(ARGVS.fi) +'\n'
    run()
    sres += AppUtils.read_File_text_2(ARGVS.fo,4)
    ARGVS.pk = ["[RELAYGROUP] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"]
    run()
    sres +='\n'
    sres += AppUtils.read_File_text_2(ARGVS.fo,4)
    AppUtils.deleteFile(ARGVS.fo)
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)

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








