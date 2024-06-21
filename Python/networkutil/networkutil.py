"""
Sample OlxAPI app: List all remote end of all branches of the selected bus
                    All taps are ignored. Close switches are included
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Released"
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
        "\n\tList all remote end of all branches of the selected bus\
         \n\tAll taps are ignored. Close switches are included")
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = "",metavar='')
PARSER_INPUTS.add_argument('-pk' , help = '*(str) Selected Bus in the 1-line diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name of the report file',default = "",metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

# CORE
def getRemoteTerminals(bhnd):
    """
    Purpose: Find all remote end of all branches (from a bus)
             All taps are ignored. Close switches are included
             =>a test function of OlxAPILib.getRemoteTerminals(hndBr,typeConsi)

        Args :
            bhnd :  bus handle

        returns :
            br_res [] : list of all branch from bus
            bus_res []: list of terminal bus

        Raises:
            OlrxAPIException
    """
    br_res = OlxAPILib.getBusEquipmentData([bhnd],TC_BRANCH)[0]
    bus_res = []
    #
    typeConsi = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3] # all type of branch
    for br1 in br_res:
        bra,bt,ba,equip = OlxAPILib.getRemoteTerminals(br1,typeConsi)
        bus_res.append(ba)
    return br_res,bus_res
#
def getRemoteTerminals_str(br_res,bus_res):
    sres = ""
    for i in range(len(br_res)):
        sres += "\n" + str(i+1).ljust(5)+ " Branch: "+OlxAPILib.fullBranchName(br_res[i])+"\n"
        if len(bus_res[i])==0:
            sres += "\t" + str(len(bus_res[i])) + " remote terminal (branch out of service)\n"
        else:
            sres += "\t" + str(len(bus_res[i])) + " remote terminals:\n"

        for j in range(len(bus_res[i])):
            if len(bus_res[i])>1:
                sres += "\t\t" + str(j+1) + " " + OlxAPILib.fullBusName(bus_res[i][j]) + "\n"
            else:
                sres += "\t\t" + OlxAPILib.fullBusName(bus_res[i][j]) + "\n"
    return sres
#
def unit_test():
    ARGVS.fi = os.path.join (PATH_FILE, "NETWORKUTIL.OLR")
    ARGVS.pk = ["[BUS] 10 'NEW HAMPSHR' 33 kV","[BUS] 28 'ARIZONA' 132 kV"]
    #
    run()
    sres = "OLR file: " + os.path.basename(ARGVS.fi) +"\n"
    sres += AppUtils.read_File_text_2(ARGVS.fo,4)
    #
    AppUtils.deleteFile(ARGVS.fo)
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)

#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.pk = ["[BUS] 6 'ARIZONA' 132 kV"]
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.txt')
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Bus: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])

#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Bus',PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    sres = 'App: ' + PY_FILE
    sres +='\nUser: ' + os.getlogin()
    sres +='\nDate: ' + time.asctime()
    sres +='\nOLR file: ' + ARGVS.fi
    print('Selected Bus: ',str(ARGVS.pk))
    for i in range(len(ARGVS.pk)):
        sres+= '\n\nSelected Bus:' + ARGVS.pk[i]
        # get handle of object picked up
        bhnd = OlxAPILib.FindObj1LPF_check(ARGVS.pk[i],TC_BUS)
        #
        br_res, bus_res = getRemoteTerminals(bhnd)
        sres += getRemoteTerminals_str(br_res,bus_res)
    #
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='', ad='_res' , ext='.txt')
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
    except Exception as er:
        AppUtils.FinishException(PY_FILE,er)




