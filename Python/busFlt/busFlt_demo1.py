"""
Purpose: Simulate 3PH and 1LN fault on the selected bus.
         Do simulation with single and double branch outages
         Record maximum fault current

         Output to a file:
             - Fault current
             - Max fault current and location
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "Yes"
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
import OlxAPILib
from OlxAPIConst import *
import AppUtils

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage=
         "\n\tSimulate 3PH and 1LN fault on the selected bus.\
         \n\tDo simulation with single and double branch outages\
         \n\tRecord maximum fault current\
         \n\tOutput to a file:\
         \n\t\t- Fault current\
         \n\t\t- Max fault current and location")
#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-pk' , help = '*(str) Selected Bus in the 1-Liner diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name of the report file',default = '',type=str,metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

# CORE--------------------------------------------------------------------------
def do1busFlt(bhnd):
    sres = ""
    # all branch connected to bhnd
    bra = OlxAPILib.getBusEquipmentData([bhnd],TC_BRANCH)[0]
    # doFault
    fltConn   = [1,0,1,0] # 3LG, 1LG
    outageOpt = [1,1,0,0] # one at a time, two at a time
    fltOpt    = [1,1]     # Bus or Close-in, Bus or Close-in w/ outage
    outageLst = []
    for br1 in bra:
        outageLst.append(br1)
    #
    fltR,fltX, clearPrev = 0,0,1
    OlxAPILib.doFault(bhnd,fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev)
    #
    OlxAPILib.pick1stFault()
    NoFaults = 0
    imax = 0
    fdesmax = ""
    while True:
        NoFaults +=1
        fdes = OlxAPILib.faultDescription()
        sres += "\n\n" + fdes
        sres += "\n                                  "
        mag,ang = OlxAPILib.getSCCurrent(hnd=HND_SC,style=4)
        for i in range(3):
            sres += "," + (str(round(mag[i],0)) +"@"  +  str(round(ang[i],1))).ljust(15)
            if mag[i]>imax:
                imax = mag[i]
                fdesmax = fdes
        #
        if not OlxAPILib.pickNextFault():
            break
    #
    sres += "\n\n" + str(NoFaults) + " Faults Simulated\n"
    sres += "\nMax fault current = " + str(round(imax,0)) + "A found at:"
    sres += "\n" +fdesmax
    #
    return sres

#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Bus',PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    # get handle of object picked up
    print("Selected BUS: "+ str(ARGVS.pk))

    bhnd = OlxAPILib.FindObj1LPF_check(ARGVS.pk[0],TC_BUS)
    #
    # sr = "\nOLR file: " + os.path.abspath(ARGVS.fi)
    sr ='\n'+ OlxAPILib.fullBusName(bhnd).ljust(25) +' : BUS FOUND'
    print(sr)
    sr += "\n\nBUS FAULT REPORT"
    sr += "\n                                ,Phase A        ,Phase B        ,Phase C"
    sr += do1busFlt(bhnd)
    #update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='res' , ad='' , ext='.csv')
    #
    AppUtils.saveString2File(ARGVS.fo,sr)
    #
    print('\nReport file had been saved in:\n%s'%ARGVS.fo)
    if ARGVS.ut ==0:
        AppUtils.launch_excel(ARGVS.fo)
    return 1
#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.csv')
        ARGVS.pk = ["[BUS] 6 'NEVADA' 132 kV"]
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Bus: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def unit_test():
    ARGVS.fi = os.path.join(PATH_FILE,"BusFlt.OLR")
    ARGVS.pk = ["[BUS] 6 'NEVADA' 132 kV"]
    sres = "OLR file: " + os.path.basename(ARGVS.fi) +"\n"
    run()
    sres += AppUtils.read_File_text_2(ARGVS.fo,1)
    #
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




