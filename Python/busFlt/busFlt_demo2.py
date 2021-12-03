"""
Purpose: Simulate faults on the selected bus.
         Record fault current, Thevenin impedance and X/R
         in a CSV file
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

#
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
PARSER_INPUTS = AppUtils.iniInput(usage= "\n\tSimulate faults on the selected bus.\
         \n\tRecord fault current, Thevenin impedance and X/R\
         \n\tin a CSV file")
#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',metavar='')
PARSER_INPUTS.add_argument('-pk' , help = '*(str) Selected Bus in the 1-Liner diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name of the report file ',default = '',metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)
#
def do1busFlt1(bhnd):
    #
    fltConn   = [1,1,1,1]
    #
    fltOpt    = [0]*15
    fltOpt[0] = 1  # Bus or Close-in
    OlxAPILib.doFault(bhnd,fltConn, fltOpt, outageOpt=[], outageLst=[], fltR=0, fltX=0, clearPrev=1)
    OlxAPILib.pick1stFault()
    #
    nRound = 2
    sres = ''
    while True:
        sres += "\n"+ OlxAPILib.faultDescription()
        #
        mag,ang = OlxAPILib.getSCCurrent(hnd=HND_SC,style=4)
        #
        for i in range(3):
            sres += "," + str(round(mag[i],nRound)) +"@"  +  str(round(ang[i],nRound))
        #
        r0 =  OlxAPILib.getEquipmentData([HND_SC],FT_dRZt)[0]
        r1 =  OlxAPILib.getEquipmentData([HND_SC],FT_dRPt)[0]
        r2 =  OlxAPILib.getEquipmentData([HND_SC],FT_dRNt)[0]
        #
        x0 =  OlxAPILib.getEquipmentData([HND_SC],FT_dXZt)[0]
        x1 =  OlxAPILib.getEquipmentData([HND_SC],FT_dXPt)[0]
        x2 =  OlxAPILib.getEquipmentData([HND_SC],FT_dXNt)[0]
        #
        xr = OlxAPILib.getEquipmentData([HND_SC],FT_dXR)[0]
        #  R0+jX0, R1+jX1, R2+jX2, X/R
        sres += "," + str(round(r0,nRound)) +"+j" + str(round(x0,nRound))
        sres += "," + str(round(r1,nRound)) +"+j" + str(round(x1,nRound))
        sres += "," + str(round(r2,nRound)) +"+j" + str(round(x2,nRound))
        sres += "," + str(round(xr,nRound))
        if not OlxAPILib.pickNextFault():
            break
    #
    return sres
#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Bus',PY_FILE,PARSER_INPUTS):
        return
    #
    sr = ''#"\nOLR file: " + os.path.abspath(ARGVS.fi)
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    print("Selected BUS: "+ str(ARGVS.pk))
    for i in range(len(ARGVS.pk)):
        # get handle of object picked up
        bhnd = OlxAPILib.FindObj1LPF_check(ARGVS.pk[i],TC_BUS)
        sb = OlxAPILib.fullBusName(bhnd).ljust(25) +' : BUS FOUND'
        sr+= "\n"+sb
        print(sb)
        #
        sr += "\nFault, PhaseA, PhaseB, PhaseC, R0+jX0, R1+jX1, R2+jX2, X/R"
        #
        sr+= do1busFlt1(bhnd) + "\n"
    #update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='res' , ad='' , ext='.csv')
    AppUtils.saveString2File(ARGVS.fo,sr)
    print('\nReport file had been saved in:\n%s'%ARGVS.fo)
    if ARGVS.ut ==0:
        AppUtils.launch_excel(ARGVS.fo)
    return 1
#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.csv')
        ARGVS.pk = ["[BUS] 6 'NEVADA' 132 kV", "[BUS] 14 'MONTANA' 33 kV"]
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Selected Bus: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])

def unit_test():
    ARGVS.fi = os.path.join(PATH_FILE,"BusFlt.OLR")
    ARGVS.pk = ["[BUS] 6 'NEVADA' 132 kV", "[BUS] 14 'MONTANA' 33 kV"]
    sr = "OLR file: " + os.path.basename(ARGVS.fi) +"\n"
    run()
    sr += AppUtils.read_File_text_2(ARGVS.fo,1)
    #
    AppUtils.deleteFile(ARGVS.fo)
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sr)

#
def main():
    if ARGVS.ut >0:
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


