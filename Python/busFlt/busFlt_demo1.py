"""
Purpose: Simulate 3PH and 1LN fault on the selected bus.
         Do simulation with single and double branch outages
         Record maximum fault current

         Output to a file:
             - Fault current
             - Max fault current and location
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "2.1.3"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os,cmath
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage =\
         "\n\tSimulate 3PH and 1LN fault on the selected bus.\
         \n\tDo simulation with single and double branch outages\
         \n\tRecord maximum fault current\
         \n\tOutput to a file:\
         \n\t\t- Fault current\
         \n\t\t- Max fault current and location"
#
PARSER_INPUTS.add_argument('-fo'  , help = ' (str) Path name of the report file',default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]


# CORE--------------------------------------------------------------------------
def do1busFlt_1(b1):
    sres = '\n\n' + b1.toString()
    # Define the fault specification
    fs1 = SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='3LG')
    OLCase.simulateFault(fs1, 1)
    fs1.fltConn = '1LG:A'
    OLCase.simulateFault(fs1, 0)

    o1 = OUTAGE(option='SINGLE')
    o1.add_outageLst(b1.TERMINAL)

    print('List Outage')
    for v1 in o1.outageLst:
        print('\t', v1.toString())

    fs1 = SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='3LG',outage=o1)


    fs1.fltConn = '3LG'
    fs1.outage.option = 'SINGLE'
    OLCase.simulateFault(fs1, 0)

    #
    fs1.fltConn = '3LG'
    fs1.outage.option = 'DOUBLE'
    OLCase.simulateFault(fs1, 0)

    #
    fs1.fltConn = '1LG:A'
    fs1.outage.option = 'SINGLE'
    OLCase.simulateFault(fs1, 0)

    #
    fs1.fltConn = '1LG:A'
    fs1.outage.option = 'DOUBLE'
    OLCase.simulateFault(fs1, 0)


    # RESULT
    imax, fdesmax = 0, ''
    for r1 in FltSimResult:
        sres += '\n\n' + r1.FaultDescription
        sres += '\n                                                                   '
        c1 = r1.current()
        for ii in range(3):
            mag = abs(c1[ii])
            sres += ','+toString(c1[ii],1, mode='polar').ljust(16)
            if mag>imax:
                imax = mag
                fdesmax = r1.FaultDescription
    sres += "\n\nMax fault current = %.0f[A]"%imax + ", found at:"
    sres += "\n" +fdesmax+'\n\n'
    return sres


def run(b01=None):
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    if b01==None:
        ba = OLCase.busPicker()
        if len(ba)==0:
            print('No Bus Selected')
            return
    else:
        ba = [b01] if type(b01)!=list else b01[:]

    sr = "\nOLR file: " + OLCase.olrFile
    sr += "\n\nBUS FAULT REPORT"
    sr += "\n                                                                    ,Phase A        ,Phase B        ,Phase C"

    for b1 in ba:
        sr += do1busFlt_1(b1)

    #update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo, fi=OLCase.olrFile , subf='' , ad='_'+PY_FILE[:-3] , ext='.csv')
    #
    AppUtils.saveString2File(ARGVS.fo,sr)
    #
    s1 = '\nReport file had been saved in:\n%s'%ARGVS.fo
    AppUtils.explorerDir(ARGVS.fo, s1, PY_FILE) #open dir of fo

#
def run_demo():
    if ARGVS.demo==1:
        fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        OLCase.open(fi,1)
        b1 = OLCase.findOBJ('BUS', ['NEVADA',132])
        #
        sMain = "\nOLR file: " + fi
        sMain+= "\n\nBus: " + b1.toString()
        sMain+= "\n\nDo you want to run this demo (Y/N)?"
        choice = AppUtils.gui_askquestion(PY_FILE+' Demo',sMain)
        if choice=='YES':
            return run(b1)
    else:
        AppUtils.demo_notFound(PY_FILE, ARGVS.demo, [1],gui_err=1)


def main():
    if ARGVS.demo!=0:
        return run_demo()
    run()


if __name__ == '__main__':
    main()



