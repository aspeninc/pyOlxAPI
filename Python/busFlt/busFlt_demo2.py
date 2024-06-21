"""
Purpose: Simulate faults on the selected bus.
         Record fault current, Thevenin impedance and X/R
         in a CSV file
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
import os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "\n\tSimulate faults on the selected bus.\
         \n\tRecord fault current, Thevenin impedance and X/R\
         \n\tin a CSV file"
#
PARSER_INPUTS.add_argument('-fo'  , help = ' (str) Path name of the report file ', default = '', metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]


# CORE--------------------------------------------------------------------------
def do1busFlt_2(b1):
    sres  = '\n\n' + b1.toString()
    sres += '\n Fault                                                           '
    sres += ',PhaseA          ,PhaseB          ,PhaseC          ,R0+jX0        ,R1+jX1        ,R2+jX2        ,X/R      ,ANSI X/R ,R0/X1    ,X0/X1'

    # Define the fault specification
    fs1 = SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='3LG')
    OLCase.simulateFault(fs1, 1)

    fs1.fltConn = '2LG:BC'
    OLCase.simulateFault(fs1, 0)

    fs1.fltConn = '1LG:A'
    OLCase.simulateFault(fs1, 0)

    fs1.fltConn = 'LL:BC'
    OLCase.simulateFault(fs1, 0)

    for r1 in FltSimResult:
        sres += '\n' + r1.FaultDescription.replace('  ',' ').ljust(65)

        c1 = r1.current()
        for ii in range(3):
            v1 = ',' + toString(c1[ii],1, mode='polar').ljust(16)
            sres += v1.ljust(15)

        #
        tv1 = r1.Thevenin
        for ii in range(3):
            v1 = ",%.2f+j%.2f"%(tv1[ii].real, tv1[ii].imag)
            sres += v1.ljust(15)

        #
        xr1 = r1.XR_RATIO
        for ii in range(4):
            try:
                v1 = ',%.2f'%xr1[ii]
            except:
                v1 = ',None'
            sres += v1.ljust(10)

    return sres

#
def run(b01=None):
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    if b01==None:
        ba = OLCase.busPicker()
        if len(ba)==0:
            print('No Bus Selected')
            return
    else:
        ba = [b01] if type(b01)!=list else b01[:]

    sr = "OLR file: " + OLCase.olrFile
    for b1 in ba:
        sr += do1busFlt_2(b1)

    #update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo, fi=OLCase.olrFile , subf='' , ad='_'+PY_FILE[:-3] , ext='.csv')
    AppUtils.saveString2File(ARGVS.fo,sr)

    sMain = 'Report file had been saved in:\n%s'%ARGVS.fo
    AppUtils.explorerDir(ARGVS.fo, sMain, PY_FILE) #open dir of fo

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



