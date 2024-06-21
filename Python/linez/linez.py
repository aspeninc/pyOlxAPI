"""
Purpose: Report total line impedance and length
         All taps are ignored. Close switches,Serie capacitor/reactor bypass are included
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "2.1.1"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "Report total line impedance and length\
        \n\tFrom a selected object (LINE/RLYGROUP/SWITCH/SERIESRC) in 1-line diagram\
        \n\tAll taps are ignored. Close switches, Serie capacitor/reactor bypass are included"
PARSER_INPUTS.add_argument('-fr' , help = ' (str) Path name report file .CSV',default = "", type=str, metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]


def run(o1=None):
    """
    run linez:
    """
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    if o1==None:
        va = OLCase.getOBJSelected() #(list) of Selected Object(s) in the 1-line diagram
        if len(va)==0 or len(va)!=1 or (type(va[0]) not in {LINE,RLYGROUP,SWITCH,SERIESRC}):
            PARSER_INPUTS.print_help()
            se = PARSER_INPUTS.usage.replace('\t','')
            if len(va)==0:
                se += '\n\nERROR: No selected Object in the 1-line diagram'
            else:
                se += '\n\nERROR: selected Objects Found:'
                for i in range (min(2,len(va))):
                    se+='\n\t'+va[i].toString()
            AppUtils.gui_error(PY_FILE, se)
            return
        o1 = va[0]

    #
    sres = PY_FILE+'\nOLR file: ' + OLCase.olrFile
    sres+= '\nSelected Object: ' + o1.toString()
    sres += run1t(o1)

    #
    ARGVS.fr = AppUtils.get_file_out(fo=ARGVS.fr, fi=OLCase.olrFile , subf='' , ad='_Report_'+PY_FILE[:-3] , ext='.csv')
    AppUtils.saveString2File(ARGVS.fr,sres)
    s1 = '\nReport file had been saved in:\n%s'%ARGVS.fr
    AppUtils.explorerDir(ARGVS.fr, s1, PY_FILE) #open dir of fo


def run1t(t0):
    res = OLCase.tapLineTool(t0)
    #
    mainLine = res['mainLine']
    localBus = res['localBus']
    remoteBus = res['remoteBus']
    localRLG = res['localRLG']
    remoteRLG = res['remoteRLG']
    Z1 = res['Z1']
    Z0 = res['Z0']
    Length = res['Length']
    #
    sres =''
    for i in range(len(mainLine)):
        sres +='\nmainLine '
        if len(mainLine)>1:
            sres += str(i+1)
        sres +=': '+toString(localBus[i]) +'-'+ toString(remoteBus[i])
        sres +='\n\tZ1[pu]   : '+toString(Z1[i])
        sres +='\n\tZ0[pu]   : '+toString(Z0[i])
        sres +='\n\tLength   : '+toString(Length[i])
        sres +='\n\tlocalRLG : '+toString(localRLG[i])
        sres +='\n\tremoteRLG: '+toString(remoteRLG[i])
        sres +=('\n\tLine section(s):')
        for v1 in mainLine[i]:
            sres +='\n\t\t'+v1.toString()
        sres+='\n'
    return sres


#
def run_demo():
    if ARGVS.demo==1:
        fi = PATH_FILE+'\\LINEZ.OLR'
        OLCase.open(fi, 1)
        o1 = OLCase.findOBJ("[RELAYGROUP] 2 'CLAYTOR' 132 kV-'BUS7' 132 kV 1 L")

        sMain = "\nOLR file: " + fi
        sMain+= "\n\nobject: " + o1.toString()
        sMain+= "\n\nDo you want to run this demo (Y/N)?"
        choice = AppUtils.gui_askquestion(PY_FILE+' Demo',sMain)
        if choice=='YES':
            return run(o1)
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])


def main():
    if ARGVS.demo>0:
        return run_demo()
    return run()


if __name__ == '__main__':
    main()





