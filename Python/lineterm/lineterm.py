"""
Purpose: List the relay group(s) on the remoted terminal(s) of the selected line or transformer

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Released"
__version__   = "2.1.1"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))


# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "Report List the Relay Group(s) on the remoted terminal(s)\
        \n\tof the selected RLYGROUP in 1-line diagram"
PARSER_INPUTS.add_argument('-fr' , help = ' (str) Path name report file .CSV',default = "", type=str, metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]


def run(o1=None):
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    if o1==None:
        va = OLCase.getOBJSelected() #(list) of Selected Object(s) in the 1-line diagram
        if len(va)==0 or len(va)!=1 or (type(va[0]) !=RLYGROUP):
            PARSER_INPUTS.print_help()
            se = PARSER_INPUTS.usage.replace('\t','')
            if len(va)==0:
                se += '\n\nERROR: No selected RLYGROUP in the 1-line diagram'
            else:
                se += '\n\nERROR: selected Object(s) Found:'
                for i in range (min(2,len(va))):
                    se+='\n\t'+va[i].toString()
            AppUtils.gui_error(PY_FILE, se)
            return
        o1 = va[0]


    res = OLCase.tapLineTool(o1)
    #
    mainLine = res['mainLine']
    localBus = res['localBus']
    remoteBus = res['remoteBus']
    localRLG = res['localRLG']
    remoteRLG = res['remoteRLG']

    sres = PY_FILE+'\nOLR file: ' + OLCase.olrFile
    sres+= '\n\nSelected Relay Group: ' + o1.toString()
    if len(remoteRLG)==0:
        sres+= '\nRemote groups : Not Found'
    else:
        sres+= '\nRemote groups(%i):'%(len(remoteRLG))
        for r1 in remoteRLG:
            sres+= '\n\t'+r1.toString()
    print(sres)

    #
    ARGVS.fr = AppUtils.get_file_out(fo=ARGVS.fr, fi=OLCase.olrFile , subf='' , ad='_Report_'+PY_FILE[:-3] , ext='.csv')
    AppUtils.saveString2File(ARGVS.fr,sres)
    s1 = '\nReport file had been saved in:\n%s'%ARGVS.fr
    AppUtils.explorerDir(ARGVS.fr, s1, PY_FILE) #open dir of fo


def run_demo():
    if ARGVS.demo==1:
        fi = PATH_FILE+'\\LINETERM.OLR'
        OLCase.open(fi, 1)
        o1 = OLCase.findOBJ("[RELAYGROUP] 2 'CLAYTOR' 132 kV-'BUS7' 132 kV 1 L")

        sMain = "\nOLR file: " + fi
        sMain+= "\n\nobject: " + o1.toString()
        sMain+= "\n\nDo you want to run this demo (Y/N)?"
        choice = AppUtils.gui_askquestion(PY_FILE+' Demo',sMain)
        if choice=="YES":
            return run(o1)
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
    return 0


def main():
    if ARGVS.demo>0:
        return run_demo()
    run()


if __name__ == '__main__':
    main()









