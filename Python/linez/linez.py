"""
Purpose: Report total line impedance and length
         All taps are ignored. Close switches,Serie capacitor/reactor bypass are included
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.3.1"

import os,sys,time,logging
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\OlxAPI\python'
#olxpathpy = os.path.split(PATH_FILE)[0]
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "\n\tReport total line impedance and length\
        \n\tFrom a Selected object (GUID) BUS/LINE/RLYGROUP/SWITCH/SERIESRC\
        \n\tAll taps are ignored. Close switches, Serie capacitor/reactor bypass are included"
PARSER_INPUTS.add_argument('-fi'        , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-pk'        , help = '*(str) Selected/GUID BUS/LINE/RLYGROUP/SWITCH/SERIESRC in the 1-line diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-olxpath'   , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
PARSER_INPUTS.add_argument('-demo'      , help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
#
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,olxpathpy)
import OlxObj
import AppUtils
#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = ARGVS.olxpath +'\\SAMPLE30.OLR'
        ARGVS.pk = "[BUS] 6 'NEVADA' 132 kV"
        #
        choice = AppUtils.ask_run_demo(PY_FILE,0,ARGVS.fi,"Selected BUS: " + str(ARGVS.pk))
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def run():
    """
    run linez:
    """
    flog = AppUtils.logger2File(PY_FILE,version = __version__)
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return

    if not AppUtils.checkInputPK(ARGVS.pk,'BUS/LINE/RLYGROUP',PY_FILE,PARSER_INPUTS):
        return
    #
    OlxObj.load_olxapi(ARGVS.olxpath) # OlxAPI module initialization
    OlxObj.OLCase.open(ARGVS.fi,1)
    #
    logging.info('\nOLR file: ' + ARGVS.fi)
    if type(ARGVS.pk)!= list:
        ARGVS.pk = [ARGVS.pk]
    #
    for i in range(len(ARGVS.pk)):
        logging.info('\nSelected Object: ' + str(ARGVS.pk[i]))
        t0 = OlxObj.OLCase.findOBJ(ARGVS.pk[i])
        run1t(t0)
    print('logFile:',flog)
    return 1
#
def run1t(t0):
    res = OlxObj.OLCase.tapLineTool(t0)
    #
    mainLine = res['mainLine']
    remoteBus = res['remoteBus']
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
        sres +=': '+OlxObj.toString(mainLine[i][0].BUS1) +'-'+ OlxObj.toString(remoteBus[i])
        sres +='\n\tZ1[pu]   : '+OlxObj.toString(Z1[i])
        sres +='\n\tZ0[pu]   : '+OlxObj.toString(Z0[i])
        sres +='\n\tLength   : '+OlxObj.toString(Length[i])
        sres +='\n\tremoteRLG: '+OlxObj.toString(remoteRLG[i])
        sres +=('\n\tLine section(s):')
        for v1 in mainLine[i]:
            sres +='\n\t\t'+v1.toString()[11:]
        sres+='\n'
    logging.info(sres)
#
def main():
    if ARGVS.demo>0:
        return run_demo()
    return run()
#
if __name__ == '__main__':
##    ARGVS.fi = os.path.join (PATH_FILE, "LINEZ.OLR")
##    ARGVS.pk = "[SWITCH] 8 'REUSENS' 132 kV-'BUS8' 132 kV 1t"
##    ARGVS.demo = 1
    try:
        main()
    except Exception as err:
        #raise Exception(err)
        AppUtils.FinishException(PY_FILE,err)




