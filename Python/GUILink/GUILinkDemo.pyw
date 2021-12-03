"""
Demo OneLiner to Python OlxAPI app GUI link
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

#
import sys,os
#
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
PARSER_INPUTS = AppUtils.iniInput(usage=\
"Demo OneLiner GUI - OlxAPI Python app link.\
 \n\tHighlight a line connected to the bus that user had selected.")
PARSER_INPUTS.add_argument('-fi'   ,help = '*(str) OLR file path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-tpath',help = '*(str) Output folder for PowerScript command file', default = '',type=str,metavar='' )
PARSER_INPUTS.add_argument('-pk'   ,help = '*(str) Selected Bus in the 1-line diagram', default = [],nargs='+',metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS)

#
def psCallBack_show_obj(hnd):
    """
    Create a callback PowerSript program in -tpath
    that will highlight the hnd object on OneLiner 1-line diagram
    """
    if hnd<=0:
        raise OlxAPI.OlxAPIException("Invalid object handle.")
    # text description of Object (line)
    sobj = OlxAPI.PrintObj1LPF(hnd)
    if sobj.find("[") < 0:
        raise OlxAPI.OlxAPIException("Invalid object handle.")

    # PowerScript program to highlight an object on the 1-line digram
    s = 'Sub main\n'
    s+= '\tobjStr$=' + '"'+sobj+'"\n'
    s+= '\tCall FindObj1LPF(objStr$, hnd&)\n'
    s+= '\tCall Locate1LObj(hnd)\n'
    s+= '\tPrint "Found " & printobj1lpf(hnd)\n'
    s+= 'End Sub'
    #
    if ARGVS.tpath=='':              # When debugging, run GUILinkDemo in the Python IDE folder
        ARGVS.tpath = get_tpath()
    #
    sfile = os.path.join(ARGVS.tpath,'pst.bas')
    AppUtils.saveString2File(sfile,s)
    print("Callback PowerScript file saved as: "+sfile)
#
def findFirstLine():
    """
    Return the handle of the 1st line connected to the selected bus on the 1-line diagram
    specified in ARGVS.pk[0] string
    """
    # find handle of BUS
    bhnd = OlxAPILib.FindObj1LPF_check(ARGVS.pk[0],TC_BUS)
    print("Selected Bus: "+ OlxAPI.FullBusName(bhnd))
    # Branch handles at selected BUS
    br = OlxAPILib.getBusEquipmentData([bhnd],TC_BRANCH)[0]
    # Branch equipment Handles
    equi = OlxAPILib.getEquipmentData(br,BR_nHandle)
    # Equipment type
    typ1  = OlxAPILib.getEquipmentData(br,BR_nType)
    for i in range(len(equi)):
        if typ1[i]==TC_LINE: # Found the 1st line
            print('Line found: '+OlxAPILib.fullBranchName_1(equi[i]))
            return equi[i]
    raise OlxAPI.OlxAPIException("There is no line at the selected Bus.")
    return 0

def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    if not AppUtils.checkInputPK(ARGVS.pk,'Bus',PY_FILE,PARSER_INPUTS):
        return  
    #
    # Open the copy of the OLR file in the temp folder, which
    # matches exactly what is currently in the GUI.
    olrFile = os.path.basename(ARGVS.fi)
    olrFileTemp = os.path.join(ARGVS.tpath,olrFile)
    if not os.path.isfile(olrFileTemp):
        olrFileTemp = ARGVS.fi
    #
    OlxAPILib.open_olrFile(olrFileTemp,ARGVS.olxpath)     

    # search for a line at the selected bus
    lhnd = findFirstLine()

    # Create the PowerScript callback 'pst.bas' in -tpath that will
    # highlight the like on the 1-line diagram
    if lhnd > 0:
        psCallBack_show_obj(lhnd)
#
if __name__ == '__main__':
    try:
        run()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)



