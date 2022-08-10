"""
Purpose:  Check phase shift of all 2-winding transformers with wye-delta connection.
          When a transformer with high side lagging the low side is found,
          this function converts it to make the high side lead.
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.2"

import sys,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)

import OlxAPI
import OlxAPILib
from OlxAPIConst import *
import AppUtils

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage=
         "\n\tCheck phase shift of all 2-winding transformers with wye-delta connection\
                        \n\tWhen a transformer with high side lagging the low side is found,\
                        \n\tthis function converts it to make the high side lead")
#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR input file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) OLR corrected file path',default = '',type=str,metavar='')
#
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

# CORE--------------------------------------------------------------------------
def run():
    """
      Check phase shift of all 2-winding transformers with wye-delta connection.
          When a transformer with high side lagging the low side is found,
          this function converts it to make the high side lead.

        Args :
            None
        Returns:
            string result
            OLR corrected file

        Raises:
            OlrxAPIException
    """
    #
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile_1(ARGVS.olxpath,ARGVS.fi,readonly=1)
    s1 = ""
    # get all xfmr
    x2a = OlxAPILib.getEquipmentHandle(TC_XFMR)
    kd = 0
    for x1 in x2a:
        configA = OlxAPILib.getEquipmentData([x1],XR_sCfg1)[0]
        configB = OlxAPILib.getEquipmentData([x1],XR_sCfg2)[0]
        tapA =  OlxAPILib.getEquipmentData([x1],XR_dTap1)[0]
        tapB =  OlxAPILib.getEquipmentData([x1],XR_dTap2)[0]
        #
        test = False

        # small - hight  =>  small - hight
        #   G   - D             G  - E
        if (tapA<tapB) and configA=="G" and configB=="D":
            test = True
            # set
            OlxAPILib.setData(x1,XR_sCfg2,"E")
            # validation
            OlxAPILib.postData(x1)

        # hight - small  =>  hight - small
        #  G    - E              G - D
        if (tapA>tapB) and configA=="G" and configB=="E":
            test = True
            # set
            OlxAPILib.setData(x1,XR_sCfg2,"D")
            # validation
            OlxAPILib.postData(x1)
        #
        if test:
             kd += 1
             s1 += '\n'+str(kd).ljust(5)  + OlxAPILib.fullBranchName(x1)
    #
    if kd>0:
        s2 = "Fixed (" + str(kd) + ") wye-delta transformers such that the high side leads the low side:"
    else:
        s2 = "All wye-delta transformers in this OLR file have phase shift with high-side leading low-side."
    #
    print(s2+s1)
    #
    if kd>0:
        ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='' , ad='_res' , ext='.OLR')
        #
        OlxAPILib.saveAsOlr(ARGVS.fo)
        if ARGVS.ut==0:
            print("All corrects had been saved in:\n" + ARGVS.fo)
            AppUtils.launch_OneLiner(ARGVS.fo)
    #
    return s2+s1
#
def run_demo():
    if ARGVS.demo==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.OLR')
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,'')
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def unit_test():
    ARGVS.fi = os.path.join (PATH_FILE, "XFMRConn.OLR")
    sres = "OLR file: "+os.path.basename(ARGVS.fi) +"\n"
    sres+= run()
    #
    s1 = ARGVS.fi
    ARGVS.fi = ARGVS.fo
    sres +="\nRe-check\n"
    sres += "OLR file: "+os.path.basename(ARGVS.fi) +"\n"
    sres+= run()
    #
    OlxAPI.CloseDataFile()
    AppUtils.deleteFile(ARGVS.fi)
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





