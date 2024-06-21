"""
Purpose:  Check the phase shift across all 2-winding transformers with wye-delta connection.
          When a transformer with high side lagging the low side is found,
          this script modifies the delta side to make the high side leads.
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "2.1.0"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "\n\tCheck phase shift of all 2-winding transformers with wye-delta connection\
                        \n\tWhen a transformer with high side lagging the low side is found,\
                        \n\tthis function converts it to make the high side lead"
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
PARSER_INPUTS.add_argument('-fo'  , help = ' (str) OLR corrected file path',default = '',type=str,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]

# CORE--------------------------------------------------------------------------
def run():
    """
      Check phase shift of all 2-winding transformers with wye-delta connection.
          When a transformer with high side lagging the low side is found,
          this function converts it to make the high side lead.

        Returns:
            string result
            OLR corrected file
    """
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    #
    x2a = []
    for x2 in OLCase.XFMR: #loop all XFMR
        configP = x2.CONFIGP
        configS = x2.CONFIGS
        tapP = x2.PRITAP
        tapS = x2.SECTAP

        # low - hight  =>  low - hight
        #   G - D            G - E
        if (tapP<tapS) and configP=='G' and configS=='D':
            x2.CONFIGS = 'E' # set data
            x2.postData()    # post data
            x2a.append(x2)

        # hight-low   =>  hight-low
        #    G - E           G - D
        if (tapP>tapS) and configP=='G' and configS=='E':
            x2.CONFIGS = 'D' # set data
            x2.postData()    # post data
            x2a.append(x2)
    #
    if len(x2a)==0:
        s1 = "All wye-delta transformers in this OLR file have the correct phase shift of the high-side leading the low-side."
        AppUtils.gui_info(PY_FILE, s1)
        return s1

    s1 = "Fixed (%i) wye-delta transformers to make the high side lead the low side:\n"%len(x2a)
    for i in range(len(x2a)):
        s1 += '\n'+str(i+1).ljust(5)  + x2a[i].toString()
    AppUtils.gui_info(PY_FILE, s1)

    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=OLCase.olrFile , subf='' , ad='_res' , ext='.OLR')
    OLCase.save(ARGVS.fo)
    return s1

#
def run_demo():
    if ARGVS.demo==1:
        fi = PATH_FILE+'\\XFMRConn.OLR'
        OLCase.open(fi,1)
        #
        sMain = "\nOLR file: " + fi
        sMain+= "\n\nDo you want to run this demo (Y/N)?"
        choice = AppUtils.gui_askquestion(PY_FILE+' Demo',sMain)
        if choice=='YES':
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE, ARGVS.demo, [1],gui_err=1)


def main():
    if ARGVS.demo!=0:
        return run_demo()
    run()


if __name__ == '__main__':
    main()






