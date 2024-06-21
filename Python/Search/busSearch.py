"""
Purpose: Test bus search function

"""
from __future__ import print_function
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "0.1.1"
__email__     = "support@aspeninc.com"
__status__    = "In development"

# IMPORT -----------------------------------------------------------------------
import sys,os,time
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
sys.path.insert(0, PATH_LIB)

#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
#
PARSER_INPUTS.usage = "\n\tTest bus search function\
                      \n\tbus defined by (name,KV)\
                      \n\t         or by (bus Number)"
PARSER_INPUTS.add_argument('-na', metavar='', help = 'bus Name'    ,default = "" , type=str)
PARSER_INPUTS.add_argument('-kv', metavar='', help = 'bus KV'      ,default = 0  , type=float)
PARSER_INPUTS.add_argument('-nu', metavar='', help = 'bus Number'  ,default = -1 , type=int)
PARSER_INPUTS.add_argument('-gui',metavar='', help = '0/1 GUI option (default=1)',default = 1, type=int)
ARGVS = PARSER_INPUTS.parse_args()


def unit_test():
    sres = "\nUNIT TEST: " + PY_FILE +"\n"
    ARGVS.fi = "SEARCH.OLR"
    ARGVS.gui = 0
    #
    ARGVS.na , ARGVS.kv = "al",32
    sres +=run()
    #
    ARGVS.na , ARGVS.kv = "alz",32
    sres +=run()
    #
    ARGVS.na , ARGVS.kv = "etc",100
    sres +=run()
    #
    ARGVS.na , ARGVS.kv = "nhs",33
    sres +=run()
    #
    ARGVS.na , ARGVS.kv = "nhs",132
    sres +=run()
    #
    ARGVS.na , ARGVS.kv = "jj",0
    sres +=run()
    #
    ARGVS.nu, ARGVS.na , ARGVS.kv = 25,"",0
    sres +=run()
    #
    ARGVS.nu, ARGVS.na , ARGVS.kv = 250,"",0
    sres +=run()
    print(sres)
    #
    OlxAPILib.unit_test_compare(PATH_FILE,PY_FILE,sres)
#
def run():
    sres  = "\nRun : " + PY_FILE
    sres += "\nOLR file : " + os.path.basename(ARGVS.fi) +"\n"

    OlxAPILib.open_olrFile(ARGVS.fi,readonly=1)
    bs = OlxAPILib.BusSearch(gui=ARGVS.gui)
    #
    bs.setBusNumber(ARGVS.nu)
    bs.setBusNameKV(ARGVS.na , ARGVS.kv)
    #
    bhnd = bs.runSearch()
    #
    sres += bs.stringResult()
    return sres

if __name__ == '__main__':
    if ARGVS.ut ==1:
        unit_test()
    else:
        sr =run()
        print(sr)



