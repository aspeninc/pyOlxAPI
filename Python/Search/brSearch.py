"""
Purpose: Test branch search function

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
PARSER_INPUTS.usage = "\n\tTest branch search function\
                       \n\tbranch defined by (bus1, bus2, id)\
                       \n\t             or by (name branch)"
PARSER_INPUTS.add_argument('-na1',metavar='', help = 'bus1 Name'    ,default = "" , type=str)
PARSER_INPUTS.add_argument('-kv1',metavar='', help = 'bus1 KV'      ,default = 0  , type=float)
PARSER_INPUTS.add_argument('-na2',metavar='', help = 'bus2 Name'    ,default = "" , type=str)
PARSER_INPUTS.add_argument('-kv2',metavar='', help = 'bus2 KV'      ,default = 0  , type=float)
#
PARSER_INPUTS.add_argument('-nu1',metavar='', help = 'bus1 Number'  ,default = 0  , type=int)
PARSER_INPUTS.add_argument('-nu2',metavar='', help = 'bus2 Number'  ,default = 0  , type=int)
PARSER_INPUTS.add_argument('-cid',metavar='', help = 'circuit ID'   ,default = "" , type=str)
#
PARSER_INPUTS.add_argument('-nbr',metavar='', help = 'name Branch'  ,default = "" , type=str)
#
PARSER_INPUTS.add_argument('-gui',metavar='', help = '0/1 GUI option (default=1)' ,default = 1 , type=int)
ARGVS = PARSER_INPUTS.parse_args()


def unit_test():
    sres = "\nUNIT TEST: " + PY_FILE +"\n"
    ARGVS.fi = "SEARCH.OLR"
    ARGVS.gui = 0
    #
    ARGVS.na1 , ARGVS.kv1 = "glen lyn",132
    ARGVS.na2 , ARGVS.kv2 = "claytor",132
    sres += run()
    #
    ARGVS.cid = "1"
    sres += run()
    #
    ARGVS.na1, ARGVS.kv1, ARGVS.na2 ,ARGVS.kv2 = "",0,"",0
    ARGVS.nu1, ARGVS.nu2 = 12,15
    ARGVS.cid = ""
    sres += run()
    #
    ARGVS.cid = "2"
    sres += run()
    #
    ARGVS.nbr = "/Nev"
    sres += run()
    #
    ARGVS.nbr = "Nev/Araz"
    sres += run()
    #
    print(sres)
    OlxAPILib.unit_test_compare(PATH_FILE,PY_FILE,sres)
#
def run():
    sres  = "\nRun : " + PY_FILE
    sres += "\nOLR file : " + os.path.basename(ARGVS.fi) +"\n"
    #
    OlxAPILib.open_olrFile(ARGVS.fi,readonly=1)
    bs = OlxAPILib.BranchSearch(gui=ARGVS.gui)
    #
    bs.setBusNum1(ARGVS.nu1)
    bs.setBusNameKV1(ARGVS.na1 , ARGVS.kv1)
    bs.setBusNum2(ARGVS.nu2)
    bs.setBusNameKV2(ARGVS.na2 , ARGVS.kv2)
    bs.setCktID(ARGVS.cid)
    bs.setNameBranch(ARGVS.nbr)
    #
    bhnd = bs.runSearch()
    #
    sres += bs.string_result()
    return sres
#
if __name__ == '__main__':
    if ARGVS.ut ==1:
        unit_test()
    else:
        sr = run()
        print(sr)



