"""
Search all Bus Equipment downstream from a RLYGroup of high voltage winding of Transformer 2/3 or PhaseShifter
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2023, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__version__   = "1.0.1"

#
import sys,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
olxpathpy = os.path.split(PATH_FILE)[0]
olxpath = olxpathpy+'\\1lpfV15'
# INPUTS cmdline ---------------------------------------------------------------
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "Search all Bus Equipment downstream from a RLYGroup of high voltage winding of Transformer 2/3 or PhaseShifter"
PARSER_INPUTS.add_argument('-fi'   ,help = '*(str) OLR file path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-pk'   ,help = '*(str) Selected Bus in the 1-line diagram', default = [],nargs='+',metavar='')
PARSER_INPUTS.add_argument('-olxpath' , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
#
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
try:
    import OlxObj
except:
    raise Exception('\nPlease check in folder olxpathpy:"%s"\n\tnot found OlxAPI Python wrapper OlxAPI.py and relevant libraries'%ARGVS.pyolxpath)
#
def search1Bus(bus1,kv1):
    ba = []
    for t1 in bus1.TERMINAL:
        for b1 in t1.BUS[1:]:
            if b1.KV<kv1:
                ba.append(b1)
    return ba
#
def search(fi,olxpath,pk):
    OlxObj.OLCase.open(fi,1,olxpath=olxpath)
    o1 = OlxObj.OLCase.findOBJ(pk)
    try:
        ba = o1.BUS
        e1 = o1.EQUIPMENT
    except:
        e1 = None
    #
    if type(e1).__name__ not in {'XFMR','XFMR3','SHIFTER'}:
        raise Exception('-pk must start on RLYGROUP at high voltage side of XFMR, XFMR3 or SHIFTER')
    #
    kv1 = ba[0].KV
    if ba[1].KV>kv1 or (len(ba)==3 and ba[2].KV>kv1):
        raise Exception('-pk must start on RLYGROUP at high voltage side of XFMR, XFMR3 or SHIFTER')
    #
    print('\nStart on ',o1.toString())
    print('      of ',e1.toString())
    #
    busRes = ba[1:]
    setAlready = set() # bus handle search already
    setRes = set([b1.HANDLE for b1 in busRes]) # set handle of bus result
    #
    for ii in range(501): # max 500
        if ii==500:
            raise Exception('Error max iteration=500')
        #
        n1 = len(busRes)
        for i in range(n1):
            h1 = busRes[i].HANDLE
            if  h1 not in setAlready:
                ba1 = search1Bus(busRes[i],kv1)
                setAlready.add(h1)
                for b1 in ba1:
                    if b1.HANDLE not in setRes:
                        busRes.append(b1)
                        setRes.add(b1.HANDLE)
        if n1==len(busRes): # not found new bus
            break
    #
    return busRes
#
def printResult(busRes):
    print('\nBUS Result(s):')
    for b1 in busRes:
        print('\t',b1.toString())
    #
    print('LOAD Result(s):')
    for b1 in busRes:
        l1 = b1.LOAD
        if l1:
            print('\t',l1.toString())
    #
    print('GENERATOR Result(s):')
    for b1 in busRes:
        for o1 in ['GEN','GENW3','GENW4','CCGEN']:
            g1 = b1.getData(o1)
            if g1!=None:
                print('\t',g1.toString())

if __name__ == '__main__':
    ARGVS.fi = 'SAMPLE30.OLR'
    ARGVS.pk = "[RELAYGROUP] 14 'MONTANA' 33 kV-'MT MV BUS A' 12.47 kV 1 T"
    #
    busRes = search(ARGVS.fi,ARGVS.olxpath,ARGVS.pk)
    #
    printResult(busRes)

