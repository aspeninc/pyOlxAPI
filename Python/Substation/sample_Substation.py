"""
Purpose:
  Substation demo
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.1.0"

import os,sys
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
#
olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\OlxAPI\python'
olxpathpy = os.path.split(PATH_FILE)[0] #delete in RELEASE
# INPUTS cmdline ---------------------------------------------------------------
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'Substation demo'
PARSER_INPUTS.add_argument('-fi'  , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpath' , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
#
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
import OlxObj
import Substation
import AppUtils
OlxObj.load_olxapi(ARGVS.olxpath)

def sample_1(fi):
    """ Substation from a BUS """
    OlxObj.OLCase.open(fi,1)
    b1 = OlxObj.OLCase.findBUS('nevada',132)
    # 3 criterias to considered 1 AC Line is an Equipment of Substation: leng or xpu or xOhm
    sub1 = Substation.substation(b1,leng=0.1,xpu=1e-4,xOhm=0.5)
    print(sub1.toString())
#
def sample_2(fi):
    """ all Substation of NETWORK """
    OlxObj.OLCase.open(fi,1)
    suba = Substation.substation(OlxObj.OLCase,leng=0.1,xpu=1e-4,xOhm=0.5)
    for a1 in suba:
        print(a1.toString())
#
def sample_3(fi):
    """ export all Substation of NETWORK: each substation to 1 file OLR """
    OlxObj.OLCase.open(fi,1)
    folder = 'XXX'
    if not os.path.isdir(folder):
        os.mkdir(folder)
    ff = folder+'\\'+os.path.basename(ARGVS.fi)[:-4]
    fo1 = AppUtils.get_file_out(fo=ff, fi=fi , subf='' , ad='', ext='.TXT')
    ff +='_substation'
    f1 = open(fo1, "w", encoding="utf-8")
    #
    suba = Substation.substation(OlxObj.OLCase,xpu=1e-4)
    ga = []# all GUID of 1st bus of substation
    ind = 1
    for sub1 in suba:
        ba = sub1.BUS
        nbXFMR = len(sub1.XFMR)
        nbXFMR3= len(sub1.XFMR3)
        nbSHIFTER = len (sub1.SHIFTER)
        nbX = nbXFMR + nbXFMR3 + nbSHIFTER
        nbEquipment = len(sub1.equipment)
        nbSWITCH = len(sub1.SWITCH)
        if (nbX >=1) or (nbSWITCH > 0):
            f1.write("\n SubNo: {} ".format(ind))
            f1.write(sub1.toString())
            ga.append(ba[0].GUID)
        ind+=1
    f1.close()
    #
    ki=1
    for g1 in ga:
        OlxObj.OLCase.open(ARGVS.fi,1)
        b1 = OlxObj.OLCase.findBUS(g1)
        sub1 = Substation.substation(b1,xpu=1e-4)
        sub1.saveAsOlr(ff+str(ki)+'.OLR',option=0)#
        ki+=1
#
def sample_4(fi):
    OlxObj.OLCase.open(fi,1)
    b1 = OlxObj.OLCase.findBUS('nevada',132)
##    b1 = OlxObj.OLCase.findBUS('BUS10',132)
    # 3 criterias to considered 1 AC Line is an Equipment of Substation: leng or xpu or xOhm
    sub1= Substation.substation(b1,leng=0.1,xpu=1e-4,xOhm=0.5)
##    print(sub1.toString())
##    print(sub1.toString())
    Substation.grpSubstation(sub1)

    return
#
if __name__ == '__main__':
    ARGVS.fi = 'sample\\SAMPLE30_3.OLR'
    if not os.path.isfile(ARGVS.fi):
       print( 'Input file is missing' )
    else:
##        sample_1(ARGVS.fi)
        sample_2(ARGVS.fi)
        sample_3(ARGVS.fi)
        sample_4(ARGVS.fi)



