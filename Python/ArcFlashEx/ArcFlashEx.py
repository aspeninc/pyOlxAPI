"""
Purpose:  Run DoArcFlash (IEEE-1584 2018) with input data from a template file
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "2.1.4"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage =\
    "\nRun DoArcFlash with input data from a template file\
    \nin CSV/Excel format with columns (fields) in following order:\
    \n\
    \nIEEE-1584 2018 Arc flash\
    \n  1.  'No.'                - Bus number  (optional)\
    \n  2.  'Bus Name'             - bus name\
    \n  3.  'kV'                 - bus kv\
    \n  4.  'Electrode config'   - 0: VCB  1: VCBB  2: HCB  3: VOA  4: HOA\
    \n  5.  'Enclosure H (in.)'  - Applicable for electrode configs VCB, VCBB, HCB\
    \n  6.  'Enclosure W (in.)'  - Applicable for electrode configs VCB, VCBB, HCB\
    \n  7.  'Enclosure D (in.)'  - Applicable for electrode configs VCB, VCBB, HCB at voltage 600V or lower\
    \n  8.  'Conductor Gap (mm)'\
    \n  9.  'Working Distance (inches)'\
    \n  10. 'Breaker interrupting Time (cycles)'\
    \n  11. 'Fault Clearing'     - FUSE: Use fuse curve\
    \n                           - FIXED: Use fixed duration\
    \n                           - FASTEST: Use fastest trip time of device in vicinity\
    \n  12. 'Fault clearing option'\
    \n                           - For clearing time option FASTEST: NO of tiers\
    \n                           - For clearing time option FIXED: Fixed delay\
    \n                           - For clearing Time Option FUSE: fuse LibraryName:CurveName"
#
ft_default = PATH_FILE+'\\InputTemplate2018.csv'
#
PARSER_INPUTS.add_argument('-ft' ,  help = '*(str) Input template file (CSV/Excel format)', default = ft_default, type=str, metavar='')
PARSER_INPUTS.add_argument('-fo' ,  help = ' (str) Path name of the report file (CSV/Excel format)', default = '', type=str, metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0, type=int, metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]
import xml.etree.ElementTree as ET


# CORE--------------------------------------------------------------------------
def getStrBus(ws,i):
    nBusNumber = int  (ws.getVal(row=i, column=1))
    sBusName   = str  (ws.getVal(row=i, column=2))
    dBusKv     = float(ws.getVal(row=i, column=3))

    #
    b1 = OLCase.findOBJ('BUS',[sBusName,dBusKv])
    if b1==None and nBusNumber>0:
        b1 = OLCase.findOBJ('BUS', nBusNumber)
    #
    if b1==None:
        sbus = "'%s',%f"%(sBusName,dBusKv)
        print(sbus.ljust(25), ' : BUS NOT FOUND')
        return ''
    sbus = "'%s',%f"%(b1.NAME, b1.KV)
    print(sbus.ljust(25), ' : BUS FOUND')
    return sbus


def run():
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened

    #
    _, er = AppUtils.checkFile(ARGVS.ft,[".XLSX",".XLS",".XLSM",".CSV"],'Input template file')
    if er:
        PARSER_INPUTS.print_help()
        AppUtils.gui_error('Arc-flash calculation', er)

    print("Template file: ", ARGVS.ft)
    print("IEEE-1584 2018")

    # update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=OLCase.olrFile , subf='' , ad='' , ext='.csv')
    #
    ws = AppUtils.ToolCSVExcell()
    ws.readFile(fi=ARGVS.ft, delim=',')
    ws2 = AppUtils.ToolCSVExcell()
    ws2.readFile(fi=ARGVS.ft, delim=',')
    #
    nSuccess = 0
    #
    for i in range(7,10000):
        if ws.getVal(row=i, column=1)==None:
            break
        #
        sbus = getStrBus(ws,i)
        if sbus:
            if str(ws.getVal(row=6, column=4)) =="Electrode config":     # standardYear = 2018
                data = ET.Element('ARCFLASHCALCULATOR2018')
                data.set('REPFILENAME',ARGVS.fo)  # fo
                data.set('OUTFILETYPE','2') # csv file
                if nSuccess==0:
                    data.set('APPENDREPORT','0')
                else:
                    data.set('APPENDREPORT','1')
                #
                data.set('SELECTEDOBJ',sbus) # bus
                data.set('ELECTRODECFG',str(ws.getVal(row=i, column=4)))
                data.set('BOXH'        ,str(ws.getVal(row=i, column=5)))
                data.set('BOXW'        ,str(ws.getVal(row=i, column=6)))
                data.set('BOXD'        ,str(ws.getVal(row=i, column=7)))
                data.set('CONDUCTORGAP',str(ws.getVal(row=i, column=8)))
                data.set('WORKDIST'    ,str(ws.getVal(row=i, column=9)))
                data.set('BRKINTTIME'  ,str(ws.getVal(row=i, column=10)))
                data.set('ARCDURATION' ,str(ws.getVal(row=i, column=11)))
                if str(ws.getVal(row=i, column=11)) =="FUSE":
                    data.set('FUSECURVE',str(ws.getVal(row=i, column=12)))
                elif str(ws.getVal(row=i, column=11)) =="FIXED":
                    data.set('ARCTIME' ,str(ws.getVal(row=i, column=12)))
                elif str(ws.getVal(row=i, column=11)) =="FASTEST":
                    data.set('DEVICETIERS',str(ws.getVal(row=i, column=12)))
            #-------------------------------------------------------------------
            elif str(ws.getVal(row=6, column=4)) =="Equipment Category": # standardYear = 2012
                raise Exception("Not yet support for standardYear = 2012")
                #
            else:
                raise Exception("Error Template File")
            #
            sInput = ET.tostring(data)
            #
            OLCase.run1LPFCommand(sInput)
            nSuccess += 1
    #
    ws.close()
    if nSuccess>0:
        sMain = "\nArc-flash calculation ran successfully on %i buses"%nSuccess
        sMain+= "\n\nReport file had been saved in:\n%s"%ARGVS.fo
    else:
        sMain="Something was not right (no bus found).\nNo arc flash calculation results."

    AppUtils.explorerDir(ARGVS.fo, sMain, PY_FILE) #open dir of fo

def run_demo():
    if ARGVS.demo==1:
        fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        OLCase.open(fi,1)
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.csv')
        ARGVS.ft = ft_default
        #
        sMain = "\nOLR file: " + fi
        sMain+= "\n\nInput template file: " + ARGVS.ft
        sMain+= "\n\nDo you want to run this demo (Y/N)?"
        choice = AppUtils.gui_askquestion(PY_FILE+' Demo',sMain)
        if choice=='YES':
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE, ARGVS.demo, [1],gui_err=1)


def main():
    if ARGVS.demo>0:
        return run_demo()
    run()


if __name__ == '__main__':
    main()



