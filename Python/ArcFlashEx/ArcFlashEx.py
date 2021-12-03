"""
Purpose:  Run DoArcFlash (IEEE-1584 2018) with input data from a template file
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

#
import sys,os
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
PARSER_INPUTS = AppUtils.iniInput(usage=
"\n\tRun DoArcFlash with input data from a template file\
\n\tin CSV/Excel format with columns (fields) in following order:\
\n\
\nIEEE-1584 2018 Arc flash\
\n  1.  'No.'                - Bus number  (optional)\
\n  2.  'Bus Name'	         - bus name\
\n  3.  'kV'	             - bus kv\
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
\n  12. 'Fault clearing option\
\n                           - For clearing time option FASTEST: NO of tiers\
\n                           - For clearing time option FIXED: Fixed delay\
\n                           - For clearing Time Option FUSE: fuse LibraryName:CurveName")
#
ft_default = os.path.join(PATH_FILE,'InputTemplate2018.xlsx')
if not os.path.isfile(ft_default):
    ft_default= ''
#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-ft' , help = '*(str) Input template file (CSV/Excel format)',default = ft_default,type=str,metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name of the report file (CSV/Excel format)',default = '',type=str,metavar='')
#
ARGVS = AppUtils.parseInput(PARSER_INPUTS,demo=1)

#
def checkInput():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return False
    #
    a,s = AppUtils.checkFile(ARGVS.ft,[".XLSX",".XLS",".XLSM",".CSV"],'Input template file')
    if not a:
        return AppUtils.FinishCheck(PY_FILE,s,PARSER_INPUTS)
    #
    return True
#
def run():
    if not checkInput():
        return
    print("Template file: " +  ARGVS.ft )
    print("IEEE-1584 2018")
    #
    import xml.etree.ElementTree as ET
    # update file out
    ARGVS.fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='' , ad='' , ext='.csv')
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    ws = AppUtils.ToolCSVExcell()
    ws.readFile(fi=ARGVS.ft,delim=',')
    #
    nSuccess = 0
    #
    for i in range(7,10000):
        if ws.getVal(row=i, column=1)==None:
            break
        #
        sbus = getStrBus(ws,i)
        if sbus !="":
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
            if OlxAPILib.run1LPFCommand(sInput):
                nSuccess += 1
    #
    if nSuccess>0:
        print("\nArc-flash calculation ran successfully on " + str(nSuccess) + " buses")
        print("Report file had been saved in:\n%s"%ARGVS.fo)

        # open res in Excel
        if ARGVS.ut ==0 :
            AppUtils.launch_excel(ARGVS.fo)
    else:
        print("Something was not right (no bus found). No arc flash calculation results.")
    return 1

#
def getStrBus(ws,i):
    nBusNumber = int  (ws.getVal(row=i, column=1))
    sBusName   = str  (ws.getVal(row=i, column=2))
    dBusKv     = float(ws.getVal(row=i, column=3))
    sbus = "'"+ sBusName + "'," + str(round(dBusKv,3))
    #
    bhnd  = OlxAPI.FindBus(sBusName,dBusKv)
    if bhnd>0:
        print(sbus.ljust(25), ' : BUS FOUND')
        return sbus
    #
    if nBusNumber>0:
        bhnd = OlxAPI.FindBusNo(nBusNumber)
    #
    if bhnd<=0:
        print(sbus.ljust(25), ' : BUS NOT FOUND')
        return ''
    #
    name1 = OlxAPILib.getEquipmentData([bhnd],BUS_sName)[0]
    kv1 = OlxAPILib.getEquipmentData([bhnd],BUS_dKVnominal)[0]
    sbus = "'"+ name1 + "'," + str(round(kv1,3))
    print(sbus.ljust(25), ' : BUS FOUND')
    return sbus
#
def unit_test():
    ARGVS.fi = os.path.join(PATH_FILE,'ARCFLASH.OLR')
    run()
    sres  = "OLR file: "+os.path.basename(ARGVS.fi) +"\n"
    sres += "Template file: "+os.path.basename(ARGVS.ft) +"\n"
    sres += AppUtils.read_File_text_2(ARGVS.fo,5)
    #
    AppUtils.deleteFile(ARGVS.fo)
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)

#
def run_demo():
    if ARGVS.demo ==1:
        ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
        ARGVS.fo = AppUtils.get_file_out(fo='' , fi=ARGVS.fi , subf='' , ad='_'+os.path.splitext(PY_FILE)[0]+'_demo' , ext='.csv')
        ARGVS.ft = os.path.join(PATH_FILE,'InputTemplate2018.xlsx')
        #
        choice = AppUtils.ask_run_demo(PY_FILE,ARGVS.ut,ARGVS.fi,"Input template file: " + ARGVS.ft)
        #
        if choice=="yes":
            return run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def main():
    if ARGVS.ut >0:
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


