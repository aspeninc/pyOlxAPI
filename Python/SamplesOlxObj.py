"""
Purpose:
  OlxObj library application examples
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.1.9"

import os,sys
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

#################################################################################
### OlxObj initialization logic
## TODO: Customize the following pathnames to match your OneLiner installation
## Full pathname of the folder where the OlxAPI.dll is located
olxpath = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
##Full pathname of the folder where the OlxAPI Python wrapper
##   OlxAPI.py and relevant libraries are located
olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python'
olxpathpy = PATH_FILE # delete in 00ship
# INPUTS cmdline ---------------------------------------------------------------
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'OlxObj library application examples'
PARSER_INPUTS.add_argument('-fi'  , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpath' , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
#
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
from OlxObj import *
load_olxapi(ARGVS.olxpath)           # OlxAPI module initialization
## End OlxObj initialization
###############################################################################

#
def sample_1_network(fi):
    """
    OLR network operations demo
    """
    print('sample_1_NetWork')
    OLCase.open(fi,1) # Load the OLR network file as read-only
    #
    print('File comments:', OLCase.COMMENT)
    print('System Base MVA:',OLCase.BASEMVA)
    # Filter the area=[1,10] and kV=132
    OLCase.applyScope(areaNum=[1,10],zoneNum=[],optionTie=0, kV=[132,132])
    #
    print('\nList of all BUSes in the scope')
    for it in OLCase.BUS:
         print( it.toString() )      # text string for object
    #
    print('\nList of all LINEs in the scope')
    for it in OLCase.LINE:
         print( it.toString() )
    OLCase.close()
#
def sample_2_Bus(fi):
    """
    BUS object demo
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_2_Bus')
    #
    b1 = OLCase.findBUS('arizona',132)     # by (name,kV)
    if b1:
        print(b1.toString())         # text string for object
    else:
        print('Bus not found')
    #
    b1 = OLCase.findBUS(28)                # by (busNumber)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OLCase.findBUS('{b4f6d06c-473b-4365-ba05-a1b47a1f9364}') # by (GUID)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OLCase.findBUS("[BUS] 28 'ARIZONA' 132 kV")       # by (STR)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OLCase.BUS[0]                    # by OLCase.BUS[i]
    print(b1.toString())
    #
    b1 = OLCase.LINE[6].BUS2              # by OLCase.LINE[i].BUS
    print(b1.toString())
    #
    print(b1.NAME,b1.kv,b1.no) # Bus name,kV,Bus Number
#
def sample_3_Line(fi):
    """
    LINE object demo
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_3_Line')
    #
    b1 = OLCase.findBUS('claytor',132)
    b2 = OLCase.findBUS('nevada',132)
    l1 = OLCase.findLINE(b1,b2,'1')      # by (BUS1, BUS2, CID)
    if l1:
        print(l1.toString())             # text string for object
    else:
        print('Line not found')
    #
    l1 = OLCase.findLINE("[LINE] 6 'NEVADA' 132 kV-28 'ARIZONA' 132 kV 1") # by (STR)
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OLCase.findLINE('{e0a47740-8f81-43e2-b567-a580e8e6a442}')         # by (GUID)
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OLCase.LINE[7]                    # by OLCase.LINE[i]
    print(l1.toString())
    #
    l1 = b1.LINE[0]                        # by BUS.LINE[i]
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OLCase.RLYGROUP[0].EQUIPMENT      # by RLYGROUP.EQUIPMENT
    print(l1.toString())
    #
    print('R=',l1.R,' X=',l1.X)            # R, X
#
def sample_4_Terminal(fi):
    """
    Demo of TERMINAL object type
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_4_Terminal')
    #
    b1 = OLCase.findBUS('claytor',132)
    b2 = OLCase.findBUS('GLEN LYN',132)
    t1 = OLCase.findTERMINAL(b1,b2,'LINE','1')
    if t1:
        print(t1.toString())
    else:
        print('TERMINAL not found')
    #
    print('\nTERMINALS on ' + b1.toString()+' with CID=1')
    for ti in b1.TERMINAL:
        if ti.CID=='1':
            print(ti.toString())
    #
    print('\nTERMINALS from CLAYTOR to GLEN LYN')
    for ti in b1.terminalTo(b2):
        print(ti.toString())
    #
    l1 = OLCase.LINE[0]
    print('\nTERMINALS on ' + l1.toString() )
    for ti in l1.TERMINAL:
        print(ti.toString())
    #
    x2 = OLCase.XFMR[0]
    print('\nTERMINALS on ' + x2.toString() )
    for ti in x2.TERMINAL:
        print(ti.toString())
    #
    x3 = OLCase.XFMR3[0]
    print('\nTERMINALS on ' + x3.toString() )
    for ti in x3.TERMINAL:
        print(ti.toString())
    #
    o1 = OLCase.SWITCH[0]
    print('\nTERMINALS on ' + o1.toString() )
    for ti in o1.TERMINAL:
        print(ti.toString())
    #
    o1 = OLCase.SHIFTER[0]
    print('\nTERMINALS on ' + o1.toString() )
    for ti in o1.TERMINAL:
        print(ti.toString())
    #
    b1 = OLCase.findBUS('claytor',132)
    b2 = OLCase.findBUS('nevada',132)
    t1 = OLCase.findTERMINAL(b1,b2,'LINE','1')
    print('\nRLYGROUPS on ' + t1.toString() )
    for r1 in t1.RLYGROUP:
        if r1:
            print(r1.toString())
    #
    print('\nOPPOSITE TMNL:')
    for ti in t1.OPPOSITE:
        print(ti.toString())
    #
    print('\nREMOTE TMNL:')    # Tap bus is ignored
    for ti in t1.REMOTE:
        print(ti.toString())
#
def sample_5_changeData(fi):
    """
    Object data update demo
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_5_changeData\n')
    b1 = OLCase.findBUS('arizona',132)
    print('Old name:', b1.NAME)
    b1.NAME = 'New ARIZONA'
    b1.postData()  # Perform validation and update object data
    print('New name:', b1.NAME)

    # Temporary directory to save result files
    import tempfile
    OLCase.save(tempfile.gettempdir()+'\\sample30x.OLR') #save as a new file
#
def sample_6_Setting(fi):
    """
    Relay setting update demo
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_6_Setting\n')
    rl = OLCase.RLYOCP[0]
    print(rl.toString())
    print('VPOL50PF=',rl.getSetting('VPOL50PF'))   # 0.0
    rl.changeSetting('VPOL50PF',0.225)
    rl.postData() # Perform validation and update object data
    print('updated: VPOL50PF=',rl.getSetting('VPOL50PF'))

    # Temporary directory to save result files
    import tempfile
    OLCase.save(tempfile.gettempdir()+'\\sample30x.OLR') #save as a new file
#
def sample_7_ClassicalFault_1(fi):
    """
    Classical Fault on Bus without Outage
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_7_ClassicalFault_1\n')
    b1 = OLCase.findBUS('claytor',132)
    b2 = OLCase.findBUS('nevada',132)
    l1 = OLCase.findLINE(b1,b2,'1')
    t1 = b1.terminalTo(b2)[0]
    ocg = OLCase.findOBJ("[DSRLYG]  Clator_NV G1@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")
    # Define the fault spec
    fs1 = SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='2LG:AB',Z=[0.1,0.2])
    # Run fault simulation
    OLCase.simulateFault(fs1,0)
    # new Fault '3LG' & Z
    fs1.fltConn = '3LG'
    fs1.Z = [0.2,0.3]
    OLCase.simulateFault(fs1,0)
    #RESULT
    for i in range(OLCase.faultNumber()):            # number of fault in buffer
        r1 = RESULT_FLT(i+1)
        print(r1.FaultDescription)
        r1.setScope(tiers=6)
        print('IABC fault=',r1.Current())         # fault current ABC form
        print('I012 t1   =',r1.Current(t1,'012')) # line (l1) current 012 form
        print('IABC l1   =',r1.Current(l1))       # line (l1) current 012 form
        print('VABC b2   =',r1.Voltage(b2))       # voltage at bus b2 (Nevada,132) ABC form
        print('V012 b2   =',r1.Voltage(b2,'012')) # voltage at bus b2 (Nevada,132) 012 form
        print('op,toc    =',r1.OpTime(ocg,1,1))   # Relay operating time
#
def sample_8_ClassicalFault_2(fi):
    """
    Classical Fault on Line with Outage
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_8_ClassicalFault_2')
    b1 = OLCase.findBUS('nevada',132)
    b2 = OLCase.findBUS('claytor',132)
    b3 = OLCase.findBUS('fieldale',132)
    t1 = OLCase.findTERMINAL(b2,b1,LINE,'1')
    l1 = OLCase.findLINE(b2,b1,'1')
    #
    o1 = OUTAGE(option='single-gnd',G=0.4)
    o1.add_outageLst(l1)
    va = o1.build_outageLst(obj=t1,tiers=2,wantedTypes=['XFMR'])
    for v1 in va:
        print(v1.toString())
    #
    fs1 = SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='2LG:AB',outage=o1)
    OLCase.simulateFault(fs1,0)
    fs1.outage.option = 'DOUBLE'
    OLCase.simulateFault(fs1,0)
    #RESULT
    for i in range(OLCase.faultNumber()):        # number of fault in buffer
        r1 = RESULT_FLT(i+1)
        print(r1.FaultDescription)
        print('IABC fault=',r1.Current())    # fault current ABC form
#
def sample_9_SimultaneousFault(fi):
    """
    Simultaneous Fault
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_9_SimultaneousFault')
    b1 = OLCase.findBUS('nevada',132)
    b2 = OLCase.findBUS('claytor',132)
    b3 = OLCase.findBUS('FIELDALE',132)
    b4 = OLCase.findBUS('TEXAS',132)
    b5 = OLCase.findBUS('TENNESSEE',132)
    b6 = OLCase.findBUS('NEW HAMPSHR',33)
    l1 = OLCase.findLINE(b2,b1,'1')
    x1 = OLCase.findXFMR(b6,b1,'1')
    rg = l1.RLYGROUP[0]
    rg2 = x1.RLYGROUP[0]
    t1 = b1.terminalTo(b2)[0]
    # DEFINE and RUN FAULT
    fs1 = SPEC_FLT.Simultaneous(obj=[b1],fltApp='Bus',fltConn='ll:bc')
    fs1 = SPEC_FLT.Simultaneous(obj=t1,fltApp='2P-open',fltConn='ac')
    fs1 = SPEC_FLT.Simultaneous(obj=[b1,b2],fltApp='Bus2Bus',fltConn='cb')
    fs1 = SPEC_FLT.Simultaneous(obj=rg,fltApp='CLOSE-IN',fltConn='3LG')
    fs2 = SPEC_FLT.Simultaneous(obj=t1,fltApp='15%',fltConn='3LG',Z=[0.1,0.2,0.5,0,0,0,0,0.5])
    #
    OLCase.simulateFault([fs1],0)
    fs1.fltConn ='1LG:A'
    fs2.fltApp  ='25%'
    OLCase.simulateFault([fs1,fs2],0)
    #RESULT
    for i in range(OLCase.faultNumber()):       # number of fault in buffer
        r1 = RESULT_FLT(i+1)
        print(r1.FaultDescription)
        print('IABC fault=',r1.Current())   # fault current ABC form
        print('VABC=',r1.Voltage(b2))       # voltage at bus b2 (Claytor,132) ABC form
#
def sample_10_SEA_1(nw):
    """
    SEA: Stepped-Event Analysis
        Single User-Defined Event
    """
    OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_10_SEA_1')
    b1 = OLCase.findBUS('nevada',132)
    b2 = OLCase.findBUS('claytor',132)
    t1 = OLCase.findTERMINAL(b1,b2,'LINE','1')
    rg = OLCase.findOBJ("[RELAYGROUP] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")

    # DEFINE and RUN FAULT
    fs1 = SPEC_FLT.SEA(obj=t1,fltApp='15%',fltConn='1LG:A',deviceOpt=[1,1,1,1,1,1,1],tiers=-5,Z=[0.1,0.2])
    OLCase.simulateFault(fs1,0)

####    fs2 = SPEC_FLT.SEA_ex(obj=b2,fltConn='1LG:A',Z=[0,0],time=0.2)
####    fs3 = SPEC_FLT.SEA_ex(obj=b2,fltConn='1LG:A',Z=[0,0],time=0.3)
##    SimulateFault([fs1])#,fs2,fs3]
##
##    #RESULT
##    for i in range(RESULT_FLT.COUNT()):# .COUNT
##        r1 = RESULT_FLT(index=i+1,tiers=5)
##        print('time=',r1.SEAtime)
##        print(r1.FaultDescription)
####        r2 = RESULT_FLT(index=i+2,tiers=5)
####        print(r2.FaultDescription)
####        print('I=',abs(r2.Current()[0]))
##        print()
##        print(r1.FaultDescription) #[fltdes,time] toString() SEAtime
##        print(r1.SEADescription)
##        print('I=',abs(r1.Current()[0])) # SEABreaker  SEADevice
##        print('V=',r1.Voltage(b1))

##        break
#
def sample_11_SEA_2(nw):
    """
    SEA: Stepped-Event Analysis
        Multiple User-Defined Event
    """
    print('\nsample_10_SEA_2')
    b1 = BUS.init('nevada',132)
    b2 = BUS.init('claytor',132)

    # DEFINE and RUN FAULT
    addParam1 = ['2LG:AB',0.02,0,0]
    fs1 = SPEC_FLT.SEA(obj=b2,fltConn='1LG:A',runOpt=[1,1,1,1,1,1,1],tiers=5,Z=[0,0],addParam=[addParam1])
    SimulateFault(fs1)

    # RESULT
    for i in range(CountFault()):
        r1 = RESULT_FLT(index=i+1,tiers=5)
        print()
        print(r1.FaultDescription)
        print(r1.SEADescription)
        print('I=',abs(r1.Current()[0]))
        print('V=',r1.Voltage(b1))
#
if __name__ == '__main__':
    ARGVS.fi ='C:\\Program Files (x86)\\ASPEN\\1LPFv15\\SAMPLE30.OLR'
    # Various examples
    sample_1_network(ARGVS.fi)
    sample_2_Bus(ARGVS.fi)
    sample_3_Line(ARGVS.fi)
    sample_4_Terminal(ARGVS.fi)
    sample_5_changeData(ARGVS.fi)
    sample_6_Setting(ARGVS.fi)
    sample_7_ClassicalFault_1(ARGVS.fi)
    sample_8_ClassicalFault_2(ARGVS.fi)
    sample_9_SimultaneousFault(ARGVS.fi)
##    sample_10_SEA_1(ARGVS.fi)
##    sample_11_SEA_2(ARGVS.fi)

