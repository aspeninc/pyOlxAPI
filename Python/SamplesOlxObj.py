"""
Purpose:
  PyOlxObj application examples
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__email__     = "support@aspeninc.com"
__status__    = "Release-candidate"
__version__   = "2.2.5"

import os,sys
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

#################################################################################
### OlxObj initialization logic
## TODO: Customize the following pathnames to match your OneLiner installation
## Full pathname of the folder where the OlxAPI.dll is located
olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15_training'
if not os.path.isdir(olxpath):
    olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
##Full pathname of the folder where the OlxAPI Python wrapper
##   OlxAPI.py and relevant libraries are located
olxpathpy = 'c:\\ASPENv15TrainingData\\OlxAPI\\Python'
if not os.path.isdir(olxpathpy):
    olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\OlxAPI\python'
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
import OlxObj
OlxObj.load_olxapi(ARGVS.olxpath)           # OlxAPI module initialization
## End OlxObj initialization
###############################################################################

#
def sample_1_network(fi):
    """
    OLR network operations demo
    """
    print('sample_1_NetWork')
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    #
    print('File comments:', OlxObj.OLCase.COMMENT)
    print('System Base MVA:',OlxObj.OLCase.BASEMVA)
    # Filter the area=[1,10] and kV=132
    OlxObj.OLCase.applyScope(areaNum=[1,10],zoneNum=[],optionTie=0, kV=[132,132])
    #
    print('\nList of all BUSes in the scope')
    for it in OlxObj.OLCase.BUS:
         print( it.toString() )      # text string for object
    #
    print('\nList of all LINEs in the scope')
    for it in OlxObj.OLCase.LINE:
         print( it.toString() )
    OlxObj.OLCase.close()
#
def sample_2_Bus(fi):
    """
    BUS object demo
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_2_Bus')
    #
    b1 = OlxObj.OLCase.findBUS('arizona',132)     # by (name,kV)
    if b1:
        print(b1.toString())         # text string for object
    else:
        print('Bus not found')
    #
    b1 = OlxObj.OLCase.findBUS(28)                # by (busNumber)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OlxObj.OLCase.findBUS('{b4f6d06c-473b-4365-ba05-a1b47a1f9364}') # by (GUID)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OlxObj.OLCase.findBUS("[BUS] 28 'ARIZONA' 132 kV")       # by (STR)
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')
    #
    b1 = OlxObj.OLCase.BUS[0]                    # by OlxObj.OLCase.BUS[i]
    print(b1.toString())
    #
    b1 = OlxObj.OLCase.LINE[0].BUS2              # by OlxObj.OLCase.LINE[i].BUS
    print(b1.toString())
    #
    print(b1.NAME,b1.kv,b1.no) # Bus name,kV,Bus Number

#
def sample_3_Line(fi):
    """
    LINE object demo
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_3_Line')
    #
    b1 = OlxObj.OLCase.findBUS('claytor',132)
    b2 = OlxObj.OLCase.findBUS('nevada',132)
    l1 = OlxObj.OLCase.findLINE(b1,b2,'1')      # by (BUS1, BUS2, CID)
    if l1:
        print(l1.toString())             # text string for object
    else:
        print('Line not found')
    #
    l1 = OlxObj.OLCase.findLINE("[LINE] 6 'NEVADA' 132 kV-28 'ARIZONA' 132 kV 1") # by (STR)
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OlxObj.OLCase.findLINE('{e0a47740-8f81-43e2-b567-a580e8e6a442}')         # by (GUID)
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OlxObj.OLCase.LINE[7]                    # by OlxObj.OLCase.LINE[i]
    print(l1.toString())
    #
    l1 = b1.LINE[0]                        # by BUS.LINE[i]
    if l1:
        print(l1.toString())
    else:
        print('Line not found')
    #
    l1 = OlxObj.OLCase.RLYGROUP[0].EQUIPMENT      # by RLYGROUP.EQUIPMENT
    print(l1.toString())
    #
    print('R=',l1.R,' X=',l1.X)            # R, X
#
def sample_4_Terminal(fi):
    """
    Demo of TERMINAL object type
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_4_Terminal')
    #
    b1 = OlxObj.OLCase.findBUS('claytor',132)
    b2 = OlxObj.OLCase.findBUS('GLEN LYN',132)
    t1 = OlxObj.OLCase.findTERMINAL(b1,b2,'LINE','1')
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
    l1 = OlxObj.OLCase.LINE[0]
    print('\nTERMINALS on ' + l1.toString() )
    for ti in l1.TERMINAL:
        print(ti.toString())
    #
    x2 = OlxObj.OLCase.XFMR[0]
    print('\nTERMINALS on ' + x2.toString() )
    for ti in x2.TERMINAL:
        print(ti.toString())
    #
    x3 = OlxObj.OLCase.XFMR3[0]
    print('\nTERMINALS on ' + x3.toString() )
    for ti in x3.TERMINAL:
        print(ti.toString())
    #
    o1 = OlxObj.OLCase.SWITCH[0]
    print('\nTERMINALS on ' + o1.toString() )
    for ti in o1.TERMINAL:
        print(ti.toString())
    #
    o1 = OlxObj.OLCase.SHIFTER[0]
    print('\nTERMINALS on ' + o1.toString() )
    for ti in o1.TERMINAL:
        print(ti.toString())
    #
    b1 = OlxObj.OLCase.findBUS('claytor',132)
    b2 = OlxObj.OLCase.findBUS('nevada',132)
    t1 = OlxObj.OLCase.findTERMINAL(b1,b2,'LINE','1')
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
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_5_changeData\n')
    b1 = OlxObj.OLCase.findBUS('arizona',132)
    print('Old name:', b1.NAME)
    b1.NAME = 'New ARIZONA'
    b1.postData()  # Perform validation and update object data
    print('New name:', b1.NAME)

    # Temporary directory to save result files
    import tempfile
    OlxObj.OLCase.save(tempfile.gettempdir()+'\\sample30x.OLR') #save as a new file
#
def sample_6_Setting(fi):
    """
    Relay setting update demo
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_6_Setting\n')
    rl = OlxObj.OLCase.RLYOCP[0]
    print(rl.toString())
    print('VPOL50PF=',rl.getSetting('VPOL50PF'))   # 0.0
    rl.changeSetting('VPOL50PF',0.225)
    rl.postData() # Perform validation and update object data
    print('updated: VPOL50PF=',rl.getSetting('VPOL50PF'))

    # Temporary directory to save result files
    import tempfile
    OlxObj.OLCase.save(tempfile.gettempdir()+'\\sample30x.OLR') #save as a new file
#
def sample_7_ClassicalFault_1(fi):
    """
    Classical Fault on Bus without Outage
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('sample_7_ClassicalFault_1\n')
    b1 = OlxObj.OLCase.findBUS('claytor',132)
    b2 = OlxObj.OLCase.findBUS('nevada',132)
    b3 = OlxObj.OLCase.findBUS('alaska',33)
    l1 = OlxObj.OLCase.findLINE(b1,b2,'1')
    t1 = b1.terminalTo(b2)[0]
    ocg = OlxObj.OLCase.findOBJ("[DSRLYG]  Clator_NV G1@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")
    # Define the fault spec
    fs1 = OlxObj.SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='2LG:AB',Z=[0.1,0.2])
    # Run fault simulation
    OlxObj.OLCase.simulateFault(fs1,0)
    ri = OlxObj.FltSimResult[0]
    # new Fault '3LG' & Z
    fs1.fltConn = '3LG'
    fs1.Z = [0.2,0.3]
    OlxObj.OLCase.simulateFault(fs1,1)
    #RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)
        r1.setScope(tiers=6)
        print('\tIABC fault=',r1.current())
        print('\tIABC t1   =',r1.current(t1))
        print('\tVABC b1   =',r1.voltage(b1))
        print('\tVABC l1   =',r1.voltage(l1))
        print('\tI012 fault=',r1.currentSeq())
        print('\tI012 t1   =',r1.currentSeq(t1))
        print('\tV012 b1   =',r1.voltageSeq(b1))
        print('\tV012 l1   =',r1.voltageSeq(l1))
        #
        print('\nop,toc    =',r1.optime(ocg,1,1))   # Relay operating time
#
def sample_8_ClassicalFault_2(fi):
    """
    Classical Fault on Line with Outage
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_8_ClassicalFault_2')
    b1 = OlxObj.OLCase.findBUS('nevada',132)
    b2 = OlxObj.OLCase.findBUS('claytor',132)
    b3 = OlxObj.OLCase.findBUS('fieldale',132)
    t1 = OlxObj.OLCase.findTERMINAL(b2,b1,'LINE','1')
    l1 = OlxObj.OLCase.findLINE(b2,b1,'1')
    #
    o1 = OlxObj.OUTAGE(option='single-gnd',G=0.4)
    o1.add_outageLst(l1)
    o1.build_outageLst(obj=t1,tiers=2,wantedTypes=['XFMR'])
    for v1 in o1.outageLst:
        print(v1.toString())
    #
    fs1 = OlxObj.SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='2LG:AB',outage=o1)
    OlxObj.OLCase.simulateFault(fs1,0)
    fs1.outage.option = 'DOUBLE'
    OlxObj.OLCase.simulateFault(fs1,0)
    #RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)
        print('\tIABC fault=',r1.current())
        print('\tI012 t1=',r1.currentSeq(t1))
#
def sample_9_SimultaneousFault(fi):
    """
    Simultaneous Fault
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_9_SimultaneousFault')
    b1 = OlxObj.OLCase.findBUS('nevada',132)
    b2 = OlxObj.OLCase.findBUS('claytor',132)
    b3 = OlxObj.OLCase.findBUS('FIELDALE',132)
    b4 = OlxObj.OLCase.findBUS('TEXAS',132)
    b5 = OlxObj.OLCase.findBUS('TENNESSEE',132)
    b6 = OlxObj.OLCase.findBUS('nevada',132)
    l1 = OlxObj.OLCase.findLINE(b2,b1,'1')
    x1 = OlxObj.OLCase.findXFMR(b6,b1,'1')
    rg = l1.RLYGROUP[0]
    rg2 = x1.RLYGROUP[0]
    t1 = b1.terminalTo(b2)[0]
    # DEFINE and RUN FAULT
    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=b1,fltApp='Bus',fltConn='ll:bc')
##    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=t1,fltApp='2P-open',fltConn='ac')
##    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=[b1,b2],fltApp='Bus2Bus',fltConn='cb')
##    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=rg,fltApp='CLOSE-IN',fltConn='3LG')
    fs2 = OlxObj.SPEC_FLT.Simultaneous(obj=t1,fltApp='15%',fltConn='3LG',Z=[0.1,0.2,0.5,0,0,0,0,0.5])
    #
    OlxObj.OLCase.simulateFault([fs1],0)
    fs2.fltApp  ='25%'
    OlxObj.OLCase.simulateFault([fs1,fs2],0)
    #RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)
        print('\tIABC fault=',r1.current())
        print('\tVABC b2   =',r1.voltage(b2))
        print('\tV012 l1   =',r1.voltageSeq(l1))
#
def sample_10_SEA_1(fi):
    """
    SEA: Stepped-Event Analysis
        Single User-Defined Event
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_10_SEA_1')
    b1 = OlxObj.OLCase.findBUS('nevada',132)
    b2 = OlxObj.OLCase.findBUS('claytor',132)
    t1 = OlxObj.OLCase.findTERMINAL(b1,b2,'LINE','1')
    t2 = OlxObj.OLCase.findTERMINAL(b2,b1,'LINE','1')
    rg = OlxObj.OLCase.findRLYGROUP("[RELAYGROUP] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")
    # DEFINE and RUN FAULT
    fs1 = OlxObj.SPEC_FLT.SEA(obj=t1,fltApp='15%',fltConn='1LG:A',deviceOpt=[1,1,1,1,1,1,1],tiers=5,Z=[0,0])
    OlxObj.OLCase.simulateFault(fs1)
    #RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)
        print('\ttime      =',r1.SEARes['time'])
        print('\tImax      =',r1.SEARes['currentMax'])
        print('\tIABC Fault=',r1.current())
        print('\tIABC t2   =',r1.current(t2))
        print('\tIABC t1   =',r1.current(t1))
        print('\tV012 b1   =',r1.voltageSeq(b1))
#
def sample_11_SEA_2(fi):
    """
    SEA: Stepped-Event Analysis
        Multiple User-Defined Event
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_11_SEA_2')
    b1 = OlxObj.OLCase.findBUS('nevada',132)
    b2 = OlxObj.OLCase.findBUS('claytor',132)
    t1 = OlxObj.OLCase.findTERMINAL(b1,b2,'LINE','1')
    t2 = OlxObj.OLCase.findTERMINAL(b2,b1,'LINE','1')
    rg = OlxObj.OLCase.findRLYGROUP("[RELAYGROUP] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")
    # DEFINE and RUN FAULT
    fs1 = OlxObj.SPEC_FLT.SEA(obj=t1,fltApp='15%',fltConn='1LG:A',deviceOpt=[1,1,1,1,1,1,1],tiers=5,Z=[0,0])
    fs2 = OlxObj.SPEC_FLT.SEA_EX(time=0.01,fltConn='2LG:AB',Z=[0.01,0.02])
    fs3 = OlxObj.SPEC_FLT.SEA_EX(time=0.02,fltConn='3LG',Z=[0.01,0.02])
    #
    OlxObj.OLCase.simulateFault([fs1,fs2,fs3])
    #RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)
        print('\ttime     =',r1.SEARes['time'])
        print('\tEventFlag=',r1.SEARes['EventFlag'])
        print('\tEventDesc=',r1.SEARes['EventDesc'])
        print('\tFaultDesc=',r1.SEARes['FaultDesc'])
        print('\tIABC t2  =',r1.current(t2))
        print('\tIABC t1  =',r1.current(t1))
        print('\tV012 b1  =',r1.voltageSeq(b1))

def sample_12_substation(fi):
    """
    Grouper bus by substation
    """
    OlxObj.OLCase.open(fi,1) # Load the OLR network file as read-only
    print('\nsample_12_substation')
    sub = OlxObj.OLCase.substationGroup()
    lstBus = sub['BUS']
    lstEquipment = sub['EQUIPMENT']
    for i in range(len(lstBus)):
        print('\nSubstation %i'%(i+1))
        for b1 in lstBus[i]:
            print('\t',b1.toString())
        for e1 in lstEquipment[i]:
            print('\t',e1.toString())
def sample_13_tapLine(fi):
    print('\nsample_13_tapLine')
    #
    OlxObj.OLCase.open(fi,1)
    #
    t1 = OlxObj.OLCase.findRLYGROUP("[RELAYGROUP] 2 'NEVADA' 132 kV-'Nev/Ariz Tap' 132 kV 1 L")
    #t1 = OlxObj.OLCase.findRLYGROUP("{0727a192-37d8-440e-9623-4c6ba99451d0}")
    res = OlxObj.OLCase.tapLineTool(t1) # Not print to stdout all details when research mainLine
    #
    mainLine = res['mainLine']
    remoteBus = res['remoteBus']
    remoteRLG = res['remoteRLG']
    Z1 = res['Z1']
    Z0 = res['Z0']
    Length = res['Length']
    #
    for i in range(len(mainLine)):
        print('mainLine: ',i+1)
        print('\tremoteBus: '+OlxObj.toString(remoteBus[i]))
        print('\tremoteRLG: '+OlxObj.toString(remoteRLG[i]))
        print('\tSegments:')
        for v1 in mainLine[i]:
            print('\t\t',v1.toString())
        print('\tZ1      : ',OlxObj.toString(Z1[i]))
        print('\tZ0      : ',OlxObj.toString(Z0[i]))
        print('\tLength  : ',OlxObj.toString(Length[i]))

if __name__ == '__main__':
    if ARGVS.fi == '':
       ARGVS.fi =  "c:\\Program Files (x86)\\ASPEN\\1LPFv15\\SAMPLE30.OLR"
    if not os.path.isfile(ARGVS.fi):
       print( 'Input file is missing' )
    else:
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
        sample_10_SEA_1(ARGVS.fi)
        sample_11_SEA_2(ARGVS.fi)
        sample_12_substation(ARGVS.fi)
        sample_13_tapLine(ARGVS.fi)

