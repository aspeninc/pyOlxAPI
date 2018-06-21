""" Python wrapper for ASPEN OlrxAPI 
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2017, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

from ctypes import *
import os.path
from pyOlrxAPIConst import *

class OlrxAPIException(Exception):
    """Raise this exeption to report OlrxAPI error messages

    """
    def __init__(self, msg):
        self.msg = msg
    def __str__(self):
        return self.msg

ASPENOlrx = None    #

def InitOlrxAPI(dllPath):
    """Initialize OlrxAPI session.

    Successfull initialization is required before any 
    other OlrxAPI call can be executed.

    Args:
        dllPath (string): Full path name with \\ as last character of
                ASPEN program folder where alrxapi.dll and related 
                program components are located

    Returns:
        None

    Raises:
        OlrxAPIException
    """
    # get the module handle and create a ctypes library object
    if not os.path.isfile(dllPath + "olrxapi.dll") :
        raise OlrxAPIException("Failed to locate: " + dllPath + "olrxapi.dll") 
    os.environ['PATH'] = os.path.abspath(dllPath) + ";" + os.environ['PATH']
    global ASPENOlrx
    ASPENOlrx = WinDLL((dllPath + "olrxapi.dll").encode('UTF-8'), use_last_error=True)
    if ASPENOlrx == None:
        raise OlrxAPIException("Failed to setup olrxapi.dll") 
    errorAPIInit = "No Error"            # This string Must match the c++ DLL code
    try:
        # Attempt to get error string
        errorAPIInit = OlrxAPIErrorString()
    except:
        pass
    if errorAPIInit <> "No Error":
        raise OlrxAPIException(errorAPIInit) 

# API function prototypes
def OlrxAPIVersion():
    """OlrxAPI engine version string in format version_major.version_minor
    """
    buf = create_string_buffer(b'\000' * 128)
    ASPENOlrx.OlrxAPIVersionInfo(buf)
    vData = buf.value.split(" ")
    return vData[2]

def OlrxAPIBuildNumber():
    """OlrxAPI engine build number
    """
    buf = create_string_buffer(b'\000' * 128)
    ASPENOlrx.OlrxAPIVersionInfo(buf)
    vData = buf.value.split(" ")
    return int(vData[4])


def OlrxAPILoadDataFile(filePath,readonly):
    """Read ASPEN OLR data file from disk

    Args:
        filePath (c_char_p) : Full path name of ASPEN OLR file.
        readonly (c_int): open in read-only mode. 1-true; 0-false

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPILoadDataFile.argtypes = [c_char_p,c_int]
    return ASPENOlrx.OlrxAPILoadDataFile(filePath,readonly)

def OlrxAPIGetEquipment(type, p_hnd):
    """Retrieves handle of the next equipment of given type in the system.
     
    This function will return the handle of all the objectsof the given type,
    one by one, in the order they are stored in the OLR fileD

    Set hnd to 0 to retrieve the first object

    Args:
        type (c_int): Object type
        p_hnd (pointer(c_int)): Object handle

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success

    """
    ASPENOlrx.OlrxAPIGetEquipment.argstype = [c_int,c_void_p]
    return ASPENOlrx.OlrxAPIGetEquipment(type, p_hnd)


def OlrxAPIEquipmentType(hnd):
    """Returns equipment of the given handle.
     
    Args:
        hdn (c_int): Object handle

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success

    """
    ASPENOlrx.OlrxAPIGetEquipment.argstype = [c_int]
    return ASPENOlrx.OlrxAPIEquipmentType(hnd)

def OlrxAPIGetData(hnd, token, dataBuf):
    """Get object data field

    Args:
        hnd (c_int):        Object handle
        token (c_int):      field token
        dataBuf (c_void_p): data buffer

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetData.argtypes = [c_int,c_int,c_void_p]
    return ASPENOlrx.OlrxAPIGetData(hnd, token, dataBuf)

def OlrxAPISetData(hnd, token, p_data):
    """Assign a value to an object data field

    Args:
        hnd (c_int): Object handle
        token (c_int): Object field token
        p_data (c_void_p): New field value

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPISetData.argstype = [c_int, c_int, c_void_p]
    return ASPENOlrx.OlrxAPISetData(hnd, token, p_data)

def OlrxAPIGetBusEquipment(hndBus, type, pHnd):
    """Retrieves the handle of the next equipment of a given type 
    that is attached to a bus.

    The equipment type can be one of the following:
    TC_GEN: 	to get the handle for the generator. 
                There can be at most one at a bus.
    TC_LOAD: 	to get the handle for the load. 
                There can be at most one at a bus.
    TC_SHUNT:	to get the handle for the shunt. 
                There can be at most one at a bus.
    TC_SVD:	    to get the handle for the switched shunt. 
                There can be at most one at a bus.
    TC_GENUNIT: 	to get the handle for the next generating unit.
    TC_LOADUNIT: 	to get the handle for the next load unit.
    TC_SHUNTUNIT: 	to get the handle for the next shunt unit.
    TC_BRANCH: 	to get the handle for the next branch object.

    Set p_hnd to zero to get the first equipment handle. 

    Args:
        hndBus (c_int): bus handle
        type (c_int): object type code
        p_hnd (byref(c_int): object handle
    
    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetBusEquipment.argtypes = [c_int,c_int,c_void_p]
    return ASPENOlrx.OlrxAPIGetBusEquipment( hndBus, type, pHnd)

def OlrxAPIPostData(hnd):
    """Perform validation and update objct data in the network database 

        Changes to the equipment data made through SetData function will 
        not be committed to the program network database until after 
        this function has been executed with success.

    Args:
        hnd (c_int): Object handle

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIPostData.argtypes = [c_int]
    return ASPENOlrx.OlrxAPIPostData(hnd)

def OlrxAPIGetSCVoltage( hnd, vdOut1, vdOut2, style ):
    """Retrieves post-fault voltage of a bus, or of connected buses of 
    a line, transformer, switch or phase shifter.

    Args:
        hnd	(c_int): object handle
        vdOut1 (c_double*9): voltage result, real part or magnitude, 
                             at equipment terminals
        vdOut2 (c_double*9): voltage result, imaginary part or angle in degree
                             at equipment terminals
        style (c_int)     : voltage result style
                                1: output 012 sequence voltage in rectangular form
                                2: output 012 sequence voltage in polar form
                                2: output ABC phase voltage in rectangular form
                                4: output ABC phase voltage in polar form

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetSCVoltage.argtypes = [c_int, c_double*9, c_double*9, c_int]
    return ASPENOlrx.OlrxAPIGetSCVoltage( hnd, vdOut1, vdOut2, style )

def OlrxAPIGetSCCurrent( hnd, vdOut1, vdOut2, style ):
    """Retrieve post fault current for a generator, load, shunt, switched shunt, 
    generating unit, load unit, shunt unit, transmission line, transformer, 
    switch or phase shifter. 
    
    You can get the total fault current by calling this function with the 
    pre-defined handle of short circuit solution, HND_SC.

    Args:
        hnd	(c_int): object handle
        vdOut1 (c_double*12): current result, real part or magnitude, into 
                             equipment terminals
        vdOut2 (c_double*12): current result, imaginary part or angle in degree,
                             into equipment terminals
        style (c_int)      : current result style
                            1: output 012 sequence current in rectangular form
                            2: output 012 sequence current in polar form
                            3: output ABC phase current in rectangular form
                            4: output ABC phase current in polar form

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetSCCurrent.argtypes = [c_int, c_double*12, c_double*12, c_int]
    return ASPENOlrx.OlrxAPIGetSCCurrent( hnd, vdOut1, vdOut2, style )

def OlrxAPIGetRelay( hndRlyGroup, hndRelay ):
    """Retrieve handle of the next relay object in a relay group.

    Set hndRelay=0 to get the first relay.

    Args:
        hndRlyGroup (c_int): Relay group handle
        hndRelay (byref(c_int)): relay handle
    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetRelay.argtypes = [c_int, c_void_p]
    return ASPENOlrx.OlrxAPIGetRelay( hndRlyGroup, hndRelay );

def OlrxAPIGetRelayTime( hndRelay, mult, trip ):
    """Return operating time for a fuse, an overcurrent relay 
    (phase or ground), or a distance relay (phase or ground)
    in fault simulation result.
    
    Args:
        hndRelay (c_int): relay object handle
        mult (c_double) : relay current multiplying factor
        trip (byref(c_double)) : relay operating time in seconds

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIGetRelayTime.argtypes = [c_int, c_double, c_void_p]
    return ASPENOlrx.OlrxAPIGetRelayTime( hndRelay, mult, trip )

def OlrxAPIRun1LPFCommand(Params):
    """Run OneLiner command

    Args: 	
        Params (c_char_p): XML string, or full path name to XML file, containing a XML node
                        - Node name: OneLiner command
                        - Node attributes: required command parameters

            Command: CHECKPRIBACKCOORD  
            Attributes:
                REPORTPATHNAME	(*) Full pathname of report file.
                OUTFILETYPE 	[2] Output file type 1- TXT; 2- CSV
                SELECTEDOBJ	Relay group to check against its backup. Must  have one of following values
                “PICKED “ 	the highlighted relaygroup on the 1-line diagram
                “BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;”  location string of the relaygroup. Format description is in OneLiner help section 10.2.
                TIERS	[0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
                AREAS	[0-9999] Comma delimited list of area numbers and ranges to check relaygroups agains backup.
                ZONES	[0-9999] Comma delimited list of zone numbers and ranges to check relaygroups agains backup. This attribute is ignored if AREAS is found.
                KVS	0-999] Comma delimited list of KV levels and ranges to check relaygroups agains backup. This attribute is ignored if SELECTEDOBJ is found.
                TAGS	Comma delimited list of tags to check relaygroups agains backup. This attribute is ignored if SELECTEDOBJ is found.
                COORDTYPE	Coordination type to check. Must  have one of following values
                “0”	OC backup/OC primary (Classical)
                “1”	OC backup/OC primary (Multi-point)
                “2”	DS backup/OC primary
                “3”	OC backup/DS primary
                “4”	DS backup/DS primary
                “5”	OC backup/Recloser primary
                “6”	All types/All types
                LINEPERCENT	Percent interval for sliding intermediate faults. This attribute is ignored if COORDTYPE is 0 or 5.
                RUNINTERMEOP	1-true; 0-false. Check  intermediate faults with end-opened. This attribute is ignored if COORDTYPE is 0 or 5.
                RUNCLOSEIN	1-true; 0-false. Check close-in fault. This attribute is ignored if COORDTYPE is 0 or 5.
                RUNCLOSEINEOP	1-true; 0-false. Check close-in fault with end-opened. This attribute is ignored if COORDTYPE is 0 or 5.
                RUNLINEEND	1-true; 0-false. Check line-end fault. This attribute is ignored if COORDTYPE is 0 or 5.
                RUNREMOTEBUS	1-true; 0-false. Check remote bus fault. This attribute is ignored if COORDTYPE is 0 or 5.
                RELAYTYPE	Relay types to check: 1-Ground; 2-Phase; 3-Both.
                FAULTTYPE	Fault  types to check: 1-3LG; 2-2LG; 4-1LF; 8-LL; or sum of values for desired selection
                OUTPUTALL	1- Include all cases in report; 0- Include only flagged cases in report
                MINCTI	Lower limit of acceptable CTI range
                MAXCTI	Upper limit of acceptable CTI range
                OUTRLYPARAMS	Include relay settings in report: 0-None; 1-OC;2-DS;3-Both
                OUTAGELINES	Run line outage contingency: 0-False; 1-True
                OUTAGEXFMRS	Run transformer outage contingency: 0-False; 1-True
                OUTAGEMULINES	Run mutual line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGEMULINESGND		Run mutual line outage and grounded contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGE2LINES	Run double line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGE1LINE1XFMR	Run double line and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0 or OUTAGEXFMRS =0
                OUTAGE2XFMRS	Run double and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEXFMRS =0
                OUTAGE3SOURCES	Outage only  3 strongest sources: 0-False; 1-True. Ignored if OUTAGEMULINES=0 and OUTAGEXFMRS =0

            Command: CHECKRELAYOPERATIONSEA   
            Attributes:
                REPORTPATHNAME	(*) Full pathname of folder for report files.
                REPORTCOMMENT		Additional comment string to include in all checking report files
                SELECTEDOBJ	Check line with selected relaygroup. Must  have one of following values
                “PICKED “ 	the highlighted relaygroup on the 1-line diagram
                “BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;”  location string of  the relaygroup. Format description is in OneLiner help section 10.2.
                TIERS	[0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
                AREAS	[0-9999] Comma delimited list of area numbers and ranges.
                ZONES	[0-9999] Comma delimited list of zone numbers and ranges. This attribute is ignored if AREAS is found.
                KVS	[0-999] Comma delimited list of KV levels and ranges. This attribute is ignored if SELECTEDOBJ is found.
                TAGS	Comma delimited list of tags. This attribute is ignored if SELECTEDOBJ is found.
                DEVICETYPE	Space delimited list of relay type types to take into consideration in stepped-events: OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF
                FAULTTYPE	Space delimited list of fault  types to take into consideration in stepped-events: 1LF, 3LG
                OUTAGELINES	Run line outage contingency: 0-False; 1-True
                OUTAGEXFMRS	Run transformer outage contingency: 0-False; 1-True
                OUTAGEMULINES	Run mutual line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGEMULINESGND		Run mutual line outage and grounded contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGE2LINES	Run double line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
                OUTAGE1LINE1XFMR	Run double line and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0 or OUTAGEXFMRS =0
                OUTAGE2XFMRS	Run double and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEXFMRS =0
                OUTAGE3SOURCES	Outage only  3 strongest sources: 0-False; 1-True. Ignored if OUTAGEMULINES=0 and OUTAGEXFMRS =0

            Command: SETGENREFANGLE    
            Attributes:
                REPORTPATHNAME	Full pathname of folder for report files.
                REFERENCEGEN	Bus name and kV of reference generator in format: 'BNAME', KV.
                EQUSOURCEOPTION	Option for calculating reference angle of equivalent sources. Must have one of the following values
                [ROTATE] 	apply delta angle of existing reference gen
                SKIP   	Leave unchanged. This option will be in effect automatically when old reference is not valid
                ASGEN  	Use angle computed for regular generator

            Command: CHECKRELAYSETTINGS    
            Attributes:

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIRun1LPFCommand.argstype = [c_char_p]
    return ASPENOlrx.OlrxAPIRun1LPFCommand(Params)

def OlrxAPIDoSteppedEvent(hnd, fltOpt, runOpt, noTiers):
    """Run stepped-event simulation

    Remarks:
        After successful call of OlrxAPIDoSteppedEvent() you must call 
        function GetSteppedEvent() to retrieve detailed result of 
        each step in the simulation.

    Args: 	
        hnd (c_int): handle of a bus or a relay group.
        fltOpt (c_double*64): simulation options
            fltOpt(1) – Fault connection code
                            1=3LG
                            2=2LG BC, 3=2LG CA, 4=2LG AB
                            5=1LG A, 5=1LG B, 6=1LG C
                            7=LL BC, 7=LL CA, 8=LL AB
            fltOpt(2) – Intermediate percent between 0.01-99.99. 0 for a 
                        close-in fault. This parameter is ignored if nDevHnd 
                        is a bus handle.
            fltOpt(3) – Fault resistance, ohm 
            fltOpt(4) – Fault reactance, ohm
            fltOpt(4+1) – Zero or Fault connection code for additional user event 
            fltOpt(4+2) – Time  of additional user event, seconds.
            fltOpt(4+3) – Fault resistance in additional user event, ohm 
            fltOpt(4+4) – Fault reactance in additional user event, ohm
            fltOpt(4+5) – Zero or Fault connection code for additional user event 
        …
        runOpt (c_int*5): simulation options flags. 1 – set; 0 - reset
            runOpt(1)  - Consider OCGnd operations
            runOpt(2)  - Consider OCPh operations
            runOpt(3)  - Consider DSGnd operations
            runOpt(4)  - Consider DSPh operations
            runOpt(5)  - Consider Protection scheme operations

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIDoSteppedEvent.argstype = [c_int,c_double*64,c_int*5,c_int]
    return ASPENOlrx.OlrxAPIDoSteppedEvent(hnd, fltOpt, runOpt, noTiers)

def OlrxAPIGetSteppedEvent( step, timeStamp, fltCurrent, userDef, eventDesc, faultDest ):
    """Retrieve detailed result of a step in stepped-event simulation

    Remarks:
        Call this function with step = 0 to get total number of steps

    Args:
        step (c_int): Sequential index of the event in the simulation.
                      (1 is the initial user-defined event)
        timeStamp (byref(c_double)): Event time in seconds
        fltCurrent (byref(c_double)): Highest phase fault current magnitude 
                                      at this step
        userDef (byref(c_double)): User defined event flag. 1= true; 0= false
        eventDesc (c_char_p): Event description string that includes list 
                              of all devices that had tripped.
        faultDesc (c_char_p): Fault description string of the event.

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    OlrxAPIGetSteppedEvent.argstype = [c_int,c_void_p,c_void_p,c_void_p,c_char_p,c_char_p]
    return ASPENOlrx.OlrxAPIGetSteppedEvent( step, timeStamp, fltCurrent, userDef, eventDesc, faultDest )

def OlrxAPIDoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
    """Simulate one or more faults

    Args:
        hnd	(c_int): handle of a bus, branch or a relay group.
        vnFltConn (c_int*4): fault connection flags. 1 – set; 0 - reset
            vnFltConn[1] – 3PH 
            vnFltConn[2] – 2LG 
            vnFltConn[3] – 1LG 
            vnFltConn[4] – LL 
        fltOpt(c_double*15): fault options flags. 1 – set; 0 - reset
            vdFltOpt(1)  - Close-in
            vdFltOpt(2)  - Close-in w/ outage
            vdFltOpt(3)  - Close-in with end opened
            vdFltOpt(4)  - Close-in with end opened w/ outage
            vdFltOpt(5)  - Remote bus
            vdFltOpt(6)  - Remote bus w/ outage
            vdFltOpt(7)  - Line end
            vdFltOpt(8)  - Line end w/ outage
            vdFltOpt(9)  - Intermediate %
            vdFltOpt(10) - Intermediate % w/ outage
            vdFltOpt(11) - Intermediate % with end opened
            vdFltOpt(12) - Intermediate % with end opened w/ outage
            vdFltOpt(13) - Auto seq. Intermediate % from (*)
            vdFltOpt(14) - Auto seq. Intermediate % to (*)
            vdFltOpt(15) – Outage line grounding admittance in mho (***).
        vnOutageLst (c_int*100): list of handles of branches to be outaged; 0 terminated
        vnOutageOpt (c_int*4):  branch outage option flags. 1 – set; 0 - reset
            vnOutageOpt(1) - one at a time
            vnOutageOpt(2) - two at a time
            vnOutageOpt(3) - all at once
            vnOutageOpt(4) – breaker failure (**)
        dFltR	[in] fault resistance, in Ohm
        dFltX	[in] fault reactance, in Ohm
        nClearPrev	[in] clear previous result flag. 1 – set; 0 - reset

    Remarks:
        (*)	To simulate a single intermediate fault without auto-sequencing, 
            set both vdFltOpt(13)and vdFltOpt(14) to zero
        (**) Set this flag to 1 to simulate breaker open failure condition 
             that caused two lines that share a common breaker to be separated 
             from the bus while still connected to each other. TC_BRANCH handle 
             of the two lines must be included in the array vnOutageLst.

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIDoFault.argtypes = [c_int,c_int*4,c_double*15,c_int*4,c_int*100,c_double,c_double,c_int]
    return ASPENOlrx.OlrxAPIDoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev)

def OlrxAPIPickFault( index, tiers ):
    """Select a specific short circuit simulation case.

    This function must be called before any post fault voltage and current results 
    and relay time can be retrieved.
    
    Args:
        index (c_int): fault number or
                        SF_FIRST: first fault
                        SF_NEXT: next fault
                        SF_PREV: previous fault
                        SF_LAST: last available fault
        tiers (c_int): number of tiers around faulted bus to compute solution results 

    Returns:
        OLRXAPI_FAILURE: Failure
        OLRXAPI_OK     : Success
    """
    ASPENOlrx.OlrxAPIPickFault.argtypes = [c_int,c_int]
    return ASPENOlrx.OlrxAPIPickFault( index, tiers)

# These API functions return string
def OlrxAPIFaultDescription(index):
    """Retrieves description string of a fault in the simulation results.

    Call this function with index=0 to get description of the
    current fault simulation.

    Args:
        index (c_int): Index of fault simulation.

    Returns:
        string (c_char_p)
    """
    ASPENOlrx.OlrxAPIFaultDescription.restype = c_char_p
    ASPENOlrx.OlrxAPIFaultDescription.argstype = [c_int]
    return ASPENOlrx.OlrxAPIFaultDescription(index)

def OlrxAPIErrorString():
    """Retrieves error message string
    Returns:
        string (c_char_p)
    """
    ASPENOlrx.OlrxAPIErrorString.restype = c_char_p
    return ASPENOlrx.OlrxAPIErrorString()

def OlrxAPIFullBusName(hnd):
    """Return string composed of name and kV of the given bus

    Args:
        hnd (c_int): Bus handle

    Returns:
        string (c_char_p)
    """
    ASPENOlrx.OlrxAPIFullBusName.restype = c_char_p
    ASPENOlrx.OlrxAPIFullBusName.argstype = [c_int]
    return ASPENOlrx.OlrxAPIFullBusName(hnd)

def OlrxAPIFullBranchName(hnd):
    """Return a string composed of Bus, Bus2, Circuit ID and type
    of a branch object

    Args:
        hnd (c_int): Branch object handle

    Returns:
        string (c_char_p)
    """
    ASPENOlrx.OlrxAPIFullBranchName.restype = c_char_p
    ASPENOlrx.OlrxAPIFullBranchName.argstype = [c_int]
    return ASPENOlrx.OlrxAPIFullBranchName(hnd)

def OlrxAPIFullRelayName(hnd):
    """Return a string composed of relay type, name and branch location

    Args:
        hnd (c_int): relay object handle

    Returns:
        string (c_char_p)
    """
    ASPENOlrx.OlrxAPIFullRelayName.restype = c_char_p
    ASPENOlrx.OlrxAPIFullRelayName.argstype = [c_int]
    return ASPENOlrx.OlrxAPIFullRelayName(hnd)
