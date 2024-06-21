""" Python wrapper for ASPEN olxapi.dll

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced Systems for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "15.12.3"
__email__     = "support@aspeninc.com"
__status__    = "Release"

from ctypes import *
import sys,os.path
from OlxAPIConst import *
ASPENOlxAPIDLL = None
ASPENOLRFILE = ''
#
def InitOlxAPI(dllPath='',prt=True):
    """ Initialize OlxAPI session.
    Successfull initialization is required before any other OlxAPI call can be executed.

    Args:
        dllPath (string): Full path name ASPEN program folder
                          where olxapi.dll and related program components are located
                          if ="" => find automatic on disc C

    return:
        None

    Raises:
         OlxAPIException
    """
    global ASPENOlxAPIDLL,ASPENOLRFILE
    if ASPENOlxAPIDLL != None:
        return
    path = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    if dllPath=='' and os.path.isfile(os.path.join(path,'olxapi.dll')):
        olxapi = os.path.join(path,'olxapi.dll')
    else:
        olxapi = getASPENFile(dllPath,'olxapi.dll')
        path = os.path.dirname(olxapi)
    if path not in sys.path:
        os.environ['PATH'] = path + ";" + os.environ['PATH']
    # Attempt to copy hasp_rt.exe to the executable path.
    # (Required in network key applications)
    ph1 = os.path.join(os.path.dirname(sys.executable),"hasp_rt.exe")
    if not os.path.isfile(ph1):
        ph2 = os.path.join(path,"hasp_rt.exe")
        if os.path.isfile(ph2):
            try:
                from shutil import copyfile
                copyfile(ph2,ph1)
            except :
                #raise Exception("\nPermission denied: " +os.path.dirname(sys.executable))
                pass
    # Load the OlxAPI.dll
    ASPENOlxAPIDLL = WinDLL(olxapi , use_last_error=True)
    if ASPENOlxAPIDLL == None:
        raise OlxAPIException("Failed to setup olxapi.dll")
    # OlxAPI.dll is hardcoded to return ErrorString() of "No Error"
    # when and only when the session is successfully initialzed and active
    errorAPIInit = "OlxAPI Init Error"
    try:
        # Attempt to get error string
        errorAPIInit = ErrorString()
    except:
        pass
    if errorAPIInit != "No Error":
        raise OlxAPIException(errorAPIInit)
    #
    buf = create_string_buffer(b'\000' * 1028)
    ASPENOlxAPIDLL.OlxAPIVersionInfo(buf)
    vData = decode(buf.value).split("\n")
    v1 = vData[0].split()
    v2 = vData[1].split()
    Version = v1[2]
    BuildNumber = v1[4]
    DBXVersion = int(v2[2])
    if prt:
        print ( "Loaded: " + olxapi )
    if OLXAPI_DBX_VER>DBXVersion:
        raise OlxAPIException('\tDBX version mismatch: %i in OlxAPI.DLL; %i in pyOlxAPI\n'%(DBXVersion,OLXAPI_DBX_VER))
    if prt:
        print ( "OlxAPI Version:"+Version+" Build: "+BuildNumber )
#
def LoadDataFile(filePath,readonly,prt=True):
    """ Read OLR data file from disk

    Args:
        filePath (str) : Full path name of OLR file.
        readonly (int/bool) : open in read-only mode. 1-true; 0-false

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
        OLXAPI_DATAFILEANOMALIES: File opened with data errors

    Sample:
        if OLXAPI_FAILURE==OlxAPI.LoadDataFile("C:\\Program Files (x86)\\ASPEN\\1LPFv15\\SAMPLE30.OLR",1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

    Remarks:
        The OneLiner’s TTY output during the data file reading operation execution
        is atomatically saved in the log file PowerScriptTTYLog.txt in the
        Windows %TEMP% folder. When this funtions return value is other than OLXAPI_OK,
        you must check the content of this log file for details of data errors and
        warnings that were found when reading the file and proceed accordingly.
    """
    __checkInit__(0)
    ASPENOlxAPIDLL.OlxAPILoadDataFile.argtypes = [c_char_p,c_int]
    r = ASPENOlxAPIDLL.OlxAPILoadDataFile( encode3(filePath) , True if readonly else False)
    if prt and r==OLXAPI_OK:
        print("File opened successfully: " + filePath)
    if prt and r==OLXAPI_DATAFILEANOMALIES:
        print("File opened with data errors: " + filePath)
    global ASPENOLRFILE
    ASPENOLRFILE = GetOlrFileName()
    return r

#
def AddDevice(tc,brHnd,brEnd,tokens,params=None):
    """ Add a new protective device in a relay group on a line, switch or transformer branch

    Args:
        tc       [in] Device type code.
        brHnd    [in] Handle number of a line, transformer, branch or relay group
        brEnd    [in] Index number of the line or transformer terminal:
                     0- Bus1 side; 1- Bus2 side; 2- Bus3 side
        tokens   [in] Array of parameter codes
        params   [in] Array of ponters to parameter values

    return:
        hnd       handle number of the new object
        0         failure

    Samples:
        RLYGROUP:
            hndLine = c_int(0)
            if OLXAPI_FAILURE==OlxAPI.FindObj1LPF("[LINE] 4 'TENNESSEE' 132 kV-12 'TEXAS' 132 kV 1", hndLine ):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            tokens = [RG_dBreakerTime]
            params = [0.2]
            hndRlygrp = OlxAPI.AddDevice(TC_RLYGROUP,hndLine,0,tokens,params)

        RLYOC:
            tokens = [OG_sID,OG_dCT,OG_vdDTPickup,OG_vdDTDelay,OG_dInst,OG_nCTLocation]
            params = ['E51G',101   ,[4.0,4.01]   ,[1.05,1.04] ,801     ,1]
            hndOC  = OlxAPI.AddDevice(TC_RLYOCG,hndLine,0,tokens,params)

        RLYDSG:
            DSParams = "Z1MAG\t1.2\t"+"Z1ANG\t75\t"+"Z0MAG\t1.2\t"+"Z0ANG\t75\t"+"k0M1\t0.8\t"+"k0A1\t-82\t"+"E32\tAUTO\t"+"E21MG\t1\t"+"E21XG\tN\t"+"Z1MG\t1.8\t"
            tokens = [DG_sID,DG_sDSType ,DG_dCT,DG_dVT,DG_sParam]
            params = ['E21G','SEL411G__',100   ,500   ,DSParams]
            hndDS  = OlxAPI.AddDevice(TC_RLYDSG,hndLine,0,tokens,params)

        SCHEME:
            tokens = [LS_sID   ,LS_sScheme,LS_sEquation,LS_vsSignalName,LS_vnSignalType          ,LS_vdSignalVar]
            params = ['SCHEME1','CUSTOM'  ,'DS+OC'     ,['OC','DS']    ,[LVT_OC_TOC,LVT_DS_ZONE1],[hndOC,hndDS]]
            hndLS  = OlxAPI.AddDevice(TC_SCHEME,hndLine,0,tokens,params)
    """
    __checkInit__()
    if type(tokens).__name__.startswith('c_long_Array') and type(params).__name__.startswith('c_void_p_Array'):
        tc1 = tc if type(tc)==c_long else c_int(tc)
        brHnd1 = brHnd if type(brHnd)==c_long else c_int(brHnd)
        brEnd1 = brEnd if type(brEnd)==c_long else c_int(brEnd)
        ASPENOlxAPIDLL.OlxAPIAddEquipment.argtypes = [c_int,c_int,c_int,c_void_p,c_void_p]
        return ASPENOlxAPIDLL.OlxAPIAddDevice(tc1,brHnd1,brEnd1,tokens,params)
    try:
        tokens1,params1 = __getTokenParam__(tokens,params)
        return AddDevice(tc,brHnd,brEnd,tokens1,params1)
    except:
        return 0

#
def AddEquipment(tc,tokens,params=None):
    """ Add a new bus or a new piece of equipment to the network.

    Args:
        tc        [in] Equpipment type code
        tokens    [in] Array of parameter codes
        params    [in] Array of ponters to parameter values

    return:
        hnd       handle number of the new object
        0         failure

    Samples:
        BUS:
            tokens = [BUS_sName,BUS_dKVnominal,OBJ_sMemo,OBJ_sGUID]
            params = ["NewBus1",132.0         ,'memoBus','{9106257f-cd51-4ec8-ae8d-24451b972323}']
            hndBus1 = OlxAPI.AddEquipment(TC_BUS,tokens,params)

        LINE:
            hndBus2 = OlxAPI.FindBus( 'CLAYTOR', 132.0 )
            tokens = [LN_nBus1Hnd,LN_nBus2Hnd,LN_sID,OBJ_sGUID                               ,LN_dX,LN_dX0,LN_vdRating]
            params = [hndBus1    ,hndBus2    ,'2'   ,'{1d1bbca4-eed6-4c4f-9f40-c86c9906f841}',0.1  ,0.112 ,[11,12,13,14]]
            hndLine = OlxAPI.AddEquipment(TC_LINE,tokens,params)

        SVD:
            tokens = [SV_nBusHnd,SV_dB,SV_vnNoStep,SV_vdBinc,SV_vdB0inc]
            params = [hndBs2    ,5.0  ,[5,2]      ,[0.1,0.2],[0.3,0.6]]
            hndSVD = OlxAPI.AddEquipment(TC_SVD,tokens,params)
    """
    __checkInit__()
    if type(tokens).__name__.startswith('c_long_Array') and type(params).__name__.startswith('c_void_p_Array'):
        tc1 = tc if type(tc)==c_long else c_int(tc)
        ASPENOlxAPIDLL.OlxAPIAddEquipment.argtypes = [c_int,c_void_p,c_void_p]
        return ASPENOlxAPIDLL.OlxAPIAddEquipment(tc1,tokens,params)
    try:
        tokens1,params1 = __getTokenParam__(tokens,params)
        return AddEquipment(tc,tokens1,params1)
    except:
        return 0

#
def BoundaryEquivalent(EquFileName, BusList, FltOpt) :
    """ Create boundary equivalent network.

    Args:
        EquFileName [in] Path name of the boundary equivalent OLR file.

        BusList   [in]   Array of handles of buses to be retained in the equivalent.
                         The list must be terminated with value -1

        FltOpt    [in] study parameters
                    FltOpt[1]  - Per-unit elimination threshold
                    FltOpt[2]  - Keep existing equipment at retained buses( 1- set; 0- reset)
                    FltOpt[3]  - Keep all existing annotations (1- set; 0-reset)

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIBoundaryEquivalent.argtypes = [c_char_p,POINTER(c_int),c_double*3]
    return ASPENOlxAPIDLL.OlxAPIBoundaryEquivalent( encode3(EquFileName) , BusList, FltOpt)

#
def BuildNumber():
    """ OlxAPI engine build number
    """
    __checkInit__(False)
    buf = create_string_buffer(b'\000' * 1028)
    ASPENOlxAPIDLL.OlxAPIVersionInfo(buf)
    vData = decode(buf.value).split(" ")
    return int(vData[4])

#
def BusPicker( title, busList, opt ):
    """  Displays the bus selection dialog.
    return:
        n: Number of buses selected
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIBusPicker.argtypes = [c_char_p,POINTER(c_int),c_int]
    return ASPENOlxAPIDLL.OlxAPIBusPicker(title,busList,opt)

#
def CloseDataFile():
    """ Close the network data file that had been loaded previously with a call to OlxAPI.LoadDataFile().

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    global ASPENOLRFILE
    ASPENOLRFILE = ''
    return ASPENOlxAPIDLL.OlxAPICloseDataFile()

#
def ComputeRelayTime(hnd, curMag, curAng, vMag, vAng,vpreMag, vpreAng, opTime, opDevice):
    """ Computes operating time for a fuse, recloser, an overcurrent relay (phase or ground),
        or a distance relay (phase or ground) at given currents and voltages.

    Args:
        hnd (c_int): relay object handle
        curMag (c_double*5) [in] array of relay current magnitude in amperes of
                    phase A, B, C, and if applicable, Io currents in neutral
                    of transformer windings P and S.
        curAng (c_double*5):  [in] array of relay current angles in degrees
        vMag (c_double*3):    [in] array of relay voltage magnitude in, line to neutral, in kV
                    of phase A, B and C
        vdVAng (c_double*3): [in] array of relay voltage angle
        vpreMag (c_double): [in] relay pre-fault positive sequence voltage magnitude in kV, line to neutral
        dVpreAng (c_double): [in] relay pre-fault positive sequence voltage angle in degrees
        opTime (pointer(c_double)): [out] relay operating time in seconds
        opDevice (c_char*128): [out] relay operation code:
                            NOP  No operation
                            ZGn  Ground distance zone n  tripped
                            ZPn  Phase distance zone n  tripped
                            Ix    Overcurrent relay operating quantity: Ia, Ib, Ic, Io, I2, 3Io, 3I2

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIMakeOutageList.argtypes = [c_int,c_double*5,c_double*5,c_double*3,c_double*3,c_double,c_double,c_void_p,c_char*128]
    return ASPENOlxAPIDLL.OlxAPIComputeRelayTime(hnd, curMag, curAng, vMag, vAng,vpreMag, vpreAng, opTime, opDevice)

#
def CreateNetwork(baseMVA):
    """ Create a new network

    Args:
        dBaseMVA  (float) : Base MVA

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    Remarks:
        This operation will invalidate all existing object handles.
    """
    global ASPENOLRFILE
    ASPENOLRFILE = 'Untitled.OLR'
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPICreateNetwork.argtypes = [c_double]
    return ASPENOlxAPIDLL.OlxAPICreateNetwork(baseMVA)

#
def decode(s,codec='ANSI'):
    """ convert bytes => string. """
    if type(s)==bytes:
        try:
            return s.decode(codec)
        except:
            #https://docs.python.org/3/library/codecs.html#standard-encodings
            for c1 in ['ANSI','ASCII','iso-8859-1','iso-8859-2','UTF-8','UTF-16','UTF-32']:
                s1 = decode(s,c1)
                if s1:
                    return s1
            return ''
    return s

#
def DeleteEquipment(hnd):
    """ Delete network object.

    Args:
        hnd (c_int/int): Object handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIDeleteEquipment.argtypes = [c_int]
    return ASPENOlxAPIDLL.OlxAPIDeleteEquipment(hnd)

#
def DoBreakerRating(Scope, RatingThreshold, OutputOpt, OptionalReport, ReportTXT, ReportCSV, ConfigFile) :
    """ Run breaker rating study.

    Args:
        Scope -          [1]: 0-IEEE;1-IEC
                         [2]: 0-All;1-Area;2-Zone;3-Selected
                         [3]: Area or zone number
                         ...: or list of bus hnd terminated with -1
        Scope[1] – Breaker rating standard: 0-ANSI/IEEE; 1-IEC.
        Scope[2] – Bus selection: 0-All buses; 1-in Area; 2-in Zone; 3- selected buses
        Scope[3] – Selected area or zone number.
        Scope[4] or list of handle number of selected buses. The last element in the list
                     must be -1.
        RatingThreshold [in] Percent rating threshold.
        OutputOpt  [in] Rating output option:
                        0- Output only overduty cases;
                        1- Output all cases;
                        OR Floating number S (0 < S < 1) -
                                 Check only breakers at buses where ratio
                                 “Bus fault current / Breaker rating” exceeds S.
        OptionalReport  [in] Integer number flag. Enable various bits to enable
                        optional sections in rating report: Bit 1- Detailed fault
                        simulation result; Bit 2- Breaker name plate data;
                        Bit 3- List of connected equipment.
        ReportTXT  [in] Full path name of text report file. Set to emty to omit text report.
        ReportCSV  [in] Full path name of CSV report file. Set to emty to omit CSV report.
        ConfigFile [in] Full path name of  breaker rating configuration file to apply in
                        this study. Set to emty to omit reading configuration file.

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIDoBreakerRating.argtypes = [POINTER(c_int),c_double,c_double,c_int,c_char_p,c_char_p,c_char_p]
    return ASPENOlxAPIDLL.OlxAPIDoBreakerRating(Scope, RatingThreshold, OutputOpt, OptionalReport,
                            encode3(ReportTXT), encode3(ReportCSV), encode3(ConfigFile))

#
def DoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
    """ Simulate one or more faults.

    Args:
        hnd    (c_int): handle of a bus, branch or a relay group.
        vnFltConn (c_int*4): fault connection flags. 1 - set; 0 - reset
            vnFltConn[1] - 3PH
            vnFltConn[2] - 2LG
            vnFltConn[3] - 1LG
            vnFltConn[4] - LL
        fltOpt(c_double*15): fault options flags. 1 - set; 0 - reset
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
            vdFltOpt(15) - Outage line grounding admittance in mho (***).
        vnOutageLst (c_int*100): list of handles of branches to be outaged; 0 terminated
        vnOutageOpt (c_int*4):  branch outage option flags. 1 - set; 0 - reset
            vnOutageOpt(1) - one at a time
            vnOutageOpt(2) - two at a time
            vnOutageOpt(3) - all at once
            vnOutageOpt(4) - breaker failure (**)
        dFltR (c_double): fault resistance, in Ohm
        dFltX (c_double)" fault reactance, in Ohm
        nClearPrev (c_int): clear previous result flag. 1 - set; 0 - reset

    Remarks:
        (*)    To simulate a single intermediate fault without auto-sequencing,
            set both vdFltOpt(13)and vdFltOpt(14) to zero
        (**) Set this flag to 1 to simulate breaker open failure condition
             that caused two lines that share a common breaker to be separated
             from the bus while still connected to each other. TC_BRANCH handle
             of the two lines must be included in the array vnOutageLst.

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIDoFault.argtypes = [c_int,c_int*4,c_double*15,c_int*4,c_int*100,c_double,c_double,c_int]
    return ASPENOlxAPIDLL.OlxAPIDoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev)

#
def DoSteppedEvent(hnd, fltOpt, runOpt, noTiers):
    """ Run stepped-event simulation.

    Remarks:
        After successful call to OlxAPIDoSteppedEvent() you must call
        the function GetSteppedEvent() to retrieve detailed results of
        each step in the simulation.
        To retrieve the fault simulation results at a network element
        in each event, call PickFault() with the event index then use
        GetSCCurrent(), GetSCVoltage() and GetRelayOperation() functions
        accordingly.

    Args:
        hnd (c_int): handle of a bus or a relay group.
        fltOpt (c_double*64): simulation options
            fltOpt(1) - Fault connection code
                            1=3LG
                            2=2LG BC, 3=2LG CA, 4=2LG AB
                            5=1LG A, 5=1LG B, 6=1LG C
                            7=LL BC, 7=LL CA, 8=LL AB
            fltOpt(2) - Intermediate percent between 0.01-99.99. 0 for a
                        close-in fault. This parameter is ignored if nDevHnd
                        is a bus handle.
            fltOpt(3) - Fault resistance, ohm
            fltOpt(4) - Fault reactance, ohm
            fltOpt(4+1) - Zero or Fault connection code for additional user event
            fltOpt(4+2) - Time  of additional user event, seconds.
            fltOpt(4+3) - Fault resistance in additional user event, ohm
            fltOpt(4+4) - Fault reactance in additional user event, ohm
            fltOpt(4+5) - Zero or Fault connection code for additional user event
        ...
        runOpt (c_int*7): simulation options flags. 1 - set; 0 - reset
            runOpt(1)  - Consider OCGnd operations
            runOpt(2)  - Consider OCPh operations
            runOpt(3)  - Consider DSGnd operations
            runOpt(4)  - Consider DSPh operations
            runOpt(5)  - Consider Protection scheme operations
            runOpt(6)  - Consider Voltage relay operations
            runOpt(7)  - Consider Differential relay operations
        nTiers: Study extent. Take into account protective devices located within this number of tiers only.

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIDoSteppedEvent.argstype = [c_int,c_double*64,c_int*7,c_int]
    return ASPENOlxAPIDLL.OlxAPIDoSteppedEvent(hnd, fltOpt, runOpt, noTiers)

#
def encode3(s,codec='ANSI'):
    """ convert string => bytes. """
    if type(s)==str:
        return s.encode(codec)
    return s

#
def EquipmentType(hnd):
    """ Get the equipment type of the given handle.

    Args:
        hnd (c_int/int): Object handle

    return:
        type      type of the given handle
        0         failure

    Sample:
        hnd = OlxAPI.FindObj1LPF("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")
        typ1 = OlxAPI.EquipmentType(busHnd) # =TC_LINE
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIEquipmentType.argstype = [c_int]
    return ASPENOlxAPIDLL.OlxAPIEquipmentType(hnd)

#
def ErrorString():
    """ Retrieves error message string.

    return:
        string (c_char_p)
    """
    ASPENOlxAPIDLL.OlxAPIErrorString.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIErrorString() )

#
def FaultDescriptionEx(index,flag):
    """ Retrieves description and fltspec strings of a fault in the simulation results buffer.
        Call this function with index=0 to get description of the current fault simulation.

    Args:
        index (c_int): Index of fault simulation.
        flag  (c_int): Output flag
               [0] Fautl description string only
               1   Plus FLTSPCVS on the last line.
                   See SIMULATEFAULT for details of FLTSPCVS string
               2   Plus FLTOPTSTR on the last line: a string containing the following
                   fault option data fields, separated by blank space:
                   PrefaultV: Prefault voltage flag
                     0 - From linear network solution
                     1 - Flat
                     2 - From load flow
                   FlatBusV: Flat prefault bus voltage
                   GenZType: Generator reactance flag
                     0 - subtransient
                     1 - transient
                     2 - synchronous
                   IgnoreFlag: a bit array
                     bit 1 - ignore load
                     bit 2 - ignore positive sequence shunt
                     bit 3 - ignore line G and B
                     bit 4 - ignore transformer B
                   GenILimitOption: generator limit option
                     0 - Ignore
                     1 - Enforce limit 1
                     2 - Enforce limit 2
                   SimulateCCGen: Simulate VCCS flag 1-true; 0-false
                   IgnoreMOV: MOV protected series capacitor iterative
                               solution flag 1-true; 0-false
                   AccelFactor: MOV protected series capacitor iterative
                               solution acceleration factor
                   MuThreshold: Ignore mutual coupling threshold
                   FaultMVAstyle: Fault MVA calculatin method
                   SimulateGenW3: Simulate type-3 wind generator flag 1-true; 0-false
                   SimulateGenW4: Simulate CIR flag 1-true; 0-false

    return:
        string (c_char_p)
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIFaultDescriptionEx.restype = c_char_p
    ASPENOlxAPIDLL.OlxAPIFaultDescriptionEx.argstype = [c_int,c_int]
    return decode (ASPENOlxAPIDLL.OlxAPIFaultDescriptionEx(index,flag) )

#
def FaultDescription(index):
    """ Retrieves description and fltspec strings of a fault in the simulation results buffer.
        Call this function with index=0 to get description of the current fault simulation.

    Args:
        index (c_int): Index of fault simulation.

    return:
        string (c_char_p)
    """
    return FaultDescriptionEx(index, 0)

#
def FindBus(name, kv):
    """ Find handle of bus with given name and kV.

    Args:
        name (str):    Bus name
        kv (float):    Bus nominal kV

    return:
        OLXAPI_FAILURE: Failure
        hnd     : bus handle

    Sample:
        busHnd = OlxAPI.FindBus('claytor', 132)
    """
    __checkInit__()
    kv = kv if type(kv)==c_double else c_double(kv)
    ASPENOlxAPIDLL.OlxAPIFindBus.argtypes = [c_char_p,c_double]
    return ASPENOlxAPIDLL.OlxAPIFindBus( encode3(name) , kv)

#
def FindEquipmentByTag(tags, devType, hnd=None):
    """ Find handle of object of devType that has the given tags.

    Args:
        tags (c_char_p): Tag string
        tags (c_int/int):    Object type: TC_BUS, TC_LOAD, TC_SHUNT, TC_GEN , TC_SVD,
                         TC_LINE, TC_XFMR, TC_XFMR3, TC_PS, TC_SCAP, TC_MU,
                         TC_RLYGROUP, TC_RLYOCG, TC_RLYOCP, TC_RLYDSG, TC_RLYDSP,
                         TC_FUSE,   TC_SWITCH, TC_RECLSRP, TC_RECLSRG or 0 (all object types).
        hnd (c_int):  Object handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
        [hnd]         : object handle in list (with hnd=None)

    Remarks:
        To get the first object in the list, call this function with *hnd equal 0.
        Call this function with devType equal 0 to search all object types.
        To get all object in list, call this function with hnd=None
    """
    if hnd==None:
        res = []
        equHnd = (c_int*1)(0)
        while OLXAPI_OK == FindEquipmentByTag(tags, devType, equHnd):
            res.append(equHnd[0])
        return res
    #
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIFindEquipmentByTag.argtypes = [c_char_p,c_int,POINTER(c_int)]
    return ASPENOlxAPIDLL.OlxAPIFindEquipmentByTag( encode3(tags) , c_int(devType), hnd)

#
def FindBusNo(no):
    """ Find handle of bus with given number.

    Args:
        no (c_int/int):    Bus number (must non zero)

    return:
        0  : Failure
        hnd: bus handle

    Sample:
        busHnd = OlxAPI.FindBusNo(2)
    """
    __checkInit__()
    no = c_int(no)
    ASPENOlxAPIDLL.OlxAPIFindBusNo.argtypes = [c_int]
    return ASPENOlxAPIDLL.OlxAPIFindBusNo(no)

#
def FindObj1LPF(obj1LPFStr, hnd=None):
    """ Returns handle number of the object in obj1LPFStr string.

    Args:
      obj1LPFStr  *(str)  : Object ID string produced by PrintObj1LPF() or the object GUID
      (hnd)       (c_int) : Object handle number.

    return:
        if hnd != None:
            OLXAPI_FAILURE: Failure
            OLXAPI_OK     : Success

        if hnd==None
            (int) Object handle number (=0 if not found)

    Samples:
        hnd = OlxAPI.FindObj1LPF("[BUS] 6 'NEVADA' 132 kV")                # by PrintObj1LPF()
        hnd = OlxAPI.FindObj1LPF("{d6a73b2c-346c-42f8-97ee-60f9a818c20c}") # by GUID

        or (with hnd!=None):
        hnd = c_int(0)
        OlxAPI.FindObj1LPF("[BUS] 6 'NEVADA' 1323 kV",hnd)
    """
    __checkInit__()
    if hnd is None:
        c_hnd = c_int(0)
        if OLXAPI_OK==FindObj1LPF(obj1LPFStr, c_hnd):
            return c_hnd.value
        return OLXAPI_FAILURE
    ASPENOlxAPIDLL.OlxAPIFindObj1LPF.argtypes = [c_char_p, POINTER(c_int)]
    return ASPENOlxAPIDLL.OlxAPIFindObj1LPF(encode3(obj1LPFStr),hnd)

#
def FullBusName(hnd):
    """ Return string composed of name and kV of the given bus.

    Args:
        hnd (c_int/int): Bus handle

    return:
        string

    Sample:
        busHnd = OlxAPI.FindBusNo(2)
        print( OlxAPI.FullBusName(busHnd) ) # "2 CLAYTOR 132.kV"
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIFullBusName.restype = c_char_p
    ASPENOlxAPIDLL.OlxAPIFullBusName.argstype = [c_int]
    return decode( ASPENOlxAPIDLL.OlxAPIFullBusName(hnd) )

#
def FullBranchName(hnd):
    """ Get a string composed of Bus, Bus2, Circuit ID and type of a branch object.

    Args:
        hnd (c_int): Branch object handle

    return:
        string (c_char_p)
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIFullBranchName.restype = c_char_p
    ASPENOlxAPIDLL.OlxAPIFullBranchName.argstype = [c_int]
    return decode( ASPENOlxAPIDLL.OlxAPIFullBranchName(hnd) )

#
def FullRelayName(hnd):
    """ Get a string composed of relay type, name and branch location.

    Args:
        hnd (c_int): relay object handle

    return:
        string (c_char_p)
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIFullRelayName.restype = c_char_p
    ASPENOlxAPIDLL.OlxAPIFullRelayName.argstype = [c_int]
    return decode( ASPENOlxAPIDLL.OlxAPIFullRelayName(hnd) )

#
def GetAreaName(no):
    """ Retrieve name of the given area number.

    Args:
        no (c_int): Area number

    return:
        name (c_char_p): Name string

    Remarks:
        if the funtion fails to execute the return string will consist of
        error message that begins with the key words: "GetAreaName failure:..."
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetAreaName.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetAreaName.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetAreaName(no) )

#
def getASPENFile(path,sfile):
    """ Get ASPEN file path (Exception if not found)

    Args:
        path : path search first
        sfile: name of file

    return:
        abs path of file
    """
    if path == "":
        if not os.path.isfile( os.path.join(path,sfile) ) :
            path = "C:\\Program Files (x86)\\ASPEN\\1LPFv15"
        if not os.path.isfile( os.path.join(path,sfile) ) :
            path = "C:\\Program Files\\ASPEN\\1LPFv15"
        if not os.path.isfile( os.path.join(path,sfile) ) :
            path = "C:\\Program Files (x86)\\ASPEN\\1LPFv15_Training"
        sf = os.path.join(path,sfile)
    else:
        sf = path
        if not os.path.isfile( sf ):
            sf = os.path.join(path,sfile)
    if os.path.isfile(sf) :
        return os.path.abspath(sf)
    else:
        raise OlxAPIException("Failed to locate: " + sf )

#
def GetBusEquipment(hndBus, TC_type, pHnd=None):
    """ Retrieves the handle of the next equipment of a given type that is attached to a bus.

    Args:
        hndBus (c_int/int)         : bus handle
        TC_type (c_int/int)        : object type code
        p_hnd (byref(c_int) /c_int): object handle

        The equipment TC_type can be one of the following:
        TC_GEN:     to get the handle for the generator.
                    There can be at most one at a bus.
        TC_LOAD:     to get the handle for the load.
                    There can be at most one at a bus.
        TC_SHUNT:    to get the handle for the shunt.
                    There can be at most one at a bus.
        TC_SVD:        to get the handle for the switched shunt.
                    There can be at most one at a bus.
        TC_GENUNIT:     to get the handle for the next generating unit.
        TC_LOADUNIT:     to get the handle for the next load unit.
        TC_SHUNTUNIT:     to get the handle for the next shunt unit.
        TC_BRANCH:     to get the handle for the next branch object.
        Set p_hnd to zero to get the first equipment handle.
        Set pHnd=None to get all equipment in list

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    Sample:
        hnd  = OlxAPI.FindBus('Claytor', 132 )
        branchHnd = c_int(0)
        while OLXAPI_OK == OlxAPI.GetBusEquipment( hnd, TC_BRANCH, branchHnd) :
            typeCode = OlxAPI.GetObjData(branchHnd,BR_nType)

        ba = OlxAPI.GetBusEquipment( hnd, TC_BRANCH) # all branch in list
    """
    if pHnd==None:
        res = []
        hnd = c_int(0)
        while OLXAPI_OK == GetBusEquipment(hndBus, TC_type,hnd):
            res.append(hnd.value)
        return res
    __checkInit__()
    pHnd1 = byref(pHnd) if type(pHnd)==c_long else pHnd
    ASPENOlxAPIDLL.OlxAPIGetBusEquipment.argtypes = [c_int,c_int,c_void_p]
    return ASPENOlxAPIDLL.OlxAPIGetBusEquipment( hndBus, TC_type, pHnd1)

#
def GetData(hnd, token, dataBuf):
    """ Get object data field. (try with GetObjData(...) )

    Args:
        hnd (c_int):        Object handle
        token (c_int):      field token
        dataBuf (c_void_p): data buffer

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetData.argtypes = [c_int,c_int,c_void_p]
    return ASPENOlxAPIDLL.OlxAPIGetData(hnd, token, dataBuf)


#
def GetEquipment(otype, p_hnd=None):
    """ Retrieves handle of the next equipment of given type in the system.
        This function will return the handle of all the objects of the given type,
        one by one, in the order they are stored in the OLR file
        Set p_hnd to 0 to retrieve the first object
        Set p_hnd=None to retrieve all in list

    Args:
        otype (c_int/int)           : Object type
        p_hnd (pointer(c_int)/c_int): Object handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    Sample:
        # One by one
        bsHnd = c_int(0)
        while OLXAPI_OK == OlxAPI.GetEquipment(TC_BUS, bsHnd):
            print('busName: ', OlxAPI.GetObjData(bsHnd,BUS_sName) )

        # All in list
        ba = OlxAPI.GetEquipment(TC_BUS)
    """
    if p_hnd==None:
        res = []
        hnd = c_int(0)
        while OLXAPI_OK == GetEquipment(otype, hnd):
            res.append(hnd.value)
        return res
    #
    __checkInit__()
    p_hnd1 = pointer(p_hnd) if type(p_hnd)==c_long else p_hnd
    ASPENOlxAPIDLL.OlxAPIGetEquipment.argstype = [c_int,c_void_p]
    return ASPENOlxAPIDLL.OlxAPIGetEquipment(otype, p_hnd1)

#
def GetLogicScheme( hndRlyGroup, hndScheme=None ):
    """ Retrieve handle of the next logic scheme object in a relay group.
        Set hndScheme=0 to get the first scheme.

    Args:
        hndRlyGroup (c_int): Relay group handle
        hndScheme (byref(c_int)): Logic scheme handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    if hndScheme==None:
        res = []
        hnd = c_int(0)
        while OLXAPI_OK == GetLogicScheme(hndRlyGroup,hnd):
            res.append(hnd.value)
        return res
    #
    __checkInit__()
    hndScheme1 = byref(hndScheme) if type(hndScheme)==c_long else hndScheme
    ASPENOlxAPIDLL.OlxAPIGetLogicScheme.argtypes = [c_int, c_void_p]
    return ASPENOlxAPIDLL.OlxAPIGetLogicScheme( hndRlyGroup, hndScheme1 );

#
def GetObjData(hnd,token):
    """ Get network object data field value

    Args:
        hnd (c_int/int)        : Object handle
        token (c_int/int/[int]): Object field token

    return:
        data : int / []*int / float / []*float / str

    Raises:
        OlxAPIException

    Samples:
        BUS:
            busName = OlxAPI.GetObjData(hndBus,BUS_sName)
            busKV = OlxAPI.GetObjData(hndBus,BUS_dKVnominal)
            print('busName, busKV :', busName, busKV)
            val = OlxAPI.GetObjData(hndBus,[OBJ_sGUID,OBJ_sTags,OBJ_sMemo,OBJ_sUDF])
            print('GUID,TAG,MEMO,UDF:',val[0],val[1],val[2],val[3])

        GENUNIT:
            X = OlxAPI.GetObjData(hndUnit,GU_vdX)           # X of genunit
            print('\tX   = ' + str(X))
    """
    if type(token) in {tuple,list}:
        return [GetObjData(hnd,t1) for t1 in token]
    #
    c_token = c_int(token) if type(token)==int else token
    vt = c_token.value//100
    #
    if c_token.value==OBJ_sGUID:
        return GetObjGUID(hnd)
    if c_token.value==OBJ_sTags:
        return GetObjTags(hnd)
    if c_token.value==OBJ_sMemo:
        return GetObjMemo(hnd)
    if c_token.value==OBJ_sUDF:
        fname = create_string_buffer(b'\000' * MXUDFNAME)
        fval = create_string_buffer(b'\000' * MXUDF)
        res,fidx = dict(),0
        while OLXAPI_OK == GetObjUDFByIndex(hnd,fidx,fname,fval):
            fidx+=1
            res[decode(fname.value)] = decode(fval.value)
        return res
    #
    c_hnd = c_int(hnd) if type(hnd)==int else hnd
    typ1 = EquipmentType(hnd)
    #
    if vt == VT_DOUBLE:
        dataBuf = c_double(0)
    elif vt == VT_INTEGER:
        dataBuf = c_int(0)
    else:
        dataBuf = create_string_buffer(b'\000' * 10*1024)
    #
    if (c_token.value==LN_nMuPairHnd and TC_LINE==typ1 ):
        res = []
        while OLXAPI_OK == GetData(c_hnd,c_token,byref(dataBuf)):
            res.append( __ProcessGetDataBuf__(dataBuf,c_token.value,c_hnd) )
        return res
    if typ1==TC_RLYDSP and c_token.value == DP_sParam:
        la = GetObjData(hnd,DP_vParamLabels)
        va = GetObjData(hnd,DP_vParams)
        res = ''
        for i in range(len(la)):
            res+=la[i]+'\t'+va[i]+'\t'
        return res
    #
    if typ1==TC_RLYDSG and c_token.value == DG_sParam:
        la = GetObjData(hnd,DG_vParamLabels)
        va = GetObjData(hnd,DG_vParams)
        res = ''
        for i in range(len(la)):
            res+=la[i]+'\t'+va[i]+'\t'
        return res

    #
    ret = GetData(c_hnd, c_token, byref(dataBuf))
    if ret==OLXAPI_FAILURE:
        if typ1==TC_LINE and c_token.value in {LN_nRlyGr1Hnd,LN_nRlyGr2Hnd}:
            return 0
        if typ1==TC_PS and c_token.value in {PS_nRlyGr1Hnd,PS_nRlyGr2Hnd}:
            return 0
        if typ1==TC_XFMR and c_token.value in {XR_nRlyGr1Hnd,XR_nRlyGr2Hnd,XR_nLTCCtrlBusHnd}:
            return 0
        if typ1==TC_XFMR3 and c_token.value in {X3_nRlyGr1Hnd,X3_nRlyGr2Hnd,X3_nRlyGr3Hnd,X3_nLTCCtrlBusHnd}:
            return 0
        if typ1==TC_SWITCH and c_token.value in {SW_nRlyGrHnd1,SW_nRlyGrHnd2}:
            return 0
        if typ1==TC_SCAP and c_token.value in {SC_nRlyGr1Hnd,SC_nRlyGr2Hnd}:
            return 0
        if typ1==TC_BRANCH and c_token.value in {BR_nRlyGrp1Hnd,BR_nRlyGrp2Hnd,BR_nRlyGrp3Hnd}:
            return 0
        if typ1==TC_RLYGROUP and c_token.value in {RG_nPrimaryHnd,RG_nBackupHnd,RG_nTripLogicHnd,RG_nReclLogicHnd}:
            return 0
        raise OlxAPIException(ErrorString())
    return __ProcessGetDataBuf__(dataBuf,c_token.value,c_hnd)

#
def GetObjGraphicData(hnd, buf):
    """ Retrieve object graphic data.

    Args:
        hnd (c_int): object handle or HND_SYS
        buf (c_int*500): output buffer for graphic data

    return:
        OLXAPI_FAILURE: Failure
        Size          : Success, size of graphic data record placed in the buffer

    Remarks:
        Object graphic data record structure
        TC_BUS: size,angle,x,y,nameX,nameX,hideID
                !!! when size is Zero, the bus had never been placed on the 1-line and
                      therefore the rest of the data are just junk
                    when size is negative, the bus is not displayed on the 1-line (hidden)
        TC_GEN, TC_GENW3, TC_GENW4, TC_CCGEN, TC_LOAD, TC_SHUNT, TC_SVD:
                x|y,angle,textX,textY
        TC_XFMR, TC_PS, TC_SCAP, TC_SWITCH, TC_DCLINE2:
                noSegs,p1X,p1Y,p2X,p2Y,p3X,p3Y,p4X,p4Y,text1X,text1Y,text2X,text2Y
                !!! noSegs determines if p3 and p4 are being used
        TC_XFMR3: noSegs,p1X,p1Y,p2X,p2Y,p3X,p3Y,p4X,p4Y,p5X,p5Y,p6X,p6Y,
                  text1X,text1Y,text2X,text2Y,text3X,text3Y
                !!! noSegs determines if p3 and p4 are being used
        TC_LINE, TC_SCAP: noSegs,p1X,p2X....,text1X,text1Y,text2X,text2Y
        TC_SYS: windowCenterX, windowCenterY, fontSizeMainWindow, fontSizeOCWindow, fontSizeDSWindow,
                xfmrStyle, colorkV1,....,colorkV13,colorIndex1,..., colorIndex13
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjGraphicData.argstype = [c_int,POINTER(c_int)]
    ASPENOlxAPIDLL.OlxAPIGetObjGraphicData.restype = c_int
    return ASPENOlxAPIDLL.OlxAPIGetObjGraphicData(hnd,buf)

#
def GetObjGUID(hnd):
    """ Retrieve name of the given area number.

    Args:
        hnd (c_int)  [in] Object handle

    return:
        GUID (c_char_p): GUID string

    Remarks:
        if the funtion fails to execute the return string will consist of
        error message that begins with the key words: "GetObjGUID failure:..."
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjGUID.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetObjGUID.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetObjGUID(hnd) )

#
def GetObjUDFByIndex(hnd,fidx,fname,fval):
    """ Retrieve user-defined field of an object.

    Args:
        hnd (c_int)             [in] Object handle
        nFIdx (c_int)           [in] Field index number (zero based)
        fname (c_char*MXUDF)    [out] Buffer for field name. Must be MXUDFNAME long. Set to NULL if not needed.
        fval (c_char*MXUDFNAME) [out] Buffer for field value. Must be MXUDF long. Set to NULL if not needed.

    return:
      OLXAPI_OK     : Success
      OLXAPI_FAILURE: Field with the given index does not exist
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjUDFByIndex.argtypes = [c_int,c_int,c_char*(MXUDFNAME+1),c_char*(MXUDF+1)]
    return ASPENOlxAPIDLL.OlxAPIGetObjUDFByIndex(hnd,fidx,fname,fval)

#
def GetObjUDF(hnd, fname, fval):
    """ Retrieve the value of a user-defined field.

    Args:
        hnd (c_int)             [in] Object handle
        fname (c_char_p)        [in] Field name
        fval (c_char*MXUDFNAME) [out] Buffer for field value. Must be MXUDF long

    return:
        OLXAPI_OK     : Success
        OLXAPI_FAILURE: Object does not have UDF Field with the given name
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjUDF.argtypes = [c_int,c_char_p,c_char*(MXUDF+1)]
    return ASPENOlxAPIDLL.OlxAPIGetObjUDF(hnd,encode3(fname),fval)

#
def GetObjJournalRecord(hnd):
    """ Retrieve journal jounal record details of a data object in the OLR file.

    Args:
        hnd (c_int): object handle

    return:
        JRec (c_char_p): String of journal record fields, separated by new line character:
            -    Create date and time
            -    Created by
            -    Last modified date and time
            -    Modified by
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjJournalRecord.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetObjJournalRecord.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetObjJournalRecord(hnd) )

#
def GetObjTags(hnd):
    """ Retrieve tag string for a bus, generator, load, shunt, switched shunt,
        transmission line, transformer, switch, phase shifter, distance relay,
        overcurrent relay, fuse, recloser, relay group.

    Args:
        hnd (c_int): object handle

    return:
        tags (c_char_p): Tag string

    Remarks:
        if the funtion fails to execute the return string will consist of
        error message that begins with the key words: "GetObjTags failure:..."
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjTags.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetObjTags.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetObjTags(hnd) )

#
def GetObjMemo(hnd):
    """ Retrieve memo string for a bus, generator, load, shunt, switched shunt,
        transmission line, transformer, switch, phase shifter, distance relay,
        overcurrent relay, fuse, recloser, relay group.

    Args:
        hnd (c_int): object handle

    return:
        memo (c_char_p): Memo string

    Remarks:
        if the funtion fails to execute the return string will consist of
        error message that begins with the key words: "GetObjTags failure:..."
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetObjMemo.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetObjMemo.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetObjMemo(hnd) )

#
def GetOlrFileName():
    """ Full path name of current OLR file.

    return:
        string : Full path name of current OLR file
        ''     : if not found (no OLR file is loaded)
    """
    __checkInit__(0)
    ASPENOlxAPIDLL.OlxAPIGetOlrFileName.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetOlrFileName() )

#
def GetPSCVoltage( hnd, vdOut1, vdOut2, style ):
    """ Retrieves pre-fault voltage of a bus, or of connected buses of
        a line, transformer, switch or phase shifter.

    Args:
        hnd    (c_int): object handle
        vdOut1 (c_double*3): voltage result, real part or magnitude,
                             at equipment terminals
        vdOut2 (c_double*3): voltage result, imaginary part or angle in degree
                             at equipment terminals
        style (c_int)     : voltage result style
                             1: output voltage in kV
                             2: output voltage in per-unit

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetPSCVoltage.argtypes = [c_int, c_double*3, c_double*3, c_int]
    return ASPENOlxAPIDLL.OlxAPIGetPSCVoltage( hnd, vdOut1, vdOut2, style )

#
def GetRelay( hndRlyGroup, hndRelay ):
    """ Retrieve handle of the next relay object in a relay group.
        Set hndRelay=0 to get the first relay.

    Args:
        hndRlyGroup (c_int): Relay group handle
        hndRelay (byref(c_int)): relay handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetRelay.argtypes = [c_int, c_void_p]
    return ASPENOlxAPIDLL.OlxAPIGetRelay( hndRlyGroup, hndRelay );

#
def GetRelayTime( hndRelay, mult, consider_signalonly, trip, device ):
    """ Retrieve operating time for a fuse, an overcurrent relay (phase or ground),
        recloser (phase or ground), or a distance relay (phase or ground)
        in fault simulation result.

    Args:
        hndRelay (c_int): [in] relay object handle
        mult (c_double) : [in] relay current multiplying factor
        consider_signalonly: [in] Consider relay element signal-only flag 1 - Yes; 0 - No
        trip (byref(c_double)) : [out] relay operating time in seconds
        device (c_char_p): [out] relay operation code. Required buffer size 128 bytes.
                     NOP  No operation.
                     ZGn  Ground distance zone n tripped.
                     ZPn  Phase distance zone n tripped.
                     TOC=value Time overcurrent element operating quantity in secondary amps
                     IOC=value Instantaneous overcurrent element operating quantity in secondary amps

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetRelayTime.argtypes = [c_int, c_double, c_void_p, c_char_p, c_int]
    return ASPENOlxAPIDLL.OlxAPIGetRelayTime( hndRelay, mult, trip, device, consider_signalonly )

#
def GetSCVoltage( hnd, vdOut1, vdOut2, style ):
    """ Retrieves post-fault voltage of a bus, or of connected buses of
        a line, transformer, switch or phase shifter.

    Args:
        hnd    (c_int): object handle
        vdOut1 (c_double*9): voltage result, real part or magnitude,
                             at equipment terminals
        vdOut2 (c_double*9): voltage result, imaginary part or angle in degree
                             at equipment terminals
        style (c_int)     : voltage result style
                                1: output 012 sequence voltage in rectangular form
                                2: output 012 sequence voltage in polar form
                                2: output ABC phase voltage in rectangular form
                                4: output ABC phase voltage in polar form

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetSCVoltage.argtypes = [c_int, c_double*9, c_double*9, c_int]
    return ASPENOlxAPIDLL.OlxAPIGetSCVoltage( hnd, vdOut1, vdOut2, style )

#
def GetSCCurrent( hnd, vdOut1, vdOut2, style ):
    """ Retrieve post fault current for a generator, load, shunt, switched shunt,
        generating unit, load unit, shunt unit, transmission line, transformer,
        switch or phase shifter.
        You can get the total fault current by calling this function with the
        pre-defined handle of short circuit solution, HND_SC.

    Args:
        hnd    (c_int): object handle
        vdOut1 (c_double*12): current result, real part or magnitude, into
                             equipment terminals
        vdOut2 (c_double*12): current result, imaginary part or angle in degree,
                             into equipment terminals
        style (c_int)      : current result style
                            1: output 012 sequence current in rectangular form
                            2: output 012 sequence current in polar form
                            3: output ABC phase current in rectangular form
                            4: output ABC phase current in polar form

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetSCCurrent.argtypes = [c_int, c_double*12, c_double*12, c_int]
    return ASPENOlxAPIDLL.OlxAPIGetSCCurrent( hnd, vdOut1, vdOut2, style )

#
def GetSteppedEvent( step, timeStamp, fltCurrent, userDef, eventDesc, faultDest ):
    """ Retrieve detailed result of a step in stepped-event simulation

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

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetSteppedEvent.argstype = [c_int,c_void_p,c_void_p,c_void_p,c_char_p,c_char_p]
    return ASPENOlxAPIDLL.OlxAPIGetSteppedEvent( step, timeStamp, fltCurrent, userDef, eventDesc, faultDest )

#
def GetZoneName(no):
    """ Retrieve name of the given area number.

    Args:
        no (c_int): Zone number

    return:
        name (c_char_p): Name string

    Remarks:
        if the funtion fails to execute the return string will consist of
        error message that begins with the key words: "GetZoneName failure:..."
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGetZoneName.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIGetZoneName.restype = c_char_p
    return decode( ASPENOlxAPIDLL.OlxAPIGetZoneName(no) )

#
def Locate1LObj(hnd,opt):
    """ Locate an object on the 1-line diagram (bus, generator, load, shunt,
        switched shunt, transmission line, transformer, switch, phase shifter, relay group)

    Args:
        hnd (c_int): Object handle number.
        opt (c_int): If the object is not visible show nearest visible one
                     1 – Set; 0 - Reset

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    """
    __checkInit__(False)
    ASPENOlxAPIDLL.OlxAPILocate1LObj.argtypes = [c_int,c_int]
    return ASPENOlxAPIDLL.OlxAPILocate1LObj( hnd, opt)

#
def MakeOutageList(hnd, maxTiers, wantedTypes, branchList, listLen):
    """ Return list of neighboring branches that can be used as outage list
        in calling the DoFault function on a bus, branch or relay group.

    Args:
        hnd    (c_int): handle of a bus, branch or a relay group.
        maxTiers (c_int): Number of tiers (must be positive)
        wantedTypes (c_int): Branch type to comsider. Sum of one or more
             following values: 1- Line; 2- 2-winding transformer;
             4- Phase shifter; 8- 3-winding transformer; 16- Switch
        branchList (c_void_p): [in] array of c_int*listLen+1 or none
                               [out] zero-terminated list of branch handles.
        listLen (pointer(c_int)): [in] pointer to c_int with length of branchList array
                                  [out] Number of outage branches found.

    Remarks:
        Calling this function with None in place of branchList will let you determine the
        number of outage branches in the number of tiers specified.

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIMakeOutageList.argtypes = [c_int,c_int,c_int,c_void_p,c_void_p]
    return ASPENOlxAPIDLL.OlxAPIMakeOutageList(hnd, maxTiers, wantedTypes, branchList, listLen)

#
def OlxAPIEliminateZZBranch( hnd, nOption, pOutBuf ):
    """ Eliminate switch and zmall impeance line and merge two end buses if the switch is closed.

    Args:
        hnd (c_int)     Switch or line handle
        nOption (c_int) Bit 1: Repeat for all connected switches and lines
                        Bit 2: Output list of retained bus handles (-1 terminated)
                        Bit 3: Output list of eliminated bus handles (-1 terminated)
                        Bit 4: Also eliminate small impedance lines
        pOutBuf   [out] Buffer for outputs. Can be NULL.

    return:
        OLXAPI_OK     : Success
        OLXAPI_FAILURE: Failure
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIEliminateZZBranch.argstype = [c_int,c_int,c_void_p]
    ASPENOlxAPIDLL.OlxAPIEliminateZZBranch.restype = c_int
    return ASPENOlxAPIDLL.OlxAPIEliminateZZBranch(hnd, nOption, byref(pOutBuf))

#
def OlxAPIGfxOp( nCmd, vnInput, vnOutput ):
    """ OlxAPIGfxOp: Various 1-line graphic operations.

    Args:
        nCmd       [in] Command code. See remark
        vnInput    [in] Input data buffer
        vnOutput   [out] Output data buffer

    return:
        OLXAPI_OK     : Success
        OLXAPI_FAILURE: Failure
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIGfxOp.argstype = [c_int,c_void_p,c_void_p]
    ASPENOlxAPIDLL.OlxAPIGfxOp.restype = c_int
    return ASPENOlxAPIDLL.OlxAPIGfxOp(nCmd, vnInput, vnOutput)

#
class OlxAPIException(Exception):
    """Raise this exeption to report OlxAPI error messages. """
    def __init__(self, msg):
        self.msg = msg
    def __str__(self):
        return self.msg

#
def PickFault( index, tiers ):
    """ Select a specific short circuit simulation case.
        This function must be called before any post fault voltage and current results
        and relay time can be retrieved.

    Args:
        index (c_int): fault number or
                        SF_FIRST: first fault
                        SF_NEXT: next fault
                        SF_PREV: previous fault
                        SF_LAST: last available fault
        tiers (c_int): number of tiers around faulted bus to compute solution results

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIPickFault.argtypes = [c_int,c_int]
    return ASPENOlxAPIDLL.OlxAPIPickFault( index, tiers)

#
def PostData(hnd):
    """ Perform validation and update objct data in the network database
        Changes to the equipment data made through SetData function will
        not be committed to the program network database until after
        this function has been executed with success.

    Args:
        hnd (c_int): Object handle

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIPostData.argtypes = [c_int]
    return ASPENOlxAPIDLL.OlxAPIPostData(hnd)

#
def PrintObj1LPF(hnd):
    """ Return a text description of network database object
       (bus, generator, load, shunt, switched shunt, transmission line,
       transformer, switch, phase shifter, distance relay,
       overcurrent relay, fuse, recloser, relay group).

    Args:
        hnd (c_int/int): object handle

    return:
        string
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIPrintObj1LPF.argstype = [c_int]
    ASPENOlxAPIDLL.OlxAPIPrintObj1LPF.restype = c_char_p
    return decode(ASPENOlxAPIDLL.OlxAPIPrintObj1LPF(hnd))

#
def ReadChangeFile(filePath):
    """ Apply ASPEN .CHF and .ADX change file.

    Args:
        filePath (str) : Full path name of the ASPEN change file

    return:
        OLXAPI_FAILURE: There were errors and/or warnings. See remarks below.
        OLXAPI_OK     : No error and warning.

    Remarks:
        This API call always resets the script engine handle table,
        which will invalidate all handles that exist up to this point in the
        scrip program execution. For this reason user cannot re-use any of
        these handles in subsequence script calculation.

        All error and warnings messges generated during the read change file
        operation are saved to a text file with the default name
        PowerScriptTTYLog.txt in the Windows %TEMP% folder. User should always
        analyze the content of this file to determine the valididy of the
        network model after the read change file operation.
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIReadChangeFile.argtypes = [c_char_p]
    return ASPENOlxAPIDLL.OlxAPIReadChangeFile( encode3(filePath) )

#
def Run1LPFCommand(Params):
    """ Run OneLiner command.

    Args:
        Params (c_char_p): XML string, or full path name to XML file, containing a XML node
                        - Node name: OneLiner command
                        - Node attributes: required command parameters

    Command:
        ARCFLASHCALCULATOR - Faults | Arc-flash hazard calculator
        Attributes:
            REPORTPATHNAME    (*) Full pathname of report file.
            APPENDREPORT            [0] Append to existing report 0-No;1-Yes
            OUTFILETYPE     [2] Output file type 1- TXT; 2- CSV
            SELECTEDOBJ    Arcflash bus. Must  have one of following values
                       "PICKED "     the highlighted bus on the 1-line diagram
                       "'BNAME1                ,KV1;’BNAME2’,KV2;..."  Bus name and nominal kV.
            TIERS    [0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
            AREAS    [0-9999] Comma delimited list of area numbers and ranges to check relaygroups agains backup.
                           This attribute is ignored if SELECTEDOBJ is found.
            ZONES    [0-9999] Comma delimited list of zone numbers and ranges to check relaygroups agains backup.
                        This attribute is ignored if AREAS or SELECTEDOBJ are found.
            KVS    [0-999] Comma delimited list of KV levels and ranges to check relaygroups agains backup.
                        This attribute is ignored if SELECTEDOBJ is found.
            TAGS    Comma delimited list of bus tags. This attribute is ignored if SELECTEDOBJ is found.
            EQUIPMENTCAT    (*) Equipment category: 0-Switch gear; 1- Cable; 2- Open air; 3- MCC                s and panelboards 1kV or lower
            GROUNDED    (*) Is the equipment grounded 0-No; 1-Yes
            ENCLOSED    (*) Is the equipment inside enclosure 0-No; 1-Yes
            CONDUCTORGAP    (*) Conductor gap in mm
            WORKDIST    (*) Working distance in inches
            ARCDURATION    Arc duration calculation method. Must have one of following values:
                      "FIXED"     Use fixed duration
                      "FUSE "     Use fuse curve
                      "FASTEST"     Use fastest trip time of device in vicinity
                      "DEVICE"     Use trip time of specified device
                      "SEA"     Use stepped-event analysis
            ARCTIME    Arc duration in second. Must be present when ARCDURATION="FIXED"
            FUSECURVE    Fuse curve for arc duration calculation. Must be present when ARCDURATION=" FUSECURVE"
            BRKINTTIME    Breaker interrupting time in cycle. Must be present when ARCDURATION=" FASTEST" and "DEVICE"
            DEVICETIERS    [1] Number of tiers. Must be present when ARCDURATION=" FASTEST" and ="SEA"
            DEVICE    String  with location of the relaygroup and the relay name
                     "BNO1;                BNAME1                ;KV1;BNO2;                BNAME2                ;KV2;                CKT                ;BTYP; RELAY_ID; ".
                     Format description of these fields are is in OneLiner help section 10.2.
            ARCTIMELIMIT    [1] Perform no energy calculation when arc duration time is longer than 2 seconds

    Command:
        BUSFAULTSUMMARY : Faults | Bus fault summary
        Attributes:
            REPORTPATHNAME= (*) full valid path to report file
            BASELINECASE= pathname of base-line bus fault summary report in CSV format
            === Only when BASELINECASE is specified
            DIFFBASE= Basis for computing current deviation: [MAX3PH1LG] or MAXPHGND
            FLAGPCNT= [15] Current deviation percent threshold.
            === Only when BASELINECASE is not specified
            BUSLIST= Bus list, one on each row in format 'BusName',kV
            BUSNOLIST= Bus number list, coma delimited.
                      This attribute is ignored when BUSLIST is specified
            === Only when no BUSLIST and BUSNOLIST is  specified
            XGND= Fault reactance X
            RGND= Fault resistance R
            NOTAP= Exclude tap buses: [1]-TRUE; 0-FALSE
            PERUNITV= Report voltage in PU
            PERUNIT= Report current in PU
            AREAS= Area number range
            ZONES= Zone number range. This attribute is ignored when AREAS is specified
            BUSNOS= Additional bus number range
            KVS= Additional bus kV range
            TAGS= Additional tag filter
            TIERS= check lines in vicinity within this tier number
            AREAS= Check all lines in area range
            ZONES= Check all lines in zone range
            KVS=   Additional KV filter

    Command:
        CHECKRELAYOPERATIONPRC023 - Check | Relay loadability
        Attributes:
            REPORTPATHNAME= (*) full valid pathname of report file
            REPORTCOMMENT= Report comment string. 255 char or shorter
            SELECTEDOBJ=
                PICKED Check devices in selected relaygroup
                BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of branch to check(OneLiner Help section 10.2)
            TIERS= check relaygroups in vicinity within this tier number
            AREAS= Check all relaygroups in area range
            ZONES= Check all relaygroups in zone range
            KVS=   Additional KV filter
            TAGS=  Additional tag filter
            USETAGFLAG= [0]-AND;[1]-OR
            DEVICETYPE= [OCP DSP] Devide type to check. Space delimited
            APPENDREPORT=    Append report file: 0-False; [1]-True
            LINERATINGTYPE=    [3] Line rating to use: 0-first; 1-second; 2-Third;    3-Fourth
            XFMRRATINGTYPE=    [2] Transformer rating to use: 0-MVA1; 1-MVA2; 2-MVA3
            FWRLOADLONLY= [0] Consider load in forward direction only
            VOLTAGEPU= [0.85] Per unit voltage
            LINECURRMULT= [1.5] Line load current multiplier
            XFMRCURRMULT= [1.5] Transformer load current multiplier
            PFANGLE= [30] Power factor angle

    Command:
        CHECKRELAYOPERATIONPRC026 - run Check | Relay performance in stable power swing (PRC-026-1)
        Attributes:
            REPORTPATHNAME= (*) full valid pathname of report file
            REPORTCOMMENT= Report comment string. 255 char or shorter
            SELECTEDOBJ=
            PICKED Check devices in selected relaygroup
            BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of line to check(Help section 10.2)
            TIERS= check relaygroups in vicinity within this tier number
            AREAS= Check all relaygroups in area range
            ZONES= Check all relaygroups in zone range
            KVS=   Additional KV filter
            TAGS=  Additional tag filter
            DEVICETYPE= [OCP DSP] Devide type to check. Space delimited
            APPENDREPORT=    Append report file: 0-False; [1]-True
            SEPARATIONANGLE=    [120] System separation angle for stable power swing calculation
            DELAYLIMIT= [15] Report violation if relay trips faster than this limit (in cycles)
            CURRMULT= [1.0] Current multiplier to apply in relay trip checking

    Command:
        CHECKPRIBACKCOORD - Check | Primary/back relay coordination
        Attributes:
            REPORTPATHNAME  (*) Full pathname of report file.
            OUTFILETYPE     [2] Output file type 1- TXT; 2- CSV
            SELECTEDOBJ    Relay group to check against its backup. Must  have one of following values
            PICKED      the highlighted relaygroup on the 1-line diagram
            "BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;"  location string of the relaygroup. Format description is in OneLiner help section 10.2.
            TIERS    [0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
            AREAS    [0-9999] Comma delimited list of area numbers and ranges to check relaygroups agains backup.
            ZONES    [0-9999] Comma delimited list of zone numbers and ranges to check relaygroups agains backup. This attribute is ignored if AREAS is found.
            KVS    0-999] Comma delimited list of KV levels and ranges to check relaygroups agains backup. This attribute is ignored if SELECTEDOBJ is found.
            TAGS    Comma delimited list of tags to check relaygroups agains backup. This attribute is ignored if SELECTEDOBJ is found.
            COORDTYPE    Coordination type to check. Must  have one of following values
            "0"    OC backup/OC primary (Classical)
            "1"    OC backup/OC primary (Multi-point)
            "2"    DS backup/OC primary
            "3"    OC backup/DS primary
            "4"    DS backup/DS primary
            "5"    OC backup/Recloser primary
            "6"    All types/All types
            LINEPERCENT    Percent interval for sliding intermediate faults. This attribute is ignored if COORDTYPE is 0 or 5.
            RUNINTERMEOP    1-true; 0-false. Check  intermediate faults with end-opened. This attribute is ignored if COORDTYPE is 0 or 5.
            RUNCLOSEIN    1-true; 0-false. Check close-in fault. This attribute is ignored if COORDTYPE is 0 or 5.
            RUNCLOSEINEOP    1-true; 0-false. Check close-in fault with end-opened. This attribute is ignored if COORDTYPE is 0 or 5.
            RUNLINEEND    1-true; 0-false. Check line-end fault. This attribute is ignored if COORDTYPE is 0 or 5.
            RUNREMOTEBUS    1-true; 0-false. Check remote bus fault. This attribute is ignored if COORDTYPE is 0 or 5.
            RELAYTYPE    Relay types to check: 1-Ground; 2-Phase; 3-Both.
            FAULTTYPE    Fault  types to check: 1-3LG; 2-2LG; 4-1LF; 8-LL; or sum of values for desired selection
            OUTPUTALL    1- Include all cases in report; 0- Include only flagged cases in report
            MINCTI    Lower limit of acceptable CTI range
            MAXCTI    Upper limit of acceptable CTI range
            OUTRLYPARAMS    Include relay settings in report: 0-None; 1-OC;2-DS;3-Both
            OUTAGELINES    Run line outage contingency: 0-False; 1-True
            OUTAGEXFMRS    Run transformer outage contingency: 0-False; 1-True
            OUTAGEMULINES    Run mutual line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGEMULINESGND        Run mutual line outage and grounded contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGE2LINES    Run double line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGE1LINE1XFMR    Run double line and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0 or OUTAGEXFMRS =0
            OUTAGE2XFMRS    Run double and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEXFMRS =0
            OUTAGE3SOURCES    Outage only  3 strongest sources: 0-False; 1-True. Ignored if OUTAGEMULINES=0 and OUTAGEXFMRS =0

    Command:
        CHECKRELAYOPERATIONSEA - Check | Relay operation using stepped-events
        Attributes:
            REPORTPATHNAME    (*) Full pathname of folder for report files.
            OPTIONSFILEPATH Full pathname of the checking option file.
            REPORTCOMMENT        Additional comment string to include in all checking report files
            SELECTEDOBJ    Check line with selected relaygroup. Must  have one of following values
                "PICKED "     the highlighted relaygroup on the 1-line diagram
                "BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;"  location string of  the relaygroup. Format description is in OneLiner help section 10.2.
            TIERS    [0] Number of tiers around selected object. This attribute is ignored if SELECTEDOBJ is not found.
            AREAS    [0-9999] Comma delimited list of area numbers and ranges.
            ZONES    [0-9999] Comma delimited list of zone numbers and ranges. This attribute is ignored if AREAS is found.
            KVS    [0-999] Comma delimited list of KV levels and ranges. This attribute is ignored if SELECTEDOBJ is found.
            TAGS    Comma delimited list of tags. This attribute is ignored if SELECTEDOBJ is found.
            DEVICETYPE    Space delimited list of relay type types to take into consideration in stepped-events: OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF
            FAULTTYPE    Space delimited list of fault  types to take into consideration in stepped-events: 1LF, 3LG
            OUTAGELINES    Run line outage contingency: 0-False; 1-True
            OUTAGEXFMRS    Run transformer outage contingency: 0-False; 1-True
            OUTAGEMULINES    Run mutual line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGEMULINESGND        Run mutual line outage and grounded contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGE2LINES    Run double line outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0
            OUTAGE1LINE1XFMR    Run double line and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEMULINES=0 or OUTAGEXFMRS =0
            OUTAGE2XFMRS    Run double and transformer outage contingency: 0-False; 1-True. Ignored if OUTAGEXFMRS =0
            OUTAGE3SOURCES    Outage only  3 strongest sources: 0-False; 1-True. Ignored if OUTAGEMULINES=0 and OUTAGEXFMRS =0
            OUTAGEPILOT    Simulate pilot outage in N and N-1 cases

    Command:
        CHECKRELAYSETTINGS - Check | Relay settings
        Attributes:
            SELECTEDOBJ=
                PICKED Check line with selected relaygroup
                BNO1;'BNAME1';KV1;BNO2;'BNAME2';KV2;'CKT';BTYP;  location string of line to check(Help section 10.2)
            TIERS= check lines in vicinity within this tier number
            AREAS= Check all lines in area range
            ZONES= Check all lines in zone range
            KVS=   Additional KV filter
            TAGS=  Additional tag filter
            REPORTPATHNAME= (*) full valid path to report folder with write access
            OPTIONSFILEPATH Full pathname of the checking option file.
            REPORTCOMMENT= Report comment string. 255 char or shorter
            FAULTTYPE= 1LG, 3LG. Fault type to check. Space delimited
            DEVICETYPE= OCG, OCP, DSG, DSP, LOGIC, VOLTAGE, DIFF Devide type to check. Space delimited
            OUTAGELINES    Run Line outage contingency: 0-False; 1-True
            OUTAGEXFMRS    Run transformer outage contingency: 0-False; 1-True
            OUTAGE3SOURCES= 1 or 0 Outage only 3 strongest sources
            OUTAGEMULINES= 1 or 0 Outage mutually coupled lines
            OUTAGEMULINESGND= 1 or 0 Outage and ground ends of mutually coupled lines
            OUTAGE2LINES= 1 or 0 Double outage lines
            OUTAGE1LINE1XFMR= 1 or 0 Double outage line and transformer
            OUTAGE2XFMR= 1 or 0 Double outage transformers

    Command:
        DIFFANDMERGE - File | Compare and Merge...
        Attributes:
            FILEPATHA     Full pathname of file A. Can be ommited when and OLR file
                         had been opened.
            FILEPATHB (*) Full pathname of file B.
            FILEPATHBASE     Full pathname of common base.
            FILEPATHDIFF     Full pathname of the ADX file with DIFF result.
            FILEPATHMERGED  Full pathname of the OLR file with merge result.
            FILEPATHCFG     Full pathname of the configuration file (see OneLiner User Manual
                           Appendix F.2 for file formal details).

    Command:
        EXPORTNETWORK = File | Export network data
        Attributes:
            FORMAT     = Output format: [DXT]-ASPEN DXT; PSSE-PSS/E Raw and Seq
                                       XML- ASPEN XML; CSV- ASPEN CSV
            SCOPE      =  Export scope: [0]-Entire network; 1-Area number; 2- Zone number
            AREANO     =  Export area number
            ZONENO      =  Export zone number
            INCLUDETIES=  Include ties: [0]-False; 1-True
            ====DXT export only:
            DXTPATHNAME= (*) full valid pathname of ouput DXT file
            ====XML export only:

            OUTPUTPATHNAME= (*) full valid pathname of ouput XML file
            DATATYPE   = Types of data to include in the export. Integer number with
                        [Bit 1]-Network; Bit 2- Relay;
            ====PSSE export only:
            RAWPATHNAME= (*) full valid pathname of ouput RAW file
            SEQPATHNAME= (*) full valid pathname of ouput SEQ file
            PSSEVER      =  [33] PSS/E version
            X3MIDBUSNO =  [18000] First fictitious bus number for 3-w transformer mid point
            NEWBUSNO   =  [15000] First bus number for buses with no bus number

    Command:
        EXPORTRELAY - Relay | Export relay
        Attributes:
            FORMAT     =  Output format: [RAT]-ASPEN RAT;
            SCOPE      =  Export scope: [0]-Entire network; 1-Area number; 2- Zone number; 3-Invicinity of a bus
            AREANO     =  Export area number (*required when SCOPE=1)
            ZONENO      =  Export zone number (*required when SCOPE=2)
            SELECTEDOBJ=  Selected bus (*required when SCOPE=3). Must be a string with following content
                         PICKED - Selected bus on the 1-line diagram
                         'BusName' kV - Bus name in single quotes and kV separated by space
            TIERS      =  [0] Number of tiers (ignored when SCOPE<>3)
            DEVICETYPE =  Device type to export. Comma delimied list of the following:
                         OC: Overcurrent
                         DS: Distance
                         RC: Recloser
                         VR: Voltage relay
                         DIFF: Differential relay
                         SCHEME: Logic scheme
                         COORDPAIR: Coorination pair
                         [OC,DS,RC,VR,DIFF,COORDPAIR,SCHEME]
            LASTCHANGEDDATE =  [01-01-1986] Cutoff last changed date
            RATPATHNAME= (*) full valid pathname of ouput RAT file

    Command:
        INSERTTAPBUS - Network | Bus | Insert tap bus
        Attributes:
            BUSNAME1=    (*) Line bus 1 name
            BUSNAME2=    (*) Line bus 2 name
            KV=    (*) Line kV
            CKTID=    (*) Line circuit ID
            PERCENT=    (*) Percent distance to tap from bus 1 (must be between 0-100)
            TAPBUSNAME=    (*) Tap bus name

    Command:
        SAVEDATAFILE - File | Save and File | Save as
        Attributes:
            PATHNAME     = Name or full pathname of new OLR file for File | Save as command.
                          If only file name is given, file will be saved in the folder
                          where the current OLR file is located.
                          If no attribute is specified, the File | Save command will be executed.

    Command:
        SAVEFLTSPEC - Save desc and spec of faults in the result buffer to XML or CSV files
        Attributes:
            REPORTPATHNAME= (*)Output file pathname
            APPEND= [0]: overwrite existing file; 1: Append to existing file;
            FORMAT= output file format [0]: XML; 1: CSV
            CLEARPREV= Clear previous results flag to be recorded in the XML output
                        [0]: no; 1: yes;
            RANGE= Comma delimited list of fault index numbers and ranges(e.g. 1,3,5-10)
                    Default: save all faults in the results buffer

    Command:
        SETGENREFANGLE - Network | Set generator reference angle
        Attributes:
            REPORTPATHNAME    Full pathname of folder for report files.
            REFERENCEGEN    Bus name and kV of reference generator in format: 'BNAME', KV.
            EQUSOURCEOPTION    Option for calculating reference angle of equivalent sources. Must have one of the following values
            [ROTATE]     apply delta angle of existing reference gen
            SKIP       Leave unchanged. This option will be in effect automatically when old reference is not valid
            ASGEN      Use angle computed for regular generator

    Command:
        SIMULATEFAULT - Run OneLiner command FAULTS  | BATCH COMMAND & FAULT SPEC FILE | EXECUTE COMMAND
        Attributes:
            <FAULT> The <SIMULATEFAULT> XML node must contain one or more children nodes <FAULT>,
             one for each fault case to be simulated.
             The <FAULT> node can include XML attributes to specify the various fault simulation options:
                   PREFAULTV: Prefault voltage flag
                     0 - From linear network solution
                     1 - Flat
                     2 - From load flow
                   VPF: Flat prefault bus voltage
                   GENZTYPE: Generator reactance flag
                     0 - subtransient
                     1 - transient
                     2 - synchronous
                   IGNORE: a bit field array
                     bit 1 - ignore load
                     bit 2 - ignore positive sequence shunt
                     bit 3 - ignore line G and B
                     bit 4 - ignore transformer B
                   GENILIMIT: generator limit option
                     0 - Ignore
                     1 - Enforce limit 1
                     2 - Enforce limit 2
                   VCCS: Simulate VCCS flag 1-true; 0-false
                   MOV: MOV protected series capacitor iterative
                        solution flag 1-true; 0-false
                   MOVITERF: MOV protected series capacitor iterative
                               solution acceleration factor
                   MINZMUPAIR: Ignore mutual coupling threshold
                   MVASTYLE: Fault MVA calculatin method
                   GENW3: Simulate type-3 wind generator flag 1-true; 0-false
                   GENW4: Simulate CIR flag 1-true; 0-false
                   FLTOPTSTR: a space delimited string with all the fault otion fields in the order listed above
                          PREFAULTV VPF GENZTYPE IGNORE GENILIMIT VCCS MOV MOVITERF MINZMUPAIR MVASTYLE GENW3 GENW4
            <FLTSPEC>  Each of the <FAULT> nodes can include up to 40 fault specifications to be
                   simulated in the case. Fault specification string in the format described in
                   OneLiner’s user manual APPENDIX I:  FAULT SPECIFICATION FILE
                   must be included as the text of children nodes <FLTSPEC>.
            FLTSPCVS= Alternatively, a tab-delimited list of all the fault spec strings can be entered as
                  the attribute "FLTSPCVS" in the <FAULT> node.
            CLEARPREV= Clear previous result buffer flag [0]: no; 1: yes;

        return:
            OLXAPI_FAILURE: Failure
            NOfault        : Total number of faults in the solution buffer

        Remarks:
            When NOfault is not the expected value, call OlxAPIErrorString() for details of simulation error

    Command:
        SNAPSPC: Run OneLiner command Diagram | Snap to state plane coordinates
        Attributes:
             CENTERX - Screen center X in SPC
             CENTERY - Screen center Y in SPC
             SCALE   - Scaling factor
             OPTIONS - Bit array options
                       Bit 1: Auto scale (compute center X, Y from
                              SPC of all objects in the network)
                       Bit 2: Place all hidden buses using SPC
                       Bit 3: Reset all existing graphic before rebuilding
                              the 1-line using SPC
        return:
             OLXAPI_FAILURE: Failure
             NOmoves       : Total number of buses placed/moved
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPIRun1LPFCommand.argstype = [c_char_p]
    return ASPENOlxAPIDLL.OlxAPIRun1LPFCommand( encode3(Params) )

#
def SaveDataFile(filePath=''):
    """ Save OLR data file to disk
        Set filePath='' to save current OLR file

    Args:
        filePath (str) : Full path name of ASPEN OLR file.

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISaveDataFile.argtypes = [c_char_p]
    return ASPENOlxAPIDLL.OlxAPISaveDataFile( encode3(GetOlrFileName() if filePath=='' else filePath) )

#
def SetData(hnd, token, p_data):
    """ Assign a value to an object data field. (try with SetObjData(...)

    Args:
        hnd (c_int): Object handle
        token (c_int): Object field token
        p_data (c_void_p): New field value

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISetData.argstype = [c_int, c_int, c_void_p]
    return ASPENOlxAPIDLL.OlxAPISetData(hnd, token, p_data)

#
def SetObjUDF(hnd, fname, fval):
    """ Set the value of a user-defined field.

    Args:
        hnd (c_int)       [in] Object handle
        fname (c_char_p)  [in] Field name
        fval (c_char_p)   [in] Field value

    return:
        OLXAPI_OK     : Success
        OLXAPI_FAILURE: Object does not have UDF Field with the given name
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISetObjUDF.argtypes = [c_int,c_char_p,c_char_p]
    return ASPENOlxAPIDLL.OlxAPISetObjUDF(hnd,encode3(fname),encode3(fval))

#
def SetObjData(hnd, token, p_data):
    """ Assign a value to an object data field.

    Args:
        hnd (c_int/int)       : Object handle
        token (c_int/int)     : Object field token
        p_data (int/float,...): New field value

    return:
        None

    Raises:
        OlxAPIException

    Samples:
        BUS:
            busNameOld = OlxAPI.GetObjData(bsHnd,BUS_sName)
            if OLXAPI_OK != OlxAPI.SetObjData(bsHnd, BUS_sName, len(busNameOld+'_xx') ):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            if OLXAPI_OK != OlxAPI.PostData(bsHnd):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            busNameNew = OlxAPI.GetObjData(bsHnd,BUS_sName)
            print('busNameOld,busNameNew: %s,%s'%(busNameOld,busNameNew))

        GENUNIT:
            Xold = OlxAPI.GetObjData(hndUnit,GU_vdX)           # X of genunit
            if OLXAPI_OK != OlxAPI.SetObjData(hndUnit, GU_vdX, [0.21,0.22,0.23,0.24,0.25]):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            if OLXAPI_OK != OlxAPI.PostData(hndUnit):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            Xnew = OlxAPI.GetObjData(hndUnit,GU_vdX)
            print('Xold Xnew: ',Xold, Xnew)
    """
    if hnd==HND_SYS and token==SY_vParams:
        if type(p_data)!=str:
            raise OlxAPIException("TypeError: an string is required (got type %s)"%type(p_data).__name__)
        params = (c_void_p*2)()
        params[0] = cast(pointer(c_char_p(encode3(p_data))),c_void_p)
        params[1] = cast(pointer(c_char_p(encode3(""))),c_void_p)  # Terminator
        #
        if OLXAPI_OK!=SetData(hnd, token, params):
            raise OlxAPIException(ErrorString())
        return
    #
    vt = token//100
    if vt == VT_STRING:
        if type(p_data)!=str:
            raise OlxAPIException("a string is required (got type %s)"%type(p_data).__name__)
        c_data = c_char_p(encode3(p_data))
    elif vt == VT_DOUBLE:
        c_data = c_double(p_data)
    elif vt == VT_INTEGER:
        c_data = c_int(p_data)
    else:
        #
        if type(p_data)!=list:
            raise OlxAPIException('a list is required (got type %s)'%type(p_data).__name__)
        tc = EquipmentType(hnd)
        #
        if vt == VT_ARRAYINT:
            p_data = [OLXAPI_DFI if p1 is None else p1 for p1 in p_data]
            if tc==TC_SCHEME and token==LS_vnSignalType:
                intArray = c_int*MXOBJPARAMS
                signalType = intArray()
                for i in range(len(p_data)):
                    signalType[i] = c_int(p_data[i])
                signalType[i+1] = c_int(0)  # terminator
                #
                if OLXAPI_OK != SetData(hnd,token,byref(pointer(signalType))):
                    raise OlxAPIException(ErrorString())
                return
            c_data = (c_int*len(p_data))(*p_data)
            if OLXAPI_OK != SetData(hnd,token,byref(pointer(c_data))):
                raise OlxAPIException(ErrorString())
            return
        #
        elif vt == VT_ARRAYSTRING:
            p_data = [OLXAPI_DFS if p1 is None else p1 for p1 in p_data]
            if tc==TC_SCHEME and token==LS_vsSignalName:
                voidpArray = c_void_p*MXOBJPARAMS
                signalName = voidpArray()
                for i in range(len(p_data)):
                    signalName[i] = cast(pointer(c_char_p(encode3(p_data[i]))),c_void_p)
                signalName[i+1] = cast(pointer(c_char_p(encode3(""))),c_void_p)
                #
                if OLXAPI_OK != SetData(hnd,token,pointer(signalName)):
                    raise OlxAPIException(ErrorString())
                return
        else:
            #
            if tc == TC_GENUNIT and (token == GU_vdR or token == GU_vdX):
                count = 5
            elif tc == TC_LOADUNIT and (token == LU_vdMW or token == LU_vdMVAR):
                count = 3
            elif tc == TC_LINE and token == LN_vdRating:
                count = 4
            elif tc == TC_BREAKER and token in {BK_vdRecloseInt1,BK_vdRecloseInt2}:
                count = 3
            elif tc == TC_BREAKER and token in {BK_vnG1DevHnd}:
                count = len(p_data)
            elif tc == TC_SCHEME and token in {LS_vdSignalVar}:
                count = len(p_data)
            elif tc == TC_MU and token in {MU_vdFrom1,MU_vdTo1,MU_vdFrom2,MU_vdTo2,MU_vdR,MU_vdX}:
                count = 5
            else:
                raise OlxAPIException('OlxAPI.SetObjData failure: unsupported data token')
            if count!=len(p_data):
                raise OlxAPIException('len data must=%i, found len=%i'%(count,len(p_data)))
            #
            p_data = [OLXAPI_DFF if p1 is None else p1 for p1 in p_data]
            c_data = (c_double*count)(*p_data)
    #
    if OLXAPI_OK!=SetData(hnd, token, byref(c_data)):
        raise OlxAPIException(ErrorString())

#
def SetObjGraphicData(hnd, buf):
    """ Set object graphic data.

    Args:
        hnd (c_int): object handle
        buf (c_int*500): graphic data

    return:
        OLXAPI_OK: Success
        OLXAPI_FAILURE: Failure

    Remarks:
        Object graphic data record structure must be the same as in GetObjGraphicData()
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISetObjGraphicData.argstype = [c_int,POINTER(c_int)]
    ASPENOlxAPIDLL.OlxAPISetObjGraphicData.restype = c_int
    return ASPENOlxAPIDLL.OlxAPISetObjGraphicData(hnd,buf)

#
def SetObjMemo(hnd,memo):
    """ Assign memo string for a bus, generator, load, shunt, switched shunt,
        transmission line, transformer, switch, phase shifter, distance relay,
        overcurrent relay, fuse, recloser, relay group.

    Args:
        hnd (c_int): object handle
        memo (c_char_p): Memo string

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    Remarks:
        Line breaks must be included in the memo string as escape character
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISetObjMemo.argstype = [c_int,c_char_p]
    ASPENOlxAPIDLL.OlxAPISetObjMemo.restype = c_int
    return ASPENOlxAPIDLL.OlxAPISetObjMemo(hnd,encode3(memo))

#
def SetObjTags(hnd,tags):
    """ Assign tag string for a bus, generator, load, shunt, switched shunt,
        transmission line, transformer, switch, phase shifter, distance relay,
        overcurrent relay, fuse, recloser, relay group.

    Args:
        hnd (c_int): object handle
        tags (c_char_p): tag string

    return:
        OLXAPI_FAILURE: Failure
        OLXAPI_OK     : Success

    Remarks:
        tags string must be terminated with ; character
    """
    __checkInit__()
    ASPENOlxAPIDLL.OlxAPISetObjTags.argstype = [c_int,c_char_p]
    ASPENOlxAPIDLL.OlxAPISetObjTags.restype = c_int
    return ASPENOlxAPIDLL.OlxAPISetObjTags(hnd,encode3(tags))

#
def UnloadOlxAPI():
    """ Unload the OlxAPI.dll

    return:
        None

    Raises:
        OlxAPIException
    """
    if sys.executable.endswith('python.exe'):
        global ASPENOlxAPIDLL
        if ASPENOlxAPIDLL != None:
            if 0!= windll.kernel32.FreeLibrary(ASPENOlxAPIDLL._handle):
                ASPENOlxAPIDLL = None
            else:
                raise OlxAPIException("Failed to unload olxapi.dll")

#
def Version():
    """ OlxAPI engine version string in format version_major.version_minor
    """
    __checkInit__(False)
    buf = create_string_buffer(b'\000' * 1028)
    ASPENOlxAPIDLL.OlxAPIVersionInfo(buf)
    vData = decode(buf.value).split(" ")
    return vData[2]

#
def __ProcessGetDataBuf__(buf,tokenV,hnd):
    """Convert GetData binary data buffer into Python object of correct type
    """
    vt = tokenV//100
    if vt == VT_STRING:
        return decode(buf.value)
    elif vt in [VT_DOUBLE,VT_INTEGER]:
        return buf.value
    #
    tc = EquipmentType(hnd)
    if tc == TC_BREAKER and tokenV in {BK_vnG1DevHnd,BK_vnG2DevHnd,BK_vnG1OutageHnd,BK_vnG2OutageHnd}:
        count = MXSBKF
        return __getValue_i__(buf,count,True)
    if tc == TC_RLYGROUP and tokenV in {RG_vnPrimaryGroup,RG_vnBackupGroup}:
        count = MXSBKF
        return __getValue_i__(buf,count,True)
    #
    if tc == TC_SVD and tokenV == SV_vnNoStep:
        count = 8
        return __getValue_i__(buf,count)
    #
    if tc in {TC_RLYDSP,TC_RLYDSG} and tokenV in {DP_vParamLabels,DG_vParamLabels,DG_vParams,DP_vParams}:
        res = decode(cast(buf,c_char_p).value).split("\t")# String with tab delimited fields
        for i in range(len(res)):
            if res[i]=='':
                return res[:i]
        return res
    #
    if tc==TC_RLYDSP and tokenV ==DP_vdParams:
        res = cast(buf,POINTER(c_double*MXDSPARAMS)).contents
        vs = GetObjData(hnd,DP_vParamLabels)
        return res[:len(vs)]
    #
    if tc ==TC_RLYDSG and tokenV ==DG_vdParams:
        res = cast(buf,POINTER(c_double*MXDSPARAMS)).contents
        vs = GetObjData(hnd,DG_vParamLabels)
        return res[:len(vs)]
    #
    if tc==TC_SCHEME:
        if tokenV ==LS_vsSignalName:
            res = (decode(cast(buf, c_char_p).value)).split("\t")
            for i in range(len(res)):
                if res[i]=='':
                    break
            return res[:i]
        if tokenV ==LS_vnSignalType:
            count = 8
            return __getValue_i__(buf,count,True)
        if tokenV ==LS_vdSignalVar:
            count = 8
            res = cast(buf,POINTER(c_double*count)).contents
            for i in range(len(res)):
                if res[i]==0.0:
                    break
            return res[:i]
    elif tc == TC_DCLINE2 and tokenV==DC_vnBridges:
        count = 2
        return __getValue_i__(buf,count)
    elif tc==TC_SYS and tokenV==SY_vParams:
        res = (decode(cast(buf, c_char_p).value)).split("\t")
        for i in range(len(res)):
            if res[i]=='' and i%2==0:
                break
        return res[:i]
    #
    elif tc == TC_GENUNIT and tokenV in {GU_vdR,GU_vdX}:
        count = 5
    elif tc == TC_LOADUNIT and tokenV in {LU_vdMW,LU_vdMVAR}:
        count = 3
    elif tc == TC_SVD and tokenV in {SV_vdBinc,SV_vdB0inc}:
        count = 8
    elif tc == TC_LINE and tokenV == LN_vdRating:
        count = 4
    elif tc == TC_RLYGROUP and tokenV == RG_vdRecloseInt:
        count = 4
    elif tc == TC_RLYOCG and tokenV == OG_vdDirSetting:
        count = 8
    elif tc == TC_RLYOCG and tokenV == OG_vdDirSettingV15:
        count = 9
    elif tc == TC_RLYOCG and tokenV in {OG_vdDTPickup,OG_vdDTDelay}:
        count = 5
    elif tc == TC_RLYOCP and tokenV == OP_vdDirSetting:
        count = 8
    elif tc == TC_RLYOCP and tokenV == OP_vdDirSettingV15:
        count = 9
    elif tc == TC_RLYOCP and tokenV in {OP_vdDTPickup,OP_vdDTDelay}:
        count = 5
    elif tc == TC_RLYDSG and tokenV == DG_vdParams:
        count = MXDSPARAMS
    elif tc == TC_RLYDSG and tokenV in {DG_vdDelay,DG_vdReach,DG_vdReach1}:
        count = MXZONE
    elif tc == TC_RLYDSP and tokenV == DP_vParams:
        count = MXDSPARAMS
    elif tc == TC_RLYDSP and tokenV in {DP_vdDelay,DP_vdReach,DP_vdReach1}:
        count = MXZONE
    elif tc == TC_CCGEN and tokenV in {CC_vdV,CC_vdI,CC_vdAng}:
        count = MAXCCV
    elif tc == TC_BREAKER and tokenV in {BK_vdRecloseInt1,BK_vdRecloseInt2}:
        count = 3
    elif tc == TC_MU and tokenV in [MU_vdX,MU_vdR,MU_vdFrom1,MU_vdFrom2,MU_vdTo1,MU_vdTo2]:
        count = 5
    elif tc == TC_DCLINE2:# and tokenV in {OlxAPIConst.DC_vdAngleMax,OlxAPIConst.DC_vdAngleMin}:
        count = 2
    else:
        count = MXDSPARAMS
    val = cast(buf,POINTER(c_double*count)).contents
    return [v for v in val]

#internal
def __checkInit__(checkOLR=True):
    if ASPENOlxAPIDLL==None:
        raise OlxAPIException('OlxAPI - olxapi.dll is not yet initialized')
    if checkOLR and ASPENOLRFILE=='':
        raise OlxAPIException('ASPEN OLR file is not yet opened or closed')

#internal
def __getValue_i__(buf,count,stop=False):
    array = []
    val = cast(buf,POINTER(c_int*count)).contents
    for ii in range(count):
        if stop and val[ii] ==0:
            break
        array.append(val[ii])
    return array

#internal
def __getTokenParam__(tokens0,params0):
    if params0==None:
        ti,pi=[],[]
        for t1,p1 in tokens0.items():
            ti.append(t1)
            pi.append(p1)
        return __getTokenParam__(ti,pi)
    #
    tokens,params =[],[]
    for i in range(len(tokens0)):
        if tokens0[i] in {OBJ_sGUID,OBJ_sTags,OBJ_sMemo}:
            tokens.append(tokens0[i])
            params.append(params0[i])
    for i in range(len(tokens0)):
        if tokens0[i] not in {OBJ_sGUID,OBJ_sTags,OBJ_sMemo}:
            tokens.append(tokens0[i])
            params.append(params0[i])
    #
    intArray = c_int*MXOBJPARAMS
    voidpArray = c_void_p*MXOBJPARAMS
    tokens1 = intArray()
    params1 = voidpArray()
    doubleArray = c_double*20
    for i in range(len(tokens)):
        tokens1[i] = tokens[i]
        vt = tokens[i]//100
        if vt==VT_STRING:
            vsi = c_char_p( encode3(params[i]) )
        elif vt==VT_DOUBLE:
            vsi = c_double(params[i])
        elif vt==VT_INTEGER:
            vsi = c_int(params[i])
        elif vt==VT_ARRAYDOUBLE:
            vsi = doubleArray()
            for j in range(len(params[i])):
                vsi[j] = params[i][j]
        elif vt==VT_ARRAYINT:
            vsi = intArray()
            for j in range(len(params[i])):
                vsi[j] = params[i][j]
        elif vt==VT_ARRAYSTRING:
            vsi = voidpArray()
            for j in range(len(params[i])):
                vsi[j] = cast(pointer(c_char_p(encode3(params[i][j]))),c_void_p)
            #vsi[j+1] = cast(pointer(c_char_p(encode3(''))),c_void_p)
        params1[i] = cast(pointer(vsi),c_void_p)
    #
    tokens1[i+1] = 0 #zero terminated list
    return tokens1,params1
