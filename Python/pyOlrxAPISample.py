"""Demo usage of ASPEN OlrxAPI in Python.
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2017, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

import sys
import os.path
from pyOlrxAPI import *
from ctypes import *


def olrx_get_equipment(args):
    """Get OLR network object handle
    """
    c_tc = c_int(args["tc"])
    c_hnd = c_int(args["hnd"])
    ret = OlrxAPIGetEquipment(c_tc, pointer(c_hnd) )
    if OLRXAPI_OK == ret:
       args["hnd"] = c_hnd.value
    return ret

def olrx_set_data(args):
    """Set network object data field
    """
    token = args["token"]
    hnd = args["hnd"]
    vt = token//100
    if vt == VT_STRING:
        c_data = c_char_p(args["data"])
    elif vt == VT_DOUBLE:
        c_data =c_double(args["data"])
    elif vt == VT_INTEGER:
        c_data = c_int(args["data"])
    else:
        tc = OlrxAPIEquipmentType(hnd)
        if tc == TC_GENUNIT and (token == GU_vdR or token == GU_vdX):
            count = 5
        elif tc == TC_LOADUNIT and (token == LU_vdMW or token == LU_vdMVAR):
            count = 3
        elif tc == TC_LINE and token == LN_vdRating:
            count = 4
        array = args["data"]
        c_data = (c_double*len(array))(*array)
    return OlrxAPISetData(c_int(hnd), c_int(token), byref(c_data))

def olrx_post_data(args):
    """Post network object data
    """
    return OlrxAPIPostData(c_int(args["hnd"]))

c_GetDataBuf = create_string_buffer(b'\000' * 10*1024*1024)    # 10 KB buffer for data
def olrx_get_data(args):
    """Get network object data field value
    """
    c_token = c_int(args["token"])
    c_hnd = c_int(args["hnd"])
    global c_GetDataBuf
    ret = OlrxAPIGetData(c_hnd, c_token, byref(c_GetDataBuf))
    if OLRXAPI_OK == ret:
        args["data"] = process_GetDataBuf(c_GetDataBuf,c_token,c_hnd)
    return ret

def process_GetDataBuf(buf,token,hnd):
    """Convert GetData binary data buffer into Python object of correct type
    """
    vt = token.value//100
    if vt == VT_STRING:
        return buf.value #
    elif vt == VT_DOUBLE:
        return cast(buf, POINTER(c_double)).contents.value #
    elif vt == VT_INTEGER:
        return cast(buf, POINTER(c_int)).contents.value #
    else:
        array = []
        tc = OlrxAPIEquipmentType(hnd)
        if tc == TC_BREAKER and (token.value == BK_vnG1DevHnd or \
            token.value == BK_vnG2DevHnd or \
            token.value == BK_vnG1OutageHnd or \
            token.value == BK_vnG2OutageHnd):
            val = cast(buf, POINTER(c_int*MXSBKF)).contents  # int array of size MXSBKF
            for ii in range(0,MXSBKF-1):
                array.append(val[ii])
                if array[ii] == 0:
                    break
        elif (tc == TC_SVD and (token.value == SV_vnNoStep)):
            val = cast(buf, POINTER(c_int*8)).contents  # int array of size 8
            for ii in range(0,7):
                array.append(val[ii])
        elif (tc == TC_RLYDSP and (token.value == DP_vParams or token.value == DP_vParamLabels)) or \
             (tc == TC_RLYDSG and (token.value == DG_vParams or token.value == DG_vParamLabels)):
            # String with tab delimited fields
            return (cast(buf, c_char_p).value).split("\t")
        else:
            if tc == TC_GENUNIT and (token.value == GU_vdR or token.value == GU_vdX):
                count = 5
            elif tc == TC_LOADUNIT and (token.value == LU_vdMW or token.value == LU_vdMVAR):
                count = 3
            elif tc == TC_SVD and (token.value == SV_vdBinc or token.value == SV_vdB0inc):
                count = 3
            elif tc == TC_LINE and token.value == LN_vdRating:
                count = 4
            elif tc == TC_RLYGROUP and token.value == RG_vdRecloseInt:
                count = 3
            elif tc == TC_RLYOCG and token.value == OG_vdDirSetting:
                count = 2
            elif tc == TC_RLYOCP and token.value == OP_vdDirSetting:
                count = 2
            elif tc == TC_RLYDSG and token.value == DG_vdParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSG and (token.value == DG_vdDelay or token.value == DG_vdReach or token.value == DG_vdReach1):
                count = MXZONE
            elif tc == TC_RLYDSP and token.value == DP_vParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSP and (token.value == DP_vdDelay or token.value == DP_vdReach or token.value == DP_vdReach1):
                count = MXZONE
            elif tc == TC_CCGEN and (token.value == CC_vdV or token.value == CC_vdI or token.value == CC_vdAng):
                count = MAXCCV
            elif tc == TC_BREAKER and (token.value == BK_vdRecloseInt1 or token.value == BK_vdRecloseInt2):
                count = 3
            val = cast(buf, POINTER(c_double*count)).contents  # double array of size count
            for ii in range(0,count-1):
                array.append(val[ii])
        return array

def run_busFault(busHnd):
    """Report fault on a bus
    """
    hnd = c_int(busHnd)
    fltConn = (c_int*4)(0,0,1,0)   # 3LG, 2LG, 1LG, LL
    fltOpt = (c_double*15)(0)
    fltOpt[0] = 1       # Bus or Close-in
    fltOpt[1] = 0       # Bus or Close-in w/ outage
    fltOpt[2] = 0       # Bus or Close-in with end opened
    fltOpt[3] = 0       # Bus or Close#-n with end opened w/ outage
    fltOpt[4] = 0       # Remote bus
    fltOpt[5] = 0       # Remote bus w/ outage
    fltOpt[6] = 0       # Line end
    fltOpt[7] = 0       # Line end w/ outage
    fltOpt[8] = 0       # Intermediate %
    fltOpt[9] = 0       # Intermediate % w/ outage
    fltOpt[10] = 0      # Intermediate % with end opened
    fltOpt[11] = 0      # Intermediate % with end opened w/ outage
    fltOpt[12] = 0      # Auto seq. Intermediate % from [*] = 0
    fltOpt[13] = 0      # Auto seq. Intermediate % to [*] = 0
    fltOpt[14] = 0      # Outage line grounding admittance in mho [***] = 0.
    outageLst = (c_int*100)(0)
    outageLst[0] = 0
    outageOpt = (c_int*4)(0)
    outageOpt[0] = 0
    fltR = c_double(0.0)
    fltX = c_double(0.0)
    clearPrev = c_int(1)
    if OLRXAPI_FAILURE == OlrxAPIDoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    if OLRXAPI_FAILURE == OlrxAPIPickFault(c_int(SF_FIRST),c_int(9)):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print "======================"
    print OlrxAPIFaultDescription(0)
    hndBr = c_int(0);
    hndBus2 = c_int(0)
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    vd9Mag = (c_double*9)(0.0)
    vd9Ang = (c_double*9)(0.0)
    while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( hnd, c_int(TC_BRANCH), byref(hndBr) ) ) :
        if ( OLRXAPI_FAILURE == OlrxAPIGetData( hndBr, c_int(BR_nBus2Hnd), byref(hndBus2) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if ( OLRXAPI_FAILURE == OlrxAPIGetSCVoltage( hndBr, vd9Mag, vd9Ang, c_int(4) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        # Voltage on bus 1
        print OlrxAPIFullBusName( hnd ), \
              "Va=", vd12Mag[0], "@", vd12Ang[0],    \
              "Vb=", vd12Mag[1], "@", vd12Ang[1],    \
              "Vc=", vd12Mag[2], "@", vd12Ang[2]
        # Voltage on bus 2
        print OlrxAPIFullBusName( hndBus2 ), \
              "Va=", vd12Mag[3], "@", vd12Ang[3],    \
              "Vb=", vd12Mag[4], "@", vd12Ang[4],    \
              "Vc=", vd12Mag[5], "@", vd12Ang[5]

        if ( OLRXAPI_FAILURE == OlrxAPIGetSCCurrent( hndBr, vd12Mag, vd12Ang, 4 ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        # Current from 1
        print "Ia=", vd12Mag[0], "@", vd12Ang[0],    \
              "Ib=", vd12Mag[1], "@", vd12Ang[1],    \
              "Ic=", vd12Mag[2], "@", vd12Ang[2]
        # Relay time
        hndRlyGroup = c_int(0)
        if ( OLRXAPI_OK == OlrxAPIGetData( hndBr, c_int(BR_nRlyGrp1Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)
        if ( OLRXAPI_OK == OlrxAPIGetData( hndBr, c_int(BR_nRlyGrp2Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)

def print_relayTime(hndRlyGroup):
    """Print operating time of all relays in a relay group
    """
    hndRelay = c_int(0)
    while OLRXAPI_OK == OlrxAPIGetRelay( hndRlyGroup, byref(hndRelay) ) :
        print OlrxAPIFullRelayName( hndRelay )
        dTime = c_double(0)
        if ( OLRXAPI_FAILURE == OlrxAPIGetRelayTime( hndRelay, c_double(1.0), byref(dTime) ) ):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print " time (s) = ", dTime.value

def run_steppedEvent(busHnd):
    """Run stepped-event simulation on a bus
    """
    hnd = c_int(busHnd)
    runOpt = (c_int*5)(1,1,1,1,1)   # OCG, OCP, DSG, DSP, SCHEME
    fltOpt = (c_double*64)(0)
    fltOpt[0] = 1       #    Fault connection code 
                        #    1=3LG
                        #    2=2LG BC, 3=2LG CA, 4=2LG AB
                        #    5=1LG A, 5=1LG B, 6=1LG C
                        #    7=LL BC, 7=LL CA, 8=LL AB

    fltOpt[1] = 0       # Intermediate percent between 0.01-99.99. 0 for a close-in fault.
                        #This parameter is ignored if nDevHnd is a bus handle.
    fltOpt[2] = 0       # Fault resistance, ohm 
    fltOpt[3] = 0       # Fault reactance, ohm 
    fltOpt[4] = 0       # Zero or Fault connection code for additional user event 
    fltOpt[4+1] = 0     # Time  of additional user event, seconds.
    fltOpt[4+2] = 0     # Fault resistance in additinoal user event, ohm 
    fltOpt[4+3] = 0     # Fault reactancein additinoal user event, ohm 
                        #....
    noTiers = c_int(5)
    if OLRXAPI_FAILURE == OlrxAPIDoSteppedEvent(hnd, fltOpt, runOpt, noTiers):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    # Call GetSteppedEvent with 0 to get total number of events simulated
    dTime = c_double(0)
    dCurrent = c_double(0)
    nUserEvent = c_int(0)
    szEventDesc = create_string_buffer(b'\000' * 512 * 4)    # 4*512 bytes buffer for event description
    szFaultDest = create_string_buffer(b'\000' * 512 * 2)    # 2*512 bytes buffer for fault description
    nSteps = OlrxAPIGetSteppedEvent( c_int(0), byref(dTime), byref(dCurrent), 
                                               byref(nUserEvent), szEventDesc, szFaultDest )
    print "Stepped-event simulation completed successfully with ", nSteps-1, " events"
    for ii in range(1, nSteps):
        OlrxAPIGetSteppedEvent( c_int(ii), byref(dTime), byref(dCurrent), 
                                          byref(nUserEvent), szEventDesc, szFaultDest )
        print "Time = ", dTime.value, " Current= ", dCurrent.value
        print cast(szFaultDest, c_char_p).value
        print cast(szEventDesc, c_char_p).value

"""
import Tkinter
import tkFileDialog
def locateASPENOlrxAPI():
    Tkinter.Tk().withdraw() # Close the root window
    opts = {}
    opts['filetypes'] = [('ASPEN OlxAPI.DLL',('olrxapi.dll'))]
    opts['title'] = 'Locate ASPEN OlrxAPI DLL'
    in_path = tkFileDialog.askopenfilename(**opts)
    return os.path.dirname(in_path) + "/"
"""

def testDoRelayCoordination():
    """Test API function OlrxAPIDoRelayCoordination
    """
    checkParams = '<CHECKPRIBACKCOORD' \
                  ' REPORTPATHNAME="c:\\000tmp\\checkcoord.csv"' \
                  ' OUTFILETYPE="1"' \
                  ' SELECTEDOBJ="6; \'NEVADA\'; 132.; 8; \'REUSENS\'; 132.; \'1\'; 1;"' \
                  ' COORDTYPE="6"' \
                  ' OUTPUTALL="1"' \
                  ' MINCTI="0.05"' \
                  ' MAXCTI="99"' \
                  ' LINEPERCENT="15"' \
                  ' RELAYTYPE="3"' \
                  ' FAULTTYPE="5"' \
                  ' />' 
    print checkParams
    if OLRXAPI_FAILURE == OlrxAPIRun1LPFCommand(checkParams):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print "Success"

def testOlrxAPI():
    try:
        if len(sys.argv) == 1:
            print "Usage: " + sys.argv[0] + " YourNetwork.olr"
            return 0
        olrFilePath = sys.argv[1]
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print "OLR file does not exit: ", olrFilePath
            return 1

        if OLRXAPI_FAILURE == OlrxAPILoadDataFile(olrFilePath,1):
            raise OlrxAPIException(OlrxAPIErrorString()) 

        testDoRelayCoordination()

#        return 0

        argsGetEquipment = {}

        # Test GetData with special handles
        argsGetData = {}
        argsGetData["hnd"] = HND_SYS
        argsGetData["token"] = SY_nNObus
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print "Number of buses: ", argsGetData["data"]
 #       """        
        # Test SetData and GetData
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0
        ii = 100
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            busHnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = busHnd
            argsGetData["token"] = BUS_dKVnominal
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busKV = argsGetData["data"]
            argsGetData["token"] = BUS_sName
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busNameOld = argsGetData["data"]
            argsSetData = {}
            argsSetData["hnd"] = busHnd
            argsSetData["token"] = BUS_sName
            argsSetData["data"] = busNameOld+str(ii+1)
            if OLRXAPI_OK != olrx_set_data(argsSetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            argsGetData["token"] = BUS_nNumber
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busNumberOld = argsGetData["data"]
            argsSetData["token"] = BUS_nNumber
            argsSetData["data"] = ii+1
            if OLRXAPI_OK != olrx_set_data(argsSetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            if OLRXAPI_OK != OlrxAPIPostData(busHnd):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            argsGetData["token"] = BUS_sName
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busNameNew = argsGetData["data"]
            argsGetData["token"] = BUS_nNumber
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busNumberNew = argsGetData["data"]
            print "Old:", busNumberOld, busNameOld, busKV, "kV -> New: ", busNumberNew, busNameNew, busKV, "kV"
            ii = ii + 1
#        return 0

        argsGetEquipment["tc"] = TC_GENUNIT
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            argsGetData = {}
            hnd = argsGetEquipment["hnd"]
            argsGetData["hnd"] = hnd
            argsGetData["token"] = GU_vdX
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print "X=", argsGetData["data"]
            argsSetData = {}
            argsSetData["hnd"] = hnd
            argsSetData["token"] = GU_vdX
            argsSetData["data"] = [0.21,0.22,0.23,0.24,0.25]
            if OLRXAPI_OK != olrx_set_data(argsSetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            if OLRXAPI_OK != OlrxAPIPostData(hnd):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print "Xnew=", argsGetData["data"]
#            return 0


        # Test GetData
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            busHnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = busHnd
            argsGetData["token"] = BUS_sName
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busName = argsGetData["data"]
            argsGetData["token"] = BUS_dKVnominal
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            busKV = argsGetData["data"]
            print busName, busKV
#        return 0

        argsGetEquipment["tc"] = TC_RLYDSP
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = DP_vParamLabels
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            val = argsGetData["data"]
            print val
 #       return 0

        argsGetEquipment["tc"] = TC_GENUNIT
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = GU_vdX
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            val = argsGetData["data"]
            print val[0], val[1]
#        return 0

        argsGetEquipment["tc"] = TC_BREAKER
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = BK_vnG1DevHnd
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            val = argsGetData["data"]
            print val[0] #, val[1]
#        return 0
#        """
        # Test Fault simulation
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
            busHnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = busHnd
            argsGetData["token"] = BUS_nNumber
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            bno = argsGetData["data"]
            argsGetData["token"] = BUS_sName
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            bname = argsGetData["data"]
            print "hnd=", busHnd, bname, bno
            if bno > 0 and bno < 999999:
                print "\n>>>>>>Test bus fault"
                run_busFault(busHnd)
                print "\n>>>>>>Test bus fault SEA"
                run_steppedEvent(busHnd)
        return 0

    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def main(argv=None):

    #
    # IMPORTANT: Successfull initialization is required before any 
    #            other OlrxAPI call can be executed.
    #
    dllPath = os.path.dirname(os.path.abspath(sys.argv[0])) + "\\..\\"
    if not os.path.isfile(dllPath + "olrxapi.dll") :
        dllPath = os.path.dirname(os.path.abspath(sys.argv[0])) + "\\..\\obj\\debug_1L\\"
        #dllPath = locateASPENOlrxAPI()
        if not os.path.isfile(dllPath + "olrxapi.dll") :
            print "olrxapi.dll not found"
            return 1
    try:
        InitOlrxAPI(dllPath)
        print "\nASPEN OlrxAPI path: ", dllPath, "\nVersion:", OlrxAPIVersion(), " Build: ", OlrxAPIBuildNumber(), "\n"
    except OlrxAPIException as err:
        print( "OlrxAPI Init Error: {0}".format(err) )
        return 1        # error:

    return testOlrxAPI()

if __name__ == '__main__':
    status = main()
    sys.exit(status)
    
