"""Demo usage of ASPEN OlrxAPI in Python.
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2018, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

import sys
import os.path
import math
import Tkinter as tk
import tkMessageBox
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

c_GetDataBuf = create_string_buffer(b'\000' * 10*1024*1024)    # 10 KB buffer for string data
c_GetDataBuf_double = c_double(0)
c_GetDataBuf_int = c_int(0)
def olrx_get_data(args):
    """Get network object data field value
    """
    c_token = c_int(args["token"])
    c_hnd = c_int(args["hnd"])
    try:
        data = args["data"]
    except:
        data = None
    dataBuf = make_GetDataBuf( c_token, data )
    ret = OlrxAPIGetData(c_hnd, c_token, byref(dataBuf))
    if OLRXAPI_OK == ret:
        args["data"] = process_GetDataBuf(dataBuf,c_token,c_hnd)
    return ret

def make_GetDataBuf(token,data):
    """Prepare correct data buffer for OlrxAPIGetData()
    """
    global c_GetDataBuf
    global c_GetDataBuf_double
    global c_GetDataBuf_int
    vt = token.value//100
    if vt == VT_DOUBLE:
        if data != None:
            try:
                c_GetDataBuf_double = c_double(data)
            except:
                pass
        return c_GetDataBuf_double
    elif vt == VT_INTEGER:
        if data != None:
            try:
                c_GetDataBuf_int = c_int(data)
            except:
                pass
        return c_GetDataBuf_int
    else:
        try:
            c_GetDataBuf.value = data
        except:
            pass
        return c_GetDataBuf

def process_GetDataBuf(buf,token,hnd):
    """Convert GetData binary data buffer into Python object of correct type
    """
    vt = token.value//100
    if vt == VT_STRING:
        return buf.value #
    elif vt == VT_DOUBLE:
        return buf.value
    elif vt == VT_INTEGER:
        return buf.value
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
            for v in val:
                array.append(v)
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
    vd3Mag = (c_double*3)(0.0)
    vd3Ang = (c_double*3)(0.0)
    while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( hnd, c_int(TC_BRANCH), byref(hndBr) ) ) :
        if ( OLRXAPI_FAILURE == OlrxAPIGetData( hndBr, c_int(BR_nBus2Hnd), byref(hndBus2) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if ( OLRXAPI_FAILURE == OlrxAPIGetPSCVoltage( hndBr, vd3Mag, vd3Ang, c_int(2) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if ( OLRXAPI_FAILURE == OlrxAPIGetSCVoltage( hndBr, vd9Mag, vd9Ang, c_int(4) ) ) :
            raise OlrxAPIException(OlrxAPIErrorString()) 
        # Voltage on bus 1
        print OlrxAPIFullBusName( hnd ), \
              "VP (pu)=", vd3Mag[0], "@", vd3Ang[0],    \
              "Va=", vd9Mag[0], "@", vd9Ang[0],    \
              "Vb=", vd9Mag[1], "@", vd9Ang[1],    \
              "Vc=", vd9Mag[2], "@", vd9Ang[2]
        # Voltage on bus 2
        print OlrxAPIFullBusName( hndBus2 ), \
              "VP (pu)=", vd3Mag[1], "@", vd3Ang[1],    \
              "Va=", vd9Mag[3], "@", vd9Ang[3],    \
              "Vb=", vd9Mag[4], "@", vd9Ang[4],    \
              "Vc=", vd9Mag[5], "@", vd9Ang[5]

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
        szDevice = create_string_buffer(b'\000' * 128)
        flag = c_int(0)
        if ( OLRXAPI_FAILURE == OlrxAPIGetRelayTime( hndRelay, c_double(1.0), flag, byref(dTime), szDevice ) ):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print (" time (s) = " + str(dTime.value) + "  device= " + szDevice.value)

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
    szEventDesc = create_string_buffer(b'\000' * 512 * 4)     # 4*512 bytes buffer for event description
    szFaultDest = create_string_buffer(b'\000' * 512 * 50)    # 50*512 bytes buffer for fault description
    nSteps = OlrxAPIGetSteppedEvent( c_int(0), byref(dTime), byref(dCurrent), 
                                               byref(nUserEvent), szEventDesc, szFaultDest )
    print "Stepped-event simulation completed successfully with ", nSteps-1, " events"
    for ii in range(1, nSteps):
        OlrxAPIGetSteppedEvent( c_int(ii), byref(dTime), byref(dCurrent), 
                                          byref(nUserEvent), szEventDesc, szFaultDest )
        print "Time = ", dTime.value, " Current= ", dCurrent.value
        print cast(szFaultDest, c_char_p).value
        print cast(szEventDesc, c_char_p).value

def branchSearch( bsBusName1, bsKV1, bsBusName2, bsKV2, sCKID ):
    hnd1    = OlrxAPIFindBus( bsBusName1, bsKV1 )
    if hnd1 == OLRXAPI_FAILURE:
            print "Bus ", bsBusName1, bsKV1, " not found"
    hnd2    = OlrxAPIFindBus( bsBusName2, bsKV2 )
    if hnd2 == OLRXAPI_FAILURE:
            print "Bus ", bsBusName2, bsKV2, " not found"

    hndBr = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( hnd1, c_int(TC_BRANCH), byref(hndBr) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = hndBr.value
        argsGetData["token"] = BR_nBus2Hnd
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            hndFarBus = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if hndFarBus == hnd2:
            argsGetData = {}
            argsGetData["hnd"] = hndBr.value
            argsGetData["token"] = BR_nType
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                brType = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            if brType == TC_LINE:
                typeCode = LN_sID
            elif brType == TC_XFMR:
                typeCode = XF_sID
            elif brType == TC_XFMR3:
                typeCode = X3_sID
            elif brType == TC_PS:
                typeCode = PS_sID
            argsGetData = {}
            argsGetData["hnd"] = hndBr.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                itemHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            argsGetData = {}
            argsGetData["hnd"] = itemHnd
            argsGetData["token"] = typeCode
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                sID = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            if sID == sCKID:
                return hndBr.value

def binarySearch(Array, nKey, nMin, nMax):
    while (nMax >= nMin):
        nMid = (nMax + nMin) / 2
        if Array[nMid] == nKey:
            return nMid
        else:
            if Array[nMid] < nKey:
                nMin = nMid + 1
            else:
                nMax = nMid - 1
    return -1

def compuOneLiner(nLineHnd, ProcessedHnd, hndOffset):
    argsGetData = {}
    argsGetData["hnd"] = nLineHnd
    argsGetData["token"] = LN_nBus1Hnd
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        Bus1Hnd = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_nBus2Hnd
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        Bus2Hnd = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_dR
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dR = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_dX
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dX = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_dR0
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dR0 = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_dX0
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dX0 = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_dLength
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dLength = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["token"] = LN_sName
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        sName = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    argsGetData["hnd"] = Bus1Hnd
    argsGetData["token"] = BUS_dKVnominal
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        dKV = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())

    BusHndList = (c_int*100)(0)

    BusListCount = 2
    if Bus1Hnd > Bus2Hnd:
        BusHndList[0] = Bus2Hnd
        BusHndList[1] = Bus1Hnd
    else:
        BusHndList[0] = Bus1Hnd
        BusHndList[1] = Bus2Hnd

    aLine1 = OlrxAPIFullBusName(Bus1Hnd) + " - " + OlrxAPIFullBusName(Bus2Hnd) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
    ProcessedHnd[nLineHnd-hndOffset] = 1

    # find tap segments on Bus1 side
    BusHnd  = Bus1Hnd
    while (True):
        LineHnd = FindTapSegmentAtBus(BusHnd, ProcessedHnd, hndOffset, sName)
        if LineHnd == 0:
            break
        ProcessedHnd[LineHnd-hndOffset] = 1
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_dR
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dX
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dR0
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dX0
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dLength
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlrxAPIFullBusName(BusHnd) + " - " + OlrxAPIFullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
        print("Segment: " + aLine)
        ProcessedHnd[LineHnd-hndOffset] = 1
        BusHndList[BusListCount] = BusHnd
        BusListCount = BusListCount+1
        BusHnd  = BusFarHnd
        BusHndList[BusListCount] = BusFarHnd
        BusListCount = BusListCount+1

    # find tap segments on Bus1 side
    BusHnd  = Bus2Hnd
    while (True):
        LineHnd = FindTapSegmentAtBus(BusHnd, ProcessedHnd, hndOffset, sName)
        if LineHnd == 0:
            break
        ProcessedHnd[LineHnd-hndOffset] = 1
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_dR
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dX
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dR0
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dX0
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_dLength
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlrxAPIFullBusName(BusHnd) + " - " + OlrxAPIFullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
        print("Segment: " + aLine)
        ProcessedHnd[LineHnd-hndOffset] = 1
        BusHndList[BusListCount] = BusHnd
        BusListCount = BusListCount+1
        BusHnd  = BusFarHnd
        BusHndList[BusListCount] = BusFarHnd
        BusListCount = BusListCount+1


    if BusListCount > 2:
        print("Segment: " + aLine1)
        Changed = 1    
        while (Changed > 0):
            Changed = 0
            for ii in range(0,BusListCount-2):
                if BusHndList[ii] > BusHndList[ii+1]:
                    nTemp = BusHndList[ii]
                    BusHndList[ii] = BusHndList[ii+1]
                    BusHndList[ii+1] = nTemp
                    Changed = 1

        for ii in range(0,BusListCount-2):
            if BusHndList[ii] == BusHndList[ii+1]:
                BusHndList[ii] = 0
                BusHndList[ii+1] = 0

        jj = 0
        nflg = 0
        for ii in range(0,BusListCount-1):
            if BusHndList[ii] > 0:
                BusHndList[jj] = BusHndList[ii]
                jj = jj + 1
            if jj == 2:
                nflg = 1 
                break

        if nflg == 1:
            aLine1 = OlrxAPIFullBusName(busHndList[0]) + " - " + OlrxAPIFullBusName(busHndList[1]) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
    print("Line: " + aLine1)

def FindTapSegmentAtBus( BusHnd, ProcessedHnd, hndOffset, sName ):
    FindTapSegmentAtBus = 0
    argsGetData = {}
    argsGetData["hnd"] = BusHnd
    argsGetData["token"] = BUS_nTapBus
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        TapCode = argsGetData["data"]
    else:
        raise OlrxAPIException(OlrxAPIErrorString())
    if TapCode == 0:
        return 0
    BranchHnd = c_int(0)
    while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nType
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            TypeCode = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if TypeCode != TC_LINE:
            continue
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nHandle
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            LineHnd = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if ProcessedHnd[LineHnd - hndOffset] == 1:
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sName
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            sNameThis = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if sNameThis == sName:
            return LineHnd
        if sNameThis[:3] == "[T]":
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sID
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            sIDThis = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if sIDThis == "T":
            continue
        FindTapSegmentAtBus = LineHnd
    return FindTapSegmentAtBus

def printImpedance(dR, dX, dKV):
    dMag = math.sqrt(dR*dR + dX*dX)*dKV*dKV/100.0
    if dR != 0.0:
        dAng = math.atan(dX/dR)*180/3.14159
    else:
        if dX > 0:
            dAng = 90
        else:
            dAng = -90
    aLine = "{0:.5f}".format(dR) + "+j" + "{0:.5f}".format(dX) + "pu(" + "{0:.2f}".format(dMag) + "@" + "{0:.2f}".format(dAng) + "Ohm"
    return aLine

def GetRemoteTerminals( BranchHnd, TermsHnd ):
    """
    Purpose: Find all remote ends of a line. All taps are ignored. Close switches are included
    Usage: 
      BranchHnd [in] branch handle of the local terminal
      TermsHnd  [in] array to hold list of branch handle at remote ends
    Return: Number of remote ends
    """
    ListSize = 0
    MXSIZE =100
    TempListSize = 1
    TempBrList = (c_int*(MXSIZE))(0)
    TempBrList[TempListSize-1] = BranchHnd    
    while (TempListSize > 0):
        NearEndHnd = TempBrList[TempListSize-1]
        TempListSize = TempListSize - 1
        ListSize = FindOppositeBranch( NearEndHnd, TermsHnd, TempBrList, TempListSize, MXSIZE )
    return ListSize

def FindOppositeBranch( NearEndBrHnd, OppositeBrList, TempBrList, TempListSize, MXSIZE ):
    TempListSize2 = 0
    ListSize = 0
    TempList2 = (c_int*(MXSIZE))(0)
    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nInService
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        nStatus = argsGetData["data"]
    if nStatus != 1:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nBus2Hnd
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        nBus2Hnd = argsGetData["data"]
    else:
        return 0
    for ii in range(0, TempListSize2):
        if TempListSize2[ii] == nBus2Hnd:
            return 0
    if (TempListSize2-1 == MXSIZE):
        print("Ran out of buffer spase. Edit script code to incread MXSIZE" )
    TempListSize2 = TempListSize2 + 1
    TempList2[TempListSize2-1] = nBus2Hnd

    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nHandle
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        nThisLineHnd = argsGetData["data"]
    else:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = nBus2Hnd
    argsGetData["token"] = BUS_nTapBus
    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
        nTapBus= argsGetData["data"]
    else:
        return 0
    nBranchHnd = c_int(0);
    if (nTapBus != 1):
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                break;
        ListSize = ListSize + 1
        OppositeBrList[ListSize-1] = nBranchHnd.value
    else:
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                continue;
            argsGetData["token"] = BR_nType
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                nType = argsGetData["data"]
            else:
                return 0
            if ((nType != TC_LINE) and (nType != TC_SWITCH)):
                continue
            if (nType == TC_SWITCH):
                argsGetData = {}
                argsGetData["hnd"] = nLindHnd
                argsGetData["token"] = SW_nInService
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
                argsGetData["token"] = SW_nStatus
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
            if (nType == TC_LINE):
                argsGetData = {}
                argsGetData["hnd"] = nLineHnd
                argsGetData["token"] = LN_nInService
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
            if (TempListSize-1 == MXSIZE):
                print("Ran out of buffer spase. Edit script code to incread MXSIZE" )
                return 0
            TempListSize = TempListSize + 1
            TempBrList[TempListSize] = nBranchHnd
    return ListSize


