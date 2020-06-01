"""
Demo usage of ASPEN OlxAPI in Python.
common utils
"""
from __future__ import print_function

__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.2.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

from ctypes import *
from OlxAPIConst import *
import OlxAPI
import math
import os,time
from datetime import datetime
import subprocess
# Tkinter python 2|3
try:
    import tkinter as tk
    import tkinter.filedialog as tkf
    import tkinter.messagebox as tkm
    from tkinter import ttk
except:
    import Tkinter as tk
    import tkFileDialog as tkf
    import tkMessageBox as tkm
    import ttk

#
def open_olrFile(olrFile,readonly):
    """
    Open OLR file in
        olrFile = (str) OLR file
    Args:
        readonly (int): open in read-only mode. 1-true; 0-false
    """
    # load dll
    OlxAPI.InitOlxAPI(OLXAPI_DLL_PATH)
    #
    olrFile1 = olrFile
    if not (olrFile.upper()).endswith(".OLR"):
        olrFile1 += ".OLR"
    #
    if OLXAPI_OK == OlxAPI.LoadDataFile(olrFile1,readonly):
         print("\tFile opened successfully: " + str(olrFile1))
    else:
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

def get_equipment(args):
    """Get OLR network object handle
    """
    c_tc = c_int(args["tc"])
    c_hnd = c_int(args["hnd"])
    ret = OlxAPI.GetEquipment(c_tc, pointer(c_hnd) )
    if OLXAPI_OK == ret:
       args["hnd"] = c_hnd.value
    return ret

def set_data(args):
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
        tc = OlxAPI.EquipmentType(hnd)
        if tc == TC_GENUNIT and (token == GU_vdR or token == GU_vdX):
            count = 5
        elif tc == TC_LOADUNIT and (token == LU_vdMW or token == LU_vdMVAR):
            count = 3
        elif tc == TC_LINE and token == LN_vdRating:
            count = 4
        array = args["data"]
        c_data = (c_double*len(array))(*array)
    return OlxAPI.SetData(c_int(hnd), c_int(token), byref(c_data))

def post_data(args):
    """Post network object data
    """
    return OlxAPI.PostData(c_int(args["hnd"]))

c_GetDataBuf = create_string_buffer(b'\000' * 10*1024*1024)    # 10 KB buffer for string data
c_GetDataBuf_double = c_double(0)
c_GetDataBuf_int = c_int(0)
def get_data(args):
    """Get network object data field value
    """
    c_token = c_int(args["token"])
    c_hnd = c_int(args["hnd"])
    try:
        data = args["data"]
    except:
        data = None
    dataBuf = make_GetDataBuf( c_token, data )
    ret = OlxAPI.GetData(c_hnd, c_token, byref(dataBuf))
    if OLXAPI_OK == ret:
        args["data"] = process_GetDataBuf(dataBuf,c_token,c_hnd)
    return ret

def make_GetDataBuf(token,data):
    """Prepare correct data buffer for OlxAPI.GetData()
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
        tc = OlxAPI.EquipmentType(hnd)
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
    if OLXAPI_FAILURE == OlxAPI.DoFault(hnd, fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    if OLXAPI_FAILURE == OlxAPI.PickFault(c_int(SF_FIRST),c_int(9)):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    print ("======================")
    print (OlxAPI.FaultDescription(0))
    hndBr = c_int(0);
    hndBus2 = c_int(0)
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    vd9Mag = (c_double*9)(0.0)
    vd9Ang = (c_double*9)(0.0)
    vd3Mag = (c_double*3)(0.0)
    vd3Ang = (c_double*3)(0.0)
    while ( OLXAPI_OK == OlxAPI.GetBusEquipment( hnd, c_int(TC_BRANCH), byref(hndBr) ) ) :
        if ( OLXAPI_FAILURE == OlxAPI.GetData( hndBr, c_int(BR_nBus2Hnd), byref(hndBus2) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ( OLXAPI_FAILURE == OlxAPI.GetPSCVoltage( hndBr, vd3Mag, vd3Ang, c_int(2) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ( OLXAPI_FAILURE == OlxAPI.GetSCVoltage( hndBr, vd9Mag, vd9Ang, c_int(4) ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        # Voltage on bus 1
        print (OlxAPI.FullBusName( hnd ), \
              "VP (pu)=", vd3Mag[0], "@", vd3Ang[0],    \
              "Va=", vd9Mag[0], "@", vd9Ang[0],    \
              "Vb=", vd9Mag[1], "@", vd9Ang[1],    \
              "Vc=", vd9Mag[2], "@", vd9Ang[2])
        # Voltage on bus 2
        print (OlxAPI.FullBusName( hndBus2 ), \
              "VP (pu)=", vd3Mag[1], "@", vd3Ang[1],    \
              "Va=", vd9Mag[3], "@", vd9Ang[3],    \
              "Vb=", vd9Mag[4], "@", vd9Ang[4],    \
              "Vc=", vd9Mag[5], "@", vd9Ang[5])

        if ( OLXAPI_FAILURE == OlxAPI.GetSCCurrent( hndBr, vd12Mag, vd12Ang, 4 ) ) :
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        # Current from 1
        print ("Ia=", vd12Mag[0], "@", vd12Ang[0],    \
              "Ib=", vd12Mag[1], "@", vd12Ang[1],    \
              "Ic=", vd12Mag[2], "@", vd12Ang[2])
        # Relay time
        hndRlyGroup = c_int(0)
        if ( OLXAPI_OK == OlxAPI.GetData( hndBr, c_int(BR_nRlyGrp1Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)
        if ( OLXAPI_OK == OlxAPI.GetData( hndBr, c_int(BR_nRlyGrp2Hnd), byref(hndRlyGroup) ) ) :
            if hndRlyGroup != 0 :
                print_relayTime(hndRlyGroup)

def print_relayTime(hndRlyGroup):
    """Print operating time of all relays in a relay group
    """
    hndRelay = c_int(0)
    while OLXAPI_OK == OlxAPI.GetRelay( hndRlyGroup, byref(hndRelay) ) :
        print (OlxAPI.FullRelayName( hndRelay ))
        dTime = c_double(0)
        szDevice = create_string_buffer(b'\000' * 128)
        flag = c_int(0)
        if ( OLXAPI_FAILURE == OlxAPI.GetRelayTime( hndRelay, c_double(1.0), flag, byref(dTime), szDevice ) ):
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        print (" time (s) = " + str(dTime.value) + "  device= " + str(szDevice.value))

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
    if OLXAPI_FAILURE == OlxAPI.DoSteppedEvent(hnd, fltOpt, runOpt, noTiers):
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    # Call GetSteppedEvent with 0 to get total number of events simulated
    dTime = c_double(0)
    dCurrent = c_double(0)
    nUserEvent = c_int(0)
    szEventDesc = create_string_buffer(b'\000' * 512 * 4)     # 4*512 bytes buffer for event description
    szFaultDest = create_string_buffer(b'\000' * 512 * 50)    # 50*512 bytes buffer for fault description
    nSteps = OlxAPI.GetSteppedEvent( c_int(0), byref(dTime), byref(dCurrent),
                                               byref(nUserEvent), szEventDesc, szFaultDest )
    print ("Stepped-event simulation completed successfully with ", nSteps-1, " events")
    for ii in range(1, nSteps):
        OlxAPI.GetSteppedEvent( c_int(ii), byref(dTime), byref(dCurrent),
                                          byref(nUserEvent), szEventDesc, szFaultDest )
        print ("Time = ", dTime.value, " Current= ", dCurrent.value)
        print (cast(szFaultDest, c_char_p).value)
        print (cast(szEventDesc, c_char_p).value)

def branchSearch( bsBusName1, bsKV1, bsBusName2, bsKV2, sCKID ):
    hnd1    = OlxAPI.FindBus( bsBusName1, bsKV1 )
    if hnd1 == OLXAPI_FAILURE:
            print ("Bus ", bsBusName1, bsKV1, " not found")
    hnd2    = OlxAPI.FindBus( bsBusName2, bsKV2 )
    if hnd2 == OLXAPI_FAILURE:
            print ("Bus ", bsBusName2, bsKV2, " not found")

    hndBr = c_int(0)
    while ( OLXAPI_OK == OlxAPI.GetBusEquipment( hnd1, c_int(TC_BRANCH), byref(hndBr) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = hndBr.value
        argsGetData["token"] = BR_nBus2Hnd
        if (OLXAPI_OK == get_data(argsGetData)):
            hndFarBus = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if hndFarBus == hnd2:
            argsGetData = {}
            argsGetData["hnd"] = hndBr.value
            argsGetData["token"] = BR_nType
            if (OLXAPI_OK == get_data(argsGetData)):
                brType = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
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
            if (OLXAPI_OK == get_data(argsGetData)):
                itemHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
            argsGetData = {}
            argsGetData["hnd"] = itemHnd
            argsGetData["token"] = typeCode
            if (OLXAPI_OK == get_data(argsGetData)):
                sID = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
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
    if (OLXAPI_OK == get_data(argsGetData)):
        Bus1Hnd = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_nBus2Hnd
    if (OLXAPI_OK == get_data(argsGetData)):
        Bus2Hnd = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dR
    if (OLXAPI_OK == get_data(argsGetData)):
        dR = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dX
    if (OLXAPI_OK == get_data(argsGetData)):
        dX = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dR0
    if (OLXAPI_OK == get_data(argsGetData)):
        dR0 = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dX0
    if (OLXAPI_OK == get_data(argsGetData)):
        dX0 = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_dLength
    if (OLXAPI_OK == get_data(argsGetData)):
        dLength = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["token"] = LN_sName
    if (OLXAPI_OK == get_data(argsGetData)):
        sName = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    argsGetData["hnd"] = Bus1Hnd
    argsGetData["token"] = BUS_dKVnominal
    if (OLXAPI_OK == get_data(argsGetData)):
        dKV = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())

    BusHndList = (c_int*100)(0)

    BusListCount = 2
    if Bus1Hnd > Bus2Hnd:
        BusHndList[0] = Bus2Hnd
        BusHndList[1] = Bus1Hnd
    else:
        BusHndList[0] = Bus1Hnd
        BusHndList[1] = Bus2Hnd

    aLine1 = OlxAPI.FullBusName(Bus1Hnd) + " - " + OlxAPI.FullBusName(Bus2Hnd) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
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
        if (OLXAPI_OK == get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX
        if (OLXAPI_OK == get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dR0
        if (OLXAPI_OK == get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX0
        if (OLXAPI_OK == get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dLength
        if (OLXAPI_OK == get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLXAPI_OK == get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLXAPI_OK == get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
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
        if (OLXAPI_OK == get_data(argsGetData)):
            dRn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX
        if (OLXAPI_OK == get_data(argsGetData)):
            dXn = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dR0
        if (OLXAPI_OK == get_data(argsGetData)):
            dR0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dX0
        if (OLXAPI_OK == get_data(argsGetData)):
            dX0n = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_dLength
        if (OLXAPI_OK == get_data(argsGetData)):
            dL = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        argsGetData["token"] = LN_nBus2Hnd
        if (OLXAPI_OK == get_data(argsGetData)):
            BusFarHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if BusFarHnd == BusHnd:
            argsGetData = {}
            argsGetData["hnd"] = LineHnd
            argsGetData["token"] = LN_nBus1Hnd
            if (OLXAPI_OK == get_data(argsGetData)):
                BusFarHnd = argsGetData["data"]
            else:
                raise OlxAPI.Exception(OlxAPI.ErrorString())
        dLength = dLength + dL
        dR  = dR  + dRn
        dX  = dX  + dXn
        dR0 = dR0 + dR0n
        dX0 = dX0 + dX0n
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) + ": Z=" + printImpedance(dRn,dXn,dKV) + " Zo=" + printImpedance(dR0n,dX0n,dKV) + " L=" + str(dL)
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
            aLine1 = OlxAPI.FullBusName(busHndList[0]) + " - " + OlxAPI.FullBusName(busHndList[1]) + ": Z=" + printImpedance(dR,dX,dKV) + " Zo=" + printImpedance(dR0,dX0,dKV) + " L=" + str(dLength)
    print("Line: " + aLine1)

def FindTapSegmentAtBus( BusHnd, ProcessedHnd, hndOffset, sName ):
    FindTapSegmentAtBus = 0
    argsGetData = {}
    argsGetData["hnd"] = BusHnd
    argsGetData["token"] = BUS_nTapBus
    if (OLXAPI_OK == get_data(argsGetData)):
        TapCode = argsGetData["data"]
    else:
        raise OlxAPI.Exception(OlxAPI.ErrorString())
    if TapCode == 0:
        return 0
    BranchHnd = c_int(0)
    while ( OLXAPI_OK == OlrxAPIGetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nType
        if (OLXAPI_OK == get_data(argsGetData)):
            TypeCode = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if TypeCode != TC_LINE:
            continue
        argsGetData = {}
        argsGetData["hnd"] = BranchHnd.value
        argsGetData["token"] = BR_nHandle
        if (OLXAPI_OK == get_data(argsGetData)):
            LineHnd = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if ProcessedHnd[LineHnd - hndOffset] == 1:
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sName
        if (OLXAPI_OK == get_data(argsGetData)):
            sNameThis = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
        if sNameThis == sName:
            return LineHnd
        if sNameThis[:3] == "[T]":
            continue
        argsGetData = {}
        argsGetData["hnd"] = LineHnd
        argsGetData["token"] = LN_sID
        if (OLXAPI_OK == get_data(argsGetData)):
            sIDThis = argsGetData["data"]
        else:
            raise OlxAPI.Exception(OlxAPI.ErrorString())
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
    aLine = "{0:.5f}".format(dR) + "+j" + "{0:.5f}".format(dX) + "pu(" + "{0:.2f}".format(dMag) + "@" + "{0:.2f}".format(dAng) + "Ohm)"
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
    if (OLXAPI_OK == get_data(argsGetData)):
        nStatus = argsGetData["data"]
    if nStatus != 1:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = NearEndBrHnd
    argsGetData["token"] = BR_nBus2Hnd
    if (OLXAPI_OK == get_data(argsGetData)):
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
    if (OLXAPI_OK == get_data(argsGetData)):
        nThisLineHnd = argsGetData["data"]
    else:
        return 0
    argsGetData = {}
    argsGetData["hnd"] = nBus2Hnd
    argsGetData["token"] = BUS_nTapBus
    if (OLXAPI_OK == get_data(argsGetData)):
        nTapBus= argsGetData["data"]
    else:
        return 0
    nBranchHnd = c_int(0);
    if (nTapBus != 1):
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLXAPI_OK == get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                break;
        ListSize = ListSize + 1
        OppositeBrList[ListSize-1] = nBranchHnd.value
    else:
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( nBus2Hnd, c_int(TC_BRANCH), byref(nBranchHnd) ) ) :
            argsGetData = {}
            argsGetData["hnd"] = nBranchHnd.value
            argsGetData["token"] = BR_nHandle
            if (OLXAPI_OK == get_data(argsGetData)):
                nLineHnd = argsGetData["data"]
            else:
                return 0
            if (nThisLineHnd == nLineHnd):
                continue;
            argsGetData["token"] = BR_nType
            if (OLXAPI_OK == get_data(argsGetData)):
                nType = argsGetData["data"]
            else:
                return 0
            if ((nType != TC_LINE) and (nType != TC_SWITCH)):
                continue
            if (nType == TC_SWITCH):
                argsGetData = {}
                argsGetData["hnd"] = nLindHnd
                argsGetData["token"] = SW_nInService
                if (OLXAPI_OK == get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
                argsGetData["token"] = SW_nStatus
                if (OLXAPI_OK == get_data(argsGetData)):
                    nStatus = argsGetData["data"]
                else:
                    return 0
                if (nStatus != 1):
                    continue
            if (nType == TC_LINE):
                argsGetData = {}
                argsGetData["hnd"] = nLineHnd
                argsGetData["token"] = LN_nInService
                if (OLXAPI_OK == get_data(argsGetData)):
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

#
dictCode_BR_sID   = {TC_LINE:LN_sID  , TC_SWITCH:SW_sID  , TC_PS: PS_sID  , TC_XFMR:XR_sID  , TC_XFMR3:X3_sID  , TC_SCAP:SC_sID   }
dictCode_BR_sName = {TC_LINE:LN_sName, TC_SWITCH:SW_sName, TC_PS: PS_sName, TC_XFMR:XR_sName, TC_XFMR3:X3_sName, TC_SCAP:SC_sName }
dictCode_BR_RAT   = {TC_LINE:1       , TC_SWITCH:7       , TC_PS: 3       , TC_XFMR:2       , TC_XFMR3:10}
dictCode_BR_nBus  = {TC_LINE:[LN_nBus1Hnd,LN_nBus2Hnd], TC_SWITCH:[SW_nBus1Hnd,SW_nBus2Hnd] , TC_PS: [PS_nBus1Hnd,PS_nBus2Hnd], \
                     TC_SCAP:[SC_nBus1Hnd,SC_nBus2Hnd], TC_XFMR:[XR_nBus1Hnd,XR_nBus2Hnd]   , TC_XFMR3 :[X3_nBus1Hnd,X3_nBus2Hnd,X3_nBus3Hnd]}

#
def getDataByBranch(hndBr,scode):
    # hndBr: branch handle
    # scode = 'sID' | 'sName'
    # return
    #
    typ  = getEquipmentData([hndBr],BR_nType)[0]
    ehnd = getEquipmentData([hndBr],BR_nHandle)[0]
    #
    if scode == "sID":
        paraCode = dictCode_BR_sID[typ]
    elif scode == "sName":
        paraCode = dictCode_BR_sName[typ]
    else:
        raise OlrxAPIException("Error scode")
    #
    return getEquipmentData([ehnd],paraCode)[0]

#
def getBusByBranch(hndBr):
    """
    Find bus of branch
        Args :
            hndBr :  branch handle

        Returns:
            bus[nBus1Hnd,nBus2Hnd, (nBus3Hnd)]

        Raises:
            OlrxAPIException
    """
    bres = []
    b1 = getEquipmentData([hndBr],BR_nBus1Hnd)[0]
    bres.append(b1)
    #
    b2 = getEquipmentData([hndBr],BR_nBus2Hnd)[0]
    bres.append(b2)
    try:
        b3 = getEquipmentData([hndBr],BR_nBus3Hnd)[0]
        bres.append(b3)
    except:
        pass
    #
    return bres

#
def getBusByEquipment(ehnd,TC_Type):
    """
    Find bus of equipment
        Args :
            ehnd :  equipment handle
            TC_Type: TC_LINE | TC_SCAP | TC_PS | TC_SWITCH | TC_XFMR | TC_XFMR3

        Returns:
            bus[nBus1Hnd,nBus2Hnd, (nBus3Hnd)]

        Raises:
            OlrxAPIException
    """
    busHnd = dictCode_BR_nBus[TC_Type]
    #
    busRes = []
    for b1 in busHnd:
        busRes.append(getEquipmentData([ehnd],b1)[0])
    #
    return busRes

#
def getEquipmentData(ehnd,paraCode):
    """
    Find data of element [] (line/bus/...)
        Args :
            ehnd []:  handle element
            nParaCode: code data (BUS_,LN_,...)

        Returns:
            data [len(bhnd)]

        Raises:
            OlrxAPIException
   """
    # get data
    res = []
    vt = paraCode//100
    val1 = setValType(vt,0)
    for ehnd1 in ehnd:
        if ( OLXAPI_FAILURE == OlxAPI.GetData( ehnd1, c_int(paraCode), byref(val1) ) ) :
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if vt==VT_STRING:
            res.append((val1.value).decode('UTF-8'))
        else:
            res.append(val1.value)
    #
    return res

#
def get1EquipmentData_try(ehnd,paraCode):
    """
    Find data of 1 element (line/bus/...) with try/except =-1 if not found
        Args :
            ehnd : handle element
            nParaCode: code data (BUS_,LN_,...)

        Returns:
            data (=-1 if not found)

        Raises:
            OlrxAPIException
    """
    try:
        val = getEquipmentData([ehnd],paraCode)[0]
        return val
    except:
        return -1

#
def get1EquipmentData_array(ehnd,paraCode):
    """
    Find array data of 1 element (line/relay group...)
        Args :
            ehnd :  handle element
            paraCode: code data (RG_,...)

        Returns:
            data array [] (array data attached to ehnd)

        Raises:
            OlrxAPIException
   """
    # get data
    res = []
    vt = paraCode//100
    val1 = setValType(vt,0)
    while ( OLXAPI_OK  == OlxAPI.GetData(ehnd,c_int(paraCode), byref(val1) ) ) :
        if vt==VT_STRING:
            res.append((val1.value).decode("UTF-8"))
        else:
            res.append(val1.value)
    #
    return res

#
def setValType(vt,value):
##    vt = paracode//100
    if vt == VT_STRING:
        if value==0:
            return create_string_buffer(b'\000' * 128)
        else:
            return c_char_p(value.encode("UTF-8"))
    elif vt == VT_DOUBLE:
       return c_double(value)
    elif vt == VT_INTEGER or vt==0:
       return c_int(value)
    #
    raise OlxAPI.OlrxAPIException("Error of paraCode")

#
def getBusEquipmentData(bhnd,paraCode):
    """
    Retrieves the handle of all equipment of a given type (paraCode)
    that is attached to bus [].

        Args :
            bhnd :  [bus handle]
            nParaCode : code data (BR_nHandle,GE_nBusHnd...)

        Returns:
           [][]  = [len(bhnd)] [len(all equipment)]

        Raises:
            OlrxAPIException
   """
    # get data
    res = []
    vt = paraCode//100
    val1 = setValType(vt,0)
    for bhnd1 in bhnd:
        r2 = []
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( bhnd1, c_int(paraCode), byref(val1) ) ) :
            if vt==VT_STRING:
                r2.append((val1.value).decode("UTF-8"))
            else:
                r2.append(val1.value)
        res.append(r2)
    return res

#
def getEquipmentHandle(TC_type):
    """
    Find all handle of element [] (line/bus/...)
        Args :
            TC_type : type element (TC_LINE, TC_BUS,...)

        Returns:
            all handle

        Raises:
            OlrxAPIException
   """
    res = []
    hndBr = c_int(0)
    while ( OLXAPI_OK == OlxAPI.GetEquipment(c_int(TC_type), byref(hndBr) )) :
        res.append(hndBr.value)
    return res

#
def branchIsInType(hndBr,typeConsi):
    """
    test if branch (in service) is in type defined by typeConsi:  [TC_LINE,TC_SWITCH,TC_SCAP, TC_PS, TC_XFMR,TC_XFMR3]
                       Close switches are tested
    """
    type1 = getEquipmentData([hndBr],BR_nType)[0]
    inSer = getEquipmentData([hndBr],BR_nInService)[0]
    #
    if inSer !=1:
        return False
    #
    if type1 not in typeConsi:
        return False

    if (TC_SWITCH in typeConsi) and (type1 ==TC_SWITCH):# if Switch
        swHnd = getEquipmentData([hndBr],BR_nHandle)[0]
        status = getEquipmentData([swHnd],SW_nStatus)[0]
        if status==0: # 0 OPEN, 1 close
            return False
    return True

#
def branchesNextToBranch(hndBr,typeConsi):
    """
    Purpose: Find list branches next to a branch
             All taps are ignored. Close switches are included
             Branches out of service are ignored.

        Args :
            hndBr :  branch handle
            typeConsi: type of branche considered
                       [TC_LINE,TC_SWITCH,TC_SCAP, TC_PS, TC_XFMR]

        returns :
            nTap: nTap of Bus2 (of hndBr)
            br_Self: branch inverse (by bus) of hndBr
            br_res []: list branch next
                                              | br_res
                                              |
                                  hndBr       |      br_res
        Illustration:       Bus1------------Bus2------------
                                  br_Self     |
                                              |
                                              | br_res
        Raises:
            OlrxAPIException
    """
    br_Self = -1
    br_res = []
    #
    b2 = getEquipmentData([hndBr],BR_nBus2Hnd)[0]
    # get nTap
    nTap = getEquipmentData([b2],BUS_nTapBus)[0]
    # equipment handle of hndBr
    equiHnd_0 = getEquipmentData([hndBr],BR_nHandle)[0]
    # all branch connect to b2
    allBr0 = getBusEquipmentData([b2],TC_BRANCH)[0]
    #
    allBr = []
    for br1 in allBr0:
        if branchIsInType(br1,typeConsi): #test branch is in type considered, Switch close
            allBr.append(br1)
    #
    equiHnd_all = getEquipmentData(allBr,BR_nHandle)
    #
    for i in range(len(allBr)):
        if equiHnd_all[i]== equiHnd_0: # branch same equipment
            br_Self = allBr[i]
        else:
            br_res.append(allBr[i])
    #
    return nTap,br_Self,br_res

#
def getOppositeBranch_1(hndBr,typeConsi): # Without XFMR3
    """
    use: lineComponents without XFMR3
    #
    Purpose: find opposite branches (start by hndBr)
            All taps are ignored. Close switches are included.
            Branches out of service are ignored.

        Args :
            hndBr :  branch handle (start)
            typeConsi: type considered as component of line
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]

        returns :
            list of opposite branches

        Raises:
            OlrxAPIException
    """
    if len(typeConsi)==0:
        typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]
    else:
        typeConsi1 = typeConsi[:]
    #
    bres = lineComponents(hndBr,typeConsi1)
    res = set()
    for bres1 in bres:
        brmote = bres1[len(bres1)-1]
        res.add(brmote)
    #
    return list(res)

#
def getOppositeBranch(hndBr,typeConsi): # all type of branch
    """
    use: getRemoteTerminals - support all type of branch
    #
    Purpose: find opposite branches (start by hndBr)
            All taps are ignored. Close switches are included.
            Branches out of service are ignored.

        Args :
            hndBr :  branch handle (start)
            typeConsi: type considered as component of line
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]

        returns :
            list of opposite branches

        Raises:
            OlrxAPIException
    """
    if len(typeConsi)==0:
        typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]
    else:
        typeConsi1 = typeConsi[:]

    ba,bra = getRemoteTerminals(hndBr,typeConsi1)
    #
    return bra

#
def lineComponents(hndBr,typeConsi): # Without XFMR3
    """
    Purpose: find list branches of line (start by hndBr)
            All taps are ignored. Close switches are included.
            Branches out of service are ignored.

        Args :
            hndBr :  branch handle (start)
            typeConsi: type considered as component of line
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]

        returns :
            list of branches

        Raises:
            OlrxAPIException
            if road found>1000 => error
    """
    if len(typeConsi)==0:
        typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]
    else:
        typeConsi1 = typeConsi[:]
    #
    b1 = getEquipmentData([hndBr],BR_nBus1Hnd)[0]
    b2 = getEquipmentData([hndBr],BR_nBus2Hnd)[0]
    nTap1 = getEquipmentData([b1],BUS_nTapBus)[0]
    typ1  = getEquipmentData([hndBr],BR_nType)[0]
    if nTap1!=0:
        raise Exception("Impossible to start lineComponents from a Tap bus:" + OlxAPI.FullBusName(b1))
    #
    if TC_XFMR3 in typeConsi1:
        raise Exception("Impossible run lineComponents with XFMR3")
    #
    if typ1 not in typeConsi1:
        return []
    #
    brA_res = [[hndBr]] # result
    bsa = [[b2]]
    bra = [hndBr]# for each direction
    #
    kmax = 0
    while True:
        bra_in = []
        for br1 in bra:
            #
            nTap,br_Self,br_res = branchesNextToBranch(br1, typeConsi1)
            if nTap ==0 or len(br_res)==0: # finish
                for i in range(len(brA_res)):
                    if br1 == brA_res[i][len(brA_res[i])-1]:
                        brA_res[i].append(br_Self)
            else:
                br0 = br_res[0]
                b20 = getEquipmentData([br0],BR_nBus2Hnd)[0]
                #
                if len(br_res)==1: # tap 2
                    #
                    for i in range(len(brA_res)):
                        if (br1 == brA_res[i][len(brA_res[i])-1]) and (b20 not in bsa[i]):
                            brA_res[i].append(br0)
                            bsa[i].append(b20)
                            bra_in.append(br0)  # continue this direction
                else: # tap>3
                    k = -1
                    for i in range(len(brA_res)):
                        if br1 == brA_res[i][len(brA_res[i])-1]:
                            k = i
                            break
                    # add more direction
                    for i in range(1,len(br_res)):
                        bri = br_res[i]
                        b2 = getEquipmentData([bri],BR_nBus2Hnd)[0]
                        if (b2 not in bsa[k]):
                            brd = []
                            brd.extend(brA_res[k])
                            brd.append(bri)
                            brA_res.append(brd)
                            #
                            bsd = []
                            bsd.extend(bsa[k])
                            bsd.append(b2)
                            bsa.append(bsd)
                            #
                            bra_in.append(bri)
                    # continue this direction
                    if (b20 not in bsa[k]):
                        brA_res[k].append(br0)
                        bsa[k].append(b20)
                        bra_in.append(br0)
        # finish or not
        if len(bra_in)==0:
            break
        # continue
        bra = []
        bra.extend(bra_in)
        # check km
        kmax +=1
        if kmax>1000:
            raise Exception("Out of range: try to check network")

    #check finish by non Tap bus
    resF = []
    for bra1 in brA_res:
        br1 = bra1[len(bra1)-1]
        b1 = getEquipmentData([br1],BR_nBus1Hnd)[0]
        nTap1 = getEquipmentData([b1],BUS_nTapBus)[0]
        if nTap1 ==0:
            resF.append(bra1)
    #debug
##    for bra1 in resF:
##        print("--")
##        for br1 in bra1:
##            print(fullBranchName(br1))
    #
    return resF

#
def isMainLine(hndLna):
    if len(hndLna)<=1:
        return True # simple line
    #
    ida = getEquipmentData(hndLna,LN_sID) # id line
    naa = getEquipmentData(hndLna,LN_sName) # name line
    for i in range(len(ida)):
        ida[i] = (ida[i].upper()).replace(" ","")
        naa[i] = (naa[i].upper()).replace(" ","")

    # METHOD 1: Enter the same name in the Name field of the line’s main segments
    for i in range(1,len(hndLna)):
        if (naa[i]!= naa[0]):
            return False

    # METHOD 2:  Include these three characters [T] (or [t]) in the tap lines name
    for i in range(1,len(hndLna)):
        if "[T]" in naa[i]:
            return False

    # METHOD 3: Give the tap lines circuit IDs that contain the capital letter T (or lowercase letter t)
    for i in range(1,len(hndLna)):
        if "T" in ida[i]:
            return False
    #
    return True

#
def getRemoteTerminals(hndBr,typeConsi):
    """
    Purpose: Find all remote end of a branch
             All taps are ignored.
             Close switches are included
             Out of service branches are ignored

        Args :
            hndBr :  branch handle
            typeConsi: type considered as component of branch
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]

        returns :
            bus_res [] list of terminal bus

        Raises:
            OlrxAPIException
    """
    inSer1 = getEquipmentData([hndBr],BR_nInService)[0]
    if inSer1 == 2:
        return [],[]
    #
    if len(typeConsi)==0:
        typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]
    else:
        typeConsi1 = typeConsi[:]
    #
    bus_res = []
    br_res  = []
    #
    setEqui = set()
    equiHnd1 = getEquipmentData([hndBr],BR_nHandle)[0]
    setEqui.add(equiHnd1)
    #
    b123 = getBusByBranch(hndBr)
    b23 = b123[1:]
    #
    while True:
        bn23_in = []
        for b2 in b23:
            #
            nTap = getEquipmentData([b2],BUS_nTapBus)[0]
            allBr1 = getBusEquipmentData([b2],TC_BRANCH)[0]
            alltyp = getEquipmentData(allBr1,BR_nType)
            allBr = []
            for i in range(len(allBr1)):
                if alltyp[i] in typeConsi1:
                    allBr.append(allBr1[i])
            #
            if nTap==0 or len(allBr)==1:
                bus_res.append(b2)
                # check
                for br1 in allBr:
                    equi1 = getEquipmentData([br1],BR_nHandle)[0]
                    if (equi1 in setEqui) and br1!=hndBr:
                        br_res.append(br1)
            else:
                for br1 in allBr:
                    equi1 = getEquipmentData([br1],BR_nHandle)[0]
                    if (equi1 not in setEqui):
                        inSer1 = getEquipmentData([br1],BR_nInService)[0]
                        type1 =  getEquipmentData([br1],BR_nType)[0]
                        statusSW = 1
                        if type1 == TC_SWITCH: # if Switch
                            swHnd = getEquipmentData([br1],BR_nHandle)[0]
                            statusSW =  getEquipmentData([swHnd],SW_nStatus)[0]
                        #
                        if statusSW==0 and len(allBr)==2: # finish
                            bus_res.append(b2)
                            for br2 in allBr:
                                if br2!=br1:
                                    br_res.append(br2)
                        elif statusSW>0 and inSer1==1 :
                            setEqui.add(equi1)
                            b123 = getBusByBranch(br1)
                            bn23_in.extend(b123[1:])
        #update b23
        b23 = list(set(bn23_in))
        if len(b23)==0:
            break
        #
##        print("tempo")
##        for br1 in br_res:
##            print(fullBranchName(br1))
    return list(set(bus_res)), list(set(br_res))

#
def saveString2File(nameFile,sres):
    try:
        text_file = open(nameFile, "w")
        text_file.write(sres)
        text_file.close()
        print("File saved as: " + nameFile)
    except:
        pass

#
def addString2File(nameFile,sres):
    try:
        text_file = open(nameFile, "a")
        text_file.write(sres)
        text_file.close()
    except:
        pass

#
def read_File_text (sfile):
    # read file text to array
    ar = []
    try:
        ins = open( sfile, "r" )
        for line in ins:
            ar.append(line.replace("\n",""))
        ins.close()
    except OSError:
        raise Exception(OSError.strerror)
    #
    return ar
#
def read_File_text_1 (sfile):
    # read file text to str
    sres = ""
    try:
        ins = open(sfile, "r" )
        sres = ins.read()
        ins.close()
    except OSError:
        raise Exception(OSError.strerror)
    #
    return sres

# to update name file output ARGVS.fo,fr,...
def updateNameFile(path,fi,exti,fo,exto):
    #
    res = fo
    if res== "":
        base = os.path.basename(fi).upper().replace(exti.upper(),"")
        if exto.upper()!= exti.upper():
            res = os.path.join (path, base + exto.upper())
        else:
            res = os.path.join (path, base + "_1"+exto.upper())
    #
    if not res.upper().endswith(exto.upper()):
        res += exto.upper()
    #
    return res

#
def read_File_csv(fileName, delim):
    res = []
    ar = read_File_text(fileName)
    #
    for a1 in ar:
        if a1.replace(" ","") != "":
            v1 = str(a1).split(delim)
            res.append(v1)
    return res

#
def read_File_excell(fileName):
    """
    read file excel
    return ws => work sheet active
    need to install openpyxl
    """
    import openpyxl
    wb = openpyxl.load_workbook(fileName)
    ws = wb.active
    return ws

#
def read_File_csv_as_Excell(fileName, delim):
    """
    read file csv
            return as work sheet excell
    need to install openpyxl
    """
    import openpyxl,csv
    wb = openpyxl.Workbook()
    ws = wb.active
    #
    with open(fileName) as f:
        reader = csv.reader(f, delimiter=delim)
        for row in reader:
            ws.append(row)
    return ws

#
def deleteFile(sfile):
    try:
        if os.path.isfile(sfile):
            os.remove(sfile)
    except:
        pass
#
def saveAsOlr(fileNew):
    fileNew = fileNew.replace("/","\\")
    if fileNew:
        if OLXAPI_FAILURE == OlxAPI.SaveDataFile(fileNew):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        return True
    return False

#
def compare2FileText(file1,file2):
    # compare 2 file Text:
    # returns: arrays of number of line different
    ar1 = read_File_text(file1)
    ar2 = read_File_text(file2)
    dif = []
    for i in range(max(len(ar1),len(ar2))):
        a1 = ""
        a2 = ""
        try:
            a1 = ar1[i]
        except:
            pass
        try:
            a2 = ar2[i]
        except:
            pass
        #
        a1 = a1.strip()
        a2 = a2.strip()
        if a1!=a2:
            dif.append(i+1)
    return dif

#
def unit_test_compare(PATH_FILE,PY_FILE,sres):
    """
    tool for unit test
    - save result to file.txt
    - compare result with REF
    """
    fileRes = os.path.join(PATH_FILE,PY_FILE.replace(".py","_ut.txt"))
    #
    saveString2File(fileRes,sres)
    # compare.
    fileREF = fileRes.replace(".txt","_REF.txt")
    #
    dif = compare2FileText(fileREF,fileRes) # return array of line number with difference

    if len(dif)==0:
        print("\nPASS unit test: "+ PY_FILE+ "("+ os.path.basename(fileRes)+ "=="+os.path.basename(fileREF)+")")
    else:
        print("\nPROBLEM unit test: " + PY_FILE+ "("+ os.path.basename(fileRes)+ "!="+os.path.basename(fileREF)+")")
        print("\tdifferences in lines: ",dif)

#
def setData(hnd,paraCode,value):
    """
     Set data for equipment (bus,line,...)
        Args :
            hnd:  handle element
            nParaCode = code data (BUS_,LN_,...)
            value: value to be set

        Returns:
            None

        Raises:
            OlrxAPIException
    """
    vt = paraCode//100
    val1 = setValType(vt,value)
    #
    if OLXAPI_OK != OlxAPI.SetData( c_int(hnd), c_int(paraCode), byref(val1) ):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

#
def postData(hnd):
    """Post network object data
    """
    # Validation
    if OLXAPI_OK !=  OlxAPI.PostData(c_int(hnd)):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return OLXAPI_OK

#
def convert_LengthUnit(len_unit1,len_unit):
    """
    convert unit from [len_unit] => [len_unit1]
    """
    if len_unit1 =="" or len_unit1 == len_unit:
        return len_unit,1
    #
    dictC = {"ft_km":0.0003048, "kt_km": 1.852  , "mi_km":1.609344, "m_km":0.001, \
             "km_ft":3280.8399, "km_kt": 0.53996, "km_mi":0.621371, "km_m":1000 , "km_km":1}
    #
    s1 = len_unit +"_km"
    s2 = "km_"+len_unit1
    val = dictC[s1] * dictC[s2]
    #
    return len_unit1, val

#
def getSCVoltage(hnd, style ):
    #style (c_int)     : voltage result style
##                    1: output 012 sequence voltage in rectangular form
##                    2: output 012 sequence voltage in polar form
##                    3: output ABC phase voltage in rectangular form
##                    4: output ABC phase voltage in polar form
    vd9Mag  = (c_double*9)(0.0)
    vd9Ang  = (c_double*9)(0.0)
    if ( OLXAPI_FAILURE == OlxAPI.GetSCVoltage( hnd, vd9Mag, vd9Ang, style) ) :
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    #
    return vd9Mag[:],vd9Ang[:]

#
def getSCVoltage_A(hnd):
    mag,ang = getSCVoltage(hnd, style=4)
    return mag[0],ang[0]

#
def getSCVoltage_p(hnd):
    mag,ang = getSCVoltage(hnd, style=2)
    return mag[1],ang[1]

#
def getSCCurrent(hnd,style):
##    style (c_int)      : current result style
##                      1: output 012 sequence current in rectangular form
##                      2: output 012 sequence current in polar form
##                      3: output ABC phase current in rectangular form
##                      4: output ABC phase current in polar form
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    if  OLXAPI_FAILURE == OlxAPI.GetSCCurrent( hnd, vd12Mag, vd12Ang, style):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    #
    return vd12Mag[:],vd12Ang[:]

#
def getSCCurrent_A(hnd):
    mag,ang = getSCCurrent(hnd,style=4)
    return vd12Mag[0],ang[0]

#
def getSCCurrent_p(hnd):
    mag,ang = getSCCurrent(hnd,style=2)
    return mag[1],ang[1]

#
def gui_Error(sTitle,sMain):
    root = tk.Tk()
    try:
        root.wm_iconbitmap(OLXAPI_DLL_PATH+"ASPEN.ico")
    except:
        pass
    #
    root.withdraw()
    tkm.showerror(sTitle,sMain)
    root.destroy()

#-----------
def getStrNow():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

#
def fullBranchName(hbr):
    sa = str(OlxAPI.FullBranchName(hbr)).rsplit()
    sres = ""
    for s1 in sa:
        sres += s1 +" "
    return sres
#
def fullBusName(bhnd):
    return OlxAPI.FullBusName(bhnd)
#
def doFault(ehnd,fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
    """
    fltConn = [3LG, 2LG, 1LG, LL]
        fltConn =[1,0,0,0]   => 3LG
        fltConn =[1,0,1,0]   => 3LG, 1LG
    fltOpt:
        fltOpt[0]      - Bus or Close-in
        fltOpt[1]      - Bus or Close-in w/ outage
        fltOpt[2]      - Bus or Close-in with end opened
        fltOpt[3]      - Bus or Close#-n with end opened w/ outage
        fltOpt[4]      - Remote bus
        fltOpt[5]      - Remote bus w/ outage
        fltOpt[6]      - Line end
        fltOpt[7]      - Line end w/ outage
        fltOpt[8]      - Intermediate %
        fltOpt[9]      - Intermediate % w/ outage
        fltOpt[10]     - Intermediate % with end opened
        fltOpt[11]     - Intermediate % with end opened w/ outage
        fltOpt[12]     - Auto seq. Intermediate % from [*] = 0
        fltOpt[13]     - Auto seq. Intermediate % to [*] = 0
        fltOpt[14]     - Outage line grounding admittance in mho [***] = 0.
    outageOpt:
        outageOpt[0]   - one at a time
        outageOpt[0]   - two at a time
        outageOpt[0]   - all at once
        outageOpt[0]   - breaker failure (**)
    outageLst:         - (c_int*100): list of handles of branches to be outaged; 0 terminated
    clearPrev:         - clear previous result flag. 1 - set; 0 - reset
    fltR, fltX:        - fault resistance,reactance, in Ohm
    """
    fltConn1   = (c_int*4)(0)
    for i in range(len(fltConn)):
        fltConn1[i] = fltConn[i]
    #
    fltOpt1    = (c_double*15)(0)
    for i in range(len(fltOpt)):
        fltOpt1[i] = fltOpt[i]
    #
    outageOpt1 = (c_int*4)(0)
    for i in range(len(outageOpt)):
        outageOpt1[i] = outageOpt[i]

    outageLst1 = (c_int*100)(0)
    for i in range(len(outageLst)):
        outageLst1[i] = outageLst[i]
    #
    fltR1 = c_double(fltR)
    fltX1 = c_double(fltX)
    #
    clearPrev1 = c_int(clearPrev)
    #
    if OLXAPI_FAILURE == OlxAPI.DoFault(ehnd, fltConn1, fltOpt1,outageOpt1, outageLst1, fltR1, fltX1, clearPrev1):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
#
def pick1stFault():
    if OLXAPI_FAILURE == OlxAPI.PickFault(c_int(SF_FIRST),c_int(9)):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
#
def pickNextFault():
    return OlxAPI.PickFault(c_int(SF_NEXT),c_int(9))

#
class Gui_Select(tk.Frame):
    def __init__(self, parent,w1,xOK,yOK,xCom,yCom,data,select):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.data = data
        self.select = select
        self.initGUI(w1,xOK,yOK,xCom,yCom)
    #
    def initGUI(self,w1,xOK,yOK,xCom,yCom):
        #
        self.cb = ttk.Combobox(self.parent,width=w1, values=self.data)
        self.cb.place(x=xCom, y=yCom)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
        #
        button_OK = tk.Button(self.parent,text =   '     OK     ',command=self.run_OK)
        button_OK.place(x=xOK, y=yOK)
    #
    def run_OK(self):
        x = self.cb.get()
        self.select.append(x)
        self.parent.destroy()

#
class BusSearch:
    """
    Bus search by (name,kV) or by bus Number
        - if (name,kV) is OK => return OlxAPI.FindBus(bname,kV)
                else: return the nearest bus (select by GUI or not)
        - if number if OK => return OlxAPI.FindBusNo(bNum)
                else: return the nearest bus (select by GUI or not)

        syntax:
            bs = OlxAPILib.BusSearch(gui) gui =1 select by GUI, gui = 0 select the nearest bus (result[0])
            bs.setBusNameKV("nhs",132)
            or bs.setBusNumber(250)
            bhnd = bs.runSearch()

        sample: [ALASKA-33, ALuBAZ-33, alasaha-33, ARKANSAS-33, ARIRONA-132]
            (ar,32) => [ARKANSAS-33]
            (ar,0)  => [ARKANSAS-33,ARIRONA-132]
            (alz,0) => [ALuBAZ-33]
            (ub,0)  => [ALuBAZ-33]
            ( ,35)  => [ALASKA-33, ALuBAZ-33, alasaha-33, ARKANSAS-33]

    """
    def __init__(self,gui):
        self.gui = gui
        self.bhnd = []
        #
        self.busNum  = -1
        self.busName = ""
        self.kV      = 0
        #
        self.flagNum = False
        self.flagNameKV = False
        #
        self.sInput = ""
        self.ba = []
    #
    def setGui(self,gui):
        self.gui = gui
    #
    def get_resultArray(self):
        """ return array result
        """
        return self.ba
    #
    def stringResult(self):
        res = ""
        if self.sInput!="":
            res = self.sInput +"\n\t"+ self.sFound
        #
        if self.sSelected != "":
            res+= "\n\t" + self.sSelected
        #
        return res

    #
    def setBusNameKV(self,busName,kV):
        if busName!="" or kV>0:
            self.busName =  busName
            self.kV = kV
            self.flagNameKV = True
            self.flagNum = False
    #
    def setBusNumber(self,busNum):
        if busNum>=0:
            self.busNum =  busNum
            self.flagNum = True
            self.flagNameKV = False
    #
    def runSearch(self):
        self.sSelected = ""
        self.exact = False
        self.ba = []
        #
        if self.flagNameKV:
            return self.__searchBy_NameKv()
        #
        if self.flagNum:
            return self.__searchBy_BusNumber()
        #
##        print("@: Not enough input for search")
        return -1
    #
    def __searchBy_BusNumber(self):
        self.sInput = "Bus search by (bus Number) = " + str(self.busNum)
        #
        bhnd = OlxAPI.FindBusNo(self.busNum)
        if bhnd>0:
            self.exact = True
            self.ba = [bhnd]
            return self.__selectEnd()
        #
        self.__getData()
        num = self.num [:] #copy
        num.append(self.busNum)
        num.sort()
        i = num.index(self.busNum)
        if i==0:
            return self.__selectEnd2([num[1]])
        if i==len(num)-1:
            return self.__selectEnd2([num[i-1]])
        #
        return self.__selectEnd2([num[i-1],num[i+1]])

    def __searchBy_NameKv(self):
        self.sInput = "Bus search by (name, kV) = " + "("+ self.busName + " , " + str(self.kV) +")"
        #
        bhnd  = OlxAPI.FindBus(self.busName,self.kV)
        if bhnd>0:
            self.exact = True
            self.ba = [bhnd]
            return self.__selectEnd()
        #
        self.__getData()
        #
        bnameUp = self.busName.upper().replace(" ","")

        # just search by KV ----------------------------------------------------
        if bnameUp=="":
            kr = self.__searchBy_Kv(self.kr0,self.kV)
            return self.__selectEnd1(kr)
        # start with& kv==
        kr = self.__searchBy_Name_Start(self.kr0,bnameUp)
        kr2 = self.__searchBy_Kv_1(kr,self.kV,deltaKv = 1.1)
        if len(kr2)>0:
            return self.__selectEnd1(kr2)
        # in & kv==
        kr3 = self.__searchBy_Name_in(bnameUp,self.kr0)
        kr4 = self.__searchBy_Kv_1(kr3,self.kV,deltaKv = 1.1)
        if len(kr4)>0:
            return self.__selectEnd1(kr4)

        # start with & kV!=
        if len(kr)>0:
            if len(kr)>1 and self.kV>0:
                kr = self.__searchBy_Kv(kr,self.kV)
            return self.__selectEnd1(kr)
        # in & kv!=
        if len(kr3)>0:
            # IN
            if len(kr3)>1 and self.kV>0:
                kr3 = self.__searchBy_Kv(kr3,self.kV)
            if len(kr3)>0:
                return self.__selectEnd1(kr3)

        #------------------------------------another algo
        for i in range(len(self.busName)):
            bname_i = self.busName[i:]
            kr = self.__searchBy_NameKv_1(bname_i,self.kV)
            if len(kr)>0:
                return self.__selectEnd1(kr)

        return self.__selectEnd1([])

    def __searchBy_NameKv_1(self,bname,kV):
        bnameUp = bname.upper().replace(" ","")
        #
        bn1 = bnameUp[:1]
        kr = self.__searchBy_Name_Start(self.kr0,bn1)
        if len(kr)==0:
            # update na
            kr1 = []
            for i in self.kr0:
                if len(self.na[i])>1:
                    self.na[i] = self.na[i][1:]
                    kr1.append(i)
            self.kr0 = kr1[:] #copy
            if len(self.kr0)==0: # not found
                return []
            #
            return self.__searchBy_NameKv_1(bname,kV)
        #
        if kV>0:
            kr = self.__searchBy_Kv(kr,kV)
        #
        if len(kr)==0:
            return []
        #
        if len(kr)==1 or len(bnameUp)==1: # finish search
            return kr
        #
        for j in range(1,len(bnameUp)):
            # search 2 start with 2s
            bn1 = bnameUp[j:j+1]
            kr2 = self.__searchBy_Name_Start(kr,bn1)
            #
            if len(kr2)==0: # start with 1s, 1s in
                kr2 = self.__searchBy_Name_in(bn1,kr)
            if len(kr2)==0:
                return kr
            #
            if len(kr2)==1 or len(bnameUp)==(j+1): # finish search
                return kr2
            kr = kr2[:] # copy
        #
        return kr
    #
    def __searchBy_Kv(self,kri,kV):
        dkv = [0.1,1,10,50,100,200,1000]
        for dkv1 in dkv:
            kr = self.__searchBy_Kv_1(kri,kV,dkv1)
            if len(kr)>0:
                return kr
        return []
    #
    def __searchBy_Kv_1(self,kri,kV,deltaKv):
        kr1 = []
        for i in kri:
            if abs(kV-self.kv[i])<deltaKv:
                kr1.append(i)
        return kr1
    #
    def __searchBy_Name_Start(self,kri,bn1):
        kr1 = []
        for i in kri:
            if (self.na[i]).startswith(bn1):
                kr1.append(i)
                self.na[i] = self.na[i][len(bn1):]
        return kr1
    #
    def __searchBy_Name_in(self,bn1,kr1):
        kr2 = []
        for i in kr1:
            idx = (self.na[i]).find(bn1)
            if idx>=0:
                kr2.append(i)
                self.na[i] = self.na[i][idx+1:]
        return kr2
    #
    def __selectEnd1(self,kr):
        self.ba = []
        for k in kr:
            self.ba.append(self.bhnd[k])
        return self.__selectEnd()
    #
    def __selectEnd2(self,num):
        self.ba = []
        for n1 in num:
            bhnd = OlxAPI.FindBusNo(n1)
            self.ba.append(bhnd)
        return self.__selectEnd()
    #
    def __selectEnd(self):
        if len(self.ba) ==0: # no result
            self.ba = [0]
        #
        if self.gui==1:
            root = tk.Tk()
            try:
                root.wm_iconbitmap(OLXAPI_DLL_PATH+"ASPEN.ico")
            except:
                pass
        #
        if len(self.ba)==1:
            if self.ba[0]==0:
                self.sFound = "No bus found!"
            else:
                if self.exact:
                    self.sFound  = "Bus found : "
                else:
                    self.sFound  = "Bus found (nearest) : "
                self.sFound += OlxAPI.FullBusName(self.ba[0])
            #
            if self.gui ==1:
                root.withdraw()
                #
                tkm.showinfo("Bus search",self.sInput + "\n"+ self.sFound)
                root.destroy()
            #
            return self.ba[0]
        # multipe result
        data = []
        for bi in self.ba:
            data.append (OlxAPI.FullBusName(bi))
        #
        self.sFound = "Bus found (nearest): "
        for d1 in data:
            self.sFound += "\n\t\t"+ d1
        #
        if self.gui==0:
            self.sSelected = "Selected : " + OlxAPI.FullBusName(self.ba[0])
            return self.ba[0]
        #
        try:
            root.wm_title("Bus search")
            ws = root.winfo_screenwidth()
            hs = root.winfo_screenheight()
            root.geometry("300x200+"+ str(int((ws-300)/2))+"+" + str(int((hs-200)/2)))
            w1 = tk.Label(root, text= self.sInput)
            w1.place(x=40, y=10)
            #
            w2 = tk.Label(root, text="Select the nearest bus")
            w2.place(x=40, y=40)
            #
            select = []
            d = Gui_Select(parent=root,w1=25,xOK=120,yOK=150,xCom=40,yCom=70,data=data,select=select)
            root.mainloop()
            i = data.index(select[0])
            self.sSelected = "Selected : " + OlxAPI.FullBusName(self.ba[i])
            return self.ba[i]
        except:
            self.sSelected = "Selected : " + OlxAPI.FullBusName(self.ba[0])
            return self.ba[0]
    #
    def __getData(self):
        if len(self.bhnd)==0:
            self.bhnd = getEquipmentHandle(TC_BUS)
            self.kv   = getEquipmentData(self.bhnd,BUS_dKVnominal)
            self.name = getEquipmentData(self.bhnd,BUS_sName     )
            for i in range(len(self.name)):
                self.name[i]= (self.name[i].upper()).replace(" ","")
            self.num  = getEquipmentData(self.bhnd,BUS_nNumber)
        #
        self.na = self.name[:] #copy
        #
        self.kr0 = []
        for i in range(len(self.na)):
            self.kr0.append(i)

#
class BranchSearch:
    """
    Branch search
        - by (bus1,bus2,CktID)
                bus defined by (name,kv) or by (bus number)
        - by (Name Branch)

    syntax:
        bs = OlxAPILib.BranchSearch(gui) ''gui = 1: select by GUI gui = 0: select the nearest branch (result[0])
        bs.setBusNameKV1(name1,kV1)
        bs.setBusNameKV2(name2,kV2)
        bs.setCktID(CktID)
        |
        bs.setBusNum1(busNum1)
        bs.setBusNum2(busNum2)

        brhnd = bs.runSearch()

    """
    def __init__(self,gui):
        self.gui = gui
        #
        self.busName1 = ""
        self.busName2 = ""
        #
        self.buskV1 = 0
        self.buskV2 = 0
        #
        self.busNum1 = -1
        self.busNum2 = -1
        #
        self.CktID = ""
        #
        self.bus1 = []
        self.bus2 = []
        self.flagBus1 = False
        self.flagBus2 = False
        self.flagNameBr = False
        self.sInput = ""
        #
        self.nameBr = ""
        self.s1 = ""
        self.s2 = ""
        self.sid = "CktID = "
        #
        self.bs = BusSearch(gui=0)
        self.bra = []
    #
    def setGui(self,gui):
        self.gui = gui
    #
    def setNameBranch(self,nameBr):
        if nameBr!="":
            self.nameBr = nameBr
            self.nameBrUp = nameBr.replace(" ","").upper()
            self.flagNameBr = True
            self.flagBus1 = False
            self.flagBus2 = False
    #
    def setCktID(self,CktID):
        self.CktID =  str(CktID)
        self.sid = "CktID = " + self.CktID
    #
    def setBusNameKV1(self,busName1,kV1):
        if busName1!="" or kV1>0:
            self.busName1 =  busName1
            self.kV1 = kV1
            #
            self.bs.setBusNameKV(busName1,kV1)
            self.bs.runSearch()
            self.bus1 = self.bs.get_resultArray()
            #
            self.flagBus1 = True
            #
            self.s1 = "bus1 (name,kV) = (" +busName1 + "," +str(kV1) + ")"
    #
    def setBusNameKV2(self,busName2,kV2):
        if busName2!="" or kV2>0:
            self.busName2 =  busName2
            self.kV2 = kV2
            #
            self.bs.setBusNameKV(busName2,kV2)
            self.bs.runSearch()
            self.bus2 = self.bs.get_resultArray()
            #
            self.flagBus2 = True
            #
            self.s2 = "bus2 (name,kV) = (" +busName2 + "," +str(kV2) + ")"
    # use intern
    def setBusHnd1(self,bhnd1):
        if bhnd1>0:
            self.bus1 = [bhnd1]
            self.flagBus1 = True
    # use intern
    def setBusHnd2(self,bhnd2):
        if bhnd2>0:
            self.bus2 = [bhnd2]
            self.flagBus2 = True
    #
    def setBusNum1(self,busNum1):
        if busNum1>=0:
            self.busNum1 =  busNum1
            #
            self.bs.setBusNumber(busNum1)
            self.bs.runSearch()
            self.bus1 = self.bs.get_resultArray()
            #
            self.flagBus1 = True
            #
            self.s1 = "bus1 (number) = (" + str(busNum1) + ")"
    #
    def setBusNum2(self,busNum2):
        if busNum2>0:
            self.busNum2 =  busNum2
            #
            self.bs.setBusNumber(busNum2)
            self.bs.runSearch()
            self.bus2 = self.bs.get_resultArray()
            #
            self.flagBus2 = True
            #
            self.s2 = "bus2 (number) = (" + str(busNum2) + ")"
    #
    def get_resultArray(self):
        return self.bra
    #
    def runSearch(self):
        self.sSelected = ""
        self.sFound = ""
        self.sSelected = ""
        self.bra = []
        #
        if self.flagBus1 and self.flagBus2:
            return self.__searchBy_bus12_sID()
        #
        if self.flagNameBr:
            return self.__searchBy_nameBr()
        #
##        print("@Branch search: Not enough input for search")
        return -1
    #
    def __searchBy_nameBr(self):
        self.sInput = ("Branch search:").ljust(80)
        self.sInput += "\n\t" + ("branch Name = "+ self.nameBr).ljust(60)
        #
        dictName = {}
        setEhnd = set()
        busa = getEquipmentHandle(TC_BUS)
        for b1 in busa:
            bra = getBusEquipmentData([b1],TC_BRANCH)[0]
            #
            ehand = getEquipmentData(bra,BR_nHandle)
            for i in range(len(ehand)):
                br1 = bra[i]
                e1 = ehand[i]
                if e1 not in setEhnd:
                    na1 = getDataByBranch(br1,"sName")
                    na1 = na1.replace(" ","").upper()
                    if na1!="":
                        dictName[br1] = na1
                    setEhnd.add(e1)
        #
        return self.__searchBy_nameBr1(dictName,self.nameBrUp)

    # search------------------------------------
    def __searchBy_nameBr1(self,dictName,nameBr):
        #exact
        for (k,v) in dictName.items():
            if v== nameBr:
                self.bra.append(k)
        # start with
        if len(self.bra)==0:
            for (k,v) in dictName.items():
                if v.startswith(nameBr):
                    self.bra.append(k)
        # in
        if len(self.bra)==0:
            for (k,v) in dictName.items():
                if nameBr in v:
                    self.bra.append(k)
        #
        if len(self.bra)==0:
            nameBr1 = nameBr[:len(nameBr)-1]
            if len(nameBr1)>0:
                return self.__searchBy_nameBr1(dictName,nameBr1)
        #
        return self.__selectEnd()

   #
    def __searchBy_bus12_sID(self):
        self.sInput = ("Branch search:").ljust(80) +"\n"
        self.sInput += ("\t"+self.s1).ljust(60) +"\n"
        self.sInput += ("\t"+self.s2).ljust(60) +"\n"
        self.sInput += ("\t"+self.sid).ljust(75)
        #
        for b1 in self.bus1:
            br = getBusEquipmentData([b1],TC_BRANCH)[0]
            for br1 in br:
                if len(self.bus2)==0:
                    self.__test_sID(br1)
                else:
                    b2t = getEquipmentData([br1],BR_nBus2Hnd)[0]
                    if b2t in self.bus2:
                        self.__test_sID(br1)
        return self.__selectEnd()
    #
    def __test_sID(self,br1):
        if self.CktID!="":
            id1 = getDataByBranch(br1,"sID")
        #
        if self.CktID == "" or id1 == self.CktID :
            self.bra.append(br1)

    def string_result(self):
        res = ""
        if self.sInput!="":
            res = self.sInput + "\n" + self.sFound
        if self.sSelected!="":
             res += "\n" + self.sSelected
        #
        return res
    #
    def __selectEnd(self):
        #
        if len(self.bra) ==0: # no result
            self.bra = [0]
        #
        if self.gui==1:
            root = tk.Tk()
            try:
                root.wm_iconbitmap(OLXAPI_DLL_PATH+"ASPEN.ico")
            except:
                pass
        #
        if len(self.bra)==1:
            if self.bra[0]==0:
                self.sFound = "No branch found!"
            else:
                self.sFound = "Branch found:\n\t" + fullBranchName(self.bra[0])
            #
            if self.gui ==1:
                root.withdraw()
                tkm.showinfo("Branch search:",self.sInput + "\n"+self.sFound)
                root.destroy()
            return self.bra[0]

        # multipe result
        data = []
        for bi in self.bra:
            data.append(fullBranchName(bi))
        #
        self.sFound = "Branch found:"
        for d1 in data:
            self.sFound += "\n\t"+ d1

        if self.gui==0:
            self.sSelected = "Selected:\n\t" + fullBranchName(self.bra[0])
            return self.bra[0]
        #
        try:
            root.wm_title("Branch search")
            ws = root.winfo_screenwidth()
            hs = root.winfo_screenheight()
            root.geometry("420x200+"+ str(int((ws-420)/2))+"+" + str(int((hs-200)/2)))
            w1 = tk.Label(root, text = self.sInput)
            w1.place(x=30, y=10)
            #
            w2 = tk.Label(root, text="Select the nearest branch")
            w2.place(x=40, y=70)
            #
            select = []
            d = Gui_Select(parent=root,w1=50,xOK=180,yOK=150,xCom=40,yCom= 90,data=data,select=select)
            root.mainloop()
            i = data.index(select[0])
            self.sSelected = "Selected:\n\t" + fullBranchName(self.bra[i])
            return self.bra[i]
        except:
            self.sSelected = "Selected:\n\t" + fullBranchName(self.bra[0])
            return self.bra[0]







