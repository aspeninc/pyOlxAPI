"""
Common ASPEN OlxAPI Python routines.
"""
from __future__ import print_function

__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.3.5"
#
import sys,os
import OlxAPI
from OlxAPIConst import *
from ctypes import *
import AppUtils

#
dictCode_BR_sID   = {TC_LINE:LN_sID  , TC_SWITCH:SW_sID  , TC_PS: PS_sID  , TC_XFMR:XR_sID  , TC_XFMR3:X3_sID  , TC_SCAP:SC_sID   }
dictCode_BR_sName = {TC_LINE:LN_sName, TC_SWITCH:SW_sName, TC_PS: PS_sName, TC_XFMR:XR_sName, TC_XFMR3:X3_sName, TC_SCAP:SC_sName }
dictCode_BR_sC    = {TC_LINE:"L"     , TC_SWITCH:"W"     , TC_PS: "P"     , TC_XFMR:"T"     , TC_XFMR3:"X"     , TC_SCAP:"S" }
dictCode_BR_RAT   = {TC_LINE:1       , TC_SWITCH:7       , TC_PS: 3       , TC_XFMR:2       , TC_XFMR3:10}
dictCode_BR_nBus  = {TC_LINE:[LN_nBus1Hnd,LN_nBus2Hnd],\
                     TC_SWITCH:[SW_nBus1Hnd,SW_nBus2Hnd],\
                     TC_PS: [PS_nBus1Hnd,PS_nBus2Hnd], \
                     TC_SCAP:[SC_nBus1Hnd,SC_nBus2Hnd],\
                     TC_XFMR:[XR_nBus1Hnd,XR_nBus2Hnd],\
                     TC_XFMR3 :[X3_nBus1Hnd,X3_nBus2Hnd,X3_nBus3Hnd]}
dictCode_RLY_sID  = {TC_RLYOCG:OG_sID  , TC_RLYOCP:OP_sID  , TC_RLYDSG: DG_sID  , TC_RLYDSP:DP_sID  , TC_FUSE:FS_sID }
dictCode_TC_Name  = {TC_BUS:'BUS' ,TC_LINE:'LINE' , TC_SWITCH:'SWITCH' , TC_PS: 'PHASE SHIFTER' ,\
                    TC_XFMR:'TRANSFORMER 2 WINDINGS' , TC_XFMR3:'TRANSFORMER 3 WINDINGS' , TC_SCAP:'SERIE CAPACITOR/REACTOR',\
                    TC_RLYGROUP:'RELAY GROUP', TC_GEN:'GENERATOR',TC_GENUNIT:'GENERATOR UNIT' }
#
dictFltConn       = {"3PH":[1,0,0,0], "2LG_BC":[0,1,0,0],"2LG_CA":[0,2,0,0],"2LG_AB":[0,3,0,0],"1LG_A":[0,0,1,0],"1LG_B":[0,0,2,0],\
                     "1LG_C":[0,0,3,0],"LL_BC":[0,0,0,1],"LL_CA":[0,0,0,2],"LL_AB":[0,0,0,3]}

dictCode_PRSys = {'BASEMVA':SY_dBaseMVA,'BUS':SY_nNObus,'GEN':SY_nNOgen ,'LOAD':SY_nNOload ,'SHUNT':SY_nNOshunt,'LINE': SY_nNOline ,
                  'SERIESRC':SY_nNOseriescap ,'XFMR':SY_nNOxfmr ,'XFMR3':SY_nNOxfmr3 ,'OPS':SY_nNOps ,'MULINE':SY_nNOmutual ,'SWITCH': SY_nNOswitch ,
            'LOADUNIT': SY_nNOloadUnit,'SVD':SY_nNOsvd ,'OC PHASE RELAY':SY_nNOrlyOCP ,'OC GROUND RELAY':SY_nNOrlyOCG ,'DS PHASE RELAY':SY_nNOrlyDSP ,'DS GROUND RELAY':SY_nNOrlyDSG ,
            'D RELAY':SY_nNOrlyD ,'V RELAY': SY_nNOrlyV ,'FUSE':SY_nNOfuse ,'GROUND RECLOSER':SY_nNOrecloserG , 'PHASE: RECLOSER':SY_nNOrecloserP ,
            'SHUNTUNIT':SY_nNOshuntUnit ,'CCGEN':SY_nNOccgen ,'BREAKER':SY_nNObreaker ,'SCHEME':SY_nNOscheme ,'IED': SY_nNOIED}
#
def open_olrFile(olrFile,dllPath='',readonly=1,prt=True):
    """
    Open OLR file in
        olrFile = (str) OLR file
        dllPath = default
    Args:
        readonly (int): open in read-only mode. 1-true; 0-false
    """
    #check file
    AppUtils.checkFileSelected(olrFile,'OLR file path')
    # load dll
    OlxAPI.InitOlxAPI(dllPath,prt)
    #
    val = OlxAPI.LoadDataFile(olrFile,readonly)
    if OLXAPI_OK == val:
        if prt:
            print( "File opened successfully: " + olrFile)
    elif OLXAPI_DATAFILEANOMALIES==val:
        if prt:
            print( "File opened successfully: " + olrFile)
            print(OlxAPI.ErrorString())
    else:
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
#
def open_olrFile_1(dllPath,olrFile,readonly=1,prt=True):
    """
    This function is deprecated
    """
    open_olrFile(olrFile,dllPath=dllPath,readonly=readonly,prt=prt)

#
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
        c_data = c_char_p(OlxAPI.encode3(args["data"]))
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

c_GetDataBuf = create_string_buffer(b'\000' * 10*1024)    # 10 KB buffer for string data
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
        args["data"] = process_GetDataBuf(dataBuf,c_token.value,c_hnd)
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

def process_GetDataBuf(buf,tokenV,hnd):
    """Convert GetData binary data buffer into Python object of correct type
    """
    vt = tokenV//100
    if vt == VT_STRING:
        return (buf.value).decode("UTF-8")
    elif vt in [VT_DOUBLE , VT_INTEGER]:
        return buf.value
    else:
        array = []
        tc = OlxAPI.EquipmentType(hnd)
        if tc == TC_BREAKER and tokenV in [BK_vnG1DevHnd,BK_vnG2DevHnd,BK_vnG1OutageHnd,BK_vnG2OutageHnd]:
            val = cast(buf, POINTER(c_int*MXSBKF)).contents  # int array of size MXSBKF
            for ii in range(0,MXSBKF-1):
                array.append(val[ii])
                if array[ii] == 0:
                    break
        #
        elif (tc == TC_SVD and tokenV == SV_vnNoStep):
            val = cast(buf, POINTER(c_int*8)).contents  # int array of size 8
            for ii in range(0,7):
                array.append(val[ii])
        elif (tc == TC_RLYDSP and tokenV in [DP_vParams,DP_vParamLabels]) or \
             (tc == TC_RLYDSG and tokenV in [DG_vParams,DG_vParamLabels]):
            # String with tab delimited fields
            return ((cast(buf, c_char_p).value).decode("UTF-8")).split("\t")
        #
        else:
            if tc == TC_GENUNIT and tokenV in [GU_vdR,GU_vdX]:
                count = 5
            elif tc == TC_LOADUNIT and tokenV in [LU_vdMW,LU_vdMVAR] :
                count = 3
            elif tc == TC_SVD and tokenV in [SV_vdBinc,SV_vdB0inc]:
                count = 3
            elif tc == TC_LINE and tokenV == LN_vdRating:
                count = 4
            elif tc == TC_RLYGROUP and tokenV == RG_vdRecloseInt:
                count = 3
            elif tc == TC_RLYOCG and tokenV == OG_vdDirSetting:
                count = 8
            elif tc == TC_RLYOCP and tokenV == OP_vdDirSetting:
                count = 8
            elif tc == TC_RLYOCG and tokenV == OG_vdDirSettingV15:
                count = 9
            elif tc == TC_RLYOCP and tokenV == OP_vdDirSettingV15:
                count = 9
            elif tc == TC_RLYOCG and tokenV in [OG_vdDTPickup,OG_vdDTDelay]:
                count = 5
            elif tc == TC_RLYOCP and tokenV in [OP_vdDTPickup,OP_vdDTDelay]:
                count = 5
            elif tc == TC_RLYDSG and tokenV == DG_vdParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSG and tokenV in [DG_vdDelay,DG_vdReach,DG_vdReach1]:
                count = MXZONE
            elif tc == TC_RLYDSP and tokenV == DP_vParams:
                count = MXDSPARAMS
            elif tc == TC_RLYDSP and tokenV in[DP_vdDelay,DP_vdReach,DP_vdReach1]:
                count = MXZONE
            elif tc == TC_CCGEN and tokenV in [CC_vdV,CC_vdI,CC_vdAng]:
                count = MAXCCV
            elif tc == TC_BREAKER and tokenV in [BK_vdRecloseInt1,BK_vdRecloseInt2]:
                count = 3
            elif tc == TC_MU and tokenV in [MU_vdX,MU_vdR,MU_vdFrom1,MU_vdFrom2,MU_vdTo1,MU_vdTo2]:
                count = 5
            else:
                count = MXDSPARAMS
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
    runOpt = (c_int*7)(1,1,1,1,1,1,1)   # OCG, OCP, DSG, DSP, SCHEME, DIFF, VRELAY
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
    #
    baseMVA = getParaSys(cod=SY_dBaseMVA)
    aLine1 = OlxAPI.FullBusName(Bus1Hnd) + " - " + OlxAPI.FullBusName(Bus2Hnd) +\
             ": Z=" + printImpedance(dR,dX,dKV,baseMVA) + " Zo=" + printImpedance(dR0,dX0,dKV,baseMVA) + " L=" + str(dLength)
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
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) +\
                ": Z=" + printImpedance(dRn,dXn,dKV,baseMVA) + " Zo=" + printImpedance(dR0n,dX0n,dKV,baseMVA) + " L=" + str(dL)
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
        aLine = OlxAPI.FullBusName(BusHnd) + " - " + OlxAPI.FullBusName(BusFarHnd) +\
                ": Z=" + printImpedance(dRn,dXn,dKV,baseMVA) + " Zo=" + printImpedance(dR0n,dX0n,dKV,baseMVA) + " L=" + str(dL)
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
            aLine1 = OlxAPI.FullBusName(busHndList[0]) + " - " + OlxAPI.FullBusName(busHndList[1]) +\
            ": Z=" + printImpedance(dR,dX,dKV,baseMVA) + " Zo=" + printImpedance(dR0,dX0,dKV,baseMVA) + " L=" + str(dLength)
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
#
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
def getRelayTime( hndRelay, mult, consider_signalonly):
    sx = create_string_buffer(b'\000' * 128)
    triptime = c_double(0.0)
    if OLXAPI_OK != OlxAPI.GetRelayTime(hndRelay,c_double(mult), c_int(consider_signalonly), byref(triptime),sx):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    #
    return triptime.value, (sx.value).decode("UTF-8")

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
            OlxAPIException
   """
    # get data
    res = []
    vt = paraCode//100
    val1 = setValType(vt,0)
    for ehnd1 in ehnd:
        if ( OLXAPI_FAILURE == OlxAPI.GetData( ehnd1, c_int(paraCode), byref(val1) ) ) :
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #
        res.append( process_GetDataBuf(val1,paraCode,ehnd1) )
    #
    return res
#
def getDataByBranch(hndBr,scode):
    # hndBr: branch handle
    # scode = 'sID' | 'sName' | 'sC'
    # return
    #
    typ  = getEquipmentData([hndBr],BR_nType)[0]
    ehnd = getEquipmentData([hndBr],BR_nHandle)[0]
    #
    if scode == "sID":
        paraCode = dictCode_BR_sID[typ]
    elif scode == "sName":
        paraCode = dictCode_BR_sName[typ]
    elif scode == "sC":
        return dictCode_BR_sC[typ]
    else:
        raise OlxAPIException("Error scode")
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
            OlxAPIException
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
            OlxAPIException
    """
    busHnd = dictCode_BR_nBus[TC_Type]
    #
    busRes = []
    for b1 in busHnd:
        busRes.append(getEquipmentData([ehnd],b1)[0])
    #
    return busRes
#
def getParaSys(cod):
    """
    get para system
    cod = SY_dBaseMVA...
    """
    argsGetData = {}
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = cod
    if OLXAPI_OK != get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return  argsGetData["data"]

#
def get1EquipmentData(ehnd,paraCode):
    """
    Find data of 1 element (line/bus/...)
        Args :
            ehnd:  handle element
            nParaCode []: code data (BUS_,LN_,...)

        Returns:
            data [len(paraCode)]

        Raises:
            OlxAPIException
   """
   # get data
    res = []
    for paraCode1 in paraCode:
        vt = paraCode1//100
        val1 = setValType(vt,0)
        if ( OLXAPI_FAILURE == OlxAPI.GetData( ehnd, c_int(paraCode1), byref(val1) ) ) :
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #
        res.append( process_GetDataBuf(val1,paraCode1,ehnd) )
    return res
#
def get1EquipmentData_try(ehnd,paraCode):
    """
    Find data of 1 element (line/bus/...) with try/except =None if not found
        Args :
            ehnd : handle element
            nParaCode: code data (BUS_,LN_,...)

        Returns:
            data (=-1 if not found)

        Raises:
            OlxAPIException
    """
    try:
        val = getEquipmentData([ehnd],paraCode)[0]
        return val
    except:
        return None

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
            OlxAPIException
   """
    # get data
    res = []
    vt = paraCode//100
    val1 = setValType(vt,0)
    while ( OLXAPI_OK  == OlxAPI.GetData(ehnd,c_int(paraCode), byref(val1) ) ) :
        res.append( process_GetDataBuf(val1,paraCode,ehnd) )
    #
    return res
#
def setValType(vt,value):
    if vt == VT_STRING:
        if value==0:
            return create_string_buffer(b'\000' * 128)
        else:
            return c_char_p( value.encode('UTF-8'))
    elif vt == VT_DOUBLE:
       return c_double(value)
    elif vt == VT_INTEGER or vt==0:
       return c_int(value)
    elif vt == VT_ARRAYDOUBLE or vt==VT_ARRAYSTRING or vt==VT_ARRAYINT:
        return c_GetDataBuf
    #
    raise OlxAPI.OlxAPIException("Error of paraCode")

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
            OlxAPIException
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
            OlxAPIException
   """
    res = []
    hndBr = c_int(0)
    while ( OLXAPI_OK == OlxAPI.GetEquipment(c_int(TC_type), byref(hndBr) )) :
        res.append(hndBr.value)
    return res
#
def getAll_Bus():
    """
    Find all handle of BUS
    """
    return getEquipmentHandle(TC_BUS)
#
def getAll_Area():
    """
    Find all handle of Area
    """
    busHnd = getAll_Bus()
    ar = getEquipmentData(busHnd,BUS_nArea)
    return list(set(ar))
#
def getAll_Zone():
    """
    Find all handle of Area
    """
    busHnd = getAll_Bus()
    ar = getEquipmentData(busHnd,BUS_nZone)
    return list(set(ar))

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
            OlxAPIException
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
            OlxAPIException
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
            OlxAPIException
    """
    bra,bt,ba,equip = getRemoteTerminals(hndBr,typeConsi)
    #
    return bra
#
def linez(hnda):
    """ Calcul total line impedance and length of multisectionLine
    Args:
        hnda: list handle of Line (must be TC_LINE,TC_SCAP,TC_SWITCH or TC_BRANCH on)
    return:
        Z1,Z0,Length
    """
    if type(hnda)!=list:
        hnda = [hnda]
    z1,z0,length = 0,0,0
    setEqui = set()
    u0 = None
    for h1 in hnda:
        typ1 = OlxAPI.EquipmentType(h1)
        if typ1==TC_BRANCH:
            e1 = getEquipmentData([h1],BR_nHandle)[0]
            typ1 = OlxAPI.EquipmentType(e1)
        else:
            e1 = h1
        if typ1 not in {TC_LINE,TC_SCAP,TC_SWITCH}:
            raise Exception('OlxAPILib.linez(hnda) invalid handle inputs: %i (must be TC_LINE,TC_SCAP,TC_SWITCH or TC_BRANCH on)'%typ1)
        if e1 not in setEqui:
            setEqui.add(e1)
            if typ1==TC_LINE:
                R = getEquipmentData([e1],LN_dR)[0]
                X = getEquipmentData([e1],LN_dX)[0]
                R0 = getEquipmentData([e1],LN_dR0)[0]
                X0 = getEquipmentData([e1],LN_dX0)[0]
                l1 = getEquipmentData([e1],LN_dLength)[0]
                u1 = getEquipmentData([e1],LN_sLengthUnit)[0]
                if u0!=None:
                    l1 *= AppUtils.convert_LengthUnit_1(u0,u1)
                else:
                    u0 = u1
            elif typ1==TC_SWITCH:
                R,R0,X,X0 = 0,0,0,0
                l1 = 0
            elif typ1==TC_SCAP:
                R = getEquipmentData([e1],SC_dR)[0]
                X = getEquipmentData([e1],SC_dX)[0]
                R0 = getEquipmentData([e1],SC_dR0)[0]
                X0 = getEquipmentData([e1],SC_dX0)[0]
                l1 = 0
            z1 += complex(R,X)
            z0 += complex(R0,X0)
            length += l1
    return z1,z0,length
#
def tapLineTool(hnd0,prt=False):
    """
    Purpose: Find main sections of Line and sum impedance(Z0,Z1),Length
        All taps are ignored. Close switches are included.
        Branches out of service are ignored
    Args:
        hnd0: handle of start BRANCH/RLYGROUP
            Exception if 1st Bus (of start branch) is a tap bus
            Exception if this BRANCH/RLYGROUP located on XFMR,XFMR3,SHIFTER
        prt: (True/False)
            option print to stdout all details when research mainLine
    return:
        res['mainLineHnd']  = [[]]     List of all Branche handle in the main-line
                                   main-line searched by 3 methods in 'Section 8.9 Help/OneLiner Help Contents'
                                METHOD 1: Enter the same name in the Name field of the line’s main segments
                                METHOD 2: Include these three characters [T] (or [t]) in the tap lines name
                                METHOD 3: Give the tap lines circuit IDs that contain letter T or t
        res['remoteBusHnd']  = []      List of remote buses handle on the mainLine: two for 2-terminal line and >2 for 3-terminal line
        res['remoteRLGHnd']  = []      List of all relay groups handle at the remote end on the mainLine
        res['Z1']            = []      List of positive sequence Impedance of the mainLine
        res['Z0']            = []      List of zero sequence Impedance of the mainLine
        res['Length']        = []      List of sum length of the mainLine
    """
    mainLineHnd,remoteBusHnd,remoteRLGHnd = findLineSections(hnd0,prt)
    z1,z0,length =[],[],[]
    for m1 in mainLineHnd:
        z1i,z0i,lengthi = linez(m1)
        z1.append(z1i)
        z0.append(z0i)
        length.append(lengthi)
    #
    res = dict()
    res['mainLineHnd'] = mainLineHnd
    res['remoteBusHnd'] = remoteBusHnd
    res['remoteRLGHnd'] = remoteRLGHnd
    res['Z0'] = z0
    res['Z1'] = z1
    res['Length'] = length
    return res
#
def findLineSections(hnd0,prt=False):
    """
    Purpose: Find main sections of Line (LINE,SERIESRC,SWITCH) from a BRANCH/RLYGROUP
        All taps are ignored. Close switches are included.
        Branches out of service are ignored
    Args:
        hnd0: handle of start BRANCH/RLYGROUP
            Exception if 1st Bus (of start branch) is a tap bus
            Exception if this BRANCH/RLYGROUP located on XFMR,XFMR3,SHIFTER
    return:
        mainLineHnd  = [[]]     List of all Branche handle in the main-line
                                    main-line searched by 3 methods in 'Section 8.9 Help/OneLiner Help Contents'
                                METHOD 1: Enter the same name in the Name field of the line’s main segments
                                METHOD 2: Include these three characters [T] (or [t]) in the tap lines name
                                METHOD 3: Give the tap lines circuit IDs that contain letter T or t
        remoteBusHnd  = []      List of remote buses handle on the mainLine: two for 2-terminal line and >2 for 3-terminal line
        remoteRLGHnd  = []      List of all relay groups handle at the remote end on the mainLine
    """
    ty1 = OlxAPI.EquipmentType(hnd0)
    if ty1 not in {TC_RLYGROUP,TC_BRANCH}:
        raise Exception("OlxAPILib.findLineSection(hnd0)\n\t hnd0 must be a handle of BRANCH or RLYGROUP")
    if ty1==TC_RLYGROUP:
        hnd0 = getEquipmentData([hnd0],RG_nBranchHnd)[0]
    b1 = getEquipmentData([hnd0],BR_nBus1Hnd)[0]
    nTap1 = getEquipmentData([b1],BUS_nTapBus)[0]
    typ1  = getEquipmentData([hnd0],BR_nType)[0]
    if nTap1!=0 :
        raise Exception("\nImpossible to start OlxAPILib.findLineSection from a Tap bus :" + OlxAPI.FullBusName(b1))
    #
    typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP]
    if typ1 not in typeConsi1:
        raise Exception("\nImpossible to start OlxAPILib.findLineSection from XFMR,XFMR3,SHIFTER: \n\t"+fullBranchName_1(hnd0))
    #
    brA_res = lineComponents(hnd0,[TC_LINE,TC_SWITCH,TC_SCAP])
    mainLineHnd = []
    remoteRLGHnd = []
    remoteBusHnd = []
    if len(brA_res)==1 and len(brA_res[0])==2:# simple result
        r1 = brA_res[0]
        b1 = getEquipmentData(r1[1:],BR_nBus1Hnd)[0]
        #nTap1 = getEquipmentData([b1],BUS_nTapBus)[0]
        mainLineHnd = [r1[:-1]]
        remoteBusHnd = [b1]
        try:
            rg1 = getEquipmentData(r1[1:],BR_nRlyGrp1Hnd)[0]
        except:
            rg1 = None
        remoteRLGHnd = [rg1]
        return mainLineHnd,remoteBusHnd,remoteRLGHnd
    # get name and ID
    ida,naa = [],[]
    for ri in brA_res:
        r1 = ri[:-1]
        ea = getEquipmentData(r1,BR_nHandle)
        ta = getEquipmentData(r1,BR_nType)
        ida1,naa1 = [],[]
        for i in range(len(ea)):
            sID = dictCode_BR_sID[ta[i]]
            sName = dictCode_BR_sName[ta[i]]
            id1 = getEquipmentData([ea[i]],sID)[0]
            na1 = getEquipmentData([ea[i]],sName)[0]
            ida1.append( id1.upper() )
            naa1.append( na1.upper() )
        ida.append(ida1)
        naa.append(naa1)
    #
    if len(brA_res)==1:
        mainLineHnd.append(brA_res[0][:-1])
    else:
        if prt:
            print('Found multi paths from [TERMINAL] '+fullBranchName_1(hnd0))
        for i in range(len(brA_res)):
            r1 = brA_res[i]
            ida1 = ida[i]
            naa1 = naa[i]
            if prt:
                print('\tPath %i:'%(i+1))
                for i1 in range(len(r1)-1):
                    print('\t\t ',str(i1+1).ljust(2),ida1[i1].ljust(3),naa1[i1].ljust(20),fullBranchName_1(r1[i1]))
            #MAIN LINE
            b1 = getEquipmentData(r1[len(r1)-1:],BR_nBus1Hnd)[0]
            nTap1 = getEquipmentData([b1],BUS_nTapBus)[0]
            t1,t2,t3 = False,True,True
            if nTap1==0: #end by no tap bus
                # METHOD 1: Enter the same name in the Name field of the line’s main segments
                setName = set()
                for n1 in naa1:
                    setName.add(n1)
                if len(setName)==1 and naa1[0]:
                    t1 = True
                # METHOD 2:  Include these three characters [T] (or [t]) in the tap lines name
                for na1 in naa1:
                    if "[T]" in na1:
                        t2=False
                # METHOD 3: Give the tap lines circuit IDs that contain letter T or t
                for id1 in ida1:
                    if "T" in id1:
                        t3 = False
                if t1 or (t2 and t3):
                    mainLineHnd.append(r1[:-1])
    #
    if len(mainLineHnd)==0:
        ma2 = []
        for i in range(len(brA_res)):
            r1 = brA_res[i]
            ida1 = ida[i]
            naa1 = naa[i]
            r2 = [r1[0]]
            for j in range(1,len(naa1)):
                if "[T]" in naa1[j] or 'T' in ida1[j]:
                    break
                else:
                    r2.append(r1[j])
            ma2.append(r2)
        ma2.sort(key=lambda x:len(x))
        mainLineHnd.append(ma2[len(ma2)-1])
    # check Tap3
    if len(mainLineHnd)>1:
        ma2 = []
        for m1 in mainLineHnd:
            b1 = getEquipmentData(m1,BR_nBus1Hnd)
            tap1 = getEquipmentData(b1[1:],BUS_nTapBus)
            for t1 in tap1:
                if t1==3:
                    ma2.append(m1)
        if len(ma2)==0:
            ma2.append(mainLineHnd[0])
        mainLineHnd = [m1 for m1 in ma2]
    #
    for m1 in mainLineHnd:
        try:
            rg1 = getEquipmentData(m1[len(m1)-1:],BR_nRlyGrp2Hnd)[0]
        except:
            rg1 = None
        remoteRLGHnd.append(rg1)
        b1 = getEquipmentData(m1[len(m1)-1:],BR_nBus2Hnd)[0]
        remoteBusHnd.append(b1)
    return mainLineHnd,remoteBusHnd,remoteRLGHnd
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

        returns : all path from this hndbr
             [ [hndbr,br1,br2] , [hndbr,br3,br4] , [hndbr,br3,br5] ]   => 3 paths
                                              Bus3 (no tap)
                                              |
                                              |br2
                                              |
                                              Bus(tap)      Bus5 (tap)
                                              |             |
                                              |br1          |br5
                                    hndBr     |     br3     |     br4
        Illustration:       Bus1--------------Bus2----------Bus---------Bus4
                            (no tap)          (tap)         (tap)       (no tap)

    """
    b2t = 0 # not test end by bus
    return lineComponents_0(hndBr,b2t,typeConsi)
#
def lineComponents_0(hndBr,b2t,typeConsi): # Without XFMR3
    """
    Purpose: find list branches of line (start by hndBr) end by bus (b2t) if not b2t<=0
            All taps are ignored. Close switches are included.
            Branches out of service are ignored.

        Args :
            hndBr :  branch handle (start)
            b2t   " bus handle (end)
            typeConsi: type considered as component of line
            [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR]

        returns :
            list of branches

        Raises:
            OlxAPIException
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
    flag1  = getEquipmentData([hndBr],BR_nInService)[0]
    if nTap1!=0 and b2t<=0:
        raise Exception("Impossible to start lineComponents from a Tap bus with no end bus\n\t" + OlxAPI.FullBusName(b1))
    #
    if flag1!=1:
        raise Exception("Impossible to start lineComponents from a out-of-service branch\n\t"+fullBranchName_1(hndBr))
    if TC_XFMR3 in typeConsi1:
        raise Exception("Impossible run lineComponents with XFMR3\n\t"+fullBranchName_1(hndBr))
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
            b2 = getEquipmentData([br1],BR_nBus2Hnd)[0]
            if b2t>0 and b2==b2t:#finish
                pass
            else:
                #
                nTap,br_Self,br_res = branchesNextToBranch(br1, typeConsi1)
                if nTap ==0 or len(br_res)==0: # finish
                    for i in range(len(brA_res)):
                        if br1 == brA_res[i][len(brA_res[i])-1] and br_Self>0:
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
            raise Exception("Out of range. Check the network for anomalies.")
    #check circle by tapbus (branche //)
    resF = []
    for bra1 in brA_res:
        if b2t<=0:
            br1 = bra1[len(bra1)-1]
            br2 = bra1[len(bra1)-2]
            equiHnd = getEquipmentData([br1,br2],BR_nHandle)
            if equiHnd[0]==equiHnd[1]:
                resF.append(bra1)
        else: # test end by b2t
            br1 = bra1[len(bra1)-1]
            b2 = getEquipmentData([br1],BR_nBus2Hnd)[0]
            if b2==b2t:
                resF.append(bra1)
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
    t1,t2,t3 = False,True,True
    # METHOD 1: Enter the same name in the Name field of the line’s main segments
    setName = set()
    for n1 in naa:
        setName.add(n1)
    if len(setName)==1 and naa[0]:
        t1 = True

    # METHOD 2:  Include these three characters [T] (or [t]) in the tap lines name

    for i in range(1,len(hndLna)):
        if "[T]" in naa[i]:
            t2 = False

    # METHOD 3: Give the tap lines circuit IDs that contain the capital letter T (or lowercase letter t)
    for i in range(1,len(hndLna)):
        if "T" in ida[i]:
            t3 = False
    #
    return t1 or (t2 and t3)

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
            br_res   [] list of branch terminal
            bus_res  [] list of terminal bus
            bus_resa [] list of all bus
            equip    [] list of all equipement

        Raises:
            OlxAPIException
    """
    inSer1 = getEquipmentData([hndBr],BR_nInService)[0]
    if inSer1 == 2:
        return [],[],[],[]
    #
    if len(typeConsi)==0:
        typeConsi1 = [TC_LINE,TC_SWITCH,TC_SCAP,TC_PS,TC_XFMR,TC_XFMR3]
    else:
        typeConsi1 = typeConsi[:]
    #
    bus_res = []
    bus_resa = []
    br_res  = []
    #
    setEqui = []
    equiHnd1 = getEquipmentData([hndBr],BR_nHandle)[0]
    setEqui.append(equiHnd1)
    #
    b123 = getBusByBranch(hndBr)
    b23 = b123[1:]
    bus_resa.append(b123[0])
    #
    while True:
        bn23_in = []
        for b2 in b23:
            bus_resa.append(b2)
            #
            nTap = getEquipmentData([b2],BUS_nTapBus)[0]
            allBr1 = getBusEquipmentData([b2],TC_BRANCH)[0]
            alltyp = getEquipmentData(allBr1,BR_nType)
            allBr = []
            for i in range(len(allBr1)):
                if alltyp[i] in typeConsi1:
                    allBr.append(allBr1[i])
            #
            if nTap==0 or len(allBr)==1:# finish
                bus_res.append(b2)
                # check
                for br1 in allBr:
                    equi1 = getEquipmentData([br1],BR_nHandle)[0]
                    if (equi1 in setEqui) and br1!=hndBr:
                        if br1 not in br_res:
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
                                    if br2 not in br_res:
                                        br_res.append(br2)
                        elif statusSW>0 and inSer1==1 :
                            setEqui.append(equi1)
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
    return br_res, list(set(bus_res)), list(set(bus_resa)), setEqui

#
def searchLines(b1,b2,cktid):
    """
    purpose: return list of branch lines beetween bus1/bus2
    args:
            - b1,b2 bus1/2 handle
            - cktid circuit id (str)
    return
            - linea : list of line handle between bus1/bus2 without test cktid
            - linea_wid : list of line handle between bus1/bus2 with test cktid
    """
    linea = []
    linea_wid = []
    if b1<=0 or b2<=0:
        return linea,linea_wid
    #
    br = getBusEquipmentData([b1],TC_BRANCH)[0]
    for br1 in br:
        bra = lineComponents_0(br1,b2,[TC_LINE])
        for bra1 in bra:
            la = getEquipmentData(bra1,BR_nHandle)
            linea.append(la)
    #
    if cktid=='': # no cktid
        return linea,linea_wid
    #
    for la1 in linea:
        test = True
        ida1 = getEquipmentData(la1,LN_sID)
        for id1 in ida1:
            if id1!=cktid:
                test = False
                break
        if test:
            linea_wid.append(la1)
    #
    return linea,linea_wid
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
            OlxAPIException
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
    return True
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
def fullBranchName_1(hbr):
    return AppUtils.deleteSpace1(OlxAPI.FullBranchName(hbr))
#
def fullBranchName(hbr):
    return OlxAPI.FullBranchName(hbr)
#
def fullBusName(bhnd):
    return OlxAPI.FullBusName(bhnd)
#
def fullBusName_1(bhnd):
    return AppUtils.deleteSpace1(OlxAPI.FullBusName(bhnd))
#
def faultDescription():
    return OlxAPI.FaultDescription(0)
#
def faultDescription_1():
    return AppUtils.deleteSpace1(OlxAPI.FaultDescription(0))
#
def version_Build():
    return "OneLiner Version :" + str(OlxAPI.Version()) + ", Build: " + str(OlxAPI.BuildNumber())
#
def doFault_2(ehnd,fltConn, fltOpt_idx, outageOpt, outageLst, fltR, fltX, clearPrev):
    # do fault with fltOpt_idx
    fltOpt = [0]*15
    for id1 in fltOpt_idx:
        fltOpt[id1] = 1
    #
    doFault(ehnd,fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev)
#
def doFault(ehnd,fltConn, fltOpt, outageOpt, outageLst, fltR, fltX, clearPrev):
    """
    fltConn = [3LG, 2LG, 1LG, LL] see getFltConn_doFault
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
        outageOpt[1]   - two at a time
        outageOpt[2]   - all at once
        outageOpt[3]   - breaker failure (**)
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
    # return False if finish
    if OLXAPI_FAILURE==OlxAPI.PickFault(c_int(SF_NEXT),c_int(9)):
        return False
    return True
#
def getAllRelay(hndRlyGroup):
    """
    get all relay handle of a relay group
    """
    ra = []
    hndRelay = c_int(0)
    while OLXAPI_OK == OlxAPI.GetRelay( hndRlyGroup, byref(hndRelay) ) :
        ra.append(hndRelay.value)
    return ra
#
def getRelayID(rlyhnd_a):
    res =[]
    for r1 in rlyhnd_a:
        tc = OlxAPI.EquipmentType(r1)
        sid = getEquipmentData([r1],dictCode_RLY_sID[tc])[0]
        res.append(sid)
    return res

def getOutageList(hnd, maxTiers):
    """
        hnd	(c_int): handle of a bus, branch or a relay group.
        wantedTypes: 1- Line; 2- 2-winding transformer;
                     4- Phase shifter; 8- 3-winding transformer; 16- Switch
                     Serie capacitor/Reactor ?
    """
    res = []
    wantedTypes = c_int(1+2+4+8+16)
    listLen = c_int(0)
    OlxAPI.MakeOutageList(hnd, c_int(maxTiers), wantedTypes, None, pointer(listLen) )
##    print( "listLen=" + str(listLen.value) )
    branchList = (c_int*(5+listLen.value))(0)
    OlxAPI.MakeOutageList(hnd, c_int(maxTiers), wantedTypes, branchList, pointer(listLen) )
    for i in range(listLen.value):
        res.append(branchList[i])
    #
    return res

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
    def runSearch_exact(self):
        if self.flagNameKV:
            bhnd  = OlxAPI.FindBus(self.busName,self.kV)
            if bhnd>0:
                return bhnd
        if self.flagNum:
            bhnd = OlxAPI.FindBusNo(self.busNum)
        return bhnd
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
            setIco(root,"","1LINER.ico")
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
            setIco(root,"","1LINER.ico")
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
#
def FindObj1LPF(obj1LPFStr):
    """
    find object by String ("[BUS] 23 'ILLINOIS' 33 kV")
    return
        hnd: object handle
        tc: object type

        return 0,0 if not found
    """
    hnd = c_int(0)
    if OLXAPI_FAILURE == OlxAPI.FindObj1LPF(obj1LPFStr,hnd):
        return 0,0
    #
    tc = OlxAPI.EquipmentType(hnd)
    #
    return hnd.value,tc

#
def FindObj1LPF_check(obj1LPFStr,tc_code):
    """
    find object by String
    if not found OBJ (tc_code= TC_BUS,TC_LINE,...) => Error
    """
    hnd,tc = FindObj1LPF(obj1LPFStr)
    if tc!=tc_code:
        try:
            name = dictCode_TC_Name[tc_code]
        except:
            name = 'OBJECT'
        #
        err = '\n'+name+ ' not found: "' + obj1LPFStr + '"'
        raise Exception(err)
    #
    return hnd
#
def BoundaryEquivalent(EquFileName, BusList, FltOpt):
    FltOpt1 = (c_double*3)(0,0,0)
    for i in range(3):
        FltOpt1[i] = c_double(FltOpt[i])
    #
    BusList1 = (c_int*len(BusList))(0)
    for i in range(len(BusList)):
        BusList1[i]= c_int(BusList[i])
    #
    OlxAPI.BoundaryEquivalent(EquFileName, BusList1, FltOpt1)
#
def FindObj(TC_type,paraCode,vals):
    """ find obj by vals
    Args: TC_type: (TC_LINE,...) equipement type
          paracode: (LN_sName,..): paratype to search
          vals [] : values to search
    returns:
        []: array of object handle

    example:
        FindObj(TC_LINE,LN_sName,["NEV/OHIO", "1565"]) => find all line with name =["NEV/OHIO", "1565"]
    """
    #
    res = [None] * len(vals)
    #
    obA = getEquipmentHandle(TC_type)
    va = getEquipmentData(obA,paraCode)
    for j in range (len(vals)):
        for i in range(len(va)):
            if vals[j]==va[i]:
                res[j] = obA[i]
                break
    return res
#
def getTest():
    """
    get para system
    cod = SY_dBaseMVA...
    """
    argsGetData = {}
    argsGetData["hnd"] = TC_SYS
    argsGetData["token"] = TC_AREA
    if OLXAPI_OK != get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return  argsGetData["data"]
#
def saveAsOlr(fileNew):
    fileNew = os.path.abspath(fileNew)
    if fileNew:
        if OLXAPI_FAILURE == OlxAPI.SaveDataFile(fileNew):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        OlxAPI.CloseDataFile()
        return True
    return False
#
def run1LPFCommand(cmdParams):
    if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(cmdParams):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    return True
#
def addString2File(nameFile,sres):
    print("addString2File: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.add2File(nameFile,sres)
#
def read_File_text (sfile):
    print("read_File_text: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.read_File_text(sfile)
#
def read_File_text_1 (sfile):
    print("read_File_text_1: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.read_File_text_1(sfile)
#
def updateNameFile(path,fi,exti,fo,exto):
    #to update name file output ARGVS.fo,fr,...
    print("@This function is deprecated, it is replaced by AppUtils.py")
    #
    res = fo
    if res== "":
        base = os.path.basename(fi).upper().replace(exti.upper(),"")
        res = os.path.join (path, base + exto.upper())
    #
    if not res.upper().endswith(exto.upper()):
        res += exto.upper()
    #
    return res
#
def read_File_csv(fileName, delim):
    print("read_File_csv: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.read_File_csv(fileName, delim)
#
def deleteFile(sfile):
    print("deleteFile: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.deleteFile(sfile)
#
def compare2FileText(file1,file2):
    print("compare2FileText: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.compare2FilesText(file1,file2)

def convert_LengthUnit(len_unit1,len_unit):
    print("convert_LengthUnit: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.convert_LengthUnit(len_unit1,len_unit)
#
def gui_Error(sTitle,sMain):
    print("gui_Error: This function is deprecated, it is replaced by AppUtils.py")
    AppUtils.gui_error(sTitle,sMain)

#-----------
def getStrNow():
    print("getStrNow: This function is deprecated, it is replaced by AppUtils.py")
    return AppUtils.getStrNow()
#
def findBusNumber(bn,bkv):
    bhnd = OlxAPI.FindBus(bn,bkv)
    if bhnd>0:
        return get1EquipmentData(bhnd,[BUS_nNumber])[0]
    return -1
