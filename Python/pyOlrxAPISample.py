"""Demo usage of ASPEN OlrxAPI in Python.
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2018, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "0.1.1"
__email__ = "support@aspeninc.com"
__status__ = "In development"

import sys
#import os.path
import os
from pyOlrxAPI import *
from ctypes import *
from pyOlrxAPILib import *

def testDoRelayCoordination():
    """Test API Run1LPFCommand CHECKPRIBACKCOORD
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

def testExportNetwork():
    """Test API Run1LPFCommand EXPORTNETWORK
    """
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
        print "File opened successfully: " + olrFilePath

        cmdParams = '<EXPORTNETWORK' \
                    ' FORMAT="DXT" DXTPATHNAME="' + olrFilePath + '.dxt"' \
                    ' />'

        print cmdParams
        if OLRXAPI_FAILURE == OlrxAPIRun1LPFCommand(cmdParams):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print "Success"
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def testExportRelay():
    """Test API Run1LPFCommand EXPORTRELAY
    """
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
        print "File opened successfully: " + olrFilePath

        cmdParams = '<EXPORTRELAY' \
                    ' FORMAT="RAT" RATPATHNAME="' + olrFilePath + '.rat"' \
                    ' SCOPE= "3" SELECTEDOBJ="\'CLAYTOR\' 132"' \
                    ' DEVICETYPE= "OC,DS"' \
                    ' />'
        print cmdParams
        if OLRXAPI_FAILURE == OlrxAPIRun1LPFCommand(cmdParams):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print "Success"
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def testReadChangeFile(changeFile):
    # Test OlrxAPIReadChangeFile
    
    argsGetData = {}
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNODSRly
    if OLRXAPI_OK != olrx_get_data(argsGetData):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    noDS = argsGetData["data"]
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNOOCRly
    if OLRXAPI_OK != olrx_get_data(argsGetData):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    noOC = argsGetData["data"]
    print "Before change file DS= ", noDS, " OC= ", noOC

    if OLRXAPI_OK != OlrxAPIReadChangeFile(changeFile):
        raise OlrxAPIException(OlrxAPIErrorString()) 

    print OlrxAPIErrorString()
    ttyLogFile = open(os.getenv('TEMP') + '\\PowerScriptTTYLog.txt', 'r')
    for line in ttyLogFile: 
        print line 
    ttyLogFile.close()

    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNODSRly
    if OLRXAPI_OK != olrx_get_data(argsGetData):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    noDS = argsGetData["data"]
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNOOCRly
    if OLRXAPI_OK != olrx_get_data(argsGetData):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    noOC = argsGetData["data"]
    print "After change file DS= ", noDS, " OC= ", noOC

def testGetDataMupair():
    try:
        if len(sys.argv) == 1:
            print "Usage: " + sys.argv[0] + " YourNetwork.olr"
            return 0
        olrFilePath = sys.argv[1]
        if not os.path.isfile(olrFilePath):
            print "OLR file does not exit: ", olrFilePath
            return 1

        if OLRXAPI_FAILURE == OlrxAPILoadDataFile(olrFilePath,1):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print "File opened successfully: " + olrFilePath
        # Test OlrxAPIFindBus()
        bsName = "CLAYTOR"
        bsKV   = 132.0
        hnd    = OlrxAPIFindBus( bsName, bsKV )
        if hnd == OLRXAPI_FAILURE:
            print "Bus ", bsName, bsKV, " not found"
            return 1        # error
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_nNumber
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        bsNo = argsGetData["data"]
        print("hnd= " + str(hnd) + ", Bus " + str(bsNo) + " " +  bsName + str(bsKV))
        branchHnd = c_int(0)
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( hnd, c_int(TC_BRANCH), byref(branchHnd) ) ) :
            print( OlrxAPIFullBranchName(branchHnd) )
            argsGetData = {}
            argsGetData["hnd"] = branchHnd.value
            argsGetData["token"] = BR_nType
            if (OLRXAPI_OK != olrx_get_data(argsGetData)):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            typeCode = argsGetData["data"]
            if typeCode == TC_LINE:
                argsGetData["hnd"] = branchHnd.value
                argsGetData["token"] = BR_nHandle
                if (OLRXAPI_OK != olrx_get_data(argsGetData)):
                    raise OlrxAPIException(OlrxAPIErrorString()) 
                hndLine = argsGetData["data"]
                argsGetData["hnd"] = hndLine
                argsGetData["token"] = LN_nMuPairHnd
                argsGetData["data"] = 0
                while OLRXAPI_OK == olrx_get_data(argsGetData):
                    hndPair = argsGetData["data"]
                    argsGetDataX = {}
                    argsGetDataX["hnd"] = hndPair
                    argsGetDataX["token"] = MU_nHndLine1
                    if (OLRXAPI_OK != olrx_get_data(argsGetDataX)):
                        raise OlrxAPIException(OlrxAPIErrorString()) 
                    hndLine1 = argsGetDataX["data"]
                    argsGetDataX["token"] = MU_nHndLine2
                    if (OLRXAPI_OK != olrx_get_data(argsGetDataX)):
                        raise OlrxAPIException(OlrxAPIErrorString()) 
                    hndLine2 = argsGetDataX["data"]
                    print ( "MU pair " + str(hndPair) + ":" )
                    print ( "  " + OlrxAPIFullBranchName(hndLine1) )
                    print ( "  " + OlrxAPIFullBranchName(hndLine2) )
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def imMag(r,i):
    return math.sqrt(r*r+i*i)
def imAng(r,i):
    return math.atan2(i,r)*180/math.pi

def testOlrxAPIComputeRelayTime():

    #Phasor #1: 2 CLAYTOR 132.kV - 6 NEVADA 132.kV 1 L
    #Interm. Fault on:   2 CLAYTOR 132.kV - 6 NEVADA 132.kV 1L 2LG 80.00% Type=B-C 
    #Voltage(kV) at 2 CLAYTOR 132.kV 	 56.8197 	 -3.1644 	 10.3955 	 1.37777 	 11.4594 	 1.6586 	
    #               78.6746 	 -0.128026 	 -26.0818 	 -37.6527 	 -18.2145 	 42.7565
    #Current(A) to 6 NEVADA 132.kV 1 L 	 269.072 	 -1148.78 	 -125.967 	 608.215 	 -102.8 	 572.603 	
    #               40.3047 	 32.0372 	 -1695.96 	 500.773 	 1347.25 	 1185
    #APPARENT Z TO FAULT (PRI.OHM) 	 50.7976 	 35.3261 	 8.1157 	 24.5978 	 -25.4112 	 55.23

    Vreal = [78.6746,-26.0818,-18.2145]
    Vimag = [-0.128026,-37.6527,42.7565]
    Ireal = [40.3047,-1695.96,1347.25]
    Iimag = [32.0372, 500.773, 1185]

    #Phasor #2: 2 CLAYTOR 132.kV
    #Bus Fault on:           2 CLAYTOR      132. kV 3LG  R=99999 X=99999 
    #Voltage(kV) at 2 CLAYTOR 132.kV 	 76.347 	 0.39916 	 -1.90034e-008 	 -9.93528e-011 	 0 	 -4.73695e-015 	 76.347 	 0.39916 	 -37.8278 	 -66.318 	 -38.5192 	 65.9188
 	#     0 	 0 	 0 	 0 	 0 	 0 	 0 	 0 	 0 	 0 	 0 	 0

    VpreMag = c_double(76.347)
    VpreAng = c_double(0.39916)

    Vmag = (c_double*3)(0.0)
    Vmag[0] = imMag(Vreal[0],Vimag[0])
    Vmag[1] = imMag(Vreal[1],Vimag[1])
    Vmag[2] = imMag(Vreal[2],Vimag[2])
    Vang = (c_double*3)(0.0)
    Vang[0] = imAng(Vreal[0],Vimag[0]) 
    Vang[1] = imAng(Vreal[1],Vimag[1]) 
    Vang[2] = imAng(Vreal[2],Vimag[2])
    print( "V=", Vmag[0], Vang[0], Vmag[1], Vang[1], Vmag[2], Vang[2] )

    Imag = (c_double*5)(0.0)
    Imag[0] = imMag(Ireal[0],Iimag[0])
    Imag[1] = imMag(Ireal[1],Iimag[1])
    Imag[2] = imMag(Ireal[2],Iimag[2])
    Iang = (c_double*5)(0.0)
    Iang[0] = imAng(Ireal[0],Iimag[0]) 
    Iang[1] = imAng(Ireal[1],Iimag[1]) 
    Iang[2] = imAng(Ireal[2],Iimag[2])
    print( "I=",Imag[0], Iang[0], Imag[1], Iang[1], Imag[2], Iang[2] )
    opTime = c_double(0.0)
    opDevice = create_string_buffer(b'\000' * 128)

    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYOCG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = OG_sID
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if "CL-G1" == argsGetData["data"]:
            if OLRXAPI_OK != OlrxAPIComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print( "OC Ground relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value )
            break
    argsGetEquipment["tc"] = TC_RLYOCP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        handle = argsGetEquipment["hnd"]
        argsGetData = {}
        argsGetData["hnd"] = handle
        argsGetData["token"] = OP_sID
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if "CL-P1" == argsGetData["data"]:
            if OLRXAPI_OK != OlrxAPIComputeRelayTime(handle,Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print( "OC Phase relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value )
            break
    argsGetEquipment["tc"] = TC_RLYDSG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = DG_sID
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if "Clator_NV G1" == argsGetData["data"]:
            if OLRXAPI_OK != OlrxAPIComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print( "DS Ground relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value )
            break
    argsGetEquipment["tc"] = TC_RLYDSP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = DP_sID
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        if "CLPhase2" == argsGetData["data"]:
            if OLRXAPI_OK != OlrxAPIComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            print( "DS Phase relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value )
            break

def testOlrxAPIMakeOutageList():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYGROUP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        handle = argsGetEquipment["hnd"]
        aTags = OlrxAPIGetObjTags(handle)
        if aTags == "testOutageList;":
            print ( "Outage list for relay group: " + OlrxAPIFullBranchName(handle) )
            maxTiers = c_int(0)
            wantedTypes = c_int(1+2+4+8)
            listLen = c_int(0)
            OlrxAPIMakeOutageList(handle, maxTiers, wantedTypes, None, pointer(listLen) )
            print( "listLen=" + str(listLen.value) )
            branchList = (c_int*(5+listLen.value))(0)
            OlrxAPIMakeOutageList(handle, maxTiers, wantedTypes, branchList, pointer(listLen) )
            for i in range(listLen.value):
                print "Branch=" + OlrxAPIFullBranchName(branchList[i])
            return


def testOlrxAPIGetSetObjTagsMemo():
    # Test GetObjTags and GetObjMemo
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_LINE
    argsGetEquipment["hnd"] = 0  # Get all lines
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        lineHnd = argsGetEquipment["hnd"]
        print OlrxAPIFullBranchName(lineHnd)
        aLine1 = OlrxAPIGetObjTags(lineHnd)
        aLine2 = OlrxAPIGetObjMemo(lineHnd)
        if (aLine1 != "") or (aLine2 != ""):
            print ( "Line: " + OlrxAPIFullBranchName(lineHnd) )
        if aLine1 != "":
            print ( "  Existing tags=" + aLine1 )
        if OLRXAPI_OK != OlrxAPISetObjTags(lineHnd, c_char_p("NewTag;" + aLine1 ) ):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        aLine1 = OlrxAPIGetObjTags(lineHnd)
        print ( "  New tags=" + aLine1 )
        if aLine2 != "":
            print ( "  Existing memo=" + aLine2 )
        if OLRXAPI_OK != OlrxAPISetObjMemo(lineHnd, c_char_p("New memo: line 1\r\nLine2\r\n" + aLine2 ) ):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        aLine2 = OlrxAPIGetObjMemo(lineHnd)
        print ( "  New memo=" + aLine2 )
        return 0

def testFaultSimulation():
    # Test Fault simulation
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        busHnd = argsGetEquipment["hnd"]
        sObj = OlrxAPIPrintObj1LPF(busHnd)
        print sObj
        if "NEVADA" in sObj:
            print "\n>>>>>>Bus fault at: " + sObj
            run_busFault(busHnd)
            #print "\n>>>>>>Test bus fault SEA"
            #run_steppedEvent(busHnd)
    return 0

def testBoundaryEquivalent(OlrFileName):
    # Test boundary equivalent network creation
    EquFileName = OlrFileName.lower().replace( ".olr", "_eq.olr" )
    FltOpt = (c_double*3)(99,0,0)
    BusList = (c_int*3)(0)
    bsName = "CLAYTOR"
    bsKV   = 132.0
    hnd    = OlrxAPIFindBus( bsName, bsKV )
    if hnd == OLRXAPI_FAILURE:
        raise OlrxAPIException("Bus ", bsName, bsKV, " not found")
    BusList[0] = c_int(hnd)
    bsName = "NEVADA"
    bsKV   = 132.0
    hnd    = OlrxAPIFindBus( bsName, bsKV )
    if hnd == OLRXAPI_FAILURE:
        raise OlrxAPIException("Bus ", bsName, bsKV, " not found")
    BusList[1] = c_int(hnd)
    BusList[2] = c_int(-1)
    if OLRXAPI_OK != OlrxAPIBoundaryEquivalent(c_char_p(EquFileName), BusList, FltOpt):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print "Success. Equivalent is  in " + EquFileName
    return 0

def testDoBreakerRating():
    # Test Breaker rating simulation
    Scope = (c_int*3)(0,1,1)
    RatingThreshold = c_double(70);
    OutputOpt = c_double(1)
    OptionalReport = c_int(1+2+4)
    ReportTXT = c_char_p("bkrratingreport.txt")
    ReportCSV = c_char_p("")
    ConfigFile = c_char_p("")
    if OLRXAPI_OK != OlrxAPIDoBreakerRating(Scope, RatingThreshold, OutputOpt, OptionalReport, 
                            ReportTXT, ReportCSV, ConfigFile):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print "Success. Report is  in " + ReportTXT.value
    return 0

def testGetData_BUS():
    # Test GetData
    argsGetEquipment = {}
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
        print( busName, busKV )
    return 0

def testGetData_DSRLY():
    argsGetEquipment = {}
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
    return 0

def testGetData_GENUNIT():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_GENUNIT
    argsGetEquipment["hnd"] = 0
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = GU_vdX
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        val = argsGetData["data"]
        print val
    return 0

def testGetData_BREAKER():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BREAKER
    argsGetEquipment["hnd"] = 0
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = BK_vnG1DevHnd
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        val = argsGetData["data"]
        print val
    return 0

def testGetData_SCHEME():
    # Using getequipment
    argsGetEquipment = {}
    argsGetData = {}
    argsGetEquipment["tc"] = TC_SCHEME
    argsGetEquipment["hnd"] = 0
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = LS_nRlyGrpHnd
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        rlyGrpHnd = argsGetData["data"]
        argsGetData["token"] = LS_sID
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        sID = argsGetData["data"]
        argsGetData["token"] = LS_sEquation
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        sEqu = argsGetData["data"]
        argsGetData["token"] = LS_sVariables
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        sVariables = argsGetData["data"]
        print 'Scheme: ' + sID + '@' + OlrxAPIFullBranchName(rlyGrpHnd) + "\n" + \
            sEqu + "\n" + sVariables
    
    # Through relay groups
    argsGetEquipment["tc"] = TC_RLYGROUP
    argsGetEquipment["hnd"] = 0
    argsGetLogicScheme = {}
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        schemeHnd = c_int(0)
        while (OLRXAPI_OK == OlrxAPIGetLogicScheme(argsGetEquipment["hnd"], byref(schemeHnd) )):
            argsGetData["hnd"] = schemeHnd.value
            argsGetData["token"] = LS_nRlyGrpHnd
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            rlyGrpHnd = argsGetData["data"]
            argsGetData["token"] = LS_sID
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            sID = argsGetData["data"]
            argsGetData["token"] = LS_sEquation
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            sEqu = argsGetData["data"]
            argsGetData["token"] = LS_sVariables
            if OLRXAPI_OK != olrx_get_data(argsGetData):
                raise OlrxAPIException(OlrxAPIErrorString()) 
            sVariables = argsGetData["data"]
            print 'Scheme: ' + sID + '@' + OlrxAPIFullBranchName(rlyGrpHnd) + "\n" + \
                sEqu + "\n" + sVariables
    return 0

def testSaveDataFile(olrFilePath):
    olrFilePath = olrFilePath.lower()
    testReadChangeFile(str(olrFilePath).replace( ".olr", ".chf"))
    olrFilePathNew = str(olrFilePath).replace( ".olr", "x.olr" )
    olrFilePathNew = olrFilePathNew.replace( ".OLR", "x.olr" )
    if OLRXAPI_FAILURE == OlrxAPISaveDataFile(olrFilePathNew):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print olrFilePathNew + " had been saved successfully"
    return 0

def testFindObj():
    # Test OlrxAPIFindBus()
    bsName = "CLAYTOR"
    bsKV   = 132.0
    hnd    = OlrxAPIFindBus( bsName, bsKV )
    if hnd == OLRXAPI_FAILURE:
        print "Bus ", bsName, bsKV, " not found"
    else:
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_nNumber
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        bsNo = argsGetData["data"]
        print "hnd= ", hnd, "Bus ", bsNo, " ", bsName, bsKV 
        print OlrxAPIPrintObj1LPF(hnd)

    # Test OlrxAPIFindBusNo()
    bsNo = 99
    hnd    = OlrxAPIFindBusNo( bsNo )
    if hnd == OLRXAPI_FAILURE:
        print "Bus ", bsNo, " not found"
    else:
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_sName
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        bsName = argsGetData["data"]
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_dKVnominal
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        bsKV = argsGetData["data"]
        print "hnd= ", hnd, " Bus ", bsNo, " ", bsName, " ", bsKV

    # Test OlrxAPIFindEquipmentByTag
    tags = c_char_p("tagS")
    equType = c_int(0)
    equHnd = (c_int*1)(0)
    count = 0
    while OLRXAPI_OK == OlrxAPIFindEquipmentByTag( tags, equType, equHnd ):
        print OlrxAPIPrintObj1LPF(equHnd[0])
        count = count + 1
    print "Objects with tag " + tags.value + ": " + str(count)
    return 0

def testDeleteEquipment(olrFilePath):
    hnd = (c_int*1)(0)
    ii = 5
    while (OLRXAPI_OK == OlrxAPIGetEquipment(TC_BUS,hnd)):
        busHnd = hnd[0]
        print "Delete " + OlrxAPIPrintObj1LPF(busHnd)
        OlrxAPIDeleteEquipment(busHnd)
        if ii == 0:
            break
        ii = ii - 1

    olrFilePathNew = olrFilePath.lower()
    olrFilePathNew = olrFilePathNew.replace( ".olr", "x.olr" )
    if OLRXAPI_FAILURE == OlrxAPISaveDataFile(olrFilePathNew):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print olrFilePathNew + " had been saved successfully"

def testGetData_SetData():

    argsGetEquipment = {}
    argsGetData = {}

    # Test GetData with special handles
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNObus
    if OLRXAPI_OK != olrx_get_data(argsGetData):
        raise OlrxAPIException(OlrxAPIErrorString()) 
    print "Number of buses: ", argsGetData["data"]

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
    
    argsGetEquipment["tc"] = TC_GENUNIT
    argsGetEquipment["hnd"] = 0
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        argsGetData = {}
        hnd = argsGetEquipment["hnd"]
        argsGetData["hnd"] = hnd
        argsGetData["token"] = GU_vdX
        if OLRXAPI_OK != olrx_get_data(argsGetData):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print OlrxAPIPrintObj1LPF(hnd) + " X=", argsGetData["data"]
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
        print OlrxAPIPrintObj1LPF(hnd) + " Xnew=", argsGetData["data"]
    return 0


def testGetData():
    #testGetData_SetData()
    #testGetData_SCHEME()
    #testGetData_BREAKER()
    #testGetData_GENUNIT()
    #testGetData_DSRLY()
    #testGetData_BUS()
    #testGetRelay()
    testGetJournalRecord()
    return 0

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
        print "File opened successfully: " + olrFilePath
        
        #testDeleteEquipment(olrFilePath)
        #testFindObj()
        #testBoundaryEquivalent(olrFilePath)
        #testDoBreakerRating()
        testGetData()
        #testFaultSimulation()
        #testOlrxAPIComputeRelayTime()
        #testOlrxAPIMakeOutageList()
        #testOlrxAPIGetSetObjTagsMemo()
        # testDoRelayCoordination()
        #testSaveDataFile(olrFilePath)
        return 0

    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def listbranch():
    """
    Print list of branches at the two ends of a line.
    User must select a relay group on the line and fill the input of function branchSearch
    """
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
        print "File opened successfully: " + olrFilePath

        targetHnd = branchSearch("CLAYTOR", 132, "NEVADA", 132, "1")
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_RLYDSP
        argsGetEquipment["hnd"] = 0    
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):            
            RGhnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = RGhnd
            argsGetData["token"] = RG_nBranchHnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                Branch1Hnd = argsGetData["data"]
                if Branch1Hnd == targetHnd:
                    break
            else:
                return 1
        if Branch1Hnd != targetHnd:
            return 1
        argsGetData = {}
        argsGetData["hnd"] = Branch1Hnd
        argsGetData["token"] = BR_nType
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            TypeCode1 = argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nHandle
        else:
            return 1

        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            FacilityHandle = argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nBus1Hnd
        else:
            return 1

        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            Bus1Hnd= argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nBus2Hnd
        else:
            return 1

        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            Bus2Hnd= argsGetData["data"]
        else:
            return 1

        if (TypeCode1 == TC_LINE):
            argsGetData = {}
            argsGetData["hnd"] = FacilityHandle
            argsGetData["token"] = LN_sName
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                LineName = argsGetData["data"]
                BusHnd = Bus1Hnd
                BranchNext = 1
                while (BranchNext != 0):
                    argsGetData = {}
                    argsGetData["hnd"] = Bus2Hnd
                    argsGetData["token"] = BUS_nTapBus
                    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                         TapCode = argsGetData["data"]
                         if TapCode == 0:
                             break
                         BranchNext = 0
                         BranchHnd  = 0
                         argsGetData = {}
                         argsGetData["hnd"] = Bus2Hnd
                         argsGetData["token"] = TC_BRANCH
                         while (OLRXAPI_OK == olrx_get_data(argsGetData)):
                             BranchHnd = argsGetData["data"]
                             argsGetData = {}
                             argsGetData["hnd"] = BranchHnd
                             argsGetData["token"] = BR_nBus2Hnd
                             if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                                 RemoteBusHnd = argsGetData["data"]
                                 if RemoteBusHnd == BusHnd:
                                     Branch2Hnd = BranchHnd
                                 else:
                                     argsGetData = {}
                                     argsGetData["hnd"] = BranchHnd
                                     argsGetData["token"] = BR_nType
                                     if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                                         TypeCode2 = argsGetData["data"]
                                         if (TypeCode2 == TC_LINE):
                                             if BranchNext == 0:
                                                 BranchNext = BranchHnd
                                                 BusHnd = Bus2Hnd
                                                 Bus2Hnd = RemoteBusHnd
                                         else:
                                             argsGetData = {}
                                             argsGetData["hnd"] = BranchHnd
                                             argsGetData["token"] = BR_nHandle
                                             if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                                                 TmpHandle = argsGetData["data"]
                                             else:
                                                 return 1
                                             argsGetData = {}
                                             argsGetData["hnd"] = TmpHandle
                                             argsGetData["token"] = LN_sName
                                             if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                                                 TmpLineName = argsGetData["data"]
                                             else:
                                                 return 1
                                             if (TmpLineName == LineName):
                                                 BranchNext = BranchHnd
                                                 BusHnd     = Bus2Hnd
                                                 Bus2Hnd    = RemoteBusHnd
        # Near bus and branches
        BusHnd = Bus1Hnd
        BusName = OlrxAPIFullBusName( BusHnd )
        BranchThis = Branch1Hnd
        List1Len = 0
        BranchHnd = c_int(0)
        OutageList1 = (c_int*20)(0)
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
            if BranchHnd.value != BranchThis:
                List1Len = List1Len + 1
                OutageList1[List1Len] = BranchHnd
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nBus2Hnd
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    RemoteBusHnd = argsGetData["data"]
                else:
                    return 1
                BusName2 = OlrxAPIFullBusName( RemoteBusHnd )
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nType
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    TypeCode = argsGetData["data"]
                else:
                    return 1

                if TypeCode == TC_LINE:
                    BranchType = "LINE"
                elif TypeCode == TC_XFMR:
                    BranchType = "XFMR"
                elif TypeCode == TC_XFMR3:
                    BranchType = "XFMR3"
                elif TypeCode == TC_PS:
                    BranchType = "SHIFTER"
                elif TypeCode == TC_SWITCH:
                    BranchType = "SWITCH"
                else:
                    BranchType = "UNKNOWN"

                aLine = "Sub " + BusName + ": " + BranchType + " to " + BusName2 + "; Handle=" + str(BranchHnd)
                print(aLine)

        # Far bus branches
        BusHnd = Bus2Hnd
        BusName = OlrxAPIFullBusName( BusHnd )
        BranchThis = Branch1Hnd
        List2Len = 0
        BranchHnd = c_int(0);
        OutageList2 = (c_int*20)(0)
        while ( OLRXAPI_OK == OlrxAPIGetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
            if BranchHnd.value != BranchThis:
                List2Len = List2Len + 1
                OutageList2[List2Len] = BranchHnd
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nBus2Hnd
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    RemoteBusHnd = argsGetData["data"]
                else:
                    return 1
                BusName2 = OlrxAPIFullBusName( RemoteBusHnd )
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nType
                if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                    TypeCode = argsGetData["data"]
                else:
                    return 1

                if TypeCode == TC_LINE:
                    BranchType = "LINE"
                elif TypeCode == TC_XFMR:
                    BranchType = "XFMR"
                elif TypeCode == TC_XFMR3:
                    BranchType = "XFMR3"
                elif TypeCode == TC_PS:
                    BranchType = "SHIFTER"
                elif TypeCode == TC_SWITCH:
                    BranchType = "SWITCH"
                else:
                    BranchType = "UNKNOWN"

                aLine = "Sub " + BusName + ": " + BranchType + " to " + BusName2 + "; Handle=" + str(BranchHnd)
                print(aLine)

        print("Found " + str(List1Len) + " branches at near end")
        print("Found " + str(List2Len) + " branches at far end")

        return 0
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def orphanbus():
    """
    Print list of buses that have no connection to any other buses
    """
    try:
        if len(sys.argv) == 1:
            print "Usage: " + sys.argv[0] + " YourNetwork.olr"
            return 0
        #olrFilePath = sys.argv[1]
        olrFilePath = 'H:\\data\\00Support&Dev\\02Developtment\\0000V15\\00Dev_V15_OLRX\\sample30_orphanbus.olr'
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
        print "File opened successfully: " + olrFilePath

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0   
        NObus = 0 
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            NObus = NObus + 1
        BusList = (c_int*(NObus))(0)  
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0  
        NObus = 0
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            BusList[NObus] = argsGetEquipment["hnd"]
            NObus = NObus + 1

        MapList = (c_int*(NObus))(0)
        for ii in range(0,NObus):
            MapList[ii] = 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_LINE
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = LN_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = LN_nBus2Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_SWITCH
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = SW_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = SW_nBus2Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_PS
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = PS_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = PS_nBus2Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_XFMR
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = XR_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = XR_nBus2Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_XFMR3
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus2Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus3Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1 

        nCount = 0
        for ii in range(0,NObus-1):
            if MapList[ii] == 1:
                busName = OlrxAPIFullBusName(BusList[ii])
                if nCount == 0:
                    print("Following bus(es) do not have connection to any other bus")
                print(busName)
                nCount = nCount + 1
   
        if nCount == 0:
            print("Found no orphan bus in this network") 
        else:
            aLine = "Found " + str(nCount) + " no orphan bus(s) in this network"
            print(aLine)
        return 0
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error

def linez_v2():
    """
    Report total line impedance and length.
    Lines with tap buses are handled correctly
    """
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
        print "File opened successfully: " + olrFilePath
        maxLines = 10000
        hndOffset = 3
        dKVMin = 0
        dKVMax = 999
        ProcessedHnd = (c_int*(maxLines))(0)
        for ii in range(0,maxLines-1):
            ProcessedHnd[ii] = 1
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_LINE
        argsGetEquipment["hnd"] = 0   
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
            PickedHnd = argsGetEquipment["hnd"]
            Index = PickedHnd - hndOffset
            if Index >= maxLines:
                print("Too many lines in this network. Edit script code to increase maxLines and try again.")
                return 1
            ProcessedHnd[Index] = 0
        root = tk.Tk()
        root.withdraw()
        if tkMessageBox.askyesno('Line Impedance', 'Do you want to print impedance of all lines in kV range: 0-999'):
            LineCount = 0
            argsGetEquipment = {}
            argsGetEquipment["tc"] = TC_LINE
            argsGetEquipment["hnd"] = 0   
            while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)): 
                PickedHnd = argsGetEquipment["hnd"]
                if ProcessedHnd[PickedHnd-hndOffset] == 0:
                    argsGetData = {}
                    argsGetData["hnd"] = argsGetEquipment["hnd"]
                    argsGetData["token"] = LN_nBus1Hnd
                    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                        Bus1Hnd = argsGetData["data"]
                    else:
                        raise OlrxAPIException(OlrxAPIErrorString())
                    argsGetData = {}
                    argsGetData["hnd"] = Bus1Hnd
                    argsGetData["token"] = BUS_dKVnominal
                    if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                        dKV = argsGetData["data"]
                    else:
                        raise OlrxAPIException(OlrxAPIErrorString())
                    if dKV >= dKVMin and dKV <= dKVMax:
                        argsGetData = {}
                        argsGetData["hnd"] = Bus1Hnd
                        argsGetData["token"] = BUS_nTapBus
                        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                            TapCode1 = argsGetData["data"]
                        else:
                            raise OlrxAPIException(OlrxAPIErrorString())
                        argsGetData = {}
                        argsGetData["hnd"] = argsGetEquipment["hnd"]
                        argsGetData["token"] = LN_nBus2Hnd
                        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                            Bus2Hnd = argsGetData["data"]
                        else:
                            raise OlrxAPIException(OlrxAPIErrorString())
                        argsGetData = {}
                        argsGetData["hnd"] = Bus2Hnd
                        argsGetData["token"] = BUS_nTapBus
                        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                            TapCode2 = argsGetData["data"]
                        else:
                            raise OlrxAPIException(OlrxAPIErrorString())
                        if TapCode1 == 0 or TapCode2 == 0:
                            compuOneLiner( argsGetEquipment["hnd"], ProcessedHnd, hndOffset )
                            LineCount = LineCount + 1

            print("line precessed.")
            root.update()
            return 0
        else:
            root.update()
            return 1 
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error  

def networkutil():
    """
    Various utility functions for traversing OneLiner network
    """
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
        print "File opened successfully: " + olrFilePath

        targetHnd = branchSearch("CLAYTOR", 132, "NEVADA", 132, "1")
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_RLYDSP
        argsGetEquipment["hnd"] = 0    
        while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):            
            RGhnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = RGhnd
            argsGetData["token"] = RG_nBranchHnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                BranchHnd = argsGetData["data"]
                if BranchHnd == targetHnd:
                    break
            else:
                return 1
        TerminalList = (c_int*(50))(0)
        nCount = GetRemoteTerminals( BranchHnd, TerminalList )

        argsGetData = {}
        argsGetData["hnd"] = BranchHnd
        argsGetData["token"] = BR_nBus1Hnd
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            nBusHnd = argsGetData["data"]
        else:
            return 1
        sText = "Local: " + OlrxAPIFullBusName(nBusHnd) + "; remote: "
        for ii in range(0,nCount):  
            BranchHnd = TerminalList[ii]
            argsGetData = {}
            argsGetData["hnd"] = BranchHnd
            argsGetData["token"] = BR_nBus1Hnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                nBusHnd = argsGetData["data"]  
            else:
                return 1
            sText = sText + " " + OlrxAPIFullBusName(nBusHnd)  
        print( sText )      
        return 0
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error 
    
def testLN_nMuPairHnd():                    
    """
    Various utility functions for traversing OneLiner network
    """
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
        print "File opened successfully: " + olrFilePath

        targetHnd = branchSearch("GLEN LYN", 132, "TEXAS", 132, "1")
        argsGetData = {}
        argsGetData["hnd"] = targetHnd 
        argsGetData["token"] = BR_nType
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            brType = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if brType == TC_LINE:
            argsGetData = {}
            argsGetData["hnd"] = targetHnd 
            argsGetData["token"] = BR_nHandle
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                itemHnd = argsGetData["data"]
            else:
                raise OlrxAPIException(OlrxAPIErrorString())
            argsGetData = {}
            argsGetData["hnd"] = itemHnd
            argsGetData["token"] = LN_nMuPairHnd
            if (OLRXAPI_OK == olrx_get_data(argsGetData)):
                nMuHnd = argsGetData["data"]
            else:
                return OlrxAPIException(OlrxAPIErrorString())
        return 0
    except  OlrxAPIException as err:
        print( "Error: {0}".format(err))
        return 1        # error 

def testGetJournalRecord():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYOCG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLRXAPI_OK == olrx_get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        print OlrxAPIPrintObj1LPF(hnd)
        JRec = OlrxAPIGetObjJournalRecord(hnd).split("\n")
        print "Created: " + JRec[0]
        print "Created by: " + JRec[1]
        print "Modified: " + JRec[2]
        print "Modified by: " + JRec[3]
    return 0

def testGetRelay():                    
    """
    Test OlrxAPIGetRelay()
    """
    try:
        """
        if len(sys.argv) == 1:
            print "Usage: " + sys.argv[0] + " YourNetwork.olr"
            return 0
        olrFilePath = sys.argv[1]
        
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        
        if not os.path.isfile(olrFilePath):
            print "OLR file does not exit: ", olrFilePath
            return 1

        if OLRXAPI_FAILURE == OlrxAPILoadDataFile(olrFilePath,1):
            raise OlrxAPIException(OlrxAPIErrorString()) 
        print "File opened successfully: " + olrFilePath
        """
        targetHnd = branchSearch("NEVADA", 132, "REUSENS", 132, "1")
        argsGetData = {}
        argsGetData["hnd"] = targetHnd 
        argsGetData["token"] = BR_nRlyGrp1Hnd
        if (OLRXAPI_OK == olrx_get_data(argsGetData)):
            nRGHnd = argsGetData["data"]
        else:
            raise OlrxAPIException(OlrxAPIErrorString())
        if OlrxAPIPickFault( SF_LAST, 9 ):
            getTime = True
        else:
            getTime = False
        hndRelay = c_int(0)
        while OLRXAPI_OK == OlrxAPIGetRelay( nRGHnd, byref(hndRelay) ) :
            print OlrxAPIFullRelayName( hndRelay )
            if TC_RLYDSP == OlrxAPIEquipmentType(hndRelay):
                sx = create_string_buffer("\0",128)
                if OLRXAPI_OK <> OlrxAPIGetData(hndRelay,DP_sParam,byref(sx)):
                    raise OlrxAPIException(OlrxAPIErrorString())
                print("Z2_Delay=" + sx.value)
            if getTime:
                triptime = c_double(0)
                if OLRXAPI_OK <> OlrxAPIGetRelayTime(hndRelay, 1.0, 1, byref(triptime),byref(sx)):
                    raise OlrxAPIException(OlrxAPIErrorString())
                print("Trip time=" + str(triptime.value) + " device=" + str(sizeof.value))
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
        dllPath = os.path.dirname(os.path.abspath(sys.argv[0])) + "\\..\\..\\obj\\debug_1L\\"
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
    #return testGetDataMupair()
    #return listbranch()
    #return orphanbus()
    #return linez_v2()
    #return networkutil()
    #return testExportNetwork()
    #testExportRelay()

if __name__ == '__main__':
    status = main()
    #sys.exit(status)
    
