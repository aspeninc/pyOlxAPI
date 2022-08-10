﻿"""Sample usage of ASPEN OlxAPI in Python.
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__version__ = "1.3.0"
__email__ = "support@aspeninc.com"
__status__ = "Release"

import sys
import os
from ctypes import *
import OlxAPI
from OlxAPIConst import *
import OlxAPILib
olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
#
# Command Line INPUTS
import argparse
INPUTS = argparse.ArgumentParser(epilog= "")
INPUTS = argparse.ArgumentParser("\tA collection of code snippets to demo usage of OlxAPI calls.")
INPUTS.add_argument('-olxpath', metavar='', help = ' (str) Full pathname of the folder, where the ASPEN OlxApi.dll is located', default = olxpath)
INPUTS.usage = '\n\tsample.py -fi "InputFile.olr"'
INPUTS.add_argument('-fi' , metavar='', help = '*(str) OLR input file',   default = "")
INPUTS.add_argument('-fib' , metavar='', help = '(str) OLR input file B',   default = "")
INPUTS.add_argument('-fo' , metavar='', help = '(str) output file',   default = "")
ARGVS = INPUTS.parse_known_args()[0]

def testGetObjGraphicData():
    """Test API GetObjGraphic
    """
    buf = (c_int*500)(0)
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        if OLXAPI_FAILURE == OlxAPI.GetObjGraphicData(hnd,buf):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        sObj1LPF = OlxAPI.decode(OlxAPI.PrintObj1LPF(hnd))
        print( sObj1LPF + ': x=' + str(buf[0]) + ' y=' + str(buf[1]) )

def testEliminateZZBranch():
    """Test API EliminateZZBranch
    """
    buf = (c_int*10)(0)
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_SWITCH
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = SW_nInService
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if 0 != argsGetData["data"]:
            sObj1LPF = OlxAPI.PrintObj1LPF(hnd)
            print(sObj1LPF)
            OlxAPI.OlxAPIEliminateZZBranch(hnd, 1+2+4, buf )
            print( "Retained buses" )
            for i in range(len(buf)):
                if buf[i]==-1:
                    break
                print(buf[i])
                print(OlxAPI.PrintObj1LPF(buf[i]))
            print( "Eliminated buses" )
            for ii in range(i+1,len(buf)):
                if buf[ii]==-1:
                    break
                print(buf[ii])
                print(OlxAPI.PrintObj1LPF(buf[ii]))

def testFindObj1LPF():
    """Test API FindObj1LPF
    """
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        sObj1LPF = OlxAPI.decode(OlxAPI.PrintObj1LPF(hnd))
        foundHnd = c_int(0)
        OlxAPI.FindObj1LPF(sObj1LPF, byref(foundHnd) )
        if hnd != foundHnd.value:
            print( "Not found: " + sObj1LPF + ". " + OlxAPI.ErrorString() )
        else:
            print( "Found: " + sObj1LPF )

def testFindObjGUID():
    """Test API FindObj1LPF with object GUID
    """
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        sObj1LPF = OlxAPI.decode(OlxAPI.PrintObj1LPF(hnd))
        sObjGUID = OlxAPI.decode(OlxAPI.GetObjGUID(hnd))
        foundHnd = c_int(0)
        OlxAPI.FindObj1LPF(sObjGUID, byref(foundHnd) )
        if hnd != foundHnd.value:
            print( "Not found: " + sObj1LPF + ". " + OlxAPI.ErrorString() )
        else:
            print( "Found GUID: " + sObjGUID + " " + sObj1LPF )

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
    print(checkParams)
    if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(checkParams):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print("Success")

def testDIFFANDMERGE():
    """Test API Run1LPFCommand DIFFANDMERGE
    """
    if ARGVS.fib == "":
        print( "Input fileB is missing" )
        return 1
    if ARGVS.fo == "":
        print( "Output file name is missing" )
        return 1
    inputParams = '<DIFFANDMERGE FILEPATHB="' + ARGVS.fib + '"' \
                  ' FILEPATHDIFF="' + ARGVS.fo + '" />'
    print(inputParams)
    if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(inputParams):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    fADX = open(ARGVS.fo,"r")
    Lines = fADX.readlines()
    count = 0
    for line in Lines:
        pos = line.find('CHANGECOUNT')
        if pos > -1:
            pos1 = line.find('"', pos+1)
            pos2 = line.find('"', pos1+1)
            changeCount = line[pos1+1:pos2]
            print("Change Count = "+changeCount)
            break
        if count > 10:
            print("Something is wrong in ADX file")
            break
        count += 1
    fADX.close()

def testCHECKRELAYOPERATIONSEA():
    """Test API Run1LPFCommand CHECKRELAYOPERATIONSEA
    """
    checkParams = '<CHECKRELAYOPERATIONSEA REPORTPATHNAME="c:\\000tmp\\" FAULTTYPE="3LG 1LF" DEVICETYPE="DSG DSP" />'
    print(checkParams)
    if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(checkParams):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print("Success")

def testExportNetwork():
    """Test API Run1LPFCommand EXPORTNETWORK
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        cmdParams = '<EXPORTNETWORK' \
                    ' FORMAT="DXT" DXTPATHNAME="' + olrFilePath + '.dxt"' \
                    ' />'

        print(cmdParams)
        if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(cmdParams):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("Success")
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def testExportRelay():
    """Test API Run1LPFCommand EXPORTRELAY
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        cmdParams = '<EXPORTRELAY' \
                    ' FORMAT="RAT" RATPATHNAME="' + olrFilePath + '.rat"' \
                    ' SCOPE= "3" SELECTEDOBJ="\'CLAYTOR\' 132"' \
                    ' DEVICETYPE= "OC,DS"' \
                    ' />'
        print(cmdParams)
        if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(cmdParams):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("Success")
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def testReadChangeFile(changeFile):
    # Test OlxAPI.ReadChangeFile

    argsGetData = {}
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNODSRly
    if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    noDS = argsGetData["data"]
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNOOCRly
    if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    noOC = argsGetData["data"]
    print("Before change file DS= ", noDS, " OC= ", noOC)

    if OLXAPI_OK != OlxAPI.ReadChangeFile(changeFile):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())

    print(OlxAPI.ErrorString())
    ttyLogFile = open(os.getenv('TEMP') + '\\PowerScriptTTYLog.txt', 'r')
    for line in ttyLogFile:
        print(line)
    ttyLogFile.close()

    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNODSRly
    if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    noDS = argsGetData["data"]
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNOOCRly
    if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    noOC = argsGetData["data"]
    print("After change file DS= ", noDS, " OC= ", noOC)

def testGetDataMupair():
    try:
        #olrFilePath = ARGVS.fi
        #if not os.path.isfile(olrFilePath):
        #    print("OLR file does not exit: ", olrFilePath)
        #    return 1

        #if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
        #    raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #print("File opened successfully: " + olrFilePath)
        # Test OlxAPI.FindBus()
        bsName = "CLAYTOR"
        bsKV   = 132.0
        hnd    = OlxAPI.FindBus( bsName, bsKV )
        if hnd == OLXAPI_FAILURE:
            print("Bus ", bsName, bsKV, " not found")
            return 1        # error
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_nNumber
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        bsNo = argsGetData["data"]
        print(("hnd= " + str(hnd) + ", Bus " + str(bsNo) + " " +  bsName + str(bsKV)))
        branchHnd = c_int(0)
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( hnd, c_int(TC_BRANCH), byref(branchHnd) ) ) :
            print(( OlxAPI.FullBranchName(branchHnd) ))
            argsGetData = {}
            argsGetData["hnd"] = branchHnd.value
            argsGetData["token"] = BR_nType
            if (OLXAPI_OK != OlxAPILib.get_data(argsGetData)):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            typeCode = argsGetData["data"]
            if typeCode == TC_LINE:
                argsGetData["hnd"] = branchHnd.value
                argsGetData["token"] = BR_nHandle
                if (OLXAPI_OK != OlxAPILib.get_data(argsGetData)):
                    raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                hndLine = argsGetData["data"]
                argsGetData["hnd"] = hndLine
                argsGetData["token"] = LN_nMuPairHnd
                argsGetData["data"] = 0
                while OLXAPI_OK == OlxAPILib.get_data(argsGetData):
                    hndPair = argsGetData["data"]
                    argsGetDataX = {}
                    argsGetDataX["hnd"] = hndPair
                    argsGetDataX["token"] = MU_nHndLine1
                    if (OLXAPI_OK != OlxAPILib.get_data(argsGetDataX)):
                        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                    hndLine1 = argsGetDataX["data"]
                    argsGetDataX["token"] = MU_nHndLine2
                    if (OLXAPI_OK != OlxAPILib.get_data(argsGetDataX)):
                        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                    hndLine2 = argsGetDataX["data"]
                    print(( "MU pair " + str(hndPair) + ":" ))
                    print(( "  " + OlxAPI.FullBranchName(hndLine1) ))
                    print(( "  " + OlxAPI.FullBranchName(hndLine2) ))
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def imMag(r,i):
    return math.sqrt(r*r+i*i)
def imAng(r,i):
    return math.atan2(i,r)*180/math.pi

def testOlxAPIComputeRelayTime():

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
    print(( "V=", Vmag[0], Vang[0], Vmag[1], Vang[1], Vmag[2], Vang[2] ))

    Imag = (c_double*5)(0.0)
    Imag[0] = imMag(Ireal[0],Iimag[0])
    Imag[1] = imMag(Ireal[1],Iimag[1])
    Imag[2] = imMag(Ireal[2],Iimag[2])
    Iang = (c_double*5)(0.0)
    Iang[0] = imAng(Ireal[0],Iimag[0])
    Iang[1] = imAng(Ireal[1],Iimag[1])
    Iang[2] = imAng(Ireal[2],Iimag[2])
    print(( "I=",Imag[0], Iang[0], Imag[1], Iang[1], Imag[2], Iang[2] ))
    opTime = c_double(0.0)
    opDevice = create_string_buffer(b'\000' * 128)

    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYOCG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = OG_sID
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if "CL-G1" == argsGetData["data"]:
            if OLXAPI_OK != OlxAPI.ComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            print(( "OC Ground relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value ))
            break
    argsGetEquipment["tc"] = TC_RLYOCP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        handle = argsGetEquipment["hnd"]
        argsGetData = {}
        argsGetData["hnd"] = handle
        argsGetData["token"] = OP_sID
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if "CL-P1" == argsGetData["data"]:
            if OLXAPI_OK != OlxAPI.ComputeRelayTime(handle,Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            print(( "OC Phase relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value ))
            break
    argsGetEquipment["tc"] = TC_RLYDSG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = DG_sID
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if "Clator_NV G1" == argsGetData["data"]:
            if OLXAPI_OK != OlxAPI.ComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            print(( "DS Ground relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value ))
            break
    argsGetEquipment["tc"] = TC_RLYDSP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = DP_sID
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if "CLPhase2" == argsGetData["data"]:
            if OLXAPI_OK != OlxAPI.ComputeRelayTime(argsGetEquipment["hnd"],Imag,Iang,Vmag,Vang,VpreMag,VpreAng,pointer(opTime),pointer(opDevice)):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            print(( "DS Phase relay: " + str(argsGetData["data"]) + " opTime=" + str(opTime.value) + " opDevice=" + opDevice.value ))
            break

def testOlxAPIMakeOutageList():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYGROUP
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        handle = argsGetEquipment["hnd"]
        aTags = OlxAPI.GetObjTags(handle)
        if aTags == "testOutageList;":
            print(( "Outage list for relay group: " + OlxAPI.FullBranchName(handle) ))
            maxTiers = c_int(0)
            wantedTypes = c_int(1+2+4+8)
            listLen = c_int(0)
            OlxAPI.MakeOutageList(handle, maxTiers, wantedTypes, None, pointer(listLen) )
            print(( "listLen=" + str(listLen.value) ))
            branchList = (c_int*(5+listLen.value))(0)
            OlxAPI.MakeOutageList(handle, maxTiers, wantedTypes, branchList, pointer(listLen) )
            for i in range(listLen.value):
                print("Branch=" + OlxAPI.FullBranchName(branchList[i]))
            return


def testOlxAPIGetSetObjTagsMemo():
    # Test GetObjTags and GetObjMemo
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_LINE
    argsGetEquipment["hnd"] = 0  # Get all lines
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        lineHnd = argsGetEquipment["hnd"]
        print(OlxAPI.FullBranchName(lineHnd))
        aLine1 = OlxAPI.GetObjTags(lineHnd)
        aLine2 = OlxAPI.GetObjMemo(lineHnd)
        if (aLine1 != "") or (aLine2 != ""):
            print(( "Line: " + OlxAPI.FullBranchName(lineHnd) ))
        if aLine1 != "":
            print(( "  Existing tags=" + aLine1 ))
        if OLXAPI_OK != OlxAPI.SetObjTags(lineHnd, c_char_p("NewTag;" + aLine1 ) ):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        aLine1 = OlxAPI.GetObjTags(lineHnd)
        print(( "  New tags=" + aLine1 ))
        if aLine2 != "":
            print(( "  Existing memo=" + aLine2 ))
        if OLXAPI_OK != OlxAPI.SetObjMemo(lineHnd, c_char_p("New memo: line 1\r\nLine2\r\n" + aLine2 ) ):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        aLine2 = OlxAPI.GetObjMemo(lineHnd)
        print(( "  New memo=" + aLine2 ))
        return 0

def testFaultSimulation():
    # Test Fault simulation
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        busHnd = argsGetEquipment["hnd"]
        sObj = str(OlxAPI.PrintObj1LPF(busHnd))
        #print(sObj)
        if sObj.find("NEVADA") > -1:
            print("\n>>>>>>Bus fault at: " + sObj)
            OlxAPILib.run_busFault(busHnd)
            print ("\n>>>>>>Test bus fault SEA")
            OlxAPILib.run_steppedEvent(busHnd)
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
    return 0

def testGetSEADeviceOp():
    sFindObj = "[BUS] 'NEVADA' 132 kV"
    busHnd = c_int(0)
    if OLXAPI_OK != OlxAPI.FindObj1LPF(sFindObj,byref(busHnd)):
        print( sFindObj + " Not found" )
        return 1
    sFindObj = "[DSRLYP]  CLPhase2@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"
    rlyHnd = c_int(0)
    if OLXAPI_OK != OlxAPI.FindObj1LPF(sFindObj,byref(rlyHnd)):
        print( sFindObj + " Not found" )
        return 1
    #print(sObj)
    print ("\n>>>>>>Run bus fault SEA at " + str(OlxAPI.PrintObj1LPF(busHnd) ) )
    print (">>>>>>Check relay operation: " + str(OlxAPI.PrintObj1LPF(rlyHnd) ) )

    OlxAPILib.run_steppedEvent(busHnd.value)
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
        print (OlxAPI.decode(cast(szFaultDest, c_char_p).value))
        print (OlxAPI.decode(cast(szEventDesc, c_char_p).value))
        # Pick corresponding fault
        if OLXAPI_OK == OlxAPI.PickFault(ii,3):
            triptime = c_double(9999)
            sDeviceOp = create_string_buffer(b'\000' * 32)
            if OLXAPI_OK == OlxAPI.GetRelayTime(rlyHnd, 1, 1, byref(triptime), sDeviceOp ):
                print( OlxAPI.decode(sDeviceOp.value) + "=" + str(triptime.value) )
    return 0

def testBoundaryEquivalent(OlrFileName):
    # Test boundary equivalent network creation
    EquFileName = OlrFileName.lower().replace( ".olr", "_eq.olr" )
    FltOpt = (c_double*3)(99,0,0)
    BusList = (c_int*3)(0)
    bsName = "CLAYTOR"
    bsKV   = 132.0
    hnd    = OlxAPI.FindBus( bsName, bsKV )
    if hnd == OLXAPI_FAILURE:
        raise OlxAPI.OlxAPIException("Bus ", bsName, bsKV, " not found")
    BusList[0] = c_int(hnd)
    bsName = "NEVADA"
    bsKV   = 132.0
    hnd    = OlxAPI.FindBus( bsName, bsKV )
    if hnd == OLXAPI_FAILURE:
        raise OlxAPI.OlxAPIException("Bus ", bsName, bsKV, " not found")
    BusList[1] = c_int(hnd)
    BusList[2] = c_int(-1)
    if OLXAPI_OK != OlxAPI.BoundaryEquivalent(c_char_p(EquFileName), BusList, FltOpt):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print("Success. Equivalent is  in " + EquFileName)
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
    if OLXAPI_OK != OlxAPI.DoBreakerRating(Scope, RatingThreshold, OutputOpt, OptionalReport,
                            ReportTXT, ReportCSV, ConfigFile):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print("Success. Report is  in " + ReportTXT.value)
    return 0

def testGetData_BUS():
    # Test GetData
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        busHnd = argsGetEquipment["hnd"]
        argsGetData = {}
        argsGetData["hnd"] = busHnd
        argsGetData["token"] = BUS_sName
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busName = argsGetData["data"]
        argsGetData["token"] = BUS_dKVnominal
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busKV = argsGetData["data"]
        print(( busName, busKV ))
    return 0

def testGetData_DSRLY():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYDSP
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = DP_vParamLabels
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print(val)
    return 0

def testGetData_GENUNIT():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_GENUNIT
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hndUnit = argsGetEquipment["hnd"]
        print( "hndUnit =" + OlxAPI.PrintObj1LPF(hndUnit) )

        argsGetData = {}
        argsGetData["hnd"] = hndUnit
        argsGetData["token"] = GU_nGenHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndGen = argsGetData["data"]
        print( "hndGen =" + OlxAPI.PrintObj1LPF(hndGen) )

        argsGetData = {}
        argsGetData["hnd"] = hndGen
        argsGetData["token"] = GE_nBusHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndBus = argsGetData["data"]
        print( "hndBus =" + OlxAPI.PrintObj1LPF(hndBus) )

        argsGetData = {}
        argsGetData["hnd"] = hndUnit
        argsGetData["token"] = GU_vdX
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print('X= ' + str(val))
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = GU_nOnline
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print('nOnline= ' + str(val))
        print(val)
    return 0

def testGetData_SHUNTUNIT():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_SHUNTUNIT
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hndUnit = argsGetEquipment["hnd"]
        print( "hndUnit =" + OlxAPI.PrintObj1LPF(hndUnit) )

        argsGetData = {}
        argsGetData["hnd"] = hndUnit
        argsGetData["token"] = SU_nShuntHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndGen = argsGetData["data"]
        print( "hndShunt =" + OlxAPI.PrintObj1LPF(hndGen) )

        argsGetData = {}
        argsGetData["hnd"] = hndGen
        argsGetData["token"] = SH_nBusHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndBus = argsGetData["data"]
        print( "hndBus =" + OlxAPI.PrintObj1LPF(hndBus) )

        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = SU_nOnline
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print('nOnline= ' + str(val))
    return 0

def testGetData_LOADUNIT():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_LOADUNIT
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hndUnit = argsGetEquipment["hnd"]
        print( "hndUnit =" + OlxAPI.PrintObj1LPF(hndUnit) )

        argsGetData = {}
        argsGetData["hnd"] = hndUnit
        argsGetData["token"] = LU_nLoadHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndGen = argsGetData["data"]
        print( "hndLoad =" + OlxAPI.PrintObj1LPF(hndGen) )

        argsGetData = {}
        argsGetData["hnd"] = hndGen
        argsGetData["token"] = LD_nBusHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        hndBus = argsGetData["data"]
        print( "hndBus =" + OlxAPI.PrintObj1LPF(hndBus) )

        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = LU_nOnline
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print('nOnline= ' + str(val))
    return 0

def testGetData_BREAKER():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BREAKER
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = BK_vnG1DevHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        val = argsGetData["data"]
        print(val)
    return 0

def testGetData_SCHEME():
    # Using getequipment
    argsGetEquipment = {}
    argsGetData = {}
    argsGetEquipment["tc"] = TC_SCHEME
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData["hnd"] = argsGetEquipment["hnd"]
        argsGetData["token"] = LS_nRlyGrpHnd
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        rlyGrpHnd = argsGetData["data"]
        argsGetData["token"] = LS_sID
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        sID = argsGetData["data"]
        argsGetData["token"] = LS_sEquation
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        sEqu = argsGetData["data"]
        argsGetData["token"] = LS_sVariables
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        sVariables = argsGetData["data"]
        print('Scheme: ' + sID + '@' + OlxAPI.FullBranchName(rlyGrpHnd) + "\n" + \
            sEqu + "\n" + sVariables)

    # Through relay groups
    argsGetEquipment["tc"] = TC_RLYGROUP
    argsGetEquipment["hnd"] = 0
    argsGetLogicScheme = {}
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        schemeHnd = c_int(0)
        while (OLXAPI_OK == OlxAPI.GetLogicScheme(argsGetEquipment["hnd"], byref(schemeHnd) )):
            argsGetData["hnd"] = schemeHnd.value
            argsGetData["token"] = LS_nRlyGrpHnd
            if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            rlyGrpHnd = argsGetData["data"]
            argsGetData["token"] = LS_sID
            if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            sID = argsGetData["data"]
            argsGetData["token"] = LS_sEquation
            if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            sEqu = argsGetData["data"]
            argsGetData["token"] = LS_sVariables
            if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            sVariables = argsGetData["data"]
            print('Scheme: ' + sID + '@' + OlxAPI.FullBranchName(rlyGrpHnd) + "\n" + \
                sEqu + "\n" + sVariables)
    return 0

def testSaveDataFile(olrFilePath):
    olrFilePath = olrFilePath.lower()
    testReadChangeFile(str(olrFilePath).replace( ".olr", ".chf"))
    olrFilePathNew = str(olrFilePath).replace( ".olr", "x.olr" )
    olrFilePathNew = olrFilePathNew.replace( ".OLR", "x.olr" )
    if OLXAPI_FAILURE == OlxAPI.SaveDataFile(olrFilePathNew):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print(olrFilePathNew + " had been saved successfully")
    return 0

def testFindObj():
    # Test OlxAPI.FindBus()
    bsName = "CLAYTOR"
    bsKV   = 132.0
    hnd    = OlxAPI.FindBus( bsName, bsKV )
    if hnd == OLXAPI_FAILURE:
        print("Bus ", bsName, bsKV, " not found")
    else:
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_nNumber
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        bsNo = argsGetData["data"]
        print("hnd= ", hnd, "Bus ", bsNo, " ", bsName, bsKV)
        print(OlxAPI.PrintObj1LPF(hnd))

    # Test OlxAPI.FindBusNo()
    bsNo = 99
    hnd    = OlxAPI.FindBusNo( bsNo )
    if hnd == OLXAPI_FAILURE:
        print("Bus ", bsNo, " not found")
    else:
        argsGetData = {}
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_sName
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        bsName = argsGetData["data"]
        argsGetData["hnd"] = hnd
        argsGetData["token"] = BUS_dKVnominal
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        bsKV = argsGetData["data"]
        print("hnd= ", hnd, " Bus ", bsNo, " ", bsName, " ", bsKV)

    # Test OlxAPI.FindEquipmentByTag
    tags = "tagS"
    equType = c_int(0)
    equHnd = (c_int*1)(0)
    count = 0
    while OLXAPI_OK == OlxAPI.FindEquipmentByTag( tags, equType, equHnd ):
        print(OlxAPI.PrintObj1LPF(equHnd[0]))
        count = count + 1
    print("Objects with tag " + tags + ": " + str(count))
    return 0

def testDeleteEquipment(olrFilePath):
    hnd = (c_int*1)(0)
    ii = 5
    while (OLXAPI_OK == OlxAPI.GetEquipment(TC_BUS,hnd)):
        busHnd = hnd[0]
        print("Delete " + OlxAPI.PrintObj1LPF(busHnd))
        OlxAPI.DeleteEquipment(busHnd)
        if ii == 0:
            break
        ii = ii - 1

    olrFilePathNew = olrFilePath.lower()
    olrFilePathNew = olrFilePathNew.replace( ".olr", "x.olr" )
    if OLXAPI_FAILURE == OlxAPI.SaveDataFile(olrFilePathNew):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print(olrFilePathNew + " had been saved successfully")

def testGetData_SetData():

    argsGetEquipment = {}
    argsGetData = {}

    # Test GetData with special handles
    argsGetData["hnd"] = HND_SYS
    argsGetData["token"] = SY_nNObus
    if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    print("Number of buses: ", argsGetData["data"])

    # Test SetData and GetData
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0
    ii = 100
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        busHnd = argsGetEquipment["hnd"]
        argsGetData = {}
        argsGetData["hnd"] = busHnd
        argsGetData["token"] = BUS_dKVnominal
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busKV = argsGetData["data"]
        argsGetData["token"] = BUS_sName
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busNameOld = argsGetData["data"]
        argsSetData = {}
        argsSetData["hnd"] = busHnd
        argsSetData["token"] = BUS_sName
        argsSetData["data"] = busNameOld+str(ii+1)
        if OLXAPI_OK != OlxAPILib.set_data(argsSetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        argsGetData["token"] = BUS_nNumber
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busNumberOld = argsGetData["data"]
        argsSetData["token"] = BUS_nNumber
        argsSetData["data"] = ii+1
        if OLXAPI_OK != OlxAPILib.set_data(argsSetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if OLXAPI_OK != OlxAPI.PostData(busHnd):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        argsGetData["token"] = BUS_sName
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busNameNew = argsGetData["data"]
        argsGetData["token"] = BUS_nNumber
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        busNumberNew = argsGetData["data"]
        print("Old:", busNumberOld, busNameOld, busKV, "kV -> New: ", busNumberNew, busNameNew, busKV, "kV")
        ii = ii + 1

    argsGetEquipment["tc"] = TC_GENUNIT
    argsGetEquipment["hnd"] = 0
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        argsGetData = {}
        hnd = argsGetEquipment["hnd"]
        argsGetData["hnd"] = hnd
        argsGetData["token"] = GU_vdX
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print(OlxAPI.PrintObj1LPF(hnd) + " X=", argsGetData["data"])
        argsSetData = {}
        argsSetData["hnd"] = hnd
        argsSetData["token"] = GU_vdX
        argsSetData["data"] = [0.21,0.22,0.23,0.24,0.25]
        if OLXAPI_OK != OlxAPILib.set_data(argsSetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if OLXAPI_OK != OlxAPI.PostData(hnd):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print(OlxAPI.PrintObj1LPF(hnd) + " Xnew=", argsGetData["data"])
    return 0


def testGetData():
    testGetData_SetData()
    #testGetData_SCHEME()
    #testGetData_BREAKER()
    #testGetData_GENUNIT()
    #testGetData_SHUNTUNIT()
    #testGetData_LOADUNIT()
    #testGetData_DSRLY()
    #testGetData_BUS()
    #testGetRelay()
    #testGetJournalRecord()
    return 0

def testOlxAPI():
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        #testGetObjGraphicData()
        #testEliminateZZBranch()
        #testGetSetUDF()
        #testFindObjGUID()
        #testFindObj1LPF()
        #testDeleteEquipment(olrFilePath)
        testFindObj()
        #testBoundaryEquivalent(olrFilePath)
        #testDoBreakerRating()
        #testGetData()
        #testFaultSimulation()
        #testGetSEADeviceOp()
        #testOlxAPIComputeRelayTime()
        #testOlxAPI.MakeOutageList()
        #testOlxAPI.GetSetObjTagsMemo()
        #testDoRelayCoordination()
        #testCHECKRELAYOPERATIONSEA()
        #testDIFFANDMERGE()
        #testSaveDataFile(olrFilePath)
        return 0

    except  OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def listbranch():
    """
    Print list of branches at the two ends of a line.
    User must select a relay group on the line and fill the input of function branchSearch
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        targetHnd = branchSearch("CLAYTOR", 132, "NEVADA", 132, "1")
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_RLYDSP
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            RGhnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = RGhnd
            argsGetData["token"] = RG_nBranchHnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
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
        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            TypeCode1 = argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nHandle
        else:
            return 1

        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            FacilityHandle = argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nBus1Hnd
        else:
            return 1

        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            Bus1Hnd= argsGetData["data"]
            argsGetData = {}
            argsGetData["hnd"] = Branch1Hnd
            argsGetData["token"] = BR_nBus2Hnd
        else:
            return 1

        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            Bus2Hnd= argsGetData["data"]
        else:
            return 1

        if (TypeCode1 == TC_LINE):
            argsGetData = {}
            argsGetData["hnd"] = FacilityHandle
            argsGetData["token"] = LN_sName
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                LineName = argsGetData["data"]
                BusHnd = Bus1Hnd
                BranchNext = 1
                while (BranchNext != 0):
                    argsGetData = {}
                    argsGetData["hnd"] = Bus2Hnd
                    argsGetData["token"] = BUS_nTapBus
                    if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                         TapCode = argsGetData["data"]
                         if TapCode == 0:
                             break
                         BranchNext = 0
                         BranchHnd  = 0
                         argsGetData = {}
                         argsGetData["hnd"] = Bus2Hnd
                         argsGetData["token"] = TC_BRANCH
                         while (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                             BranchHnd = argsGetData["data"]
                             argsGetData = {}
                             argsGetData["hnd"] = BranchHnd
                             argsGetData["token"] = BR_nBus2Hnd
                             if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                                 RemoteBusHnd = argsGetData["data"]
                                 if RemoteBusHnd == BusHnd:
                                     Branch2Hnd = BranchHnd
                                 else:
                                     argsGetData = {}
                                     argsGetData["hnd"] = BranchHnd
                                     argsGetData["token"] = BR_nType
                                     if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
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
                                             if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                                                 TmpHandle = argsGetData["data"]
                                             else:
                                                 return 1
                                             argsGetData = {}
                                             argsGetData["hnd"] = TmpHandle
                                             argsGetData["token"] = LN_sName
                                             if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                                                 TmpLineName = argsGetData["data"]
                                             else:
                                                 return 1
                                             if (TmpLineName == LineName):
                                                 BranchNext = BranchHnd
                                                 BusHnd     = Bus2Hnd
                                                 Bus2Hnd    = RemoteBusHnd
        # Near bus and branches
        BusHnd = Bus1Hnd
        BusName = OlxAPI.FullBusName( BusHnd )
        BranchThis = Branch1Hnd
        List1Len = 0
        BranchHnd = c_int(0)
        OutageList1 = (c_int*20)(0)
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
            if BranchHnd.value != BranchThis:
                List1Len = List1Len + 1
                OutageList1[List1Len] = BranchHnd
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nBus2Hnd
                if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                    RemoteBusHnd = argsGetData["data"]
                else:
                    return 1
                BusName2 = OlxAPI.FullBusName( RemoteBusHnd )
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nType
                if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
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
        BusName = OlxAPI.FullBusName( BusHnd )
        BranchThis = Branch1Hnd
        List2Len = 0
        BranchHnd = c_int(0);
        OutageList2 = (c_int*20)(0)
        while ( OLXAPI_OK == OlxAPI.GetBusEquipment( BusHnd, c_int(TC_BRANCH), byref(BranchHnd) ) ) :
            if BranchHnd.value != BranchThis:
                List2Len = List2Len + 1
                OutageList2[List2Len] = BranchHnd
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nBus2Hnd
                if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                    RemoteBusHnd = argsGetData["data"]
                else:
                    return 1
                BusName2 = OlxAPI.FullBusName( RemoteBusHnd )
                argsGetData = {}
                argsGetData["hnd"] = BranchHnd.value
                argsGetData["token"] = BR_nType
                if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
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

        print(("Found " + str(List1Len) + " branches at near end"))
        print(("Found " + str(List2Len) + " branches at far end"))

        return 0
    except  OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def orphanbus():
    """
    Print list of buses that have no connection to any other buses
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0
        NObus = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            NObus = NObus + 1
        BusList = (c_int*(NObus))(0)
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_BUS
        argsGetEquipment["hnd"] = 0
        NObus = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            BusList[NObus] = argsGetEquipment["hnd"]
            NObus = NObus + 1

        MapList = (c_int*(NObus))(0)
        for ii in range(0,NObus):
            MapList[ii] = 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_LINE
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = LN_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = LN_nBus2Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_SWITCH
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = SW_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = SW_nBus2Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_PS
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = PS_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = PS_nBus2Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_XFMR
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = XR_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = XR_nBus2Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1

        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_XFMR3
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus2Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1
            argsGetData = {}
            argsGetData["hnd"] = argsGetEquipment["hnd"]
            argsGetData["token"] = X3_nBus3Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                BusHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            nIdx = binarySearch(BusList, BusHnd, 0, NObus-1)
            if nIdx > -1:
                MapList[nIdx] = 0
            else:
                return 1

        nCount = 0
        for ii in range(0,NObus-1):
            if MapList[ii] == 1:
                busName = OlxAPI.FullBusName(BusList[ii])
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
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def linez_v2():
    """
    Report total line impedance and length.
    Lines with tap buses are handled correctly
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)
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
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
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
            while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
                PickedHnd = argsGetEquipment["hnd"]
                if ProcessedHnd[PickedHnd-hndOffset] == 0:
                    argsGetData = {}
                    argsGetData["hnd"] = argsGetEquipment["hnd"]
                    argsGetData["token"] = LN_nBus1Hnd
                    if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                        Bus1Hnd = argsGetData["data"]
                    else:
                        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                    argsGetData = {}
                    argsGetData["hnd"] = Bus1Hnd
                    argsGetData["token"] = BUS_dKVnominal
                    if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                        dKV = argsGetData["data"]
                    else:
                        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                    if dKV >= dKVMin and dKV <= dKVMax:
                        argsGetData = {}
                        argsGetData["hnd"] = Bus1Hnd
                        argsGetData["token"] = BUS_nTapBus
                        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                            TapCode1 = argsGetData["data"]
                        else:
                            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                        argsGetData = {}
                        argsGetData["hnd"] = argsGetEquipment["hnd"]
                        argsGetData["token"] = LN_nBus2Hnd
                        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                            Bus2Hnd = argsGetData["data"]
                        else:
                            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                        argsGetData = {}
                        argsGetData["hnd"] = Bus2Hnd
                        argsGetData["token"] = BUS_nTapBus
                        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                            TapCode2 = argsGetData["data"]
                        else:
                            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                        if TapCode1 == 0 or TapCode2 == 0:
                            compuOneLiner( argsGetEquipment["hnd"], ProcessedHnd, hndOffset )
                            LineCount = LineCount + 1

            print("line precessed.")
            root.update()
            return 0
        else:
            root.update()
            return 1
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def networkutil():
    """
    Various utility functions for traversing OneLiner network
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        targetHnd = branchSearch("CLAYTOR", 132, "NEVADA", 132, "1")
        argsGetEquipment = {}
        argsGetEquipment["tc"] = TC_RLYDSP
        argsGetEquipment["hnd"] = 0
        while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
            RGhnd = argsGetEquipment["hnd"]
            argsGetData = {}
            argsGetData["hnd"] = RGhnd
            argsGetData["token"] = RG_nBranchHnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
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
        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            nBusHnd = argsGetData["data"]
        else:
            return 1
        sText = "Local: " + OlxAPI.FullBusName(nBusHnd) + "; remote: "
        for ii in range(0,nCount):
            BranchHnd = TerminalList[ii]
            argsGetData = {}
            argsGetData["hnd"] = BranchHnd
            argsGetData["token"] = BR_nBus1Hnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                nBusHnd = argsGetData["data"]
            else:
                return 1
            sText = sText + " " + OlxAPI.FullBusName(nBusHnd)
        print( sText )
        return 0
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def testLN_nMuPairHnd():
    """
    Various utility functions for traversing OneLiner network
    """
    try:
        olrFilePath = ARGVS.fi
        """
        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))
        """
        if not os.path.isfile(olrFilePath):
            print("OLR file does not exit: ", olrFilePath)
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("File opened successfully: " + olrFilePath)

        targetHnd = branchSearch("GLEN LYN", 132, "TEXAS", 132, "1")
        argsGetData = {}
        argsGetData["hnd"] = targetHnd
        argsGetData["token"] = BR_nType
        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            brType = argsGetData["data"]
        else:
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if brType == TC_LINE:
            argsGetData = {}
            argsGetData["hnd"] = targetHnd
            argsGetData["token"] = BR_nHandle
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                itemHnd = argsGetData["data"]
            else:
                raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            argsGetData = {}
            argsGetData["hnd"] = itemHnd
            argsGetData["token"] = LN_nMuPairHnd
            if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
                nMuHnd = argsGetData["data"]
            else:
                return OlxAPIException(OlxAPI.ErrorString())
        return 0
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def testGetJournalRecord():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_RLYOCG
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        print(OlxAPI.PrintObj1LPF(hnd))
        JRec = OlxAPI.GetObjJournalRecord(hnd).split("\n")
        print("Created: " + JRec[0])
        print("Created by: " + JRec[1])
        print("Modified: " + JRec[2])
        print("Modified by: " + JRec[3])
    return 0

def testGetRelay():
    """
    Test OlxAPI.GetRelay()
    """
    try:
        """
        olrFilePath = ARGVS.fi

        if (not os.path.isfile(olrFilePath)):
            Tkinter.Tk().withdraw() # Close the root window
            opts = {}
            opts['filetypes'] = [('ASPEN OneLiner file',('.olr'))]
            opts['title'] = 'Open OneLiner Network'
            olrFilePath = str(tkFileDialog.askopenfilename(**opts))

        if not os.path.isfile(olrFilePath):
            print "OLR file does not exit: ", olrFilePath
            return 1

        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(olrFilePath,1):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print "File opened successfully: " + olrFilePath
        """
        targetHnd = OlxAPILib.branchSearch("NEVADA", 132, "REUSENS", 132, "1")
        if targetHnd == None:
            raise OlxAPI.OlxAPIException("Branch not found: NEVADA-REUSENS 132kV 1L")
        argsGetData = {}
        argsGetData["hnd"] = targetHnd
        argsGetData["token"] = BR_nRlyGrp1Hnd
        if (OLXAPI_OK == OlxAPILib.get_data(argsGetData)):
            nRGHnd = argsGetData["data"]
        else:
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if OlxAPI.PickFault( SF_LAST, 9 ):
            getTime = True
        else:
            getTime = False
        hndRelay = c_int(0)
        while OLXAPI_OK == OlxAPI.GetRelay( nRGHnd, byref(hndRelay) ) :
            print(OlxAPI.PrintObj1LPF( hndRelay ))
            if TC_RLYDSP == OlxAPI.EquipmentType(hndRelay):
                sx = create_string_buffer(b"Z2_Delay",128)
                if OLXAPI_OK != OlxAPI.GetData(hndRelay,DP_sParam,byref(sx)):
                    raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                print((b"Z2_Delay=" + sx.value))
            elif TC_RLYOCG == OlxAPI.EquipmentType(hndRelay):
                argsGetData = {}
                argsGetData["hnd"] = hndRelay.value
                argsGetData["token"] = OG_vdDirSetting
                if OLXAPI_OK != OlxAPILib.get_data(argsGetData):
                    raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                print("OG_vdDirSetting=" + str(argsGetData["data"]))
            if getTime:
                triptime = c_double(0)
                if OLXAPI_OK != OlxAPI.GetRelayTime(hndRelay, 1.0, 1, byref(triptime),byref(sx)):
                    raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
                print(("Trip time=" + str(triptime.value) + " device=" + str(sizeof.value)))
        return 0
    except OlxAPI.OlxAPIException as err:
        print(( "Error: {0}".format(err)))
        return 1        # error

def testGetSetUDF():
    argsGetEquipment = {}
    argsGetEquipment["tc"] = TC_BUS
    argsGetEquipment["hnd"] = 0  # Get the first object
    while (OLXAPI_OK == OlxAPILib.get_equipment(argsGetEquipment)):
        hnd = argsGetEquipment["hnd"]
        sOutLine = OlxAPI.decode(OlxAPI.PrintObj1LPF(hnd)) + ":"
        Fidx = 0
        FName = create_string_buffer(b'\000' * MXUDFNAME)
        FValue = create_string_buffer(b'\000' * MXUDF)
        while OLXAPI_OK == OlxAPI.GetObjUDFByIndex(hnd,Fidx,FName,FValue):
            sOutLine = sOutLine + " [" + OlxAPI.decode(FName.value) + "]=|" + OlxAPI.decode(FValue.value)
            if FValue.value == "":
                FValue.value = b"New" + OlxAPI.decode(FName.value)
            else:
                FValue.value = b""
            if OLXAPI_OK != OlxAPI.SetObjUDF(hnd,FName,FValue):
                raise OlxAPI.OlxAPIException(OlxAPI.GetLastError())
            OlxAPI.GetObjUDF(hnd,FName,FValue)
            sOutLine = sOutLine + "|===>|" + OlxAPI.decode(FValue.value) + "|"
            Fidx = Fidx + 1
        print( sOutLine )
    return 0

def main(argv=None):
    #ARGVS.fi = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\Sample30.OLR'
    #
    # IMPORTANT: Successfull initialization is required before any
    #            other OlxAPI call can be executed.
    #
    if ARGVS.fi == "":
        print( "Input file is missing. Try samples.py -h" )
        return 1
    try:
        OlxAPI.InitOlxAPI(ARGVS.olxpath)
    except OlxAPI.OlxAPIException as err:
        print(( "OlxAPI Init Error: {0}".format(err) ))
        return 1        # error:


    return testOlxAPI()
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

