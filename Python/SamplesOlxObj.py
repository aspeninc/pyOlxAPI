"""
Purpose:
  PyOlxObj application examples
"""
__author__ = "ASPEN Inc."
__copyright__ = "Copyright 2023, Advanced System for Power Engineering Inc."
__license__ = "All rights reserved"
__category__ = "Common"
__email__ = "support@aspeninc.com"
__status__ = "Release"
__version__ = "3.4.5"

import sys
import os
import OlxObj
import argparse
PATH_FILE, _ = os.path.split(os.path.abspath(__file__))

INPUTS = argparse.ArgumentParser('\tA collection of code snippets to demo usage of OlxObj calls.')
INPUTS.add_argument('-fi', metavar='', help='*(str) OLR input file', default='')
INPUTS.add_argument('-olxpath', metavar='', help=' (str) Full pathname of the folder, where the ASPEN OlxApi.dll is located', default='')
ARGVS = INPUTS.parse_known_args()[0]

if PATH_FILE != 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python': # ASPEN dev
    if ARGVS.olxpath == '':
        ARGVS.olxpath = PATH_FILE+'\\1lpfV15'


def sample_1_network(olxpath, fi):
    """ OLR network operations demo

        #
        OlxObj.load_olxapi(olxpath)   # Initialize the OlxAPI module (in olxpath folder)
        OlxObj.OLCase.open(fi, 1)     # Load the OLR network file (fi) as read-only
        OlxObj.OLCase.save()          # Save current OLR data file
        OlxObj.OLCase.save(newFile)   # Save current OLR data file as newFile
        OlxObj.OLCase.close()         # Close the network data file that had been loaded previously

        # Scope for NETWORK access
        OlxObj.OLCase.applyScope()    # Scope for NETWORK access
        OlxObj.OLCase.resetScope()    # reset Scope: All NETWORK Access

        # new File
        OlxObj.OLCase.create_Network(baseMVA)  # Create a new network with baseMVA

        # NETWORK Data:
            'AREA','AREANO','BASEMVA','BREAKER','BUS','CCGEN','COMMENT','DCLINE2','FUSE','GEN',
            'GENUNIT','GENW3','GENW4','KV','LINE','LOAD','LOADUNIT','MULINE','OBJCOUNT','RECLSR','RLYD',
            'RLYDS','RLYDSG','RLYDSP','RLYGROUP','RLYOC','RLYOCG','RLYOCP','RLYV','SCHEME','SERIESRC',
            'SHIFTER','SHUNT','SHUNTUNIT','SVD','SWITCH','XFMR','XFMR3','ZCORRECT','ZONE','ZONENO'

        OlxObj.OLCase.BUS                        # List of all BUSes in the scope
        OlxObj.OLCase.getData('BUS')
        OlxObj.OLCase.getData(['XFMR','XFMR3'])  # List of all 2,3 windings Transformer in the scope
        #
        OlxObj.OLCase.getData('XXX')             # help details in python message error
    """

    print('sample_1_NetWork')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    # Scope for NETWORK access
    OlxObj.OLCase.applyScope(areaNum='1,2,3,5-7', zoneNum=[1, 3], optionTie=0, kV=[132, 132])
    #
    print('\nList of all BUSes in the scope')
    for it in OlxObj.OLCase.GEN:
        print(it.toString())      # text string for object
    #
    for it in OlxObj.OLCase.getData(['XFMR', 'XFMR3','GENERATOR','RLYOC']):
        print(it.toString())
    #
    OlxObj.OLCase.resetScope()
    print('\nList of all RLYDS in NETWORK')
    for it in OlxObj.OLCase.RLYDS:
        print(it.toString())
    #
    OlxObj.OLCase.close()      # Close the network data file that had been loaded previously


def sample_2_Bus(olxpath, fi):
    """ BUS object demo

        # various ways to find a BUS
        b1 = OlxObj.OLCase.findOBJ('BUS',['arizona',132])                          # by (name,kV)
        b1 = OlxObj.OLCase.findOBJ('BUS',28)                                       # by (busNumber)
        b1 = OlxObj.OLCase.findOBJ('BUS','{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}') # by (GUID)
        b1 = OlxObj.OLCase.findOBJ('{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}')
        b1 = OlxObj.OLCase.findOBJ('BUS',"[BUS] 28 'ARIZONA' 132 kV")              # by (STR)
        b1 = OlxObj.OLCase.findOBJ("[BUS] 28 'ARIZONA' 132 kV")
        #
        b1 = OlxObj.OLCase.findBUS(['arizona',132])                                # by (name,kV)
        b1 = OlxObj.OLCase.findBUS(28)                                             # by (busNumber)
        b1 = OlxObj.OLCase.findBUS('{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}')       # by (GUID)
        b1 = OlxObj.OLCase.findBUS("[BUS] 28 'ARIZONA' 132 kV")                    # by (STR)

        # BUS can be accessed by related object
        OlxObj.OLCase.BUS[0]
        OlxObj.OLCase.LINE[0].BUS1

        # BUS data:
            'ANGLEP','AREA','AREANO','BREAKER','BUS','CCGEN','DCLINE2','GEN','GENUNIT','GENW3','GENW4','GUID','JRNL',
            'KEYSTR','KV','KVP','LINE','LOAD','LOADUNIT','LOCATION','MEMO','MIDPOINT','NAME','NO','PARAMSTR',
            'PYTHONSTR','RLYGROUP','SERIESRC','SHIFTER','SHUNT','SHUNTUNIT','SLACK','SPCX','SPCY','SUBGRP','SVD',
            'SWITCH','TAGS','TAP','TERMINAL','VISIBLE','XFMR','XFMR3','ZONE','ZONENO'
        #
        print(b1.NAME,b1.kv,b1.no)       # Bus name,kV,Bus Number
        print(b1.getData('NAME'))        # get Bus Name
        print(b1.getData(['NAME','KV'])) # Bus name,kV
        print(b1.getData())              # (dict) get all Data of Bus

        # methods:
        b1.equals(b2)                    #(bool) comparison of 2 objects
        b1.delete()                      # delete object
        b1.isInList(la)                  # check if object in in List/Set/tuple of object

        b1.toString()                    # (str) text description/composed of object
        b1.copyFrom(b2)                  # copy Data from another object
        b1.changeData(sParam,value)      # change Data (sample_6_changeData())
        b1.postData()                    # Perform validation and update object data in the network database
        b1.getAttributes()               # [str] list attributes of object

        b1.getData('XXX')                # or b1.xx => help details in message error

        # Add BUS:
        params = {'GUID': '{e87fbec5-c4f9-47fe-96d3-8f24ccb101ff}', 'AREANO': 2, 'LOCATION': 'CLAYTOR', 'MEMO': b'm1', 'MIDPOINT': 0,
             'NO': 0, 'SLACK': 1, 'SPCX': 15, 'SPCY': 25, 'SUBGRP': 10, 'TAGS': 'tag1;tag2;', 'TAP': 0, 'ZONENO': 2}
        keys = ['NEW OHIO', 132]
        o1 = OlxObj.OLCase.addOBJ('BUS', keys, params)
    """

    print('sample_2_Bus')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    # various ways to find a BUS
    b1 = OlxObj.OLCase.findOBJ('BUS', ['arizona', 132])   # by (name,kV)
    b1 = OlxObj.OLCase.findOBJ('BUS', 28)                 # by (busNumber)
    b1 = OlxObj.OLCase.findOBJ('BUS', '{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}')  # by (GUID)
    b1 = OlxObj.OLCase.findOBJ('{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}')
    b1 = OlxObj.OLCase.findOBJ('BUS', "[BUS] 28 'ARIZONA' 132 kV")               # by (STR)
    b1 = OlxObj.OLCase.findOBJ("[BUS] 28 'ARIZONA' 132 kV")

    # print string for object, None if not found
    if b1:
        print(b1.toString())
    else:
        print('Bus not found')

    # BUS can be accessed by related object
    b1 = OlxObj.OLCase.BUS[0]                    # by OlxObj.OLCase.BUS[i]
    b1 = OlxObj.OLCase.LINE[0].BUS2              # by OlxObj.OLCase.LINE[i].BUS
    print(b1.toString())

    # Bus Data
    print(b1.NAME, b1.kv, b1.no)        # Bus name,kV,Bus Number
    print(b1.getData('NAME'))           # get Bus Name
    print(b1.getData(['NAME', 'KV']))   # Bus name,kV
    da = b1.getData()                   # get all Data of Bus
    print(da['NAME'], da['KV'])

    # b1.getData('XXX')                  # or b1.xx => help details in message error

    # Add BUS
    params = {'GUID': '{e87fbec5-c4f9-47fe-96d3-8f24ccb101ff}', 'AREANO': 2, 'LOCATION': 'CLAYTOR', 'MEMO': b'm1', 'MIDPOINT': 0,
             'NO': 0, 'SLACK': 1, 'SPCX': 15, 'SPCY': 25, 'SUBGRP': 10, 'TAGS': 'tag1;tag2;', 'TAP': 0, 'ZONENO': 2}
    keys = ['NEW OHIO', 132]
    o1 = OlxObj.OLCase.addOBJ('BUS', keys, params)


def sample_3_Line(olxpath, fi):
    """ AC LINE object demo
        # various ways to find a LINE ----------------------------------------------------------------------
        l1 = OlxObj.OLCase.findOBJ('LINE',[['claytor',132],['nevada',132],'1'])     # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findOBJ('LINE',[2,6,'1'])                                # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findOBJ('LINE',[2,['nevada',132],'1'])                   # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findOBJ('LINE',[OlxObj.BUS(2),OlxObj.BUS(6),'1'])        # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findOBJ('LINE',"{86b0bd1d-838a-479d-aa6d-7509505bd244}") # by GUID
        l1 = OlxObj.OLCase.findOBJ("{86b0bd1d-838a-479d-aa6d-7509505bd244}")
        l1 = OlxObj.OLCase.findOBJ('LINE',"[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") # by STR
        l1 = OlxObj.OLCase.findOBJ("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")
        #
        l1 = OlxObj.OLCase.findLINE([['claytor',132],['nevada',132],'1'])     # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findLINE([2,6,'1'])                                # by (busNumber1,busNumber2,CID)
        l1 = OlxObj.OLCase.findLINE([2,['nevada',132],'1'])                   # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findLINE([OlxObj.BUS(2),OlxObj.BUS(6),'1'])        # by (bus1,bus2,CID)
        l1 = OlxObj.OLCase.findLINE("{86b0bd1d-838a-479d-aa6d-7509505bd244}")        # by GUID
        l1 = OlxObj.OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") # by STR

        # LINE can be accessed by related object -----------------------------------------------------------------
        l1 = OlxObj.OLCase.LINE[7]   # from NETWORK (OlxObj.OLCase)
        l1 = b1.LINE[0]              # from BUS
        l1 = t1.EQUIPMENT            # from TERMINAL
        l1 = r1.EQUIPMENT            # from RLYGROUP

        # LINE Data --------------------------------------------------------------------------------------------
            'B1','B10','B2','B20','BRCODE','BUS','BUS1','BUS2','CID','DATEOFF','DATEON','EQUIPMENT','FLAG','G1',
            'G10','G2','G20','GUID','I2T','JRNL','KEYSTR','LN','MEMO','METEREDEND','MULINE','NAME','PARAMSTR',
            'PYTHONSTR','R','R0','RATG','RLYGROUP','RLYGROUP1','RLYGROUP2','TAGS','TERMINAL','TIE','TYPE',
            'UNIT','X','X0'
        #
        print('R=',l1.R,' X=',l1.X, ' R0',l1.R0)  # Line R,X,R0
        print( l1.getData('R') )                  # Line R
        print( l1.getData(['R','X']) )            # Line R,X
        print( l1.getData() )                     # Line all Data
        #
        l1.MULINE                                 # list of All MULINE of LINE

        # some methods--------------------------------------------------------------------------------------------
        l1.equals(l2)                    #(bool) comparison of 2 objects
        l1.delete()                      # delete object
        l1.isInList(la)                  # check if object in in List/Set/tuple of object
        l1.postData()                    # Perform validation and update object data in the network database
        l1.toString()                    # (str) text description/composed of object
        l1.copyFrom(l2)                  # copy Data from another object
        l1.changeData(sParam,value)      # change Data (sample_6_changeData())
        l1.getAttributes()               # [str] list attributes of object
        #
        l1.getData('XXX')                # or l1.xx => help details in message error

        # Add LINE -----------------------------------------------------------------------------------------------
        params = {'GUID': '{5f166074-9912-4cd1-a3cd-2e87dd035f84}', 'B1': 0.0104, 'B10': 0.0104, 'B2': 0.0104, 'B20': 0.0104,
             'DATEOFF': '2020/10/30', 'DATEON': '2018/10/30', 'FLAG': 1, 'G1': 0.22, 'G10': 0.23, 'G2': 0.33, 'G20': 0.34, 'I2T': 18,
             'LN': 0, 'MEMO': b'ffe', 'NAME': 'Clav/Fiel', 'R': 0.0472, 'R0': 0.0472, 'RATG': [14, 15, 16, 17], 'TAGS': 'ts1;', 'TIE': 1,
             'TYPE': '', 'UNIT': 'kt', 'X': 0.1983, 'X0': 0.1983}
        keys = [['claytor', 132], ['nevada', 132], '10'] # [bus1,bus2,CID]
        o1 = OlxObj.OLCase.addOBJ('LINE', keys, params)
    """

    print('sample_3_Line')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    # various ways to find a LINE
    l1 = OlxObj.OLCase.findOBJ('LINE', [['claytor', 132], ['nevada', 132], '1'])     # by (bus1,bus2,CID) BUS by (name,kV)
    l1 = OlxObj.OLCase.findOBJ('LINE', [2, 6, '1'])                                  # BUS by (name,kV)
    l1 = OlxObj.OLCase.findOBJ('LINE', [2, ['nevada', 132], '1'])
    l1 = OlxObj.OLCase.findOBJ('LINE', [OlxObj.BUS(2), OlxObj.BUS(6), '1'])
    l1 = OlxObj.OLCase.findOBJ('LINE', "{86b0bd1d-838a-479d-aa6d-7509505bd244}")         # by GUID
    l1 = OlxObj.OLCase.findOBJ("{86b0bd1d-838a-479d-aa6d-7509505bd244}")
    l1 = OlxObj.OLCase.findOBJ('LINE', "[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")  # by STR
    l1 = OlxObj.OLCase.findOBJ("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")
    #
    l1 = OlxObj.OLCase.findLINE([['claytor', 132], ['nevada', 132], '1'])         # by (bus1,bus2,CID)
    l1 = OlxObj.OLCase.findLINE([2, 6, '1'])
    l1 = OlxObj.OLCase.findLINE([2, ['nevada', 132], '1'])
    l1 = OlxObj.OLCase.findLINE([OlxObj.BUS(2), OlxObj.BUS(6), '1'])
    l1 = OlxObj.OLCase.findLINE("{86b0bd1d-838a-479d-aa6d-7509505bd244}")         # by GUID
    l1 = OlxObj.OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")  # by STR

    # print string for object, None if not found
    if l1:
        print(l1.toString())
    else:
        print('Line not found')

    # LINE can be accessed by related object
    l1 = OlxObj.OLCase.LINE[7]
    b1 = OlxObj.BUS(2)
    l1 = b1.LINE[0]
    l1 = OlxObj.OLCase.RLYGROUP[0].EQUIPMENT
    print(l1.toString())

    # LINE Data
    print('R=', l1.R, ' X=', l1.X, ' R0', l1.R0)  # Line R,X,R0
    print(l1.getData('R'))                        # Line R
    print(l1.getData(['R', 'X']))                 # Line R,X
    da = l1.getData()                             # Line all Data
    print('R=', da['R'], ' X=', da['X'])

    # MULINE
    l1 = OlxObj.OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")
    print('All MULINE of LINE: '+l1.toString())
    for mu1 in l1.MULINE:
        print('\t'+mu1.toString())

    #l1.getData('xx')                        # or l1.xx =>help details in message error

    # Add LINE
    params = {'GUID': '{5f166074-9912-4cd1-a3cd-2e87dd035f84}', 'B1': 0.0104, 'B10': 0.0104, 'B2': 0.0104, 'B20': 0.0104,
             'DATEOFF': '2020/10/30', 'DATEON': '2018/10/30', 'FLAG': 1, 'G1': 0.22, 'G10': 0.23, 'G2': 0.33, 'G20': 0.34, 'I2T': 18,
             'LN': 0, 'MEMO': b'ffe', 'NAME': 'Clav/Fiel', 'R': 0.0472, 'R0': 0.0472, 'RATG': [14, 15, 16, 17], 'TAGS': 'ts1;', 'TIE': 1,
             'TYPE': '', 'UNIT': 'kt', 'X': 0.1983, 'X0': 0.1983}
    keys = [['claytor', 132], ['nevada', 132], '10'] # [bus1,bus2,CID]
    o1 = OlxObj.OLCase.addOBJ('LINE', keys, params)


def sample_4_Terminal(olxpath, fi):
    """ TERMINAL object demo

        # various ways to find a TERMINAL
        t1 = OlxObj.OLCase.findOBJ('TERMINAL',[['claytor',132],['nevada',132],'1','LINE'])  # by [bus1,bus2,CID,BRCODE]
        t1 = OlxObj.OLCase.findOBJ('TERMINAL',[2,6,'1','Line'])
        #
        t1 = OlxObj.OLCase.findTERMINAL([['claytor',132],['nevada',132],'1','LINE'])
        t1 = OlxObj.OLCase.findTERMINAL([2,6,'1','Line'])
        #
        x2 = OlxObj.OLCase.findOBJ('XFMR',"[XFORMER] 12 'VERMONT' 33 kV-4 'TENNESSEE' 132 kV 1")
        t1 = x2.TERMINAL[0]


        # TERMINAL Data:
           'BUS','BUS1','BUS2','BUS3','CID','EQUIPMENT','FLAG','KEYSTR','OPPOSITE','REMOTE',
           'RLYGROUP','RLYGROUP1','RLYGROUP2','RLYGROUP3'

        #
        for ti in t1.OPPOSITE:     # opposite to this TERMINAL on the EQUIPMENT
            print(ti.toString())

        for ti in t1.REMOTE:       # Tap bus is ignored
            print(ti.toString())

        #t1.getData('xx')                # or t1.xx help message data of TERMINAL
    """

    print('sample_4_Terminal')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    # various ways to find a TERMINAL
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [['claytor', 132], ['nevada', 132], '1', 'LINE'])  # by [bus1,bus2,CID,BRCODE]
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [2, 6, '1', 'L'])

    t1 = OlxObj.OLCase.findTERMINAL([['claytor', 132], ['nevada', 132], '1', 'LINE'])
    t1 = OlxObj.OLCase.findTERMINAL([2, 6, '1', 'L'])

    # print string for object, None if not found
    if t1:
        print(t1.toString())
    else:
        print('TERMINAL not found')

    # TERMINAL can be accessed by related object
    b1 = OlxObj.OLCase.findBUS(['claytor', 132])
    t1 = b1.TERMINAL[1]
    print(t1.toString())

    l1 = OlxObj.OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")
    print(l1.toString())
    for t1 in l1.TERMINAL:
        print('\t', t1.toString())

    x3 = OlxObj.OLCase.findXFMR3("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-11 'ROANOKE' 13.8 kV 1")
    print(x3.toString())
    for t1 in x3.TERMINAL:
        print('\t', t1.toString())

    # OPPOSITE
    t1 = x3.TERMINAL[0]
    print('\nOPPOSITE TERMINAL:', t1.toString())  # opposite to this TERMINAL on the EQUIPMENT
    for ti in t1.OPPOSITE:
        print('\t', ti.toString())

    #REMOTE (Tap bus is ignored)
    t1 = OlxObj.OLCase.findTERMINAL([['nevada', 132], ['Tap Bus', 132], '1', 'L'])
    print('\nREMOTE TERMINAL:', t1.toString())    # Tap bus is ignored
    for ti in t1.REMOTE:
        print('\t', ti.toString())

    # t1.getData('xx')                # or t1.xx help message data of TERMINAL


def sample_5_changeData(olxpath, fi):
    """
    Object data update demo
        b1.ZONENO = 10                                    # change Data by change Attribute
        b1.changeData('ZONENO',10)                        # by changeData(name,value)
        b1.changeData(['ZONENO','AREANO'],[10,20])        # by changeData(nameLst,valueLst)
        b1.changeData({'MEMO': 'memo1', 'TAGS': 'tag1'})  # by changeData(dict)
        #
        b1.changeData('xx',1)                             #or b1.xx =1  => help in message error
    """

    print('sample_5_changeData\n')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    b1 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])

    # various ways to change data
    b1.ZONENO = 10
    b1.changeData('ZONENO', 10)
    b1.changeData(['ZONENO', 'AREANO'], [10, 20])
    b1.changeData({'MEMO': 'memo1', 'TAGS': 'tag1'})
    #
    b1.postData()  # Perform validation and update object data
    print('ZONENO, AREANO, MEMO, TAGS = %i, %i, %s, %s'%(b1.ZONENO, b1.AREANO,b1.MEMO, b1.getData('TAGS')))

    #b1.changeData('xx',1)                  #or b1.xx =1  => help in message error


def sample_6_Setting(olxpath, fi):
    """ Relay setting get/update demo (OC/DS relays)

        rl = OlxObj.OLCase.findOBJ('RLYDSG', [['NEVADA',132],['REUSENS',132],'1','L','NV_Reusen G1'])
        rl.getSetting('PT ratio')
        rl.changeSetting('PT ratio','101:1')
        rl.changeSetting({'CT ratio':'101:1'})
        rl.postData()


        rl.getSetting('xx')              #=> help in message error
        rl.changeSetting('xx','1')       #=> help in message error
    """

    print('sample_6_Setting\n')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    rl = OlxObj.OLCase.findOBJ('RLYDSG', [['NEVADA', 132], ['REUSENS', 132], '1', 'L', 'NV_Reusen G1'])
    print(rl.toString())

    # getSetting
    print('PT ratio =', rl.getSetting('PT ratio'))
    print('CT ratio =', rl.getSetting()['CT ratio'])

    # changeSetting
    rl.changeSetting('PT ratio', '101:1')   # settingName, value
    rl.changeSetting({'CT ratio':'101:1'})  # dict
    rl.postData()  # Perform validation and update object data
    print('Update PT ratio, CT ratio = %s, %s'%(rl.getSetting('PT ratio'), rl.getSetting('CT ratio')))

    #rl.getSetting('xx')              #=> help in message error
    #rl.changeSetting('xx','1')       #=> help in message error


def sample_7_ClassicalFault_1(olxpath, fi):
    """ Classical Fault on Bus without Outage

        fs1 = OlxObj.SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='2LG:AB', Z=[0.1,0.2])
        OlxObj.OLCase.simulateFault(fs1,0)

        for r1 in OlxObj.FltSimResult:
            print(r1.FaultDescription)
            r1.setScope(tiers=1)
            print('\tIABC fault=',r1.current())
            print('\tIABC t1   =',r1.current(t1))
            print('\tVABC b1   =',r1.voltage(b1))
            print('\top,toc    =',r1.optime(rg1,1,1),rg1.toString())
            #
            r1.xx      #=> help in message error
    """

    print('sample_7_ClassicalFault_1\n')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    b1 = OlxObj.OLCase.findOBJ('BUS', ['nevada', 132])
    b2 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])
    l1 = OlxObj.OLCase.findOBJ('LINE', [b1, b2, '1'])
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [b1, b2, '1', 'L'])
    rl1 = OlxObj.OLCase.findOBJ('RLYDSP', [b2, b1, '1', 'L', 'CLPhase2'])
    rl2 = OlxObj.OLCase.findOBJ('RLYOCP', [['ohio', 132], b1, '1', 'L', 'OH-P1'])

    # Define the fault specification
    fs1 = OlxObj.SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='2LG:AB', Z=[0.1, 0.2])

    # Run fault simulation
    OlxObj.OLCase.simulateFault(fs1, 1)  # (0/1) clear previous result flag

    # new Fault '3LG' & Z
    fs1.obj = b2
    fs1.fltConn = '3LG'
    fs1.z = [0.2, 0.3]
    OlxObj.OLCase.simulateFault(fs1, 0)  # (0/1) clear previous result flag

    # RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)                 # Description of Fault
        # set scope for compute solution results
        r1.setScope(tiers=1)
        print('\tIABC fault=', r1.current())        # ABC form current at Fault

        # ABC form current on Terminal
        print('\tIABC t1   =', r1.current(t1))
        print('\tVABC b1   =', r1.voltage(b1))      # ABC form voltage at BUS
        print('\tVABC l1   =', r1.voltage(l1))      # ABC form voltage at LINE

        # 012 sequence form current(zero/positive/negative)
        print('\tI012 fault=', r1.currentSeq())

        # 012 sequence form current on Terminal
        print('\tI012 t1   =', r1.currentSeq(t1))

        # 012 sequence form voltage at BUS
        print('\tV012 b1   =', r1.voltageSeq(b1))

        # 012 sequence form voltage at LINE
        print('\tV012 l1   =', r1.voltageSeq(l1))

        #
        print('\top,toc    =', r1.optime(rl1, 1, 1), rl1.toString())  # Relay operating time
        print('\top,toc    =', r1.optime(rl2, 1, 1), rl2.toString())  # Relay operating time

        # Thevenin impedance
        print('\tThevenin Zp,Zn,Z0 =', r1.Thevenin)

        # Short circuit MVA
        print('\tMVA =', r1.MVA)
        print('\tXR_ratio [X/R,ANSI X/R,R0/X1,X0/X1] =', r1.XR_ratio)  # Short circuit X/R ratio

        #r1.xx      #=> help in message error


def sample_8_ClassicalFault_2(olxpath,fi):
    """ Classical Fault on Line with Outage
    """

    print('\nsample_8_ClassicalFault_2')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    b1 = OlxObj.OLCase.findOBJ('BUS', ['nevada', 132])
    b2 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [b2, b1, '1', 'LINE'])
    l1 = OlxObj.OLCase.findOBJ('LINE', [b2, b1, '1'])

    # Define the fault outage
    o1 = OlxObj.OUTAGE(option='single-gnd', G='0.4')
    o1.add_outageLst(l1)
    o1.build_outageLst(obj=t1, tiers= '2', wantedTypes=['T'])
    print('List Outage')
    for v1 in o1.outageLst:
        print('\t', v1.toString())

    # Define the fault specification
    fs1 = OlxObj.SPEC_FLT.Classical(
        obj=b1, fltApp='BUS', fltConn='2LG:AB', outage=o1)

    # Run fault simulation
    OlxObj.OLCase.simulateFault(fs1, 1)   # (0/1) clear previous result flag
    o1.option = 'DOUBLE'  # or fs1.outage.option = 'DOUBLE'

    OlxObj.OLCase.simulateFault(fs1, 0)   # (0/1) clear previous result flag

    # RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)                # Description of Fault
        print('\tIABC fault=', r1.current())       # ABC form current at fault

        # ABC form current on Terminal t1
        print('\tIABC t1=', r1.current(t1))

        # 012 sequence form current (zero/positive/negative) at fault
        print('\tI012 fault=', r1.currentSeq())

        # 012 sequence form current on Terminal t1
        print('\tI012 t1=', r1.currentSeq(t1))

        #r1.xx      #=> help in message error


def sample_9_SimultaneousFault(olxpath, fi):
    """ Simultaneous Fault
    """

    print('\nsample_9_SimultaneousFault')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    b1 = OlxObj.OLCase.findOBJ('BUS', ['nevada', 132])
    b2 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])
    b3 = OlxObj.OLCase.findOBJ('BUS', ['FIELDALE', 132])
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [b2, b1, '1', 'L'])
    l1 = OlxObj.OLCase.findOBJ('LINE', [b2, b1, '1'])
    rg = OlxObj.OLCase.findOBJ('RLYGROUP', [b1, b2, '1', 'L'])

    # Define the fault specification
    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=b1, fltApp='Bus', fltConn='ll:bc')
    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=t1, fltApp='2P-open', fltConn='ac')
    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=[b1, b3], fltApp='Bus2Bus', fltConn='cb')
    fs1 = OlxObj.SPEC_FLT.Simultaneous(obj=rg, fltApp='CLOSE-IN', fltConn='3LG')
    fs2 = OlxObj.SPEC_FLT.Simultaneous(obj=t1, fltApp='15%', fltConn='3LG', Z=[0.1, 0.2, 0.5, 0, 0, 0, 0, 0.5])

    # Run fault simulation
    OlxObj.OLCase.simulateFault([fs1], 1)
    OlxObj.OLCase.simulateFault([fs1, fs2], 0)

    # RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)                  # Description of Fault

        # ABC form current at fault
        print('\tIABC fault=', r1.current())
        print('\tVABC b2   =', r1.voltage(b2))       # ABC form Voltage at BUS

        # 012 sequence form Voltage at LINE
        print('\tV012 l1   =', r1.voltageSeq(l1))

        #r1.xx      #=> help in message error


def sample_10_SEA_1(olxpath,fi):
    """
    SEA: Stepped-Event Analysis
        Single User-Defined Event
    """

    print('\nsample_10_SEA_1')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    b1 = OlxObj.OLCase.findOBJ('BUS', ['nevada', 132])
    b2 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [b1, b2, '1', 'L'])
    t2 = OlxObj.OLCase.findOBJ('TERMINAL', [b2, b1, '1', 'L'])
    rg = OlxObj.OLCase.findOBJ('RLYGROUP', [2, 6, '1', 'L'])

    # Define the fault specification
    fs1 = OlxObj.SPEC_FLT.SEA(obj=t1, fltApp='15%', fltConn='1LG:A', deviceOpt=[1, 1, 1, 1, 1, 1, 1], tiers=5, Z=[0, 0])

    # Run fault simulation
    OlxObj.OLCase.simulateFault(fs1)

    # RESULT
    for r1 in OlxObj.FltSimResult:
        # Description of Fault
        print(r1.FaultDescription)

        # SEA event time stamp [s]
        print('\ttime      =', r1.SEARes['time'])

        # Highest phase fault current magnitude at this step
        print('\tImax      =', r1.SEARes['currentMax'])

        # ABC form current at fault
        print('\tIABC Fault=', r1.current())

        # ABC form current om Terminal
        print('\tIABC t2   =', r1.current(t2))

        # 012 sequence form Voltage at BUS
        print('\tV012 b1   =', r1.voltageSeq(b1))

        # r1.xx      #=> help in message error


def sample_11_SEA_2(olxpath, fi):
    """
    SEA: Stepped-Event Analysis
        Multiple User-Defined Event
    """

    print('\nsample_11_SEA_2')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    b1 = OlxObj.OLCase.findOBJ('BUS', ['nevada', 132])
    b2 = OlxObj.OLCase.findOBJ('BUS', ['claytor', 132])
    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [b1, b2, '1', 'L'])
    t2 = OlxObj.OLCase.findOBJ('TERMINAL', [b2, b1, '1', 'L'])
    rg = OlxObj.OLCase.findOBJ('RLYGROUP', [2, 6, '1', 'L'])

    # Define the fault specification
    fs1 = OlxObj.SPEC_FLT.SEA(obj=t1, fltApp='15%', fltConn='1LG:A', deviceOpt=[1, 1, 1, 1, 1, 1, 1], tiers=5, Z=[0, 0])
    fs2 = OlxObj.SPEC_FLT.SEA_EX(time=0.01, fltConn='2LG:AB', Z=[0.01, 0.02])
    fs3 = OlxObj.SPEC_FLT.SEA_EX(time=0.02, fltConn='3LG', Z=[0.01, 0.02])

    # Run fault simulation
    OlxObj.OLCase.simulateFault([fs1, fs2, fs3])

    # RESULT
    for r1 in OlxObj.FltSimResult:
        print(r1.FaultDescription)                      # Description of Fault

        # SEA event time stamp [s]
        print('\ttime     =', r1.SEARes['time'])

        # (0/1) flag showing is this is an user-defined event
        print('\tEventFlag=', r1.SEARes['EventFlag'])

        # (str) Event Description
        print('\tEventDesc=', r1.SEARes['EventDesc'])

        # ABC form current om Terminal
        print('\tIABC t2  =', r1.current(t2))

        # 012 sequence form Voltage at BUS
        print('\tV012 b1  =', r1.voltageSeq(b1))


def sample_12_tapLine(olxpath,fi):

    print('\nsample_12_tapLine')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    t1 = OlxObj.OLCase.findOBJ('TERMINAL', [6, ['Tap Bus', 132], '1', 'L'])
    res = OlxObj.OLCase.tapLineTool(t1)
    #
    mainLine = res['mainLine']      # List of all TERMINAL in the main-line
    allPath = res['allPath']        # List of all TERMINAL of all path
    localBus = res['localBus']      # List of local buses on the mainLine
    remoteBus = res['remoteBus']    # List of remote BUSES on the mainLine

    # List of all relay groups at the local end on the mainLine
    localRLG = res['localRLG']

    # List of all RLYGROUP at the remote end on the mainLine
    remoteRLG = res['remoteRLG']

    # positive sequence Impedance of the mainLine
    Z1 = res['Z1']
    Z0 = res['Z0']                  # zero sequence Impedance of the mainLine
    Length = res['Length']          # sum length of the mainLine

    #
    for i in range(len(mainLine)):
        print('mainLine: ', i+1)
        print('\tlocalBus : '+OlxObj.toString(localBus[i]))
        print('\tremoteBus: '+OlxObj.toString(remoteBus[i]))
        print('\tlocalRLG : '+OlxObj.toString(localRLG[i]))
        print('\tremoteRLG: '+OlxObj.toString(remoteRLG[i]))
        print('\tSegments:')
        for v1 in mainLine[i]:
            print('\t\t', v1.toString())
        print('\tZ1      : ', OlxObj.toString(Z1[i]))
        print('\tZ0      : ', OlxObj.toString(Z0[i]))
        print('\tLength  : ', OlxObj.toString(Length[i]))

    #
    for i in range(len(allPath)):
        print('Segments of path (%i)' % (i+1))
        for v1 in allPath[i]:
            print('\t', v1.toString())


def sample_13_addObj_1(olxpath,fi):
    """
    Add Object demo:
        BUS:      obn = OlxObj.OLCase.addOBJ('BUS', keys, params)
        LINE:     obn = OlxObj.OLCase.addOBJ('LINE', keys, params)
        XFMR3:    obn = OlxObj.OLCase.addOBJ('XFMR3', keys, params)
        RLYGROUP: obn = OlxObj.OLCase.addOBJ('RLYGROUP', keys, params)
        RLYV:     obn = OlxObj.OLCase.addOBJ('RLYV', keys, params)
    """

    print('\nsample_13_addObj_1')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    # BUS
    params = {'AREANO': 3, 'LOCATION': 'CLAYTOR', 'MEMO': 'm1', 'MIDPOINT': 1, 'NO': 0, 'SLACK': 1,
             'SPCX': 15, 'SPCY': 25, 'SUBGRP': 10, 'TAGS': 'tag1;tag2;', 'TAP': 0, 'ZONENO': 2, 'NAME': 5}
    keys = ['fieldale_1', '132.']
    o1 = OlxObj.OLCase.addOBJ('BUS', keys, params)
    print('\tNew BUS:', o1.toString())


    # LINE
    params = {'B1': '0.0104', 'B10': 0.0104, 'B2': 0.0104, 'B20': 0.0104, 'DATEOFF': '2020/10/30', 'DATEON': '2018/10/30', 'FLAG': 1, 'G1': 0.22,
             'G10': 0.23, 'G2': 0.33, 'G20': 0.34, 'I2T': 18, 'LN': 0, 'MEMO': 'memo', 'NAME': 'Clav/Fiel', 'R': 0.0472, 'R0': 0.0472,
             'RATG': [14, 15, 16, 17], 'TAGS': 'ts1;', 'TIE': 1, 'TYPE': '', 'UNIT': 'kt', 'X': '0.1983', 'X0': 0.1983}
    keys = [['claytor', '132'], ['nevada', '132'], '10']
    o1 = OlxObj.OLCase.addOBJ('LINE', keys, params)
    print('\tNew LINE:', o1.toString())


    # XFMR3
    params = {'GUID': '{9106257f-cd51-4ec8-ae8d-24451b972323}', 'AUTOX': 1, 'B': -0.561, 'B0': -0.57, 'BASEMVA': 101,
             'CONFIGP': 'G', 'CONFIGS': 'G', 'CONFIGST': 'G', 'CONFIGT': 'D', 'CONFIGTT': 'D', 'DATEOFF': '2016/5/1',
             'DATEON': '2015/4/1', 'FICTBUSNO': 0, 'FLAG': 1, 'GANGED': 1, 'LTCCENTER': 132, 'LTCCTRL': "[BUS] 28 'ARIZONA' 132 kV",
             'LTCSIDE': 2, 'LTCSTEP': 0.00625, 'LTCTYPE': 0, 'MAXTAP': 1.5, 'MAXVW': 1.5, 'MEMO': b'gth', 'MINTAP': 0.51,
             'MINVW': 0.51, 'MVA1': 15, 'MVA2': 16, 'MVA3': 17, 'NAME': 'Nev/NH/Rnk', 'PRIORITY': 0, 'PRITAP': 33, 'RG1': 0.21,
             'RG2': 0, 'RG3': 0, 'RGN': 0, 'RMG0': 0.33, 'RPM0': 0.33, 'RPS': 0.11, 'RPS0': 0.2, 'RPT': 0.12, 'RPT0': 0.2, 'RSM0': 0.33,
             'RST': 0.13, 'RST0': 0.24, 'SECTAP': 132, 'TAGS': 't1;t2;', 'TERTAP': 13.8, 'XG1': 0.22, 'XG2': 0, 'XG3': 0, 'XGN': 0,
             'XMG0': 0.33, 'XPM0': 0.33, 'XPS': 0.32, 'XPS0': 0.26, 'XPT': 0.42, 'XPT0': 0.45, 'XSM0': 0.33, 'XST': 0.33,
             'XST0': 0.45, 'Z0METHOD': 1}
    keys = [['nevada', 132], ['NEW HAMPSHR', 33], ['ROANOKE', 13.8], '50']
    o1 = OlxObj.OLCase.addOBJ('XFMR3', keys, params)
    print('\tNew XFMR3:', o1.toString())


    # GENUNIT
    params = {'GUID': '{8378e9b7-22a9-49e0-8b49-e3974e63ccf3}', 'DATEOFF': '2017/1/5', 'DATEON': '2016/1/1', 'FLAG': 1,
             'MVARATE': 100, 'PMAX': 50, 'PMIN': 20, 'QMAX': 40, 'QMIN': -30, 'R': [0.021, 0.023, 0.022, 0.024, 0.025], 'RG': 0.14,
             'SCHEDP': 40, 'SCHEDQ': 30, 'TAGS': 'tag1;', 'X': [0.24, 0.26, 0.25, 0.27, 0.28], 'XG': 0.15}
    keys = [['OHIO', '132'], 2]
    o1 = OlxObj.OLCase.addOBJ('GENUNIT', keys, params)
    print('\tNew GENUNIT:', o1.toString())


    # RLYGROUP
    params = {'GUID': '{48c97fdd-5fab-4cbd-8601-18b5b50856e6}', 'BACKUP': [],
             'INTRPTIME': 0.5, 'MEMO': 'grtg', 'PRIMARY': [], 'RECLSRTIME': [2.11, 1, 0.5, 0]}
    keys = [['NEVADA', 132], ['OHIO', 132], '1', 'L']
    o1 = OlxObj.OLCase.addOBJ('RLYGROUP', keys, params)
    print('\tNew RLYGROUP:', o1.toString())


    # RLYV
    params = {'GUID': '{0fc246b1-63f6-4266-adab-82fdb0b38928}', 'ASSETID': 'aid1', 'DATEOFF': '2016/6/7', 'DATEON': '2015/5/6',
             'FLAG': '1', 'MEMO': b'vm1', 'OPQTY': 1, 'OVDELAYTD': 0, 'OVINST': 0, 'OVPICKUP': 0.1, 'PACKAGE': 1, 'PTR': 1,
              'SGLONLY': [1, 0, 1,0],'TAGS': 'g1;', 'UVDELAYTD': 0, 'UVINST': 0, 'UVPICKUP': 0.12}
    keys = [['NEVADA', 132], ['OHIO', 132], '1', 'L', 'v1']
    o1 = OlxObj.OLCase.addOBJ('RLYV', keys, params)
    print('\tNew RLYV:', o1.toString())


def sample_14_addObj_2(olxpath,fi):
    """
    Add Object demo:
        RLYOCG
        RLYDSG
        RECLSR
        SCHEME
    """

    print('\nsample_14_addObj_2')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    # RLYOCG
    params = {'GUID': '{413a2dac-d26a-4ae0-aa0f-dd69053c6aa3}', 'ASSETID': 'aid122', 'ASYM': 0, 'CT': 100, 'CTLOC': 1, 'DATEOFF': '2018/7/8',
             'DATEON': '2015/3/5', 'DTDELAY': [0.2, 0.1, 0, 0, 0], 'DTDIR': 0, 'DTPICKUP': [23, 25, 0, 0, 0], 'DTTIMEADD': 0, 'DTTIMEMULT': 1,
             'FLAG': 1, 'FLATINST': 0, 'INSTSETTING': 2600, 'LIBNAME': 'ABB', 'MEMO': b'bgn', 'MINTIME': 0, 'OCDIR': 1, 'OPI': 0,
             'PACKAGE': 0, 'PICKUPTAP': 2.5, 'POLAR': 2, 'SGNL': 0, 'TAGS': 'tggg1;', 'TAPTYPE': 'W1', 'TDIAL': 2, 'TIMEADD': 0,
             'TIMEMULT': 1, 'TIMERESET': 0, 'TYPE': 'CO-5'}
    settings = {'PTR: PT ratio': '98', 'Z2F: Fwd threshold': '0.68', 'Z2R: Rev threshold': '0.3', 'a2: I2/I1 restr.factor': '0.1',
               'k2: i2/I0 restr.factor': '0.1', 'Z1ANG: Line Z1 angle': '78', '50QF: Fwd 3I2 pickup': '0.42', '50QR: Rev 3I2 pickup': '0.52'}
    keys = [['NEVADA', 132], ['OHIO', 132], '1', 'L', 'NV-G10']
    o1 = OlxObj.OLCase.addOBJ('RLYOCG', keys, params, settings)
    print('\tNew RLYOCG:', o1.keystr)
    print('\n\tparams:', o1.paramstr)
    print('\n\tsettings:', o1.settingstr)


    # RLYDSG
    params = {'GUID': '{1a065b6a-6bbf-4c01-93f2-53832e17297a}', 'ASSETID': 'fg', 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'DSTYPE': 'CEY-Type',
             'FLAG': 1, 'TYPE': 'CEYX', 'LIBNAME': 'COOPER', 'MEMO': b'bt', 'PACKAGE': 0, 'SNLZONE': 1, 'TAGS': '', 'VTBUS': "[BUS] 6 'NEVADA' 132 kV",
             'Z2OCPICKUP': 3.1, 'Z2OCTD': 1.15, 'Z2OCTYPE': '240-91-16540'}
    settings = {'CT ratio': '500:2', 'CT polarity is reversed': 'No', 'PT ratio': '101:1', 'Min I': '0.2', 'Zone 1 delay': '0.2', 'K1 Mag': '0.59',
               'K1 Ang': '0.58', 'K2 Mag': '0.51', 'K2 Ang': '0.65', 'Z1_Imp.': '0.54', 'Z1_Ang.': '75', 'Z2_Offset Imp.': '0.56',
               'Z2_Offset Ang.': '0.65', 'Z2_Imp.': '0.48', 'Z2_Ang.': '75', 'Z2_Delay': '0.66', 'Z3_Offset Imp.': '0.22', 'Z3_Offset Ang.': '0.64',
               'Z3_Imp.': '0.77', 'Z3_Ang.': '75.6', 'Z3_Delay': '-0.6', 'Z3_Frwrd(1)/Rev(0)': '1'}
    keys = [['NEVADA', 132], ['OHIO', 132], '1', 'L', 'DSG20']
    o1 = OlxObj.OLCase.addOBJ('RLYDSG', keys, params, settings)
    print('\tNew RLYDSG:', o1.toString())
    print('\n\tparams:', o1.paramstr)
    print('\n\tsettings:', o1.settingstr)


    # RECLSR
    params = {'GUID': '{1555ccea-8595-4846-a171-579c4d490959}', 'ASSETID': 'aid11', 'BYTADD': '2', 'BYTMULT': 2, 'DATEOFF': '2018/1/2', 'DATEON': '2015/6/7',
             'FASTOPS': 2, 'FLAG': 1, 'GR_FASTTYPE': '240-91-56-05', 'GR_INST': 0.56, 'GR_INSTDELAY': 0.57, 'GR_MEMO': b'gth', 'GR_MINTIMEF': 0.12,
             'GR_MINTIMES': 0.14, 'GR_MINTRIPF': 1.52, 'GR_MINTRIPS': 1.42, 'GR_SLOWTYPE': '240-91-56-06', 'GR_TIMEADDF': 0.34, 'GR_TIMEADDS': 0.36,
             'GR_TIMEMULTF': 1.62, 'GR_TIMEMULTS': 1.64, 'INTRPTIME': 0.34, 'LIBNAME': 'COOPER', 'PH_FASTTYPE': '240-91-56-02', 'PH_INST': 0.54,
             'PH_INSTDELAY': 0.55, 'PH_MEMO': b'rgg', 'PH_MINTIMEF': 0.11, 'PH_MINTIMES': 0.13, 'PH_MINTRIPF': 1.51, 'PH_MINTRIPS': 1.41,
             'PH_SLOWTYPE': '240-91-56-01', 'PH_TIMEADDF': 0.33, 'PH_TIMEADDS': 0.35, 'PH_TIMEMULTF': 1.61, 'PH_TIMEMULTS': 1.63, 'RATING': 1.2,
             'RECLOSE1': 0.15, 'RECLOSE2': 0.25, 'RECLOSE3': 0.35, 'TAGS': 'gjh;', 'TOTALOPS': 4}
    keys = [['NEVADA', 132], ['REUSENS', 132], '1', 'L', 'fd4']
    o1 = OlxObj.OLCase.addOBJ('RECLSR', keys, params)
    print('New RECLSR:', o1.toString())
    print('\n\tparams:', o1.paramstr)


    # SCHEME
    params = {'GUID':'{a2f5356f-819e-416d-be69-034503de8675}','ASSETID':'aied', 'DATEOFF':'N/A', 'DATEON':'N/A', 'FLAG':1,
         'MEMO':b'rf', 'SIGNALONLY':'1', 'TAGS':'t1', 'TYPE':'PUTT'}
    logics = {'EQUATION': 'RU_NEAR + (RU_FAR @ TS * RO_NEAR) +RO_NEAR',
        'RU_NEAR': ['INST/DT PICKUP', "[OCRLYG]  FL-G1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"],
        'RU_FAR': ['OPEN OP.', "[TERMINAL] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"],
        'TS': 0.43,
        'RO_NEAR': ['PICKUP 3I2', "[DEVICEVR]  rlv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"]}
    keys = [['FIELDALE', 132], ['CLAYTOR', 132], 1, 'L', 'scheme10']
    o1 = OlxObj.OLCase.addOBJ('SCHEME', keys, params, logics)
    print('New SCHEME:', o1.toString())
    print('\n\tparams:', o1.paramstr)
    print('\n\tlogics:', o1.logicstr)


def sample_15_XFMR(olxpath, fi):
    """ (2 Windings Transformer) XFMR object demo

        # various ways to find a XFMR
        x2 = OlxObj.OLCase.findOBJ('XFMR',"_{4D49357D-E376-4776-8618-D47CA9490EC3}")              # GUID
        x2 = OlxObj.OLCase.findOBJ("_{4D49357D-E376-4776-8618-D47CA9490EC3}")                     # GUID
        x2 = OlxObj.OLCase.findOBJ('XFMR',"[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1") # STR
        x2 = OlxObj.OLCase.findOBJ("[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1")        # STR
        x2 = OlxObj.OLCase.findOBJ('XFMR',[['NEVADA',132],['NEW HAMPSHR',33],'1'])                # [b1,b2,CID]
        x2 = OlxObj.OLCase.findOBJ('XFMR',[6,['NEW HAMPSHR',33],'1'])                             # [b1,b2,CID]
        x2 = OlxObj.OLCase.findOBJ('XFMR',["[BUS] 6 'NEVADA' 132 kV",['NEW HAMPSHR',33],'1'])     # [b1,b2,CID]


        # XFMR can be accessed by related object
        x2 = OlxObj.OLCase.XFMR[0]     # from NETWORK (OlxObj.OLCase)
        x2 = b1.XFMR[0]                # from BUS (b1)
        x2 = t1.EQUIPMENT              # from TERMINAL (t1)
        x2 = r1.EQUIPMENT              # from RLYGROUP


        # XFMR Data :'AUTOX','B','B0','B1','B10','B2','B20','BASEMVA','BRCODE','BUS','BUS1','BUS2','CID',
                     'CONFIGP','CONFIGS','CONFIGST','DATEOFF','DATEON','EQUIPMENT','FLAG','G1','G10','G2',
                     'G20','GANGED','GUID','JRNL','KEYSTR','LTCCENTER','LTCCTRL','LTCSIDE','LTCSTEP','LTCTYPE',
                     'MAXTAP','MAXVW','MEMO','METEREDEND','MINTAP','MINVW','MVA','MVA1','MVA2','MVA3','NAME',
                     'PARAMSTR','PRIORITY','PRITAP','PYTHONSTR','R','R0','RG1','RG2','RGN','RLYGROUP',
                     'RLYGROUP1','RLYGROUP2','SECTAP','TAGS','TERMINAL','TIE','X','X0','XG1','XG2','XGN'

        #
        print(x2.NAME,x2.R,x2.R0)            # (str,float,float) XFMR Name,R,R0
        print(x2.getData('NAME'))            # (str) XFMR Name
        print(x2.getData(['NAME','R','R0'])) # (dict) XFMR Name,R,R0
        print(x2.getData())                  # (dict) get all Data of XFMR


        # some methods
        x2.equals(x22)                   #(bool) comparison of 2 objects
        x2.delete()                      # delete object
        x2.isInList(la)                  # check if object in in List/Set/tuple of object
        x2.toString()                    # (str) text description/composed of object
        x2.copyFrom(x22)                 # copy Data from another XFMR
        x2.changeData(sParam,value)      # change Data (sample_6_changeData())
        x2.postData()                    # Perform validation and update object data in the network database
        x2.getAttributes()               # [str] list attributes of object
        #
        x2.getData('XXX')                # help details in message error


        # Add XFMR:
        params = {'GUID':'{1d49357d-e376-4776-8618-d47ca9490ec3}','AUTOX':1,'B':-0.64,'B0':-0.6,'B1':0.14,'B10':0.16,'B2':0.11,
            'B20':0.11,'BASEMVA':20,'CONFIGP':'G','CONFIGS':'E','CONFIGST':'D','DATEOFF':'2015/9/1','DATEON':'2016/9/1',
            'FLAG':1,'G1':0.13,'G10':0.15,'G2':0.11,'G20':0.11,'GANGED':0,'LTCCENTER':132,'LTCCTRL':None,'LTCSIDE':0,
            'LTCSTEP':0.00625,'LTCTYPE':0,'MAXTAP':1.5,'MAXVW':1.5,'MEMO':b'th 22','MINTAP':0.51,'MINVW':0.51,'MVA1':20,
            'MVA2':26,'MVA3':32,'NAME':'a2','PRIORITY':1,'PRITAP':132,'R':0.1,'R0':0.12,'RG1':0.1,'RG2':0,'RGN':0,
            'SECTAP':33,'TAGS':'xt1;xt2;','TIE':2,'X':0,'X0':0.11,'XG1':0.1,'XG2':0,'XGN':0}
        keys = [['NEVADA', 132], ['NEW HAMPSHR', 33], '10']
        o1 = OlxObj.OLCase.addOBJ('XFMR', keys, params)
    """

    print('\nsample_15_XFMR')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    x2 = OlxObj.OLCase.findOBJ('XFMR', ["[BUS] 6 'NEVADA' 132 kV", ['NEW HAMPSHR', 33], '1'])
    if x2!=None:
        print(x2.toString())
        print('NAME, X: %s, %f'%(x2.name, x2.x))
        print(x2.getData(['NAME', 'X', 'X0']))


    # Add
    params = {'GUID': '{1d49357d-e376-4776-8618-d47ca9490ec3}', 'AUTOX': 1, 'B': -0.64, 'B0': -0.6, 'B1': 0.14, 'B10': 0.16, 'B2': 0.11,
             'B20': 0.11, 'BASEMVA': 20, 'CONFIGP': 'G', 'CONFIGS': 'E', 'CONFIGST': 'D', 'DATEOFF': '2015/9/1', 'DATEON': '2016/9/1',
             'FLAG': 1, 'G1': 0.13, 'G10': 0.15, 'G2': 0.11, 'G20': 0.11, 'GANGED': 0, 'LTCCENTER': 132, 'LTCCTRL': None, 'LTCSIDE': 0,
             'LTCSTEP': 0.00625, 'LTCTYPE': 0, 'MAXTAP': 1.5, 'MAXVW': 1.5, 'MEMO': b'th 22', 'MINTAP': 0.51, 'MINVW': 0.51, 'MVA1': 20,
             'MVA2': 26, 'MVA3': 32, 'NAME': 'a2', 'PRIORITY': 1, 'PRITAP': 132, 'R': 0.1, 'R0': 0.12, 'RG1': 0.1, 'RG2': 0, 'RGN': 0,
             'SECTAP': 33, 'TAGS': 'xt1;xt2;', 'TIE': 2, 'X': 0, 'X0': 0.11, 'XG1': 0.1, 'XG2': 0, 'XGN': 0}
    keys = [['NEVADA', 132], ['NEW HAMPSHR', 33], '10']
    o1 = OlxObj.OLCase.addOBJ('XFMR', keys, params)
    print('New XFMR: ', o1.toString())


def sample_16_XFMR3(olxpath, fi):
    """ (3 Windings Transformer) XFMR3 object demo

        # various ways to find a XFMR3
        x3 = OlxObj.OLCase.findOBJ('XFMR3',"{ce406fc7-6e2d-41af-8adb-9ca8375d7916}")              # GUID
        x3 = OlxObj.OLCase.findOBJ("{ce406fc7-6e2d-41af-8adb-9ca8375d7916}")                      # GUID
        x3 = OlxObj.OLCase.findOBJ('XFMR3',"[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-11 'ROANOKE' 13.8 kV 1") # STR
        x3 = OlxObj.OLCase.findOBJ("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-11 'ROANOKE' 13.8 kV 1")         # STR
        x3 = OlxObj.OLCase.findOBJ('XFMR3',[['NEVADA',132],['NEW HAMPSHR',33],'1'])               # [b1,b2,CID]
        x3 = OlxObj.OLCase.findOBJ('XFMR3',[6,10,11,'1'])                                         # [b1,b2,b3,CID]
        x3 = OlxObj.OLCase.findOBJ('XFMR3',["[BUS] 6 'NEVADA' 132 kV",11,'1'])                    # [b1,b2,CID]


        # XFMR3 can be accessed by related object
        x3 = OlxObj.OLCase.XFMR3[0]     # from NETWORK (OlxObj.OLCase)
        x3 = b1.XFMR3[0]                # from BUS (b1)
        x3 = t1.EQUIPMENT               # from TERMINAL (t1)
        x3 = r1.EQUIPMENT               # from RLYGROUP (r1)

        # XFMR3 Data : 'AUTOX','B','B0','BASEMVA','BRCODE','BUS','BUS1','BUS2','BUS3','CID','CONFIGP','CONFIGS','CONFIGST',
            'CONFIGT','CONFIGTT','DATEOFF','DATEON','EQUIPMENT','FICTBUSNO','FLAG','GANGED','GUID','JRNL','KEYSTR','LTCCENTER',
            'LTCCTRL','LTCSIDE','LTCSTEP','LTCTYPE','MAXTAP','MAXVW','MEMO','MINTAP','MINVW','MVA1','MVA2','MVA3','NAME',
            'PARAMSTR','PRIORITY','PRITAP','PYTHONSTR','RG1','RG2','RG3','RGN','RLYGROUP','RLYGROUP1','RLYGROUP2','RLYGROUP3',
            'RMG0','RPM0','RPS','RPS0','RPT','RPT0','RSM0','RST','RST0','SECTAP','TAGS','TERMINAL','TERTAP','XG1','XG2','XG3',
            'XGN','XMG0','XPM0','XPS','XPS0','XPT','XPT0','XSM0','XST','XST0','Z0METHOD'
        #
        print(x3.NAME,x3.XPT,x3.XPT0)            # (str,float,float) XFMR3 Name,XPT,XPT0
        print(x3.getData('NAME'))                # (str) XFMR3 Name
        print(x3.getData(['NAME','XPT','XPT0'])) # (dict) XFMR3 Name,XPT,XPT0
        print(x3.getData())                      # (dict) get all Data of XFMR3

        # some methods--------------------------------------------------------------------------------------------
        x3.equals(x32)                   # (bool) comparison of 2 objects
        x3.delete()                      # delete object
        x3.isInList(la)                  # check if object in in List/Set/tuple of object
        x3.toString()                    # (str) text description/composed of object
        x3.copyFrom(x32)                 # copy Data from another object
        x3.changeData(sParam,value)      # change Data (sample_6_changeData())
        x3.postData()                    # Perform validation and update object data in the network database
        x3.getAttributes()               # [str] list attributes of object
        #
        x3.getData('XXX')                # help details in message error

        # Add XFMR3
        params = {'GUID':'{1106257f-cd51-4ec8-ae8d-24451b972323}','AUTOX':1,'B':-0.561,'B0':-0.57,'BASEMVA':101,'CONFIGP':'G',
            'CONFIGS':'G','CONFIGST':'G','CONFIGT':'D','CONFIGTT':'D','DATEOFF':'','DATEON':'','FICTBUSNO':0,'FLAG':1,'GANGED':1,
            'LTCCENTER':132,'LTCCTRL':"[BUS] 28 'ARIZONA' 132 kV",'LTCSIDE':2,'LTCSTEP':0.00625,'LTCTYPE':0,'MAXTAP':1.5,
            'MAXVW':1.5,'MEMO':b'gth','MINTAP':0.51,'MINVW':0.51,'MVA1':15,'MVA2':16,'MVA3':17,'NAME':'Nev/NH/Rnk',
            'PRIORITY':0,'PRITAP':33,'RG1':0.21,'RG2':0,'RG3':0,'RGN':0,'RMG0':0.33,'RPM0':0.33,'RPS':0.11,'RPS0':0.2,
            'RPT':0.12,'RPT0':0.2,'RSM0':0.33,'RST':0.13,'RST0':0.24,'SECTAP':132,'TAGS':'t1;t2;','TERTAP':13.8,'XG1':0.22,
            'XG2':0,'XG3':0,'XGN':0,'XMG0':0.33,'XPM0':0.33,'XPS':0.32,'XPS0':0.26,'XPT':0.42,'XPT0':0.45,'XSM0':0.33,
            'XST':0.33,'XST0':0.45,'Z0METHOD':1}
        keys = [['NEVADA', 132], ['NEW HAMPSHR', 33], ['ROANOKE', 13.8], '10']
        o1 = OlxObj.OLCase.addOBJ('XFMR3', keys, params)
    """

    print('\nsample_16_XFMR3')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    x3 = OlxObj.OLCase.findOBJ("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-11 'ROANOKE' 13.8 kV 1")
    if x3!=None:
        print(x3.toString())
        print('NAME, XPT: %s, %f'%(x3.name, x3.xpt))
        print(x3.getData(['NAME', 'XPT', 'XPT0']))

    # Add
    params = {'GUID': '{1106257f-cd51-4ec8-ae8d-24451b972323}', 'AUTOX': 1, 'B': -0.561, 'B0': -0.57, 'BASEMVA': 101, 'CONFIGP': 'G',
             'CONFIGS': 'G', 'CONFIGST': 'G', 'CONFIGT': 'D', 'CONFIGTT': 'D', 'DATEOFF': '', 'DATEON': '', 'FICTBUSNO': 0, 'FLAG': 1, 'GANGED': 1,
             'LTCCENTER': 132, 'LTCCTRL': "[BUS] 28 'ARIZONA' 132 kV", 'LTCSIDE': 2, 'LTCSTEP': 0.00625, 'LTCTYPE': 0, 'MAXTAP': 1.5,
             'MAXVW': 1.5, 'MEMO': b'gth', 'MINTAP': 0.51, 'MINVW': 0.51, 'MVA1': 15, 'MVA2': 16, 'MVA3': 17, 'NAME': 'Nev/NH/Rnk',
             'PRIORITY': 0, 'PRITAP': 33, 'RG1': 0.21, 'RG2': 0, 'RG3': 0, 'RGN': 0, 'RMG0': 0.33, 'RPM0': 0.33, 'RPS': 0.11, 'RPS0': 0.2,
             'RPT': 0.12, 'RPT0': 0.2, 'RSM0': 0.33, 'RST': 0.13, 'RST0': 0.24, 'SECTAP': 132, 'TAGS': 't1;t2;', 'TERTAP': 13.8, 'XG1': 0.22,
             'XG2': 0, 'XG3': 0, 'XGN': 0, 'XMG0': 0.33, 'XPM0': 0.33, 'XPS': 0.32, 'XPS0': 0.26, 'XPT': 0.42, 'XPT0': 0.45, 'XSM0': 0.33,
             'XST': 0.33, 'XST0': 0.45, 'Z0METHOD': 1}
    keys = [['NEVADA', 132], ['NEW HAMPSHR', 33], ['ROANOKE', 13.8], '10']
    o1 = OlxObj.OLCase.addOBJ('XFMR3', keys, params)
    print('New XFMR3: ', o1.toString())


def sample_17_SHIFTER(olxpath, fi):
    """ (Phase Shifter) SHIFTER object demo

        # various ways to find a SHIFTER
        xp = OlxObj.OLCase.findOBJ('SHIFTER',"{0234d855-b935-47a5-ae05-682c243442e0}")             # GUID
        xp = OlxObj.OLCase.findOBJ("{0234d855-b935-47a5-ae05-682c243442e0}")                       # GUID
        xp = OlxObj.OLCase.findOBJ('SHIFTER',"[SHIFTER] 4 'TENNESSEE' 132 kV-6 'NEVADA' 132 kV 1") # STR
        xp = OlxObj.OLCase.findOBJ("[SHIFTER] 4 'TENNESSEE' 132 kV-6 'NEVADA' 132 kV 1")           # STR
        xp = OlxObj.OLCase.findOBJ('SHIFTER',[['TENNESSEE',132],['NEVADA',132],'1'])               # [b1,b2,CID]
        xp = OlxObj.OLCase.findOBJ('SHIFTER',[4,['NEVADA',132],'1'])                               # [b1,b2,CID]
        xp = OlxObj.OLCase.findOBJ('SHIFTER',["[BUS] 4 'TENNESSEE' 132 kV",6,'1'])                 # [b1,b2,CID]


        # SHIFTER can be accessed by related object
        xp = OlxObj.OLCase.SHIFTER[0]     # from NETWORK (OlxObj.OLCase)
        xp = b1.SHIFTER[0]                # from BUS (b1)
        xp = t1.EQUIPMENT                 # from TERMINAL (t1)
        xp = r1.EQUIPMENT                 # from RLYGROUP

        # SHIFTER Data : 'ANGMAX','ANGMIN','BASEMVA','BN1','BN2','BP1','BP2','BRCODE','BUS','BUS1','BUS2','BZ1','BZ2','CID',
            'CNTL','DATEOFF','DATEON','EQUIPMENT','FLAG','GN1','GN2','GP1','GP2','GUID','GZ1','GZ2','JRNL','KEYSTR','MEMO',
            'MVA1','MVA2','MVA3','MWMAX','MWMIN','NAME','PARAMSTR','PYTHONSTR','RLYGROUP','RLYGROUP1','RLYGROUP2','RN','RP',
            'RZ','SHIFTANGLE','TAGS','TERMINAL','XN','XP','XZ','ZCORRECTNO'


        print(xp.NAME,xp.XN,xp.XP)            # (str,float,float) SHIFTER Name,XN,XP
        print(xp.getData('NAME'))             # (str) SHIFTER Name
        print(xp.getData(['NAME','XN','XP'])) # (dict) SHIFTER Name,XN,XP
        print(xp.getData())                   # (dict) get all Data of SHIFTER


        # some methods
        xp.equals(xp2)                   # (bool) comparison of 2 objects
        xp.delete()                      # delete object
        xp.isInList(la)                  # check if object in in List/Set/tuple of object
        xp.toString()                    # (str) text description/composed of object
        xp.copyFrom(xp2)                 # copy Data from another object
        xp.changeData(sParam,value)      # change Data (sample_6_changeData())
        xp.postData()                    # Perform validation and update object data in the network database
        xp.getAttributes()               # [str] list attributes of object
        #
        xp.getData('xx')                 # help details in message error


        # add SHIFTER
        params = {'GUID':'{e115233c-20ab-43c5-a7c4-3bf29f0f890f}','ANGMAX':10.2,'ANGMIN':2.2,'BASEMVA':100,'BN1':0.13,\
            'BN2':-0.0052,'BP1':0.2,'BP2':-0.0051,'BZ1':0.16,'BZ2':-0.0053,'CNTL':0,'DATEOFF':'','DATEON':'',\
            'FLAG':1,'GN1':0.12,'GN2':0.14,'GP1':0.2,'GP2':0.12,'GZ1':0.15,'GZ2':0.17,'MEMO':b'fef','MVA1':5,'MVA2':6,\
            'MVA3':7,'MWMAX':5.2,'MWMIN':4,'NAME':'gyg','RN':0.014,'RP':0.013,'RZ':0.034,'SHIFTANGLE':5.1,\
            'TAGS':'t1;t2;','XN':0.046,'XP':0.047,'XZ':0.12,'ZCORRECTNO':1}
        keys = [['TENNESSEE', 132], ['NEVADA', 132], '5']
        o1 = OlxObj.OLCase.addOBJ('SHIFTER', keys, params)
    """

    print('\nsample_17_SHIFTER')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    xp = OlxObj.OLCase.findOBJ('SHIFTER', ["[BUS] 4 'TENNESSEE' 132 kV", 6, '1'])
    if xp!=None:
        print(xp.toString())
        print(xp.getData(['NAME', 'XN', 'XP']))

    # Add
    params = {'GUID': '{e115233c-20ab-43c5-a7c4-3bf29f0f890f}', 'ANGMAX': 10.2, 'ANGMIN': 2.2, 'BASEMVA': 100, 'BN1': 0.13,
             'BN2': -0.0052, 'BP1': 0.2, 'BP2': -0.0051, 'BZ1': 0.16, 'BZ2': -0.0053, 'CNTL': 0, 'DATEOFF': '', 'DATEON': '',
             'FLAG': 1, 'GN1': 0.12, 'GN2': 0.14, 'GP1': 0.2, 'GP2': 0.12, 'GZ1': 0.15, 'GZ2': 0.17, 'MEMO': b'fef', 'MVA1': 5, 'MVA2': 6,
             'MVA3': 7, 'MWMAX': 5.2, 'MWMIN': 4, 'NAME': 'gyg', 'RN': 0.014, 'RP': 0.013, 'RZ': 0.034, 'SHIFTANGLE': 5.1,
             'TAGS': 't1;t2;', 'XN': 0.046, 'XP': 0.047, 'XZ': 0.12, 'ZCORRECTNO': 1}
    keys = [['TENNESSEE', 132], ['NEVADA', 132], '5']
    o1 = OlxObj.OLCase.addOBJ('SHIFTER', keys, params)
    print('New SHIFTER: ', o1.toString())


def sample_18_SERIESRC(olxpath, fi):
    """ (Series capacitor/reactor) SERIESRC object demo

        # various ways to find a SERIESRC
        xc = OlxObj.OLCase.findOBJ('SERIESRC',"{fc28c824-4415-4869-ab6a-ac7527d484eb}")           # GUID
        xc = OlxObj.OLCase.findOBJ("{fc28c824-4415-4869-ab6a-ac7527d484eb}")                      # GUID
        xc = OlxObj.OLCase.findOBJ('SERIESRC',"[SERIESRC] 5 'FIELDALE' 132 kV-7 'OHIO' 132 kV 2") # STR
        xc = OlxObj.OLCase.findOBJ("[SERIESRC] 5 'FIELDALE' 132 kV-7 'OHIO' 132 kV 2")            # STR
        xc = OlxObj.OLCase.findOBJ('SERIESRC',[['FIELDALE',132],['OHIO',132],'2'])                # [b1,b2,CID]
        xc = OlxObj.OLCase.findOBJ('SERIESRC',[5,['OHIO',132],'2'])                               # [b1,b2,CID]
        xc = OlxObj.OLCase.findOBJ('SERIESRC',["[BUS] 5 'FIELDALE' 132 kV",7,'2'])                # [b1,b2,CID]


        # SERIESRC can be accessed by related object-
        xc = OlxObj.OLCase.SERIESRC[0]     # from NETWORK (OlxObj.OLCase)
        xc = b1.SERIESRC[0]                # from BUS (b1)
        xc = t1.EQUIPMENT                  # from TERMINAL (t1)
        xc = r1.EQUIPMENT                  # from RLYGROUP

        # SERIESRC Data : 'BRCODE','BUS','BUS1','BUS2','CID','DATEOFF','DATEON','EQUIPMENT','FLAG','GUID','IPR','JRNL','KEYSTR',
                'MEMO','NAME','PARAMSTR','PYTHONSTR','R','RLYGROUP','RLYGROUP1','RLYGROUP2','SCOMP','TAGS','TERMINAL','X'
        #
        print(xc.NAME,xc.R,xc.X)             # (str,float,float) SERIESRC Name,R,X
        print(xc.getData('NAME'))            # (str) SERIESRC Name
        print(xc.getData(['NAME','R','X']))  # (dict) SERIESRC Name,R,X
        print(xc.getData())                  # (dict) get all Data of SERIESRC


        # some methods:
        xc.equals(xc2)                   # (bool) comparison of 2 objects
        xc.delete()                      # delete object
        xc.isInList(la)                  # check if object in in List/Set/tuple of object
        xc.toString()                    # (str) text description/composed of object
        xc.copyFrom(xc2)                 # copy Data from another object
        xc.changeData(sParam,value)      # change Data (sample_6_changeData())
        xc.postData()                    # Perform validation and update object data in the network database
        xc.getAttributes()               # [str] list attributes of object
        #
        xc.getData('xxx')                # help details in message error


        # Add SERIESRC
        params = {'GUID':'{12607a9f-2ea9-4795-8e81-a7013b7f1d8e}','DATEOFF':'2016/10/31','DATEON':'2015/10/30','FLAG':1,
                'IPR':0,'MEMO':b'rg','NAME':'se1','R':1e-05,'SCOMP':1,'TAGS':'tt','X':0.0002}
        keys = [['FIELDALE', 132], ['OHIO', 132], '10']
        o1 = OlxObj.OLCase.addOBJ('SERIESRC', keys, params)
    """

    print('\nsample_18_SERIESRC')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only


    xc = OlxObj.OLCase.findOBJ('SERIESRC', ["[BUS] 5 'FIELDALE' 132 kV", 7, '2'])
    if xc!=None:
        print(xc.toString())
        print(xc.getData(['NAME', 'R', 'X']))


    # Add SERIESRC
    params = {'GUID': '{12607a9f-2ea9-4795-8e81-a7013b7f1d8e}', 'DATEOFF': '2016/10/31', 'DATEON': '2015/10/30', 'FLAG': 1,
             'IPR': 0, 'MEMO': b'rg', 'NAME': 'se1', 'R': 1e-05, 'SCOMP': 1, 'TAGS': 'tt', 'X': 0.0002}
    keys = [['FIELDALE', 132], ['OHIO', 132], '10']
    o1 = OlxObj.OLCase.addOBJ('SERIESRC', keys, params)
    print('New SERIESRC: ', o1.toString())


def sample_19_SWITCH(olxpath, fi):
    """ (Switch) SWITCH object demo
        # various ways to find a SWITCH
        xc = OlxObj.OLCase.findOBJ('SWITCH',"{a81a1a4b-8ace-4da2-be5c-c00231dfd8d8}")          # GUID
        xc = OlxObj.OLCase.findOBJ("{a81a1a4b-8ace-4da2-be5c-c00231dfd8d8}")                   # GUID
        xc = OlxObj.OLCase.findOBJ('SWITCH',"[SWITCH] 5 'FIELDALE' 132 kV-7 'OHIO' 132 kV 1")  # STR
        xc = OlxObj.OLCase.findOBJ("[SWITCH] 5 'FIELDALE' 132 kV-7 'OHIO' 132 kV 1")           # STR
        xc = OlxObj.OLCase.findOBJ('SWITCH',[['FIELDALE',132],['OHIO',132],'1'])               # [b1,b2,CID]
        xc = OlxObj.OLCase.findOBJ('SWITCH',[5,['OHIO',132],'1'])                              # [b1,b2,CID]
        xc = OlxObj.OLCase.findOBJ('SWITCH',["[BUS] 5 'FIELDALE' 132 kV",7,'1'])               # [b1,b2,CID]


        # SWITCH can be accessed by related object
        xc = OlxObj.OLCase.SWITCH[0]     # from NETWORK (OlxObj.OLCase)
        xc = b1.SWITCH[0]                # from BUS (b1)
        xc = t1.EQUIPMENT                # from TERMINAL (t1)
        xc = r1.EQUIPMENT                # from RLYGROUP

        # SWITCH Data : 'BRCODE','BUS','BUS1','BUS2','CID','DATEOFF','DATEON','DEFAULT','EQUIPMENT','FLAG','GUID','JRNL',
                'KEYSTR','MEMO','NAME','PARAMSTR','PYTHONSTR','RATING','RLYGROUP','RLYGROUP1','RLYGROUP2','STAT','TAGS','TERMINAL'
        #
        print(xc.NAME,xc.CID,xc.FLAG)             # (str,str,int) SWITCH Name,CID,FLAG
        print(xc.getData('NAME'))                 # (str) SWITCH Name
        print(xc.getData(['NAME','CID','FLAG']))  # (dict) SWITCH Name,CID,FLAG
        print(xc.getData())                       # (dict) get all Data of SWITCH


        # some methods
        xc.equals(xc2)                   # (bool) comparison of 2 objects
        xc.delete()                      # delete object
        xc.isInList(la)                  # check if object in in List/Set/tuple of object
        xc.toString()                    # (str) text description/composed of object
        xc.copyFrom(xc2)                 # copy Data from another object
        xc.changeData(sParam,value)      # change Data (sample_6_changeData())
        xc.postData()                    # Perform validation and update object data in the network database
        xc.getAttributes()               # [str] list attributes of object
        #
        xc.getData('XXX')                # help details in message error


        # Add SWITCH
        params = {'GUID':'{ad7edd55-ff9d-47df-b178-fda6cfdc9898}','DATEOFF':'2016/10/30','DATEON':'2015/10/30','DEFAULT':1,
                'FLAG':1,'MEMO':b'ht','NAME':'sw2','RATING':105,'STAT':0,'TAGS':'t1;t2;'}
        keys = [['FIELDALE', 132], ['OHIO', 132], '11']
        o1 = OlxObj.OLCase.addOBJ('SWITCH', keys, params)
    """

    print('\nsample_19_SWITCH')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    xc = OlxObj.OLCase.findOBJ('SWITCH', ["[BUS] 5 'FIELDALE' 132 kV", 7, '1'])
    if xc!=None:
        print(xc.toString())
        print(xc.getData(['NAME', 'CID', 'FLAG']))

    # Add
    params = {'GUID': '{ad7edd55-ff9d-47df-b178-fda6cfdc9898}', 'DATEOFF': '2016/10/30', 'DATEON': '2015/10/30', 'DEFAULT': 1,
             'FLAG': 1, 'MEMO': b'ht', 'NAME': 'sw2', 'RATING': 105, 'STAT': 0, 'TAGS': 't1;t2;'}
    keys = [['FIELDALE', 132], ['OHIO', 132], '11']
    o1 = OlxObj.OLCase.addOBJ('SWITCH', keys, params)
    print('New SWITCH: ', o1.toString())


def sample_20_BREAKER(olxpath,fi):
    """ (Breaker Rating) BREAKER object demo

        # various ways to find a BREAKER
        bk = OlxObj.OLCase.findOBJ('BREAKER',"{9dfa81a8-deb5-4118-9b04-2e23b458d255}")  # GUID
        bk = OlxObj.OLCase.findOBJ("{9dfa81a8-deb5-4118-9b04-2e23b458d255}")            # GUID
        bk = OlxObj.OLCase.findOBJ('BREAKER',"[BREAKER]  1E82A@ 6 'NEVADA' 132 kV")     # STR
        bk = OlxObj.OLCase.findOBJ("[BREAKER]  1E82A@ 6 'NEVADA' 132 kV")               # STR
        bk = OlxObj.OLCase.findOBJ('BREAKER',['NEVADA',132,'1E82A'])                    # [bus,Name]
        bk = OlxObj.OLCase.findOBJ('BREAKER',[6,'1E82A'])                               # [bus,Name]


        # BREAKER can be accessed by related object
        bk = OlxObj.OLCase.BREAKER[0]     # from NETWORK (OlxObj.OLCase)
        bk = b1.BREAKER[0]                # from BUS (b1)


        # BREAKER Data : 'BUS','CPT1','CPT2','FLAG','G1OUTAGES','G2OUTAGES','GROUPTYPE1','GROUPTYPE2','GUID','INTRATING',
                'INTRTIME','JRNL','K','KEYSTR','MEMO','MRATING','NACD','NAME','NODERATE','OBJLST1','OBJLST2','OPKV',
                'OPS1','OPS2','PARAMSTR','PYTHONSTR','RATEDKV','RATINGTYPE','RECLOSE1','RECLOSE2','TAGS'
        #
        print(bk.NAME,bk.OPKV,bk.RATEDKV)             # (str,float,float) BREAKER Name,OPKV,RATEDKV
        print(bk.getData('NAME'))                     # (str) BREAKER Name
        print(bk.getData(['NAME','OPKV','RATEDKV']))  # (dict) BREAKER Name,OPKV,RATEDKV
        print(bk.getData())                           # (dict) get all Data of BREAKER


        # some methods
        bk.equals(bk2)                   # (bool) comparison of 2 objects
        bk.delete()                      # delete object
        bk.isInList(la)                  # check if object in in List/Set/tuple of object
        bk.toString()                    # (str) text description/composed of object
        bk.copyFrom(bk2)                 # copy Data from another object
        bk.getAttributes()               # [str] list attributes of object
        bk.changeData(sParam,value)      # change Data (sample_6_changeData())
        bk.postData()                    # Perform validation and update object data in the network database
        #
        bk.getData('XXX')                # help details in message error


        # Add BREAKER
        params = {'GUID':'{1dfa81a8-deb5-4118-9b04-2e23b458d255}','CPT1':4,'CPT2':4,'FLAG':1,'G1OUTAGES':[],'G2OUTAGES':[],
            'GROUPTYPE1':0,'GROUPTYPE2':0,'INTRATING':3000,'INTRTIME':5,'K':1.2,'MEMO':b'mm','MRATING':20000,'NACD':1,
            'NODERATE':0,'OBJLST1':["[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1"],'OBJLST2':["[LINE] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1",
            "[LINE] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1"],'OPKV':130,'OPS1':3,'OPS2':3,'RATEDKV':139,'RATINGTYPE':1,'RECLOSE1':[8,30,0],'RECLOSE2':[8,30,0],'TAGS':'tt'}
        keys = ['NEVADA', 132, 'newBK']
        o1 = OlxObj.OLCase.addOBJ('BREAKER', keys, params)
    """

    print('\nsample_20_BREAKER')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    bk = OlxObj.OLCase.findOBJ('BREAKER', ['NEVADA', 132, '1E82A'])
    if bk!=None:
        print(bk.toString())
        print(bk.getData(['NAME', 'OPKV', 'RATEDKV']))

    # Add
    params = {'GUID': '{1dfa81a8-deb5-4118-9b04-2e23b458d255}', 'CPT1': 4, 'CPT2': 4, 'FLAG': 1, 'G1OUTAGES': [],
             'G2OUTAGES': [],'GROUPTYPE1': 0, 'GROUPTYPE2': 0, 'INTRATING': 3000, 'INTRTIME': 5, 'K': 1.2,
             'MEMO': b'mm', 'MRATING': 20000, 'NACD': 1,'NODERATE': 0,
             'OBJLST1': ["[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1"],
             'OBJLST2': ["[LINE] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1","[LINE] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1"],
             'OPKV': 130, 'OPS1': 3, 'OPS2': 3, 'RATEDKV': 139, 'RATINGTYPE': 1, 'RECLOSE1': [8, 30, 0], 'RECLOSE2': [8, 30, 0], 'TAGS': 'tt'}
    keys = [['NEVADA', 132], 'newBrk']
    o1 = OlxObj.OLCase.addOBJ('BREAKER', keys, params)
    print('New BREAKER: ', o1.toString())
    print(o1.pythonstr)


def sample_21_MULINE(olxpath,fi):
    """ (Mutual Pair) MULINE object demo
        # various ways to find a MULINE ----------------------------------------------------------------------------
        m1 = OlxObj.OLCase.findOBJ('MULINE',"{f1aba445-7598-4c1c-82a5-b908f34b5692}")  # GUID
        m1 = OlxObj.OLCase.findOBJ("{f1aba445-7598-4c1c-82a5-b908f34b5692}")  # GUID
        m1 = OlxObj.OLCase.findOBJ('MULINE',"[MUPAIR] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1|8 'REUSENS' 132 kV-28 'ARIZONA' 132 kV 1")  # STR
        m1 = OlxObj.OLCase.findOBJ("[MUPAIR] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1|8 'REUSENS' 132 kV-28 'ARIZONA' 132 kV 1")           # STR
        m1 = OlxObj.OLCase.findOBJ('MULINE',["[LINE] 2 'CLAYTOR' 132 kV-5 'FIELDALE' 132 kV 1","[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1"]) # [l1,l2]
        m1 = OlxObj.OLCase.findOBJ('MULINE',[[2,5,'1'],[2,6,'1']])  # [l1,l2]
        #
        m1 = OlxObj.OLCase.findMULINE("{f1aba445-7598-4c1c-82a5-b908f34b5692}")        # GUID
        m1 = OlxObj.OLCase.findMULINE("[MUPAIR] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1|8 'REUSENS' 132 kV-28 'ARIZONA' 132 kV 1")         # STR
        m1 = OlxObj.OLCase.findMULINE(["[LINE] 2 'CLAYTOR' 132 kV-5 'FIELDALE' 132 kV 1","[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1"]) # [l1,l2]
        m1 = OlxObj.OLCase.findMULINE([[2,5,'1'],[2,6,'1']])  # [l1,l2]

        # MULINE can be accessed by related object ------------------------------------------------------------------
        m1 = OlxObj.OLCase.MULINE[0]     # from NETWORK (OlxObj.OLCase)
        m1 = l1.MULINE[0]                # from LINE (l1)

        # MULINE Data : 'FROM1','FROM2','GUID','JRNL','KEYSTR','LINE1','LINE2','MEMO','PARAMSTR','PYTHONSTR','R','TAGS','TO1','TO2','X'
        #
        print(m1.R,m1.X)              # (float,float) MULINE R,X
        print(m1.getData('R'))        # ([]) MULINE R
        print(m1.getData(['R','X']))  # (dict) MULINE R,X
        print(m1.getData())           # (dict) get all Data of MULINE

        # some methods--------------------------------------------------------------------------------------------
        m1.equals(m2)                 # (bool) comparison of 2 objects
        m1.delete()                   # delete object
        m1.isInList(la)               # check if object in in List/Set/tuple of object
        m1.postData()                 # Perform validation and update object data in the network database
        m1.toString()                 # (str) text description/composed of object
        m1.copyFrom(m2)               # copy Data from another object
        m1.changeData(sParam,value)   # change Data (sample_6_changeData())
        m1.getAttributes()            # [str] list attributes of object
        #
        m1.getData('XXX')             # help details in message error

        # Add MULINE -----------------------------------------------------------------------------------------------
        params = {'GUID': '{1f8ba9f6-dcc5-4556-b3df-7e99991c4a91}', 'FROM1': [0, 0, 0, 0, 0], 'FROM2': [0, 0, 0, 0, 0], 'MEMO': b'mm',
             'R': [-0.01, 0, 0, 0, 0], 'TAGS': 'tt', 'TO1': [100, 0, 0, 0, 0], 'TO2': [100, 0, 0, 0, 0], 'X': [-0.02, 0, 0, 0, 0]}
        keys = [[['CLAYTOR', 132], ['TENNESSEE', 132], '1'], [ ['FIELDALE', 132], ['CLAYTOR', 132],'1']]
        o1 = OlxObj.OLCase.addOBJ('MULINE', keys, params)
    """

    print('\nsample_21_MULINE')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findMULINE([[2, 5, '1'], [2, 6, '1']])
    print(o1.getData(['R', 'X']))
    print(o1.toString())
    params = {'ORIENTLINE1': [2, 5, '1'],
            'ORIENTLINE2': [6, 2, '1'],
        'R': [-0.0111, 0, 0, 0, 0],
        'X': [-0.0222, 0, 0, 0, 0]}
    o1.changeData(params)
    o1.postData()
    print('newRX: ', o1.getData(['R', 'X']))

    # Add
    params = {'GUID': '{1f8ba9f6-dcc5-4556-b3df-7e99991c4a91}', 'FROM1': [0, 0, 0, 0, 0], 'FROM2': [0, 0, 0, 0, 0], 'MEMO': b'mm',
             'R': [0.01, 0, 0, 0, 0], 'TAGS': 'tt', 'TO1': [100, 0, 0, 0, 0], 'TO2': [100, 0, 0, 0, 0], 'X': [0.02, 0, 0, 0, 0]}
    keys = [[['CLAYTOR', 132], ['TENNESSEE', 132], '1'], [['CLAYTOR', 132], ['FIELDALE', 132], '1']]
    o1 = OlxObj.OLCase.addOBJ('MULINE', keys, params)
    print('New MULINE: ', o1.toString())


def sample_22_GEN(olxpath, fi):
    """ (Generator) GEN object demo

        # various ways to find a GEN
        g1 = OlxObj.OLCase.findOBJ('GEN',"{53363af8-293c-4565-9d36-0d598f619075}") # GUID
        g1 = OlxObj.OLCase.findOBJ("{53363af8-293c-4565-9d36-0d598f619075}")       # GUID
        g1 = OlxObj.OLCase.findOBJ('GEN',"[GENERATOR] 5 'FIELDALE' 132 kV")        # STR
        g1 = OlxObj.OLCase.findOBJ("[GENERATOR] 5 'FIELDALE' 132 kV")              # STR
        g1 = OlxObj.OLCase.findOBJ('GEN',['FIELDALE',132])                         # BUS = [name,kv] (or STR or GUID or Bus Number )
        g1 = OlxObj.OLCase.findOBJ('GEN',5)                                        # BUS = Bus Number
        g1 = OlxObj.OLCase.findOBJ('GEN',"[BUS] 5 'FIELDALE' 132 kV")              # BUS = Bus STR


        # GEN can be accessed by related object
        g1 = OlxObj.OLCase.GEN[0]     # from NETWORK (OlxObj.OLCase)
        g1 = b1.GEN                   # from BUS (b1)
        g1 = gu1.GEN                  # from GENUNIT (gu1)

        # GEN Data :'BUS','CNTBUS','FLAG','GENUNIT','GUID','ILIMIT1','ILIMIT2','JRNL','KEYSTR','MEMO','PARAMSTR','PGEN',
                    'PYTHONSTR','QGEN','REFANGLE','REFV','REG','SCHEDP','SCHEDQ','SCHEDV','TAGS'
        #
        print(g1.PGEN,g1.QGEN)               # (float,float) GEN PGEN,QGEN
        print(g1.getData('PGEN'))            # (str) GEN Name
        print(g1.getData(['PGEN','QGEN']))   # (dict) GEN Name,R,R0
        print(g1.getData())                  # (dict) get all Data of GEN

        # some methods--------------------------------------------------------------------------------------------
        g1.equals(g2)                    # (bool) comparison of 2 objects
        g1.delete()                      # delete object
        g1.isInList(la)                  # check if object in in List/Set/tuple of object
        g1.toString()                    # (str) text description/composed of object
        g1.copyFrom(g2)                  # copy Data from another object
        g1.getAttributes()               # [str] list attributes of object
        g1.changeData(sParam,value)      # change Data (sample_6_changeData())
        g1.postData()                    # Perform validation and update object data in the network database
        #
        g1.getData('XXX')                # help details in message error

        # Add GEN -----------------------------------------------------------------------------------------------
        params = {'GUID':'{13363af8-293c-4565-9d36-0d598f619071}','CNTBUS':"[BUS] 5 'FIELDALE' 132 kV",'FLAG':1,'ILIMIT1':0,
               'ILIMIT2':0,'MEMO':b'mm','REFANGLE':0,'REFV':1,'REG':0,'SCHEDV':1,'TAGS':'tt'}
        keys = ['OHIO', 132]
        o1 = OlxObj.OLCase.addOBJ('GEN', keys, params)
    """

    print('\nsample_22_GEN')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    g1 = OlxObj.OLCase.findOBJ('GEN', ['FIELDALE', 132])
    if g1!=None:
        print(g1.toString())
        print(g1.getData(['PGEN','QGEN']))
        print(g1.PGEN, g1.QGEN)

    # Add
    params = {'GUID': '{13363af8-293c-4565-9d36-0d598f619071}', 'CNTBUS': "[BUS] 5 'FIELDALE' 132 kV", 'FLAG': 1, 'ILIMIT1': 10,
             'ILIMIT2': 0, 'MEMO': b'mm', 'REFANGLE': 0, 'REFV': 1, 'REG': 0, 'SCHEDV': 1, 'TAGS': 'tt'}
    keys = ['OHIO', '132']
    o1 = OlxObj.OLCase.addOBJ('GEN', keys, params)
    print('New GEN: ', o1.toString())


def sample_23_GENUNIT(olxpath, fi):
    """ (Generator unit) GENUNIT object demo

        # various ways to find a GENUNIT
        gu1 = OlxObj.OLCase.findOBJ('GENUNIT',"{46f3688d-5131-47ce-9dc0-a5a83f96bf77}") # GUID
        gu1 = OlxObj.OLCase.findOBJ("{46f3688d-5131-47ce-9dc0-a5a83f96bf77}")           # GUID
        gu1 = OlxObj.OLCase.findOBJ('GENUNIT',"[GENUNIT]  1@5 'FIELDALE' 132 kV")       # STR
        gu1 = OlxObj.OLCase.findOBJ("[GENUNIT]  1@5 'FIELDALE' 132 kV")                 # STR
        gu1 = OlxObj.OLCase.findOBJ('GENUNIT',['FIELDALE',132,'1'])                     # [BUS,CID] BUS = [name,kv] (or STR or GUID or Bus Number )
        gu1 = OlxObj.OLCase.findOBJ('GENUNIT',[5,'1'])                                  # [BUS,CID] BUS = Bus Number
        gu1 = OlxObj.OLCase.findOBJ('GENUNIT',["[BUS] 5 'FIELDALE' 132 kV",'1'])        # [BUS,CID] BUS = Bus STR


        # GENUNIT can be accessed by related object
        gu1 = OlxObj.OLCase.GENUNIT[0]     # from NETWORK (OlxObj.OLCase)
        gu1 = b1.GENUNIT[0]                # from BUS (b1)
        gu1 = g1.GENUNIT[0]                # from GEN (g1)

        # GENUNIT Data :'BUS','CID','DATEOFF','DATEON','FLAG','GEN','GUID','JRNL','KEYSTR','MVARATE','PARAMSTR',
                        'PMAX','PMIN','PYTHONSTR','QMAX','QMIN','R','RG','SCHEDP','SCHEDQ','TAGS','X','XG'


        print(gu1.R,gu1.X)              # (float,float) GENUNIT R,X
        print(gu1.getData('R'))         # ([]) GENUNIT R
        print(gu1.getData(['R','X']))   # (dict) GENUNIT R,X
        print(gu1.getData())            # (dict) get all Data of GENUNIT

        # some methods
        gu1.equals(gu2)                   # (bool) comparison of 2 objects
        gu1.delete()                      # delete object
        gu1.isInList(la)                  # check if object in in List/Set/tuple of object
        gu1.toString()                    # (str) text description/composed of object
        gu1.copyFrom(gu2)                 # copy Data from another object
        gu1.getAttributes()               # [str] list attributes of object
        gu1.changeData(sParam,value)      # change Data (sample_6_changeData())
        gu1.postData()                    # Perform validation and update object data in the network database
        #
        gu1.getData('XXX')                # help details in message error

        # Add GENUNIT
        params = {'GUID':'{16f3688d-5131-47ce-9dc0-a5a83f96bf71}','DATEOFF':'N/A','DATEON':'N/A','FLAG':1,'MVARATE':100,
            'PMAX':9999,'PMIN':-9999,'QMAX':9999,'QMIN':-9999,'R':[0,0,0,0,0],'RG':0,'SCHEDP':0,'SCHEDQ':0,'TAGS':'tt',
            'X':[0.1,0.1,0.1,0.1,0.1],'XG':0}
        keys = ['FIELDALE', 132, '12']
        o1 = OlxObj.OLCase.addOBJ('GENUNIT', keys, params)
    """

    print('\nsample_23_GENUNIT')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    gu1 = OlxObj.OLCase.findOBJ('GENUNIT', ['FIELDALE', '132', 1])
    if gu1!=None:
        print(gu1.toString())
        print(gu1.R, gu1.X)
        print(gu1.getData(['R','X']))

    #
    params = {'GUID': '{16f3688d-5131-47ce-9dc0-a5a83f96bf71}', 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'FLAG': 1, 'MVARATE': 100,
             'PMAX': 9999, 'PMIN': '-9999', 'QMAX': 9999, 'QMIN': -9999, 'R': [0, 0, 0, 0, 0], 'RG': 0,
             'SCHEDP': 0, 'SCHEDQ': 0, 'TAGS': 'tt','X': [0.1, 0.1, 0.1, 0.1, 0.1], 'XG': 0}
    keys = [['FIELDALE', 132], '12']
    o1 = OlxObj.OLCase.addOBJ('GENUNIT', keys, params)
    print('New GENUNIT: ', o1.toString())


def sample_24_GENW3(olxpath, fi):
    """ (Type-3 Wind Plants) GENW3 object demo

        # various ways to find a GENW3
        g3 = OlxObj.OLCase.findOBJ('GENW3',"{62be5ee7-5d4a-455c-b148-67144d0fc205}") # GUID
        g3 = OlxObj.OLCase.findOBJ("{62be5ee7-5d4a-455c-b148-67144d0fc205}")         # GUID
        g3 = OlxObj.OLCase.findOBJ('GENW3',"[GENW3] 10 'NEW HAMPSHR' 33 kV")         # STR
        g3 = OlxObj.OLCase.findOBJ("[GENW3] 10 'NEW HAMPSHR' 33 kV")                 # STR
        g3 = OlxObj.OLCase.findOBJ('GENW3',['NEW HAMPSHR',33])                       # BUS = [name,kv] (or STR or GUID or Bus Number )
        g3 = OlxObj.OLCase.findOBJ('GENW3',10)                                       # BUS = Bus Number
        g3 = OlxObj.OLCase.findOBJ('GENW3',"[BUS] 10 'NEW HAMPSHR' 33 kV")           # BUS = Bus STR


        # GENW3 can be accessed by related object
        g3 = OlxObj.OLCase.GENW3[0]     # from NETWORK (OlxObj.OLCase)
        g3 = b1.GENW3                   # from BUS (b1)

        # GENW3 Data :'BUS','CBAR','DATEOFF','DATEON','FLAG','FLRZ','GUID','IMAXG','IMAXR','JRNL','KEYSTR','LLR','LLS','LM','MEMO',
                      'MVA','MW','MWR','PARAMSTR','PYTHONSTR','RR','RS','SLIP','TAGS','UNITS','VMAX','VMIN'
        #
        print(g3.MVA,g3.MW)               # (float,float) GENW3 MVA,MW
        print(g3.getData('MVA'))          # (float) GENW3 MVA
        print(g3.getData(['MVA','MW']))   # (dict) GENW3 MVA,MW
        print(g3.getData())               # (dict) get all Data of GENW3

        # some methods
        g3.equals(g32)                   # (bool) comparison of 2 objects
        g3.delete()                      # delete object
        g3.isInList(la)                  # check if object in in List/Set/tuple of object
        g3.toString()                    # (str) text description/composed of object
        g3.copyFrom(g32)                 # copy Data from another object
        g3.getAttributes()               # [str] list attributes of object
        g3.changeData(sParam,value)      # change Data (sample_6_changeData())
        g3.postData()                    # Perform validation and update object data in the network database
        #
        g3.getData('XXX')                # help details in message error

        # Add GENW3
        params = {'GUID':'{12be5ee7-5d4a-455c-b148-67144d0fc201}','CBAR':0,'DATEOFF':'N/A','DATEON':'N/A','FLAG':1,
            'FLRZ':-20,'IMAXG':0.35,'IMAXR':1.1,'LLR':0.2,'LLS':0.2,'LM':3,'MEMO':b'mm','MVA':5,'MW':4,'MWR':4,
            'RR':0.03,'RS':0.03,'SLIP':-0.2,'TAGS':'tt','UNITS':1,'VMAX':1.2,'VMIN':0.01}
        keys = ['TENNESSEE', 132]
        o1 = OlxObj.OLCase.addOBJ('GENW3', keys, params)
    """

    print('\nsample_24_GENW3')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    g3 = OlxObj.OLCase.findGENW3("[BUS] 10 'NEW HAMPSHR' 33 kV")
    if g3!=None:
        print(g3.toString())
        print(g3.MVA, g3.MW)
        print(g3.getData(['MVA', 'MW']))

    # Add
    params = {'GUID': '{12be5ee7-5d4a-455c-b148-67144d0fc201}', 'CBAR': 0, 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'FLAG': 1,
             'FLRZ': -20, 'IMAXG': 0.35, 'IMAXR': 1.1, 'LLR': 0.2, 'LLS': 0.2, 'LM': 3, 'MEMO': b'mm', 'MVA': 5, 'MW': 4, 'MWR': 4,
             'RR': 0.03, 'RS': 0.03, 'SLIP': -0.2, 'TAGS': 'tt', 'UNITS': 1, 'VMAX': 1.2, 'VMIN': 0.01}
    keys = ['TENNESSEE', '132']
    o1 = OlxObj.OLCase.addOBJ('GENW3', keys, params)
    print('New GENW3: ', o1.paramstr)


def sample_25_GENW4(olxpath, fi):
    """ (Converter-Interfaced Resources) GENW4 object demo

        # various ways to find a GENW4
        g4 = OlxObj.OLCase.findOBJ('GENW4',"{df7317c9-3b83-4d59-9aa4-5a55a7f20746}") # GUID
        g4 = OlxObj.OLCase.findOBJ("{df7317c9-3b83-4d59-9aa4-5a55a7f20746}")         # GUID
        g4 = OlxObj.OLCase.findOBJ('GENW4',"[GENW4] 28 'ARIZONA' 132 kV")            # STR
        g4 = OlxObj.OLCase.findOBJ("[GENW4] 28 'ARIZONA' 132 kV")                    # STR
        g4 = OlxObj.OLCase.findOBJ('GENW4',['ARIZONA',132])                          # BUS = [name,kv] (or STR or GUID or Bus Number )
        g4 = OlxObj.OLCase.findOBJ('GENW4',28)                                       # BUS = Bus Number
        g4 = OlxObj.OLCase.findOBJ('GENW4',"[BUS] 28 'ARIZONA' 132 kV")              # BUS = Bus STR


        # GENW4 can be accessed by related object
        g4 = OlxObj.OLCase.GENW4[0]     # from NETWORK (OlxObj.OLCase)
        g4 = b1.GENW4                   # from BUS (b1)

        # GENW4 Data :'BUS','CTRLMETHOD','DATEOFF','DATEON','FLAG','GUID','JRNL','KEYSTR','MAXI','MAXILOW','MEMO','MVA','MVAR',
                    'MW','PARAMSTR','PYTHONSTR','SLOPE','SLOPENEG','TAGS','UNITS','VLOW','VMAX','VMIN'
        #
        print(g4.MVA,g4.MW)               # (float,float) GENW4 MVA,MW
        print(g4.getData('MVA'))          # (float) GENW4 MVA
        print(g4.getData(['MVA','MW']))   # (dict) GENW4 MVA,MW
        print(g4.getData())               # (dict) get all Data of GENW4

        # some methods
        g4.equals(g42)                   # (bool) comparison of 2 objects
        g4.delete()                      # delete object
        g4.isInList(la)                  # check if object in in List/Set/tuple of object
        g4.postData()                    # Perform validation and update object data in the network database
        g4.toString()                    # (str) text description/composed of object
        g4.copyFrom(g42)                 # copy Data from another object
        g4.changeData(sParam,value)      # change Data (sample_6_changeData())
        g4.getAttributes()               # [str] list attributes of object
        #
        g4.getData('XXX')                # help details in message error

        # Add GENW4
        params = {'GUID':'{df7317c1-3b83-4d59-9aa4-5a55a7f20741}','CTRLMETHOD':0,'DATEOFF':'N/A','DATEON':'N/A',
                'FLAG':1,'MAXI':1.1,'MAXILOW':1.1,'MEMO':b'mm','MVA':5,'MVAR':0,'MW':4,'SLOPE':2,'SLOPENEG':0,
                'TAGS':'tt','UNITS':1,'VLOW':0.5,'VMAX':1.2,'VMIN':0}
        keys = ['NEVADA', 132]
        o1 = OlxObj.OLCase.addOBJ('GENW4', keys, params)
    """

    print('\nsample_25_GENW4')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    g4 = OlxObj.OLCase.findOBJ('GENW4', "[BUS] 28 'ARIZONA' 132 kV")
    if g4!=None:
        print(g4.toString())
        print(g4.getData(['MVA', 'MW']))

    # Add
    params = {'GUID': '{df7317c1-3b83-4d59-9aa4-5a55a7f20741}', 'CTRLMETHOD': 0, 'DATEOFF': 'N/A', 'DATEON': 'N/A',
             'FLAG': 1, 'MAXI': 1.1, 'MAXILOW': 1.1, 'MEMO': b'mm', 'MVA': 5, 'MVAR': 0, 'MW': 4, 'SLOPE': 2,
             'SLOPENEG': 0, 'TAGS': 'tt', 'UNITS': 1, 'VLOW': 0.5, 'VMAX': 1.2, 'VMIN': 0}
    keys = ['NEVADA', 132]
    o1 = OlxObj.OLCase.addOBJ('GENW4', keys, params)
    print('New GENW4: ', o1.paramstr)


def sample_26_CCGEN(olxpath, fi):
    """ (Converter-Interfaced Resources) CCGEN object demo

        # various ways to find a CCGEN
        gc = OlxObj.OLCase.findOBJ('CCGEN',"{47ac9db2-ca74-4316-bb74-86a1e8aed0c8}") # GUID
        gc = OlxObj.OLCase.findOBJ("{47ac9db2-ca74-4316-bb74-86a1e8aed0c8}")         # GUID
        gc = OlxObj.OLCase.findOBJ('CCGEN',"[CCGENUNIT] 12 'TEXAS' 132 kV")          # STR
        gc = OlxObj.OLCase.findOBJ("[CCGENUNIT] 12 'TEXAS' 132 kV")                  # STR
        gc = OlxObj.OLCase.findOBJ('CCGEN',['TEXAS',132])                            # BUS = [name,kv] (or STR or GUID or Bus Number )
        gc = OlxObj.OLCase.findOBJ('CCGEN',12)                                       # BUS = Bus Number
        gc = OlxObj.OLCase.findOBJ('CCGEN',"[BUS] 12 'TEXAS' 132 kV")                # BUS = Bus STR

        # CCGEN can be accessed by related object
        gc = OlxObj.OLCase.CCGEN[0]     # from NETWORK (OlxObj.OLCase)
        gc = b1.CCGEN                   # from BUS (b1)

        # CCGEN Data : 'A','BLOCKPHASE','BUS','DATEOFF','DATEON','FLAG','GUID','I','JRNL','KEYSTR','MEMO','MVARATE',
                       'PARAMSTR','PYTHONSTR','TAGS','V','VLOC','VMAXMUL','VMIN'
        #
        print(gc.I,gc.V,gc.A)             # ([],[],[]) CCGEN I,V,A
        print(gc.getData('I'))            # ([]) CCGEN I
        print(gc.getData(['V','I','A']))  # (dict) CCGEN I,V,A
        print(gc.getData())               # (dict) get all Data of CCGEN

        # some methods
        gc.equals(gc2)                   # (bool) comparison of 2 objects
        gc.delete()                      # delete object
        gc.isInList(la)                  # check if object in in List/Set/tuple of object
        gc.toString()                    # (str) text description/composed of object
        gc.copyFrom(gc2)                 # copy Data from another object
        gc.getAttributes()               # [str] list attributes of object
        gc.changeData(sParam,value)      # change Data (sample_6_changeData())
        gc.postData()                    # Perform validation and update object data in the network database

        #
        gc.getData('XXX')                # help details in message error

        # Add CCGEN
        params = {'GUID':'{17ac9db2-ca74-4316-bb74-86a1e8aed0c1}','A':[0.2,0.1,0,0,0,0,0,0,0,0],'BLOCKPHASE':0,'DATEOFF':'N/A',
                'DATEON':'N/A','FLAG':1,'I':[2.2,2.1,2,0,0,0,0,0,0,0],'MEMO':b'mm','MVARATE':5,'TAGS':'tt',
                'V':[1.03,1.02,1,0,0,0,0,0,0,0],'VLOC':0,'VMAXMUL':1.05,'VMIN':0.05}
        keys = ['GLEN LYN', 132]
        o1 = OlxObj.OLCase.addOBJ('CCGEN', keys, params)
    """

    print('\nsample_26_CCGEN')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    gc = OlxObj.OLCase.findOBJ('CCGEN', ['TEXAS', 132])
    if gc!=None:
        print(gc.toString())
        print(gc.I, gc.V, gc.A)
        print(gc.getData(['V', 'A']))

    # Add
    params = {'GUID': '{17ac9db2-ca74-4316-bb74-86a1e8aed0c1}', 'A': [0.2, 0.1, 0, 0, 0, 0, 0, 0, 0, 0], 'BLOCKPHASE': 0, 'DATEOFF': 'N/A',
             'DATEON': 'N/A', 'FLAG': 1, 'I': [2.2, 2.1, 2, 0, 0, 0, 0, 0, 0, 0], 'MEMO': 'memo', 'MVARATE': 5, 'TAGS': 'tt',
             'V': [1.03, 1.02, 1, 0, 0, 0, 0, 0, 0, 0], 'VLOC': 0, 'VMAXMUL': 1.05, 'VMIN': 0.05}
    keys = ['GLEN LYN', 132]
    o1 = OlxObj.OLCase.addOBJ('CCGEN', keys, params)
    print('New CCGEN: ', o1.paramstr)


def sample_27_LOAD(olxpath, fi):
    """ (Load) LOAD object demo

        # various ways to find a LOAD
        lo = OlxObj.OLCase.findOBJ('LOAD',"{6189f131-2137-4eac-8435-a0b49a94395e}") # GUID
        lo = OlxObj.OLCase.findOBJ("{6189f131-2137-4eac-8435-a0b49a94395e}")        # GUID
        lo = OlxObj.OLCase.findOBJ('LOAD',"[LOAD] 4 'TENNESSEE' 132 kV")            # STR
        lo = OlxObj.OLCase.findOBJ("[LOAD] 4 'TENNESSEE' 132 kV")                   # STR
        lo = OlxObj.OLCase.findOBJ('LOAD',['TENNESSEE',132])                        # BUS = [name,kv] (or STR or GUID or Bus Number )
        lo = OlxObj.OLCase.findOBJ('LOAD',4)                                        # BUS = Bus Number
        lo = OlxObj.OLCase.findOBJ('LOAD',"[BUS] 4 'TENNESSEE' 132 kV")             # BUS = Bus STR

        # LOAD can be accessed by related object
        lo = OlxObj.OLCase.LOAD[0]     # from NETWORK (OlxObj.OLCase)
        lo = b1.LOAD                   # from BUS (b1)
        lo = lu1.LOAD                  # from LOADUNIT (lu1)

        # LOAD Data : 'BUS','FLAG','GUID','JRNL','KEYSTR','LOADUNIT','MEMO','P','PARAMSTR','PYTHONSTR','Q','TAGS','UNGROUNDED'

        print(lo.P,lo.Q)                 # (float,float) LOAD P,Q
        print(lo.getData('P'))           # (float) LOAD P
        print(lo.getData(['P','Q']))     # (dict) LOAD P,Q
        print(lo.getData())              # (dict) get all Data of LOAD

        # some methods
        lo.equals(lo2)                   #(bool) comparison of 2 objects
        lo.delete()                      # delete object
        lo.isInList(la)                  # check if object in in List/Set/tuple of object
        lo.toString()                    # (str) text description/composed of object
        lo.copyFrom(lo2)                 # copy Data from another object
        lo.getAttributes()               # [str] list attributes of object
        lo.changeData(sParam,value)      # change Data (sample_6_changeData())
        lo.postData()                    # Perform validation and update object data in the network database

        #
        lo.getData('XXX')                # help details in message error

        # Add LOAD
        params = {'GUID':'{1189f131-2137-4eac-8435-a0b49a94391e}','FLAG':1,'MEMO':b'mm','TAGS':'tt','UNGROUNDED':0}
        keys = ['texas', 132]
        o1 = OlxObj.OLCase.addOBJ('LOAD', keys, params)
    """

    print('\nsample_27_LOAD')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    lo = OlxObj.OLCase.findOBJ("[LOAD] 4 'TENNESSEE' 132 kV")
    if lo!=None:
        print(lo.toString())
        print(lo.P, lo.FLAG)
        print(lo.getData(['P', 'FLAG']))

    # Add
    params = {'GUID': '{1189f131-2137-4eac-8435-a0b49a94391e}','FLAG': 1, 'MEMO': 'mm', 'TAGS': 'tt', 'UNGROUNDED': 0}
    keys = ['texas', 132]
    o1 = OlxObj.OLCase.addOBJ('LOAD', keys, params)
    print('New LOAD: ', o1.paramstr)


def sample_28_LOADUNIT(olxpath, fi):
    """ (Load unit) LOADUNIT object demo

        # various ways to find a LOADUNIT
        lu1 = OlxObj.OLCase.findOBJ('LOADUNIT',"{1a6fb93e-fa85-4380-96eb-05e30bcb2078}") # GUID
        lu1 = OlxObj.OLCase.findOBJ("{1a6fb93e-fa85-4380-96eb-05e30bcb2078}")            # GUID
        lu1 = OlxObj.OLCase.findOBJ('LOADUNIT',"[LOADUNIT]  1@4 'TENNESSEE' 132 kV")     # STR
        lu1 = OlxObj.OLCase.findOBJ("[LOADUNIT]  1@4 'TENNESSEE' 132 kV")                # STR
        lu1 = OlxObj.OLCase.findOBJ('LOADUNIT',['TENNESSEE',132,'1'])                    # [BUS,CID] BUS = [name,kv] (or STR or GUID or Bus Number )
        lu1 = OlxObj.OLCase.findOBJ('LOADUNIT',[4,'1'])                                  # [BUS,CID] BUS = Bus Number
        lu1 = OlxObj.OLCase.findOBJ('LOADUNIT',["[BUS] 4 'TENNESSEE' 132 kV",'1'])       # [BUS,CID] BUS = Bus STR


        # LOADUNIT can be accessed by related object
        lu1 = OlxObj.OLCase.LOADUNIT[0]     # from NETWORK (OlxObj.OLCase)
        lu1 = b1.LOADUNIT[0]                # from BUS (b1)
        lu1 = lo.LOADUNIT[0]                # from LOAD (lo)

        # LOADUNIT Data : 'BUS','CID','DATEOFF','DATEON','FLAG','GUID','JRNL','KEYSTR','LOAD','MVAR','MW','P',
                            'PARAMSTR','PYTHONSTR','Q','TAGS'

        print(lu1.P,lu1.Q)              # (float,float) LOADUNIT P,Q
        print(lu1.getData('P'))         # (float) LOADUNIT P
        print(lu1.getData(['P','Q']))   # (dict) LOADUNIT P,Q
        print(lu1.getData())            # (dict) get all Data of LOADUNIT

        # some methods
        lu1.equals(lu2)                   # (bool) comparison of 2 objects
        lu1.delete()                      # delete object
        lu1.isInList(la)                  # check if object in in List/Set/tuple of object
        lu1.toString()                    # (str) text description/composed of object
        lu1.copyFrom(lu2)                 # copy Data from another object
        lu1.getAttributes()               # [str] list attributes of object
        lu1.changeData(sParam,value)      # change Data (sample_6_changeData())
        lu1.postData()                    # Perform validation and update object data in the network database

        #
        lu1.getData('XXX')                # help details in message error

        # Add LOADUNIT
        params = {'GUID':'{5a6fb93e-fa85-4380-96eb-05e30bcb2071}','DATEOFF':'N/A','DATEON':'N/A','FLAG':1,'MVAR':[1.6,0,0],'MW':[7.6,0,0],'TAGS':'tt'}
        keys = ['texas', 132, '1']
        o1 = OlxObj.OLCase.addOBJ('LOADUNIT', keys, params)
    """

    print('\nsample_28_LOADUNIT')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    lu1 = OlxObj.OLCase.findOBJ("[LOADUNIT]  1@4 'TENNESSEE' 132 kV")
    if lu1!=None:
        print(lu1.toString())
        print(lu1.P, lu1.MVAR)
        print(lu1.getData(['P', 'MVAR']))

    # Add
    params = {'GUID': '{5a6fb93e-fa85-4380-96eb-05e30bcb2071}', 'DATEOFF': 'N/A',
             'DATEON': 'N/A', 'FLAG': 1, 'MVAR': [1.6, 0, 0], 'MW': [7.6, 0, 0], 'TAGS': 'tt'}
    keys = ['texas', 132, '1']
    o1 = OlxObj.OLCase.addOBJ('LOADUNIT', keys, params)
    print('New LOADUNIT: ', o1.paramstr)


def sample_29_SHUNT(olxpath, fi):
    """ (shunt) SHUNT object demo

        # various ways to find a SHUNT
        sh1 = OlxObj.OLCase.findOBJ('SHUNT',"{7a5293e3-1cc4-4795-a267-36bb1c6702c8}") # GUID
        sh1 = OlxObj.OLCase.findOBJ("{7a5293e3-1cc4-4795-a267-36bb1c6702c8}")         # GUID
        sh1 = OlxObj.OLCase.findOBJ('SHUNT',"[SHUNT] 11 'ROANOKE' 13.8 kV")           # STR
        sh1 = OlxObj.OLCase.findOBJ("[SHUNT] 11 'ROANOKE' 13.8 kV")                   # STR
        sh1 = OlxObj.OLCase.findOBJ('SHUNT',['ROANOKE',13.8])                         # BUS = [name,kv] (or STR or GUID or Bus Number )
        sh1 = OlxObj.OLCase.findOBJ('SHUNT',11)                                       # BUS = Bus Number
        sh1 = OlxObj.OLCase.findOBJ('SHUNT',"[BUS] 11 'ROANOKE' 13.8 kV")             # BUS = Bus STR

        # SHUNT can be accessed by related object
        sh1 = OlxObj.OLCase.SHUNT[0]     # from NETWORK (OlxObj.OLCase)
        sh1 = b1.SHUNT                   # from BUS (b1)
        sh1 = su1.SHUNT                  # from SHUNTUNIT (su1)

        # SHUNT Data : 'BUS','FLAG','GUID','JRNL','KEYSTR','MEMO','PARAMSTR','PYTHONSTR','SHUNTUNIT','TAGS'
        #
        print(sh1.FLAG,sh1.MEMO)             # (int str) SHUNT FLAG,MEMO
        print(sh1.getData('FLAG'))           # (int) SHUNT FLAG
        print(sh1.getData(['FLAG','MEMO']))  # (dict) SHUNT FLAG,MEMO
        print(sh1.getData())                 # (dict) get all Data of SHUNT

        # some methods
        sh1.equals(sh2)                   # (bool) comparison of 2 objects
        sh1.delete()                      # delete object
        sh1.isInList(la)                  # check if object in in List/Set/tuple of object
        sh1.toString()                    # (str) text description/composed of object
        sh1.getAttributes()               # [str] list attributes of object
        sh1.copyFrom(sh2)                 # copy Data from another object
        sh1.changeData(sParam,value)      # change Data (sample_6_changeData())
        sh1.postData()                    # Perform validation and update object data in the network database
        #
        sh1.getData('XXX')                # help details in message error

        # Add SHUNT
        params = {'GUID':'{6a5293e3-1cc4-4795-a267-36bb1c6702c1}','FLAG':1,'MEMO':b'mm','TAGS':'tt'}
        keys = ['claytor', 132]
        o1 = OlxObj.OLCase.addOBJ('SHUNT', keys, params)
    """

    print('\nsample_29_SHUNT')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    sh1 = OlxObj.OLCase.findOBJ("[SHUNT] 11 'ROANOKE' 13.8 kV")
    if sh1!=None:
        print(sh1.toString())
        print(sh1.FLAG, sh1.MEMO)
        for u1 in sh1.SHUNTUNIT:
            print('\t', u1.toString())

    # Add
    params = {'GUID': '{6a5293e3-1cc4-4795-a267-36bb1c6702c1}', 'FLAG': 1, 'MEMO': 'mm1', 'TAGS': 'tt'}
    keys = ['claytor', 132]
    o1 = OlxObj.OLCase.addOBJ('SHUNT', keys, params)
    print('New SHUNT: ', o1.paramstr)


def sample_30_SHUNTUNIT(olxpath, fi):
    """ (Shunt unit) SHUNTUNIT object demo

        # various ways to find a SHUNTUNIT
        su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT',"{1676c57f-eb1f-437d-8ef7-91afc0f7b4ba}") # GUID
        su1 = OlxObj.OLCase.findOBJ("{1676c57f-eb1f-437d-8ef7-91afc0f7b4ba}")             # GUID
        su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT',"[CAPUNIT]  1@11 'ROANOKE' 13.8 kV")      # STR
        su1 = OlxObj.OLCase.findOBJ("[CAPUNIT]  1@11 'ROANOKE' 13.8 kV")                  # STR
        su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT',['ROANOKE',13.8,'1'])                     # [BUS,CID] BUS = [name,kv] (or STR or GUID or Bus Number )
        su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT',[11,'1'])                                 # [BUS,CID] BUS = Bus Number
        su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT',["[BUS] 11 'ROANOKE' 13.8 kV",'1'])       # [BUS,CID] BUS = Bus STR

        # SHUNTUNIT can be accessed by related object
        su1 = OlxObj.OLCase.SHUNTUNIT[0]     # from NETWORK (OlxObj.OLCase)
        su1 = b1.SHUNTUNIT[0]                # from BUS (b1)
        su1 = sh.SHUNTUNIT[0]                # from SHUNT (sh)

        # SHUNTUNIT Data : 'B','B0','CID','DATEOFF','DATEON','FLAG','G','G0','GUID','JRNL','KEYSTR','PARAMSTR',
                            'PYTHONSTR','SHUNT','TAGS','TX3'
        #
        print(su1.B,su1.B0)             # (float,float) SHUNTUNIT B,B0
        print(su1.getData('B'))         # (float) SHUNTUNIT B
        print(su1.getData(['B','B0']))  # (dict) SHUNTUNIT B,B0
        print(su1.getData())            # (dict) get all Data of SHUNTUNIT

        # some methods
        su1.equals(su2)                   # (bool) comparison of 2 objects
        su1.delete()                      # delete object
        su1.isInList(la)                  # check if object in in List/Set/tuple of object
        su1.getAttributes()               # [str] list attributes of object
        su1.toString()                    # (str) text description/composed of object
        su1.copyFrom(su2)                 # copy Data from another object
        su1.changeData(sParam,value)      # change Data (sample_6_changeData())
        su1.postData()                    # Perform validation and update object data in the network database

        #
        su1.getData('XXX')                # help details in message error

        # Add SHUNTUNIT
        params = {'GUID':'{2876c57f-eb1f-437d-8ef7-91afc0f7b4ba}','B':0,'B0':0,'DATEOFF':'N/A','DATEON':'N/A','FLAG':1,'G':0,'G0':0,'TAGS':'','TX3':0}
        keys = ['ROANOKE', 13.8, '2']
        o1 = OlxObj.OLCase.addOBJ('SHUNTUNIT', keys, params)
    """

    print('\nsample_30_SHUNTUNIT')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    su1 = OlxObj.OLCase.findOBJ('SHUNTUNIT', ['ROANOKE', 13.8, '1'])
    if su1!=None:
        print(su1.toString())
        print(su1.B, su1.B0)
        print(su1.getData(['B', 'B0']))

    # Add
    params = {'GUID': '{2876c57f-eb1f-437d-8ef7-91afc0f7b4ba}', 'B': 0.5, 'B0': 0.52,
             'DATEOFF': 'N/A', 'DATEON': 'N/A', 'FLAG': 1, 'G': 0, 'G0': 0, 'TAGS': '', 'TX3': 0}
    keys = ['ROANOKE', 13.8, '2']
    o1 = OlxObj.OLCase.addOBJ('SHUNTUNIT', keys, params)
    print('New SHUNTUNIT: ', o1.paramstr)


def sample_31_SVD(olxpath, fi):
    """ (Switched Shunt) SVD object demo

        # various ways to find a SVD
        sv1 = OlxObj.OLCase.findOBJ('SVD',"{fbaf4cb3-1b5d-42d1-b9e9-584c15f1d056}") # GUID
        sv1 = OlxObj.OLCase.findOBJ("{fbaf4cb3-1b5d-42d1-b9e9-584c15f1d056}")       # GUID
        sv1 = OlxObj.OLCase.findOBJ('SVD',"[SVD] 11 'ROANOKE' 13.8 kV")             # STR
        sv1 = OlxObj.OLCase.findOBJ("[SVD] 11 'ROANOKE' 13.8 kV")                   # STR
        sv1 = OlxObj.OLCase.findOBJ('SVD',['ROANOKE',13.8])                         # BUS = [name,kv] (or STR or GUID or Bus Number )
        sv1 = OlxObj.OLCase.findOBJ('SVD',11)                                       # BUS = Bus Number
        sv1 = OlxObj.OLCase.findOBJ('SVD',"[BUS] 11 'ROANOKE' 13.8 kV")             # BUS = Bus STR

        # SVD can be accessed by related object
        sv1 = OlxObj.OLCase.SVD[0]     # from NETWORK (OlxObj.OLCase)
        sv1 = b1.SVD                   # from BUS (b1)

        # SVD Data : 'B','B0','BUS','B_USE','CNTBUS','DATEOFF','DATEON','FLAG','GUID','JRNL','KEYSTR','MEMO',
                    'MODE','PARAMSTR','PYTHONSTR','STEP','TAGS','VMAX','VMIN'
        #
        print(sv1.B,sv1.B0)             # (float,float) SVD B,B0
        print(sv1.getData('B'))         # (float) SVD B
        print(sv1.getData(['B','B0']))  # (dict) SVD B,B0
        print(sv1.getData())            # (dict) get all Data of SVD

        # some methods
        sv1.equals(sv2)                   # (bool) comparison of 2 objects
        sv1.delete()                      # delete object
        sv1.isInList(la)                  # check if object in in List/Set/tuple of object
        sv1.toString()                    # (str) text description/composed of object
        sv1.copyFrom(sv2)                 # copy Data from another object
        sv1.getAttributes()               # [str] list attributes of object
        sv1.changeData(sParam,value)      # change Data (sample_6_changeData())
        sv1.postData()                    # Perform validation and update object data in the network database
        #
        sv1.getData('XXX')                # help details in message error

        # Add SVD
        params = {'GUID':'{fbaf4cb3-1b5d-42d1-b9e9-584c15f1d051}','B':[2,1,0,0,0,0,0,0],'B0':[2,1,0,0,0,0,0,0],
                'B_USE':2,'CNTBUS':"[BUS] 11 'ROANOKE' 13.8 kV",'DATEOFF':'N/A','DATEON':'N/A','FLAG':1,'MEMO':b'mm',
                'MODE':1,'STEP':[1,2,0,0,0,0,0,0],'TAGS':'tt','VMAX':1.02,'VMIN':0.98}
        keys = ['nevada', 132]
        o1 = OlxObj.OLCase.addOBJ('SVD', keys, params)
    """

    print('\nsample_31_SVD')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('SVD', ['ROANOKE', 13.8])
    if o1!=None:
        print(o1.toString())
        print(o1.B, o1.B0)
        print(o1.getData(['B', 'B0']))

    # Add
    params = {'GUID': '{fbaf4cb3-1b5d-42d1-b9e9-584c15f1d051}', 'B': [2, 1, 0, 0, 0, 0, 0, 0], 'B0': [2, 1, 0, 0, 0, 0, 0, 0],
             'B_USE': 2, 'CNTBUS': "[BUS] 11 'ROANOKE' 13.8 kV", 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'FLAG': 1, 'MEMO': b'mm',
             'MODE': 1, 'STEP': [1, 2, 0, 0, 0, 0, 0, 0], 'TAGS': 'tt', 'VMAX': 1.02, 'VMIN': 0.98}
    keys = ['nevada', 132]
    o1 = OlxObj.OLCase.addOBJ('SVD', keys, params)
    print('New SVD: ', o1.paramstr)


def sample_32_RLYGROUP(olxpath, fi):
    """ (Relay Group) RLYGROUP object demo

        # various ways to find a RLYGROUP
        rg1 = OlxObj.OLCase.findOBJ('RLYGROUP',"{e6d85c69-bda7-44cf-9d25-e0d0885fc681}")                # GUID
        rg1 = OlxObj.OLCase.findOBJ("{e6d85c69-bda7-44cf-9d25-e0d0885fc681}")                           # GUID
        rg1 = OlxObj.OLCase.findOBJ('RLYGROUP',"[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") # STR
        rg1 = OlxObj.OLCase.findOBJ("[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")            # STR
        rg1 = OlxObj.OLCase.findOBJ('RLYGROUP',[6,8,'1','L'])                                           # [b1,b2,CID,BRCODE]
                # BRCODE='X','T','P','L','DC','S','W','XFMR3','XFMR','SHIFTER','LINE','DCLINE2','SERIESRC','SWITCH'
        rg1 = OlxObj.OLCase.findOBJ('RLYGROUP',[['NEVADA',132],10,'1','XFMR'])                          # [b1,b2,CID,BRCODE]

        # RLYGROUP can be accessed by related object
        rg1 = OlxObj.OLCase.RLYGROUP[0]     # from NETWORK (OlxObj.OLCase)
        rg1 = b1.RLYGROUP[0]                # from BUS (b1)
        rg1 = t1.RLYGROUP                   # from TERMINAL (t1)
        rg1 = l1.RLYGROUP1                  # from EQUIPMENT  (l1 in XFMR3,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH)
        rg1 = rl1.RLYGROUP                  # from RELAY (rl1 in RLYOCG,RLYOCP,FUSE,RLYDSG,RLYDSP,RECLSR.RLYD,RLYV)

        # RLYGROUP Data : 'BACKUP','EQUIPMENT','FLAG','FUSE','GUID','INTRPTIME','JRNL','KEYSTR','LOGICRECL','LOGICTRIP','MEMO',
                          'NOTE','OPFLAG','PARAMSTR','PRIMARY','PYTHONSTR','RECLSR','RECLSRTIME','RLYD','RLYDS','RLYDSG',
                          'RLYDSP','RLYOC','RLYOCG','RLYOCP','RLYV','SCHEME','TAGS'
        #
        print(rg1.FLAG,rg1.OPFLAG)               # (int,int) RLYGROUP FLAG,OPFLAG
        print(rg1.getData('FLAG'))               # (int) RLYGROUP FLAG
        print(rg1.getData(['FLAG','OPFLAG']))    # (dict) RLYGROUP FLAG,OPFLAG
        print(rg1.getData())                     # (dict) get all Data of RLYGROUP

        # some methods
        rg1.equals(rg2)                   # (bool) comparison of 2 objects
        rg1.delete()                      # delete object
        rg1.isInList(la)                  # check if object in in List/Set/tuple of object
        rg1.toString()                    # (str) text description/composed of object
        rg1.copyFrom(rg2)                 # copy Data from another object
        rg1.getAttributes()               # [str] list attributes of object
        rg1.changeData(sParam,value)      # change Data (sample_6_changeData())
        rg1.postData()                    # Perform validation and update object data in the network database
        #
        rg1.getData('XXX')                # help details in message error

        # Add RLYGROUP
        params = {'GUID':'{fce7f292-92bd-457d-a6a4-f9f948127b01}','BACKUP':[],'INTRPTIME':0,'MEMO':b'mm','PRIMARY':[],'RECLSRTIME':[0,0,0,0],'TAGS':'tt'}
        keys = [['nevada',132],['ohio',132],'1','LINE']
        o1 = OlxObj.OLCase.addOBJ('RLYGROUP', keys, params)
    """

    print('\nsample_32_RLYGROUP')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYGROUP', [['NEVADA', 132], 10, '1', 'XFMR'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.RECLSRTIME)
        print(o1.getData(['FLAG', 'RECLSRTIME']))

    # Add
    params = {'GUID': '{fce7f292-92bd-457d-a6a4-f9f948127b01}', 'BACKUP': [], 'INTRPTIME': 0,
             'MEMO': 'memo', 'PRIMARY': [], 'RECLSRTIME': [1.0, 0, 0, 0], 'TAGS': 'tt'}
    keys = [['nevada', 132], ['ohio', 132], '1', 'LINE']
    o1 = OlxObj.OLCase.addOBJ('RLYGROUP', keys, params)
    print('New RLYGROUP: ', o1.paramstr)


def sample_33_RLYD(olxpath, fi):
    """ (Differential relay) RLYD object demo

        # various ways to find a RLYD
        rd1 = OlxObj.OLCase.findOBJ('RLYD',"[DEVICEDIFF]  rd1@6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1 L") # STR
        rd1 = OlxObj.OLCase.findOBJ('RLYD',"{66a98db8-7000-4262-be41-b0b67ebddcd2}")                   # GUID
        rd1 = OlxObj.OLCase.findOBJ('RLYD',[6,['TAP BUS',132],'1','L','rld1'])                         # [b1,b2,CID,BRCODE,NAME]

        # RLYD can be accessed by related object
        rd1 = OlxObj.OLCase.RLYD[0]      # from NETWORK (OlxObj.OLCase)
        rd1 = rg1.RLYD[0]                # from RLYGROUP (rg1)

        # RLYD Data : 'ASSETID','BUS','CTGRP1','CTR1','DATEOFF','DATEON','EQUIPMENT','FLAG','GUID','ID','IMIN3I0','IMIN3I2','IMINPH',
                'JRNL','KEYSTR','MEMO','PACKAGE','PARAMSTR','PYTHONSTR','RLYGROUP','RMTE1','RMTE2','SGLONLY','TAGS','TLCCV3I0',
                'TLCCVI2','TLCCVPH','TLCTD3I0','TLCTDI2','TLCTDPH'
        #
        print(rd1.FLAG,rd1.ID)                   # (int,str) RLYD FLAG,ID
        print(rd1.getData('FLAG'))               # (int) RLYD FLAG
        print(rd1.getData(['FLAG','ID']))        # (dict) RLYD FLAG,ID
        print(rd1.getData())                     # (dict) get all Data of RLYD

        # some methods
        rd1.equals(rd2)                   # (bool) comparison of 2 objects
        rd1.delete()                      # delete object
        rd1.isInList(la)                  # check if object in in List/Set/tuple of object
        rd1.toString()                    # (str) text description/composed of object
        rd1.copyFrom(rd2)                 # copy Data from another object
        rd1.getAttributes()               # [str] list attributes of object
        rd1.changeData(sParam,value)      # change Data (sample_6_changeData())
        rd1.postData()                    # Perform validation and update object data in the network database


        rd1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        rd1.getData('XXX')                # help details in message error

        # Add RLYD
        params = {'GUID':'{16a98db8-7000-4262-be41-b0b67ebddcd1}','ASSETID':'','CTGRP1':"[RELAYGROUP] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1 L",\
                'CTR1':1,'DATEOFF':'N/A','DATEON':'N/A','FLAG':1,'IMIN3I0':0.3,'IMIN3I2':0.3,'IMINPH':0.2,'MEMO':b'mm','PACKAGE':1,'RMTE1':None,\
                'RMTE2':"[DEVICEDIFF]  rd3@28 'ARIZONA' 132 kV-'TAP BUS' 132 kV 1 L",'SGLONLY':0,'TAGS':'tt','TLCTD3I0':0,'TLCTDI2':0,'TLCTDPH':0}
        keys = [['nevada', 132], ['TAP BUS', 132], '1', 'LINE', 'rld10']
        o1 = OlxObj.OLCase.addOBJ('RLYD', keys, params)
    """

    print('\nsample_33_RLYD')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYD', [['nevada', 132], ['TAP BUS', 132], '1', 'L', 'rld1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.ID)

    # Add
    params = {'GUID': '{16a98db8-7000-4262-be41-b0b67ebddcd1}', 'ASSETID': '', 'CTGRP1': "[RELAYGROUP] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1 L",
             'CTR1': 10, 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'FLAG': 1, 'IMIN3I0': 0.3, 'IMIN3I2': 0.3, 'IMINPH': 0.2,
             'MEMO': 'mm', 'PACKAGE': 1, 'RMTE1': None,'RMTE2': "[DEVICEDIFF]  rd3@28 'ARIZONA' 132 kV-'TAP BUS' 132 kV 1 L",
             'SGLONLY': 0, 'TAGS': 'tt', 'TLCTD3I0': 0, 'TLCTDI2': 0, 'TLCTDPH': 0}
    keys = [['nevada', 132], ['TAP BUS', 132], '1', 'LINE', 'rld10']
    o1 = OlxObj.OLCase.addOBJ('RLYD', keys, params)
    print('New RLYD: ', o1.paramstr)


def sample_34_RLYV(olxpath, fi):
    """ (Voltage relay) RLYV object demo

        # various ways to find a RLYV
        rv1 = OlxObj.OLCase.findOBJ('RLYV',"{4cb670f4-c444-49bf-afe2-0d08a84e62cd}")                     # GUID
        rv1 = OlxObj.OLCase.findOBJ("{4cb670f4-c444-49bf-afe2-0d08a84e62cd}")                            # GUID
        rv1 = OlxObj.OLCase.findOBJ('RLYV',"[DEVICEVR]  rv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L") # STR
        rv1 = OlxObj.OLCase.findOBJ("[DEVICEVR]  rv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L")        # STR
        rv1 = OlxObj.OLCase.findOBJ('RLYV',[5,['CLAYTOR',132],'1','L','rlv1'])                           # [b1,b2,CID,BRCODE,NAME]

        # RLYV can be accessed by related object
        rv1 = OlxObj.OLCase.RLYV[0]      # from NETWORK (OlxObj.OLCase)
        rv1 = rg1.RLYV[0]                # from RLYGROUP (rg1)

        # RLYV Data : 'ASSETID','BUS','DATEOFF','DATEON','EQUIPMENT','FLAG','GUID','ID','JRNL','KEYSTR','MEMO','OPQTY','OVCVR','OVDELAYTD',
                'OVINST','OVPICKUP','PACKAGE','PARAMSTR','PTR','PYTHONSTR','RLYGROUP','SGLONLY','TAGS','UVCVR','UVDELAYTD','UVINST','UVPICKUP'
        #
        print(rv1.FLAG,rv1.ID)                   # (int,str) RLYV FLAG,ID
        print(rv1.getData('FLAG'))               # (int) RLYV FLAG
        print(rv1.getData(['FLAG','ID']))        # (dict) RLYV FLAG,ID
        print(rv1.getData())                     # (dict) get all Data of RLYV

        # some methods
        rv1.equals(rv2)                   # (bool) comparison of 2 objects
        rv1.delete()                      # delete object
        rv1.isInList(la)                  # check if object in in List/Set/tuple of object
        rv1.toString()                    # (str) text description/composed of object
        rv1.copyFrom(rv2)                 # copy Data from another object
        rv1.getAttributes()               # [str] list attributes of object
        rv1.changeData(sParam,value)      # change Data (sample_6_changeData())
        rv1.postData()                    # Perform validation and update object data in the network database

        rv1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        rv1.getData('XXX')                # help details in message error

        # Add RLYV
        params = {'GUID':'{1cb670f4-c444-49bf-afe2-0d08a84e62cd}','ASSETID':'','DATEOFF':'N/A','DATEON':'N/A',
            'FLAG':1,'MEMO':b'rv1mm','OPQTY':1,'OVDELAYTD':0,'OVINST':0,'OVPICKUP':110,'PACKAGE':1,'PTR':100,
            'SGLONLY':[1,1,1,1],'TAGS':'tt','UVDELAYTD':0,'UVINST':0,'UVPICKUP':100}
        keys = [['FIELDALE', 132], ['CLAYTOR', 132], '1', 'L', 'rlv2']
        o1 = OlxObj.OLCase.addOBJ('RLYV', keys, params)
    """

    print('\nsample_34_RLYV')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYV', [['FIELDALE', 132], ['CLAYTOR', 132], '1', 'L', 'rlv1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.ID)
        print(o1.getData(['OVPICKUP', 'UVPICKUP']))

    # Add
    params = {'GUID': '{1cb670f4-c444-49bf-afe2-0d08a84e62cd}', 'ASSETID': '', 'DATEOFF': 'N/A', 'DATEON': 'N/A',
             'FLAG': 1, 'MEMO': b'rv1mm', 'OPQTY': 1, 'OVDELAYTD': 0, 'OVINST': 0, 'OVPICKUP': 110, 'PACKAGE': 1, 'PTR': 100,
             'SGLONLY': [0, 1, 1, 1], 'TAGS': 'tt', 'UVDELAYTD': 0, 'UVINST': 0, 'UVPICKUP': 100}
    keys = [['FIELDALE', '132'], ['CLAYTOR', '132'], '1', 'L', 'rlv2']
    o1 = OlxObj.OLCase.addOBJ('RLYV', keys, params)
    print('New RLYV: ', o1.paramstr)


def sample_35_FUSE(olxpath,fi):
    """ (Fuse) FUSE object demo

        # various ways to find a FUSE
        fu1 = OlxObj.OLCase.findOBJ('FUSE',"{20949e1d-a4e4-403f-b235-155315d8daef}")                 # GUID
        fu1 = OlxObj.OLCase.findOBJ("{20949e1d-a4e4-403f-b235-155315d8daef}")                        # GUID
        fu1 = OlxObj.OLCase.findOBJ('FUSE',"[FUSE]  NV Fuse@6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1 L") # STR
        fu1 = OlxObj.OLCase.findOBJ("[FUSE]  NV Fuse@6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1 L")        # STR
        fu1 = OlxObj.OLCase.findOBJ('FUSE',[6,['TAP BUS',132],'1','L','NV Fuse'])                    # [b1,b2,CID,BRCODE,NAME]

        # FUSE can be accessed by related object
        fu1 = OlxObj.OLCase.FUSE[0]      # from NETWORK (OlxObj.OLCase)
        fu1 = rg1.FUSE[0]                # from RLYGROUP (rg1)

        # FUSE Data : 'ASSETID','BUS','DATEOFF','DATEON','EQUIPMENT','FLAG','FUSECURDIV','FUSECVE','GUID','ID','JRNL','KEYSTR',
                'LIBNAME','MEMO','PACKAGE','PARAMSTR','PYTHONSTR','RATING','RLYGROUP','TAGS','TIMEMULT','TYPE'
        #
        print(fu1.FLAG,fu1.RATING)               # (int,float) FUSE FLAG,RATING
        print(fu1.getData('FLAG'))               # (int) FUSE FLAG
        print(fu1.getData(['FLAG','RATING']))    # (dict) FUSE FLAG,RATING
        print(fu1.getData())                     # (dict) get all Data of FUSE

        # some methods
        fu1.equals(fu2)                   # (bool) comparison of 2 objects
        fu1.delete()                      # delete object
        fu1.isInList(la)                  # check if object in in List/Set/tuple of object
        fu1.toString()                    # (str) text description/composed of object
        fu1.copyFrom(fu2)                 # copy Data from another object
        fu1.getAttributes()               # [str] list attributes of object
        fu1.changeData(sParam,value)      # change Data (sample_6_changeData())
        fu1.postData()                    # Perform validation and update object data in the network database
        #
        fu1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        fu1.getData('XXX')                # help details in message error

        # Add FUSE
        params = {'GUID':'{10949e1d-a4e4-403f-b235-155315d8daef}','ASSETID':'','DATEOFF':'N/A','DATEON':'N/A',
                'FLAG':1,'FUSECURDIV':20,'FUSECVE':1,'LIBNAME':'ABB','MEMO':b'mm','PACKAGE':0,'RATING':0,
                'TAGS':'tt','TIMEMULT':1,'TYPE':'CLE1-15-030'}
        keys = [['nevada',132], ['TAP BUS',132], '1', 'LINE', 'NV fuse2']
        o1 = OlxObj.OLCase.addOBJ('FUSE', keys, params)
    """

    print('\nsample_35_FUSE')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('FUSE', [['nevada', 132], ['TAP BUS', 132], '1', 'L', 'NV Fuse'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.RATING)
        print(o1.getData(['TIMEMULT', 'FUSECVE']))

    # Add
    params = {'GUID': '{10949e1d-a4e4-403f-b235-155315d8daef}', 'ASSETID': '', 'DATEOFF': 'N/A', 'DATEON': 'N/A',
             'FLAG': 1, 'FUSECURDIV': '20', 'FUSECVE': 1, 'LIBNAME': 'ABB', 'MEMO': b'mm', 'PACKAGE': 0, 'RATING': 0,
             'TAGS': 'tt', 'TIMEMULT': 1, 'TYPE': 'CLE1-15-030'}
    keys = [['nevada', 132], ['TAP BUS', 132], '1', 'LINE', 'NV fuse2']
    o1 = OlxObj.OLCase.addOBJ('FUSE', keys, params)
    print('New FUSE: ', o1.paramstr)


def sample_36_RLYDSG(olxpath, fi):
    """ (DS ground relay) RLYDSG object demo

        # various ways to find a RLYDSG
        dg1 = OlxObj.OLCase.findOBJ('RLYDSG',"{a8452803-56fa-452e-bc3e-58be21e453b5}")                          # GUID
        dg1 = OlxObj.OLCase.findOBJ("{a8452803-56fa-452e-bc3e-58be21e453b5}")                                   # GUID
        dg1 = OlxObj.OLCase.findOBJ('RLYDSG',"[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") # STR
        dg1 = OlxObj.OLCase.findOBJ("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")          # STR
        dg1 = OlxObj.OLCase.findOBJ('RLYDSG',[6,['REUSENS',132],'1','L','NV_Reusen G1'])                        # [b1,b2,CID,BRCODE,NAME]

        # RLYDSG can be accessed by related object
        dg1 = OlxObj.OLCase.RLYDSG[0]      # from NETWORK (OlxObj.OLCase)
        dg1 = rg1.RLYDSG[0]                # from RLYGROUP (rg1)

        # RLYDSG Data : 'ASSETID','BUS','DATEOFF','DATEON','DSTYPE','EQUIPMENT','FLAG','GUID','ID','ID2','JRNL','KEYSTR','LIBNAME',
                        'MEMO','PACKAGE','PARAMSTR','PYTHONSTR','RLYGROUP','SETTINGNAME','SETTINGSTR','SNLZONE','TAGS','VTBUS',
                        'Z2OCPICKUP','Z2OCTD','Z2OCTYPE'
        #
        print(dg1.FLAG,dg1.DSTYPE)               # (int,str) RLYDSG FLAG,DSTYPE
        print(dg1.getData('FLAG'))               # (int) RLYDSG FLAG
        print(dg1.getData(['FLAG','DSTYPE']))    # (dict) RLYDSG FLAG,DSTYPE
        print(dg1.getData())                     # (dict) get all Data of RLYDSG

        # setting of RLYDSG
        print(dg1.getSetting('CT ratio'))              # RLYDSG get Setting 'CT ratio'
        print(dg1.getSetting(['CT ratio','PT ratio'])) # RLYDSG get Setting 'CT ratio','PT ratio'
        print(dg1.getSetting())                        # get all Setting of RLYDSG
        print(dg1.SETTINGNAME)                         # get all Setting name of RLYDSG

        # some methods
        dg1.equals(dg2)                   # (bool) comparison of 2 objects
        dg1.delete()                      # delete object
        dg1.isInList(la)                  # check if object in in List/Set/tuple of object
        dg1.toString()                    # (str) text description/composed of object
        dg1.copyFrom(dg2)                 # copy Data from another object
        dg1.getAttributes()               # [str] list attributes of object
        dg1.changeData(sParam,value)      # change Data (sample_6_changeData())
        dg1.postData()                    # Perform validation and update object data in the network database
        #
        dg1.changeSetting(name,value)                          # change Setting (sample_6_Setting())
        dg1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        dg1.getData('XXX')                # help details in message error
        dg1.getSetting('XXX')

        # Add RLYDSG
        params = {'GUID': '{a8452803-56fa-452e-bc3e-58be21e453b1}', 'ASSETID': '', 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'DSTYPE': 'CEY-Type',
             'FLAG': 1, 'TYPE': 'CEYX', 'LIBNAME': '', 'MEMO': b'mm', 'PACKAGE': 0, 'SNLZONE': 0, 'STARTZ2FLAG': 1, 'TAGS': 'tt',
             'VTBUS': "[BUS] 6 'NEVADA' 132 kV",'Z2OCPICKUP': 0, 'Z2OCTD': 0, 'Z2OCTYPE': '__Fixed'}
        settings = {'CT ratio': '100:1', 'CT polarity is reversed': 'No', 'PT ratio': '100:1', 'Min I': '0.',
               'Zone 1 delay': '0.','K1 Mag': '0.', 'K1 Ang': '0.', 'K2 Mag': '0.', 'K2 Ang': '0.', 'Z1_Imp.': '5',
               'Z1_Ang.': '75', 'Z2_Offset Imp.': '0','Z2_Offset Ang.': '75', 'Z2_Imp.': '10', 'Z2_Ang.': '75',
               'Z2_Delay': '0.6', 'Z3_Offset Imp.': '0', 'Z3_Offset Ang.': '75',
               'Z3_Imp.': '16', 'Z3_Ang.': '75', 'Z3_Delay': '1', 'Z3_Frwrd(1)/Rev(0)': '1'}
        keys = [['nevada', 132], ['REUSENS', 132], '1', 'LINE', 'ds g10']
        o1 = OlxObj.OLCase.addOBJ('DSRLYG', keys, params, settings)
    """

    print('\nsample_36_RLYDSG')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYDS', [['nevada', 132], ['REUSENS', 132], '1', 'L', 'NV_Reusen G1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.DSTYPE)
        print(o1.getSetting('CT ratio'))
        print('All Setting: ', o1.getSetting())

    # Add
    params = {'GUID': '{a8452803-56fa-452e-bc3e-58be21e453b1}', 'ASSETID': '', 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'DSTYPE': 'CEY-Type',
             'FLAG': 1, 'TYPE': 'CEYX', 'LIBNAME': '', 'MEMO': b'mm', 'PACKAGE': '0', 'SNLZONE': 0,'STARTZ2FLAG': 1, 'TAGS': 'tt',
             'VTBUS': "[BUS] 6 'NEVADA' 132 kV",'Z2OCPICKUP': 0, 'Z2OCTD': 0, 'Z2OCTYPE': '__Fixed'}
    settings = {'CT ratio': '100:1', 'CT polarity is reversed': 'No', 'PT ratio': '100:1', 'Min I': '0.',
               'Zone 1 delay': '0.','K1 Mag': '0.', 'K1 Ang': '0.', 'K2 Mag': '0.', 'K2 Ang': '0.', 'Z1_Imp.': '5',
               'Z1_Ang.': '75', 'Z2_Offset Imp.': '0','Z2_Offset Ang.': '75', 'Z2_Imp.': '10', 'Z2_Ang.': '75',
               'Z2_Delay': '0.6', 'Z3_Offset Imp.': '0', 'Z3_Offset Ang.': '75',
               'Z3_Imp.': '16', 'Z3_Ang.': '75', 'Z3_Delay': '1', 'Z3_Frwrd(1)/Rev(0)': '1'}
    keys = [['nevada', 132], ['REUSENS', 132], 1, 'L', 'ds g10']
    o1 = OlxObj.OLCase.addOBJ('RLYDSG', keys, params, settings)
    print('New RLYDSG: ', o1.paramstr)


def sample_37_RLYDSP(olxpath, fi):
    """ (DS phase relay) RLYDSP object demo

        # various ways to find a RLYDSP
        dp1 = OlxObj.OLCase.findOBJ('RLYDSP',"{0bd3f991-00d6-48ff-9aca-50ee86f5b22b}")                      # GUID
        dp1 = OlxObj.OLCase.findOBJ("{0bd3f991-00d6-48ff-9aca-50ee86f5b22b}")                               # GUID
        dp1 = OlxObj.OLCase.findOBJ('RLYDSP',"[DSRLYP]  NVPhase1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") # STR
        dp1 = OlxObj.OLCase.findOBJ("[DSRLYP]  NVPhase1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")          # STR
        dp1 = OlxObj.OLCase.findOBJ('RLYDSP',[6,['REUSENS',132],'1','L','NVPhase1'])                        # [b1,b2,CID,BRCODE,NAME]

        # RLYDSP can be accessed by related object
        dp1 = OlxObj.OLCase.RLYDSP[0]      # from NETWORK (OlxObj.OLCase)
        dp1 = rg1.RLYDSP[0]                # from RLYGROUP (rg1)

        # RLYDSP Data : 'ASSETID','BUS','DATEOFF','DATEON','DSTYPE','EQUIPMENT','FLAG','GUID','ID','ID2','KEYSTR','LIBNAME','MEMO','PACKAGE',
            'PARAMSTR','PYTHONSTR','RLYGROUP','SETTINGNAME','SETTINGSTR','SNLZONE','TAGS','VTBUS','Z2OCPICKUP','Z2OCTD','Z2OCTYPE'
        #
        print(dp1.FLAG,dp1.DSTYPE)               # (int,str) RLYDSP FLAG,DSTYPE
        print(dp1.getData('FLAG'))               # (int) RLYDSP FLAG
        print(dp1.getData(['FLAG','DSTYPE']))    # (dict) RLYDSP FLAG,DSTYPE
        print(dp1.getData())                     # (dict) get all Data of RLYDSP

        # setting of RLYDSP
        print(dp1.getSetting('CT ratio'))              # RLYDSP get Setting 'CT ratio'
        print(dp1.getSetting(['CT ratio','PT ratio'])) # RLYDSP get Setting 'CT ratio','PT ratio'
        print(dp1.getSetting())                        # get all Setting of RLYDSP
        print(dp1.SETTINGNAME)                         # get all Setting name of RLYDSP

        # some methods
        dp1.equals(dp2)                   # (bool) comparison of 2 objects
        dp1.delete()                      # delete object
        dp1.isInList(la)                  # check if object in in List/Set/tuple of object
        dp1.toString()                    # (str) text description/composed of object
        dp1.copyFrom(dp2)                 # copy Data from another object
        dp1.getAttributes()               # [str] list attributes of object
        dp1.changeData(sParam,value)      # change Data (sample_6_changeData())
        dp1.postData()                    # Perform validation and update object data in the network database

        #
        dp1.changeSetting(name,value)                          # change Setting (sample_6_Setting())
        dp1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        dp1.getData('XXX')                # help details in message error
        dp1.getSetting('XXX')

        # Add RLYDSP
        params = {'GUID':'{1bd3f991-00d6-48ff-9aca-50ee86f5b22b}','ASSETID':'','DATEOFF':'N/A','DATEON':'N/A','DSTYPE':'KD-Type',
            'FLAG':1,'TYPE':'KD-10','LIBNAME':'','MEMO':b'mm','PACKAGE':0,'SNLZONE':0, ,'STARTZ2FLAG': 1,'TAGS':'tt','VTBUS':"[BUS] 6 'NEVADA' 132 kV",
            'Z2OCPICKUP':0,'Z2OCTD':0,'Z2OCTYPE':'__Fixed'}
        settings = {'CT ratio':'500:1','CT polarity is reversed':'No','PT ratio':'500:1','Min I':'0.','Zone 1 delay':'0.',
            'Z1_Imp. 3-Ph':'5','Z1_Imp. Ph-Ph':'5','Z1_Ang. 3-Ph':'75','Z1_Ang. Ph-Ph':'75','Z2_Inc. Bus(0 or 1)':'0',
            'Z2_Imp. 3-Ph':'12','Z2_Imp. Ph-Ph':'12','Z2_Ang. 3-Ph':'75','Z2_Ang. Ph-Ph':'75','Z2_Delay':'0.5',
            'Z3_Inc. Bus(0 or 1)':'0','Z3_Imp. 3-Ph':'18','Z3_Imp. Ph-Ph':'18','Z3_Ang. 3-Ph':'75','Z3_Ang. Ph-Ph':'75',
            'Z3_Delay':'1','Z3_Frwrd(1)/Rev(0)':'1'}
        keys = [['nevada', 132], ['REUSENS', 132], '1', 'LINE', 'ds p11']
        o1 = OlxObj.OLCase.addOBJ('RLYDSP', keys, params, settings)
    """

    print('\nsample_37_RLYDSP')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYDSP', [['NEVADA', 132], ['REUSENS', 132], '1', 'L', 'NVPhase1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.DSTYPE)
        print(o1.getSetting(['CT ratio', 'PT ratio']))
        print('All Setting: ', o1.getSetting())

    # Add
    params = {'GUID': '{1bd3f991-00d6-48ff-9aca-50ee86f5b22b}', 'ASSETID': '', 'DATEOFF': 'N/A', 'DATEON': 'N/A', 'DSTYPE': 'KD-Type',
             'FLAG': 1, 'TYPE': 'KD-10', 'LIBNAME': '', 'MEMO': b'mm', 'PACKAGE': 0, 'SNLZONE': 0,'STARTZ2FLAG': 1, 'TAGS': 'tt', 'VTBUS': "[BUS] 6 'NEVADA' 132 kV",
             'Z2OCPICKUP': 0, 'Z2OCTD': 0, 'Z2OCTYPE': '__Fixed'}
    settings = {'CT ratio': '500:1', 'CT polarity is reversed': 'No', 'PT ratio': '500:1', 'Min I': '0.', 'Zone 1 delay': '0.',
               'Z1_Imp. 3-Ph': '5', 'Z1_Imp. Ph-Ph': '5', 'Z1_Ang. 3-Ph': '75', 'Z1_Ang. Ph-Ph': '75', 'Z2_Inc. Bus(0 or 1)': '0',
               'Z2_Imp. 3-Ph': '12', 'Z2_Imp. Ph-Ph': '12', 'Z2_Ang. 3-Ph': '75', 'Z2_Ang. Ph-Ph': '75', 'Z2_Delay': '0.5',
               'Z3_Inc. Bus(0 or 1)': '0', 'Z3_Imp. 3-Ph': '18', 'Z3_Imp. Ph-Ph': '18', 'Z3_Ang. 3-Ph': '75', 'Z3_Ang. Ph-Ph': '75',
               'Z3_Delay': '1', 'Z3_Frwrd(1)/Rev(0)': '1'}
    keys = [['NEVADA', 132], ['REUSENS', 132], '1', 'LINE', 'ds p11']
    o1 = OlxObj.OLCase.addOBJ('RLYDSP', keys, params, settings)
    print('New RLYDSP: ', o1.paramstr)

def sample_38_RLYOCG(olxpath, fi):
    """ (OC ground relay) RLYOCG object demo

        # various ways to find a RLYOCG
        og1 = OlxObj.OLCase.findOBJ('RLYOCG',"{c92a3720-7ae9-4cec-aef8-d139a9ee5b96}")                   # GUID
        og1 = OlxObj.OLCase.findOBJ("{c92a3720-7ae9-4cec-aef8-d139a9ee5b96}")                            # GUID
        og1 = OlxObj.OLCase.findOBJ('RLYOCG',"[OCRLYG]  NV-G2@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L") # STR
        og1 = OlxObj.OLCase.findOBJ("[OCRLYG]  NV-G2@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L")          # STR
        og1 = OlxObj.OLCase.findOBJ('RLYOCG',[6,['CLAYTOR',132],'1','L','NV-G2'])                        # [b1,b2,CID,BRCODE,NAME]

        # RLYOCG can be accessed by related object
        og1 = OlxObj.OLCase.RLYOCG[0]      # from NETWORK (OlxObj.OLCase)
        og1 = rg1.RLYOCG[0]                # from RLYGROUP (rg1)

        # RLYOCG Data : 'ASSETID','ASYM','BUS','CT','CTSTR','CTLOC','DATEOFF','DATEON','DTDELAY','DTDIR','DTPICKUP','DTTIMEADD','DTTIMEMULT',
                'EQUIPMENT','FLAG','FLATINST','GUID','ID','INSTSETTING','JRNL','KEYSTR','LIBNAME','MEMO','MINTIME','OCDIR','OPI',
                'PACKAGE','PARAMSTR','PICKUPTAP','POLAR','PYTHONSTR','RLYGROUP','SETTINGNAME','SETTINGSTR','SGNL','TAGS','TAPTYPE',
                'TDIAL','TIMEADD','TIMEMULT','TIMERESET','TYPE'
        #
        print(og1.FLAG,og1.POLAR)               # (int,int) RLYOCG FLAG,POLAR
        print(og1.getData('FLAG'))              # (int) RLYOCG FLAG
        print(og1.getData(['FLAG','POLAR']))    # (dict) RLYOCG FLAG,POLAR
        print(og1.getData())                    # (dict) get all Data of RLYOCG

        # setting of RLYOCG
        print(og1.getSetting('Forward pickup'))                     # RLYOCG get Setting 'Forward pickup'
        print(og1.getSetting(['Forward pickup','Reverse pickup']))  # RLYOCG get Setting 'Forward pickup','Reverse pickup'
        print(og1.getSetting())                                     # get all Setting of RLYOCG
        print(og1.SETTINGNAME)                                      # get all Setting name of RLYOCG

        # some methods
        og1.equals(dg2)                   # (bool) comparison of 2 objects
        og1.delete()                      # delete object
        og1.isInList(la)                  # check if object in in List/Set/tuple of object
        og1.toString()                    # (str) text description/composed of object
        og1.copyFrom(dg2)                 # copy Data from another object
        og1.getAttributes()               # [str] list attributes of object
        og1.changeData(sParam,value)      # change Data (sample_6_changeData())
        og1.postData()                    # Perform validation and update object data in the network database

        #
        og1.changeSetting(name,value)                          # change Setting (sample_6_Setting())
        og1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        og1.getData('XXX')                # help details in message error
        og1.getSetting('XXX')

        # Add RLYOCG
        params = {'GUID':'{c92a3720-7ae9-4cec-aef8-d139a9ee5b91}','ASSETID':'','ASYM':0,'CTSTR': '82:2',,'CTLOC':1,'DATEOFF':'N/A',
            'DATEON':'N/A','DTDELAY':[0,0,0,0,0],'DTDIR':1,'DTPICKUP':[0,0,0,0,0],'DTTIMEADD':0,'DTTIMEMULT':1,'FLAG':1,
            'FLATINST':0,'INSTSETTING':2000,'LIBNAME':'GE','MEMO':b'mm','MINTIME':0,'OCDIR':1,'OPI':0,'PACKAGE':0,'PICKUPTAP':1,
            'POLAR':0,'SGNL':0,'TAGS':'tt','TAPTYPE':'Two','TDIAL':1,'TIMEADD':0,'TIMEMULT':1,'TIMERESET':0,'TYPE':'IAC-77'}
        settings = {'Characteristic angle':'89','Forward pickup':'0','Reverse pickup':'0'}
        keys = [['NEVADA', 132], ['CLAYTOR', 132], '1', 'LINE', 'newOCG']
        o1 = OlxObj.OLCase.addOBJ('RLYOCG', keys, params, settings)
    """

    print('\nsample_38_RLYOCG')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYOCG', [['NEVADA', 132], ['CLAYTOR', 132], '1', 'L', 'NV-G2'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.POLAR)
        print(o1.getSetting('Forward pickup'))
        print('All Setting: ', o1.getSetting())

    # Add
    params = {'GUID': '{c92a3720-7ae9-4cec-aef8-d139a9ee5b91}', 'ASSETID': '', 'ASYM': 0,'CTSTR': '82:2', 'CTLOC': 1, 'DATEOFF': 'N/A',
             'DATEON': 'N/A', 'DTDELAY': [1., 0, 0, 0, 0], 'DTDIR': 1, 'DTPICKUP': [1.0, 0, 0, 0, 0], 'DTTIMEADD': 0, 'DTTIMEMULT': 1, 'FLAG': 1,
             'FLATINST': 0, 'INSTSETTING': 2000, 'LIBNAME': 'GE', 'MEMO': b'mm', 'MINTIME': 0, 'OCDIR': 1, 'OPI': 0, 'PACKAGE': 0, 'PICKUPTAP': 1,
             'POLAR': 0, 'SGNL': 0, 'TAGS': 'tt', 'TAPTYPE': 'Two', 'TDIAL': 1, 'TIMEADD': 0, 'TIMEMULT': 1, 'TIMERESET': 0, 'TYPE': 'IAC-77'}
    settings = {'Characteristic angle': '89', 'Forward pickup': '0', 'Reverse pickup': '0'}
    keys = [['NEVADA', 132], ['CLAYTOR', 132], '1', 'LINE', 'newOCG']
    o1 = OlxObj.OLCase.addOBJ('RLYOCG', keys, params, settings)
    print('New RLYOCG: ', o1.paramstr)
    print(o1.settingstr)


def sample_39_RLYOCP(olxpath, fi):
    """ (OC phase relay) RLYOCP object demo

        # various ways to find a RLYOCP
        op1 = OlxObj.OLCase.findOBJ('RLYOCP',"{b7934abd-7e72-46f2-92f8-7972bf85691d}")                   # GUID
        op1 = OlxObj.OLCase.findOBJ("{b7934abd-7e72-46f2-92f8-7972bf85691d}")                            # GUID
        op1 = OlxObj.OLCase.findOBJ('RLYOCP',"[OCRLYP]  CL-P1@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") # STR
        op1 = OlxObj.OLCase.findOBJ("[OCRLYP]  CL-P1@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")          # STR
        op1 = OlxObj.OLCase.findOBJ('RLYOCP',[['CLAYTOR',132],6,'1','L','CL-P1'])                        # [b1,b2,CID,BRCODE,NAME]

        # RLYOCP can be accessed by related object
        op1 = OlxObj.OLCase.RLYOCP[0]      # from NETWORK (OlxObj.OLCase)
        op1 = rg1.RLYOCP[0]                # from RLYGROUP (rg1)

        # RLYOCP Data : 'ASSETID','ASYM','BUS','CT','CTCONNECT','CTSTR','DATEOFF','DATEON','DTDELAY','DTDIR','DTPICKUP','DTTIMEADD','DTTIMEMULT',
            'EQUIPMENT','FLAG','FLATINST','GUID','ID','INSTSETTING','JRNL','KEYSTR','LIBNAME','MEMO','MINTIME','OCDIR','PACKAGE',
            'PARAMSTR','PICKUPTAP','POLAR','PYTHONSTR','RLYGROUP','SETTINGNAME','SETTINGSTR','SGNL','TAGS','TAPTYPE','TDIAL','TIMEADD',
            'TIMEMULT','TIMERESET','TYPE','VOLTCONTROL','VOLTPERCENT'
        #
        print(op1.FLAG,op1.POLAR)               # (int,int) RLYOCP FLAG,POLAR
        print(op1.getData('FLAG'))              # (int) RLYOCP FLAG
        print(op1.getData(['FLAG','POLAR']))    # (dict) RLYOCP FLAG,POLAR
        print(op1.getData())                    # (dict) get all Data of RLYOCP

        # setting of RLYOCP
        print(op1.getSetting('Forward pickup'))                     # RLYOCP get Setting 'Forward pickup'
        print(op1.getSetting(['Forward pickup','Reverse pickup']))  # RLYOCP get Setting 'Forward pickup','Reverse pickup'
        print(op1.getSetting())                                     # get all Setting of RLYOCP
        print(op1.SETTINGNAME)                                      # get all Setting name of RLYOCP

        # some methods
        op1.equals(op2)                   # (bool) comparison of 2 objects
        op1.delete()                      # delete object
        op1.isInList(la)                  # check if object in in List/Set/tuple of object
        op1.postData()                    # Perform validation and update object data in the network database
        op1.toString()                    # (str) text description/composed of object
        op1.copyFrom(op2)                 # copy Data from another object
        op1.changeData(sParam,value)      # change Data (sample_6_changeData())
        op1.getAttributes()               # [str] list attributes of object
        #
        op1.changeSetting(name,value)                          # change Setting (sample_6_Setting())
        op1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        op1.getData('XXX')                # help details in message error
        op1.getSetting('XXX')

        # Add RLYOCP
        params = {'GUID':'{b1934abd-7e72-46f2-92f8-7972bf85691d}','ASSETID':'','ASYM':0,'CTCONNECT':0,'CTSTR':'500:5','DATEOFF':'N/A',
            'DATEON':'N/A','DTDELAY':[0,0,0,0,0],'DTDIR':1,'DTPICKUP':[0,0,0,0,0],'DTTIMEADD':0,'DTTIMEMULT':1,'FLAG':1,
            'FLATINST':0,'INSTSETTING':2150,'LIBNAME':'ABB','MEMO':b'mm','MINTIME':0,'OCDIR':1,'PACKAGE':0,'PICKUPTAP':2.5,
            'POLAR':0,'SGNL':0,'TAGS':'tt','TAPTYPE':'Two','TDIAL':0.5,'TIMEADD':0,'TIMEMULT':1,'TIMERESET':0,'TYPE':'CO-5',
            'VOLTCONTROL':0,'VOLTPERCENT':0}
        settings = {'Characteristic angle':'89','Forward pickup':'0','Reverse pickup':'0'}
        keys = [['CLAYTOR', 132], ['NEVADA', 132], '1', 'LINE', 'new OCP']
        o1 = OlxObj.OLCase.addOBJ('RLYOCP', keys, params, settings)
    """

    print('\nsample_39_RLYOCP')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RLYOCP',[['CLAYTOR', 132], 6, '1', 'L', 'CL-P1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.POLAR)
        print(o1.getSetting(['Forward pickup', 'Reverse pickup']))
        print(o1.getSetting())

    # Add
    params = {'GUID': '{b1934abd-7e72-46f2-92f8-7972bf85691d}', 'ASSETID': '', 'ASYM': 0,'CTCONNECT': 0,'CTSTR': '500:5', 'DATEOFF': 'N/A',
             'DATEON': 'N/A', 'DTDELAY': [0, 0, 0, 0, 0], 'DTDIR': 1, 'DTPICKUP': [0, 0, 0, 0, 0], 'DTTIMEADD': 0, 'DTTIMEMULT': 1, 'FLAG': 1,
             'FLATINST': 0, 'INSTSETTING': 2150, 'LIBNAME': 'ABB', 'MEMO': b'mm', 'MINTIME': 0, 'OCDIR': 1, 'PACKAGE': 0, 'PICKUPTAP': 2.5,
             'POLAR': 0, 'SGNL': 0, 'TAGS': 'tt', 'TAPTYPE': 'Two', 'TDIAL': 0.5, 'TIMEADD': 0, 'TIMEMULT': 1, 'TIMERESET': 0, 'TYPE': 'CO-5',
             'VOLTCONTROL': 1, 'VOLTPERCENT': 0}
    settings = {'Characteristic angle': '89', 'Forward pickup': '0', 'Reverse pickup': '0'}
    keys = [['CLAYTOR', 132], ['NEVADA', 132], '1', 'LINE', 'new OCP']
    o1 = OlxObj.OLCase.addOBJ('RLYOCP', keys, params, settings)
    print('New RLYOCP: ', o1.paramstr)
    print(o1.settingstr)

def sample_40_RECLSR(olxpath, fi):
    """ (Recloser) RECLSR object demo

        # various ways to find a RECLSR
        re1 = OlxObj.OLCase.findOBJ('RECLSR',"{6a224bbc-80e0-43c4-9f2d-ee8c756def26}")                          # GUID
        re1 = OlxObj.OLCase.findOBJ("{6a224bbc-80e0-43c4-9f2d-ee8c756def26}")                                   # GUID
        re1 = OlxObj.OLCase.findOBJ('RECLSR',"[RECLSRP]  reclsr1_P@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L") # STR
        re1 = OlxObj.OLCase.findOBJ("[RECLSRP]  reclsr1_P@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L")          # STR
        re1 = OlxObj.OLCase.findOBJ('RECLSR',[5,['CLAYTOR',132],'1','L','reclsr1'])                             # [b1,b2,CID,BRCODE,NAME]

        # RECLSR can be accessed by related object
        re1 = OlxObj.OLCase.RECLSR[0]      # from NETWORK (OlxObj.OLCase)
        re1 = rg1.RECLSR[0]                # from RLYGROUP (rg1)

        # RECLSR Data : 'ASSETID','BUS','BYTADD','BYTMULT','DATEOFF','DATEON','EQUIPMENT','FASTOPS','FLAG','GR_FASTTYPE','GR_INST',
            'GR_INSTDELAY','GR_MEMO','GR_MINTIMEF','GR_MINTIMES','GR_MINTRIPF','GR_MINTRIPS','GR_SLOWTYPE','GR_TIMEADDF','GR_TIMEADDS',
            'GR_TIMEMULTF','GR_TIMEMULTS','GUID','ID','INTRPTIME','KEYSTR','LIBNAME','PARAMSTR','PH_FASTTYPE','PH_INST','PH_INSTDELAY',
            'PH_MEMO','PH_MINTIMEF','PH_MINTIMES','PH_MINTRIPF','PH_MINTRIPS','PH_SLOWTYPE','PH_TIMEADDF','PH_TIMEADDS','PH_TIMEMULTF',
            'PH_TIMEMULTS','PYTHONSTR','RECLOSE1','RECLOSE2','RECLOSE3','RLYGROUP','TAGS','TOTALOPS'
        #
        print(re1.FLAG,re1.BYTADD)                # (int,int) RECLSR FLAG,BYTADD
        print(re1.getData('FLAG'))                # (int) RECLSR FLAG
        print(re1.getData(['FLAG','BYTADD']))     # (dict) RECLSR FLAG,BYTADD
        print(re1.getData())                      # (dict) get all Data of RECLSR

        # some methods
        re1.equals(re2)                   # (bool) comparison of 2 objects
        re1.delete()                      # delete object
        re1.isInList(la)                  # check if object in in List/Set/tuple of object
        re1.toString()                    # (str) text description/composed of object
        re1.copyFrom(re2)                 # copy Data from another object
        re1.getAttributes()               # [str] list attributes of object
        re1.changeData(sParam,value)      # change Data (sample_6_changeData())
        re1.postData()                    # Perform validation and update object data in the network database

        #
        re1.computeRelayTime(current,voltage,preVoltage)       # Computes operating time at given currents and voltages
        #
        re1.getData('XXX')                # help details in message error

        # Add RECLSR
        params = {'GUID':'{1555ccea-8595-4846-a171-579c4d490959}','ASSETID':'aid11','BYTADD':2,'BYTMULT':2,'DATEOFF':'2018/1/2','DATEON':'2015/6/7',\
            'FASTOPS':2,'FLAG':1,'GR_FASTTYPE':'240-91-56-05','GR_INST':0.56,'GR_INSTDELAY':0.57,'GR_MEMO':b'gth','GR_MINTIMEF':0.12,\
            'GR_MINTIMES':0.14,'GR_MINTRIPF':1.52,'GR_MINTRIPS':1.42,'GR_SLOWTYPE':'240-91-56-06','GR_TIMEADDF':0.34,'GR_TIMEADDS':0.36,\
            'GR_TIMEMULTF':1.62,'GR_TIMEMULTS':1.64,'INTRPTIME':0.34,'LIBNAME':'COOPER','PH_FASTTYPE':'240-91-56-02','PH_INST':0.54,\
            'PH_INSTDELAY':0.55,'PH_MEMO':b'rgg','PH_MINTIMEF':0.11,'PH_MINTIMES':0.13,'PH_MINTRIPF':1.51,'PH_MINTRIPS':1.41,\
            'PH_SLOWTYPE':'240-91-56-01','PH_TIMEADDF':0.33,'PH_TIMEADDS':0.35,'PH_TIMEMULTF':1.61,'PH_TIMEMULTS':1.63,'RATING':1.2,\
            'RECLOSE1':0.15,'RECLOSE2':0.25,'RECLOSE3':0.35,'TAGS':'gjh;','TOTALOPS':4}
        keys = [['NEVADA', 132], ['REUSENS', 132], '1', 'L', 'reclsrNEW']
        o1 = OlxObj.OLCase.addOBJ('RECLSR', keys, params)
    """

    print('\nsample_40_RECLSR')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('RECLSR', [['FIELDALE', 132], ['CLAYTOR', 132], '1', 'L', 'reclsr1'])
    if o1!=None:
        print(o1.toString())
        print(o1.FLAG, o1.BYTADD)
        print('All Param: ', o1.getData())

    # Add
    params = {'GUID': '{1555ccea-8595-4846-a171-579c4d490959}', 'ASSETID': 'aid11', 'BYTADD': 2, 'BYTMULT': 2, 'DATEOFF': '2018/1/2', 'DATEON': '2015/6/7',
             'FASTOPS': 2, 'FLAG': 1, 'GR_FASTTYPE': '240-91-56-05', 'GR_INST': 0.56, 'GR_INSTDELAY': 0.57, 'GR_MEMO': b'gth', 'GR_MINTIMEF': 0.12,
             'GR_MINTIMES': 0.14, 'GR_MINTRIPF': 1.52, 'GR_MINTRIPS': 1.42, 'GR_SLOWTYPE': '240-91-56-06', 'GR_TIMEADDF': 0.34, 'GR_TIMEADDS': 0.36,
             'GR_TIMEMULTF': 1.62, 'GR_TIMEMULTS': 1.64, 'INTRPTIME': 0.34, 'LIBNAME': 'COOPER', 'PH_FASTTYPE': '240-91-56-02', 'PH_INST': 0.54,
             'PH_INSTDELAY': 0.55, 'PH_MEMO': b'rgg', 'PH_MINTIMEF': 0.11, 'PH_MINTIMES': 0.13, 'PH_MINTRIPF': 1.51, 'PH_MINTRIPS': 1.41,
             'PH_SLOWTYPE': '240-91-56-01', 'PH_TIMEADDF': 0.33, 'PH_TIMEADDS': 0.35, 'PH_TIMEMULTF': 1.61, 'PH_TIMEMULTS': 1.63, 'RATING': 1.2,
             'RECLOSE1': 0.15, 'RECLOSE2': 0.25, 'RECLOSE3': 0.35, 'TAGS': 'gjh;', 'TOTALOPS': 4}
    keys = [['NEVADA', 132], ['REUSENS', 132], '1', 'L', 'reclsrNEW']
    o1 = OlxObj.OLCase.addOBJ('RECLSR', keys, params)
    print('New RECLSR: ', o1.keystr)


def sample_41_SCHEME(olxpath, fi):
    """ (Logic scheme) SCHEME object demo

        # various ways to find a SCHEME
        ls1 = OlxObj.OLCase.findOBJ('SCHEME',"{e0f5356f-819e-416d-be69-034503de8675}")                      # GUID
        ls1 = OlxObj.OLCase.findOBJ("{e0f5356f-819e-416d-be69-034503de8675}")                               # GUID
        ls1 = OlxObj.OLCase.findOBJ('SCHEME',"[PILOT]  scheme1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L") # STR
        ls1 = OlxObj.OLCase.findOBJ("[PILOT]  scheme1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L")          # STR
        ls1 = OlxObj.OLCase.findOBJ('SCHEME',[['FIELDALE',132],2,'1','L','scheme1'])                        # [b1,b2,CID,BRCODE,NAME]

        # SCHEME can be accessed by related object
        ls1 = OlxObj.OLCase.SCHEME[0]      # from NETWORK (OlxObj.OLCase)
        ls1 = rg1.SCHEME[0]                # from RLYGROUP (rg1)

        # SCHEME Data : 'ASSETID','DATEOFF','DATEON','EQUATION','FLAG','GUID','ID','JRNL','KEYSTR','LOGICVARNAME','LOGICVARSTR',
                'MEMO','PARAMSTR','PYTHONSTR','RLYGROUP','SIGNALONLY','TAGS','TYPE'
        #
        print(ls1.FLAG,ls1.EQUATION)               # (int,str) SCHEME FLAG,EQUATION
        print(ls1.getData('FLAG'))                 # (int) SCHEME FLAG
        print(ls1.getData(['FLAG','EQUATION']))    # (dict) SCHEME FLAG,EQUATION
        print(ls1.getData())                       # (dict) get all Data of SCHEME


        # logic of SCHEME
        print(ls1.getLogic('RU_NEAR'))              # SCHEME get Logic var 'RU_NEAR'
        print(ls1.getLogic(['RU_NEAR','RO_NEAR']))  # SCHEME get Logic var 'RU_NEAR','RO_NEAR'
        print(ls1.getLogic())                       # get all Logic var of SCHEME
        print(ls1.LOGICVARNAME)                     # get all Logic var name of SCHEME


        # some methods
        ls1.equals(ls2)                   # (bool) comparison of 2 objects
        ls1.delete()                      # delete object
        ls1.isInList(la)                  # check if object in in List/Set/tuple of object
        ls1.toString()                    # (str) text description/composed of object
        ls1.copyFrom(ls2)                 # copy Data from another object
        ls1.getAttributes()               # [str] list attributes of object
        ls1.changeData(sParam,value)      # change Data (sample_6_changeData())
        ls1.postData()                    # Perform validation and update object data in the network database

        #
        ls1.changeLogic(nameVar,value)    # change logic var (nameVar,value)
        ls1.setLogic(logic)               # set Logic of SCHEME (all logic + EQUATION)
        #
        ls1.getData('XXX')                # help details in message error
        ls1.getLogic('XXX')

        # Add SCHEME
        params = {'GUID':'{a2f5356f-819e-416d-be69-034503de8675}','ASSETID':'aied', 'DATEOFF':'N/A', 'DATEON':'N/A', 'FLAG':1,
            'MEMO':b'rf', 'SIGNALONLY':'1', 'TAGS':'t1', 'TYPE':'PUTT'}
        logics = {'EQUATION': 'RU_NEAR + (RU_FAR @ TS * RO_NEAR) +RO_NEAR',
            'RU_NEAR': ['INST/DT PICKUP', "[OCRLYG]  FL-G1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"],
            'RU_FAR': ['OPEN OP.', "[TERMINAL] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"],
            'TS': 0.43,
            'RO_NEAR': ['PICKUP 3I2', "[DEVICEVR]  rlv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"]}
        keys = [['FIELDALE', 132], ['CLAYTOR', 132], 1, 'L', 'scheme10']
        o1 = OlxObj.OLCase.addOBJ('SCHEME', keys, params, logics)
    """

    print('\nsample_41_SCHEME')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    o1 = OlxObj.OLCase.findOBJ('SCHEME', [['FIELDALE',132], ['CLAYTOR', 132],'1','L','scheme1'])
    if o1!=None:
        print(o1.toString())
        print(o1.getData())
        o1.TYPE = 'POTT'
        o1.EQUATION = 'RU_NEAR + (RU_FAR @ TS * RO_NEAR) + RO_NEAR'
        o1.changeLogic('RO_NEAR', ['OV INST.', "[DEVICEVR]  rlv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"])
        o1.postData()
        print(o1.equation,o1.TYPE)
        print(o1.logicstr)
        print(o1.getLogic('RO_NEAR'))
        print(o1.getLogic(['RU_NEAR','RO_NEAR']))
        print(o1.getLogic())

    # Add
    params = {'GUID':'{a2f5356f-819e-416d-be69-034503de8675}','ASSETID':'aied', 'DATEOFF':'N/A', 'DATEON':'N/A', 'FLAG':1,
         'MEMO':b'rf', 'SIGNALONLY':'1', 'TAGS':'t1', 'TYPE':'PUTT'}
    logics = {'EQUATION': 'RU_NEAR + (RU_FAR @ TS * RO_NEAR) +RO_NEAR',
        'RU_NEAR': ['INST/DT PICKUP', "[OCRLYG]  FL-G1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"],
        'RU_FAR': ['OPEN OP.', "[TERMINAL] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"],
        'TS': 0.43,
        'RO_NEAR': ['PICKUP 3I2', "[DEVICEVR]  rlv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"]}
    keys = [['FIELDALE', 132], ['CLAYTOR', 132], 1, 'L', 'scheme10']
    o1 = OlxObj.OLCase.addOBJ('SCHEME', keys, params, logics)
    print('New SCHEME:', o1.toString())
    print('logics:', o1.logicstr)


def sample_42_SystemParams(olxpath, fi):
    """
    set/get System Parameters
    """

    print('\nsample_42_SystemParams')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    # get SystemParams
    ps = OlxObj.OLCase.getSystemParams()
    print(ps)

    # change SystemParams
    OlxObj.OLCase.setSystemParams({'fSmallX': 0.0002, 'bSimulateCCGen': 1})
    ps = OlxObj.OLCase.getSystemParams()
    print(ps['fSmallX'])
    print(ps['bSimulateCCGen'])


def sample_43_UDF(olxpath, fi):

    print('\nsample_43_UDF')
    OlxObj.load_olxapi(olxpath) # Initialize the OlxAPI module (in olxpath folder)
    OlxObj.OLCase.open(fi, 1)   # Load the OLR network file (fi) as read-only

    # add UDF for BUS
    OlxObj.OLCase.addUDFTemplate('BUS', ['BUSUDF1', 'busudf1', 1])
    OlxObj.OLCase.addUDFTemplate('BUS', ['BUSUDF2', 'busudf2', 2])
    #
    OlxObj.OLCase.addUDFTemplate('LINE', ['LINEUDF2', 'lineudf', 2])
    #
    udf = OlxObj.OLCase.getUDFTemplate()
    print(udf['BUS'])
    print(udf['LINE'])


def run():
    # Various examples
##    sample_1_network(ARGVS.olxpath, ARGVS.fi)
##    sample_2_Bus(ARGVS.olxpath, ARGVS.fi)
##    sample_3_Line(ARGVS.olxpath, ARGVS.fi)
##    sample_4_Terminal(ARGVS.olxpath, ARGVS.fi)
##    sample_5_changeData(ARGVS.olxpath, ARGVS.fi)
##    sample_6_Setting(ARGVS.olxpath, ARGVS.fi)
##    sample_7_ClassicalFault_1(ARGVS.olxpath, ARGVS.fi)
##    sample_8_ClassicalFault_2(ARGVS.olxpath, ARGVS.fi)
##    sample_9_SimultaneousFault(ARGVS.olxpath, ARGVS.fi)
##    sample_10_SEA_1(ARGVS.olxpath, ARGVS.fi)
##    sample_11_SEA_2(ARGVS.olxpath, ARGVS.fi)
##    sample_12_tapLine(ARGVS.olxpath, ARGVS.fi)
##    sample_13_addObj_1(ARGVS.olxpath, ARGVS.fi)
##    sample_14_addObj_2(ARGVS.olxpath, ARGVS.fi)
##    sample_15_XFMR(ARGVS.olxpath, ARGVS.fi)
##    sample_16_XFMR3(ARGVS.olxpath, ARGVS.fi)
##    sample_17_SHIFTER(ARGVS.olxpath, ARGVS.fi)
##    sample_18_SERIESRC(ARGVS.olxpath, ARGVS.fi)
##    sample_19_SWITCH(ARGVS.olxpath, ARGVS.fi)
##    sample_20_BREAKER(ARGVS.olxpath, ARGVS.fi)
    sample_21_MULINE(ARGVS.olxpath, ARGVS.fi)
##    sample_22_GEN(ARGVS.olxpath, ARGVS.fi)
##    sample_23_GENUNIT(ARGVS.olxpath, ARGVS.fi)
##    sample_24_GENW3(ARGVS.olxpath, ARGVS.fi)
##    sample_25_GENW4(ARGVS.olxpath, ARGVS.fi)
##    sample_26_CCGEN(ARGVS.olxpath, ARGVS.fi)
##    sample_27_LOAD(ARGVS.olxpath, ARGVS.fi)
##    sample_28_LOADUNIT(ARGVS.olxpath, ARGVS.fi)
##    sample_29_SHUNT(ARGVS.olxpath, ARGVS.fi)
##    sample_30_SHUNTUNIT(ARGVS.olxpath, ARGVS.fi)
##    sample_31_SVD(ARGVS.olxpath, ARGVS.fi)
##    sample_32_RLYGROUP(ARGVS.olxpath, ARGVS.fi)
##    sample_33_RLYD(ARGVS.olxpath, ARGVS.fi)
##    sample_34_RLYV(ARGVS.olxpath, ARGVS.fi)
##    sample_35_FUSE(ARGVS.olxpath, ARGVS.fi)
##    sample_36_RLYDSG(ARGVS.olxpath, ARGVS.fi)
##    sample_37_RLYDSP(ARGVS.olxpath, ARGVS.fi)
##    sample_38_RLYOCG(ARGVS.olxpath, ARGVS.fi)
##    sample_39_RLYOCP(ARGVS.olxpath, ARGVS.fi)
##    sample_40_RECLSR(ARGVS.olxpath, ARGVS.fi)
##    sample_41_SCHEME(ARGVS.olxpath, ARGVS.fi)
    #
##    sample_42_SystemParams(ARGVS.olxpath, ARGVS.fi)
##    sample_43_UDF(ARGVS.olxpath, ARGVS.fi)
    return

if __name__ == '__main__':
    if not os.path.isfile(ARGVS.fi):
        ARGVS.fi = "SAMPLE12.OLR"
    #
    run()
