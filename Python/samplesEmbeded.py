""" OneLiner embeded Python application examples

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced Systems for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "1.1.1"
__email__     = "support@aspeninc.com"
__status__    = "In-development"

def testBusPicker():
    dlgTitle = "Pick Buses"
    Opt = (c_int)(0)
    BusList = (c_int*100)(0)
    if OLXAPI_OK == OlxAPI.BusPicker(c_char_p(OlxAPI.encode3("Pick Buses")), BusList,Opt):
        SelectedBuses = 'You selected:'
        for i in range(len(BusList)):
            if BusList[i]<=0:
                break
            SelectedBuses = SelectedBuses + '\n' + OlxAPI.PrintObj1LPF(BusList[i])
        print( SelectedBuses )
    else:
        print( 'You pressed Cancel' )

def testLocate1LObj():
    hnd = c_int(0)
    OlxAPI.GetEquipment(TC_BUS,hnd)
    print( OlxAPI.PrintObj1LPF(hnd) )
    if OLXAPI_OK == OlxAPI.Locate1LObj(hnd,0):
       print( 'Found: ' + OlxAPI.PrintObj1LPF(hnd) )
    else:
       print( 'Not found: ' + OlxAPI.PrintObj1LPF(hnd) )
    return 0

testBusPicker()
#testLocate1LObj()

