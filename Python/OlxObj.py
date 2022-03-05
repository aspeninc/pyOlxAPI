"""
ASPEN OLX OBJECT LIBRARY
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__email__     = "support@aspeninc.com"
__status__    = "Release candidate"
__version__   = "2.2.0"
#
import os
import OlxAPI
import OlxAPILib
import OlxAPIConst
from ctypes import c_int, byref,c_double,cast,POINTER,c_char_p,create_string_buffer,pointer

#
def getVersion():
    """ (str) OlxObj library version"""
    return __version__
#
def load_olxapi(dllPath,verbose=True):
    """ Initialize the OlxAPI module.
    Args:
        dllPath (string) : Full path name of the disk folder where the olxapi.dll is located
        verbose (bool)   : Option to print OlxAPI module infos to the consol (True/False)
    Remark: Successfull initialization is required before the OlxObj library objects and
    functions can be utilized.
    """
    OlxAPI.InitOlxAPI(dllPath,verbose)
    v = OlxAPI.Version()
    va = v.split('.')
    if int(va[0])<=15 and int(va[1])<5:
        raise Exception('Incorrect OlxAPI.dll version: found '+v+va[0]+'.'+va[1]+'. Required V15.5 or higher.')
#
def toString(v,nRound=5):
    """ convert object/value to String """
    if v is None:
        return 'None'
    t = type(v)
    if t==str:
        return "'%s'"%v.replace('\n',' ')
    if t==float:
        s1 ='%.'+str(nRound)+'g'
        return s1 % v
    if t==complex:
        if v.imag>=0:
            return '('+ toString(v.real,nRound)+'+' + toString(v.imag,nRound)+'j)'
        return '('+ toString(v.real,nRound) + toString(v.imag,nRound)+'j)'
    try:
        return v.toString()
    except:
        pass
    if t in {list,tuple,set}:
        s1=''
        for v1 in v:
            s1+=toString(v1,nRound)+','
        if v:
            s1 = s1[:-1]
        if t==list:
            return '['+s1+']'
        elif t==tuple:
            return '('+s1+')'
        else:
            return '{'+s1+'}'
    return str(v)
#
__K_INI_NETWORK__ = 0
__CURRENT_FILE_IDX__ = 0
__INDEX_FAULT__ = -1
__COUNT_FAULT__ = 0
__TIERS_FAULT__ = 9
__INDEX_SIMUL__ = 1
__TYPEF_SIMUL__ = ''
#
class RESULT_FLT:
    def __init__(self,index):
        """ Construct a fault solution case
        Args:
            -index (int): case index   """
        if __COUNT_FAULT__==0:
            raise Exception('\nNo fault simulation result buffer is empty.')
        if type(index)==RESULT_FLT:
            index = index.__index__
        if type(index)!=int or index<=0 or index>__COUNT_FAULT__:
            se = '\nRESULT_FLT(index)'
            se += '\n  Case index'
            se += '\n\tRequired           : (int) 1-%i'%__COUNT_FAULT__
            se += '\n\t'+__getErrValue__(int,index)
            raise ValueError(se)
        self.__index__ = index
        self.__tiers__ = __TIERS_FAULT__
        self.__index_simul__ = __INDEX_SIMUL__
        if __TYPEF_SIMUL__=='SEA':
            self.__SEAResult__ = __getSEA_Result__(index)
            self.__fdes__ = str(index)+'. '+self.__SEAResult__['FaultDesc']
        else:
            self.__fdes__ = OlxAPI.FaultDescriptionEx(index,0)
    #
    def __check__(self):
        if self.__index_simul__!=__INDEX_SIMUL__:
            se = '\nAll previous RESULT_FLT are not available'
            se += '\n\t   by a SEA simulation or Classical/Simultaneous with clearPrev=1 has been done'
            se += '\n\tor by OLCase.close()'
            se += '\n\tor by OLCase.open() to open a other OLR file'
            raise Exception(se)
    #
    def __pick__(self):
        self.__check__()
        global __INDEX_FAULT__
        if __INDEX_FAULT__!=self.__index__:
            if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.PickFault(c_int(self.__index__),c_int(self.__tiers__)):
                raise Exception('\nError PickFault index=%i, with index available: 1-'%self.__index__+str(__COUNT_FAULT__))
            __INDEX_FAULT__ = self.__index__
    #
    def setScope(self,tiers=9,buses=[]):
        """ tiers (int): number of tiers around faulted bus to compute solution results
        """
        self.__check__()
        self.__tiers__ = tiers
        if type(tiers)!=int or tiers<0:
            se ='\nRESULT_FLT.setScope()\n\ttiers: an integer>=0 is required (got type %s) tiers='%type(tiers).__name__+str(tiers)
            raise TypeError(se)
    @property
    def FaultDescription(self):
        """ (str) Fault description string """
        return self.__fdes__
    @property
    def SEARes(self):
        """ (dict) SEA Results
                SEARes['time']      : (float) SEA event time stamp [s]
                SEARes['currentMax']: (float) Highest phase fault current magnitude at this step
                SEARes['EventFlag'] : (0/1)   flag showing is this is an user-defined event
                SEARes['EventDesc'] : (str)   Event Description
                SEARes['FaultDesc'] : (str)   Fault Description
                SEARes['breaker']   : []      List breaker that operated in the current SEA event
                SEARes['device']    : []      List device that tripped in the current SEA event """
        try:
            return self.__SEAResult__
        except:
            return None
    @property
    def tiers(self):
        """ (int) Number of tiers in this solution result scope """
        return self.__tiers__
    #
    def current(self,obj=None):
        """ (ABC) Retrieve ABC PHASE post fault current on a object or at the fault
        Args:
            -obj: object None,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,GEN,GENUNIT,GENW3,GENW4,CCGEN
                         LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
        Return:[complex]
            -obj=None
                    => [IA,IB,IC] total fault current
            -obj=GEN,GENUNIT,GENW3,GENW4,CCGEN,LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
                    => [IA,IB,IC]
            -obj=LINE,DCLINE2,SERIESRC,SWITCH
                    => [IA1,IB1,IC1, IA2,IB2,IC2] current from BUS1, BUS2
            -obj=XFMR,SHIFTER
                    => [IA1,IB1,IC1,IN1, IA2,IB2,IC2,IN2]
                      with IN1,IN2 are Neutral current of winding on Bus1, Bus2
            -obj=XFMR3
                    => [IA1,IB1,IC1,IN1, IA2,IB2,IC2,IN2, IA3,IB3,IC3,Id3]
                       with IN1,IN2 are Neutral current of winding on Bus1, Bus2
                            Id3 circulating current on Bus 3 (3-W Transformer only: Delta)
        """
        if obj is not None and type(obj) not in __OLXOBJ_IFLT__:
            se= '\nRESULT_FLT.current(obj) unsupported object'
            se +='\n\tSupported object: None'
            for vi in __OLXOBJ_IFLT__:
                se+=','+vi.__name__
            se +='\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        self.__pick__()
        return __currentFault__(obj,style=3)
    #
    def currentSeq(self,obj=None):
        """ (012) Retrieve 012 SEQUENCE post fault current on a object or at the fault
        Args:
            -obj: object None,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,GEN,GENUNIT,GENW3,GENW4,CCGEN
                         LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
        Return:[complex]
            -obj=None
                    => [I0,I1,I2] total fault current
            -obj=GEN,GENUNIT,GENW3,GENW4,CCGEN,LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
                    => [I0,I1,I2]
            -obj=LINE,DCLINE2,SERIESRC,SWITCH
                    => [I0_1,I1_1,I2_1 , I0_2,I1_2,I2_2] current from BUS1, BUS2
            -obj=XFMR,SHIFTER
                    => [I0_1,I1_1,I2_1,IN1 , I0_2,I1_2,I2_2,IN2]
                      with IN1,IN2 are Neutral current of winding on Bus1, Bus2
            -obj=XFMR3
                    => [I0_1,I1_1,I2_1,IN1 , I0_2,I1_2,I2_2,IN2 , I0_3,I1_3,I2_3,Id3]
                       with IN1,IN2 are Neutral current of winding on Bus1, Bus2
                            Id3 circulating current on Bus 3 (3-W Transformer only: Delta)
        """
        if obj is not None and type(obj) not in __OLXOBJ_IFLT__:
            se= '\nRESULT_FLT.currentSeq(obj) unsupported object'
            se +='\n\tSupported object: None'
            for vi in __OLXOBJ_IFLT__:
                se+=','+vi.__name__
            se +='\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        self.__pick__()
        return __currentFault__(obj,style=1)
    #
    def voltage(self,obj):
        """ (ABC) Retrieve ABC PHASE post-fault voltage of a bus,
                           or of connected buses of a line, transformer, switch or phase shifter
        Args:
            -obj = object BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3,TERMINAL
        return: [voltage complexe]
        -obj=BUS
                => [VA,VB,VC]
        -obj=XFMR3
                => [VA1,VB1,VC1 , VA2,VB2,VC2 , VA3,VB3,VC3]
        -obj=XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,TERMINAL
                => [VA1,VB1,VC1 , VA2,VB2,VC2]
        """
        self.__pick__()
        #
        if type(obj) not in {BUS,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,TERMINAL}:
            se= '\nRESULT_FLT.voltage(obj) unsupported object'
            se +='\n\tSupported object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3,TERMINAL'
            se +='\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        return __voltageFault__(obj,style=3)
    #
    def voltageSeq(self,obj):
        """ (012) Retrieve 012 SEQUENCE post-fault voltage of a bus,
                           or of connected buses of a line, transformer, switch or phase shifter
        Args:
            -obj = object BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3
        return: [voltage complexe]
        -obj=BUS
                => [V0,V1,V2]
        -obj=XFMR3
                => [V0_1,V1_1,V2_1 , V0_2,V1_2,V2_2 , V0_3,V1_3,V2_3]
        -obj=XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH
                => [V0_1,V1_1,V2_1 , V0_2,V1_2,V2_2]
        """
        self.__pick__()
        #
        if type(obj) not in {BUS,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH}:
            se= '\nRESULT_FLT.voltageSeq(obj) unsupported object'
            se +='\n\tSupported object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'
            se +='\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        return __voltageFault__(obj,style=1)
    #
    def optime(self,obj,mult,signalonly):
        """ Operating time of a protective device or logic scheme in a fault
        Args:
            obj : a protective device or logic scheme
            mult: Relay current multiplying factor
            signalonly: [int] Consider relay element signal-only flag 1 - Yes; 0 - No

        return: time,code
            - time: relay operating time in seconds
            - code: relay operation code. Required buffer size: 128 bytes.
                    NOP  No operation.
                    ZGn  Ground distance zone n tripped.
                    ZPn  Phase distance zone n tripped.
                    TOC=value Time overcurrent element operating quantity in secondary amps
                    IOC=value Instantaneous overcurrent element operating quantity in secondary amps
        """
        self.__pick__()
        # check obj
        if type(obj) not in {FUSE,RLYOCG,RLYOCP,RLYDSG,RLYDSP,RECLSR,RLYD,RLYV,SCHEME} or\
              type(mult) not in {float, int} or signalonly not in {0,1}:
            se = '\nRESULT_FLT.optime(obj,mult,signalonly)'
            se += '\n\t             Required  Found'
            se +='\n\tobj          RELAY     '+type(obj).__name__.ljust(8) +' with RELAY=FUSE,RLYOCG,RLYOCP,RLYDSG,RLYDSP,RECLSR,RLYD,RLYV'
            se +='\n\tmult         float     '+str(mult).ljust(8)
            se +=' Relay current multiplying factor'
            se +='\n\tsignalonly   0,1       ' +str(signalonly).ljust(8)
            se +=' Consider relay element signal-only flag'
            raise ValueError(se)
        #
        sx = create_string_buffer(b'\000' * 128)
        triptime = c_double(0.0)
        if OlxAPIConst.OLXAPI_OK != OlxAPI.GetRelayTime(obj.__hnd__,c_double(mult), c_int(signalonly), byref(triptime),sx):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #
        return triptime.value, (sx.value).decode("UTF-8")
#
FltSimResult = []
#
class OUTAGE:
    def __init__(self,option,G=0):
        """ Construct equipment outage contingencies
        -option: (str) Branch outage option code
            'SINGLE'     : one at a time
            'SINGLE-GND' : (LINE only) one at a time with ends grounded
            'DOUBLE'     : two at a time
            'ALL'        : all at once
        -G: Admittance Outage line grounding (mho) (option=SINGLE-GND)
        """
        self.outageLst = []
        self.option = option
        self.G = G
        self.__check1__()
    #
    def __check1__(self):
        se = '\nOUTAGE(option,G)'
        if type(self.option)!=str or self.option.upper() not in {'SINGLE','SINGLE-GND','DOUBLE','ALL'}:
            se+= '\n  option : (str)Branch outage option code'
            se+= '\n\tRequired: (str)'
            se+= '\n\t        SINGLE     : one at a time'
            se+= '\n\t        SINGLE-GND : (LINE only) one at a time with ends grounded'
            se+= '\n\t        DOUBLE     : two at a time'
            se+= '\n\t        ALL        : all at once'
            se+= '\n\t' +__getErrValue__(int,self.option)
            raise ValueError(se)
        if (type(self.G) not in {float,int} or self.G<0) and self.option.upper()=='SINGLE-GND':
            se+= '\n  G : Admittance Outage line grounding (mho) (option=SINGLE-GND)'
            se+= '\n\tRequired: >0'
            se+= '\n\t'+__getErrValue__(float,self.G)
            raise ValueError(se)
    #
    def __check2__(self,obj,tiers,wantedTypes):
        se ='\nOUTAGE.build_outageLst(obj,tiers,wantedTypes)'
        if obj is None or type(obj) not in {BUS,RLYGROUP,LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH,TERMINAL}:
            se+= '\nobj : object for Outage option'
            se+= '\n\tRequired: BUS or TERMINAL or RLYGROUP or EQUIPMENT (LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH)'
            se+= '\n\t'+__getErrValue__('',obj)
            raise ValueError(se)
        __check_currFileIdx__(obj)
        if tiers is None or type(tiers)  not in {float,int} or tiers<0:
            se+= '\n  tiers : Number of tiers-neighboring'
            se+= '\n\tRequired: (int)>0'
            se+= '\n\t'+__getErrValue__(int,tiers)
            raise ValueError(se)
        #
        flag = type(wantedTypes)==list
        if flag:
            for w1 in wantedTypes:
                if type(w1)!=str or w1 not in {'LINE','XFMR','XFMR3','SHIFTER','SWITCH'}:
                    flag = False
        if not flag or wantedTypes is None  :
            se+= '\n  wantedTypes: Branch type to consider'
            se+= "\n\tRequired: (list)[(str/obj)] in ['LINE','XFMR','XFMR3','SERIESRC','SHIFTER','SWITCH']"
            se+= '\n\tFound (ValueError) :'+str(wantedTypes)
            raise ValueError(se)
    #
    def add_outageLst(self,la):
        se = '\nOUTAGE.add_outageLst(la)'
        if type(la)==list:
            for l1 in la:
                if type(l1) not in {LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH}:
                    se+= "\n\tRequired: (list)[(str/obj)] in [LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH]"
                    se+= '\n\t'+__getErrValue__(list,la)
                    raise ValueError(se)
                __check_currFileIdx__(l1)
                if not l1.isInList(self.outageLst):
                    self.outageLst.append(l1)
            return
        if type(la) not in {LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH}:
            se+= "\n\tRequired: (list)[(str/obj)] in [LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH]"
            se+= '\n\t'+__getErrValue__(None,la)
            raise ValueError(se)
        __check_currFileIdx__(la)
        if not la.isInList(self.outageLst):
            self.outageLst.append(la)
    #
    def build_outageLst(self,obj,tiers,wantedTypes):
        """ Return list of neighboring branches
        Args:
        -obj: BUS or TERMINAL or RLYGROUP or EQUIPMENT (LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH)
        -tiers: Number of tiers-neighboring (must be positive)
        -wantedTypes [(str/obj)]: Branch type to consider
                ['LINE','XFMR','XFMR3','SERIESRC','SHIFTER','SWITCH'] """
        wantedTypes1 = []
        if type(wantedTypes)==list:
            for w1 in wantedTypes:
                if type(w1)==str:
                    wantedTypes1.append(w1.upper())
                else:
                    try:
                        wantedTypes1.append(w1.__name__)
                    except:
                        wantedTypes1.append(w1)
        else:
            wantedTypes1 = wantedTypes
        self.__check1__()
        self.__check2__(obj,tiers,wantedTypes1)
        #
        vd = {'LINE':1,'XFMR':2,'XFMR3':8,'SHIFTER':4,'SWITCH':16}
        if len(wantedTypes1)==0:
            vw = 1+2+4+8+16
        else:
            wa = set()
            vw = 0
            for w1 in wantedTypes1:
                wa.add(w1.upper())
            for w1 in wa:
                vw += vd[w1]
        #
        wantedT = c_int(vw)
        listLen = c_int(0)
        hnd = obj.__hnd__
        OlxAPI.MakeOutageList(hnd, c_int(tiers-1), wantedT, None, pointer(listLen) )
        branchList = (c_int*(5+listLen.value))(0)
        OlxAPI.MakeOutageList(hnd, c_int(tiers-1), wantedT, branchList, pointer(listLen) )
        for i in range(listLen.value):
            r1 = __getOBJ__(branchList[i])
            if not r1.isInList(self.outageLst):
                self.outageLst.append(r1)
    #
    def toString(self):
        res = '{option='+str(self.option)
        if self.option.upper()=='SINGLE-GND':
            res += ',G='+str(self.G)
        res += ',List=['
        for o1 in self.outageLst:
            res+=type(o1).__name__ +','
        res = res[:-1] +'](len=%i)}'%len(self.outageLst)
        return res
#
class SPEC_FLT:
    def __init__(self,para,typ):
        super().__setattr__('__paraInput__', para)
        super().__setattr__('__type__', typ)
    #
    def __setattr__(self, sPara, value):
        sPara1 = sPara
        for k in self.__paraInput__.keys():
            if k.upper()==sPara.upper():
                self.__paraInput__[k] = value
                sPara1 = k
                break
        #
        if self.__type__=='Classical':
            if sPara1 not in ['obj','fltApp','fltConn','Z','outage']:
                se = "\nSPEC_FLT.Classical(obj,runOpt,fltConn,Z,fParam,outageOpt,outageLst)\n  has no attribute '%s'"%sPara
                se += '\n  All attributes available:'
                se += '\n\tobj       : (obj) Classical fault object: BUS or RLYGROUP or TERMINAL'
                se += "\n\tfltApp    : (str) Fault application: 'Bus','Close-in','Close-in-EO','Remote-bus','Line-end','xx%','xx%-EO'"
                se += '\n\tfltConn   : (str) Fault connection code:'+str(__OLXOBJ_fltConn__[:7])[:-1]
                se += ',\n                                            '+str(__OLXOBJ_fltConn__[7:])[1:]
                se += '\n\tZ         : ([R,X]) Fault impedances in Ohm'
                se += '\n\toutage    : (OUTAGE) outage option'
                raise AttributeError(se)
        elif self.__type__=='Simultaneous': # obj,fltApp,fltConn,Z
            if sPara1 not in ['obj','fltApp','fltConn','Z']:
                se = "\nSPEC_FLT.Simultaneous(obj,fltApp,fltConn,Z)\n  has no attribute '%s'"%sPara
                se += '\n  All attributes available:'
                se += '\n\tobj       : (obj) Simultaneous fault object: BUS,[BUS,BUS],RLYGROUP,TERMINAL'
                se += "\n\tfltApp    : (str) Simulation option code: 'BUS','CLOSE-IN','BUS2BUS','LINE-END','OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN','xx%'"
                se += '\n\tfltConn   : (str) Fault connection code:'+str(__OLXOBJ_fltConn__[:7])[:-1]
                se += ',\n                                            '+str(__OLXOBJ_fltConn__[7:])[1:]
                se += '\n\tZ         : ([]) Fault impedances in Ohm'
                raise AttributeError(se)

        elif self.__type__=='SEA':
            if sPara1 not in ['obj','runOpt','fltConn','Z','fParam','addParam']:
                se = "\nSPEC_FLT.SEA(obj,fltConn,runOpt,tiers,Z,fParam,addParam)\n  has no attribute '%s'"%sPara
                se += '\n  All attributes available:'
                se += '\n\tobj       : (obj) SEA object'
                se += '\n\tfltConn   : (str) Fault connection code'
                se += '\n\trunOpt    : ([]) simulation option falg'
                se += '\n\tZ         : ([]) Fault impedances in Ohm'
                se += '\n\tfParam    : (float) Intermediate fault location in %'
                se += '\n\taddParam  : ([]) additionnal event for Multiple User-Defined Event'
                raise AttributeError(se)
    #
    def __getattr__(self, sPara):
        for k in self.__paraInput__.keys():
            if k.upper()==sPara.upper():
                return self.__paraInput__[k]
        if self.__type__=='Classical':
            se = "\nSPEC_FLT.Classical has no attribute '%s'"%sPara
        elif self.__type__=='Simultaneous':
            se = "\nSPEC_FLT.Simultaneous has no attribute '%s'"%sPara
        elif self.__type__=='SEA':
            se = "\nSPEC_FLT.SEA has no attribute '%s'"%sPara
        raise AttributeError(se)
    #
    def Classical(obj,fltApp,fltConn,Z=None,outage=None):
        """ Define a specfication for classical fault
        Args = (obj,fltApp,fltConn,Z,outage)
        -obj: Classical fault object
            BUS or RLYGROUP or TERMINAL
        -fltApp: (str) Fault application
            'Bus'
            'Close-in'
            'Close-in-EO' (EO: End Opened)
            'Remote-bus'
            'Line-end'
            'xx%'         (xx:Intermediate fault location %)
            'xx%-EO'
        -fltConn: (str) Fault connection code
            '3LG'
            '2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA'
            '1LG:A' ,'1LG:B' ,'1LG:C'
            'LL:BC' ,'LL:CA' ,'LL:AB' ,'LL:CB' ,'LL:AC' ,'LL:BA'
        -Z : [R,X] Fault Impedance (Ohm)
        -outage : (OUTAGE) outage option
        """
        #
        para = dict()
        para['obj'] = obj
        para['fltApp'] = fltApp
        para['fltConn'] = fltConn
        para['Z'] = Z
        para['outage'] = outage
        #
        return SPEC_FLT(para,'Classical')
    #
    def Simultaneous(obj,fltApp,fltConn,Z=None):
        """Define a specfication for Simultaneous fault:
        Args = (obj,fltApp,fltConn,Z)
        -obj: Simultaneous fault object
            BUS       : fltApp='Bus'
            [BUS,BUS] : fltApp='Bus2Bus'
            RLYGROUP or TERMINAL (on a LINE)      : fltApp='xx%' (with xx: Intermediate fault location)
            RLYGROUP or TERMINAL : fltApp='Close-In','Line-End','Outage','1P-Open','2P-Open','3P-Open'
        -fltApp: (str) Fault application code
            'Bus'
            'Close-In'
            'Bus2Bus'
            'Line-End'
            'xx%'           (xx:Intermediate fault location %)
            'Outage'
            '1P-Open','2P-Open','3P-Open'
        -fltConn: (str) Fault connection code
            '3LG'
            '2LG:BC','2LG:CA','2LG:AB'
            '1LG:A' ,'1LG:B' ,'1LG:C'
            'LL:BC' ,'LL:CA' ,'LL:AB'
            'AA','BB','CC','AB','AC','BC','BA','CA','CB' : for BUS2BUS
            'A','B','C'                   : for 1P-Open
            'AB','AC','BC'                : for 2P-Open
        -Z: Fault Impedance (Ohm)
            [RA,XA,RB,XB,RC,XC,RG,XG]: runOpt!='Bus2Bus'
            [R,X] : fltApp ='Bus2Bus'
        """
        #
        para = dict()
        para['obj'] = obj
        para['fltApp'] = fltApp
        para['fltConn'] = fltConn
        para['Z'] = Z
        #
        return SPEC_FLT(para,'Simultaneous')
    #
    def SEA_EX(time,fltConn,Z):
        """ Stepped-Event Analysis with Additional Event
        Args = (time,obj,fltApp,fltConn,Z)
        - time: [s] time of Additional Event
        - fltConn: (str) Fault connection code
            '3LG'
            '2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA'
            '1LG:A' ,'1LG:B' ,'1LG:C'
            'LL:BC' ,'LL:CA' ,'LL:AB' ,'LL:CB' ,'LL:AC' ,'LL:BA'
        - Z: [R,X]: Fault Impedance (Ohm)
        """
        para = dict()
        para['time'] = time
        para['fltConn'] = fltConn
        para['Z'] = Z
        return SPEC_FLT(para,'SEA_EX')
    #
    def SEA(obj,fltApp,fltConn,deviceOpt,tiers,Z=None):
        """ Stepped-Event Analysis
        Args = (obj,fltApp,fltConn,deviceOpt,tiers,Z)
        - obj: Stepped-Event Analysis object
            BUS or RLYGROUP or TERMINAL
        - fltApp: (str) Fault application code
            'Bus'
            'Close-In'
            'xx%' (xx: Intermediate fault location %)
        - fltConn: (str) Fault connection code
            '3LG'
            '2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA'
            '1LG:A' ,'1LG:B' ,'1LG:C'
            'LL:BC' ,'LL:CA' ,'LL:AB' ,'LL:CB' ,'LL:AC' ,'LL:BA'
        - Z: [R,X]: Fault Impedance (Ohm)
        - deviceOpt: []*7 simulation options
            deviceOpt[0] = 1/0 Consider OCGnd operations
            deviceOpt[1] = 1/0 Consider OCPh operations
            deviceOpt[2] = 1/0 Consider DSGnd operations
            deviceOpt[3] = 1/0 Consider DSPh operations
            deviceOpt[4] = 1/0 Consider Protection scheme operations
            deviceOpt[5] = 1/0 Consider Voltage relay operations
            deviceOpt[6] = 1/0 Consider Differential relay operations
        - tiers: Study extent. Take into account protective devices located within this number of tiers only.
        """
        para = dict()
        para['obj'] = obj
        para['fltApp'] = fltApp
        para['fltConn'] = fltConn
        para['deviceOpt'] = deviceOpt
        para['tiers'] = tiers
        para['Z'] = Z
        return SPEC_FLT(para,'SEA')
    #
    def checkData(self):
        if self.__type__=='Classical':
            __checkClassical__(self)
        elif self.__type__=='Simultaneous':
            __checkSimultaneous__(self)
        elif self.__type__=='SEA':
            __checkSEA__(self)
        elif self.__type__=='SEA_EX':
            __checkSEA_EX__(self)
    #
    def getData(self):
        if self.__type__=='Classical':
            para1 = __getClassical__(self)
        elif self.__type__=='Simultaneous':
            para1 = __getSimultaneous__(self)
        elif self.__type__=='SEA':
            para1 = __getSEA__(self)
        elif self.__type__=='SEA_EX':
            para1 = __getSEA_EX__(self)
        else:
            raise Exception('Error type of SPEC_FLT')
        return para1
    #
    def toString(self):
        try:
            obj = self.__paraInput__['obj']
            sp = 'obj='+type(obj).__name__
        except:
            sp = ''
        for k,v in self.__paraInput__.items():
            if sp:
                sp+=', '
            if k not in {'obj','outage'}:
                if k=='Z' and v is None:
                    sp += k+'=[0]*'
                else:
                    sp += k+'='+str(v).upper()
            if k=='outage' and v is not None:
                sp += 'outage='+v.toString()
        return self.__type__+': '+sp
#
class __DATAABSTRACT__:
    def __init__(self,hnd):
        super().__setattr__('__hnd__', hnd)
        super().__setattr__('__currFileIdx__',__CURRENT_FILE_IDX__)
        #
        va1,va2 = set(),set()
        flag1,flag2 = True,False
        for v1 in dir(self):
            if v1.startswith('__') and v1.endswith('__'):
                flag1 = False
            elif not flag1:
                flag2 = True
            if flag1:
                va1.add(v1)
            if flag2:
                va2.add(v1)
        #
        if type(self) in {GENUNIT,SHUNTUNIT,LOADUNIT}:
            va1.discard('MEMO')
        if type(self)==TERMINAL:
            for v1 in ['GUID','TAGS','JRNL','MEMO']:
                va1.discard(v1)
        #
        va1 = list(va1)
        va1.sort()
        va2.discard('init')
        va2 = list(va2)
        va2.sort()
        super().__setattr__('__paraUDF__', None)
        super().__setattr__('__allAttributes__', va1)
        super().__setattr__('__allMethods__', va2)
    #
    def __getName_udf__(self):
        if self.__paraUDF__ is None:
            paraUDF = []
            if type(self)!=DCLINE2: # bug udf dcline
                for idx in range(200):
                    fname= create_string_buffer(b'\000' * OlxAPIConst.MXUDFNAME)
                    fval = create_string_buffer(b'\000' * OlxAPIConst.MXUDF)
                    try:
                        if OlxAPIConst.OLXAPI_OK == OlxAPI.GetObjUDFByIndex(self.__hnd__,idx,fname,fval):
                            paraUDF.append(OlxAPI.decode(fname.value))
                        else:
                            break
                    except:
                        pass
            super().__setattr__('__paraUDF__', paraUDF)
            self.__allAttributes__.extend(paraUDF)
    #
    def __setattr__(self, sPara, value):
        """
        change data for object
            sPara = str of parameter
            value : new value
        """
        self.changeData(sPara, value)
    #
    def __getattr__(self, name):
        return self.getData(name)
    #
    @property
    def HANDLE(self):
        """ (int) return handle of object"""
        return self.getData('HANDLE')
    @property
    def JRNL(self):
        """ (str) Creation and modification log records"""
        return self.getData('JRNL')
    @property
    def GUID(self):
        """ (str) GUID for object"""
        return self.getData('GUID')
    @property
    def MEMO(self):
        """ (str) Memo for object (Line breaks must be included in the memo string as escape charater \n) """
        return self.getData('MEMO')
    @property
    def TAGS(self):
        """ (str) Tags for object """
        return self.getData('TAGS')
    #
    def equals(self,o2):
        """ return (bool) comparison of 2 objects """
        try:
            return self.__hnd__==o2.__hnd__
        except:
            return False
    #
    def isInList(self,la):
        """ check if object in in List/Set of object """
        __check_currFileIdx__(self)
        if type(la) in {list,set}:
            for li in la:
                __check_currFileIdx__(li)
                try:
                    if self.__hnd__==li.__hnd__:
                        return True
                except:
                    pass
        return False
    #
    def changeData(self,sPara, value):
        """ change Data method """
        __check_currFileIdx__(self)
        #
        if type(sPara)==list:
            for i in range(len(sPara)):
                self.changeData(sPara[i],value[i])
            return
        #
        sname = type(self).__name__
        hnd = self.__hnd__
        if type(sPara)!=str :
            se = "\nin call %s.changeData(sPara,value)"%sname
            se+= "\n\ttype(sPara)    : str"
            se+= "\n\t"+__getErrValue__(str,sPara)
            raise ValueError(se)
        #
        super().__setattr__('__spara__', sPara)
        sPara =  sPara.upper()
        # GUID/JournalRecord
        if sPara in {'GUID','JRNL','HANDLE'}:
            raise Exception('\nWrite Access = NO with %s.%s'%(sname,self.__spara__))

        try:
            paraCode = __OLXOBJ_PARA__[sname][sPara][0]
        except:
            if type(self)==RECLSR:
                try:
                    if sPara.startswith('PH_') :
                        paraCode = __OLXOBJ_PARA__[sname][sPara[3:]][0]
                    elif sPara.startswith('GR_') :
                        paraCode = __OLXOBJ_PARA__[sname][sPara[3:]][1]
                    else:
                        paraCode = 0
                except:
                    paraCode = 0
            else:
                paraCode = 0
        t0 = __getTypeParaCode__(paraCode)
        #
        if paraCode>0:
            if (t0 != type(value)) and not (t0==float and type(value)==int):
                se = "\nin call %s.%s = "%(sname,sPara)
                se+= ("'"+str(value)+"'") if type(value)==str else str(value)
                se+= "\n\ttype(value) required : %s"%t0.__name__
                se+= "\n\tfound                : "+type(value).__name__
                raise TypeError(se)
            val1 = __setValue__(hnd,paraCode,value)
            try:
                if OlxAPIConst.OLXAPI_OK == OlxAPI.SetData( c_int(hnd), c_int(paraCode), byref(val1) ):
                    return
            except:
                pass
            raise Exception('\nWrite Access = NO with %s.%s'%(sname,sPara))
        # Memo
        if sPara=='MEMO':
            if type(self) in {GENUNIT,SHUNTUNIT,LOADUNIT}:
                __errorsPara__(self,sPara)
            #
            if type(self).__name__!=RECLSR:
                if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.SetObjMemo(hnd,value):
                    se = OlxAPI.ErrorString()
                    if se=='SetObjMemo failure: Invalid Device Type':
                        raise Exception('\nWrite Access = NO with '+ sname+'.'+sPara)
                    else:
                        raise OlxAPI.OlxAPIException('\n'+self.toString()+ '\n'+se)
                return
            if type(value)!=list or len(value)!=2:
                se = "\nin call %s.%s = "%(sname,str(sPara)) + str(value)
                se+= "\n\ttype(value) required : [memoP,memoG]"
                se+= "\n\tfound                : "+str(value)
                raise TypeError(se)
            #
            if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.SetObjMemo(hnd,str(value[0])): # RECLSRP
                se = OlxAPI.ErrorString()
                if se=='SetObjMemo failure: Invalid Device Type':
                    raise Exception('\nWrite Access = NO with '+ sname+'.'+sPara)
                else:
                    raise OlxAPI.OlxAPIException('\n'+self.toString()+ '\n'+se)
            if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.SetObjMemo(hnd+1,str(value[1])): # RECLSRG
                se = OlxAPI.ErrorString()
                if se=='SetObjMemo failure: Invalid Device Type':
                    raise Exception('\nWrite Access = NO with '+ sname+'.'+sPara)
                else:
                    raise OlxAPI.OlxAPIException('\n'+self.toString()+ '\n'+se)
            return
        # Tags
        if sPara=='TAGS':
            if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.SetObjTags(hnd,value):
                raise OlxAPI.OlxAPIException('\n'+self.toString()+ '\n'+OlxAPI.ErrorString())
            return
        # UDF
        self.__getName_udf__()
        if sPara in self.__paraUDF__:
            if type(value)!=str:
                se = "\nin call %s.%s = "%(sname,str(sPara)) + str(value)
                se+= "\n\ttype(value) required : str"
                se+= "\n\tfound                : "+type(value).__name__
                raise TypeError(se)
            if OlxAPIConst.OLXAPI_OK==OlxAPI.SetObjUDF(hnd, sPara, value):
                return
            se = "\nError in call %s.%s = "%(sname,str(sPara)) + str(value)
            raise Exception(se)
        # error
        __errorsPara__(self,sPara)
    #
    def getData(self,sPara=None):
        """
        return data of object
            sPara = None or str or list(str) parameter
                    try ='' to print in console all sPara avalaible
        """
        __check_currFileIdx__(self)
        #
        if sPara is None:
            sPara = self.getAttributes()
            return dict(zip(sPara, self.getData(sPara)))
        #
        if type(sPara) in {list,set,tuple}:
            return [self.getData(sp1) for sp1 in sPara]
        #
        if type(sPara)!=str:
            se = "\nin call %s.getData(sPara)"%type(self).__name__
            se+= "\n\tRequired sPara     : None or str or list/set/tuple of str"
            se+= "\n\t"+__getErrValue__(str,sPara)
            raise TypeError(se)
        #
        hnd = self.__hnd__
        super().__setattr__('__spara__', sPara)
        sPara = sPara.upper()
        if sPara=='HANDLE':
            return hnd
        #
        typ1 = type(self)
        if typ1==RLYGROUP:
            if sPara in {'BACKUP','PRIMARY'} :
                res = []
                val1 = c_int(0)
                paraCode = __OLXOBJ_PARA__['RLYGROUP'][sPara][0]
                while OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(hnd,c_int(paraCode),byref(val1)):
                    res.append(RLYGROUP(hnd=val1.value))
                return res
            if sPara=='TERMINAL':
                return TERMINAL(hnd=__getDatai__(hnd,OlxAPIConst.RG_nBranchHnd))
        #
        if typ1==RECLSR:
            p1,p2 = sPara[3:],sPara[:3]
            if p2=='GR_':
                paraCode = __OLXOBJ_PARA__['RECLSR'][p1][1]
                res = __getData__(hnd+1,paraCode)
                return __getOBJ__(res,sPara=sPara)
            if p2=='PH_':
                paraCode = __OLXOBJ_PARA__['RECLSR'][p1][0]
                res = __getData__(hnd,paraCode)
                return __getOBJ__(res,sPara=sPara)
        #
        try:
            paraCode = __OLXOBJ_PARA__[typ1.__name__][sPara][0]
            if paraCode!=0:
                res = __getData__(hnd,paraCode)
##                if typ1==BUS and sPara=='SLACK':# System slack bus flag: 1-yes; 0-no
##                    return 1 if res>0 else 0
                return __getOBJ__(res,sPara=sPara)
        except:
            pass
        if sPara=='JRNL':
            return OlxAPI.GetObjJournalRecord(hnd)
        if sPara=='GUID':
            return OlxAPI.GetObjGUID(hnd)
        if sPara=='MEMO':
            s1 = OlxAPI.GetObjMemo(hnd)
            if s1=='GetObjMemo failure: Invalid Device Type':
                s1= ''
            if typ1==RECLSR:
                s2 =  OlxAPI.GetObjMemo(hnd+1)
                if s2=='GetObjMemo failure: Invalid Device Type':
                    s2= ''
                return [s1,s2]
            else:
                return s1
        if sPara=='TAGS':
            return OlxAPI.GetObjTags(hnd)
        #
        if sPara not in self.getAttributes():
            return __errorsPara__(self,sPara)
        #
        if sPara in self.__paraUDF__: # UDF
            fval = create_string_buffer(b'\000' * OlxAPIConst.MXUDF)
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetObjUDF(hnd, sPara, fval) :
                return OlxAPI.decode(fval.value)
        #
        return getattr(self, sPara)
    #
    def postData(self):
        """ Perform validation and update object data in the network database """
        __check_currFileIdx__(self)
        if OlxAPIConst.OLXAPI_OK != OlxAPI.PostData(c_int(self.__hnd__)):
            raise OlxAPI.OlxAPIException('\n'+self.toString()+ '\n'+OlxAPI.ErrorString())
        return True
    #
    def getAttributes(self):
        """ [str] list of all attributes """
        __check_currFileIdx__(self)
        self.__getName_udf__()
        return self.__allAttributes__
    #
    def toString(self,option=0):
        """ (str) text description/composed of object
        option = 0 (default)
            return (str) text composed for object
                BUS   : name and kV
                EQUIPMENT: Bus, Bus2, (Bus3), Circuit ID and type of a branch
                RELAY : relay type, name and branch location
        option = 1
            return (str) text description of object
        option = 10:
            str all attribute of object with comment of attribute
        option = 11:
            str all attribute of object without comment of attribute, without JRNL
        """
        __check_currFileIdx__(self)
        typ = type(self)
        if option==0:
            s1 = OlxAPI.PrintObj1LPF(self.__hnd__)
            if typ!=TERMINAL:
                return s1
            # Recompose the obj string:
            #   Replace object type inside [] with TERMINAL
            #   and append the branch code letter to the end
            id1 = s1.find('[')
            id2 = s1.find(']')
            if id1>=0 and id2>=0:
                s1 = s1[:id1+1]+'TERMINAL'+s1[id2:]
                tc1 =  __getDatai__(self.__hnd__,OlxAPIConst.BR_nType)
                dictL = {OlxAPIConst.TC_LINE:' L',OlxAPIConst.TC_DCLINE2:' DC',OlxAPIConst.TC_SCAP:' S',OlxAPIConst.TC_PS:' P',\
                         OlxAPIConst.TC_XFMR:' T', OlxAPIConst.TC_XFMR3:' X',OlxAPIConst.TC_SWITCH:' W'}
                s1 += dictL[tc1]
            return s1
        if option==1:
            if typ==BUS:
                return OlxAPI.FullBusName(self.__hnd__)
            elif typ in __OLXOBJ_EQUIPMENTO__:
                return OlxAPI.FullBranchName(self.__hnd__)
            elif typ in __OLXOBJ_RELAYO__ or typ==RLYGROUP:
                return OlxAPI.FullRelayName(self.__hnd__)
            s1 = OlxAPI.PrintObj1LPF(self.__hnd__)
            id2 = s1.find(']')
            return s1[id2+1:]
        # all data available for Object
        if option == 10 or option== 11:
            sres = '\n'+self.toString()
            vals = self.getData()
            if option== 11:
                del vals['HANDLE']
                try:
                    del vals['JRNL']
                except:
                    pass
            for k,v in vals.items():
                if type(v)==list:
                    sres +='\n'+ k.ljust(15) +' : ' +toString(v)
            #
            for k,v in vals.items():
                if type(v)!=list:
                    s1 = k.ljust(15) +' : ' +toString(v).ljust(15)
                    s2=''
                    if option==10:
                        try:
                            s2 = __OLXOBJ_PARA__[type(self).__name__][k][1]
                        except:
                            s2 = ''
                        #
                        if type(self) ==RECLSR:
                            p2 = k[:3]
                            if p2=='GR_':
                                s2 += __OLXOBJ_PARA__[type(self).__name__][k[3:]][3]
                            elif p2=='PH_':
                                s2 += __OLXOBJ_PARA__[type(self).__name__][k[3:]][2]
                    if len(s1)<37:
                        sres +='\n'+s1+'   '+s2
                    else:
                        if s2:
                            sres +='\n'+s1+'\n                                    '+s2
                        else:
                            sres +='\n'+s1
            #
            if type(self) in {RLYOCG,RLYOCP,RLYDSG,RLYDSP}:
                sres+=__printRLYSetting__(self,cmt=(option==10))
            return sres
        raise ValueError ('object.toString(option) with option =0,1,10,11\n\tFound option='+str(option))
#
class NETWORK:
    def __init__(self,val1=None,val2=None):
        if val1 is not None or val2 is not None:
            se = '\nThe NETWORK class can no longer be instantiated outside the OlxObj module version 2.1.4 or above.'
            se+= '\nYou must use the global object OLCase instead.'
            raise Exception(se)
        global __K_INI_NETWORK__
        __K_INI_NETWORK__ +=1
        if __K_INI_NETWORK__>1:
            raise Exception('\nCan not init class NETWORK(), try to use OLRCase.')
        #
        va1,va2 = [],[]
        flag1,flag2 = True,False
        for v1 in dir(self):
            if v1.startswith('__') and v1.endswith('__'):
                flag1 = False
            elif not flag1:
                flag2 = True
            if flag1:
                va1.append(v1)
            if flag2:
                va2.append(v1)
        va1.sort()
        va2.sort()
        super().__setattr__('__allAttributes__',va1)
        super().__setattr__('__allMethods__',va2)
        super().__setattr__('__scope__', {'isFullNetWork':True})
    #
    def open(self,olrFile,readonly):
        super().__setattr__('olrFile',olrFile)
        try:
            OlxAPI.CloseDataFile()
        except:
            pass
        #
        if type(olrFile)!=str or not os.path.isfile(olrFile):
            raise ValueError("File not found: "+str(olrFile))
        if not olrFile.upper().endswith('.OLR'):
            raise ValueError("\nError file format (.OLR) : "+olrFile)
        #
        if OlxAPI.ASPENOlxAPIDLL is None:
            raise Exception('\nNot yet init OlxAPI, try to load_olxapi(dllPath) before to open a NetWork')
        val = OlxAPI.LoadDataFile(olrFile,True if readonly else False)
        if OlxAPIConst.OLXAPI_OK == val:
            print( "File opened successfully: " + olrFile)
        elif OlxAPIConst.OLXAPI_DATAFILEANOMALIES==val:
            print( "File opened successfully: " + olrFile)
            print(OlxAPI.ErrorString())
        else:
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        #
        global __CURRENT_FILE_IDX__,__INDEX_SIMUL__,__COUNT_FAULT__,FltSimResult
        __CURRENT_FILE_IDX__ = abs(__CURRENT_FILE_IDX__) + 1
        __INDEX_SIMUL__ +=1
        __COUNT_FAULT__ = 0
        FltSimResult.clear()
        #
        super().__setattr__('__currFileIdx__',__CURRENT_FILE_IDX__)
        super().__setattr__('__scope__', {'isFullNetWork':True})
        super().__setattr__('__allBus__',None)
        base = __getData__(OlxAPIConst.HND_SYS,OlxAPIConst.SY_dBaseMVA)
        super().__setattr__('__basemva__',base)
    #
    def __getattr__(self, name):
        return self.getData(name)
    #
    def close(self):
        """ Close the network data file that had been loaded previously return 0 if OK """
        __check_currFileIdx__(self)
        if OlxAPIConst.OLXAPI_OK ==OlxAPI.CloseDataFile():
            global __CURRENT_FILE_IDX__
            if __CURRENT_FILE_IDX__>0:
                __CURRENT_FILE_IDX__ *=-1
            return 0
        return 1
    #
    def save(self,fileNew=''):
        """
        Save ASPEN OLR data file
            fileNew : name path of new file
                if fileNew=='' (default) => save current file
        """
        __check_currFileIdx__(self)
        if fileNew:
            self.olrFile = os.path.abspath(fileNew)
        #
        if OlxAPIConst.OLXAPI_FAILURE == OlxAPI.SaveDataFile(self.olrFile):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
            return 1
        print( "File saved successfully: " + self.olrFile)
        return 0
    #
    def findOBJ(self,objStr):
        """
        find object by GUID or String produced by toString(0)
            objStr = GUID or STR
            return None if not found
        """
        __check_currFileIdx__(self)
        if type(objStr)==list:
            res = []
            for o1 in objStr:
                res.append( self.findObj(o1) )
            return res
        #
        if type(objStr)!=str:
            return None
        #
        hnd = c_int(0)
        if OlxAPIConst.OLXAPI_FAILURE == OlxAPI.FindObj1LPF(objStr,hnd):
            return None
        #
        tc = OlxAPI.EquipmentType(hnd)
        return __getOBJ__(hnd.value,tc=tc)
    #
    def getAttributes(self):
        """ [str] list of all attributes """
        __check_currFileIdx__(self)
        return self.__allAttributes__
    #
    def getData(self,sPara=None):
        """
        return data of NETWORK by sPara
        """
        __check_currFileIdx__(self)
        if sPara is None:
            sPara = self.getAttributes()
            return dict(zip(sPara, self.getData(sPara)))
        if type(sPara) in {list,set,tuple}:
            return [self.getData(sp1) for sp1 in sPara]
        #
        if type(sPara)!=str:
            se = "\nin call OLCase.getData(sPara)"
            se+= "\n\tRequire  sPara: None or str or list(str)"
            se+= "\n\t"+__getErrValue__(str,sPara)
            raise TypeError(se)
        #
        super().__setattr__('__spara__', sPara)
        sPara = sPara.upper()
        #
        if sPara in __OLXOBJ_CONST__.keys():
            return __getEquipment__(sPara,self.__scope__)
        #
        if sPara=='OBJCOUNT':
            res = {}
            for sPara in __OLXOBJ_PARA__['SYS'].keys():
                paraCode = __OLXOBJ_PARA__['SYS'][sPara][0]
                res[sPara] =  __getData__(OlxAPIConst.HND_SYS,paraCode)
            return res
        if sPara=='BASEMVA':
            return self.__basemva__
        if sPara=='COMMENT':
            return __getData__(OlxAPIConst.HND_SYS,OlxAPIConst.SY_sFComment)
        if sPara=='RLYOC':
            res = self.getData('RLYOCG')
            res.extend(self.getData('RLYOCP'))
            return res
        if sPara=='RLYDS':
            res = self.getData('RLYDSG')
            res.extend(self.getData('RLYDSP'))
            return res
        #
        if sPara not in self.getAttributes():
            se='\nAll methods for OLCase: '
            for sp in self.__allMethods__:
                se+=sp +'(),'
            se=se[:-1]+'\n\nAll attributes for OLCase:'
            for sp in self.__allAttributes__:
                if sp in __OLXOBJ_CONST__.keys():
                    se+= '\n'+sp.ljust(15)+' : '+__OLXOBJ_CONST__[sp][2] +' in NETWORK with Scope'
                elif sp in __OLXOBJ_PARA__['SYS2'].keys():
                    se+= '\n'+sp.ljust(15)+' : '+__OLXOBJ_PARA__['SYS2'][sp][1]
                else:
                    se+= '\n'+sp.ljust(15)
            se+= "\nAttributeError  : OLCase has no attribute/method '%s'" %self.__spara__
            raise AttributeError (se)
        return getattr(self, sPara)
    #
    def applyScope(self,areaNum=[],zoneNum=[],optionTie=0,kV=[]):
        """Scope for NETWORK Access
            areaNum  []: List of area Number
            zoneNum  []: List of zone Number
            optionTie  : 0-strictly in areaNum/zoneNum
                         1- with tie
                         2- only tie
            kV       []: [kVmin,kVmax]
        """
        self.__scope__ = dict()
        self.__scope__['areaNum'] = areaNum
        self.__scope__['zoneNum'] = zoneNum
        self.__scope__['optionTie'] = optionTie
        self.__scope__['kV'] = kV
        self.__scope__['isFullNetWork'] = not(areaNum or zoneNum or kV)
    #
    def resetScope(self):
        """
        reset Scope: All NETWORK Access
        """
        self.__scope__ = {'isFullNetWork':True}
    #
    def findBUS(self,val1=None,val2=None):
        """ find BUS (return None if not found):
                OLCase.findBUS(name,kV)    OLCase.findBUS('arizona',132) or BUS(132,'arizona')
                OLCase.findBUS(busNo)      OLCase.findBUS(28)
                OLCase.findBUS(GUID)       OLCase.findBUS('{b4f6d06c-473b-4365-ba05-a1b47a1f9364}'
                OLCase.findBUS(STR)        OLCase.findBUS("[BUS] 28 'ARIZONA' 132 kV") """
        try:
            return BUS(val1,val2)
        except:
            if not(((type(val1)==str or int) and val2 is None) or (type(val1)==str and (type(val2)==float or int))\
                or (type(val2)==str and (type(val1)==float or int))):
                    se='\nValueError OLCase.findBUS(...)'
                    se +='\n\tRequired: (name,kV) or (busNo) or (GUID) or (STR)'
                    se +='\n\tFound   : ('+type(val1).__name__
                    if val2 is not None:
                        se+= ','+type(val2).__name__
                    se +=')'
                    raise ValueError(se)
            __check_currFileIdx1__()
            return None
    #
    def findLINE(self,val1=None,val2=None,val3=None):
        """ find LINE (return None if not found):
                OLCase.findLINE(b1,b2,CID)   OLCase.findLINE(b1,b2,'1')
                OLCase.findLINE(GUID)        OLCase.findLINE('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                OLCase.findLINE(STR)         OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        try:
            return LINE(val1,val2,val3)
        except:
            #__checkArgsBr__('findLINE',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findDCLINE(self,val1=None,val2=None,val3=None):
        """ find DCLINE (return None if not found):
                OLCase.findDCLINE(b1,b2,CID)   OLCase.findDCLINE(b1,b2,'1')
                OLCase.findDCLINE(GUID)        OLCase.findDCLINE('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                OLCase.findDCLINE(STR)         OLCase.findDCLINE("[DCLINE2] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        try:
            return DCLINE2(val1,val2,val3)
        except:
            #__checkArgsBr__('findDCLINE',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findXFMR(self,val1=None,val2=None,val3=None):
        """ find XFMR (return None if not found):
                OLCase.findXFMR(b1,b2,CID)   OLCase.findXFMR(b1,b2,'1')
                OLCase.findXFMR(GUID)        OLCase.findXFMR('{e417041a-ef1b-46c2-82f2-9828d22d88b8}'
                OLCase.findXFMR(STR)         OLCase.findXFMR("[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1") """
        try:
            return XFMR(val1,val2,val3)
        except:
            #__checkArgsBr__('findXFMR',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findXFMR3(self,val1=None,val2=None,val3=None):
        """ find XFMR3 (return None if not found):
                OLCase.findXFMR3(b1,b2,CID)   OLCase.findXFMR3(b1,b2,'1')
                OLCase.findXFMR3(GUID)        OLCase.findXFMR3('{4a6b16d7-9469-4361-9d8e-60caf451f50a}'
                OLCase.findXFMR3(STR)         OLCase.findXFMR3("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-'DOT BUS' 13.8 kV 1") """
        try:
            return XFMR3(val1,val2,val3)
        except:
            #__checkArgsBr__('findXFMR3',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findSHIFTER(self,val1=None,val2=None,val3=None):
        """ find SHIFTER (return None if not found):
                OLCase.findSHIFTER(b1,b2,CID)   OLCase.findSHIFTER(b1,b2,'1')
                OLCase.findSHIFTER(GUID)        OLCase.findSHIFTER('{1bee3fcd-2663-4d2a-9bbe-9ff12f2240c1}'
                OLCase.findSHIFTER(STR)         OLCase.findSHIFTER("[SHIFTER] 4 'TENNESSEE' 132 kV-6 'NEVADA' 132 kV 1") """
        try:
            return SHIFTER(val1,val2,val3)
        except:
            #__checkArgsBr__('findSHIFTER',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findSERIESRC(self,val1=None,val2=None,val3=None):
        """ find SERIESRC (return None if not found):
                OLCase.findSERIESRC(b1,b2,CID)   OLCase.findSERIESRC(b1,b2,'1')
                OLCase.findSERIESRC(GUID)        OLCase.findSERIESRC('{1bee3fcd-2663-4d2a-9bbe-9ff12f2240c1}'
                OLCase.findSERIESRC(STR)         OLCase.findSERIESRC("[SERIESRC] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        try:
            return SERIESRC(val1,val2,val3)
        except:
            #__checkArgsBr__('findSERIESRC',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findSWITCH(self,val1=None,val2=None,val3=None):
        """ find SWITCH (return None if not found):
                OLCase.findSWITCH(b1,b2,CID)   OLCase.findSWITCH(b1,b2,'1')
                OLCase.findSWITCH(GUID)        OLCase.findSWITCH('{1bee3fcd-2663-4d2a-9bbe-9ff12f2240c1}'
                OLCase.findSWITCH(STR)         OLCase.findSWITCH("[SWITCH] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        try:
            return SWITCH(val1,val2,val3)
        except:
            #__checkArgsBr__('findSWITCH',val1,val2,val3)
            __check_currFileIdx1__()
            return None
    #
    def findTERMINAL(self,b1=None,b2=None,sType=None,CID=None):
        """ find TERMINAL (return None if not found):
            OLCase.findTERMINAL(b1,b2,sType,CID)   OLCase.findTERMINAL(b1,b2,'LINE','1')
                b1,b2 : (BUS)
                sType : (str/obj) 'XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH'
                CID   : (str) """
        sType = TERMINAL.__checkInputs__(b1,b2,sType,CID,'\nOLCase.findTERMINAL(b1,b2,sType,CID)')
        try:
            return TERMINAL(b1,b2,sType,CID)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findGEN(self,val1=None,val2=None):
        """ find GEN (return None if not found):
                OLCase.findGEN(GUID)        OLCase.findGEN('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findGEN(STR)         OLCase.findGEN("[GENERATOR] 2 'CLAYTOR' 132 kV")
                OLCase.findGEN(name,kV)     OLCase.findGEN('CLAYTOR',132)   """
        try:
            return GEN(val1,val2)
        except:
            __checkArgs1__('findGEN',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findGENW3(self,val1=None,val2=None):
        """ find GENW3 (return None if not found):
                OLCase.findGENW3(GUID)        OLCase.findGENW3('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findGENW3(STR)         OLCase.findGENW3("[GENW3] 'BUS3' 33 kV")
                OLCase.findGENW3(name,kV)     OLCase.findGENW3('CLAYTOR',132)   """
        try:
            return GENW3(val1,val2)
        except:
            __checkArgs1__('findGENW3',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findGENW4(self,val1=None,val2=None):
        """ find GENW4 (return None if not found):
                OLCase.findGENW4(GUID)        OLCase.findGENW4('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findGENW4(STR)         OLCase.findGENW4("[GENW4] 'BUS3' 33 kV")
                OLCase.findGENW4(name,kV)     OLCase.findGENW4('CLAYTOR',132)   """
        try:
            return GENW4(val1,val2)
        except:
            __checkArgs1__('findGENW4',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findCCGEN(self,val1=None,val2=None):
        """ find CCGEN (return None if not found):
                OLCase.findCCGEN(GUID)        OLCase.findCCGEN('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findCCGEN(STR)         OLCase.findCCGEN("[CCGENUNIT] 'BUS5' 13 kV")
                OLCase.findCCGEN(name,kV)     OLCase.findCCGEN('CLAYTOR',132)   """
        try:
            return CCGEN(val1,val2)
        except:
            __checkArgs1__('findCCGEN',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findSHUNT(self,val1=None,val2=None):
        """ find SHUNT (return None if not found):
                OLCase.findSHUNT(GUID)        OLCase.findSHUNT('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findSHUNT(STR)         OLCase.findSHUNT("[SHUNT] 21 'IOWA' 33 kV")
                OLCase.findSHUNT(name,kV)     OLCase.findSHUNT('CLAYTOR',132)   """
        try:
            return SHUNT(val1,val2)
        except:
            __checkArgs1__('findSHUNT',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findSVD(self,val1=None,val2=None):
        """ find SVD (return None if not found):
                OLCase.findSVD(GUID)        OLCase.findSVD('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findSVD(STR)         OLCase.findSVD("[SVD] 3 'TEXAS' 132 kV")
                OLCase.findSVD(name,kV)     OLCase.findSVD('TEXAS',132)   """
        try:
            return SVD(val1,val2)
        except:
            __checkArgs1__('findSVD',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findLOAD(self,val1=None,val2=None):
        """ find LOAD (return None if not found):
                OLCase.findLOAD(GUID)        OLCase.findLOAD('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                OLCase.findLOAD(STR)         OLCase.findLOAD("[LOAD] 17 'WASHINGTON' 33 kV")
                OLCase.findLOAD(name,kV)     OLCase.findLOAD('WASHINGTON',33)   """
        try:
            return LOAD(val1,val2)
        except:
            __checkArgs1__('findLOAD',val1,val2)
            __check_currFileIdx1__()
            return None
    #
    def findGENUNIT(self,val1):
        """ find GENUNIT (return None if not found):
                OLCase.findGENUNIT(GUID)    OLCase.findGENUNIT('{b933fc57-691f-4b24-a5bf-b9000c082d14}'
                OLCase.findGENUNIT(STR)     OLCase.findGENUNIT("[GENUNIT]  1@2 'CLAYTOR' 132 kV") """
        try:
            return GENUNIT(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findLOADUNIT(self,val1):
        """ find LOADUNIT (return None if not found):
                OLCase.findLOADUNIT(GUID)    OLCase.findLOADUNIT('{61e0ebc9-96ef-4cec-b8ed-d91ece2d9b3c}'
                OLCase.findLOADUNIT(STR)     OLCase.findLOADUNIT("[LOADUNIT]  1@17 'WASHINGTON' 33 kV") """
        try:
            return LOADUNIT(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findSHUNTUNIT(self,val1):
        """ find SHUNTUNIT (return None if not found):
                OLCase.findSHUNTUNIT(GUID)    OLCase.findSHUNTUNIT('{61e0ebc9-96ef-4cec-b8ed-d91ece2d9b3c}'
                OLCase.findSHUNTUNIT(STR)     OLCase.findSHUNTUNIT("[CAPUNIT]  1@21 'IOWA' 33 kV")  """
        try:
            return SHUNTUNIT(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYGROUP(self,val1):
        """ find RLYGROUP (return None if not found):
                OLCase.findRLYGROUP(GUID)  OLCase.findRLYGROUP('{369fce04-353b-4c81-8e8e-9c4d97784206}'
                OLCase.findRLYGROUP(STR)   OLCase.findRLYGROUP("[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")  """
        try:
            return RLYGROUP(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYOCG(self,val1):
        """ find RLYOCG (return None if not found):
                OLCase.findRLYOCG(GUID)  OLCase.findRLYOCG('{b4e006f6-6ccb-497e-8828-8938df46bd4a}'
                OLCase.findRLYOCG(STR)   OLCase.findRLYOCG("[OCRLYG]  NV-G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        try:
            return RLYOCG(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYOCP(self,val1):
        """ find RLYOCP (return None if not found):
                OLCase.findRLYOCP(GUID)  OLCase.findRLYOCP('{b4e006f6-6ccb-497e-8828-8938df46bd4a}'
                OLCase.findRLYOCP(STR)   OLCase.findRLYOCP("[OCRLYP]  NV-P1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        try:
            return RLYOCP(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findFUSE(self,val1):
        """ find FUSE (return None if not found):
                OLCase.findFUSE(GUID)  OLCase.findFUSE('{d67118b8-bb8c-4f37-9765-bf3c49cb2cba}'
                OLCase.findFUSE(STR)   OLCase.findFUSE("[FUSE]  NV Fuse@6 'NEVADA' 132 kV-4 'TENNESSEE' 132 kV 1 P") """
        try:
            return FUSE(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYDSG(self,val1):
        """ find RLYDSG (return None if not found):
                OLCase.findRLYDSG(GUID)  OLCase.findRLYDSG('{a004766f-125d-47b6-b9cd-70cb254f2a59}'
                OLCase.findRLYDSG(STR)   OLCase.findRLYDSG("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")"""
        try:
            return RLYDSG(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYDSP(self,val1):
        """ find RLYDSP (return None if not found):
                OLCase.findRLYDSP(GUID)  OLCase.findRLYDSP('{a004766f-125d-47b6-b9cd-70cb254f2a59}'
                OLCase.findRLYDSP(STR)   OLCase.findRLYDSP("[DSRLYP]  GCXTEST@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L")"""
        try:
            return RLYDSP(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYD(self,val1):
        """ find RLYD (return None if not found):
                OLCase.findRLYD(GUID)  OLCase.findRLYD('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                OLCase.findRLYD(STR)   OLCase.findRLYD("[DEVICEDIFF]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")"""
        try:
            return RLYD(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRLYV(self,val1):
        """ find RLYV (return None if not found):
                OLCase.findRLYV(GUID)  OLCase.findRLYV('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                OLCase.findRLYV(STR)   OLCase.findRLYV("[DEVICEVR]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L")"""
        try:
            return RLYV(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findRECLSR(self,val1):
        """ find RECLSR (return None if not found):
                OLCase.findRECLSR(GUID)  OLCase.findRECLSR('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                OLCase.findRECLSR(STR)   OLCase.findRECLSR("[RECLSRP]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        try:
            return RECLSR(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findSCHEME(self,val1):
        """ find SCHEME (return None if not found):
                OLCase.findSCHEME(GUID)  OLCase.findSCHEME('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                OLCase.findSCHEME(STR)   OLCase.findSCHEME("[PILOT]  272-POTT-SEL421G@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        try:
            return SCHEME(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findZCORRECT(self,val1):
        """ find ZCORRECT (return None if not found):
                OLCase.findZCORRECT(GUID)  OLCase.findZCORRECT('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                OLCase.findZCORRECT(STR)   OLCase.findZCORRECT("[ZCORRECT]...") """
        try:
            return ZCORRECT(val1)
        except:
            __check_currFileIdx1__()
            return None
    #
    def findBREAKER(self,val1):
        """ find ZCORRECT (return None if not found):
                OLCase.findBREAKER(GUID)  OLCase.findBREAKER('{369fce04-353b-4c81-8e8e-9c4d97784206}'
                OLCase.findBREAKER(STR)   OLCase.findBREAKER("[BREAKER]  1E82A@ 6 'NEVADA' 132 kV") """
        try:
            return BREAKER(val1)
        except:
            __check_currFileIdx1__()
            return None
#$ASPEN$ CODE GENERATED AUTOMATIC by OlxObj\genCodeOlxObj.py START

    @property
    def BASEMVA(self):
        """ (float) NETWORK System MVA base """
        return self.getData('BASEMVA')
    @property
    def COMMENT(self):
        """ (str) NETWORK comment """
        return self.getData('COMMENT')
    @property
    def OBJCOUNT(self):
        """ (dict) Number of object in NETWORK """
        return self.getData('OBJCOUNT')
    @property
    def BREAKER(self):
        """ [BREAKER] List Breakers in NETWORK """
        res = []
        for b1 in self.getData('BREAKER'):
            res.append(BREAKER(hnd=b1.__hnd__))
        return res
    @property
    def BUS(self):
        """ [BUS] List Buses in NETWORK """
        res = []
        for b1 in self.getData('BUS'):
            res.append(BUS(hnd=b1.__hnd__))
        return res

    @property
    def CCGEN(self):
        """ [CCGEN] List Voltage Controlled Current Sources in NETWORK """
        res = []
        for b1 in self.getData('CCGEN'):
            res.append(CCGEN(hnd=b1.__hnd__))
        return res
    @property
    def DCLINE2(self):
        """ [DCLINE2] List DC Lines in NETWORK """
        res = []
        for b1 in self.getData('DCLINE2'):
            res.append(DCLINE2(hnd=b1.__hnd__))
        return res
    @property
    def FUSE(self):
        """ [FUSE] List Fuses in NETWORK """
        res = []
        for b1 in self.getData('FUSE'):
            res.append(FUSE(hnd=b1.__hnd__))
        return res
    @property
    def GEN(self):
        """ [GEN] List Generators in NETWORK """
        res = []
        for b1 in self.getData('GEN'):
            res.append(GEN(hnd=b1.__hnd__))
        return res
    @property
    def GENUNIT(self):
        """ [GENUNIT] List Generator units in NETWORK """
        res = []
        for b1 in self.getData('GENUNIT'):
            res.append(GENUNIT(hnd=b1.__hnd__))
        return res
    @property
    def GENW3(self):
        """ [GENW3] List Type-3 Wind Plants in NETWORK """
        res = []
        for b1 in self.getData('GENW3'):
            res.append(GENW3(hnd=b1.__hnd__))
        return res
    @property
    def GENW4(self):
        """ [GENW4] List Converter-Interfaced Resources in NETWORK """
        res = []
        for b1 in self.getData('GENW4'):
            res.append(GENW4(hnd=b1.__hnd__))
        return res
    @property
    def LINE(self):
        """ [LINE] List Transmission Lines in NETWORK """
        res = []
        for b1 in self.getData('LINE'):
            res.append(LINE(hnd=b1.__hnd__))
        return res
    @property
    def LOAD(self):
        """ [LOAD] List Loads in NETWORK """
        res = []
        for b1 in self.getData('LOAD'):
            res.append(LOAD(hnd=b1.__hnd__))
        return res
    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List Load units in NETWORK """
        res = []
        for b1 in self.getData('LOADUNIT'):
            res.append(LOADUNIT(hnd=b1.__hnd__))
        return res
    @property
    def MULINE(self):
        """ [MULINE] List Mutual Pairs in NETWORK """
        res = []
        for b1 in self.getData('MULINE'):
            res.append(MULINE(hnd=b1.__hnd__))
        return res
    @property
    def RECLSR(self):
        """ [RECLSR] List Reclosers in NETWORK """
        res = []
        for b1 in self.getData('RECLSR'):
            res.append(RECLSR(hnd=b1.__hnd__))
        return res
    @property
    def RLYD(self):
        """ [RLYD] List Differential relays in NETWORK """
        res = []
        for b1 in self.getData('RLYD'):
            res.append(RLYD(hnd=b1.__hnd__))
        return res
    @property
    def RLYDSG(self):
        """ [RLYDSG] List DS Ground relays in NETWORK """
        res = []
        for b1 in self.getData('RLYDSG'):
            res.append(RLYDSG(hnd=b1.__hnd__))
        return res
    @property
    def RLYDSP(self):
        """ [RLYDSP] List DS Phase relays in NETWORK """
        res = []
        for b1 in self.getData('RLYDSP'):
            res.append(RLYDSP(hnd=b1.__hnd__))
        return res
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List Relay groups in NETWORK """
        res = []
        for b1 in self.getData('RLYGROUP'):
            res.append(RLYGROUP(hnd=b1.__hnd__))
        return res
    @property
    def RLYOCG(self):
        """ [RLYOCG] List OC Ground relays in NETWORK """
        res = []
        for b1 in self.getData('RLYOCG'):
            res.append(RLYOCG(hnd=b1.__hnd__))
        return res
    @property
    def RLYOCP(self):
        """ [RLYOCP] List OC Phase relays in NETWORK """
        res = []
        for b1 in self.getData('RLYOCP'):
            res.append(RLYOCP(hnd=b1.__hnd__))
        return res
    @property
    def RLYV(self):
        """ [RLYV] List Voltage relays in NETWORK """
        res = []
        for b1 in self.getData('RLYV'):
            res.append(RLYV(hnd=b1.__hnd__))
        return res
    @property
    def SCHEME(self):
        """ [SCHEME] List Logic schemes in NETWORK """
        res = []
        for b1 in self.getData('SCHEME'):
            res.append(SCHEME(hnd=b1.__hnd__))
        return res
    @property
    def SERIESRC(self):
        """ [SERIESRC] List Series capacitor/reactor in NETWORK """
        res = []
        for b1 in self.getData('SERIESRC'):
            res.append(SERIESRC(hnd=b1.__hnd__))
        return res
    @property
    def SHIFTER(self):
        """ [SHIFTER] List Phase Shifters in NETWORK """
        res = []
        for b1 in self.getData('SHIFTER'):
            res.append(SHIFTER(hnd=b1.__hnd__))
        return res
    @property
    def SHUNT(self):
        """ [SHUNT] List Shunts in NETWORK """
        res = []
        for b1 in self.getData('SHUNT'):
            res.append(SHUNT(hnd=b1.__hnd__))
        return res
    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List Shunt units in NETWORK """
        res = []
        for b1 in self.getData('SHUNTUNIT'):
            res.append(SHUNTUNIT(hnd=b1.__hnd__))
        return res
    @property
    def SVD(self):
        """ [SVD] List Switched Shunt in NETWORK """
        res = []
        for b1 in self.getData('SVD'):
            res.append(SVD(hnd=b1.__hnd__))
        return res
    @property
    def SWITCH(self):
        """ [SWITCH] List Switches in NETWORK """
        res = []
        for b1 in self.getData('SWITCH'):
            res.append(SWITCH(hnd=b1.__hnd__))
        return res
    @property
    def XFMR(self):
        """ [XFMR] List 2-Windings Transformers in NETWORK """
        res = []
        for b1 in self.getData('XFMR'):
            res.append(XFMR(hnd=b1.__hnd__))
        return res
    @property
    def XFMR3(self):
        """ [XFMR3] List 3-Windings Transformers in NETWORK """
        res = []
        for b1 in self.getData('XFMR3'):
            res.append(XFMR3(hnd=b1.__hnd__))
        return res
    @property
    def ZCORRECT(self):
        """ [ZCORRECT] List Impedance correction tables in NETWORK """
        res = []
        for b1 in self.getData('ZCORRECT'):
            res.append(ZCORRECT(hnd=b1.__hnd__))
        return res
    @property
    def RLYOC(self):
        """ [RLYOCG+RLYOCP] List Over Current Relays in NETWORK """
        res = self.RLYOCG
        res.extend(self.RLYOCP)
        return res
    @property
    def RLYDS(self):
        """ [RLYDSG+RLYDSP] List Distance Relays in NETWORK """
        res = self.RLYDSG
        res.extend(self.RLYDSP)
        return res
    @property
    def AREA(self):
        """ [AREA] NETWORK list of Area """
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        seta = set()
        for b1 in self.__allBus__:
            seta.add(b1.getData('AREANO'))
        return [AREA(a1) for a1 in sorted(list(seta))]
    @property
    def AREANO(self):
        """ [int] NETWORK list of Area Number """
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        seta = set()
        for b1 in self.__allBus__:
            seta.add(b1.getData('AREANO'))
        return sorted(list(seta))
    @property
    def ZONE(self):
        """ [ZONE] NETWORK list of Zone """
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        seta = set()
        for b1 in self.__allBus__:
            seta.add(b1.getData('ZONENO'))
        return [ZONE(a1) for a1 in sorted(list(seta))]
    @property
    def ZONENO(self):
        """ [init] NETWORK list of Zone Number """
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        seta = set()
        for b1 in self.__allBus__:
            seta.add(b1.getData('ZONENO'))
        return sorted(list(seta))
    @property
    def KV(self):
        """ [float] NETWORK list of KV Nominal """
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        seta = set()
        for b1 in self.__allBus__:
            seta.add(b1.KV)
        return sorted(list(seta),reverse=True)
    #
    def simulateFault(self,specFlt,clearPrev=0):
        """ Run simulation of a classical fault, simultaneous fault or stepped-event analysis
        Args:
            specFlt: Classical/Simultaneous/Stepped-Event Analysis specfication(s)
                SPEC_FLT: 'Classical' or 'Simultaneous' or 'SEA' (Stepped-Event Analysis)
                [SPEC_FLT] : Simultaneous
                [sp_0,sp_i,...] sp_0=SPEC_FLT('SEA') sp_i=SPEC_FLT('SEA_EX')
            clearPrev : (0/1) clear previous result flag
        """
        # check
        se = '\nOLCase.SimulateFault(specFlt,clearPrev)'
        if type(clearPrev) != int or clearPrev not in {0,1}:
            se+= '\n  clearPrev : clear previous result flag'
            se+= '\n\tRequired          : (int) 0 or 1 '
            if type(clearPrev)==int:
                se+= "\n\tFound (ValueError): %i"%clearPrev
            else:
                se+= '\n\tFound (ValueError): (%s)'%type(clearPrev).__name__ + ' ' +str(clearPrev)
            raise ValueError(se)
        #
        if type(specFlt)!=SPEC_FLT or specFlt.__type__ not in {'Classical','Simultaneous','SEA'}:
            flag = type(specFlt)==list and len(specFlt)>0
            if flag:
                flag1 = True
                for s1 in specFlt:
                    if type(s1)!=SPEC_FLT or s1.__type__ !='Simultaneous':
                        flag1 = False
                flag2 = type(specFlt[0])==SPEC_FLT and specFlt[0].__type__=='SEA'
                for i in range(1,len(specFlt)):
                    if type(specFlt[i])!=SPEC_FLT or specFlt[i].__type__ !='SEA_EX':
                        flag2 = False
                flag = flag1 or flag2
            #
            if not flag:
                se += '\n  specFlt : Fault Specfication'
                se += '\n\tRequired : '
                se += "\n\t\t   SPEC_FLT ('Classical','Simultaneous','SEA')"
                se += "\n\t\tor [SPEC_FLT] ('Simultaneous')"
                se += "\n\t\tor [sp_0,sp_i,...] sp_0=SPEC_FLT('SEA') sp_i=SPEC_FLT('SEA_EX')"
                se += '\n\tFound (ValueError) : '
                if type(specFlt)!= list:
                    se+= type(specFlt).__name__
                else:
                    se+='['
                    for sp1 in specFlt:
                        if type(sp1)==SPEC_FLT:
                            se+="SPEC_FLT('%s')"%sp1.__type__ +','
                        else:
                            se+=type(sp1).__name__+','
                    se=se[:-1]+']'
                raise ValueError(se)
        #
        __runSimulate__(specFlt,clearPrev)
    #
    def substationGroup(self,xsmall=None,lsmall=None):# xsmall for 100m, 0.1*0.4=>0.04Ohm
        """
            Grouper Bus and Equipment by substation
        return:
            dict['BUS']      : list of Bus by substation
            dict['EQUIPMENT] : list of Equipmement by substation
        """
##        xsmall=0.04
##        lsmall=0.1 # xsmall for 100m, 0.1*0.4=>0.04Ohm
        lstBus,lstEquipment = [],[]
        #
        if self.__allBus__ is None:
            super().__setattr__('__allBus__',self.getData('BUS'))
        #
        setBusHnd,setEquiHnd = set(),set()
        for b1 in self.__allBus__:
            if b1.__hnd__ not in setBusHnd:
                resB1 = [b1]
                setEqui1 = []
                b11 = __getSubStation__(b1,setEqui1,setEquiHnd,xsmall,lsmall)
                resB1.extend( b11 )
                #
                lstBus.append(resB1)
                lstEquipment.append(setEqui1)
                #
                for bi in resB1:
                    setBusHnd.add(bi.__hnd__)
        return {'BUS': lstBus,'EQUIPMENT':lstEquipment}
    #
    def findLineSections(self,t0,prt=False):
        """
        Purpose: Find all section (LINE,SERIESRC,SWITCH) from a TERMINAL/RLYGROUP
                All taps are ignored. Close switches are included.
                Branches out of service are ignored
        Args:
            t0: start TERMINAL/RLYGROUP
                Exception if:
                    1st Bus (of t0) is a TAP Bus
                    t0 located on XFMR,XFMR3,SHIFTER
        return:
            mainLine  = [[]]     List of all TERMINAL in the main-line of method 1,2,3 in Help 8.9
                                        METHOD 1: Enter the same name in the Name field of the line’s main segments
                                        METHOD 2: Include these three characters [T] (or [t]) in the tap lines name
                                        METHOD 3: Give the tap lines circuit IDs that contain letter T or t
            remoteBus  = []      List of remote BUSES: two for 2-terminal line and >2 for 3-terminal line
            remoteRLG  = []      List of all RLYGROUP at the remote end
        """
        if type(t0) not in {TERMINAL,RLYGROUP}:
            raise Exception('OLCase.findLineSection(t0) with t0 is a TERMINAL or RLYGROUP\n\tFound:'+toString(t0))
        mainLineHnd,remoteBusHnd,remoteRLGHnd = OlxAPILib.findLineSections(t0.HANDLE,prt)
        mainLine = self.toOBJ(mainLineHnd)
        remoteBus = self.toOBJ(remoteBusHnd)
        remoteRLG = self.toOBJ(remoteRLGHnd)
        return mainLine,remoteBus,remoteRLG
    #
    def tapLineTool(self,t0,prt=False):
        """
        Purpose: Find main sections of Line and sum impedance(Z0,Z1) and Length
            All taps are ignored. Close switches are included.
            Branches out of service are ignored
        Args:
            t0: start TERMINAL/RLYGROUP
                Exception if:
                    1st Bus (of t0) is a TAP Bus
                    t0 located on XFMR,XFMR3,SHIFTER
            prt: (True/False)
                option print to stdout all details when research mainLine
        return:
            res['mainLine']   = [[]] List of all TERMINAL in the main-line of method 1,2,3 in Help 8.9
                                        METHOD 1: Enter the same name in the Name field of the line’s main segments
                                        METHOD 2: Include these three characters [T] (or [t]) in the tap lines name
                                        METHOD 3: Give the tap lines circuit IDs that contain letter T or t
            res['remoteBus']  = []   List of remote BUSES on the mainLine: two for 2-terminal line and >2 for 3-terminal line
            res['remoteRLG']  = []   List of all RLYGROUP at the remote end on the mainLine
            res['Z1']         = []   List of positive sequence Impedance of the mainLine
            res['Z0']         = []   List of zero sequence Impedance of the mainLine
            res['Length']     = []   List of sum length of the mainLine
        """
        if type(t0) not in {TERMINAL,RLYGROUP}:
            raise Exception('\nOLCase.tapLineTool(t0) with t0 is a TERMINAL or RLYGROUP\n\tFound:'+toString(t0))
        ra = OlxAPILib.tapLineTool(t0.HANDLE,prt)
        res = dict()
        res['mainLine'] = self.toOBJ(ra['mainLineHnd'])
        res['remoteBus'] = self.toOBJ(ra['remoteBusHnd'])
        res['remoteRLG'] = self.toOBJ(ra['remoteRLGHnd'])
        res['Z1'] = ra['Z1']
        res['Z0'] = ra['Z0']
        res['Length'] = ra['Length']
        return res
    #
    def toOBJ(self,handle):
        """ Convert handle => OBJECT
            return None if not found
        example: (with 4,5 are handle of BUS)
            4       => BUS
            [4,5]   => [BUS, BUS]
            [[4,5]] => [[BUS, BUS]]
        """
        if type(handle)==list:
            return [self.toOBJ(h1) for h1 in handle]
        try:
            tc1 = OlxAPI.EquipmentType(handle)
            return __getOBJ__(handle,tc=tc1)
        except:
            return None
#
OLCase = NETWORK()
#
class AREA:
    def __init__(self,no):
        super().__setattr__('__no__', no)
        super().__setattr__('__currFileIdx__',__CURRENT_FILE_IDX__)
    #
    @property
    def NO(self):
        """ (int) AREA Number """
        __check_currFileIdx__(self)
        return self.__no__
    @property
    def NAME(self):
        """ (str) AREA Name """
        __check_currFileIdx__(self)
        return ''
    @property
    def SLACKBUS(self):
        """ (BUS) AREA Slack Bus """
        __check_currFileIdx__(self)
        return None
    @property
    def EXPORT(self):
        """ (float) AREA Export MW """
        __check_currFileIdx__(self)
        return 0
    @property
    def GEN(self):
        """ (float) AREA Generation MW """
        __check_currFileIdx__(self)
        return 0
    @property
    def LOAD(self):
        """ (float) AREA Load MW """
        __check_currFileIdx__(self)
        return 0
    #
    def equals(self,o2):
        """ return (bool) comparison of 2 objects """
        if type(o2)==AREA:
            __check_currFileIdx__(self)
            __check_currFileIdx__(o2)
            return self.__no__==o2.__no__
        return False
    #
    def toString(self,option=0):
        if option==0:
            return '[AREA] %i %s'%(self.__no__,self.NAME)
        return '%i %s'%(self.__no__,self.NAME)
#
class ZONE:
    def __init__(self,no):
        super().__setattr__('__no__', no)
        super().__setattr__('__currFileIdx__',__CURRENT_FILE_IDX__)
    #
    @property
    def NO(self):
        """ (int) ZONE Number """
        __check_currFileIdx__(self)
        return self.__no__
    @property
    def NAME(self):
        """ (str) ZONE Name """
        __check_currFileIdx__(self)
        return ''
    #
    def equals(self,o2):
        """ return (bool) comparison of 2 objects """
        if type(o2)==ZONE:
            __check_currFileIdx__(self)
            __check_currFileIdx__(o2)
            return self.__no__==o2.__no__
        return False
    #
    def toString(self,option=0):
        if option==0:
            return '[ZONE] %i %s'%(self.__no__,self.NAME)
        return '%i %s'%(self.__no__,self.NAME)
#
class BUS(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ BUS constructor (Exception if not found):
                BUS(name,kV)     BUS('arizona',132) or BUS(132,'arizona')
                BUS(busNo)       BUS(28)
                BUS(GUID)        BUS('{b4f6d06c-473b-4365-ba05-a1b47a1f9364}'
                BUS(STR)         BUS("[BUS] 28 'ARIZONA' 132 kV") """
        __checkArgs__([val1,val2,hnd],'BUS')
        if hnd is None:
            if val2 is None:
                if type(val1)==int:
                    hnd = OlxAPI.FindBusNo(val1)
                else:
                    hnd = __findObjHnd__(val1,BUS)
            elif type(val1)==str and (type(val2)==float or int):
                hnd = OlxAPI.FindBus(val1,val2)
            elif type(val2)==str and (type(val1)==float or int):
                hnd = OlxAPI.FindBus(val2,val1)
            else:
                hnd = 0
            __initFailCheck__(hnd,'BUS',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('BUS')
    @property
    def ANGLEP(self):
        """ (float) Bus voltage angle (degree) from a power flow solution """
        return self.getData('ANGLEP')
    @property
    def AREANO(self):
        """ (int) Bus area number """
        return self.getData('AREANO')
    @property
    def AREA(self):
        """ (AREA) Bus area """
        return AREA(self.getData('AREANO'))
    @property
    def VISIBLE(self):
        """ (int) Bus hide ID flag: 1-visible; -2-hidden; 0-not yet placed """
        return self.getData('VISIBLE')
    @property
    def KV(self):
        """ (float) Bus voltage nominal """
        return self.getData('KV')
    @property
    def KVP(self):
        """ (float) Bus voltage magnitude (kV) from a power flow solution """
        return self.getData('KVP')
    @property
    def LOCATION(self):
        """ (str) Bus location """
        return self.getData('LOCATION')
    @property
    def MEMO(self):
        """ (str) Bus memo """
        return self.getData('MEMO')
    @property
    def NAME(self):
        """ (str) Bus name """
        return self.getData('NAME')
    @property
    def NO(self):
        """ (int) Bus number """
        return self.getData('NO')
    @property
    def SLACK(self):
        """ (int) System slack bus flag: 1-yes; 0-no """
        return self.getData('SLACK')
    @property
    def SPCX(self):
        """ (float) Bus state plane coordinate - X """
        return self.getData('SPCX')
    @property
    def SPCY(self):
        """ (float) Bus state plane coordinate - Y """
        return self.getData('SPCY')
    @property
    def SUBGRP(self):
        """ (int) Bus substation group """
        return self.getData('SUBGRP')
    @property
    def TAP(self):
        """ (int) Tap bus flag: 0-no; 1-tap bus; 3-tap bus of 3-terminal line """
        return self.getData('TAP')
    @property
    def ZONENO(self):
        """ (int) Bus zone number """
        return self.getData('ZONENO')
    @property
    def ZONE(self):
        """ (ZONE) Bus ZONE """
        return ZONE(self.getData('ZONENO'))
    @property
    def BREAKER(self):
        """ [BREAKER] List Breakers connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'BREAKER'):
            res.append(BREAKER(hnd=h1))
        return res
    @property
    def CCGEN(self):
        """ CCGEN-Voltage Controlled Current Sources connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'CCGEN')
        try:
            return CCGEN(hnd=g[0])
        except:
            return None
    @property
    def DCLINE2(self):
        """ [DCLINE2] List DC Lines connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'DCLINE2'):
            res.append(DCLINE2(hnd=h1))
        return res
    @property
    def GEN(self):
        """ GEN-Generator connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'GEN')
        try:
            return GEN(hnd=g[0])
        except:
            return None
    @property
    def GENW3(self):
        """ GENW3-Type-3 Wind Plants connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'GENW3')
        try:
            return GENW3(hnd=g[0])
        except:
            return None
    @property
    def GENW4(self):
        """ GENW4-Converter-Interfaced Resources connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'GENW4')
        try:
            return GENW4(hnd=g[0])
        except:
            return None
    @property
    def LINE(self):
        """ [LINE] List Transmission Lines connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'LINE'):
            res.append(LINE(hnd=h1))
        return res
    @property
    def LOAD(self):
        """ LOAD-Load connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'LOAD')
        try:
            return LOAD(hnd=g[0])
        except:
            return None
    @property
    def SHIFTER(self):
        """ [SHIFTER] List Phase Shifters connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'SHIFTER'):
            res.append(SHIFTER(hnd=h1))
        return res
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List Relay groups connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'RLYGROUP'):
            res.append(RLYGROUP(hnd=h1))
        return res
    @property
    def SERIESRC(self):
        """ [SERIESRC] List Series capacitor/reactor connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'SERIESRC'):
            res.append(SERIESRC(hnd=h1))
        return res
    @property
    def SHUNT(self):
        """ SHUNT-Shunt connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'SHUNT')
        try:
            return SHUNT(hnd=g[0])
        except:
            return None
    @property
    def SVD(self):
        """ SVD-Switched Shunt connected to BUS
            None if not found """
        g = __getBusEquipmentHnd__(self,'SVD')
        try:
            return SVD(hnd=g[0])
        except:
            return None
    @property
    def SWITCH(self):
        """ [SWITCH] List Switches connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'SWITCH'):
            res.append(SWITCH(hnd=h1))
        return res
    @property
    def XFMR(self):
        """ [XFMR] List 2-Windings Transformers connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'XFMR'):
            res.append(XFMR(hnd=h1))
        return res
    @property
    def XFMR3(self):
        """ [XFMR3] List 3-Windings Transformers connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'XFMR3'):
            res.append(XFMR3(hnd=h1))
        return res
    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List Load units connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'LOADUNIT'):
            res.append(LOADUNIT(hnd=h1))
        return res
    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List Shunt units connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'SHUNTUNIT'):
            res.append(SHUNTUNIT(hnd=h1))
        return res
    @property
    def GENUNIT(self):
        """ [GENUNIT] List Generator units connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'GENUNIT'):
            res.append(GENUNIT(hnd=h1))
        return res
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs connected to BUS """
        res = []
        for h1 in __getBusEquipmentHnd__(self,'TERMINAL'):
            res.append(TERMINAL(hnd=h1))
        return res
    #
    def terminalTo(self,b2):
        """ return [TERMINAL] List TERMINALs from this BUS to BUS b2 """
        if type(b2)!=BUS:
            se = '\nBUS.terminalTo(b2)'
            se+= "\n\tRequired (b2)      : (BUS)"
            se+= '\n\t' +__getErrValue__(BUS,b2)
            raise ValueError(se)
        __check_currFileIdx__(self)
        __check_currFileIdx__(b2)
        res = []
        hnd1,hnd2 = self.__hnd__,b2.__hnd__
        val1 = c_int(0)
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd1,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            val2 = __getDatai__(val1.value,OlxAPIConst.BR_nBus2Hnd)
            val3 = __getDatai__(val1.value,OlxAPIConst.BR_nBus3Hnd)
            if val2==hnd2 or val3==hnd2:
                res.append( TERMINAL(hnd=val1.value) )
        return res
#
class GEN(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ GEN constructor (Exception if not found):
                GEN(GUID)        GEN('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                GEN(STR)         GEN("[GENERATOR] 2 'CLAYTOR' 132 kV")
                GEN(name,kV)     GEN('CLAYTOR',132)  """
        __checkArgs__([val1,val2,hnd],'GEN')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,GEN)
            else:
                try:
                    hnd = BUS(val1,val2).GEN[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'GEN',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('GEN')
    @property
    def BUS(self):
        """ (BUS) Generators connected BUS """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def CNTBUS(self):
        """ (BUS) Generators controlled BUS """
        return BUS(hnd=self.getData('CNTBUS').__hnd__)
    @property
    def FLAG(self):
        """ (int) Generator in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ILIMIT1(self):
        """ (float) Generator current limit 1 """
        return self.getData('ILIMIT1')
    @property
    def ILIMIT2(self):
        """ (float) Generator current limit 2 """
        return self.getData('ILIMIT2')
    @property
    def PGEN(self):
        """ (float) Generator MW (load flow solution) """
        return self.getData('PGEN')
    @property
    def QGEN(self):
        """ (float) Generator MVAR (load flow solution) """
        return self.getData('QGEN')
    @property
    def REFANGLE(self):
        """ (float) Generator reference angle """
        return self.getData('REFANGLE')
    @property
    def REFV(self):
        """ (float) Generator internal voltage source per unit magnitude """
        return self.getData('REFV')
    @property
    def REG(self):
        """ (int) Generator regulation flag: 1- PQ; 0- PV """
        return self.getData('REG')
    @property
    def SCHEDP(self):
        """ (float) Generator scheduled P """
        return self.getData('SCHEDP')
    @property
    def SCHEDQ(self):
        """ (float) Generator scheduled Q """
        return self.getData('SCHEDQ')
    @property
    def SCHEDV(self):
        """ (float) Generator scheduled V """
        return self.getData('SCHEDV')
    @property
    def GENUNIT(self):
        """ [GENUNIT] List GEN. units in GEN """
        res = []
        for g1 in self.BUS.GENUNIT:
            res.append(GENUNIT(hnd=g1.__hnd__))
        return res
#
class GENUNIT(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ GENUNIT constructor (Exception if not found):
                GENUNIT(GUID)        GENUNIT('{b933fc57-691f-4b24-a5bf-b9000c082d14}'
                GENUNIT(STR)         GENUNIT("[GENUNIT]  1@2 'CLAYTOR' 132 kV") """
        __checkArgs__([val1,hnd],'GENUNIT')
        if hnd is None:
            hnd = __findObjHnd__(val1,GENUNIT)
            __initFailCheck__(hnd,'GENUNIT',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('GENUNIT')
    @property
    def CID(self):
        """ (str) Generator unit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Generator unit out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Generator unit in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Generator unit in-service flag 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def GEN(self):
        """ (GEN) Generator unit generator """
        return GEN(hnd=self.getData('GEN').__hnd__)
    @property
    def MVARATE(self):
        """ (float) Generator unit rating MVA """
        return self.getData('MVARATE')
    @property
    def PMAX(self):
        """ (float) Generator unit max MW """
        return self.getData('PMAX')
    @property
    def PMIN(self):
        """ (float) Generator unit min MW """
        return self.getData('PMIN')
    @property
    def QMAX(self):
        """ (float) Generator unit max MVAR """
        return self.getData('QMAX')
    @property
    def QMIN(self):
        """ (float) Generator unit min MVAR """
        return self.getData('QMIN')
    @property
    def R(self):
        """ [float]*5 Generator unit resistances
            [subtransient, synchronous, transient, negative sequence, zero sequence] """
        return self.getData('R')
    @property
    def RG(self):
        """ (float) Generator unit grounding resistance in Ohm (do not multiply by 3) """
        return self.getData('RG')
    @property
    def SCHEDP(self):
        """ (float) Generator unit scheduled P """
        return self.getData('SCHEDP')
    @property
    def SCHEDQ(self):
        """ (float) Generator unit scheduled Q """
        return self.getData('SCHEDQ')
    @property
    def X(self):
        """ [float]*5 Generator unit reactances
            [subtransient, synchronous, transient, negative sequence, zero sequence] """
        return self.getData('X')
    @property
    def XG(self):
        """ (float) Generator unit grounding reactance in Ohm (do not multiply by 3) """
        return self.getData('XG')
    @property
    def BUS(self):
        """ BUS of GEN unit """
        return BUS(hnd=self.GEN.BUS.__hnd__)
#
class GENW3(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ GENW3 constructor (Exception if not found):
                GENW3(GUID)        GENW3('{7e5f2278-6566-42c6-9b33-a3f27f4464f1}
                GENW3(STR)         GENW3("[GENW3] 4 'TENNESSEE' 132 kV")
                GENW3(name,kV)     GENW3('CLAYTOR',132)  """
        __checkArgs__([val1,val2,hnd],'GENW3')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,GENW3)
            else:
                try:
                    hnd = BUS(val1,val2).GENW3[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'GENW3',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('GENW3')
    @property
    def BUS(self):
        """ (BUS) Generators Type 3 connected BUS """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def CBAR(self):
        """ (int) Generator Type 3 Crowbarred flag: 1-yes; 0-no """
        return self.getData('CBAR')
    @property
    def DATEOFF(self):
        """ (str) Generator Type 3 out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Generator Type 3 in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Generator Type 3 unit in-service flag 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def FLRZ(self):
        """ (float) Generator Type 3 Filter X in pu """
        return self.getData('FLRZ')
    @property
    def IMAXG(self):
        """ (float) Generator Type 3 Grid side limit in pu """
        return self.getData('IMAXG')
    @property
    def IMAXR(self):
        """ (float) Generator Type 3 rotor side limit in pu """
        return self.getData('IMAXR')
    @property
    def LLR(self):
        """ (float) Generator Type 3 Rotor leakage L in pu """
        return self.getData('LLR')
    @property
    def LLS(self):
        """ (float) Generator Type 3 Stator leakage L in pu """
        return self.getData('LLS')
    @property
    def LM(self):
        """ (float) Generator Type 3 Mutual L in pu """
        return self.getData('LM')
    @property
    def MVA(self):
        """ (float) Generator Type 3 MVA unit rated """
        return self.getData('MVA')
    @property
    def MW(self):
        """ (float) Generator Type 3 unit MW generation """
        return self.getData('MW')
    @property
    def MWR(self):
        """ (float) Generator Type 3 MW unit rated """
        return self.getData('MWR')
    @property
    def RR(self):
        """ (float) Generator Type 3 Rotor resistance in pu """
        return self.getData('RR')
    @property
    def RS(self):
        """ (float) Generator Type 3 Stator resistance in pu """
        return self.getData('RS')
    @property
    def SLIP(self):
        """ (float) Generator Type 3 Slip at rated kW """
        return self.getData('SLIP')
    @property
    def UNITS(self):
        """ (int) Generator Type 3 number of units """
        return self.getData('UNITS')
    @property
    def VMAX(self):
        """ (float) Generator Type 3 maximum voltage limit in pu """
        return self.getData('VMAX')
    @property
    def VMIN(self):
        """ (float) Generator Type 3 minimum voltage limit in pu """
        return self.getData('VMIN')
#
class GENW4(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ GENW4 constructor (Exception if not found):
                GENW4(GUID)        GENW4('{7e5f2278-6566-42c6-9b33-a3f27f4464f1}
                GENW4(STR)         GENW4("[GENW4] 4 'TENNESSEE' 132 kV")
                GENW4(name,kV)     GENW4('CLAYTOR',132) """
        __checkArgs__([val1,val2,hnd],'GENW4')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,GENW4)
            else:
                try:
                    hnd = BUS(val1,val2).GENW4[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'GENW4',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('GENW4')
    @property
    def BUS(self):
        """ (BUS) Generators Type 4 connected BUS """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def CTRLMETHOD(self):
        """ (int) Generator Type 4 control method """
        return self.getData('CTRLMETHOD')
    @property
    def DATEOFF(self):
        """ (str) Generator Type 4 out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Generator Type 4 in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Generator Type 4 unit in-service flag 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def MAXI(self):
        """ (float) Generator Type 4 Max current """
        return self.getData('MAXI')
    @property
    def MAXILOW(self):
        """ (float) Generator Type 4 Max current reduce to """
        return self.getData('MAXILOW')
    @property
    def MVA(self):
        """ (float) Generator Type 4 Unit MVA rating """
        return self.getData('MVA')
    @property
    def MVAR(self):
        """ (float) Generator Type 4 Unit MVAR """
        return self.getData('MVAR')
    @property
    def MW(self):
        """ (float) Generator Type 4 Unit MW generation or consumption """
        return self.getData('MW')
    @property
    def SLOPE(self):
        """ (float) Generator Type 4 Slope of +seq """
        return self.getData('SLOPE')
    @property
    def SLOPENEG(self):
        """ (float) Generator Type 4 Slope of -seq """
        return self.getData('SLOPENEG')
    @property
    def UNITS(self):
        """ (int) Generator Type 4 number of units """
        return self.getData('UNITS')
    @property
    def VLOW(self):
        """ (float) Generator Type 4 When +seq(pu)> """
        return self.getData('VLOW')
    @property
    def VMAX(self):
        """ (float) Generator Type 4 maximum voltage limit """
        return self.getData('VMAX')
    @property
    def VMIN(self):
        """ (float) Generator Type 4 minimum voltage limit """
        return self.getData('VMIN')
#
class CCGEN(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ CCGEN constructor (Exception if not found):
                CCGEN(GUID)        CCGEN('{5c69a8c7-6e91-43dc-a540-7368bcca6c9e}')
                CCGEN(STR)         CCGEN("[CCGENUNIT] 'BUS5' 13 kV") """
        __checkArgs__([val1,val2,hnd],'CCGEN')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,CCGEN)
            else:
                try:
                    hnd = BUS(val1,val2).CCGEN[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'CCGEN',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('CCGEN')
    @property
    def A(self):
        """ [float]*10 Voltage controlled current source angle """
        return self.getData('A')
    @property
    def BLOCKPHASE(self):
        """ (int) Voltage controlled current source number block on phase """
        return self.getData('BLOCKPHASE')
    @property
    def BUS(self):
        """ (BUS) Voltage controlled current source BUS """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def DATEOFF(self):
        """ (str) Voltage controlled current source out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Voltage controlled current source in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Voltage controlled current source in-service flag: 1-true; 2-false """
        return self.getData('FLAG')
    @property
    def I(self):
        """ [float]*10 Voltage controlled current source current """
        return self.getData('I')
    @property
    def MVARATE(self):
        """ (float) Voltage controlled current source MVA rating """
        return self.getData('MVARATE')
    @property
    def V(self):
        """ [float]*10 Voltage controlled current source voltage """
        return self.getData('V')
    @property
    def VLOC(self):
        """ (int) Voltage controlled current source voltage measurement location 0-Device terminal; 1-Network side of transformer """
        return self.getData('VLOC')
    @property
    def VMAXMUL(self):
        """ (float) Voltage controlled current source maximum voltage limit in pu """
        return self.getData('VMAXMUL')
    @property
    def VMIN(self):
        """ (float) Voltage controlled current source minimum voltage limit in pu """
        return self.getData('VMIN')
#
class XFMR(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ XFMR constructor (Exception if not found):
                XFMR(b1,b2,CID)   XFMR(b1,b2,'1')
                XFMR(GUID)        XFMR('{e417041a-ef1b-46c2-82f2-9828d22d88b8}'
                XFMR(STR)         XFMR("[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'XFMR')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,XFMR)
            else:
                hnd = __findBr2Hnd__('XFMR',val1,val2,val3)
            __initFailCheck__(hnd,'XFMR',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('XFMR')
    @property
    def AUTOX(self):
        """ (int) 2-winding transformer auto transformer flag:1-true;0-false """
        return self.getData('AUTOX')
    @property
    def B(self):
        """ (float) 2-winding transformer B """
        return self.getData('B')
    @property
    def B0(self):
        """ (float) 2-winding transformer Bo """
        return self.getData('B0')
    @property
    def B1(self):
        """ (float) 2-winding transformer B1 """
        return self.getData('B1')
    @property
    def B10(self):
        """ (float) 2-winding transformer B10 """
        return self.getData('B10')
    @property
    def B2(self):
        """ (float) 2-winding transformer B2 """
        return self.getData('B2')
    @property
    def B20(self):
        """ (float) 2-winding transformer B20 """
        return self.getData('B20')
    @property
    def BASEMVA(self):
        """ (float) 2-winding transformer base MVA for per-unit quantities """
        return self.getData('BASEMVA')
    @property
    def BUS1(self):
        """ (BUS) 2-winding transformer bus 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) 2-winding transformer bus 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def CID(self):
        """ (str) 2-winding transformer circuit ID """
        return self.getData('CID')
    @property
    def CONFIGP(self):
        """ (str) 2-winding transformer primary winding config """
        return self.getData('CONFIGP')
    @property
    def CONFIGS(self):
        """ (str) 2-winding transformer secondary winding config """
        return self.getData('CONFIGS')
    @property
    def CONFIGST(self):
        """ (str) 2-winding transformer secondary winding config in test """
        return self.getData('CONFIGST')
    @property
    def DATEOFF(self):
        """ (str) 2-winding transformer out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) 2-winding transformer in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) 2-winding transformer in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def G1(self):
        """ (float) 2-winding transformer G1 """
        return self.getData('G1')
    @property
    def G10(self):
        """ (float) 2-winding transformer G10 """
        return self.getData('G10')
    @property
    def G2(self):
        """ (float) 2-winding transformer G2 """
        return self.getData('G2')
    @property
    def G20(self):
        """ (float) 2-winding transformer G20 """
        return self.getData('G20')
    @property
    def GANGED(self):
        """ (int) 2-winding transformer LTC tag ganged flag: 0-False; 1-True """
        return self.getData('GANGED')
    @property
    def LTCCENTER(self):
        """ (float) 2-winding transformer LTC center tap """
        return self.getData('LTCCENTER')
    @property
    def LTCCTRL(self):
        """ (str) 2-winding transformer ID string of the bus whose voltage magnitude is to be regulated by the LTC """
        return self.getData('LTCCTRL')
    @property
    def LTCSIDE(self):
        """ (int) 2-winding transformer LTC side: 1; 2; 0 """
        return self.getData('LTCSIDE')
    @property
    def LTCSTEP(self):
        """ (float) 2-winding transformer LTC step size """
        return self.getData('LTCSTEP')
    @property
    def LTCTYPE(self):
        """ (int) 2-winding transformer LTC type: 1- control voltage; 2- control MVAR """
        return self.getData('LTCTYPE')
    @property
    def MAXTAP(self):
        """ (float) 2-winding transformer LTC max tap """
        return self.getData('MAXTAP')
    @property
    def MAXVW(self):
        """ (float) 2-winding transformer LTC min controlled quantity limit """
        return self.getData('MAXVW')
    @property
    def METEREDEND(self):
        """ (int) 2-winding transformer metered bus 1-at Bus1; 2 at Bus2; 0 XFMR in a single area """
        return self.getData('METEREDEND')
    @property
    def MINTAP(self):
        """ (float) 2-winding transformer LTC min tap """
        return self.getData('MINTAP')
    @property
    def MINVW(self):
        """ (float) 2-winding transformer LTC max controlled quantity limit """
        return self.getData('MINVW')
    @property
    def MVA1(self):
        """ (float) 2-winding transformer MVA1 """
        return self.getData('MVA1')
    @property
    def MVA2(self):
        """ (float) 2-winding transformer MVA2 """
        return self.getData('MVA2')
    @property
    def MVA3(self):
        """ (float) 2-winding transformer MVA3 """
        return self.getData('MVA3')
    @property
    def NAME(self):
        """ (str) 2-winding transformer name """
        return self.getData('NAME')
    @property
    def PRIORITY(self):
        """ (int) 2-winding transformer LTC adjustment priority 0-Normal, 1-Medieum, 2-High """
        return self.getData('PRIORITY')
    @property
    def PRITAP(self):
        """ (float) 2-winding transformer primary tap """
        return self.getData('PRITAP')
    @property
    def R(self):
        """ (float) 2-winding transformer R """
        return self.getData('R')
    @property
    def R0(self):
        """ (float) 2-winding transformer Ro """
        return self.getData('R0')
    @property
    def RG1(self):
        """ (float) 2-winding transformer Rg1 """
        return self.getData('RG1')
    @property
    def RG2(self):
        """ (float) 2-winding transformer Rg2 """
        return self.getData('RG2')
    @property
    def RGN(self):
        """ (float) 2-winding transformer Rgn """
        return self.getData('RGN')
    @property
    def SECTAP(self):
        """ (float) 2-winding transformer secondary tap """
        return self.getData('SECTAP')
    @property
    def X(self):
        """ (float) 2-winding transformer X """
        return self.getData('X')
    @property
    def X0(self):
        """ (float) 2-winding transformer Xo """
        return self.getData('X0')
    @property
    def XG1(self):
        """ (float) 2-winding transformer Xg1 """
        return self.getData('XG1')
    @property
    def XG2(self):
        """ (float) 2-winding transformer Xg2 """
        return self.getData('XG2')
    @property
    def XGN(self):
        """ (float) 2-winding transformer Xgn """
        return self.getData('XGN')
    @property
    def BUS(self):
        """ [BUS] List Buses of XFMR"""
        return [self.BUS1,self.BUS2]
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of XFMR
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of XFMR
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of XFMR """
        return [self.RLYGROUP1,self.RLYGROUP2]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of XFMR """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class XFMR3(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ XFMR3 constructor (Exception if not found):
                XFMR3(b1,b2,CID)  XFMR3(b1,b2,'1')
                XFMR3(GUID)          XFMR3('{4a6b16d7-9469-4361-9d8e-60caf451f50a}'
                XFMR3(STR)           XFMR3("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-'DOT BUS' 13.8 kV 1")  """
        __checkArgs__([val1,val2,val3,hnd],'XFMR3')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,XFMR3)
            else:
                hnd = __findBr2Hnd__('XFMR3',val1,val2,val3)
            __initFailCheck__(hnd,'XFMR3',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,val4=None,hnd=None):
        __oldinit214__('XFMR3')
    @property
    def AUTOX(self):
        """ (int) 3-winding transformer auto transformer flag:1-true;0-false """
        return self.getData('AUTOX')
    @property
    def B(self):
        """ (float) 3-winding transformer B """
        return self.getData('B')
    @property
    def B0(self):
        """ (float) 3-winding transformer B0 """
        return self.getData('B0')
    @property
    def BASEMVA(self):
        """ (float) 3-winding transformer base MVA for per-unit quantities """
        return self.getData('BASEMVA')
    @property
    def BUS1(self):
        """ (BUS) 3-winding transformer bus 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) 3-winding transformer bus 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS3(self):
        """ (BUS) 3-winding transformer bus 3 """
        return BUS(hnd=self.getData('BUS3').__hnd__)
    @property
    def BUS(self):
        """ return [BUS] List Buses of XFMR3 """
        return [self.BUS1,self.BUS2,self.BUS3]
    @property
    def CID(self):
        """ (str) 3-winding transformer circuit ID """
        return self.getData('CID')
    @property
    def CONFIGP(self):
        """ (str) 3-winding transformer primary winding """
        return self.getData('CONFIGP')
    @property
    def CONFIGS(self):
        """ (str) 3-winding transformer secondary winding """
        return self.getData('CONFIGS')
    @property
    def CONFIGST(self):
        """ (str) 3-winding transformer secondary winding in test """
        return self.getData('CONFIGST')
    @property
    def CONFIGT(self):
        """ (str) 3-winding transformer tertiary winding """
        return self.getData('CONFIGT')
    @property
    def CONFIGTT(self):
        """ (str) 3-winding transformer tertiary winding in test """
        return self.getData('CONFIGTT')
    @property
    def DATEOFF(self):
        """ (str) 3-winding transformer out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) 3-winding transformer in service date """
        return self.getData('DATEON')
    @property
    def FICTBUSNO(self):
        """ (int) 3-winding transformer Fiction bus Number """
        return self.getData('FICTBUSNO')
    @property
    def FLAG(self):
        """ (int) 3-winding transformer in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def GANGED(self):
        """ (int) 3-winding transformer LTC tag ganged flag: 0-False; 1-True """
        return self.getData('GANGED')
    @property
    def LTCCENTER(self):
        """ (float) 3-winding transformer LTC center tap """
        return self.getData('LTCCENTER')
    @property
    def LTCCTRL(self):
        """ (str) 3-winding transformer ID string of the bus whose voltage magnitude is to be regulated by the LTC """
        return self.getData('LTCCTRL')
    @property
    def LTCSIDE(self):
        """ (int) 3-winding transformer LTC side: 1; 2; 0 """
        return self.getData('LTCSIDE')
    @property
    def LTCSTEP(self):
        """ (float) 3-winding transformer LTC step size """
        return self.getData('LTCSTEP')
    @property
    def LTCTYPE(self):
        """ (int) 3-winding transformer LTC type: 1- control voltage; 2- control MVAR """
        return self.getData('LTCTYPE')
    @property
    def MAXTAP(self):
        """ (float) 3-winding transformer LTC max tap """
        return self.getData('MAXTAP')
    @property
    def MAXVW(self):
        """ (float) 3-winding transformer LTC min controlled quantity limit """
        return self.getData('MAXVW')
    @property
    def MINTAP(self):
        """ (float) 3-winding transformer LTC min tap """
        return self.getData('MINTAP')
    @property
    def MINVW(self):
        """ (float) 3-winding transformer LTC controlled quantity limit """
        return self.getData('MINVW')
    @property
    def MVA1(self):
        """ (float) 3-winding transformer MVA1 """
        return self.getData('MVA1')
    @property
    def MVA2(self):
        """ (float) 3-winding transformer MVA2 """
        return self.getData('MVA2')
    @property
    def MVA3(self):
        """ (float) 3-winding transformer MVA3 """
        return self.getData('MVA3')
    @property
    def NAME(self):
        """ (str) 3-winding transformer name """
        return self.getData('NAME')
    @property
    def PRIORITY(self):
        """ (int) 3-winding transformer LTC adjustment priority 0-Normal, 1-Medieum, 2-High """
        return self.getData('PRIORITY')
    @property
    def PRITAP(self):
        """ (float) 3-winding transformer primary tap """
        return self.getData('PRITAP')
    @property
    def RG1(self):
        """ (float) 3-winding transformer Rg1 """
        return self.getData('RG1')
    @property
    def RG2(self):
        """ (float) 3-winding transformer Rg2 """
        return self.getData('RG2')
    @property
    def RG3(self):
        """ (float) 3-winding transformer Rg3 """
        return self.getData('RG3')
    @property
    def RGN(self):
        """ (float) 3-winding transformer Rgn """
        return self.getData('RGN')
    @property
    def RMG0(self):
        """ (float) 3-winding transformer RMG0 """
        return self.getData('RMG0')
    @property
    def RPM0(self):
        """ (float) 3-winding transformer RPM0 """
        return self.getData('RPM0')
    @property
    def RPS(self):
        """ (float) 3-winding transformer Rps """
        return self.getData('RPS')
    @property
    def RPS0(self):
        """ (float) 3-winding transformer R0ps """
        return self.getData('RPS0')
    @property
    def RPT(self):
        """ (float) 3-winding transformer Rpt """
        return self.getData('RPT')
    @property
    def RPT0(self):
        """ (float) 3-winding transformer R0pt """
        return self.getData('RPT0')
    @property
    def RSM0(self):
        """ (float) 3-winding transformer RSM0 """
        return self.getData('RSM0')
    @property
    def RST(self):
        """ (float) 3-winding transformer Rst """
        return self.getData('RST')
    @property
    def RST0(self):
        """ (float) 3-winding transformer R0st """
        return self.getData('RST0')
    @property
    def SECTAP(self):
        """ (float) 3-winding transformer secondary tap """
        return self.getData('SECTAP')
    @property
    def TERTAP(self):
        """ (float) 3-winding transformer tertiary tap """
        return self.getData('TERTAP')
    @property
    def XG1(self):
        """ (float) 3-winding transformer Xg1 """
        return self.getData('XG1')
    @property
    def XG2(self):
        """ (float) 3-winding transformer Xg2 """
        return self.getData('XG2')
    @property
    def XG3(self):
        """ (float) 3-winding transformer Xg3 """
        return self.getData('XG3')
    @property
    def XGN(self):
        """ (float) 3-winding transformer Xgn """
        return self.getData('XGN')
    @property
    def XMG0(self):
        """ (float) 3-winding transformer XMG0 """
        return self.getData('XMG0')
    @property
    def XPM0(self):
        """ (float) 3-winding transformer XPM0 """
        return self.getData('XPM0')
    @property
    def XPS(self):
        """ (float) 3-winding transformer Xps """
        return self.getData('XPS')
    @property
    def XPS0(self):
        """ (float) 3-winding transformer X0ps """
        return self.getData('XPS0')
    @property
    def XPT(self):
        """ (float) 3-winding transformer Xpt """
        return self.getData('XPT')
    @property
    def XPT0(self):
        """ (float) 3-winding transformer X0pt """
        return self.getData('XPT0')
    @property
    def XSM0(self):
        """ (float) 3-winding transformer XSM0 """
        return self.getData('XSM0')
    @property
    def XST(self):
        """ (float) 3-winding transformer Xst """
        return self.getData('XST')
    @property
    def XST0(self):
        """ (float) 3-winding transformer X0st """
        return self.getData('XST0')
    @property
    def Z0METHOD(self):
        """ (int) 3-winding transformer Z0 method 1-Short circuit impedance; 2-Classical T model impedance """
        return self.getData('Z0METHOD')
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of XFMR3
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of XFMR3
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP3(self):
        """ RLYGROUP3 of XFMR3
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP3').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of XFMR3 """
        return [self.RLYGROUP1,self.RLYGROUP2,self.RLYGROUP3]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of XFMR3 """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class SHIFTER(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ SHIFTER constructor (Exception if not found):
                SHIFTER(b1,b2,CID)   SHIFTER(b1,b2,'1')
                SHIFTER(GUID)        SHIFTER('{1bee3fcd-2663-4d2a-9bbe-9ff12f2240c1}'
                SHIFTER(STR)         SHIFTER("[SHIFTER] 4 'TENNESSEE' 132 kV-6 'NEVADA' 132 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'SHIFTER')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,SHIFTER)
            else:
                hnd = __findBr2Hnd__('SHIFTER',val1,val2,val3)
            __initFailCheck__(hnd,'SHIFTER',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('SHIFTER')
    @property
    def ANGMAX(self):
        """ (float) Phase shifter shift angle max """
        return self.getData('ANGMAX')
    @property
    def ANGMIN(self):
        """ (float) Phase shifter shift angle min """
        return self.getData('ANGMIN')
    @property
    def BASEMVA(self):
        """ (float) Phase shifter BaseMVA """
        return self.getData('BASEMVA')
    @property
    def BN1(self):
        """ (float) Phase shifter Bn1 """
        return self.getData('BN1')
    @property
    def BN2(self):
        """ (float) Phase shifter Bn2 """
        return self.getData('BN2')
    @property
    def BP1(self):
        """ (float) Phase shifter Bp1 """
        return self.getData('BP1')
    @property
    def BP2(self):
        """ (float) Phase shifter Bp2 """
        return self.getData('BP2')
    @property
    def BUS1(self):
        """ (BUS) Phase shifter bus 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) Phase shifter bus 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS(self):
        """ return [BUS] List Buses of SHIFTER """
        return [self.BUS1,self.BUS2]
    @property
    def BZ1(self):
        """ (float) Phase shifter Bz1 """
        return self.getData('BZ1')
    @property
    def BZ2(self):
        """ (float) Phase shifter Bz2 """
        return self.getData('BZ2')
    @property
    def CID(self):
        """ (str) Phase shifter circuit ID """
        return self.getData('CID')
    @property
    def CNTL(self):
        """ (int) Phase shifter control mode 0-Fixed, 1-automatically control real power flow  """
        return self.getData('CNTL')
    @property
    def DATEOFF(self):
        """ (str) Phase shifter out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Phase shifter in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Phase shifter in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def GN1(self):
        """ (float) Phase shifter Gn1 """
        return self.getData('GN1')
    @property
    def GN2(self):
        """ (float) Phase shifter Gn2 """
        return self.getData('GN2')
    @property
    def GP1(self):
        """ (float) Phase shifter Gp1 """
        return self.getData('GP1')
    @property
    def GP2(self):
        """ (float) Phase shifter Gp2 """
        return self.getData('GP2')
    @property
    def GZ1(self):
        """ (float) Phase shifter Gz1 """
        return self.getData('GZ1')
    @property
    def GZ2(self):
        """ (float) Phase shifter Gz2 """
        return self.getData('GZ2')
    @property
    def MVA1(self):
        """ (float) Phase shifter MVA1 """
        return self.getData('MVA1')
    @property
    def MVA2(self):
        """ (float) Phase shifter MVA2 """
        return self.getData('MVA2')
    @property
    def MVA3(self):
        """ (float) Phase shifter MVA3 """
        return self.getData('MVA3')
    @property
    def MWMAX(self):
        """ (float) Phase shifter MW max """
        return self.getData('MWMAX')
    @property
    def MWMIN(self):
        """ (float) Phase shifter MW min """
        return self.getData('MWMIN')
    @property
    def NAME(self):
        """ (str) Phase shifter name """
        return self.getData('NAME')
    @property
    def RN(self):
        """ (float) Phase shifter Rn """
        return self.getData('RN')
    @property
    def RP(self):
        """ (float) Phase shifter Rp """
        return self.getData('RP')
    @property
    def RZ(self):
        """ (float) Phase shifter Rz """
        return self.getData('RZ')
    @property
    def SHIFTANGLE(self):
        """ (float) Phase shifter shift angle """
        return self.getData('SHIFTANGLE')
    @property
    def XN(self):
        """ (float) Phase shifter Xn """
        return self.getData('XN')
    @property
    def XP(self):
        """ (float) Phase shifter Xp """
        return self.getData('XP')
    @property
    def XZ(self):
        """ (float) Phase shifter Xz """
        return self.getData('XZ')
    @property
    def ZCORRECTNO(self):
        """ (int) Phase shifter correct table number """
        return self.getData('ZCORRECTNO')
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of SHIFTER
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of SHIFTER
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of SHIFTER """
        return [self.RLYGROUP1,self.RLYGROUP2]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of SHIFTER """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class LINE(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ LINE constructor (Exception if not found):
                LINE(b1,b2,CID)   LINE(b1,b2,'1')
                LINE(GUID)        LINE('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                LINE(STR)         LINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'LINE')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,LINE)
            else:
                hnd = __findBr2Hnd__('LINE',val1,val2,val3)
            __initFailCheck__(hnd,'LINE',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('LINE')
    @property
    def B1(self):
        """ (float) Line B1 , in pu """
        return self.getData('B1')
    @property
    def B10(self):
        """ (float) Line B10, in pu """
        return self.getData('B10')
    @property
    def B2(self):
        """ (float) Line B2, in pu """
        return self.getData('B2')
    @property
    def B20(self):
        """ (float) Line B20, in pu """
        return self.getData('B20')
    @property
    def BUS1(self):
        """ (BUS) Line bus 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) Line bus 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of LINE """
        return [self.BUS1,self.BUS2]
    @property
    def CID(self):
        """ (str) Line circuit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Line out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Line in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Line in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def G1(self):
        """ (float) Line G1, in pu """
        return self.getData('G1')
    @property
    def G10(self):
        """ (float) Line G10, in pu """
        return self.getData('G10')
    @property
    def G2(self):
        """ (float) Line G2, in pu """
        return self.getData('G2')
    @property
    def G20(self):
        """ (float) Line G20, in pu """
        return self.getData('G20')
    @property
    def I2T(self):
        """ (float) I^2T rating of line in ampere^2 Sec. """
        return self.getData('I2T')
    @property
    def LN(self):
        """ (float) Line length """
        return self.getData('LN')
    @property
    def MULINE(self):
        """ (MULINE) Line mutual pair """
        return self.getData('MULINE')
    @property
    def NAME(self):
        """ (str) Line name """
        return self.getData('NAME')
    @property
    def R(self):
        """ (float) Line R, in pu """
        return self.getData('R')
    @property
    def R0(self):
        """ (float) Line Ro, in pu """
        return self.getData('R0')
    @property
    def RATG(self):
        """ [float]*4 Line ratings """
        return self.getData('RATG')
    @property
    def METEREDEND(self):
        """ (int) Line meteted flag: 1- at Bus1; 2-at Bus2; 0-line is in a single area; """
        return self.getData('METEREDEND')
    @property
    def TYPE(self):
        """ (str) Line table type in ['','Delta Dove','Horz. Dove','Vert. Dove] """
        return self.getData('TYPE')
    @property
    def UNIT(self):
        """ (str) Line length unit in [ft,kt,mi,m,km] """
        return self.getData('UNIT')
    @property
    def X(self):
        """ (float) Line X, in pu """
        return self.getData('X')
    @property
    def X0(self):
        """ (float) Line Xo, in pu """
        return self.getData('X0')
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of LINE
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of LINE
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of LINE """
        return [self.RLYGROUP1,self.RLYGROUP2]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of LINE """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class DCLINE2(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ DCLINE2 constructor (Exception if not found):
                DCLINE2(b1,b2,CID)   DCLINE2(b1,b2,'1')
                DCLINE2(GUID)        DCLINE2('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                DCLINE2(STR)         DCLINE2("[DCLINE2] 5 'FIELDALE' 132 kV-'BUS2' 132 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'DCLINE2')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,DCLINE2)
            else:
                hnd = __findBr2Hnd__('DCLINE2',val1,val2,val3)
            __initFailCheck__(hnd,'DCLINE2',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('DCLINE2')
    @property
    def ANGMAX(self):
        """ [float]*2 DC Line Angle max """
        return self.getData('ANGMAX')
    @property
    def ANGMIN(self):
        """ [float]*2 DC Line Angle min """
        return self.getData('ANGMIN')
    @property
    def BRIDGE(self):
        """ [int]*2 DC Line No. of bridges """
        return self.getData('BRIDGE')
    @property
    def BUS1(self):
        """ (BUS) DC Line BUS 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) DC Line BUS 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of DCLINE2 """
        return [self.BUS1,self.BUS2]
    @property
    def CID(self):
        """ (string) DC Line circuit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (string) DC Line out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (string) DC Line in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) DC Line in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def LINER(self):
        """ (float) DC Line R(Ohm) """
        return self.getData('LINER')
    @property
    def MARGIN(self):
        """ (float) DC Line Margin """
        return self.getData('MARGIN')
    @property
    def MODE(self):
        """ (int) DC Line control mode """
        return self.getData('MODE')
    @property
    def MVA(self):
        """ [float]*2 DC Line MVA rating """
        return self.getData('MVA')
    @property
    def NAME(self):
        """ (string) DC Line name """
        return self.getData('NAME')
    @property
    def NOMKV(self):
        """ [float]*2 DC Line nomibal KV on dc side """
        return self.getData('NOMKV')
    @property
    def R(self):
        """ [float]*2 DC Line XFMR R in pu """
        return self.getData('R')
    @property
    def TAP(self):
        """ [float]*2 DC Line Tap """
        return self.getData('TAP')
    @property
    def TAPMAX(self):
        """ [float]*2 DC Line Tap max  """
        return self.getData('TAPMAX')
    @property
    def TAPMIN(self):
        """ [float]*2 DC Line Tap min """
        return self.getData('TAPMIN')
    @property
    def TAPSTEP(self):
        """ [float]*2 DC Line Tap step size """
        return self.getData('TAPSTEP')
    @property
    def TARGET(self):
        """ (float) DC Line target MW """
        return self.getData('TARGET')
    @property
    def METEREDEND(self):
        """ (int) DC Line meteted flag: 1- at Bus1; 2-at Bus2; 0-line is in a single area; """
        return self.getData('METEREDEND')
    @property
    def VSCHED(self):
        """ (float) DC Line Schedule DC volts in pu """
        return self.getData('VSCHED')
    @property
    def X(self):
        """ [float]*2 DC Line XFMR X in pu """
        return self.getData('X')
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of DCLINE2 """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class MULINE(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ MULINE constructor (Exception if not found):
                MULINE(GUID)   MULINE('{03d127b7-2363-4642-b70e-0fb6e9636813}'
                MULINE(STR)    MULINE("[MUPAIR] 6 'NEVADA' 132 kV-'BUS0' 132 kV 1|8 'REUSENS' 132 kV-28 'ARIZONA' 132 kV 1") """
        __checkArgs__([val1,hnd],'MULINE')
        if hnd is None:
            hnd = __findObjHnd__(val1,MULINE)
            __initFailCheck__(hnd,'MULINE',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('MULINE')
    @property
    def FROM1(self):
        """ [float]*5 Mutual coupling pair line 1 From percent """
        return self.getData('FROM1')
    @property
    def FROM2(self):
        """ [float]*5 Mutual coupling pair line 2 From percent """
        return self.getData('FROM2')
    @property
    def LINE1(self):
        """ (LINE) Mutual coupling pair line 1 """
        try:
            return LINE(hnd=self.getData('LINE1').__hnd__)
        except:
            return None
    @property
    def LINE2(self):
        """ (LINE) Mutual coupling pair line 2 """
        try:
            return LINE(hnd=self.getData('LINE2').__hnd__)
        except:
            return None
    @property
    def R(self):
        """ [float]*5 Mutual coupling pair R """
        return self.getData('R')
    @property
    def TO1(self):
        """ [float]*5 Mutual coupling pair line1 To percent """
        return self.getData('TO1')
    @property
    def TO2(self):
        """ [float]*5 Mutual coupling pair line2 To percent """
        return self.getData('TO2')
    @property
    def X(self):
        """ [float]*5 Mutual coupling pair X """
        return self.getData('X')
#
class SERIESRC(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ SERIESRC constructor (Exception if not found):
                SERIESRC(b1,b2,CID)   SERIESRC(b1,b2,'1')
                SERIESRC(GUID)        SERIESRC('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                SERIESRC(STR)         SERIESRC("[SERIESRC] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'SERIESRC')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,SERIESRC)
            else:
                hnd = __findBr2Hnd__('SERIESRC',val1,val2,val3)
            __initFailCheck__(hnd,'SERIESRC',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('SERIESRC')
    @property
    def BUS1(self):
        """ (BUS) Series capacitor/reactor BUS 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) Series capacitor/reactor BUS 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of SERIESRC """
        return [self.BUS1,self.BUS2]
    @property
    def CID(self):
        """ (str) Series capacitor/reactor circuit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Series capacitor/reactor out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Series capacitor/reactor in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Series capacitor/reactor in-service flag: 1- active; 2- out-of-service; 3- bypassed """
        return self.getData('FLAG')
    @property
    def IPR(self):
        """ (float) Series capacitor/reactor protective level current """
        return self.getData('IPR')
    @property
    def NAME(self):
        """ (str) Series capacitor/reactor name """
        return self.getData('NAME')
    @property
    def R(self):
        """ (float) Series capacitor/reactor R """
        return self.getData('R')
    @property
    def SCOMP(self):
        """ (int) Series capacitor/reactor bypassed flag 1- no bypassed; 2-yes bypassed  """
        return self.getData('SCOMP')
    @property
    def X(self):
        """ (float) Series capacitor/reactor X """
        return self.getData('X')
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of SERIESRC
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of SERIESRC
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of SERIESRC """
        return [self.RLYGROUP1,self.RLYGROUP2]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of SERIESRC """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class SWITCH(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,val3=None,hnd=None):
        """ SWITCH constructor (Exception if not found):
                SWITCH(b1,b2,CID)   SWITCH(b1,b2,'1')
                SWITCH(GUID)        SWITCH('{b9467f64-1833-431d-8bf2-28daca231d8a}'
                SWITCH(STR)         SWITCH("[SWITCH] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") """
        __checkArgs__([val1,val2,val3,hnd],'SWITCH')
        if hnd is None:
            if val2 is None and val3 is None:
                hnd = __findObjHnd__(val1,SWITCH)
            else:
                hnd = __findBr2Hnd__('SWITCH',val1,val2,val3)
            __initFailCheck__(hnd,'SWITCH',[val1,val2,val3])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,val3=None,hnd=None):
        __oldinit214__('SWITCH')
    @property
    def BUS1(self):
        """ (BUS) Switch BUS 1 """
        return BUS(hnd=self.getData('BUS1').__hnd__)
    @property
    def BUS2(self):
        """ (BUS) Switch BUS 2 """
        return BUS(hnd=self.getData('BUS2').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of SWITCH """
        return [self.BUS1,self.BUS2]
    @property
    def CID(self):
        """ (str) Switch Id """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Switch out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Switch in service date """
        return self.getData('DATEON')
    @property
    def DEFAULT(self):
        """ (int) Switch default position flag: 1- normaly open; 2- normaly close; 0-Not defined """
        return self.getData('DEFAULT')
    @property
    def FLAG(self):
        """ (int) Switch in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def NAME(self):
        """ (str) Switch name """
        return self.getData('NAME')
    @property
    def RATING(self):
        """ (float) Switch current rating """
        return self.getData('RATING')
    @property
    def STAT(self):
        """ (int) Switch position flag: 7- close; 0- open """
        return self.getData('STAT')
    @property
    def RLYGROUP1(self):
        """ RLYGROUP1 of SWITCH
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None
    @property
    def RLYGROUP2(self):
        """ RLYGROUP2 of SWITCH
            None if not found """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None
    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List RLYGROUPs of SWITCH """
        return [self.RLYGROUP1,self.RLYGROUP2]
    @property
    def TERMINAL(self):
        """ [TERMINAL] List TERMINALs of SWITCH """
        res = []
        for ti in __get_OBJTERMINAL__(self):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
class LOAD(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ LOAD constructor (Exception if not found):
                LOAD(GUID)        LOAD('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                LOAD(STR)         LOAD("[LOAD] 17 'WASHINGTON' 33 kV")
                LOAD(name,kV)     LOAD('WASHINGTON',33)  """
        __checkArgs__([val1,val2,hnd],'LOAD')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,LOAD)
            else:
                try:
                    hnd = BUS(val1,val2).LOAD[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'LOAD',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('LOAD')
    @property
    def BUS(self):
        """ (BUS) Load bus """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def FLAG(self):
        """ (int) Load in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def P(self):
        """ (float) Total load MW (load flow solution) """
        return self.getData('P')
    @property
    def Q(self):
        """ (float) Total load MVAR (load flow solution) """
        return self.getData('Q')
    @property
    def UNGROUNDED(self):
        """ (int) Load UnGrounded 1-UnGrounded 0-Grounded """
        return self.getData('UNGROUNDED')
    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List LOAD. units in LOAD """
        res = []
        for ti in self.BUS.LOADUNIT:
            res.append(LOADUNIT(hnd=ti.__hnd__))
        return res
#
class LOADUNIT(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ LOADUNIT constructor (Exception if not found):
                LOADUNIT(GUID)        LOADUNIT('{61e0ebc9-96ef-4cec-b8ed-d91ece2d9b3c}'
                LOADUNIT(STR)         LOADUNIT("[LOADUNIT]  1@17 'WASHINGTON' 33 kV") """
        __checkArgs__([val1,hnd],'LOADUNIT')
        if hnd is None:
            hnd = __findObjHnd__(val1,LOADUNIT)
            __initFailCheck__(hnd,'LOADUNIT',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('LOADUNIT')
    @property
    def CID(self):
        """ (str) Load unit circuit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Load unit out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Load unit in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Load unit in-service flag: 1-active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def LOAD(self):
        """ (LOAD) Load unit load """
        return LOAD(hnd=self.getData('LOAD').__hnd__)
    @property
    def MVAR(self):
        """ [float]*3 Load unit MVARs: [const. P, const. I, const. Z] """
        return self.getData('MVAR')
    @property
    def MW(self):
        """ [float]*3 Load unit MWs: [const. P, const. I, const. Z] """
        return self.getData('MW')
    @property
    def P(self):
        """ (float) Load unit MW (load flow solution) """
        return self.getData('P')
    @property
    def Q(self):
        """ (float) Load unit MVAR (load flow solution) """
        return self.getData('Q')
    @property
    def BUS(self):
        """ (BUS) Bus of Load unit """
        return BUS(hnd=self.LOAD.BUS.__hnd__)
#
class SHUNT(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ SHUNT constructor (Exception if not found):
                SHUNT(GUID)        SHUNT('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                SHUNT(STR)         SHUNT("[SHUNT] 21 'IOWA' 33 kV")
                SHUNT(name,kV)     SHUNT('CLAYTOR',132)  """
        __checkArgs__([val1,val2,hnd],'SHUNT')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,SHUNT)
            else:
                try:
                    hnd = BUS(val1,val2).SHUNT[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'SHUNT',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('SHUNT')
    @property
    def BUS(self):
        """ (BUS) Shunt Bus """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def FLAG(self):
        """ (int) Shunt in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List SHUNT. units in SHUNT """
        res = []
        for ti in self.BUS.SHUNTUNIT:
            res.append(SHUNTUNIT(hnd=ti.__hnd__))
        return res
#
class SHUNTUNIT(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ SHUNTUNIT constructor (Exception if not found):
                SHUNTUNIT(GUID)        SHUNTUNIT('{61e0ebc9-96ef-4cec-b8ed-d91ece2d9b3c}'
                SHUNTUNIT(STR)         SHUNTUNIT("[CAPUNIT]  1@21 'IOWA' 33 kV") """
        __checkArgs__([val1,hnd],'SHUNTUNIT')
        if hnd is None:
            hnd = __findObjHnd__(val1,SHUNTUNIT)
            __initFailCheck__(hnd,'SHUNTUNIT',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('SHUNTUNIT')
    @property
    def B(self):
        """ (float) Shunt unit succeptance (positive sequence) """
        return self.getData('B')
    @property
    def B0(self):
        """ (float) Shunt unit succeptance (zero sequence) """
        return self.getData('B0')
    @property
    def CID(self):
        """ (str) Shunt unit ID """
        return self.getData('CID')
    @property
    def DATEOFF(self):
        """ (str) Shunt unit out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Shunt unit in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Shunt unit in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def G(self):
        """ (float) Shunt unit conductance (positive sequence) """
        return self.getData('G')
    @property
    def G0(self):
        """ (float) Shunt unit conductance (zero sequence) """
        return self.getData('G0')
    @property
    def SHUNT(self):
        """ (SHUNT) Shunt unit of shunt """
        return SHUNT(hnd=self.getData('SHUNT').__hnd__)
    @property
    def TX3(self):
        """ (int) Shunt unit 3-winding transformer flag: 1-true;0-false"""
        return self.getData('TX3')
    @property
    def BUS(self):
        """ (BUS) Bus of Shunt unit """
        return BUS(hnd=self.SHUNT.BUS.__hnd__)
#
class SVD(__DATAABSTRACT__):
    def __init__(self,val1=None,val2=None,hnd=None):
        """ SVD constructor (Exception if not found):
                SVD(GUID)        SVD('{124fe30b-2b6a-4ac7-9e16-0d4da642a6e9}'
                SVD(STR)         SVD("[SVD] 3 'TEXAS' 132 kV")
                SVD(name,kV)     SVD('TEXAS',132)  """
        __checkArgs__([val1,val2,hnd],'SVD')
        if hnd is None:
            if val2 is None:
                hnd = __findObjHnd__(val1,SVD)
            else:
                try:
                    hnd = BUS(val1,val2).SVD[0].__hnd__
                except:
                    hnd = 0
            __initFailCheck__(hnd,'SVD',[val1,val2])
        super().__init__(hnd)
    #
    def init(val1=None,val2=None,hnd=None):
        __oldinit214__('SVD')
    @property
    def B(self):
        """ [float]*8 SVD increment admitance """
        return self.getData('B')
    @property
    def B0(self):
        """ [float]*8 SVD increment zero admitance """
        return self.getData('B0')
    @property
    def BUS(self):
        """ (BUS) SVD Bus """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def B_USE(self):
        """ (float) SVD admitance in use """
        return self.getData('B_USE')
    @property
    def CNTBUS(self):
        """ (BUS) SVD controled Bus """
        return BUS(hnd=self.getData('CNTBUS').__hnd__)
    @property
    def DATEOFF(self):
        """ (str) SVD out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) SVD in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) SVD in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def MODE(self):
        """ (int) SVD control mode: 0-Fixed; 1-Discrete; 2-Continous """
        return self.getData('MODE')
    @property
    def STEP(self):
        """ [int]*8 SVD number of step """
        return self.getData('STEP')
    @property
    def VMAX(self):
        """ (float) SVD max V """
        return self.getData('VMAX')
    @property
    def VMIN(self):
        """ (float) SVD min V """
        return self.getData('VMIN')
#
class BREAKER(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ BREAKER constructor (Exception if not found):
                BREAKER(GUID)        BREAKER('{369fce04-353b-4c81-8e8e-9c4d97784206}'
                BREAKER(STR)         BREAKER("[BREAKER]  1E82A@ 6 'NEVADA' 132 kV") """
        __checkArgs__([val1,hnd],'BREAKER')
        if hnd is None:
            hnd = __findObjHnd__(val1,BREAKER)
            __initFailCheck__(hnd,'BREAKER',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('BREAKER')
    @property
    def BUS(self):
        """ (BUS) Breaker Bus """
        return BUS(hnd=self.getData('BUS').__hnd__)
    @property
    def CPT1(self):
        """ (float) Breaker contact parting time for group 1 (cycles) """
        return self.getData('CPT1')
    @property
    def CPT2(self):
        """ (float) Breaker contact parting time for group 2 (cycles) """
        return self.getData('CPT2')
    @property
    def FLAG(self):
        """ (int) Breaker In-service flag: 1-true; 2-false; """
        return self.getData('FLAG')
    @property
    def G1OUTAGES(self):
        """ [EQUIPMENT]*10 Breaker protected equipment group 1 List of additional outage """
        return self.getData('G1OUTAGES')
    @property
    def G2OUTAGES(self):
        """ [EQUIPMENT]*10 Breaker protected equipment group 2 List of additional outage """
        return self.getData('G2OUTAGES')
    @property
    def GROUPTYPE1(self):
        """ (int) Breaker group 1 interrupting current: 1-max current; 0-group current """
        return self.getData('GROUPTYPE1')
    @property
    def GROUPTYPE2(self):
        """ (int) Breaker group 1 interrupting current: 1-max current; 0-group current """
        return self.getData('GROUPTYPE2')
    @property
    def INTRATING(self):
        """ (float) Breaker interrupting rating """
        return self.getData('INTRATING')
    @property
    def INTRTIME(self):
        """ (float) Breaker interrupting time (cycles) """
        return self.getData('INTRTIME')
    @property
    def K(self):
        """ (float) Breaker kV range factor """
        return self.getData('K')
    @property
    def MRATING(self):
        """ (float) Breaker momentary rating """
        return self.getData('MRATING')
    @property
    def NACD(self):
        """ (float) Breaker no-ac-decay ratio """
        return self.getData('NACD')
    @property
    def NAME(self):
        """ (str) Breaker name (ID) """
        return self.getData('NAME')
    @property
    def NODERATE(self):
        """ (int) Breaker do not derate in reclosing operation flag: 1-true; 0-false; """
        return self.getData('NODERATE')
    @property
    def OBJLST1(self):
        """ [EQUIPMENT]*10 Breaker protected equipment group 1 List of equipment """
        return self.getData('OBJLST1')
    @property
    def OBJLST2(self):
        """ [EQUIPMENT]*10 Breaker protected equipment group 2 List of equipment """
        return self.getData('OBJLST2')
    @property
    def OPKV(self):
        """ (float) Breaker operating kV """
        return self.getData('OPKV')
    @property
    def OPS1(self):
        """ (int) Breaker total operations for group 1 """
        return self.getData('OPS1')
    @property
    def OPS2(self):
        """ (int) Breaker total operations for group 2 """
        return self.getData('OPS2')
    @property
    def RATEDKV(self):
        """ (float) Breaker max design kV """
        return self.getData('RATEDKV')
    @property
    def RATINGTYPE(self):
        """ (int) Breaker rating type: 0- symmetrical current basis;1- total current basis; 2- IEC """
        return self.getData('RATINGTYPE')
    @property
    def RECLOSE1(self):
        """ [float]*3 Breaker reclosing intervals for group 1 (s) """
        return self.getData('RECLOSE1')
    @property
    def RECLOSE2(self):
        """ [float]*3 Breaker reclosing intervals for group 2 (s) """
        return self.getData('RECLOSE2')
#
class RLYGROUP(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYGROUP constructor (Exception if not found):
                RLYGROUP(GUID)  RLYGROUP('{369fce04-353b-4c81-8e8e-9c4d97784206}'
                RLYGROUP(STR)   RLYGROUP("[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYGROUP')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYGROUP)
            __initFailCheck__(hnd,'RLYGROUP',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYGROUP')
    @property
    def BACKUP(self):
        """ [RLYGROUP] List of RLYGROUP-Downstream (this RLYGROUP backups) """
        res = []
        for h1 in self.getData('BACKUP'):
            res.append(RLYGROUP(hnd=h1.__hnd__))
        return res
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) Relay group equipment located on """
        return self.getData('EQUIPMENT')
    @property
    def FLAG(self):
        """ (int) Relay group in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def INTRPTIME(self):
        """ (float) Relay group interrupting time (cycles) """
        return self.getData('INTRPTIME')
    @property
    def LOGICRECL(self):
        """ (SCHEME) Relay group reclose logic scheme
            None if not found """
        try:
            return SCHEME(hnd=self.getData('LOGICRECL').__hnd__)
        except:
            return None
    @property
    def LOGICTRIP(self):
        """ (SCHEME) Relay group trip logic scheme
            None if not found """
        try:
            return SCHEME(hnd=self.getData('LOGICTRIP').__hnd__)
        except:
            return None
    @property
    def NOTE(self):
        """ (str) Relay group annotation """
        return self.getData('NOTE')
    @property
    def OPFLAG(self):
        """ (int) Relay group total operations """
        return self.getData('OPFLAG')
    @property
    def PRIMARY(self):
        """ [RLYGROUP] List of RLYGROUP-Upstream (backups for this RLYGROUP) """
        res = []
        for h1 in self.getData('PRIMARY'):
            res.append(RLYGROUP(hnd=h1.__hnd__))
        return res
    @property
    def RECLSRTIME(self):
        """ [float]*4 Relay group reclosing intervals """
        return self.getData('RECLSRTIME')
    @property
    def RLYOCG(self):
        """ [RLYOCG] List OC Ground relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYOCG'):
            res.append(RLYOCG(hnd=h1))
        return res
    @property
    def RLYOCP(self):
        """ [RLYOCP] List OC Phase relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYOCP'):
            res.append(RLYOCP(hnd=h1))
        return res
    @property
    def FUSE(self):
        """ [FUSE] List Fuses that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'FUSE'):
            res.append(FUSE(hnd=h1))
        return res
    @property
    def RLYDSG(self):
        """ [RLYDSG] List DS Ground relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYDSG'):
            res.append(RLYDSG(hnd=h1))
        return res
    @property
    def RLYDSP(self):
        """ [RLYDSP] List DS Phase relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYDSP'):
            res.append(RLYDSP(hnd=h1))
        return res
    @property
    def RLYD(self):
        """ [RLYD] List Differential relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYD'):
            res.append(RLYD(hnd=h1))
        return res
    @property
    def RLYV(self):
        """ [RLYV] List Voltage relays that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RLYV'):
            res.append(RLYV(hnd=h1))
        return res
    @property
    def RECLSR(self):
        """ [RECLSR] List Reclosers that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'RECLSR'):
            res.append(RECLSR(hnd=h1))
        return res
    @property
    def RELAY(self):
        """ [RELAY] List All Relays that is attached to RLYGROUP """
        return __getRLYGROUP_RLY__(self)
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYGROUP
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'BUS'):
            res.append(BUS(hnd=h1))
        return res
    @property
    def SCHEME(self):
        """ [SCHEME] List SCHEMEs that is attached to RLYGROUP """
        res = []
        for h1 in __getRLYGROUP_OBJ__(self,'SCHEME'):
            res.append(SCHEME(hnd=h1))
        return res
    @property
    def RLYOC(self):
        """ [RLYOCG+RLYOCP] List OC Relays that is attached to RLYGROUP """
        res = self.RLYOCG
        res.extend(self.RLYOCP)
        return res
    @property
    def RLYDS(self):
        """ [RLYDSG+RLYDSP] List DS Relays that is attached to RLYGROUP """
        res = self.RLYDSG
        res.extend(self.RLYDSP)
        return res
    @property
    def TERMINAL(self):
        """ (TERMINAL) TERMINAL of RLYGROUP """
        try:
            return TERMINAL(hnd=self.getData('TERMINAL').__hnd__)
        except:
            return None
#
class RLYOCG(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYOCG constructor (Exception if not found):
                RLYOCG(GUID)  RLYOCG('{b4e006f6-6ccb-497e-8828-8938df46bd4a}'
                RLYOCG(STR)   RLYOCG("[OCRLYG]  NV-G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYOCG')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYOCG)
            __initFailCheck__(hnd,'RLYOCG',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYOCG')
    @property
    def ASSETID(self):
        """ (str) OC ground relay asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) OC ground relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) OC ground relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) OC ground relay in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) OC ground relay ID """
        return self.getData('ID')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) OC ground relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYOCG
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYOCG located on """
        return self.RLYGROUP.EQUIPMENT
    #
    def getSetting(self,sPara=None):
        """ (dict/) Setting (sPara) of this RLYOCG """
        return __getRelaySetting__(self,sPara)
    #
    def changeSetting(self,sPara,val):
        """ (None) change Setting with (sPara,val) of this RLYOCG """
        return __changeRLYSetting__(self,sPara,val)
    #
    def printSetting(self):
        """ (None) Print to stdout all Setting of this RLYOCG """
        return print(__printRLYSetting__(self))
#
class RLYOCP(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYOCP constructor (Exception if not found):
                RLYOCP(GUID)  RLYOCP('{fd8d6203-9db9-4e2a-9d6a-d6f3b150d5bb}'
                RLYOCP(STR)   RLYOCP("[OCRLYP]  NV-P1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYOCP')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYOCP)
            __initFailCheck__(hnd,'RLYOCP',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYOCP')
    @property
    def ASSETID(self):
        """ (str) OC phase relay asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) OC phase relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) OC phase relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) OC phase relay in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) OC phase relay ID """
        return self.getData('ID')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) OC phase relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYOCP
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYOCP located on """
        return self.RLYGROUP.EQUIPMENT
    #
    def getSetting(self,sPara=None):
        """ (dict/) Setting (sPara) of this RLYOCP """
        return __getRelaySetting__(self,sPara)
    #
    def changeSetting(self,sPara,val):
        """ (None) change Setting with (sPara,val) of this RLYOCP """
        return __changeRLYSetting__(self,sPara,val)
    #
    def printSetting(self):
        """ (None) Print to stdout all Setting of this RLYOCP """
        return print(__printRLYSetting__(self))
#
class FUSE(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ FUSE constructor (Exception if not found):
                FUSE(GUID)  FUSE('{d67118b8-bb8c-4f37-9765-bf3c49cb2cba}'
                FUSE(STR)   FUSE("[FUSE]  NV Fuse@6 'NEVADA' 132 kV-4 'TENNESSEE' 132 kV 1 P") """
        __checkArgs__([val1,hnd],'FUSE')
        if hnd is None:
            hnd = __findObjHnd__(val1,FUSE)
            __initFailCheck__(hnd,'FUSE',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('FUSE')
    @property
    def ASSETID(self):
        """ (str) Fuse asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) Fuse out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Fuse in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Fuse in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def FUSECURDIV(self):
        """ (float) Fuse current divider """
        return self.getData('FUSECURDIV')
    @property
    def FUSECVE(self):
        """ (int) Fuse Compute time using flag: 1- minimum melt; 2- Total clear """
        return self.getData('FUSECVE')
    @property
    def ID(self):
        """ (str) Fuse name(ID) """
        return self.getData('ID')
    @property
    def LIBNAME(self):
        """ (str) Fuse Library """
        return self.getData('LIBNAME')
    @property
    def PACKAGE(self):
        """ (int) Fuse Package option """
        return self.getData('PACKAGE')
    @property
    def RATING(self):
        """ (float) Fuse Rating """
        return self.getData('RATING')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Fuse relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def TIMEMULT(self):
        """ (float) Fuse time multiplier """
        return self.getData('TIMEMULT')
    @property
    def TYPE(self):
        """ (str) Fuse type """
        return self.getData('TYPE')
    @property
    def BUS(self):
        """ [BUS] List Buses of this FUSE
                BUS[0]          : Bus Local
                BUS[1],(BUS[2]) : Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this FUSE located on """
        return self.RLYGROUP.EQUIPMENT
#
class RLYDSG(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYDSG constructor (Exception if not found):
                RLYDSG(GUID)  RLYDSG('{a004766f-125d-47b6-b9cd-70cb254f2a59}'
                RLYDSG(STR)   RLYDSG("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYDSG')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYDSG)
            __initFailCheck__(hnd,'RLYDSG',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYDSG')
    @property
    def ASSETID(self):
        """ (str) DS ground relay asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) DS ground relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) DS ground relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) DS ground relay in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) DS ground relay name(ID) """
        return self.getData('ID')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) DS ground relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def TYPE(self):
        """ (str) DS ground relay type (ID2) """
        return self.getData('TYPE')
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYDSG
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYDSG located on """
        return self.RLYGROUP.EQUIPMENT
    #
    def getDSSettingName(self):
        """ [str] All specific setting name of this RLYDSG """
        return __getRelaySettingName__(self)
    #
    def getSetting(self,sPara=None):
        """ (dict/) Setting (sPara) of this RLYDSG """
        return __getRelaySetting__(self,sPara)
    #
    def changeSetting(self,sPara,val):
        """ (None) change Setting with (sPara,val) of this RLYDSG """
        return __changeRLYSetting__(self,sPara,val)
    #
    def printSetting(self):
        """ (None) Print to stdout all Setting of this RLYDSG """
        return __printRLYSetting__(self)
#
class RLYDSP(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYDSP constructor (Exception if not found):
                RLYDSP(GUID)  RLYDSP('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                RLYDSP(STR)   RLYDSP("[DSRLYP]  GCXTEST@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYDSP')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYDSP)
            __initFailCheck__(hnd,'RLYDSP',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYDSP')
    @property
    def ASSETID(self):
        """ (str) DS phase relay asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) DS phase relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) DS phase relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) DS phase relay in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) DS phase relay ID """
        return self.getData('ID')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) DS phase relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def TYPE(self):
        """ (str) DS phase relay type (ID2) """
        return self.getData('TYPE')
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYDSP
                BUS[0]         : Bus Local
                BUS[1],(BUS[2] : Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYDSP located on """
        return self.RLYGROUP.EQUIPMENT
    #
    def getDSSettingName(self):
        """ [str] All specific setting name  of this RLYDSP """
        return __getRelaySettingName__(self)
    #
    def getSetting(self,sPara=None):
        """ (dict/) Setting (sPara) of this RLYDSP """
        return __getRelaySetting__(self,sPara)
    #
    def changeSetting(self,sPara,val):
        """ (None) Change Setting with (sPara,val) of this RLYDSP """
        return __changeRLYSetting__(self,sPara,val)
    #
    def printSetting(self):
        """ (None) Print to stdout all Setting of this RLYDSP """
        return __printRLYSetting__(self)
#
class RLYD(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYD constructor (Exception if not found):
                RLYD(GUID)  RLYD('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                RLYD(STR)   RLYD("[DEVICEDIFF]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYD')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYD)
            __initFailCheck__(hnd,'RLYD',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYD')
    @property
    def ASSETID(self):
        """ (str) Differential relay asset ID """
        return self.getData('ASSETID')
    @property
    def CTGRP1(self):
        """ (RLYGROUP) Differential relay local current input 1 """
        return self.getData('CTGRP1')
    @property
    def CTR1(self):
        """ (float) Differential relay CTR1 """
        return self.getData('CTR1')
    @property
    def DATEOFF(self):
        """ (str) Differential relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Differential relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Differential relay in service: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) Differential relay ID (NAME) """
        return self.getData('ID')
    @property
    def IMIN3I0(self):
        """ (float) Differential relay minimum enable differential current (3I0) """
        return self.getData('IMIN3I0')
    @property
    def IMIN3I2(self):
        """ (float) Differential relay minimum enable differential current (3I2) """
        return self.getData('IMIN3I2')
    @property
    def IMINPH(self):
        """ (float) Differential relay minimum enable differential current (phase) """
        return self.getData('IMINPH')
    @property
    def PACKAGE(self):
        """ (int) Differential relay Package option """
        return self.getData('PACKAGE')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Differential relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def RMTE1(self):
        """ (EQUIPMENT) Differential relay remote device 1 """
        return self.getData('RMTE1')
    @property
    def RMTE2(self):
        """ (EQUIPMENT) Differential relay remote device 2 """
        return self.getData('RMTE2')
    @property
    def SGLONLY(self):
        """ (int) Differential relay signal only : 1-true; 0-false """
        return self.getData('SGLONLY')
    @property
    def TLCCV3I0(self):
        """ (str) Differential relay tapped load coordination curve (I0) """
        return self.getData('TLCCV3I0')
    @property
    def TLCCVI2(self):
        """ (str) Differential relay tapped load coordination curve (I2) """
        return self.getData('TLCCVI2')
    @property
    def TLCCVPH(self):
        """ (str) Differential relay tapped load coordination curve (phase) """
        return self.getData('TLCCVPH')
    @property
    def TLCTD3I0(self):
        """ (float) Differential relay tapped load coordination delay (I0) """
        return self.getData('TLCTD3I0')
    @property
    def TLCTDI2(self):
        """ (float) Differential relay tapped load coordination delay (I2) """
        return self.getData('TLCTDI2')
    @property
    def TLCTDPH(self):
        """ (float) Differential relay tapped load coordination delay (phase) """
        return self.getData('TLCTDPH')
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYD
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYD located on """
        return self.RLYGROUP.EQUIPMENT
#
class RLYV(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RLYV constructor (Exception if not found):
                RLYV(GUID)  RLYV('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                RLYV(STR)   RLYV("[DEVICEVR]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RLYV')
        if hnd is None:
            hnd = __findObjHnd__(val1,RLYV)
            __initFailCheck__(hnd,'RLYV',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RLYV')
    @property
    def ASSETID(self):
        """ (str) Voltage relay asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) Voltage relay out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Voltage relay in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Voltage relay in service: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) Voltage relay ID (NAME) """
        return self.getData('ID')
    @property
    def OPQTY(self):
        """ (int) Voltage relay operate on voltage option: 1-Phase-to-Neutral; 2- Phase-to-Phase; 3-3V0;4-V1;5-V2;6-VA;7-VB;8-VC;9-VBC;10-VAB;11-VCA """
        return self.getData('OPQTY')
    @property
    def OVCVR(self):
        """ (str) Voltage relay over-voltage element curve """
        return self.getData('OVCVR')
    @property
    def OVDELAYTD(self):
        """ (float) Voltage relay over-voltage delay """
        return self.getData('OVDELAYTD')
    @property
    def OVINST(self):
        """ (float) Voltage relay over-voltage instant pickup (V) """
        return self.getData('OVINST')
    @property
    def OVPICKUP(self):
        """ (float) Voltage relay over-voltage pickup (V) """
        return self.getData('OVPICKUP')
    @property
    def PACKAGE(self):
        """ (int) Voltage relay Package option """
        return self.getData('PACKAGE')
    @property
    def PTR(self):
        """ (float) Voltage relay PT ratio """
        return self.getData('PTR')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Voltage relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def SGLONLY(self):
        """ (int) Voltage relay signal only 0-No check; 1-Check over-voltage instant;2-Check over-voltage delay;4-Check under-voltage instant; 8-Check under-voltage delay  """
        return self.getData('SGLONLY')
    @property
    def UVCVR(self):
        """ (str) Voltage relay under-voltage element curve """
        return self.getData('UVCVR')
    @property
    def UVDELAYTD(self):
        """ (float) Voltage relay under-voltage delay """
        return self.getData('UVDELAYTD')
    @property
    def UVINST(self):
        """ (float) Voltage relay under-voltage instant pickup (V) """
        return self.getData('UVINST')
    @property
    def UVPICKUP(self):
        """ (float) Voltage relay under-voltage pickup (V) """
        return self.getData('UVPICKUP')
    @property
    def BUS(self):
        """ [BUS] List Buses of this RLYV
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RLYV located on """
        return self.RLYGROUP.EQUIPMENT
#
class RECLSR(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ RECLSR constructor (Exception if not found):
                RECLSR(GUID)  RECLSR('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                RECLSR(STR)   RECLSR("[RECLSRP]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'RECLSR')
        if hnd is None:
            hnd = __findObjHnd__(val1,RECLSR)
            __initFailCheck__(hnd,'RECLSR',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('RECLSR')
    @property
    def ASSETID(self):
        """ (str) Recloser AssetID """
        return self.getData('ASSETID')
    @property
    def BYTADD(self):
        """ (int) Recloser Time adder modifies """
        return self.getData('BYTADD')
    @property
    def BYTMULT(self):
        """ (int) Recloser Time multiplier modifies """
        return self.getData('BYTMULT')
    @property
    def DATEOFF(self):
        """ (str) Recloser out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Recloser in service date """
        return self.getData('DATEON')
    @property
    def FASTOPS(self):
        """ (int) Recloser number of fast operations """
        return self.getData('FASTOPS')
    @property
    def PH_FASTTYPE(self):
        """ (str) Recloser-Phase fast curve """
        return self.getData('PH_FASTTYPE')
    @property
    def GR_FASTTYPE(self):
        """ (str) Recloser-Ground fast curve """
        return self.getData('GR_FASTTYPE')
    @property
    def FLAG(self):
        """ (int) Recloser in service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) Recloser ID """
        return self.getData('ID')
    @property
    def PH_INST(self):
        """ (float) Recloser-Phase high current trip """
        return self.getData('PH_INST')
    @property
    def GR_INST(self):
        """ (float) Recloser-Ground high current trip """
        return self.getData('GR_INST')
    @property
    def PH_INSTDELAY(self):
        """ (float) Recloser-Phase high current trip delay """
        return self.getData('PH_INSTDELAY')
    @property
    def GR_INSTDELAY(self):
        """ (float) Recloser-Ground high current trip delay """
        return self.getData('GR_INSTDELAY')
    @property
    def INTRPTIME(self):
        """ (float) Recloser interrupting time (s) """
        return self.getData('INTRPTIME')
    @property
    def LIBNAME(self):
        """ (str) Recloser Library """
        return self.getData('LIBNAME')
    @property
    def PH_MINTIMEF(self):
        """ (float) Recloser-Phase fast curve minimum time """
        return self.getData('PH_MINTIMEF')
    @property
    def GR_MINTIMEF(self):
        """ (float) Recloser-Ground fast curve minimum time """
        return self.getData('GR_MINTIMEF')
    @property
    def PH_MINTIMES(self):
        """ (float) Recloser-Phase slow curve minimum time """
        return self.getData('PH_MINTIMES')
    @property
    def GR_MINTIMES(self):
        """ (float) Recloser-Ground slow curve minimum time """
        return self.getData('GR_MINTIMES')
    @property
    def PH_MINTRIPF(self):
        """ (float) Recloser-Phase fast curve pickup """
        return self.getData('PH_MINTRIPF')
    @property
    def GR_MINTRIPF(self):
        """ (float) Recloser-Ground fast curve pickup """
        return self.getData('GR_MINTRIPF')
    @property
    def PH_MINTRIPS(self):
        """ (float) Recloser-Phase slow curve pickup """
        return self.getData('PH_MINTRIPS')
    @property
    def GR_MINTRIPS(self):
        """ (float) Recloser-Ground slow curve pickup """
        return self.getData('GR_MINTRIPS')
    @property
    def RECLOSE1(self):
        """ (float) Recloser reclosing interval 1 """
        return self.getData('RECLOSE1')
    @property
    def RECLOSE2(self):
        """ (float) Recloser reclosing interval 2 """
        return self.getData('RECLOSE2')
    @property
    def RECLOSE3(self):
        """ (float) Recloser reclosing interval 3 """
        return self.getData('RECLOSE3')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Recloser relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def PH_SLOWTYPE(self):
        """ (str) Recloser-Phase slow curve """
        return self.getData('PH_SLOWTYPE')
    @property
    def GR_SLOWTYPE(self):
        """ (str) Recloser-Ground slow curve """
        return self.getData('GR_SLOWTYPE')
    @property
    def PH_TIMEADDF(self):
        """ (float) Recloser-Phase fast curve time adder """
        return self.getData('PH_TIMEADDF')
    @property
    def GR_TIMEADDF(self):
        """ (float) Recloser-Ground fast curve time adder """
        return self.getData('GR_TIMEADDF')
    @property
    def PH_TIMEADDS(self):
        """ (float) Recloser-Phase slow curve time adder """
        return self.getData('PH_TIMEADDS')
    @property
    def GR_TIMEADDS(self):
        """ (float) Recloser-Ground slow curve time adder """
        return self.getData('GR_TIMEADDS')
    @property
    def PH_TIMEMULTF(self):
        """ (float) Recloser-Phase fast curve time multiplier """
        return self.getData('PH_TIMEMULTF')
    @property
    def GR_TIMEMULTF(self):
        """ (float) Recloser-Ground fast curve time multiplier """
        return self.getData('GR_TIMEMULTF')
    @property
    def PH_TIMEMULTS(self):
        """ (float) Recloser-Phase slow curve time multiplier """
        return self.getData('PH_TIMEMULTS')
    @property
    def GR_TIMEMULTS(self):
        """ (float) Recloser-Ground slow curve time multiplier """
        return self.getData('GR_TIMEMULTS')
    @property
    def TOTALOPS(self):
        """ (int) Recloser total operations to locked out """
        return self.getData('TOTALOPS')
    @property
    def BUS(self):
        """ [BUS] List Buses of this RECLSR
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this RECLSR located on """
        return self.RLYGROUP.EQUIPMENT
#
class SCHEME(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ SCHEME constructor (Exception if not found):
                SCHEME(GUID)  SCHEME('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                SCHEME(STR)   SCHEME("[PILOT]  272-POTT-SEL421G@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") """
        __checkArgs__([val1,hnd],'SCHEME')
        if hnd is None:
            hnd = __findObjHnd__(val1,SCHEME)
            __initFailCheck__(hnd,'SCHEME',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('SCHEME')
    @property
    def ASSETID(self):
        """ (str) Logic scheme asset ID """
        return self.getData('ASSETID')
    @property
    def DATEOFF(self):
        """ (str) Logic scheme out of service date """
        return self.getData('DATEOFF')
    @property
    def DATEON(self):
        """ (str) Logic scheme in service date """
        return self.getData('DATEON')
    @property
    def FLAG(self):
        """ (int) Logic scheme in-service flag: 1- active; 2- out-of-service """
        return self.getData('FLAG')
    @property
    def ID(self):
        """ (str) Logic scheme ID """
        return self.getData('ID')
    @property
    def NAME(self):
        """ (str) Logic scheme name """
        return self.getData('NAME')
    @property
    def PL_LOGICEQU(self):
        """ (str) Logic scheme equation """
        return self.getData('PL_LOGICEQU')
    @property
    def PL_LOGICTERM(self):
        """ (str) Logic scheme variables details (one variable per line in the format: name=description) """
        return self.getData('PL_LOGICTERM')
    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Logic scheme relay group """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)
    @property
    def SGLONLY(self):
        """ (int) Logic scheme signal only """
        return self.getData('SIGNALONLY')
    @property
    def BUS(self):
        """ [BUS] List Buses of this SCHEME
                BUS[0]         : Bus Local
                BUS[1],(BUS[2]): Bus Remote(s) """
        return self.RLYGROUP.BUS
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that this SCHEME located on """
        return self.RLYGROUP.EQUIPMENT
#
class ZCORRECT(__DATAABSTRACT__):
    def __init__(self,val1=None,hnd=None):
        """ ZCORRECT constructor (Exception if not found):
                ZCORRECT(GUID)  ZCORRECT('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}'
                ZCORRECT(STR)   ZCORRECT("[ZCORRECT]...") """
        __checkArgs__([val1,hnd],'ZCORRECT')
        if hnd is None:
            hnd = __findObjHnd__(val1,ZCORRECT)
            __initFailCheck__(hnd,'ZCORRECT',[val1])
        super().__init__(hnd)
    #
    def init(val1=None,hnd=None):
        __oldinit214__('ZCORRECT')

#$ASPEN$ CODE GENERATED AUTOMATIC by OlxObj\genCodeOlxObj.py :STOP

class TERMINAL(__DATAABSTRACT__):
    def __init__(self,b1=None,b2=None,sType=None,CID=None,hnd=None):
        """ TERMINAL constructor (Exception if not found):
            TERMINAL(b1,b2,sType,CID)   TERMINAL(b1,b2,'LINE','1')
                b1,b2 : (BUS)
                sType : (str/obj) 'XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH'
                CID   : (str) """
        if hnd is None:
            sType = TERMINAL.__checkInputs__(b1,b2,sType,CID,'\nConstructor: TERMINAL(b1,b2,sType,CID)')
            hnd = __findTerminalHnd__(b1,b2,sType,CID)
            __initFailCheck__(hnd,'TERMINAL',[b1,b2,sType,CID])
        super().__init__(hnd)
    #
    def __checkInputs__(b1,b2,sType,CID,s0):
        if b1==None or b2==None or sType==None or CID==None:
            se='\nmissing required arguments in'+s0
            se+= '\n\tb1,b2 : (BUS)'
            se+= "\n\tsType : (str/obj) 'XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH'"
            se+= "\n\tCID   : (str) "
            raise Exception (se)
        if type(b1)!=BUS or type(b2)!=BUS:
            se =s0+'\n  b1,b2: BUS'
            se +="\n\tRequired : (BUS)"
            if type(b1)!=BUS:
                se +='\n\tb1 '+__getErrValue__(BUS,b1)
            if type(b2)!=BUS:
                se +='\n\tb2 '+__getErrValue__(BUS,b2)
            raise ValueError(se)
        if type(CID)!=str:
            se =s0+'\n  CID: Circuit ID'
            se +="\n\tRequired : (str)"
            se +='\n\t'+__getErrValue__(str,CID)
            raise ValueError(se)
        try:
            if sType in __OLXOBJ_EQUIPMENT__:
                return sType
            if sType in __OLXOBJ_EQUIPMENTO__:
                return sType.__name__
        except:
            pass
        se +=s0+'\n  sType: (str/obj) type equipment of terminal'
        se +="\n\tRequired : (str/obj) in "+str(__OLXOBJ_EQUIPMENTL__)
        se +='\n\t'+__getErrValue__(str,sType)
        raise ValueError(se)
    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that attached this TERMINAL """
        return __getTERMINAL_OBJ__(self,'EQUIPMENT')
    @property
    def BUS(self):
        """ [BUS] List Buses that attached this TERMINAL """
        res = []
        for bi in __getTERMINAL_OBJ__(self,'BUS'):
            res.append(BUS(hnd=bi.__hnd__))
        return res
    @property
    def BUS1(self):
        """ (BUS) BUS1 that attached this TERMINAL """
        h1 = __getDatai__(self.__hnd__,OlxAPIConst.BR_nBus1Hnd)
        return BUS(hnd=h1)
    @property
    def BUS23(self):
        """ [BUS] BUS2,3 that attached this TERMINAL """
        h2 = __getDatai__(self.__hnd__,OlxAPIConst.BR_nBus2Hnd)
        h3 = __getDatai__(self.__hnd__,OlxAPIConst.BR_nBus3Hnd)
        if h3 is None:
            return [BUS(hnd=h2)]
        return [BUS(hnd=h2),BUS(hnd=h3)]
    @property
    def RLYGROUP(self):
        """ [RLYGROUP]*3 List RLYGROUPs that attached to this TERMINAL
                RLYGROUP[0] : local RLYGROUP, =None if not found
                RLYGROUP[1] : opposite RLYGROUP =None if not found
                RLYGROUP[2] : opposite RLYGROUP =None if not found or not on XFMR3"""
        res = []
        for ri in __getTERMINAL_OBJ__(self,'RLYGROUP'):
            if ri is not None:
                res.append(RLYGROUP(hnd=ri.__hnd__) )
            else:
                res.append(None)
        return res
    @property
    def FLAG(self):
        """ (int) TERMINAL in-service flag: 1- active; 2- out-of-service """
        return __getTERMINAL_OBJ__(self,'FLAG')
    @property
    def CID(self):
        """ (str) TERMINAL ID """
        return self.EQUIPMENT.CID
    @property
    def OPPOSITE(self):
        """ [TERMINAL] List TERMINALs that opposite to this TERMINAL on the EQUIPMENT """
        res = []
        for ti in __getTERMINAL_OBJ__(self,'OPPOSITE'):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
    @property
    def REMOTE(self):
        """ [TERMINAL] List TERMINALs remote to this TERMINAL
                All taps are ignored.
                Close switches are included
                Out of service branches are ignored """
        res = []
        for ti in __getTERMINAL_OBJ__(self,'REMOTE'):
            res.append(TERMINAL(hnd=ti.__hnd__))
        return res
#
__OLXOBJ_CONST__ = {\
                'BUS'       :[OlxAPIConst.TC_BUS      ,'BUS_','[BUS] List of Buses'],
                'GEN'       :[OlxAPIConst.TC_GEN      ,'GE_' ,'[GEN] List of Generators'],
                'GENUNIT'   :[OlxAPIConst.TC_GENUNIT  ,'GU_' ,'[GENUNIT] List of Generator units'],
                'GENW3'     :[OlxAPIConst.TC_GENW3    ,'G3_' ,'[GENW3] List of Type-3 Wind Plants'],
                'GENW4'     :[OlxAPIConst.TC_GENW4    ,'G4_' ,'[GENW4] List of Converter-Interfaced Resources'],
                'CCGEN'     :[OlxAPIConst.TC_CCGEN    ,'CC_' ,'[CCGEN] List of Voltage Controlled Current Sources'],
                'XFMR'      :[OlxAPIConst.TC_XFMR     ,'XR_' ,'[XFMR] List of 2-Windings Transformers'],
                'XFMR3'     :[OlxAPIConst.TC_XFMR3    ,'X3_' ,'[XFMR3] List of 3-Windings Transformers'],
                'SHIFTER'   :[OlxAPIConst.TC_PS       ,'PS_' ,'[SHIFTER] List of Phase Shifters'],
                'LINE'      :[OlxAPIConst.TC_LINE     ,'LN_' ,'[LINE] List of Transmission Lines'],
                'DCLINE2'   :[OlxAPIConst.TC_DCLINE2  ,'DC_' ,'[DCLINE2] List of DC Lines'],
                'MULINE'    :[OlxAPIConst.TC_MU       ,'MU_' ,'[MULINE] List of Mutual Pairs'],
                'SERIESRC'  :[OlxAPIConst.TC_SCAP     ,'SC_' ,'[SERIESRC] List of Series capacitor/reactor'],
                'SWITCH'    :[OlxAPIConst.TC_SWITCH   ,'SW_' ,'[SWITCH] List of Switches'],
                'LOAD'      :[OlxAPIConst.TC_LOAD     ,'LD_' ,'[LOAD] List of Loads'],
                'LOADUNIT'  :[OlxAPIConst.TC_LOADUNIT ,'LU_' ,'[LOADUNIT] List of Load units'],
                'SHUNT'     :[OlxAPIConst.TC_SHUNT    ,'SH_' ,'[SHUNT] List of Shunts'],
                'SHUNTUNIT' :[OlxAPIConst.TC_SHUNTUNIT,'SU_' ,'[SHUNTUNIT] List of Shunt units'],
                'SVD'       :[OlxAPIConst.TC_SVD      ,'SV_' ,'[SVD] List of Switched Shunt'],
                'BREAKER'   :[OlxAPIConst.TC_BREAKER  ,'BK_' ,'[BREAKER] List of Breakers'],
                'RLYGROUP'  :[OlxAPIConst.TC_RLYGROUP ,'RG_' ,'[RLYGROUP] List of Relay groups'],
                'RLYOCG'    :[OlxAPIConst.TC_RLYOCG   ,'OG_' ,'[RLYOCG] List of OC Ground relays'],
                'RLYOCP'    :[OlxAPIConst.TC_RLYOCP   ,'OP_' ,'[RLYOCP] List of OC Phase relays'],
                'RLYOC'     :[0                       ,'OP_' ,'[RLYOCG/RLYOCP] List of OC relays'],
                'FUSE'      :[OlxAPIConst.TC_FUSE     ,'FS_' ,'[FUSE] List of Fuses'],
                'RLYDSG'    :[OlxAPIConst.TC_RLYDSG   ,'DG_' ,'[RLYDSG] List of DS Ground relays'],
                'RLYDSP'    :[OlxAPIConst.TC_RLYDSP   ,'DP_' ,'[RLYDSP] List of DS Phase relays'],
                'RLYDS'     :[0                       ,'DP_' ,'[RLYDSG/RLYDSP] List of DS relays'],
                'RLYD'      :[OlxAPIConst.TC_RLYD     ,'RD_' ,'[RLYD] List of Differential relays'],
                'RLYV'      :[OlxAPIConst.TC_RLYV     ,'RV_' ,'[RLYV] List of Voltage relays'],
                'RECLSR'    :[OlxAPIConst.TC_RECLSRP  ,'CP_' ,'[RECLSR] List of Reclosers'],
                'SCHEME'    :[OlxAPIConst.TC_SCHEME   ,'LS_' ,'[SCHEME] List of Logic schemes '],
                'ZCORRECT'  :[OlxAPIConst.TC_ZCORRECT ,'ZC_' ,'[ZCORRECT] List of Impedance correction table'],
                'TERMINAL'  :[OlxAPIConst.TC_BRANCH   ,'BR_' ,'[TERMINAL] List of TERMINAL']}
#
__OLXOBJ_CONST1__ = {\
                OlxAPIConst.TC_BUS    :['BUS',BUS],\
                OlxAPIConst.TC_GEN    :['GEN',GEN],\
                OlxAPIConst.TC_GENUNIT:['GENUNIT',GENUNIT],\
                OlxAPIConst.TC_GENW3  :['GENW3',GENW3],\
                OlxAPIConst.TC_GENW4  :['GENW4',GENW4],\
                OlxAPIConst.TC_CCGEN  :['CCGEN',CCGEN],\
                OlxAPIConst.TC_XFMR   :['XFMR',XFMR],\
                OlxAPIConst.TC_XFMR3  :['XFMR3',XFMR3],\
                OlxAPIConst.TC_PS     :['SHIFTER',SHIFTER],
                OlxAPIConst.TC_LINE   :['LINE',LINE],
                OlxAPIConst.TC_DCLINE2:['DCLINE2',DCLINE2],
                OlxAPIConst.TC_MU     :['MULINE',MULINE],
                OlxAPIConst.TC_SCAP   :['SERIESRC',SERIESRC],
                OlxAPIConst.TC_SWITCH :['SWITCH',SWITCH],
                OlxAPIConst.TC_LOAD   :['LOAD',LOAD],
                OlxAPIConst.TC_LOADUNIT:['LOADUNIT',LOADUNIT],
                OlxAPIConst.TC_SHUNT  :['SHUNT',SHUNT],
                OlxAPIConst.TC_SHUNTUNIT:['SHUNTUNIT',SHUNTUNIT],
                OlxAPIConst.TC_SVD    :['SVD',SVD],
                OlxAPIConst.TC_BREAKER:['BREAKER',BREAKER],
                OlxAPIConst.TC_RLYGROUP:['RLYGROUP',RLYGROUP],
                OlxAPIConst.TC_RLYOCG :['RLYOCG',RLYOCG],
                OlxAPIConst.TC_RLYOCP :['RLYOCP',RLYOCP],
                OlxAPIConst.TC_FUSE   :['FUSE',FUSE],
                OlxAPIConst.TC_RLYDSG :['RLYDSG',RLYDSG],
                OlxAPIConst.TC_RLYDSP :['RLYDSP',RLYDSP],
                OlxAPIConst.TC_RLYD   :['RLYD',RLYD],
                OlxAPIConst.TC_RLYV   :['RLYV',RLYV],
                OlxAPIConst.TC_RECLSRP:['RECLSR',RECLSR],
                OlxAPIConst.TC_SCHEME :['SCHEME',SCHEME],
                OlxAPIConst.TC_ZCORRECT:['ZCORRECT',ZCORRECT]}
#
__OLXOBJ_EQUIPMENT__  = {'XFMR3','XFMR','SHIFTER','LINE','DCLINE2','SERIESRC','SWITCH'}
__OLXOBJ_EQUIPMENTL__ = ['LINE','XFMR3','XFMR','SHIFTER','DCLINE2','SERIESRC','SWITCH']
__OLXOBJ_EQUIPMENTO__ = {XFMR3,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH}
__OLXOBJ_RELAY__  = {'RLYOCG','RLYOCP','FUSE','RLYDSG','RLYDSP','RLYD','RLYV','RECLSR'}
__OLXOBJ_RELAYO__ = {RLYOCG,RLYOCP,FUSE,RLYDSG,RLYDSP,RLYD,RLYV,RECLSR}
__OLXOBJ_BUS__    = {'BREAKER','CCGEN', 'DCLINE2','GEN','GENW3','GENW4','LINE','LOAD','SHIFTER','RLYGROUP',\
                 'SERIESRC','SHUNT','SVD','SWITCH','XFMR','XFMR3','LOADUNIT','SHUNTUNIT','GENUNIT','TERMINAL'}
__OLXOBJ_BUS1__  = {'LOAD','SHUNT','SVD','GEN','GENW3','GENW4','CCGEN'}
__OLXOBJ_LIST__  = ['BUS', 'GEN', 'GENUNIT', 'GENW3', 'GENW4', 'CCGEN', 'XFMR', 'XFMR3', 'SHIFTER', 'LINE', 'DCLINE2', 'MULINE', \
                 'SERIESRC', 'SWITCH', 'LOAD', 'LOADUNIT', 'SHUNT', 'SHUNTUNIT', 'SVD', 'BREAKER', 'RLYGROUP', 'RLYOCG', 'RLYOCP',\
                 'FUSE', 'RLYDSG', 'RLYDSP', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME','ZCORRECT']
__OLXOBJ_IFLT__ = [XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,GEN,GENUNIT,GENW3,GENW4,CCGEN,\
                   LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL]
#
__OLXOBJ_RLYSET__ = {}
__OLXOBJ_PARA__ = {}
__OLXOBJ_PARA__['SYS2'] = {\
                'BASEMVA'   :[OlxAPIConst.SY_dBaseMVA    ,'(float) NETWORK System MVA base'],\
                'COMMENT'   :[OlxAPIConst.SY_sFComment   ,'(str) NETWORK comment'],\
                'OBJCOUNT'  :[0                          ,'(dict) number of object in all network']}
__OLXOBJ_PARA__['SYS'] = {\
                'DCLINE'    :[OlxAPIConst.SY_nNODCLine2  ,'(int) NETWORK number of DC transmission lines'],
                'IED'       :[OlxAPIConst.SY_nNOIED      ,'(int) NETWORK number of IED'],
                'AREA'      :[OlxAPIConst.SY_nNOarea     ,'(int) NETWORK number of Area'],
                'BREAKER'   :[OlxAPIConst.SY_nNObreaker  ,'(int) NETWORK number of Breakers'],
                'BUS'       :[OlxAPIConst.SY_nNObus      ,'(int) NETWORK number of Buses'],
                'CCGEN'     :[OlxAPIConst.SY_nNOccgen    ,'(int) NETWORK number of Voltage controlled current sources'],
                'FUSE'      :[OlxAPIConst.SY_nNOfuse     ,'(int) NETWORK number of Fuses'],
                'GEN'       :[OlxAPIConst.SY_nNOgen      ,'(int) NETWORK number of Generators'],
                'GENUNIT'   :[OlxAPIConst.SY_nNOgenUnit  ,'(int) NETWORK number of Generator units'],
                'GENW3'     :[OlxAPIConst.SY_nNOgenW3    ,'(int) NETWORK number of Type-3 Wind Plant'],
                'GENW4'     :[OlxAPIConst.SY_nNOgenW4    ,'(int) NETWORK number of Converter-Interfaced Resource'],
                'LINE'      :[OlxAPIConst.SY_nNOline     ,'(int) NETWORK number of Transmission lines'],
                'LOAD'      :[OlxAPIConst.SY_nNOload     ,'(int) NETWORK number of Loads'],
                'LOADUNIT'  :[OlxAPIConst.SY_nNOloadUnit ,'(int) NETWORK number of Load units'],
                'LTC'       :[OlxAPIConst.SY_nNOltc      ,'(int) NETWORK number of Load Tap Changer of 2-winding transformers'],
                'LTC3'      :[OlxAPIConst.SY_nNOltc3     ,'(int) NETWORK number of Load Tap Changer of 3-winding transformers'],
                'MULINE'    :[OlxAPIConst.SY_nNOmuPair   ,'(int) NETWORK number of Mutual pair'],
                'SHIFTER'   :[OlxAPIConst.SY_nNOps       ,'(int) NETWORK number of Phase shifter'],
                'RECLSR'   :[OlxAPIConst.SY_nNOrecloserP,'(int) NETWORK number of Reclosers'],
                'RLYD'      :[OlxAPIConst.SY_nNOrlyD     ,'(int) NETWORK number of Differential relays'],
                'RLYDSG'    :[OlxAPIConst.SY_nNOrlyDSG   ,'(int) NETWORK number of DS ground relays'],
                'RLYDSP'    :[OlxAPIConst.SY_nNOrlyDSP   ,'(int) NETWORK number of DS phase relays'],
                'RLYGROUP'  :[OlxAPIConst.SY_nNOrlyGroup ,'(int) NETWORK number of Relay Group'],
                'RLYOCG'    :[OlxAPIConst.SY_nNOrlyOCG   ,'(int) NETWORK number of OC ground relays'],
                'RLYOCP'    :[OlxAPIConst.SY_nNOrlyOCP   ,'(int) NETWORK number of OC phase relays'],
                'RLYV'      :[OlxAPIConst.SY_nNOrlyV     ,'(int) NETWORK number of Voltage relays'],
                'SCHEME'    :[OlxAPIConst.SY_nNOscheme   ,'(int) NETWORK number of Logic schemes'],
                'SERIESRC'  :[OlxAPIConst.SY_nNOseriescap,'(int) NETWORK number of Series reactors/capacitors'],
                'SHUNT'     :[OlxAPIConst.SY_nNOshunt    ,'(int) NETWORK number of Shunts'],
                'SHUNTUNIT' :[OlxAPIConst.SY_nNOshuntUnit,'(int) NETWORK number of Shunt units'],
                'SVD'       :[OlxAPIConst.SY_nNOsvd      ,'(int) NETWORK number of Switched shunts'],
                'SWITCH'    :[OlxAPIConst.SY_nNOswitch   ,'(int) NETWORK number of Switches'],
                'XFMR'      :[OlxAPIConst.SY_nNOxfmr     ,'(int) NETWORK number of 2-winding transformers'],
                'XFMR3'     :[OlxAPIConst.SY_nNOxfmr3    ,'(int) NETWORK number of 3-winding transformers'],
                'ZCORRECT'  :[OlxAPIConst.SY_nNOzCorrect ,'(int) NETWORK number of Impedance correction table'],
                'ZONE'      :[OlxAPIConst.SY_nNOzone     ,'(int) NETWORK number of Zone']}
__OLXOBJ_PARA__['BUS'] = {
                'ANGLEP'  :[OlxAPIConst.BUS_dAngleP    ,'(float) Bus voltage angle (degree) from a power flow solution'],
                'KVP'     :[OlxAPIConst.BUS_dKVP       ,'(float) Bus voltage magnitude (kV) from a power flow solution'],
                'KV'      :[OlxAPIConst.BUS_dKVnominal ,'(float) Bus voltage nominal'],
                'SPCX'    :[OlxAPIConst.BUS_dSPCx      ,'(float) Bus state plane coordinate - X'],
                'SPCY'    :[OlxAPIConst.BUS_dSPCy      ,'(float) Bus state plane coordinate - Y'],
                'AREANO'  :[OlxAPIConst.BUS_nArea      ,'(int) Bus area number'],
                'NO'      :[OlxAPIConst.BUS_nNumber    ,'(int) Bus number'],
                'SLACK'   :[OlxAPIConst.BUS_nSlack     ,'(int) System slack bus flag: 1-yes; 0-no'],
                'SUBGRP'  :[OlxAPIConst.BUS_nSubGroup  ,'(int) Bus substation group'],
                'TAP'     :[OlxAPIConst.BUS_nTapBus    ,'(int) Tap bus flag: 0-no; 1-tap bus; 3-tap bus of 3-terminal line'],
                'VISIBLE' :[OlxAPIConst.BUS_nVisible   ,'(int) Bus hide ID flag: 1-visible; -2-hidden; 0-not yet placed'],
                'ZONENO'  :[OlxAPIConst.BUS_nZone      ,'(int) Bus zone number'],
                'LOCATION':[OlxAPIConst.BUS_sLocation  ,'(str) Bus location'],
                'NAME'    :[OlxAPIConst.BUS_sName      ,'(str) Bus name'],
                #
                'TERMINAL'        : [0, '[TERMINAL] List TERMINALs connected to BUS'],
                'BREAKER'         : [0, '[BREAKER] List Breakers connected to BUS'],
                'CCGEN'           : [0, '(CCGEN) Voltage controlled current sources connected to BUS'],
                'DCLINE2'         : [0, '[DCLINE2] List DC transmission lines connected to BUS'],
                'GEN'             : [0, '(GEN) Generator connected to BUS'],
                'GENUNIT'         : [0, '[GENUNIT] List Generator units connected to BUS'],
                'GENW3'           : [0, '(GENW3) Type-3 Wind Plants connected to BUS'],
                'GENW4'           : [0, '(GENW4) Converter-Interfaced Resource connected to BUS'],
                'LINE'            : [0, '[LINE] List AC transmission lines connected to BUS'],
                'LOAD'            : [0, '(LOAD) Load connected to BUS'],
                'LOADUNIT'        : [0, '[LOADUNIT] List Load units connected to BUS'],
                'RLYGROUP'        : [0, '[RLYGROUP] List Relays Group connected to BUS'],
                'SERIESRC'        : [0, '[SERIESRC] List Series reactors/capacitors connected to BUS'],
                'SHIFTER'         : [0, '[SHIFTER] List Shifters connected to BUS'],
                'SHUNT'           : [0, '(SHUNT) Shunt connected to BUS'],
                'SHUNTUNIT'       : [0, '[SHUNTUNIT] List Shunt units connected to BUS'],
                'SVD'             : [0, '(SVD) Switched Shunts connected to BUS'],
                'SWITCH'          : [0, '[SWITCH] List Switches connected to BUS'],
                'XFMR'            : [0, '[XFMR] List 2-winding transformers connected to BUS'],
                'XFMR3'           : [0, '[XFMR3] List 3-winding transformers connected to BUS']}
__OLXOBJ_PARA__['TERMINAL'] = {\
                'BUS'      : [0, '[BUS] List Buses that attached this TERMINAL'],
                'CID'      : [0, '(str) TERMINAL circuit ID'],
                'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that attached this TERMINAL'],
                'FLAG'     : [0, '(int) TERMINAL in-service flag: 1- active; 2- out-of-service'],
                'OPPOSITE' : [0, '[TERMINAL] List TERMINALs that opposite to this TERMINAL on the EQUIPMENT'],
                'REMOTE'   : [0, '[TERMINAL] List TERMINALs remote to this TERMINAL\
                                    \n\t                             All taps are ignored.\
                                    \n\t                             Close switches are included\
                                    \n\t                             Out of service branches are ignored'],
                'RLYGROUP' : [0, '[RLYGROUP] List RLYGROUPs that attached to this TERMINAL \
                                    \n\t                             RLYGROUP[0] : local RLYGROUP, =None if not found\
                                    \n\t                             RLYGROUP[i] : remote RLYGROUP \
                                    \n\t                             All taps are ignored.\
                                    \n\t                             Close switches are included\
                                    \n\t                             Out of service branches are ignored']}
__OLXOBJ_PARA__['GEN'] = {\
                'ILIMIT1'  :[OlxAPIConst.GE_dCurrLimit1 ,'(float) Generator current limit 1'],
                'ILIMIT2'  :[OlxAPIConst.GE_dCurrLimit2 ,'(float) Generator current limit 2'],
                'PGEN'     :[OlxAPIConst.GE_dPgen       ,'(float) Generator MW (load flow solution)'],
                'QGEN'     :[OlxAPIConst.GE_dQgen       ,'(float) Generator MVAR (load flow solution)'],
                'REFANGLE' :[OlxAPIConst.GE_dRefAngle   ,'(float) Generator reference angle'],
                'SCHEDP'   :[OlxAPIConst.GE_dScheduledP ,'(float) Generator scheduled P'],
                'SCHEDQ'   :[OlxAPIConst.GE_dScheduledQ ,'(float) Generator scheduled Q'],
                'SCHEDV'   :[OlxAPIConst.GE_dScheduledV ,'(float) Generator scheduled V'],
                'REFV'     :[OlxAPIConst.GE_dVSourcePU  ,'(float) Generator internal voltage source per unit magnitude'],
                'FLAG'     :[OlxAPIConst.GE_nActive     ,'(int) Generator in-service flag: 1- active; 2- out-of-service'],
                'REG'      :[OlxAPIConst.GE_nFixedPQ    ,'(int) Generator regulation flag: 1- PQ; 0- PV'],
                'BUS'      :[OlxAPIConst.GE_nBusHnd     ,'(BUS) Generators connected BUS'],
                'CNTBUS'   :[OlxAPIConst.GE_nCtrlBusHnd ,'(BUS) Generators controlled BUS'],
                'GENUNIT'  :[0                          ,'[GENUNIT] List of Generators unit']}
__OLXOBJ_PARA__['GENUNIT'] = {
                'MVARATE':[OlxAPIConst.GU_dMVArating  ,'(float) Generator unit rating MVA'],
                'PMAX'   :[OlxAPIConst.GU_dPmax       ,'(float) Generator unit max MW'],
                'PMIN'   :[OlxAPIConst.GU_dPmin       ,'(float) Generator unit min MW'],
                'QMAX'   :[OlxAPIConst.GU_dQmax       ,'(float) Generator unit max MVAR'],
                'QMIN'   :[OlxAPIConst.GU_dQmin       ,'(float) Generator unit min MVAR'],
                'RG'     :[OlxAPIConst.GU_dRz         ,'(float) Generator unit grounding resistance in Ohm (do not multiply by 3)'],
                'XG'     :[OlxAPIConst.GU_dXz         ,'(float) Generator unit grounding reactance in Ohm (do not multiply by 3)'],
                'SCHEDP' :[OlxAPIConst.GU_dSchedP     ,'(float) Generator unit scheduled P'],
                'SCHEDQ' :[OlxAPIConst.GU_dSchedQ     ,'(float) Generator unit scheduled Q'],
                'FLAG'   :[OlxAPIConst.GU_nOnline     ,'(int) Generator unit in-service flag 1- active; 2- out-of-service'],
                'CID'    :[OlxAPIConst.GU_sID         ,'(str) Generator unit ID'],
                'DATEOFF':[OlxAPIConst.GU_sOffDate    ,'(str) Generator unit out of service date'],
                'DATEON' :[OlxAPIConst.GU_sOnDate     ,'(str) Generator unit in service date'],
                'R'      :[OlxAPIConst.GU_vdR         ,'[float]*5 Generator unit resistances: [subtransient, synchronous, transient, negative sequence, zero sequence]'],
                'X'      :[OlxAPIConst.GU_vdX         ,'[float]*5 Generator unit reactances : [subtransient, synchronous, transient, negative sequence, zero sequence]'],
                'GEN'    :[OlxAPIConst.GU_nGenHnd     ,'(GEN) Generator unit Generator'],
                'BUS'    :[0                          ,'(BUS) Generator unit BUS']}
__OLXOBJ_PARA__['GENW3'] = {\
                'DATEON' :[OlxAPIConst.G3_sOnDate      ,'(str) Generator Type 3 in service date'],
                'DATEOFF':[OlxAPIConst.G3_sOffDate     ,'(str) Generator Type 3 out of service date'],
                'MVA'    :[OlxAPIConst.G3_dUnitRatedMVA,'(float) Generator Type 3 MVA unit rated'],
                'MWR'    :[OlxAPIConst.G3_dUnitRatedMW ,'(float) Generator Type 3 MW unit rated'],
                'MW'     :[OlxAPIConst.G3_dUnitMW      ,'(float) Generator Type 3 unit MW generation'],
                'IMAXR'  :[OlxAPIConst.G3_dImaxRsc     ,'(float) Generator Type 3 rotor side limit in pu'],
                'IMAXG'  :[OlxAPIConst.G3_dImaxGsc     ,'(float) Generator Type 3 Grid side limit in pu'],
                'VMAX'   :[OlxAPIConst.G3_dMaxV        ,'(float) Generator Type 3 maximum voltage limit in pu'],
                'VMIN'   :[OlxAPIConst.G3_dMinV        ,'(float) Generator Type 3 minimum voltage limit in pu'],
                'RR'     :[OlxAPIConst.G3_dRr          ,'(float) Generator Type 3 Rotor resistance in pu'],
                'LLR'    :[OlxAPIConst.G3_dLlr         ,'(float) Generator Type 3 Rotor leakage L in pu'],
                'RS'     :[OlxAPIConst.G3_dRs          ,'(float) Generator Type 3 Stator resistance in pu'],
                'LLS'    :[OlxAPIConst.G3_dLls         ,'(float) Generator Type 3 Stator leakage L in pu'],
                'LM'     :[OlxAPIConst.G3_dLm          ,'(float) Generator Type 3 Mutual L in pu'],
                'SLIP'   :[OlxAPIConst.G3_dSlipRated   ,'(float) Generator Type 3 Slip at rated kW'],
                'FLRZ'   :[OlxAPIConst.G3_dZfilter     ,'(float) Generator Type 3 Filter X in pu'],
                'FLAG'   :[OlxAPIConst.G3_nInService   ,'(int) Generator Type 3 unit in-service flag 1- active; 2- out-of-service'],
                'UNITS'  :[OlxAPIConst.G3_nNOUnits     ,'(int) Generator Type 3 number of units'],
                'CBAR'   :[OlxAPIConst.G3_nCrowbared   ,'(int) Generator Type 3 Crowbarred flag: 1-yes; 0-no'],
                'BUS'    :[OlxAPIConst.G3_nBusHnd      ,'(BUS) Generators Type 3 connected BUS']}
__OLXOBJ_PARA__['GENW4'] = {\
                'DATEON'    :[OlxAPIConst.G4_sOnDate        ,'(str) Generator Type 4 in service date'],
                'DATEOFF'   :[OlxAPIConst.G4_sOffDate       ,'(str) Generator Type 4 out of service date'],
                'MVA'       :[OlxAPIConst.G4_dUnitRatedMVA  ,'(float) Generator Type 4 Unit MVA rating'],
                'MW'        :[OlxAPIConst.G4_dUnitMW        ,'(float) Generator Type 4 Unit MW generation or consumption'],
                #'MVAR'      :[OlxAPIConst.G4_dUnitMVAR      ,'(float) Generator Type 4 Unit MVAR'],#ccc
                #'SCHEDP'    :[OlxAPIConst.G4_dUnitMW        ,'(float) Generator Type 4 scheduled P'],#ccc
                #'SCHEDQ'    :[OlxAPIConst.G4_dUnitMW        ,'(float) Generator Type 4 scheduled Q'],#ccc
                'VLOW'      :[OlxAPIConst.G4_dVlow          ,'(float) Generator Type 4 When +seq(pu)>'],
                'MAXI'      :[OlxAPIConst.G4_dMaxI          ,'(float) Generator Type 4 Max current'],
                'MAXILOW'   :[OlxAPIConst.G4_dMaxIlow       ,'(float) Generator Type 4 Max current reduce to'],
                'MVAR'      :[OlxAPIConst.G4_dUnitMVAR      ,'(float) Generator Type 4 Unit MVAR'],
                'SLOPE'     :[OlxAPIConst.G4_dSlopePos      ,'(float) Generator Type 4 Slope of +seq'],
                'SLOPENEG'  :[OlxAPIConst.G4_dSlopeNeg      ,'(float) Generator Type 4 Slope of -seq'],
                'VMAX'      :[OlxAPIConst.G4_dMaxV          ,'(float) Generator Type 4 maximum voltage limit'],
                'VMIN'      :[OlxAPIConst.G4_dMinV          ,'(float) Generator Type 4 minimum voltage limit'],
                'FLAG'      :[OlxAPIConst.G4_nInService     ,'(int) Generator Type 4 unit in-service flag 1- active; 2- out-of-service'],
                'UNITS'     :[OlxAPIConst.G4_nNOUnits       ,'(int) Generator Type 4 number of units'],
                'CTRLMETHOD':[OlxAPIConst.G4_nControlMethod ,'(int) Generator Type 4 control method'],
                'BUS'       :[OlxAPIConst.G4_nBusHnd        ,'(BUS) Generators Type 4 connected BUS']}
__OLXOBJ_PARA__['CCGEN'] = {\
                'MVARATE'   :[OlxAPIConst.CC_dMVArating    ,'(float) Voltage controlled current source MVA rating'],
                'VMAXMUL'   :[OlxAPIConst.CC_dVmax         ,'(float) Voltage controlled current source maximum voltage limit in pu'],
                'VMIN'      :[OlxAPIConst.CC_dVmin         ,'(float) Voltage controlled current source minimum voltage limit in pu'],
                'BLOCKPHASE':[OlxAPIConst.CC_nBlockOnPhaseV,'(int) Voltage controlled current source number block on phase'],
                'FLAG'      :[OlxAPIConst.CC_nInService    ,'(int) Voltage controlled current source in-service flag: 1-true; 2-false'],
                'VLOC'      :[OlxAPIConst.CC_nVloc         ,'(int) Voltage controlled current source voltage measurement location 0-Device terminal; 1-Network side of transformer'],
                'DATEOFF'   :[OlxAPIConst.CC_sOffDate      ,'(str) Voltage controlled current source out of service date'],
                'DATEON'    :[OlxAPIConst.CC_sOnDate       ,'(str) Voltage controlled current source in service date'],
                'A'         :[OlxAPIConst.CC_vdAng         ,'[float]*10 Voltage controlled current source angle'],
                'I'         :[OlxAPIConst.CC_vdI           ,'[float]*10 Voltage controlled current source current'],
                'V'         :[OlxAPIConst.CC_vdV           ,'[float]*10 Voltage controlled current source voltage'],
                'BUS'       :[OlxAPIConst.CC_nBusHnd       ,'(BUS) Voltage controlled current source bus']}
__OLXOBJ_PARA__['XFMR'] = {\
                'B'         :[OlxAPIConst.XR_dB            ,'(float) 2-winding transformer B'],
                'B0'        :[OlxAPIConst.XR_dB0           ,'(float) 2-winding transformer Bo'],
                'B1'        :[OlxAPIConst.XR_dB1           ,'(float) 2-winding transformer B1'],
                'B10'       :[OlxAPIConst.XR_dB10          ,'(float) 2-winding transformer B10'],
                'B2'        :[OlxAPIConst.XR_dB2           ,'(float) 2-winding transformer B2'],
                'B20'       :[OlxAPIConst.XR_dB20          ,'(float) 2-winding transformer B20'],
                'BASEMVA'   :[OlxAPIConst.XR_dBaseMVA      ,'(float) 2-winding transformer base MVA for per-unit quantities'],
                'G1'        :[OlxAPIConst.XR_dG1           ,'(float) 2-winding transformer G1'],
                'G10'       :[OlxAPIConst.XR_dG10          ,'(float) 2-winding transformer G10'],
                'G2'        :[OlxAPIConst.XR_dG2           ,'(float) 2-winding transformer G2'],
                'G20'       :[OlxAPIConst.XR_dG20          ,'(float) 2-winding transformer G20'],
                'LTCCENTER' :[OlxAPIConst.XR_dLTCCenterTap ,'(float) 2-winding transformer LTC center tap'],
                'LTCSTEP'   :[OlxAPIConst.XR_dLTCstep      ,'(float) 2-winding transformer LTC step size'],
                'MVA1'      :[OlxAPIConst.XR_dMVA1         ,'(float) 2-winding transformer MVA1'],
                'MVA2'      :[OlxAPIConst.XR_dMVA2         ,'(float) 2-winding transformer MVA2'],
                'MVA3'      :[OlxAPIConst.XR_dMVA3         ,'(float) 2-winding transformer MVA3'],
                'MAXTAP'    :[OlxAPIConst.XR_dMaxTap       ,'(float) 2-winding transformer LTC max tap'],
                'MAXVW'     :[OlxAPIConst.XR_dMaxVW        ,'(float) 2-winding transformer LTC min controlled quantity limit'],
                'MINTAP'    :[OlxAPIConst.XR_dMinTap       ,'(float) 2-winding transformer LTC min tap'],
                'MINVW'     :[OlxAPIConst.XR_dMinVW        ,'(float) 2-winding transformer LTC max controlled quantity limit'],
                'PRITAP'    :[OlxAPIConst.XR_dPriTap       ,'(float) 2-winding transformer primary tap'],
                'R'         :[OlxAPIConst.XR_dR            ,'(float) 2-winding transformer R'],
                'R0'        :[OlxAPIConst.XR_dR0           ,'(float) 2-winding transformer Ro'],
                'RG1'       :[OlxAPIConst.XR_dRG1          ,'(float) 2-winding transformer Rg1'],
                'RG2'       :[OlxAPIConst.XR_dRG2          ,'(float) 2-winding transformer Rg2'],
                'RGN'       :[OlxAPIConst.XR_dRGN          ,'(float) 2-winding transformer Rgn'],
                'SECTAP'    :[OlxAPIConst.XR_dSecTap       ,'(float) 2-winding transformer secondary tap'],
                'X'         :[OlxAPIConst.XR_dX            ,'(float) 2-winding transformer X'],
                'X0'        :[OlxAPIConst.XR_dX0           ,'(float) 2-winding transformer Xo'],
                'XG1'       :[OlxAPIConst.XR_dXG1          ,'(float) 2-winding transformer Xg1'],
                'XG2'       :[OlxAPIConst.XR_dXG2          ,'(float) 2-winding transformer Xg2'],
                'XGN'       :[OlxAPIConst.XR_dXGN          ,'(float) 2-winding transformer Xgn'],
                'AUTOX'     :[OlxAPIConst.XR_nAuto         ,'(int) 2-winding transformer auto transformer flag:1-true;0-false'],
                'FLAG'      :[OlxAPIConst.XR_nInService    ,'(int) 2-winding transformer in-service flag: 1- active; 2- out-of-service'],
                'GANGED'    :[OlxAPIConst.XR_nLTCGanged    ,'(int) 2-winding transformer LTC tag ganged flag: 0-False; 1-True'],
                'PRIORITY'  :[OlxAPIConst.XR_nLTCPriority  ,'(int) 2-winding transformer LTC adjustment priority 0-Normal, 1-Medieum, 2-High'],
                'LTCSIDE'   :[OlxAPIConst.XR_nLTCside      ,'(int) 2-winding transformer LTC side: 1; 2; 0'],
                'LTCTYPE'   :[OlxAPIConst.XR_nLTCtype      ,'(int) 2-winding transformer LTC type: 1- control voltage; 2- control MVAR'],
                'LTCCTRL'   :[OlxAPIConst.XR_nLTCCtrlBusHnd,'(str) 2-winding transformer ID string of the bus whose voltage magnitude is to be regulated by the LTC'],
                'CONFIGP'   :[OlxAPIConst.XR_sCfgP         ,'(str) 2-winding transformer primary winding config'],
                'CONFIGS'   :[OlxAPIConst.XR_sCfgS         ,'(str) 2-winding transformer secondary winding config'],
                'CONFIGST'  :[OlxAPIConst.XR_sCfgST        ,'(str) 2-winding transformer secondary winding config in test'],
                'CID'       :[OlxAPIConst.XR_sID           ,'(str) 2-winding transformer circuit ID'],
                'NAME'      :[OlxAPIConst.XR_sName         ,'(str) 2-winding transformer name'],
                'DATEOFF'   :[OlxAPIConst.XR_sOffDate      ,'(str) 2-winding transformer out of service date'],
                'DATEON'    :[OlxAPIConst.XR_sOnDate       ,'(str) 2-winding transformer in service date'],
                'METEREDEND':[OlxAPIConst.XR_nMeteredEnd   ,'(int) 2-winding transformer metered bus 1-at Bus1; 2 at Bus2; 0 XFMR in a single area'],
                'BUS1'      :[OlxAPIConst.XR_nBus1Hnd      ,'(BUS) 2-winding transformer bus 1'],
                'BUS2'      :[OlxAPIConst.XR_nBus2Hnd      ,'(BUS) 2-winding transformer bus 2'],
                'RLYGROUP1' :[OlxAPIConst.XR_nRlyGr1Hnd    ,'(RLYGROUP) 2-winding transformer side 1 Relay Group'],
                'RLYGROUP2' :[OlxAPIConst.XR_nRlyGr2Hnd    ,'(RLYGROUP) 2-winding transformer side 2 Relay Group']}
__OLXOBJ_PARA__['XFMR3'] = {\
                'B'        :[OlxAPIConst.X3_dB            ,'(float) 3-winding transformer B'],
                'B0'       :[OlxAPIConst.X3_dB0           ,'(float) 3-winding transformer B0'],
                'BASEMVA'  :[OlxAPIConst.X3_dBaseMVA      ,'(float) 3-winding transformer base MVA for per-unit quantities'],
                'LTCCENTER':[OlxAPIConst.X3_dLTCCenterTap ,'(float) 3-winding transformer LTC center tap'],
                'LTCSTEP'  :[OlxAPIConst.X3_dLTCstep      ,'(float) 3-winding transformer LTC step size'],
                'LTCSIDE'  :[OlxAPIConst.X3_nInService    ,'(int) 3-winding transformer LTC side: 1; 2; 0'], #ccc
                'LTCTYPE'  :[OlxAPIConst.X3_nInService    ,'(int) 3-winding transformer LTC type: 1- control voltage; 2- control MVAR'],#ccc
                'LTCCTRL'  :[OlxAPIConst.X3_nLTCCtrlBusHnd,'(str) 3-winding transformer ID string of the bus whose voltage magnitude is to be regulated by the LTC'],
                'MVA1'     :[OlxAPIConst.X3_dMVA1         ,'(float) 3-winding transformer MVA1'],
                'MVA2'     :[OlxAPIConst.X3_dMVA2         ,'(float) 3-winding transformer MVA2'],
                'MVA3'     :[OlxAPIConst.X3_dMVA3         ,'(float) 3-winding transformer MVA3'],
                'MAXTAP'   :[OlxAPIConst.X3_dMaxTap       ,'(float) 3-winding transformer LTC max tap'],
                'MAXVW'    :[OlxAPIConst.X3_dMaxVW        ,'(float) 3-winding transformer LTC min controlled quantity limit'],
                'MINTAP'   :[OlxAPIConst.X3_dMinTap       ,'(float) 3-winding transformer LTC min tap'],
                'MINVW'    :[OlxAPIConst.X3_dMinVW        ,'(float) 3-winding transformer LTC controlled quantity limit'],
                'PRITAP'   :[OlxAPIConst.X3_dPriTap       ,'(float) 3-winding transformer primary tap'],
                'RPS0'     :[OlxAPIConst.X3_dR0ps         ,'(float) 3-winding transformer R0ps'],
                'RPT0'     :[OlxAPIConst.X3_dR0pt         ,'(float) 3-winding transformer R0pt'],
                'RST0'     :[OlxAPIConst.X3_dR0st         ,'(float) 3-winding transformer R0st'],
                'RG1'      :[OlxAPIConst.X3_dRG1          ,'(float) 3-winding transformer Rg1'],
                'RG2'      :[OlxAPIConst.X3_dRG2          ,'(float) 3-winding transformer Rg2'],
                'RG3'      :[OlxAPIConst.X3_dRG3          ,'(float) 3-winding transformer Rg3'],
                'RGN'      :[OlxAPIConst.X3_dRGN          ,'(float) 3-winding transformer Rgn'],
                'RPS'      :[OlxAPIConst.X3_dRps          ,'(float) 3-winding transformer Rps'],
                'RPT'      :[OlxAPIConst.X3_dRpt          ,'(float) 3-winding transformer Rpt'],
                'RST'      :[OlxAPIConst.X3_dRst          ,'(float) 3-winding transformer Rst'],
                'SECTAP'   :[OlxAPIConst.X3_dSecTap       ,'(float) 3-winding transformer secondary tap'],
                'TERTAP'   :[OlxAPIConst.X3_dTerTap       ,'(float) 3-winding transformer tertiary tap'],
                'XPS0'     :[OlxAPIConst.X3_dX0ps         ,'(float) 3-winding transformer X0ps'],
                'XPT0'     :[OlxAPIConst.X3_dX0pt         ,'(float) 3-winding transformer X0pt'],
                'XST0'     :[OlxAPIConst.X3_dX0st         ,'(float) 3-winding transformer X0st'],
                'XG1'      :[OlxAPIConst.X3_dXG1          ,'(float) 3-winding transformer Xg1'],
                'XG2'      :[OlxAPIConst.X3_dXG2          ,'(float) 3-winding transformer Xg2'],
                'XG3'      :[OlxAPIConst.X3_dXG3          ,'(float) 3-winding transformer Xg3'],
                'XGN'      :[OlxAPIConst.X3_dXGN          ,'(float) 3-winding transformer Xgn'],
                'XPS'      :[OlxAPIConst.X3_dXps          ,'(float) 3-winding transformer Xps'],
                'XPT'      :[OlxAPIConst.X3_dXpt          ,'(float) 3-winding transformer Xpt'],
                'XST'      :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer Xst'],
                'RPM0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer RPM0'], #ccc
                'XPM0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer XPM0'], #ccc
                'RSM0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer RSM0'], #ccc
                'XSM0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer XSM0'], #ccc
                'RMG0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer RMG0'], #ccc
                'XMG0'     :[OlxAPIConst.X3_dXst          ,'(float) 3-winding transformer XMG0'], #ccc
                'AUTOX'    :[OlxAPIConst.X3_nAuto         ,'(int) 3-winding transformer auto transformer flag:1-true;0-false'],
                'FLAG'     :[OlxAPIConst.X3_nInService    ,'(int) 3-winding transformer in-service flag: 1- active; 2- out-of-service'],
                'Z0METHOD' :[OlxAPIConst.X3_nInService    ,'(int) 3-winding transformer Z0 method 1-Short circuit impedance; 2-Classical T model impedance'], #ccc
                'FICTBUSNO':[OlxAPIConst.X3_nFictBusNo    ,'(int) 3-winding transformer Fiction bus Number'],
                'GANGED'   :[OlxAPIConst.X3_nLTCGanged    ,'(int) 3-winding transformer LTC tag ganged flag: 0-False; 1-True'],
                'PRIORITY' :[OlxAPIConst.X3_nLTCPriority  ,'(int) 3-winding transformer LTC adjustment priority 0-Normal, 1-Medieum, 2-High'],
                'CONFIGP'  :[OlxAPIConst.X3_sCfgP         ,'(str) 3-winding transformer primary winding'],
                'CONFIGS'  :[OlxAPIConst.X3_sCfgS         ,'(str) 3-winding transformer secondary winding'],
                'CONFIGST' :[OlxAPIConst.X3_sCfgST        ,'(str) 3-winding transformer secondary winding in test'],
                'CONFIGT'  :[OlxAPIConst.X3_sCfgT         ,'(str) 3-winding transformer tertiary winding'],
                'CONFIGTT' :[OlxAPIConst.X3_sCfgTT        ,'(str) 3-winding transformer tertiary winding in test'],
                'CID'      :[OlxAPIConst.X3_sID           ,'(str) 3-winding transformer circuit ID'],
                'NAME'     :[OlxAPIConst.X3_sName         ,'(str) 3-winding transformer name'],
                'DATEOFF'  :[OlxAPIConst.X3_sOffDate      ,'(str) 3-winding transformer out of service date'],
                'DATEON'   :[OlxAPIConst.X3_sOnDate       ,'(str) 3-winding transformer in service date'],
                'BUS1'     :[OlxAPIConst.X3_nBus1Hnd      ,'(BUS) 3-winding transformer bus 1'],
                'BUS2'     :[OlxAPIConst.X3_nBus2Hnd      ,'(BUS) 3-winding transformer bus 2'],
                'BUS3'     :[OlxAPIConst.X3_nBus3Hnd      ,'(BUS) 3-winding transformer bus 3'],
                'RLYGROUP1':[OlxAPIConst.X3_nRlyGr1Hnd    ,'(RLYGROUP) 3-winding transformer Relay Group 1'],
                'RLYGROUP2':[OlxAPIConst.X3_nRlyGr2Hnd    ,'(RLYGROUP) 3-winding transformer Relay Group 2'],
                'RLYGROUP3':[OlxAPIConst.X3_nRlyGr3Hnd    ,'(RLYGROUP) 3-winding transformer Relay Group 3']}
__OLXOBJ_PARA__['SHIFTER'] = {\
                'SHIFTANGLE':[OlxAPIConst.PS_dAngle      ,'(float) Phase shifter shift angle'],
                'ANGMAX'    :[OlxAPIConst.PS_dAngleMax   ,'(float) Phase shifter shift angle max'],
                'ANGMIN'    :[OlxAPIConst.PS_dAngleMin   ,'(float) Phase shifter shift angle min'],
                'BN1'       :[OlxAPIConst.PS_dBN1        ,'(float) Phase shifter Bn1'],
                'BN2'       :[OlxAPIConst.PS_dBN2        ,'(float) Phase shifter Bn2'],
                'BP1'       :[OlxAPIConst.PS_dBP1        ,'(float) Phase shifter Bp1'],
                'BP2'       :[OlxAPIConst.PS_dBP2        ,'(float) Phase shifter Bp2'],
                'BZ1'       :[OlxAPIConst.PS_dBZ1        ,'(float) Phase shifter Bz1'],
                'BZ2'       :[OlxAPIConst.PS_dBZ2        ,'(float) Phase shifter Bz2'],
                'BASEMVA'   :[OlxAPIConst.PS_dBaseMVA    ,'(float) Phase shifter BaseMVA'],
                'GN1'       :[OlxAPIConst.PS_dGN1        ,'(float) Phase shifter Gn1'],
                'GN2'       :[OlxAPIConst.PS_dGN2        ,'(float) Phase shifter Gn2'],
                'GP1'       :[OlxAPIConst.PS_dGP1        ,'(float) Phase shifter Gp1'],
                'GP2'       :[OlxAPIConst.PS_dGP2        ,'(float) Phase shifter Gp2'],
                'GZ1'       :[OlxAPIConst.PS_dGZ1        ,'(float) Phase shifter Gz1'],
                'GZ2'       :[OlxAPIConst.PS_dGZ2        ,'(float) Phase shifter Gz2'],
                'MVA1'      :[OlxAPIConst.PS_dMVA1       ,'(float) Phase shifter MVA1'],
                'MVA2'      :[OlxAPIConst.PS_dMVA2       ,'(float) Phase shifter MVA2'],
                'MVA3'      :[OlxAPIConst.PS_dMVA3       ,'(float) Phase shifter MVA3'],
                'MWMAX'     :[OlxAPIConst.PS_dMWmax      ,'(float) Phase shifter MW max'],
                'MWMIN'     :[OlxAPIConst.PS_dMWmin      ,'(float) Phase shifter MW min'],
                'RN'        :[OlxAPIConst.PS_dRN         ,'(float) Phase shifter Rn'],
                'RP'        :[OlxAPIConst.PS_dRP         ,'(float) Phase shifter Rp'],
                'RZ'        :[OlxAPIConst.PS_dRZ         ,'(float) Phase shifter Rz'],
                'XN'        :[OlxAPIConst.PS_dXN         ,'(float) Phase shifter Xn'],
                'XP'        :[OlxAPIConst.PS_dXP         ,'(float) Phase shifter Xp'],
                'XZ'        :[OlxAPIConst.PS_dXZ         ,'(float) Phase shifter Xz'],
                'CNTL'      :[OlxAPIConst.PS_nControlMode,'(int) Phase shifter control mode 0-Fixed, 1-automatically control real power flow '],
                'FLAG'      :[OlxAPIConst.PS_nInService  ,'(int) Phase shifter in-service flag: 1- active; 2- out-of-service'],
                'ZCORRECTNO':[OlxAPIConst.PS_nInService  ,'(int) Phase shifter correct table number'], #ccc
                'CID'       :[OlxAPIConst.PS_sID         ,'(str) Phase shifter circuit ID'],
                'NAME'      :[OlxAPIConst.PS_sName       ,'(str) Phase shifter name'],
                'DATEOFF'   :[OlxAPIConst.PS_sOffDate    ,'(str) Phase shifter out of service date'],
                'DATEON'    :[OlxAPIConst.PS_sOnDate     ,'(str) Phase shifter in service date'],
                'BUS1'      :[OlxAPIConst.PS_nBus1Hnd    ,'(BUS) Phase shifter bus 1'],
                'BUS2'      :[OlxAPIConst.PS_nBus2Hnd    ,'(BUS) Phase shifter bus 2'],
                'RLYGROUP1' :[OlxAPIConst.PS_nRlyGr1Hnd  ,'(RLYGROUP) Phase shifter Relay Group 1'],
                'RLYGROUP2' :[OlxAPIConst.PS_nRlyGr2Hnd  ,'(RLYGROUP) Phase shifter Relay Group 2']}
__OLXOBJ_PARA__['LINE'] = {\
                'B1'       :[OlxAPIConst.LN_dB1         ,'(float) Line B1 , in pu'],
                'B10'      :[OlxAPIConst.LN_dB10        ,'(float) Line B10, in pu'],
                'B2'       :[OlxAPIConst.LN_dB2         ,'(float) Line B2, in pu'],
                'B20'      :[OlxAPIConst.LN_dB20        ,'(float) Line B20, in pu'],
                'G1'       :[OlxAPIConst.LN_dG1         ,'(float) Line G1, in pu'],
                'G10'      :[OlxAPIConst.LN_dG10        ,'(float) Line G10, in pu'],
                'G2'       :[OlxAPIConst.LN_dG2         ,'(float) Line G2, in pu'],
                'G20'      :[OlxAPIConst.LN_dG20        ,'(float) Line G20, in pu'],
                'LN'       :[OlxAPIConst.LN_dLength     ,'(float) Line length'],
                'R'        :[OlxAPIConst.LN_dR          ,'(float) Line R, in pu'],
                'R0'       :[OlxAPIConst.LN_dR0         ,'(float) Line Ro, in pu'],
                'X'        :[OlxAPIConst.LN_dX          ,'(float) Line X, in pu'],
                'X0'       :[OlxAPIConst.LN_dX0         ,'(float) Line Xo, in pu'],
                'FLAG'     :[OlxAPIConst.LN_nInService  ,'(int) Line in-service flag: 1- active; 2- out-of-service'],
                'METEREDEND':[OlxAPIConst.LN_nMeteredEnd ,'(int) Line meteted flag: 1- at Bus1; 2-at Bus2; 0-line is in a single area;'], #ccc
                'I2T'      :[OlxAPIConst.LN_dI2T        ,'(float) I^2T rating of line in ampere^2 Sec.'], #ccc
                'CID'      :[OlxAPIConst.LN_sID         ,'(str) Line circuit ID'],
                'UNIT'     :[OlxAPIConst.LN_sLengthUnit ,'(str) Line length unit in [ft,kt,mi,m,km]'],
                'NAME'     :[OlxAPIConst.LN_sName       ,'(str) Line name'],
                'DATEOFF'  :[OlxAPIConst.LN_sOffDate    ,'(str) Line out of service date'],
                'DATEON'   :[OlxAPIConst.LN_sOnDate     ,'(str) Line in service date'],
                'TYPE'     :[OlxAPIConst.LN_sType       ,"(str) Line table type in ['','Delta Dove','Horz. Dove','Vert. Dove]"],
                'RATG'     :[OlxAPIConst.LN_vdRating    ,'[float]*4 Line ratings'],
                'BUS1'     :[OlxAPIConst.LN_nBus1Hnd    ,'(BUS) Line bus 1'],
                'BUS2'     :[OlxAPIConst.LN_nBus2Hnd    ,'(BUS) Line bus 2'],
                'MULINE'   :[OlxAPIConst.LN_nMuPairHnd  ,'(MULINE) Line mutual pair'],
                'RLYGROUP1':[OlxAPIConst.LN_nRlyGr1Hnd  ,'(RLYGROUP) Line Relay Group 1'],
                'RLYGROUP2':[OlxAPIConst.LN_nRlyGr2Hnd  ,'(RLYGROUP) Line Relay Group 2']}
__OLXOBJ_PARA__['DCLINE2'] = {\
                'NAME'   :[OlxAPIConst.DC_sName         ,'(string) DC Line name'],
                'CID'    :[OlxAPIConst.DC_sID           ,'(string) DC Line circuit ID'],
                'DATEON' :[OlxAPIConst.DC_sOnDate       ,'(string) DC Line in service date'],
                'DATEOFF':[OlxAPIConst.DC_sOffDate      ,'(string) DC Line out of service date'],
                'TARGET' :[OlxAPIConst.DC_dControlTarget,'(float) DC Line target MW'],
                'MARGIN' :[OlxAPIConst.DC_dControlMargin,'(float) DC Line Margin'],
                'VSCHED' :[OlxAPIConst.DC_dVdcSched     ,'(float) DC Line Schedule DC volts in pu'],
                'LINER'  :[OlxAPIConst.DC_dDClineR      ,'(float) DC Line R(Ohm)'],
                'FLAG'   :[OlxAPIConst.DC_nInService    ,'(int) DC Line in-service flag: 1- active; 2- out-of-service'],
                'METEREDEND':[OlxAPIConst.DC_nMeteredEnd,'(int) DC Line meteted flag: 1- at Bus1; 2-at Bus2; 0-line is in a single area;'], #ccc
                'MODE'   :[OlxAPIConst.DC_nInService    ,'(int) DC Line control mode'], #ccc
                'ANGMAX' :[OlxAPIConst.DC_vdAngleMax    ,'[float]*2 DC Line Angle max'],
                'ANGMIN' :[OlxAPIConst.DC_vdAngleMin    ,'[float]*2 DC Line Angle min'],
                'TAPMAX' :[OlxAPIConst.DC_vdTapMax      ,'[float]*2 DC Line Tap max '],
                'TAPMIN' :[OlxAPIConst.DC_vdTapMax      ,'[float]*2 DC Line Tap min'],
                'TAPSTEP':[OlxAPIConst.DC_vdTapStepSize ,'[float]*2 DC Line Tap step size'],
                'TAP'    :[OlxAPIConst.DC_vdTapRatio    ,'[float]*2 DC Line Tap'],
                'MVA'    :[OlxAPIConst.DC_vdMVA         ,'[float]*2 DC Line MVA rating'],
                'NOMKV'  :[OlxAPIConst.DC_vdNomKV       ,'[float]*2 DC Line nomibal KV on dc side'],
                'R'      :[OlxAPIConst.DC_vdXfmrR       ,'[float]*2 DC Line XFMR R in pu'],
                'X'      :[OlxAPIConst.DC_vdXfmrX       ,'[float]*2 DC Line XFMR X in pu'],
                'BRIDGE' :[OlxAPIConst.DC_vnBridges     ,'[int]*2 DC Line No. of bridges'],
                'BUS1'   :[OlxAPIConst.DC_nBus1Hnd      ,'(BUS) DC Line bus 1'],
                'BUS2'   :[OlxAPIConst.DC_nBus2Hnd      ,'(BUS) DC Line bus 2']}
__OLXOBJ_PARA__['SERIESRC'] = {\
                'IPR'      :[OlxAPIConst.SC_dIpr      ,'(float) Series capacitor/reactor protective level current'],
                'SCOMP'    :[OlxAPIConst.SC_nSComp    ,'(int) Series capacitor/reactor bypassed flag 1- no bypassed; 2-yes bypassed '],
                'R'        :[OlxAPIConst.SC_dR        ,'(float) Series capacitor/reactor R'],
                'X'        :[OlxAPIConst.SC_dX        ,'(float) Series capacitor/reactor X'],
                'FLAG'     :[OlxAPIConst.SC_nInService,'(int) Series capacitor/reactor in-service flag: 1- active; 2- out-of-service; 3- bypassed'],
                'CID'      :[OlxAPIConst.SC_sID       ,'(str) Series capacitor/reactor circuit ID'],
                'NAME'     :[OlxAPIConst.SC_sName     ,'(str) Series capacitor/reactor name'],
                'DATEOFF'  :[OlxAPIConst.SC_sOffDate  ,'(str) Series capacitor/reactor out of service date'],
                'DATEON'   :[OlxAPIConst.SC_sOnDate   ,'(str) Series capacitor/reactor in service date'],
                'BUS1'     :[OlxAPIConst.SC_nBus1Hnd  ,'(BUS) Series capacitor/reactor bus 1'],
                'BUS2'     :[OlxAPIConst.SC_nBus2Hnd  ,'(BUS) Series capacitor/reactor bus 2'],
                'RLYGROUP1':[OlxAPIConst.SC_nRlyGr1Hnd,'(RLYGROUP) Series capacitor/reactor Relay Group 1'],
                'RLYGROUP2':[OlxAPIConst.SC_nRlyGr2Hnd,'(RLYGROUP) Series capacitor/reactor Relay Group 2']}
__OLXOBJ_PARA__['SWITCH'] = {\
                'RATING'   :[OlxAPIConst.SW_dRating   ,'(float) Switch current rating'],
                'DEFAULT'  :[OlxAPIConst.SW_nDefault  ,'(int) Switch default position flag: 1- normaly open; 2- normaly close; 0-Not defined'],
                'FLAG'     :[OlxAPIConst.SW_nInService,'(int) Switch in-service flag: 1- active; 2- out-of-service'],
                'STAT'     :[OlxAPIConst.SW_nStatus   ,'(int) Switch position flag: 7- close; 0- open'],
                'CID'      :[OlxAPIConst.SW_sID       ,'(str) Switch Id'],
                'NAME'     :[OlxAPIConst.SW_sName     ,'(str) Switch name'],
                'DATEOFF'  :[OlxAPIConst.SW_sOffDate  ,'(str) Switch out of service date'],
                'DATEON'   :[OlxAPIConst.SW_sOnDate   ,'(str) Switch in service date'],
                'BUS1'     :[OlxAPIConst.SW_nBus1Hnd  ,'(BUS) Switch bus 1'],
                'BUS2'     :[OlxAPIConst.SW_nBus2Hnd  ,'(BUS) Switch bus 2'],
                'RLYGROUP1':[OlxAPIConst.SW_nRlyGrHnd1,'(RLYGROUP) Switch Relay Group 1'],
                'RLYGROUP2':[OlxAPIConst.SW_nRlyGrHnd2,'(RLYGROUP) Switch Relay Group 2']}
__OLXOBJ_PARA__['LOAD'] = {\
                'P'         :[OlxAPIConst.LD_dPload     ,'(float) Total load MW (load flow solution)'],
                'Q'         :[OlxAPIConst.LD_dQload     ,'(float) Total load MVAR (load flow solution)'],
                'FLAG'      :[OlxAPIConst.LD_nActive    ,'(int) Load in-service flag: 1- active; 2- out-of-service'],
                'UNGROUNDED':[OlxAPIConst.LD_nUnGrounded,'(int) Load UnGrounded 1-UnGrounded 0-Grounded'], # #
                'BUS'       :[OlxAPIConst.LD_nBusHnd    ,'(BUS) Load bus'],
                'LOADUNIT'  :[0                         ,'[LOADUNIT] List LOAD. units in LOAD']}
__OLXOBJ_PARA__['LOADUNIT'] = {\
                'P'      :[OlxAPIConst.LU_dPload  ,'(float) Load unit MW (load flow solution)'],
                'Q'      :[OlxAPIConst.LU_dQload  ,'(float) Load unit MVAR (load flow solution)'],
                'FLAG'   :[OlxAPIConst.LU_nOnline ,'(int) Load unit in-service flag: 1-active; 2- out-of-service'],
                'CID'    :[OlxAPIConst.LU_sID     ,'(str) Load unit circuit ID'],
                'DATEOFF':[OlxAPIConst.LU_sOffDate,'(str) Load unit out of service date'],
                'DATEON' :[OlxAPIConst.LU_sOnDate ,'(str) Load unit in service date'],
                'MVAR'   :[OlxAPIConst.LU_vdMVAR  ,'[float]*3 Load unit MVARs: [const. P, const. I, const. Z]'],
                'MW'     :[OlxAPIConst.LU_vdMW    ,'[float]*3 Load unit MWs: [const. P, const. I, const. Z]'],
                'LOAD'   :[OlxAPIConst.LU_nLoadHnd,'(LOAD) Load unit load'],
                'BUS'    :[0                      ,'(BUS) Load unit BUS']}
__OLXOBJ_PARA__['SHUNT'] = {\
                'FLAG'   :[OlxAPIConst.SH_nActive ,'(int) Shunt in-service flag: 1- active; 2- out-of-service'],
                'BUS'    :[OlxAPIConst.SH_nBusHnd ,'(BUS) Shunt Bus']}
__OLXOBJ_PARA__['SHUNTUNIT'] = {\
                'B'      :[OlxAPIConst.SU_dB       ,'(float) Shunt unit succeptance (positive sequence)'],
                'B0'     :[OlxAPIConst.SU_dB0      ,'(float) Shunt unit succeptance (zero sequence)'],
                'G'      :[OlxAPIConst.SU_dG       ,'(float) Shunt unit conductance (positive sequence)'],
                'G0'     :[OlxAPIConst.SU_dG0      ,'(float) Shunt unit conductance (zero sequence)'],
                'TX3'    :[OlxAPIConst.SU_n3WX     ,'(int) Shunt unit 3-winding transformer flag: 1-true;0-false'],
                'FLAG'   :[OlxAPIConst.SU_nOnline  ,'(int) Shunt unit in-service flag: 1- active; 2- out-of-service'],
                'CID'    :[OlxAPIConst.SU_sID      ,'(str) Shunt unit ID'],
                'DATEOFF':[OlxAPIConst.SU_sOffDate ,'(str) Shunt unit out of service date'],
                'DATEON' :[OlxAPIConst.SU_sOnDate  ,'(str) Shunt unit in service date'],
                'SHUNT'  :[OlxAPIConst.SU_nShuntHnd,'(SHUNT) Shunt unit shunt']}
__OLXOBJ_PARA__['SVD'] = {\
                'B_USE'  :[OlxAPIConst.SV_dB         ,'(float) SVD admitance in use'],
                'VMAX'   :[OlxAPIConst.SV_dVmax      ,'(float) SVD max V'],
                'VMIN'   :[OlxAPIConst.SV_dVmin      ,'(float) SVD min V'],
                'FLAG'   :[OlxAPIConst.SV_nActive    ,'(int) SVD in-service flag: 1- active; 2- out-of-service'],
                'MODE'   :[OlxAPIConst.SV_nCtrlMode  ,'(int) SVD control mode: 0-Fixed; 1-Discrete; 2-Continous'],
                'DATEOFF':[OlxAPIConst.SV_sOffDate   ,'(str) SVD out of service date'],
                'DATEON' :[OlxAPIConst.SV_sOnDate    ,'(str) SVD in service date'],
                'B0'    :[OlxAPIConst.SV_vdB0inc    ,'[float]*8 SVD increment zero admitance'],
                'B'     :[OlxAPIConst.SV_vdBinc     ,'[float]*8 SVD increment admitance'],
                'STEP'   :[OlxAPIConst.SV_vnNoStep   ,'[int]*8 SVD number of step'],
                'BUS'    :[OlxAPIConst.SV_nBusHnd    ,'(BUS) SVD Bus'],
                'CNTBUS' :[OlxAPIConst.SV_nCtrlBusHnd,'(BUS) SVD controled Bus']}
__OLXOBJ_PARA__['BREAKER'] = {\
                'CPT1'      :[OlxAPIConst.BK_dCPT1        ,'(float) Breaker contact parting time for group 1 (cycles)'],
                'CPT2'      :[OlxAPIConst.BK_dCPT2        ,'(float) Breaker contact parting time for group 2 (cycles)'],
                'INTRTIME'  :[OlxAPIConst.BK_dCycles      ,'(float) Breaker interrupting time (cycles)'],
                'K'         :[OlxAPIConst.BK_dK           ,'(float) Breaker kV range factor'],
                'NACD'      :[OlxAPIConst.BK_dNACD        ,'(float) Breaker no-ac-decay ratio'],
                'OPKV'      :[OlxAPIConst.BK_dOperatingKV ,'(float) Breaker operating kV'],
                'RATEDKV'   :[OlxAPIConst.BK_dRatedKV     ,'(float) Breaker max design kV'],
                'INTRATING' :[OlxAPIConst.BK_dRating1     ,'(float) Breaker interrupting rating'],
                'MRATING'   :[OlxAPIConst.BK_dRating2     ,'(float) Breaker momentary rating'],
                'NODERATE'  :[OlxAPIConst.BK_nDontDerate  ,'(int) Breaker do not derate in reclosing operation flag: 1-true; 0-false;'],
                'FLAG'      :[OlxAPIConst.BK_nInService   ,'(int) Breaker In-service flag: 1-true; 2-false;'],
                'GROUPTYPE1':[OlxAPIConst.BK_nInterrupt1  ,'(int) Breaker group 1 interrupting current: 1-max current; 0-group current'],
                'GROUPTYPE2':[OlxAPIConst.BK_nInterrupt2  ,'(int) Breaker group 1 interrupting current: 1-max current; 0-group current'],
                'RATINGTYPE':[OlxAPIConst.BK_nRatingType  ,'(int) Breaker rating type: 0- symmetrical current basis;1- total current basis; 2- IEC'],
                'OPS1'      :[OlxAPIConst.BK_nTotalOps1   ,'(int) Breaker total operations for group 1'],
                'OPS2'      :[OlxAPIConst.BK_nTotalOps2   ,'(int) Breaker total operations for group 2'],
                'NAME'      :[OlxAPIConst.BK_sID          ,'(str) Breaker name (ID)'],
                'RECLOSE1'  :[OlxAPIConst.BK_vdRecloseInt1,'[float]*3 Breaker reclosing intervals for group 1 (s)'],
                'RECLOSE2'  :[OlxAPIConst.BK_vdRecloseInt2,'[float]*3 Breaker reclosing intervals for group 2 (s)'],
                'BUS'       :[OlxAPIConst.BK_nBusHnd      ,'(BUS) Breaker bus'],
                'OBJLST1'   :[OlxAPIConst.BK_vnG1DevHnd   ,'[EQUIPMENT]*10 Breaker protected equipment group 1 List of equipment'],
                'G1OUTAGES' :[OlxAPIConst.BK_vnG1OutageHnd,'[EQUIPMENT]*10 Breaker protected equipment group 1 List of additional outage'],
                'OBJLST2'   :[OlxAPIConst.BK_vnG2DevHnd   ,'[EQUIPMENT]*10 Breaker protected equipment group 2 List of equipment'],
                'G2OUTAGES' :[OlxAPIConst.BK_vnG2OutageHnd,'[EQUIPMENT]*10 Breaker protected equipment group 2 List of additional outage']}
__OLXOBJ_PARA__['MULINE'] = {\
                'FROM1'     :[OlxAPIConst.MU_vdFrom1  ,'[float]*5 Mutual coupling pair line 1 From percent'],
                'FROM2'     :[OlxAPIConst.MU_vdFrom2  ,'[float]*5 Mutual coupling pair line 2 From percent'],
                'TO1'       :[OlxAPIConst.MU_vdTo1    ,'[float]*5 Mutual coupling pair line1 To percent'],
                'TO2'       :[OlxAPIConst.MU_vdTo2    ,'[float]*5 Mutual coupling pair line2 To percent'],
                'R'         :[OlxAPIConst.MU_vdR      ,'[float]*5 Mutual coupling pair R'],
                'X'         :[OlxAPIConst.MU_vdX      ,'[float]*5 Mutual coupling pair X'],
                'LINE1'     :[OlxAPIConst.MU_nHndLine1,'(LINE) Mutual coupling pair line 1'],
                'LINE2'     :[OlxAPIConst.MU_nHndLine2,'(LINE) Mutual coupling pair line 2']}
#
__OLXOBJ_PARA__['RLYGROUP'] = {\
                'INTRPTIME' :[OlxAPIConst.RG_dBreakerTime ,'(float) Relay group interrupting time (cycles)'],
                'FLAG'      :[OlxAPIConst.RG_nInService   ,'(int) Relay group in-service flag: 1- active; 2- out-of-service'],
                'OPFLAG'    :[OlxAPIConst.RG_nOps         ,'(int) Relay group total operations'],
                'NOTE'      :[OlxAPIConst.RG_sNote        ,'(str) Relay group annotation'],
                'RECLSRTIME':[OlxAPIConst.RG_vdRecloseInt ,'[float]*4 Relay group reclosing intervals'],
                'BACKUP'    :[OlxAPIConst.RG_nBackupHnd   ,'(RLYGROUP) Relay group back up group'],
                'EQUIPMENT' :[OlxAPIConst.RG_nBranchHnd   ,'(EQUIPMENT) Relay group branch'],
                'PRIMARY'   :[OlxAPIConst.RG_nPrimaryHnd  ,'(RLYGROUP) Relay group primary group'],
                'LOGICRECL' :[OlxAPIConst.RG_nReclLogicHnd,'(SCHEME) Relay group reclose logic scheme'],
                'LOGICTRIP' :[OlxAPIConst.RG_nTripLogicHnd,'(SCHEME) Relay group trip logic scheme']}
#
__OLXOBJ_PARA__['RLYOCG'] = {\
                'ID'        :[OlxAPIConst.OG_sID          ,'(str) OC ground relay ID'],
                'ASSETID'   :[OlxAPIConst.OG_sAssetID     ,'(str) OC ground relay asset ID'],
                'DATEOFF'   :[OlxAPIConst.OG_sOffDate     ,'(str) OC ground relay out of service date'],
                'DATEON'    :[OlxAPIConst.OG_sOnDate      ,'(str) OC ground relay in service date'],
                'FLAG'      :[OlxAPIConst.OG_nInService   ,'(int) OC ground relay in-service flag: 1- active; 2- out-of-service'],
                'RLYGROUP'  :[OlxAPIConst.OG_nRlyGrHnd    ,'(RLYGROUP) OC ground relay group']}
#
__OLXOBJ_RLYSET__['RLYOCG'] = {\
                'CT'         :[OlxAPIConst.OG_dCT          ,'(float) OC ground relay CT ratio'],
                'OPI'        :[OlxAPIConst.OG_nOperateOn   ,'(int) OC ground relay Operate On: 0-3I0; 1-3I2; 2-I0; 3-I2'],
                'ASYM'       :[OlxAPIConst.OG_nDCOffset    ,'(int) OC ground relay sentitive to DC offset:1-true; 0-false'],
                'FLATINST'   :[OlxAPIConst.OG_nFlatDelay   ,'(int) OC ground relay flat definite time delay flag: 1-true; 0-false'],
                'SGNL'       :[OlxAPIConst.OG_nSignalOnly  ,'(int) OC ground relay signal only: 1-INST;2-OC, 4-DT1'],
                'OCDIR'      :[OlxAPIConst.OG_nDirectional ,'(int) OC ground relay directional flag: 0=None;1=Fwd.;2=Rev.'],
                'DTDIR'      :[OlxAPIConst.OG_nIDirectional,'(int) OC ground relay Inst. Directional flag: 0=None;1=Fwd.;2=Rev.'],
                'POLAR'      :[OlxAPIConst.OG_nPolar       ,'(int) OC ground relay polar option 0=V0,I0;1=V2,I2;2=SEL V2;3=SEL V0'],
                'PICKUPTAP'  :[OlxAPIConst.OG_dTap         ,'(float) OC ground relay Pickup (A)'],
                'TDIAL'      :[OlxAPIConst.OG_dTDial       ,'(float) OC ground relay time dial'],
                'TIMEADD'    :[OlxAPIConst.OG_dTimeAdd     ,'(float) OC ground relay time adder'],
                'TIMEMULT'   :[OlxAPIConst.OG_dTimeMult    ,'(float) OC ground relay time multiplier'],
                'TIMERESET'  :[OlxAPIConst.OG_dResetTime   ,'(float) OC ground relay reset time'],
                'DTTIMEADD'  :[OlxAPIConst.OG_dTimeAdd2    ,'(float) OC ground relay time adder for INST/DTx'],
                'DTTIMEMULT' :[OlxAPIConst.OG_dTimeMult2   ,'(float) OC ground relay time multiplier for INST/DTx'],
                'INSTSETTING':[OlxAPIConst.OG_dInst        ,'(float) OC ground relay instantaneous setting'],
                'DTPICKUP'   :[OlxAPIConst.OG_vdDTPickup   ,'[float]*5 OC ground relay Pickups Sec.A '],
                'DTDELAY'    :[OlxAPIConst.OG_vdDTDelay    ,'[float]*5 OC ground relay Delays seconds'],
                'MINTIME'    :[OlxAPIConst.OG_dMinTripTime ,'(float) OC ground relay minimum trip time'],
                'TYPE'       :[OlxAPIConst.OG_sType        ,'(str) OC ground relay type'],
                'TAPTYPE'    :[OlxAPIConst.OG_sTapType     ,'(str) OC ground relay tap type'],
                'LIBNAME'    :[OlxAPIConst.OG_sLibrary     ,'(str) OC ground relay Library'],
                'PACKAGE'    :[OlxAPIConst.OG_nPackage     ,'(int) OC ground relay Package option'],
                # polar = [0,1]
                'CHARANGLE'  :[OlxAPIConst.OG_vdDirSettingV15 ,[0,1],1,'(float) OC ground relay direction: Characteristic angle (deg.)'],
                'VPOL50PF'   :[OlxAPIConst.OG_vdDirSettingV15 ,[0,1],2,'(float) OC ground relay direction: Forward pickup (Sec.A)'],
                'VPOL50PR'   :[OlxAPIConst.OG_vdDirSettingV15 ,[0,1],3,'(float) OC ground relay direction: Reverse pickup (Sec.A)'],
                # polar = [2]
                '32QZ2'      :[OlxAPIConst.OG_vdDirSettingV15 ,[2],2,'(float) OC ground relay direction: Z2F Fwd threshold (Sec.ohm)'],
                '32Q50Q'     :[OlxAPIConst.OG_vdDirSettingV15 ,[2],7,'(float) OC ground relay direction: 50QF: Fwd 3I2 pickup(Sec.A)'],
                '32QZ2R'     :[OlxAPIConst.OG_vdDirSettingV15 ,[2],3,'(float) OC ground relay direction: Z2R: Rev threshold (Sec.ohm)'],#ccc
                '32Q50QR'    :[OlxAPIConst.OG_vdDirSettingV15 ,[2],8,'(float) OC ground relay direction: 50QR: Rev 3I2 pickup(Sec.A)'],#ccc
                '32QA2'      :[OlxAPIConst.OG_vdDirSettingV15 ,[2],4,'(float) OC ground relay direction: a2: I2/I1 restr.factor'],
                '32QK2'      :[OlxAPIConst.OG_vdDirSettingV15 ,[2],5,'(float) OC ground relay direction: k2: i2/I0 restr.factor'],
                '32QZ1ANG'   :[OlxAPIConst.OG_vdDirSettingV15 ,[2],6,'(float) OC ground relay direction: Z1ANG: Line Z1 angle (Deg.)'],
                '32QPTR'     :[OlxAPIConst.OG_vdDirSettingV15 ,[2],1,'(float) OC ground relay direction: PTR: PT ratio'],
                # polar = [3]
                '32GZ0'      :[OlxAPIConst.OG_vdDirSettingV15 ,[3],2,'(float) OC ground relay direction: Z0F Fwd threshold (Sec.ohm)'],
                '32G50G'     :[OlxAPIConst.OG_vdDirSettingV15 ,[3],6,'(float) OC ground relay direction: 50GF: Fwd 3Io pickup(Sec.A)'],
                '32GZ0R'     :[OlxAPIConst.OG_vdDirSettingV15 ,[3],3,'(float) OC ground relay direction: Z0R: Rev threshold (Sec.ohm)'],#ccc
                '32G50GR'    :[OlxAPIConst.OG_vdDirSettingV15 ,[3],7,'(float) OC ground relay direction: 50GR: Rev 3Io pickup(Sec.A)'],#ccc
                '32GA0'      :[OlxAPIConst.OG_vdDirSettingV15 ,[3],4,'(float) OC ground relay direction: a0: I0/I1 restr.factor'],
                '32GZ0ANG'   :[OlxAPIConst.OG_vdDirSettingV15 ,[3],5,'(float) OC ground relay direction: Z0ANG: Line Z0 angle (Deg.)'],
                '32GPTR'     :[OlxAPIConst.OG_vdDirSettingV15 ,[3],1,'(float) OC ground relay direction: PTR: PT ratio']}
#
__OLXOBJ_PARA__['RLYOCP'] = {\
                'ID'         :[OlxAPIConst.OP_sID           ,'(str) OC phase relay ID'],
                'ASSETID'    :[OlxAPIConst.OP_sAssetID      ,'(str) OC phase relay asset ID'],
                'DATEOFF'    :[OlxAPIConst.OP_sOffDate      ,'(str) OC phase relay out of service date'],
                'DATEON'     :[OlxAPIConst.OP_sOnDate       ,'(str) OC phase relay in service date'],
                'FLAG'       :[OlxAPIConst.OP_nInService    ,'(int) OC phase relay in-service flag: 1- active; 2- out-of-service'],
                'RLYGROUP'   :[OlxAPIConst.OP_nRlyGrHnd     ,'(RLYGROUP) OC phase relay group']}
#
__OLXOBJ_RLYSET__['RLYOCP'] = {\
                'CT'         :[OlxAPIConst.OP_dCT           ,'(float) OC phase relay CT ratio'],
                'CTCONNECT'  :[OlxAPIConst.OP_nByCTConnect  ,'(int) OC phase reley CT connection: 0- Wye; 1-Delta'],
                'ASYM'       :[OlxAPIConst.OP_nDCOffset     ,'(int) OC phase relay sentitive to DC offset: 1-true; 0-false'],
                'FLATINST'   :[OlxAPIConst.OP_nFlatDelay    ,'(int) OC phase relay flat delay: 1-true; 0-false'],
                'SGNL'       :[OlxAPIConst.OP_nSignalOnly   ,'(int) OC phase relay signal only:  1-INST;2-OC, 4-DT1'],
                'OCDIR'      :[OlxAPIConst.OP_nDirectional  ,'(int) OC phase relay directional flag: 0=None;1=Fwd.;2=Rev.'],
                'DTDIR'      :[OlxAPIConst.OP_nIDirectional ,'(int) OC phase relay Inst. Directional flag: 0=None;1=Fwd.;2=Rev.'],
                'POLAR'      :[OlxAPIConst.OP_nPolar        ,'(int) OC_phase relay polar option'],
                'VOLTCONTROL':[OlxAPIConst.OP_nVoltControl  ,'(int) OC phase relay voltage controlled or restrained'],
                'PICKUPTAP'  :[OlxAPIConst.OP_dTap          ,'(float) OC phase relay tap Ampere'],
                'TDIAL'      :[OlxAPIConst.OP_dTDial        ,'(float) OC phase relay time dial'],
                'TIMEADD'    :[OlxAPIConst.OP_dTimeAdd      ,'(float) OC phase relay time adder'],
                'TIMEMULT'   :[OlxAPIConst.OP_dTimeMult     ,'(float) OC phase relay time multiplier'],
                'TIMERESET'  :[OlxAPIConst.OP_dResetTime    ,'(float) OC phase relay reset time'],
                'VOLTPERCENT':[OlxAPIConst.OP_dVCtrlRestPcnt,'(float) OC_phase relay voltage controlled or restrained percentage'],
                'DTTIMEADD'   :[OlxAPIConst.OP_dTimeAdd2    ,'(float) OC phase relay time adder for INST/DTx'],
                'DTTIMEMULT'  :[OlxAPIConst.OP_dTimeMult2   ,'(float) OC phase relay time multiplier for INST/DTx'],
                'INSTSETTING':[OlxAPIConst.OP_dInst         ,'(float) OC phase relay instantaneous setting'],
                'DTPICKUP'   :[OlxAPIConst.OP_vdDTPickup    ,'[float]*8 OC phase relay Pickup'],
                'DTDELAY'    :[OlxAPIConst.OP_vdDTDelay     ,'[float]*8 OC phase relay Delay' ],
                'MINTIME'    :[OlxAPIConst.OP_dMinTripTime  ,'(float) OC phase relay minimum trip time'],
                'TYPE'       :[OlxAPIConst.OP_sType         ,'(str) OC phase relay type'],
                'TAPTYPE'    :[OlxAPIConst.OP_sTapType      ,'(str) OC phase relay tap type'],
                'LIBNAME'    :[OlxAPIConst.OP_sLibrary      ,'(str) OC phase relay Library'],
                'PACKAGE'    :[OlxAPIConst.OP_nPackage      ,'(int) OC phase relay Package option'],
                # polar = [0]
                'CHARANGLE'  :[OlxAPIConst.OP_vdDirSettingV15 ,[0],1,'(float) OC phase relay direction: Characteristic angle (deg.)'],
                'VPOL50PF'   :[OlxAPIConst.OP_vdDirSettingV15 ,[0],2,'(float) OC phase relay direction: Forward pickup (Sec.A)'],
                'VPOL50PR'   :[OlxAPIConst.OP_vdDirSettingV15 ,[0],3,'(float) OC phase relay direction: Reverse pickup (Sec.A)'],
                # polar = [2]
                '32QZ2'      :[OlxAPIConst.OP_vdDirSettingV15 ,[2],2,'(float) OC phase relay direction: Z2F Fwd threshold (Sec.ohm)'],
                '32Q50Q'     :[OlxAPIConst.OP_vdDirSettingV15 ,[2],7,'(float) OC phase relay direction: 50QF: Fwd 3I2 pickup(Sec.A)'],
                '32QZ2R'     :[OlxAPIConst.OP_vdDirSettingV15 ,[2],3,'(float) OC phase relay direction: Z2R: Rev threshold (Sec.ohm)'],#ccc
                '32Q50QR'    :[OlxAPIConst.OP_vdDirSettingV15 ,[2],8,'(float) OC phase relay direction: 50QR: Rev 3I2 pickup(Sec.A)'],#ccc
                '32QA2'      :[OlxAPIConst.OP_vdDirSettingV15 ,[2],4,'(float) OC phase relay direction: a2: I2/I1 restr.factor'],
                '32QK2'      :[OlxAPIConst.OP_vdDirSettingV15 ,[2],5,'(float) OC phase relay direction: k2: i2/I0 restr.factor'],
                '32QZ1ANG'   :[OlxAPIConst.OP_vdDirSettingV15 ,[2],6,'(float) OC phase relay direction: Z1ANG: Line Z1 angle (Deg.)'],
                '32QPTR'     :[OlxAPIConst.OP_vdDirSettingV15 ,[2],1,'(float) OC phase relay direction: PTR: PT ratio']}
#
__OLXOBJ_PARA__['FUSE'] = {\
                'ID'        :[OlxAPIConst.FS_sID       ,'(str) Fuse name(ID)'],
                'ASSETID'   :[OlxAPIConst.FS_sAssetID  ,'(str) Fuse asset ID'],
                'DATEOFF'   :[OlxAPIConst.FS_sOffDate  ,'(str) Fuse out of service date'],
                'DATEON'    :[OlxAPIConst.FS_sOnDate   ,'(str) Fuse in service date'],
                'FLAG'      :[OlxAPIConst.FS_nInService,'(int) Fuse in-service flag: 1- active; 2- out-of-service'],
                'FUSECURDIV':[OlxAPIConst.FS_dCT       ,'(float) Fuse current divider'],
                'RATING'    :[OlxAPIConst.FS_dRating   ,'(float) Fuse Rating'],
                'TIMEMULT'  :[OlxAPIConst.FS_dTimeMult ,'(float) Fuse time multiplier'],
                'FUSECVE'   :[OlxAPIConst.FS_nCurve    ,'(int) Fuse Compute time using flag: 1- minimum melt; 2- Total clear'],
                'PACKAGE'   :[OlxAPIConst.FS_nPackage  ,'(int) Fuse Package option'],
                'LIBNAME'   :[OlxAPIConst.FS_sLibrary  ,'(str) Fuse Library'],
                'TYPE'      :[OlxAPIConst.FS_sType     ,'(str) Fuse type'],
                'RLYGROUP'  :[OlxAPIConst.FS_nRlyGrHnd ,'(RLYGROUP) Fuse relay group']}
#
__OLXOBJ_PARA__['RLYDSG'] = {\
                'ID'          :[OlxAPIConst.DG_sID         ,'(str) DS ground relay name(ID)'],
                'TYPE'        :[OlxAPIConst.DG_sType       ,'(str) DS ground relay type (ID2)'],
                'ASSETID'     :[OlxAPIConst.DG_sAssetID    ,'(str) DS ground relay asset ID'],
                'DATEOFF'     :[OlxAPIConst.DG_sOffDate    ,'(str) DS ground relay out of service date'],
                'DATEON'      :[OlxAPIConst.DG_sOnDate     ,'(str) DS ground relay in service date'],
                'FLAG'        :[OlxAPIConst.DG_nInService  ,'(int) DS ground relay in-service flag: 1- active; 2- out-of-service'],
                'RLYGROUP'    :[OlxAPIConst.DG_nRlyGrHnd   ,'(RLYGROUP) DS ground relay group']}
#
__OLXOBJ_RLYSET__['RLYDSG'] = {\
                'DSTYPE'      :[OlxAPIConst.DG_sDSType     ,'(str) DS ground relay type name'],
                'Z2OCTYPE'    :[OlxAPIConst.DG_sZ2OCCurve  ,'(str) DS ground zone 2 OC supervision type name'],
                'CT'          :[OlxAPIConst.DG_dCT         ,'(float) DS ground relay CT ratio'],
                'VT'          :[OlxAPIConst.DG_dVT         ,'(float) DS ground relay VT ratio'],
                'VTBUS'       :[OlxAPIConst.DG_nVTBus      ,'(BUS) DS ground relay VT at Bus'],
                'Z2OCPICKUP'  :[OlxAPIConst.DG_dZ2OCPickup ,'(float) DS ground relay Z2 OC supervision Pickup(A)'],
                'KANG1'       :[OlxAPIConst.DG_dKang       ,'(float) DS ground relay Zero sequence compensation factor K - angle'],
                'KMAG1'       :[OlxAPIConst.DG_dKmag       ,'(float) DS ground relay Zero sequence compensation factor K - magnitude'],
                'KANG'        :[OlxAPIConst.DG_dKangEx     ,'(float) DS ground relay positive sequence compensation factor K - angle'],
                'KMAG'        :[OlxAPIConst.DG_dKmagEx     ,'(float) DS ground relay positive sequence compensation factor K - magnitude'],
                'MINI'        :[OlxAPIConst.DG_dMinI       ,'(float) DS ground relay Min I'],
                'Z2OCTD'      :[OlxAPIConst.DG_dZ2OCTD     ,'(float) DS ground relay Z2 OC supervision time dial'],
                'SNLZONE'     :[OlxAPIConst.DG_nSignalOnly ,'(int) DS ground signal-only zone'],
                'OCLIBNAME'   :[OlxAPIConst.DG_sLibrary    ,'(str) DS ground relay Library'],
                'PACKAGE'     :[OlxAPIConst.DG_nPackage    ,'(int) DS ground relay Package option'],
                'DELAY'       :[OlxAPIConst.DG_vdDelay     ,'[float]*8 DS ground relay zone delay'],
                'REARCH'      :[OlxAPIConst.DG_vdReach     ,'[float]*8 DS ground relay zone reach'],
                'REARCH1'     :[OlxAPIConst.DG_vdReach1    ,'[float]*8 DS ground relay zone reach 1']}
#
__OLXOBJ_PARA__['RLYDSP'] = {\
                'ID'          :[OlxAPIConst.DP_sID         ,'(str) DS phase relay ID'],
                'Z2OCTYPE'    :[OlxAPIConst.DP_sZ2OCCurve  ,'(str) DS ground zone 2 OC supervision type name'],
                'TYPE'        :[OlxAPIConst.DP_sType       ,'(str) DS phase relay type (ID2)'],
                'ASSETID'     :[OlxAPIConst.DP_sAssetID    ,'(str) DS phase relay asset ID'],
                'DATEOFF'     :[OlxAPIConst.DP_sOffDate    ,'(str) DS phase relay out of service date'],
                'DATEON'      :[OlxAPIConst.DP_sOnDate     ,'(str) DS phase relay in service date'],
                'FLAG'        :[OlxAPIConst.DP_nInService  ,'(int) DS phase relay in-service flag: 1- active; 2- out-of-service'],
                'RLYGROUP'    :[OlxAPIConst.DP_nRlyGrHnd   ,'(RLYGROUP) DS phase relay group']}

__OLXOBJ_RLYSET__['RLYDSP'] = {\
                'CT'          :[OlxAPIConst.DP_dCT         ,'(float) DS phase relay CT ratio'],
                'VT'          :[OlxAPIConst.DP_dVT         ,'(float) DS phase relay VT ratio'],
                'VTBUS'       :[OlxAPIConst.DP_nVTBus      ,'(float) DS phase relay VT at Bus'],
                'Z2OCPICKUP'  :[OlxAPIConst.DP_dZ2OCPickup ,'(float) DS ground relay Z2 OC supervision Pickup(A)'],
                'MINI'        :[OlxAPIConst.DP_dMinI       ,'(float) DS phase relay Min I'],
                'Z2OCTD'      :[OlxAPIConst.DP_dZ2OCTD     ,'(float) DS ground relay Z2 OC supervision time dial'],
                'PACKAGE'     :[OlxAPIConst.DP_nPackage    ,'(int) DS phase relay Package option'],
                'SNLZONE'     :[OlxAPIConst.DP_nSignalOnly ,'(int) DS phase signal-only zone flag'],
                'DSTYPE'      :[OlxAPIConst.DP_sDSType     ,'(str) DS phase relay type name'],
                'OCLIBNAME'   :[OlxAPIConst.DP_sLibrary    ,'(str) DS phase relay Library'],
                'DELAY'       :[OlxAPIConst.DP_vdDelay     ,'[float]*8 DS phase relay zone delay'],
                'REACH'       :[OlxAPIConst.DP_vdReach     ,'[float]*8 DS phase relay zone reach'],
                'REACH1'      :[OlxAPIConst.DP_vdReach1    ,'[float]*8 DS phase relay alternat zone reach']}
#
__OLXOBJ_PARA__['RLYD'] = {\
                'ID'          :[OlxAPIConst.RD_sID          ,'(str) Differential relay ID (NAME)'],
                'ASSETID'     :[OlxAPIConst.RD_sAssetID     ,'(str) Differential relay asset ID'],
                'DATEOFF'     :[OlxAPIConst.RD_sOffDate     ,'(str) Differential relay out of service date'],
                'DATEON'      :[OlxAPIConst.RD_sOnDate      ,'(str) Differential relay in service date'],
                'FLAG'        :[OlxAPIConst.RD_nInService   ,'(int) Differential relay in service: 1- active; 2- out-of-service'],
                'PACKAGE'     :[OlxAPIConst.RD_nInService   ,'(int) Differential relay Package option'],
                'RLYGROUP'    :[OlxAPIConst.RD_nRlyGrpHnd   ,'(RLYGROUP) Differential relay group'],
                'CTR1'        :[OlxAPIConst.RD_dCTR1        ,'(float) Differential relay CTR1'],
                'IMIN3I0'     :[OlxAPIConst.RD_dPickup3I0   ,'(float) Differential relay minimum enable differential current (3I0)'],
                'IMIN3I2'     :[OlxAPIConst.RD_dPickup3I2   ,'(float) Differential relay minimum enable differential current (3I2)'],
                'IMINPH'      :[OlxAPIConst.RD_dPickupPh    ,'(float) Differential relay minimum enable differential current (phase)'],
                'TLCTD3I0'    :[OlxAPIConst.RD_dTLCTDDelayI0,'(float) Differential relay tapped load coordination delay (I0)'],
                'TLCTDI2'     :[OlxAPIConst.RD_dTLCTDDelayI2,'(float) Differential relay tapped load coordination delay (I2)'],
                'TLCTDPH'     :[OlxAPIConst.RD_dTLCTDDelayPh,'(float) Differential relay tapped load coordination delay (phase)'],
                'SGLONLY'     :[OlxAPIConst.RD_nSignalOnly  ,'(int) Differential relay signal only : 1-true; 0-false'],
                'TLCCV3I0'    :[OlxAPIConst.RD_sTLCCurveI0  ,'(str) Differential relay tapped load coordination curve (I0)'],
                'TLCCVI2'     :[OlxAPIConst.RD_sTLCCurveI2  ,'(str) Differential relay tapped load coordination curve (I2)'],
                'TLCCVPH'     :[OlxAPIConst.RD_sTLCCurvePh  ,'(str) Differential relay tapped load coordination curve (phase)'],
                'CTGRP1'      :[OlxAPIConst.RD_nLocalCTHnd1 ,'(RLYGROUP) Differential relay local current input 1'],
                'RMTE1'       :[OlxAPIConst.RD_nRmeDevHnd1  ,'(EQUIPMENT) Differential relay remote device 1'],
                'RMTE2'       :[OlxAPIConst.RD_nRmeDevHnd2  ,'(EQUIPMENT) Differential relay remote device 2']}
#
__OLXOBJ_PARA__['RLYV'] = {\
                'ID'        :[OlxAPIConst.RV_sID         ,'(str) Voltage relay ID (NAME)'],
                'ASSETID'   :[OlxAPIConst.RV_sAssetID    ,'(str) Voltage relay asset ID'],
                'DATEOFF'   :[OlxAPIConst.RV_sOffDate    ,'(str) Voltage relay out of service date'],
                'DATEON'    :[OlxAPIConst.RV_sOnDate     ,'(str) Voltage relay in service date'],
                'FLAG'      :[OlxAPIConst.RV_nInService  ,'(int) Voltage relay in service: 1- active; 2- out-of-service'],
                'PACKAGE'   :[OlxAPIConst.RV_nInService  ,'(int) Voltage relay Package option'],
                'PTR'       :[OlxAPIConst.RV_dCTR        ,'(float) Voltage relay PT ratio'],
                'OVINST'    :[OlxAPIConst.RV_dOVIPickup  ,'(float) Voltage relay over-voltage instant pickup (V)'],
                'UVINST'    :[OlxAPIConst.RV_dUVIPickup  ,'(float) Voltage relay under-voltage instant pickup (V)'],
                'OVPICKUP'  :[OlxAPIConst.RV_dOVTPickup  ,'(float) Voltage relay over-voltage pickup (V)'],
                'UVPICKUP'  :[OlxAPIConst.RV_dUVTPickup  ,'(float) Voltage relay under-voltage pickup (V)'],
                'UVDELAYTD' :[OlxAPIConst.RV_dUVTDelay   ,'(float) Voltage relay under-voltage delay'],
                'OVDELAYTD' :[OlxAPIConst.RV_dOVTDelay   ,'(float) Voltage relay over-voltage delay'],
                'SGLONLY'   :[OlxAPIConst.RV_nSignalOnly ,'(int) Voltage relay signal only 0-No check; 1-Check over-voltage instant;2-Check over-voltage delay;4-Check under-voltage instant; 8-Check under-voltage delay '],
                'OPQTY'     :[OlxAPIConst.RV_nVoltOperate,'(int) Voltage relay operate on voltage option: 1-Phase-to-Neutral; 2- Phase-to-Phase; 3-3V0;4-V1;5-V2;6-VA;7-VB;8-VC;9-VBC;10-VAB;11-VCA'],
                'OVCVR'     :[OlxAPIConst.RV_sOVCurve    ,'(str) Voltage relay over-voltage element curve'],
                'UVCVR'     :[OlxAPIConst.RV_sUVCurve    ,'(str) Voltage relay under-voltage element curve'],
                'RLYGROUP'  :[OlxAPIConst.RV_nRlyGrpHnd  ,'(RLYGROUP) Voltage relay group']}
#
__OLXOBJ_PARA__['RECLSR'] = {\
                'ID'          :[OlxAPIConst.CP_sID         ,'(str) Recloser ID'],
                'ASSETID'     :[OlxAPIConst.CP_sAssetID    ,'(str) Recloser AssetID'],
                'DATEOFF'     :[OlxAPIConst.CP_sOffDate    ,'(str) Recloser out of service date'],
                'DATEON'      :[OlxAPIConst.CP_sOnDate     ,'(str) Recloser in service date'],
                'FLAG'        :[OlxAPIConst.CP_nInService  ,'(int) Recloser in service flag: 1- active; 2- out-of-service'],
                'RLYGROUP'    :[OlxAPIConst.CP_nRlyGrHnd   ,'(RLYGROUP) Recloser relay group'],
                'TOTALOPS'    :[OlxAPIConst.CP_nTotalOps   ,'(int) Recloser total operations to locked out'],
                'FASTOPS'     :[OlxAPIConst.CP_nFastOps    ,'(int) Recloser number of fast operations'],
                'RECLOSE1'    :[OlxAPIConst.CP_dRecIntvl1  ,'(float) Recloser reclosing interval 1'],
                'RECLOSE2'    :[OlxAPIConst.CP_dRecIntvl2  ,'(float) Recloser reclosing interval 2'],
                'RECLOSE3'    :[OlxAPIConst.CP_dRecIntvl3  ,'(float) Recloser reclosing interval 3'],
                'BYTADD'      :[OlxAPIConst.CP_nTAddAppl   ,'(int) Recloser Time adder modifies'],
                'BYTMULT'     :[OlxAPIConst.CP_nTMultAppl  ,'(int) Recloser Time multiplier modifies'],
                'LIBNAME'     :[OlxAPIConst.CP_sLibrary    ,'(str) Recloser Library'],
                'INTRPTIME'   :[OlxAPIConst.CP_dIntrTime   ,'(float) Recloser interrupting time (s)'],
                #
                'INST'        :[OlxAPIConst.CP_dHiAmps     ,OlxAPIConst.CG_dHiAmps     ,'(float) Recloser-Phase high current trip','(float) Recloser-Ground high current trip'],
                'INSTDELAY'   :[OlxAPIConst.CP_dHiAmpsDelay,OlxAPIConst.CG_dHiAmpsDelay,'(float) Recloser-Phase high current trip delay','(float) Recloser-Ground high current trip delay'],
                'MINTIMEF'    :[OlxAPIConst.CP_dMinTF      ,OlxAPIConst.CG_dMinTF      ,'(float) Recloser-Phase fast curve minimum time','(float) Recloser-Ground fast curve minimum time'],
                'MINTIMES'    :[OlxAPIConst.CP_dMinTS      ,OlxAPIConst.CG_dMinTS      ,'(float) Recloser-Phase slow curve minimum time','(float) Recloser-Ground slow curve minimum time'],
                'MINTRIPF'    :[OlxAPIConst.CP_dPickupF    ,OlxAPIConst.CG_dPickupF    ,'(float) Recloser-Phase fast curve pickup','(float) Recloser-Ground fast curve pickup'],
                'MINTRIPS'    :[OlxAPIConst.CP_dPickupS    ,OlxAPIConst.CG_dPickupS    ,'(float) Recloser-Phase slow curve pickup','(float) Recloser-Ground slow curve pickup'],
                'TIMEADDF'    :[OlxAPIConst.CP_dTimeAddF   ,OlxAPIConst.CG_dTimeAddF   ,'(float) Recloser-Phase fast curve time adder','(float) Recloser-Ground fast curve time adder'],
                'TIMEADDS'    :[OlxAPIConst.CP_dTimeAddS   ,OlxAPIConst.CG_dTimeAddS   ,'(float) Recloser-Phase slow curve time adder','(float) Recloser-Ground slow curve time adder'],
                'TIMEMULTF'   :[OlxAPIConst.CP_dTimeMultF  ,OlxAPIConst.CG_dTimeMultF  ,'(float) Recloser-Phase fast curve time multiplier','(float) Recloser-Ground fast curve time multiplier'],
                'TIMEMULTS'   :[OlxAPIConst.CP_dTimeMultS  ,OlxAPIConst.CG_dTimeMultS  ,'(float) Recloser-Phase slow curve time multiplier','(float) Recloser-Ground slow curve time multiplier'],
                'FASTTYPE'    :[OlxAPIConst.CP_sTypeFast   ,OlxAPIConst.CG_sTypeFast   ,'(str) Recloser-Phase fast curve','(str) Recloser-Ground fast curve'],
                'SLOWTYPE'    :[OlxAPIConst.CP_sTypeSlow   ,OlxAPIConst.CG_sTypeSlow   ,'(str) Recloser-Phase slow curve','(str) Recloser-Ground slow curve']}
#
__OLXOBJ_PARA__['SCHEME'] = {\
                'FLAG'        :[OlxAPIConst.LS_nInService ,'(int) Logic scheme in-service flag: 1- active; 2- out-of-service'],
                'SIGNALONLY'  :[OlxAPIConst.LS_nSignalOnly,'(int) Logic scheme signal only'],
                'ASSETID'     :[OlxAPIConst.LS_sAssetID   ,'(str) Logic scheme asset ID'],
                'PL_LOGICEQU' :[OlxAPIConst.LS_sEquation  ,'(str) Logic scheme equation'],
                'ID'          :[OlxAPIConst.LS_sID        ,'(str) Logic scheme ID'],
                'DATEOFF'     :[OlxAPIConst.LS_sOffDate   ,'(str) Logic scheme out of service date'],
                'DATEON'      :[OlxAPIConst.LS_sOnDate    ,'(str) Logic scheme in service date'],
                'NAME'        :[OlxAPIConst.LS_sScheme    ,'(str) Logic scheme name'],
                'PL_LOGICTERM':[OlxAPIConst.LS_sVariables ,'(str) Logic scheme variables details (one variable per line in the format: name=description)'],
                'RLYGROUP'    :[OlxAPIConst.LS_nRlyGrpHnd ,'(RLYGROUP) Logic scheme relay group']}
__OLXOBJ_PARA__['ZCORRECT'] = {}
__OLXOBJ_PARA__['AREA'] = {}
__OLXOBJ_PARA__['ZONE'] = {}
__OLXOBJ_PARA__['FT'] = {\
                'Mva'         :[OlxAPIConst.FT_dMVA     ,'(float) Fault MVA'],
                'RNt'         :[OlxAPIConst.FT_dRNt     ,'(float) Thevenin equivalent negative sequence resistance'],
                'RPt'         :[OlxAPIConst.FT_dRPt     ,'(float) Thevenin equivalent positive sequence resistance'],
                'RZt'         :[OlxAPIConst.FT_dRZt     ,'(float) Thevenin equivalent zero sequence resistance'],
                'Rt'          :[OlxAPIConst.FT_dRt      ,'(float) '], # #
                'XNt'         :[OlxAPIConst.FT_dXNt     ,'(float) Thevenin equivalent negative sequence reactance'],
                'XPt'         :[OlxAPIConst.FT_dXPt     ,'(float) Thevenin equivalent positive sequence reactance'],
                'Xr'          :[OlxAPIConst.FT_dXR      ,'(float) X/R ratio at fault point'],
                'XrANSI'      :[OlxAPIConst.FT_dXRANSI  ,'(float) ANSI X/R ratio at fault point'],
                'XZt'         :[OlxAPIConst.FT_dXZt     ,'(float) Thevenin equivalent zero sequence reactance'],
                'Xt'          :[OlxAPIConst.FT_dXt      ,'(float) '], # #
                'Nofaults'    :[OlxAPIConst.FT_nNOfaults,'(int) ']} # #
#
__OLXOBJ_oHND__ = {'BUS','BUS1','BUS2','BUS3','CNTBUS','LTCCTRL','VTBUS','RLYGROUP','RLYGROUP1','RLYGROUP2','RLYGROUP3',
               'CTGRP1','LOAD','SHUNT','GEN','LINE1','LINE2','EQUIPMENT','MULINE','PRIMARY','BACKUP','LOGICTRIP',
               'LOGICRECL','OBJLST1','OBJLST2','G1DEVS','G2DEVS','G1OUTAGES','G2OUTAGES','RMTE1','RMTE2'}
__OLXOBJ_fltConn__ = ['3LG','2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA',\
                 '1LG:A','1LG:B','1LG:C','LL:BC','LL:CA','LL:AB','LL:CB','LL:AC','LL:BA']
__OLXOBJ_fltConnSEA__ = {'3LG':1,\
                '2LG:BC':2,'2LG:CA':3,'2LG:AB':4,'2LG:CB':2,'2LG:AC':3,'2LG:BA':4,\
                '1LG:A':5,'1LG:B':6,'1LG:C':7,\
                'LL:BC':8,'LL:CA':9 ,'LL:AB':10,'LL:CB':8,'LL:AC':9 ,'LL:BA':10}
__OLXOBJ_fltConnCLS__ ={'3LG':[0,1],\
               '2LG:BC':[1,1],'2LG:CA':[1,2],'2LG:AB':[1,3],'2LG:CB':[1,1],'2LG:AC':[1,2],'2LG:BA':[1,3],\
               '1LG:A':[2,1] ,'1LG:B':[2,2] ,'1LG:C':[2,3],\
               'LL:BC':[3,1] ,'LL:CA':[3,2] ,'LL:AB':[3,3] ,'LL:CB':[3,1] ,'LL:AC':[3,2] ,'LL:BA':[3,3]}
__OLXOBJ_fltConnSIM__ = {'3LG':[0,0],\
               '2LG:BC':[1,1],'2LG:CA':[1,2],'2LG:AB':[1,0],'2LG:CB':[1,1],'2LG:AC':[1,2],'2LG:BA':[1,0],\
               '1LG:A':[2,0] ,'1LG:B':[2,1] ,'1LG:C':[2,2],\
               'LL:BC':[3,1] ,'LL:CA':[3,2] ,'LL:AB':[3,0] ,'LL:CB':[3,1] ,'LL:AC':[3,2] ,'LL:BA':[3,0]}
#
def __getEquipment__(sType,scope=None):
    """
    return : OBJ in system with filter criteria
            - areaNum
            - zoneNum
            - kV = [kVmin kVmax]
    """
    if type(sType)==list:
        res = []
        for sType1 in sType:
            res.extend (__getEquipment__(sType1,scope))
        return res
    #
    try:
        tc = __OLXOBJ_CONST__[sType][0]
    except:
        se = '\nString parameter available for __getEquipment__(str):\n'
        se += str(__OLXOBJ_LIST__) +'\n'
        se += "\nNot found: '%s'"%sType
        raise Exception(se)
    #
    res = []
    hnd = c_int(0)
    while OlxAPIConst.OLXAPI_OK == OlxAPI.GetEquipment(c_int(tc), byref(hnd)):
        r = __getOBJ__(hnd.value,tc=tc)
        # test Scope
        if __isInScope__(r,scope):
            res.append(r)
    return res
#
def __getBusEquipment__(bus, sType):
    """
    Retrieves all equipment of a given type sType that is attached to bus
    - bus : BUS
    - sType : in OLXOBJ_BUS
    """
    res = []
    if type(sType)==list:
        for sType1 in sType:
            res.extend(__getBusEquipment__(bus,sType1))
        return res
    #
    if sType not in __OLXOBJ_BUS__:
        se = '\nString parameter available for __getBusEquipment__(bus,sType):\n'
        se += "\nNot found: '%s'"%sType
        raise Exception(se)
    #
    __check_currFileIdx__(bus)
    hnd = bus.__hnd__
    val1,val2 = c_int(0), c_int(0)
    if sType=='TERMINAL':
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            o1 = __getOBJ__(val1.value,tc=OlxAPIConst.TC_BRANCH)
            res.append(o1)
        return res
    #
    if sType=='RLYGROUP':
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(val1.value,c_int(OlxAPIConst.BR_nRlyGrp1Hnd),byref(val2)):
                rg1 = RLYGROUP(hnd=val2.value)
                res.append(rg1)
        return res
    #
    tc = __OLXOBJ_CONST__[sType][0]
    if sType in __OLXOBJ_EQUIPMENT__ :
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            e1 = __getDatai__(val1.value,OlxAPIConst.BR_nHandle)
            tc1 = OlxAPI.EquipmentType(e1)
            if tc1==tc:
                o1 = __getOBJ__(e1,tc=tc)
                res.append(o1)
        return res
    else:
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(tc),byref(val1)):
            o1 = __getOBJ__(val1.value,tc=tc)
            res.append(o1)
            if sType in __OLXOBJ_BUS1__:
                break
    return res
#
def __getBusEquipmentHnd__(bus, sType):
    """
    Retrieves all equipment hnd of a given type sType that is attached to bus
    - bus : BUS
    - sType : in __OLXOBJ_BUS__
    """
    if sType not in __OLXOBJ_BUS__:
        se = '\nString parameter available for __getBusEquipmentHnd__(bus,sType):\n'
        se += "\nNot found: '%s'"%sType
        raise Exception(se)
    #
    res = []
    if type(bus)==BUS:
        hnd = bus.__hnd__
        __check_currFileIdx__(bus)
    else:
        hnd = bus
    val1,val2 = c_int(0), c_int(0)
    if sType=='TERMINAL':
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            res.append(val1.value)
    #
    elif sType=='RLYGROUP':
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(val1.value,c_int(OlxAPIConst.BR_nRlyGrp1Hnd),byref(val2)):
                res.append(val2.value)
    else:
        tc = __OLXOBJ_CONST__[sType][0]
        if sType in __OLXOBJ_EQUIPMENT__ :
            while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(OlxAPIConst.TC_BRANCH),byref(val1)):
                e1 = __getDatai__(val1.value,OlxAPIConst.BR_nHandle)
                tc1 = OlxAPI.EquipmentType(e1)
                if tc1==tc:
                    res.append(e1)
        else:
            while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd,c_int(tc),byref(val1)):
                res.append(val1.value)
                if sType in __OLXOBJ_BUS1__:
                    break
    return res
#
def __getDatai__(hnd,paraCode):
    val1 = c_int(0)
    if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.GetData(hnd,c_int(paraCode),byref(val1)):
        return None
    return val1.value
#
def __getData__(hnd,paraCode):
    """
    Find data of element (line/bus/...)
        Args :
            hnd     : ([] or INT) handle element
            paraCode: ([] or INT) code data (BUS_,LN_,...)
        Returns:
            data    : SCALAR or []  or [][]
        Raises:
            OlxAPIException
    """
    if type(hnd)!=list and type(paraCode)!=list:
        vt = paraCode//100
        if vt == OlxAPIConst.VT_STRING:
            val1 = create_string_buffer(b'\000' * 128)
        elif vt in [0,OlxAPIConst.VT_INTEGER]:
            val1 = c_int(0)
        elif vt == OlxAPIConst.VT_DOUBLE:
            val1 = c_double(0)
        elif vt in [OlxAPIConst.VT_ARRAYDOUBLE,OlxAPIConst.VT_ARRAYSTRING,OlxAPIConst.VT_ARRAYINT]:
            val1 = create_string_buffer(b'\000' * 10*1024)
        else:
            OlxAPI.OlxAPIException("Error of paraCode")
        #
        if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.GetData(hnd,c_int(paraCode),byref(val1)):
            return None
        #
        return __getValue__(hnd,paraCode,val1)
    #
    res = []
    if type(hnd)==list :
        if type(paraCode)!=list:
            for hnd1 in hnd:
                res.append(__getData__(hnd1,paraCode))
        else:
            for hnd1 in hnd:
                r1 = []
                for paraCode1 in paraCode:
                    r1.append(__getData__(hnd1,paraCode1))
                res.append(r1)
    else:
        for paraCode1 in paraCode:
            res.append(__getData__(hnd,paraCode1))
    return res
#
def __getOBJ__(hnd,tc=None,sPara=None):
    if hnd==None:
        return None
    if tc!=None:
        if tc==OlxAPIConst.TC_BUS:
            return BUS(hnd=hnd)
        elif tc==OlxAPIConst.TC_GEN:
            return GEN(hnd=hnd)
        elif tc==OlxAPIConst.TC_GENUNIT:
            return GENUNIT(hnd=hnd)
        elif tc==OlxAPIConst.TC_GENW3:
            return GENW3(hnd=hnd)
        elif tc==OlxAPIConst.TC_GENW4:
            return GENW4(hnd=hnd)
        elif tc==OlxAPIConst.TC_CCGEN:
            return CCGEN(hnd=hnd)
        elif tc==OlxAPIConst.TC_XFMR:
            return XFMR(hnd=hnd)
        elif tc==OlxAPIConst.TC_XFMR3:
            return XFMR3(hnd=hnd)
        elif tc==OlxAPIConst.TC_PS:
            return SHIFTER(hnd=hnd)
        elif tc==OlxAPIConst.TC_LINE:
            return LINE(hnd=hnd)
        elif tc==OlxAPIConst.TC_DCLINE2:
            return DCLINE2(hnd=hnd)
        elif tc==OlxAPIConst.TC_SCAP:
            return SERIESRC(hnd=hnd)
        elif tc==OlxAPIConst.TC_SWITCH:
            return SWITCH(hnd=hnd)
        elif tc==OlxAPIConst.TC_MU:
            return MULINE(hnd=hnd)
        elif tc==OlxAPIConst.TC_LOAD:
            return LOAD(hnd=hnd)
        elif tc==OlxAPIConst.TC_LOADUNIT:
            return LOADUNIT(hnd=hnd)
        elif tc==OlxAPIConst.TC_SHUNT:
            return SHUNT(hnd=hnd)
        elif tc==OlxAPIConst.TC_SHUNTUNIT:
            return SHUNTUNIT(hnd=hnd)
        elif tc==OlxAPIConst.TC_SVD:
            return SVD(hnd=hnd)
        elif tc==OlxAPIConst.TC_BREAKER:
            return BREAKER(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYGROUP:
            return RLYGROUP(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYOCG:
            return RLYOCG(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYOCP:
            return RLYOCP(hnd=hnd)
        elif tc==OlxAPIConst.TC_FUSE:
            return FUSE(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYDSG:
            return RLYDSG(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYDSP:
            return RLYDSP(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYD:
            return RLYD(hnd=hnd)
        elif tc==OlxAPIConst.TC_RLYV:
            return RLYV(hnd=hnd)
        elif tc==OlxAPIConst.TC_RECLSR:
            return RECLSR(hnd=hnd)
        elif tc==OlxAPIConst.TC_BRANCH:
            return TERMINAL(hnd=hnd)
        else:
            raise Exception('error TC:'+str(tc))
    else:
        if type(sPara)!=list:
            if sPara not in __OLXOBJ_oHND__ and sPara!=None :
                return hnd
            #
            if type(hnd)!=list:
                if hnd<=0:
                    return None
                tc1 = OlxAPI.EquipmentType(hnd)
                if tc1==OlxAPIConst.TC_BRANCH:
                    hnd = __getDatai__(hnd,OlxAPIConst.BR_nHandle)
                    tc1 = OlxAPI.EquipmentType(hnd)
                return __getOBJ__(hnd,tc=tc1)
            #
            res = []
            for hnd1 in hnd:
                if hnd1>0:
                    tc1 = OlxAPI.EquipmentType(hnd1)
                    if tc1==OlxAPIConst.TC_BRANCH:
                        hnd1 = __getDatai__(hnd1,OlxAPIConst.BR_nHandle)
                        tc1 = OlxAPI.EquipmentType(hnd1)
                    res.append(__getOBJ__(hnd1,tc=tc1))
            return res
        else:
            res = []
            for i in range(len(sPara)):
                res.append(__getOBJ__(hnd[i],sPara=sPara[i]))
            return res
#
def __setValue__(hnd,paraCode,value):
    vt = paraCode//100
    if vt == OlxAPIConst.VT_STRING:
        return c_char_p( value.encode('UTF-8'))
    elif vt in [0,OlxAPIConst.VT_INTEGER]:
       return c_int(value)
    elif vt == OlxAPIConst.VT_DOUBLE:
       return c_double(value)
    elif vt in [OlxAPIConst.VT_ARRAYDOUBLE,OlxAPIConst.VT_ARRAYSTRING,OlxAPIConst.VT_ARRAYINT]:
##        tc = OlxAPI.EquipmentType(hnd)
##        if tc == TC_GENUNIT and paraCode in [ GU_vdR , GU_vdX]:
##            count = 5
##        elif tc == TC_LOADUNIT and paraCode in [OlxAPIConst.LU_vdMW,LU_vdMVAR]:
##            count = 3
##        elif tc == TC_LINE and paraCode == LN_vdRating:
##            count = 4
        #
        return (c_double*len(value))(*value)
    #
    raise OlxAPI.OlxAPIException("Error of paraCode")
#
def __getValue_i__(buf,count,stop=False):
    array = []
    val = cast(buf, POINTER(c_int*count)).contents
    for ii in range(count):
        array.append(val[ii])
        if stop and val[ii] ==0:
            break
    return array
#
def __getValue__(hnd,tokenV,buf):
    """ Convert GetData binary data buffer into Python object of correct type """
    vt = tokenV//100
    if vt == OlxAPIConst.VT_STRING:
        return OlxAPI.decode(buf.value)
    elif vt in [OlxAPIConst.VT_DOUBLE,OlxAPIConst.VT_INTEGER]:
        return buf.value
    #
    array = []
    tc = OlxAPI.EquipmentType(hnd)
    if tc == OlxAPIConst.TC_BREAKER and tokenV in {OlxAPIConst.BK_vnG1DevHnd,OlxAPIConst.BK_vnG2DevHnd,OlxAPIConst.BK_vnG1OutageHnd,OlxAPIConst.BK_vnG2OutageHnd}:
        count = OlxAPIConst.MXSBKF
        return __getValue_i__(buf,count,True)
    #
    if tc == OlxAPIConst.TC_SVD and tokenV == OlxAPIConst.SV_vnNoStep:
        count = 8
        return __getValue_i__(buf,count)
    #
    if tc in {OlxAPIConst.TC_RLYDSP,OlxAPIConst.TC_RLYDSG} and tokenV in {OlxAPIConst.DP_vParams,OlxAPIConst.DP_vParamLabels}:
        # String with tab delimited fields
        return ((cast(buf, c_char_p).value).decode("UTF-8")).split("\t")
    if tc == OlxAPIConst.TC_DCLINE2 and tokenV==OlxAPIConst.DC_vnBridges:
        count = 2
        return __getValue_i__(buf,count)
    #
    if tc == OlxAPIConst.TC_GENUNIT and tokenV in {OlxAPIConst.GU_vdR,OlxAPIConst.GU_vdX}:
        count = 5
    elif tc == OlxAPIConst.TC_LOADUNIT and tokenV in {OlxAPIConst.LU_vdMW,OlxAPIConst.LU_vdMVAR}:
        count = 3
    elif tc == OlxAPIConst.TC_SVD and tokenV in {OlxAPIConst.SV_vdBinc,OlxAPIConst.SV_vdB0inc}:
        count = 8
    elif tc == OlxAPIConst.TC_LINE and tokenV == OlxAPIConst.LN_vdRating:
        count = 4
    elif tc == OlxAPIConst.TC_RLYGROUP and tokenV == OlxAPIConst.RG_vdRecloseInt:
        count = 4
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV == OlxAPIConst.OG_vdDirSetting:
        count = 8
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV == OlxAPIConst.OG_vdDirSettingV15:
        count = 9
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV in {OlxAPIConst.OG_vdDTPickup,OlxAPIConst.OG_vdDTDelay}:
        count = 5
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV == OlxAPIConst.OP_vdDirSetting:
        count = 8
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV == OlxAPIConst.OP_vdDirSettingV15:
        count = 9
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV in {OlxAPIConst.OP_vdDTPickup,OlxAPIConst.OP_vdDTDelay}:
        count = 5
    elif tc == OlxAPIConst.TC_RLYDSG and tokenV == OlxAPIConst.DG_vdParams:
        count = OlxAPIConst.MXDSPARAMS
    elif tc == OlxAPIConst.TC_RLYDSG and tokenV in {OlxAPIConst.DG_vdDelay,OlxAPIConst.DG_vdReach,OlxAPIConst.DG_vdReach1}:
        count = OlxAPIConst.MXZONE
    elif tc == OlxAPIConst.TC_RLYDSP and tokenV == OlxAPIConst.DP_vParams:
        count = OlxAPIConst.MXDSPARAMS
    elif tc == OlxAPIConst.TC_RLYDSP and tokenV in {OlxAPIConst.DP_vdDelay,OlxAPIConst.DP_vdReach,OlxAPIConst.DP_vdReach1}:
        count = OlxAPIConst.MXZONE
    elif tc == OlxAPIConst.TC_CCGEN and tokenV in {OlxAPIConst.CC_vdV,OlxAPIConst.CC_vdI,OlxAPIConst.CC_vdAng}:
        count = OlxAPIConst.MAXCCV
    elif tc == OlxAPIConst.TC_BREAKER and tokenV in {OlxAPIConst.BK_vdRecloseInt1,OlxAPIConst.BK_vdRecloseInt2}:
        count = 3
    elif tc == OlxAPIConst.TC_MU and tokenV in [OlxAPIConst.MU_vdX,OlxAPIConst.MU_vdR,OlxAPIConst.MU_vdFrom1,OlxAPIConst.MU_vdFrom2,OlxAPIConst.MU_vdTo1,OlxAPIConst.MU_vdTo2]:
        count = 5
    elif tc == OlxAPIConst.TC_DCLINE2:# and tokenV in {OlxAPIConst.DC_vdAngleMax,OlxAPIConst.DC_vdAngleMin}:
        count = 2
    else:
        count = OlxAPIConst.MXDSPARAMS
    val = cast(buf, POINTER(c_double*count)).contents  # double array of size count
    for v in val:
        array.append(v)
    return array
#
def __findObjHnd__(val1,typ):
    if type(val1)==typ:
        __check_currFileIdx__(val1)
        return val1.__hnd__
    try:
        hnd = c_int(0)
        if OlxAPIConst.OLXAPI_OK==OlxAPI.FindObj1LPF(val1,hnd):
            return hnd.value
    except:
        pass
    return 0
#
def __findTerminalHnd__(b1,b2,sType,CID):
    if type(b1)!=BUS or type(b2)!=BUS or type(sType)!=str or sType not in __OLXOBJ_EQUIPMENT__ or type(CID)!=str:
        return 0
    #
    __check_currFileIdx__(b1)
    __check_currFileIdx__(b2)
    val1 = c_int(0)
    hnd1,hnd2 = b1.__hnd__,b2.__hnd__
    while OlxAPIConst.OLXAPI_OK==OlxAPI.GetBusEquipment(hnd1, c_int(OlxAPIConst.TC_BRANCH), byref(val1)):
        val2 = __getDatai__(val1.value,OlxAPIConst.BR_nBus2Hnd)
        val3 = __getDatai__(val1.value,OlxAPIConst.BR_nBus3Hnd)
        if val2==hnd2 or val3==hnd2:
            e1 = __getDatai__(val1.value,OlxAPIConst.BR_nHandle)
            tc1 = OlxAPI.EquipmentType(e1)
            o1 = __getOBJ__(e1,tc=tc1)
            if type(o1).__name__==sType:
                if o1.CID==CID:
                    return val1.value
    return 0
#
def __findBr2Hnd__(type1,val1,val2,val3):
    if type1 not in __OLXOBJ_EQUIPMENT__:
        return 0
    if type(val1)==BUS and type(val2)==BUS and type(val3)==str:
        b1,b2,sid = val1,val2,val3
    elif type(val1)==BUS and type(val3)==BUS and type(val2)==str:
        b1,b2,sid = val1,val3,val2
    elif type(val2)==BUS and type(val3)==BUS and type(val1)==str:
        b1,b2,sid = val2,val3,val1
    else:
        return 0
    #
    if b1.__hnd__<=0 or b2.__hnd__<=0:
        return 0
    #
    __check_currFileIdx__(b1)
    __check_currFileIdx__(b2)
    #
    bra = __getBusEquipment__(b1,type1)
    for br1 in bra:
        if b2.isInList(br1.BUS):
            if sid=='':
                return br1.__hnd__
            if br1.CID==sid:
                return br1.__hnd__
    return 0
#
def __findBr3Hnd__(type1,val1,val2,val3,val4):
    if type1!='XFMR3':
        return 0
    if type(val1)==BUS and type(val2)==BUS and type(val3)==BUS:
        b1,b2,b3 = val1,val2,val3
        sid = '' if val4==None else str(val4)
    elif type(val1)==BUS and type(val2)==BUS and type(val4)==BUS:
        b1,b2,b3 = val1,val2,val4
        sid = '' if val3==None else str(val3)
    elif type(val1)==BUS and type(val3)==BUS and type(val4)==BUS:
        b1,b2,b3 = val1,val3,val4
        sid = '' if val2==None else str(val2)
    elif type(val4)==BUS and type(val2)==BUS and type(val3)==BUS:
        b1,b2,b3 = val2,val3,val4
        sid = '' if val1==None else str(val1)
    else:
        return 0
    #
    if b1.__hnd__<=0 or b2.__hnd__<=0 or b3.__hnd__<=0:
        return 0
    #
    __check_currFileIdx__(b1)
    __check_currFileIdx__(b2)
    __check_currFileIdx__(b3)
    #
    bra = __getBusEquipment__(b1,type1)
    for br1 in bra:
        ba = br1.BUS
        if b2.isInList(ba) and b3.isInList(ba) :
            if sid=='':
                return br1.__hnd__
            if sid==br1.CID:
                return br1.__hnd__
    return 0
#
def __checkArgs__(args,sname):
    for a1 in args:
        if a1 is None:
            return
    if len(args)==2:
        se = '\n%s() takes 1 argument but 2 were given'%sname
    else:
        se = '\n%s() takes from 1 to %i arguments but %i were given'%(sname,len(args)-1,len(args))
    raise Exception(se)
#
def __checkArgsBr__(sf,val1,val2,val3):
    if not( type(val1)==str or (type(val1)==BUS and type(val2)==BUS and type(val3)==str)\
        or (type(val1)==BUS and type(val3)==BUS and type(val2)==str)\
        or (type(val2)==BUS and type(val3)==BUS and type(val1)==str)):
        se = '\nValueError OLCase.%s(...)'%sf
        se +='\n\tRequired: (b1,b2,CID) or (GUID) or (STR)'
        se +='\n\tFound   : ('+type(val1).__name__
        if val2 is not None:
            se+= ','+type(val2).__name__
        if val3 is not None:
            se+= ','+type(val3).__name__
        se +=')'
        raise ValueError(se)
#
def __checkArgs1__(sf,val1,val2):
    if not((type(val1)==str and val2 is None) or (type(val1)==str and (type(val2)==float or int))\
        or (type(val2)==str and (type(val1)==float or int))):
        se='\nValueError OLCase.%s(...)'%sf
        se +='\n\tRequired: (name,kV) or (GUID) or (STR)'
        se +='\n\tFound   : ('+type(val1).__name__
        if val2 is not None:
            se+= ','+type(val2).__name__
        se +=')'
        raise ValueError(se)
#
def __initFailCheck__(hnd,sType,vala):
    if hnd==0 or OlxAPI.EquipmentType(hnd)!=__OLXOBJ_CONST__[sType][0]:
        __check_currFileIdx1__()
        se = '\n'+ sType +'('
        for v1 in vala:
            if v1!=None:
                if type(v1)==str:
                    se +='"'+v1 +'" , '
                elif type(v1)==BUS:
                    if v1.__hnd__>0:
                        se += v1.toString()+' , '
                    else:
                        se += 'BUS NOT FOUND,'
                else:
                    se +=str(v1) +' , '
        #
        if se.endswith('('):
            se += ') missing required argument'
        else:
            if se.endswith(' , '):
                se = se[:-3]
            se += '): NOT FOUND'
        raise Exception(se)
#
def __initFail2__(sType,hnd):
    __check_currFileIdx1__()
    se = '\n'+ sType +'(hnd='+str(hnd) +'): NOT FOUND'
    raise Exception(se)
#
def __oldinit214__(so):
    """not support .init, internal usage"""
    se= '\nThe %s.init(..) can no longer be used by the OlxObj module version 2.1.4 or above.'%so
    se+='\nYou must use OLCase.find%s(...) instead.'%so
    raise Exception(se)
#
def __getTERMINAL_OBJ__(ob, sType):
    __check_currFileIdx__(ob)
    hnd = ob.__hnd__
    if sType=='BUS':
        res = []
        h1 = __getDatai__(hnd,OlxAPIConst.BR_nBus1Hnd)
        res.append(BUS(hnd=h1))
        #
        h2 = __getDatai__(hnd,OlxAPIConst.BR_nBus2Hnd)
        res.append (BUS(hnd=h2))
        h3 = __getDatai__(hnd,OlxAPIConst.BR_nBus3Hnd)
        if h3 is not None:
            res.append(BUS(hnd=h3))
        return res
    if sType=='EQUIPMENT':
        e1 = __getDatai__(hnd,OlxAPIConst.BR_nHandle)
        tc1 = OlxAPI.EquipmentType(e1)
        return __getOBJ__(e1,tc1)
    if sType=='RLYGROUP':
        res = []
        for c1 in [OlxAPIConst.BR_nRlyGrp1Hnd,OlxAPIConst.BR_nRlyGrp2Hnd,OlxAPIConst.BR_nRlyGrp3Hnd]:
            r1 = __getDatai__(hnd,c1)
            try:
                res.append( RLYGROUP(hnd=r1) )
            except:
                res.append(None)
        return res
    if sType=='FLAG':
        return __getDatai__(hnd,OlxAPIConst.BR_nInService)
    if sType=='OPPOSITE':
        res = []
        for t1 in ob.EQUIPMENT.TERMINAL:
            if t1.__hnd__!=hnd:
                res.append(t1)
        return res
    if sType=='REMOTE':
        """All taps are ignored.
             Close switches are included
             Out of service branches are ignored"""
        #import OlxAPILib
        #ba,_,_,_ = OlxAPILib.getRemoteTerminals(hnd,[])
        ba = __getRemoteTerminals__(hnd)
        res = []
        for b1 in ba:
            res.append(TERMINAL(hnd=b1))
        return res
#
def __getRemoteTerminals__(hnd0):
    """ Purpose: Find all remote end of a branch
             All taps are ignored.
             Close switches are included
             Out of service branches are ignored
             if branch located on tapbus => return [] } """
    st0 = __getDatai__(hnd0,OlxAPIConst.BR_nInService)
    if st0==2: # br out of service
        return []
    b1 = __getDatai__(hnd0,OlxAPIConst.BR_nBus1Hnd) # bus1 of br
    tap1 = __getDatai__(b1,OlxAPIConst.BUS_nTapBus)
    if tap1!=0: # br located on tapbus
        return []
    ty0 = __getDatai__(hnd0,OlxAPIConst.BR_nType)
    e0 = __getDatai__(hnd0,OlxAPIConst.BR_nHandle)
    if ty0==OlxAPIConst.TC_SWITCH:
        sw0 = __getDatai__(e0,OlxAPIConst.SW_nStatus)
        if sw0 ==0: # switch open
            return []
    #
    res = []
    setBus = {b1}
    setEqui = {e0}
    #
    h2 = __getDatai__(hnd0,OlxAPIConst.BR_nBus2Hnd)
    h3 = __getDatai__(hnd0,OlxAPIConst.BR_nBus3Hnd)
    b23 = {h2}
    if h3 is not None:
        b23.add(h3)
    #
    while True:
        b23in = set()
        setBus.update(b23)
        for b2 in b23:
            nTap = __getDatai__(b2,OlxAPIConst.BUS_nTapBus)
            if nTap==0:# finish
                bra = __getBusEquipmentHnd__(b2,'TERMINAL')
                for br1 in bra:
                    e1 = __getDatai__(br1,OlxAPIConst.BR_nHandle)
                    if e1 in setEqui:
                        res.append(br1)
            else: # continue
                bra = __getBusEquipmentHnd__(b2,'TERMINAL')
                for br1 in bra:
                    st1 = __getDatai__(br1,OlxAPIConst.BR_nInService)
                    if st1!=2: # in service branch
                        e1 = __getDatai__(br1,OlxAPIConst.BR_nHandle)
                        ty1 = __getDatai__(br1,OlxAPIConst.BR_nType)
                        stw = 7
                        if ty1==OlxAPIConst.TC_SWITCH:
                            stw = __getDatai__(e1,OlxAPIConst.SW_nStatus)
                        if stw >0 and e1 not in setEqui:
                            setEqui.add(e1)
                            h2i = __getDatai__(br1,OlxAPIConst.BR_nBus2Hnd)
                            h3i = __getDatai__(br1,OlxAPIConst.BR_nBus3Hnd)
                            if h2i not in setBus:
                                b23in.add(h2i)
                            if h3i is not None and h3i not in setBus:
                                b23in.add(h3i)
        #
        if len(b23in)==0:
            break
        b23.clear()
        b23.update(b23in)
    return res
#
def __get_OBJTERMINAL__(ob):
    """ return list TERMINAL of Object"""
    res = []
    for b1 in ob.BUS:
        for t1 in b1.TERMINAL:
            if t1.EQUIPMENT.__hnd__ == ob.__hnd__:
                res.append(t1)
    return res
#
def __getRLYGROUP_OBJ__(rg, sType):
    """
    Retrieves all object (RLY+BUS+SCHEME) that is attached to RLYGROUP
    - rg : RLYGROUP
    - sType : OLXOBJ_RELAY+BUS
    """
    __check_currFileIdx__(rg)
    val1,val2 = c_int(0),c_int(0)
    hnd = rg.__hnd__
    res = []
    if sType=='BUS':
        if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(hnd,c_int(OlxAPIConst.RG_nBranchHnd),byref(val1)):
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(val1.value,c_int(OlxAPIConst.BR_nBus1Hnd),byref(val2)):
                res.append(val2.value)
            else:
                raise Exception('\nBRANCH not found for :'+rg.toString())
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(val1.value,c_int(OlxAPIConst.BR_nBus2Hnd),byref(val2)):
                res.append(val2.value)
            if OlxAPIConst.OLXAPI_OK==OlxAPI.GetData(val1.value, c_int(OlxAPIConst.BR_nBus3Hnd),byref(val2)):
                res.append(val2.value)
        else:
            raise Exception('\nBRANCH not found for :'+rg.toString())
    else:
        while OlxAPIConst.OLXAPI_OK==OlxAPI.GetRelay(hnd,byref(val1)):
            tc1 = OlxAPI.EquipmentType(val1.value)
            if type(sType)in {list,set}:
                for sType1 in sType:
                    if tc1 == __OLXOBJ_CONST__[sType1][0]:
                        res.append(val1.value)
            else:
                if tc1 == __OLXOBJ_CONST__[sType][0]:
                    res.append(val1.value)
    return res
#
def __getRLYGROUP_RLY__(rg):
    """
    Retrieves all RELAY object that is attached to RLYGROUP
    - rg : RLYGROUP
    """
    val1 = c_int(0)
    __check_currFileIdx__(rg)
    hnd = rg.__hnd__
    res = []
    while OlxAPIConst.OLXAPI_OK==OlxAPI.GetRelay(hnd,byref(val1)):
        tc1 = OlxAPI.EquipmentType(val1.value)
        if tc1!=OlxAPIConst.TC_RECLSRG:
            r1 = __getOBJ__(val1.value,tc1)
            res.append(r1)
    return res
#
def __changeRLYSetting__(rl,sPara,value):
    __check_currFileIdx__(rl)
    if type(rl) not in {RLYOCG,RLYOCP,RLYDSG,RLYDSP}:
        raise AttributeError("changeSetting() method work only with relay type 'RLYOCG','RLYOCP','RLYDSG','RLYDSP'")
    #
    if type(sPara)!=str :
        se = "\nin call %s.changeSetting(sPara=%s,val=%s) "%(type(rl).__name__,str(sPara), str(value))
        se+= "\n\ttype(sPara): str"
        se+= "\n\tfound      : "+type(sPara).__name__
        raise TypeError(se)
    if type(rl) in {RLYOCG,RLYOCP}:
        __changeRLYSettingOC__(rl,sPara,value)
    else:
        __changeRLYSettingDS__(rl,sPara,value)
#
def __changeRLYSettingOC__(rl,sPara,value):
    #
    if sPara not in __OLXOBJ_RLYSET__[type(rl).__name__].keys():
        polar = __getRelaySetting__(rl,'POLAR')
        __errorParaSettingOC__(rl,polar,sPara,'changeSetting',value)
    #
    paraCode = __OLXOBJ_RLYSET__[type(rl).__name__][sPara]
    if len(paraCode) ==2:
        t0 = __getTypeParaCode__(paraCode[0])
        if (t0 != type(value)) and not (t0==float and type(value)==int):
            se = "\nin call %s.changeSetting('%s',value)"%(type(rl).__name__,sPara)
            se+= "\n\ttype(value) required : %s"%t0.__name__
            se+= "\n\tfound                : "+type(value).__name__
            raise TypeError(se)
        #
        val1 = __setValue__(rl.__hnd__,paraCode[0],value)
        if OlxAPIConst.OLXAPI_OK==OlxAPI.SetData(c_int(rl.__hnd__),c_int(paraCode[0]),byref(val1)):
            return
        se = "\nWrite Access = NO %s.changeSetting(sPara=%s,val=%s) "%(type(rl).__name__,str(sPara), str(value))
        raise Exception(se)
    else:
        polar = __getRelaySetting__(rl,'POLAR')
        if polar in paraCode[1]:
            if type(value).__name__ not in {'int','float'}:
                se = "\nin call %s.changeSetting('%s',value)"%(type(rl).__name__,sPara)
                se+= "\n\ttype(value) required : float"
                se+= "\n\tfound                : "+type(value).__name__
                raise TypeError(se)
            #
            id1 = paraCode[2]
            valo = [0]*8
            valo[id1] = value
            #
            val1 = __setValue__(rl.__hnd__,paraCode[0],valo)
            if OlxAPIConst.OLXAPI_OK==OlxAPI.SetData(c_int(rl.__hnd__),c_int(0x10000*(id1+1)+paraCode[0]), byref(val1)):
                return
            se = "\nWrite Access = NO %s.changeSetting(sPara=%s,val=%s) "%(type(rl).__name__,str(sPara), str(value))
            raise Exception(se)
        else:
            __errorParaSettingOC__(rl,polar,sPara,'changeSetting',value)
    return
#
def __changeRLYSettingDS__(rl,sPara,value):
    if sPara in __OLXOBJ_RLYSET__[type(rl).__name__].keys():
        paraCode = __OLXOBJ_RLYSET__[type(rl).__name__][sPara][0]
        t0 = __getTypeParaCode__(paraCode)
        #
        if (t0 != type(value)) and not (t0==float and type(value)==int):
            se = "\nin call %s.changeSetting('%s',value)"%(type(rl).__name__,sPara)
            se+= "\n\ttype(value) required : %s"%t0.__name__
            se+= "\n\tfound                : "+type(value).__name__
            raise TypeError(se)
        #
        val1 = __setValue__(rl.__hnd__,paraCode,value)
        if OlxAPIConst.OLXAPI_OK == OlxAPI.SetData( c_int(rl.__hnd__), c_int(paraCode), byref(val1) ):
            return
        se = "\nWrite Access = NO with %s.changeSetting('%s',%s)"%(type(rl).__name__,sPara,str(value))
        raise Exception(se)
    #
    if type(rl)==RLYDSG:
        nC = __getDatai__(rl.__hnd__,OlxAPIConst.DG_nParamCount)
        label = __getData__(rl.__hnd__,OlxAPIConst.DG_vParamLabels)
    else:
        nC = __getDatai__(rl.__hnd__,OlxAPIConst.DP_nParamCount)
        label = __getData__(rl.__hnd__,OlxAPIConst.DP_vParamLabels)
    #
    if sPara in label[:nC]:
        vs = sPara +'\t' +str(value)
        val = c_char_p(vs.encode('UTF-8'))
        if OlxAPIConst.OLXAPI_OK == OlxAPI.SetData(c_int(rl.__hnd__),c_int(OlxAPIConst.DG_sParam),byref(val)):
            return
        se = "\nWrite Access = NO with %s.changeSetting('%s',%s)"%(type(rl).__name__,sPara,str(value))
        raise Exception(se)
    # Error
    se = '\nAll Setting for %s : '%(type(rl).__name__) + rl.toString()
    for a1,val in __OLXOBJ_RLYSET__[type(rl).__name__].items():
        se+= '\n' + a1.ljust(15)+' : '+ val[1]
    for i in range(nC):
        se+= '\n' + label[i].ljust(15)
    #
    se += "\n\n%s.changeSetting('%s',%s)"%(type(rl).__name__,sPara,str(value))
    se+= "\nAttributeError  : '%s'" %(sPara)
    raise AttributeError (se)
#
def __printRLYSettingOC__(rl,cmt):
    polar = __getRelaySetting__(rl,'POLAR')
    sres = '\nSetting for : '+rl.toString()
    ak = list(__OLXOBJ_RLYSET__[type(rl).__name__].keys())
    ak.sort(reverse = True)
    for a1 in ak:
        val = __OLXOBJ_RLYSET__[type(rl).__name__][a1]
        if len(val)>2 :
            if polar in val[1]:
                v1 = __getRelaySetting__(rl,a1)
                sres+= '\n' + a1.ljust(15)+' : '+toString(v1).ljust(15)+'\t'+ val[3]
        else:
            v1 = __getRelaySetting__(rl,a1)
            sres+= '\n' + a1.ljust(15)+' : '+toString(v1).ljust(15)+'\t'+ val[1]
    return sres
#
def __printRLYSettingDS__(rl,cmt):
    sres = '\nSetting for : '+rl.toString()
    va = __getRelaySetting__(rl,'')
    ka = list(va.keys())
    ka.sort()
    for k in ka:
        s1= k.ljust(15)+' : '+toString(va[k]).ljust(15)
        if k in __OLXOBJ_RLYSET__[type(rl).__name__].keys():
            if len(s1)<37:
                s1+='  '+__OLXOBJ_RLYSET__[type(rl).__name__][k][1]
            else:
                s1+='\n                                   '+__OLXOBJ_RLYSET__[type(rl).__name__][k][1]
        sres+='\n'+s1
    return sres
#
def __printRLYSetting__(rl,cmt):
    if type(rl) in {RLYOCG,RLYOCP}:
        return __printRLYSettingOC__(rl,cmt)
    elif type(rl) in {RLYDSG,RLYDSP}:
        return __printRLYSettingDS__(rl,cmt)
#
def __getRelaySetting__(rl,sPara):
    __check_currFileIdx__(rl)
    if type(sPara) in {list,set}:
        res = []
        for sp1 in sPara:
            res.append(__getRelaySetting__(rl,sp1))
        return res
    # --------------
    if type(rl) not in {RLYOCG,RLYOCP,RLYDSG,RLYDSP}:
        raise AttributeError("getSetting() method work only with relay type RLYOCG,RLYOCP,RLYDSG,RLYDSP")
    #
    if type(rl) in {RLYOCG,RLYOCP}:
        return __getRelaySettingOC__(rl,sPara)
    return __getRelaySettingDS__(rl,sPara)
#
def __getRelaySettingName__(rl):
    __check_currFileIdx__(rl)
    if type(rl) not in {RLYDSG,RLYDSP}:#'RLYOCG','RLYOCP',
        raise AttributeError("getDSSettingName() method work only with relay type RLYDSG,RLYDSP")
    #
    if type(rl)==RLYDSG:
        nC = __getDatai__(rl.__hnd__,OlxAPIConst.DG_nParamCount)
        label = __getData__(rl.__hnd__,OlxAPIConst.DG_vParamLabels)
    else:
        nC = __getDatai__(rl.__hnd__,OlxAPIConst.DP_nParamCount)
        label = __getData__(rl.__hnd__,OlxAPIConst.DP_vParamLabels)
    return label[:nC]
#
def __getRelaySettingOC__(rl,sPara):# {'RLYOCG','RLYOCP'}
    #
    if sPara==None or sPara =='':
        sPara = __OLXOBJ_RLYSET__[type(rl).__name__].keys()
        res = dict()
        for sp1 in sPara:
            try:
                res[sp1] = __getRelaySetting__(rl,sp1)
            except:
                pass
        return res
    #
    if sPara not in __OLXOBJ_RLYSET__[type(rl).__name__].keys():
        polar = __getRelaySetting__(rl,'POLAR')
        __errorParaSettingOC__(rl,polar,sPara)
    #------------------
    cc = __OLXOBJ_RLYSET__[type(rl).__name__][sPara]
    paraCode = cc[0]
    if len(cc)==2:
        return __getData__(rl.__hnd__,paraCode)
    #
    polar = __getRelaySetting__(rl,'POLAR')
    #
    if polar not in cc[1]:
        __errorParaSettingOC__(rl,polar,sPara)
    #
    r = __getData__(rl.__hnd__,paraCode)
    k = cc[2]
    return r[k]
#
def __errorParaSettingOC__(rl,polar,sPara,sc ='',value =None):
    se = '\nAll Setting (POLAR=%i) for %s : '%(polar,type(rl).__name__) + rl.toString()
    for a1,val in __OLXOBJ_RLYSET__[type(rl).__name__].items():
        if len(val)>2 :
            if polar in val[1]:
                se+= '\n' + a1.ljust(15)+' : '+ val[3]
        else:
            se+= '\n' + a1.ljust(15)+' : '+ val[1]
    if sc:
        se += "\n\n%s.changeSetting('%s',%s)"%(type(rl).__name__,sPara,str(value))
    else:
        se += "\n\n%s.getSetting('%s')"%(type(rl).__name__,sPara)
    se+= "\nAttributeError : '%s' (POLAR=%i)" %(sPara,polar)
    raise AttributeError (se)
#
def __getRelaySettingDS__(rl,sPara):    # {'RLYDSG','RLYDSP'}
    hnd1 = rl.__hnd__
    if sPara==None or sPara =='':
        sPara1 = list(__OLXOBJ_RLYSET__[type(rl).__name__].keys())
        if type(rl)==RLYDSG:
            nC = __getDatai__(hnd1,OlxAPIConst.DG_nParamCount)
            label = __getData__(hnd1,OlxAPIConst.DG_vParamLabels)
        else:
            nC = __getDatai__(hnd1,OlxAPIConst.DP_nParamCount)
            label = __getData__(hnd1,OlxAPIConst.DP_vParamLabels)
        sPara1.extend(label[:nC])
        #
        res = dict()
        for sp1 in sPara1:
            try:
                res[sp1] = __getRelaySetting__(rl,sp1)
            except:
                pass
        return res
    #
    try:
        paraCode = __OLXOBJ_RLYSET__[type(rl).__name__][sPara]
    except:
        paraCode = None
    #
    if paraCode is not None:
        res = __getData__(hnd1,paraCode[0])
        return __getOBJ__(res,sPara=sPara)
    #
    sPara1 = sPara.upper()
    if type(rl)==RLYDSG:
        nC = __getDatai__(hnd1,OlxAPIConst.DG_nParamCount)
        label = __getData__(hnd1,OlxAPIConst.DG_vParamLabels)
        valF = __getData__(hnd1,OlxAPIConst.DG_vdParams)
    else:
        nC = __getDatai__(hnd1,OlxAPIConst.DP_nParamCount)
        label = __getData__(hnd1,OlxAPIConst.DP_vParamLabels)
        valF = __getData__(hnd1,OlxAPIConst.DP_vdParams)
    #
    for i in range(nC):
        if sPara1 ==label[i].upper():
            return valF[i]
    # Error
    se = '\nAll Setting for %s : '%(type(rl).__name__) + rl.toString()
    for a1,val in __OLXOBJ_RLYSET__[type(rl).__name__].items():
        se+= '\n' + a1.ljust(15)+' : '+ val[1]
    for i in range(nC):
        se+= '\n' + label[i].ljust(15)
    se += "\n\n%s.getSetting('%s')"%(type(rl).__name__,sPara)
    se+= "\nAttributeError : '%s'" %(sPara)
    raise AttributeError (se)
#
def __check_currFileIdx1__():
    if __CURRENT_FILE_IDX__==0:
        raise Exception ('\nASPEN OLR file is not yet opened')
    if __CURRENT_FILE_IDX__<0:
        raise Exception ('\nASPEN OLR file is closed')
#
def __check_currFileIdx__(ob):
    if __CURRENT_FILE_IDX__==0:
        raise Exception ('\nASPEN OLR file is not yet opened')
    if __CURRENT_FILE_IDX__!=ob.__currFileIdx__:
        raise Exception('\nASPEN OLR related to Object (%s) is closed'%type(ob).__name__)
#
def __isInScope__(ob,scope):
    """
    test if this object is in scope:
    """
    if scope is None or scope['isFullNetWork']:
        return True
    #
    typ = type(ob)
    if typ==BUS:
        return __busIsInScope__(ob,scope['areaNum'],scope['zoneNum'],scope['kV'])
    elif typ==XFMR:
        b1,b2 = ob.BUS1,ob.BUS2
        kv1,kv2 = b1.KV,b2.KV
        if kv1==kv2:
            return __brIsInScope__([b1,b2],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],scope['kV']])
        if kv1>kv2:
            return __brIsInScope__([b1,b2],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],[]])
        return __brIsInScope__([b1,b2],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[[],scope['kV']])
    elif typ==XFMR3:
        b1,b2,b3 = ob.BUS1,ob.BUS2,ob.BUS3
        kv1,kv2,kv3 = b1.KV,b2.KV,b3.KV
        if kv1>=kv2 and kv1>=kv3:
            return __brIsInScope__([b1,b2,b3],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],[],[]])
        elif kv2>=kv1 and kv2>=kv3:
            return __brIsInScope__([b1,b2,b3],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[[],scope['kV'],[]])
        return __brIsInScope__([b1,b2,b3],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[[],[],scope['kV']])
    elif typ==MULINE:
        return __brIsInScope__([ob.LINE1.BUS1,ob.LINE1.BUS2],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],scope['kV']]) \
            or __brIsInScope__([ob.LINE2.BUS1,ob.LINE2.BUS2],scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],scope['kV']])
    elif typ in __OLXOBJ_EQUIPMENTO__:
        return __brIsInScope__(ob.BUS,scope['areaNum'],scope['zoneNum'],scope['optionTie'],[scope['kV'],scope['kV']])
    elif typ not in __OLXOBJ_EQUIPMENTO__ and typ!=MULINE:
        return __busIsInScope__(ob.BUS,scope['areaNum'],scope['zoneNum'],scope['kV'])
    raise Exception ('Error Type not in : '+str(__OLXOBJ_LIST__))
#
def __busIsInScope__(b1,areaNum,zoneNum,kV):
    """
    test if bus is in scope:
        areaNum []:  List of area Number
        zoneNum []:  List of zone Number
        kV      []: [kVmin,kVmax]
    """
    if areaNum:
        if b1.AREANO not in areaNum:
            return False
    if zoneNum:
        if b1.ZONENO not in zoneNum:
            return False
    if kV:
        kv1 = b1.KV
        if kv1<kV[0] or kv1>kV[1]:
            return False
    return True
#
def __printArray__(ar):
    s1= ''
    for a1 in ar:
        s1+= str(a1)+' '
    print(s1)
#
def __getTypeParaCode__(paraCode):
    vt = paraCode//100
    if vt == OlxAPIConst.VT_STRING:
        return str
    elif vt == OlxAPIConst.VT_INTEGER:
        return int
    elif vt == OlxAPIConst.VT_DOUBLE:
        return float
    return list
#
def __currentFault__(obj,style):
    #
    if obj is None:
        hnd = OlxAPIConst.HND_SC
    else:
        hnd = obj.__hnd__
    #
    vd12Mag = (c_double*12)(0.0)
    vd12Ang = (c_double*12)(0.0)
    #
    if OlxAPIConst.OLXAPI_FAILURE == OlxAPI.GetSCCurrent( hnd, vd12Mag, vd12Ang, c_int(style)):
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    #
    if obj is None or type(obj)==TERMINAL:
        return __resultComplex__(vd12Mag[:3],vd12Ang[:3])
    if type(obj)==XFMR3:
        return __resultComplex__(vd12Mag[:12],vd12Ang[:12])
    if type(obj)==XFMR:
        return __resultComplex__(vd12Mag[:8],vd12Ang[:8])
    if type(obj) in {SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
        return __resultComplex__(vd12Mag[:6],vd12Ang[:6])
    return __resultComplex__(vd12Mag[:4],vd12Ang[:4])
#
def __voltageFault__(obj,style):
    vd9Mag  = (c_double*9)(0.0)
    vd9Ang  = (c_double*9)(0.0)
    if ( OlxAPIConst.OLXAPI_FAILURE == OlxAPI.GetSCVoltage(obj.__hnd__, vd9Mag, vd9Ang, c_int(style)) ) :
        raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
    #
    if type(obj) in {XFMR, SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH}:
        return __resultComplex__(vd9Mag[:6],vd9Ang[:6])
    if type(obj)==XFMR3:
        return __resultComplex__(vd9Mag[:9],vd9Ang[:9])
    return __resultComplex__(vd9Mag[:3],vd9Ang[:3])
#
def __resultComplex__(v1,v2):
    res = []
    for i in range(len(v1)):
        res.append(complex(v1[i],v2[i]))
    return res
#
def __brIsInScope__(ba,areaNum,zoneNum,optionTie,kVa):
    if optionTie == 0: # 0-strictly in areaNum/zoneNum
        for i in range(len(ba)):
            if not __busIsInScope__(ba[i],areaNum,zoneNum,kVa[i]):
                return False
        return True
    elif optionTie == 1: # 1- with tie
        for i in range(len(ba)):
            if __busIsInScope__(ba[i],areaNum,zoneNum,kVa[i]):
                return True
        return False
    elif optionTie == 2: # 2- only tie
        haveT,haveF = False,False
        for i in range(len(ba)):
            if __busIsInScope__(ba[i],areaNum,zoneNum,kVa[i]):
                haveT = True
            else:
                haveF = True
        return haveT and haveF
    else:
        raise Exception ('Error value of optionTie = [0,1,2]')
#
def __setValType__(vt,value):
    if vt == OlxAPIConst.VT_STRING:
        if value==0:
            return create_string_buffer(b'\000' * 128)
        else:
            return c_char_p(value.encode('UTF-8'))
    elif vt == OlxAPIConst.VT_DOUBLE:
       return c_double(value)
    elif vt == OlxAPIConst.VT_INTEGER or vt==0:
       return c_int(value)
    elif vt in {OlxAPIConst.VT_ARRAYDOUBLE,OlxAPIConst.VT_ARRAYSTRING,OlxAPIConst.VT_ARRAYINT}:
        return create_string_buffer(b'\000' * 10*1024)
    #
    raise OlxAPI.OlxAPIException("Error of paraCode")
#
def __getSobjLst__(obj):
    """Internal usage"""
    if type(obj)!=list:
        s1=type(obj).__name__
    else:
        s1='['
        for o1 in obj:
            s1+=type(o1).__name__+','
        if len(s1)>1:
            s1=s1[:-1]
        s1+=']'
    return s1
#
def __getErrValue__(st,obj):
    """Internal usage"""
    if type(obj) == st:
        if st==str:
            return "Found (ValueError) : '%s'"%obj
        if st==list:
            return 'Found (ValueError) : '+__getSobjLst__(obj) +' len=%i'%len(obj)
        return "Found (ValueError) : "+str(obj)
    else:
        se= 'Found (ValueError) : (%s) '%type(obj).__name__
        if type(obj)==list:
            se+='['
            for o1 in obj:
                se+= type(o1).__name__ +','
            if len(obj)>0:
                se =se[:-1]
            se+=']'
        else:
            if type(obj)==str:
                se += "'"+str(obj)+"'"
            else:
                se += str(obj)
        return se
#
def __getClassical__(sp):
    obj = sp.__paraInput__['obj']
    fltApp = (sp.__paraInput__['fltApp']).upper().replace(' ','')
    fltConn = (sp.__paraInput__['fltConn']).upper().replace(' ','')
    Z = sp.__paraInput__['Z']
    outage = sp.__paraInput__['outage']
    #
    para1 =  dict()
    if Z is None:
        Z = [0,0]
    #
    __check_currFileIdx__(obj)
    para1['hnd'] = obj.__hnd__
    para1['R'] = c_double(Z[0])
    para1['X'] = c_double(Z[1])
    #
    para1['fltConn'] = (c_int*4)(0)
    v1 = __OLXOBJ_fltConnCLS__[fltConn]
    para1['fltConn'][v1[0]] = v1[1]
    #
    para1['outageLst'] = (c_int*100)(0)
    para1['outageOpt'] = (c_int*4)(0)
    fltOpt = (c_double*15)(0)
    #
    flagOutage = False
    if outage is not None:
        option = outage.option.upper()
        outage.__check1__()
        la = outage.outageLst
        for i in range(len(la)):
            __check_currFileIdx__(la[i])
            para1['outageLst'][i]= la[i].__hnd__
            flagOutage = True
        if flagOutage:
            if option in {'SINGLE','SINGLE-GND'}:
                para1['outageOpt'][0] = 1
            elif option=='DOUBLE':
                para1['outageOpt'][1] = 1
            elif option=='ALL':
                para1['outageOpt'][2] = 1
            else:
                raise Exception('error value of OUTAGE.option')
        if option=='SINGLE-GND':
            fltOpt[14] = outage.G  # @AS
    #
    dfltApp = {'BUS':0,'CLOSE-IN':0,'CLOSE-IN-WEO':2,'REMOTE-BUS':4,'LINE-END':6}
    if fltApp in dfltApp.keys():
        vs = dfltApp[fltApp]
        if flagOutage:
            fltOpt[vs+1] = 1
        else:
            fltOpt[vs] = 1
    else:
        percent = float(fltApp[:fltApp.index('%')])
        if 'WEO' in fltApp:
            if flagOutage:
                fltOpt[11] = percent
            else:
                fltOpt[10] = percent
        else:
            if flagOutage:
                fltOpt[9] = percent
            else:
                fltOpt[8] = percent
    para1['fltOpt'] = fltOpt
    #
    return para1
#
def __getSimultaneous__(sp):
    obj = sp.__paraInput__['obj']
    fltApp = sp.__paraInput__['fltApp'].upper()
    fltConn = sp.__paraInput__['fltConn'].upper()
    Z = sp.__paraInput__['Z']
    #
    Z1 = [0]*8
    try:
        for i in range(len(Z)):
            Z1[i] = Z[i]
    except:
        pass
    #
    vBRTYPE = {LINE:1,SERIESRC:1,XFMR:2,SHIFTER:3,SWITCH:7,XFMR3:10}
    #
    para1 = dict()
    para1['FRA'] = str(Z1[0])
    para1['FXA'] = str(Z1[1])
    para1['FRB'] = str(Z1[2])
    para1['FXB'] = str(Z1[3])
    para1['FRC'] = str(Z1[4])
    para1['FXC'] = str(Z1[5])
    para1['FRG'] = str(Z1[6])
    para1['FXG'] = str(Z1[7])
    para1['FPARAM'] = '0'
    para1['FLTCONN'] = '0'
    para1['PHASETYPE'] = '0'
    #
    dfltApp = {'BUS':0,'CLOSE-IN':1,'BUS2BUS':2,'LINE-END':3,'OUTAGE':5,'1P-OPEN':6,'2P-OPEN':7,'3P-OPEN':8}
    if fltApp in dfltApp.keys():
        para1['FLTAPPL'] = str(dfltApp[fltApp])
    else:
        para1['FLTAPPL'] = '4'
        para1['FPARAM'] = fltApp[:fltApp.index('%')]
    #
    if fltApp=='1P-OPEN':
        dfltConn = {'A':0,'B':1,'C':2}
        para1['PHASETYPE'] = str(dfltConn[fltConn])
    elif fltApp=='2P-OPEN':
        dfltConn = {'AB':0,'AB':0,'BC':1,'CB':1,'AC':2,'CA':2}
        para1['PHASETYPE'] = str(dfltConn[fltConn])
    elif fltApp in {'BUS','CLOSE-IN','LINE-END'}:
        v1 = __OLXOBJ_fltConnSIM__[fltConn]
        para1['FLTCONN'] = str(v1[0])
        para1['PHASETYPE'] = str(v1[1])
    elif fltApp=='BUS2BUS':
        dfltConn = {'AA':[0,0],'BB':[0,3],'CC':[0,5],'AB':[0,1],'AC':[0,2],'BC':[0,4],'BA':[0,1],'CA':[0,2],'CB':[0,4]}
        para1['PHASETYPE'] = str(dfltConn[fltConn][1])
    #
    if fltApp=='BUS':
        para1['BUS1NAME'] = obj.NAME
        para1['BUS1KV'] = str(obj.KV)
    elif fltApp in {'OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN'}:
        if type(obj) in {TERMINAL,RLYGROUP}:
            obj = obj.EQUIPMENT
        bus1,bus2 = obj.BUS1,obj.BUS2
        para1['BUS1NAME'] = bus1.NAME
        para1['BUS1KV'] = str(bus1.KV)
        para1['BUS2NAME'] = bus2.NAME
        para1['BUS2KV'] = str(bus2.KV)
        para1['CKTID'] = str(obj.CID)
        #
        para1['BRTYPE'] = str(vBRTYPE[type(obj)])
    elif fltApp=='BUS2BUS':
        if fltConn in {'BA','CA','CB'}:
            b1,b2 = obj[1],obj[0]
        else:
            b1,b2 = obj[0],obj[1]
        para1['BUS1NAME'] = b1.NAME
        para1['BUS1KV'] = str(b1.KV)
        para1['BUS2NAME'] = b2.NAME
        para1['BUS2KV'] = str(b2.KV)
    elif fltApp in {'CLOSE-IN','LINE-END'} or '%' in fltApp:
        ba = obj.BUS
        e1 = obj.EQUIPMENT
        #
        para1['BUS1NAME'] = ba[0].NAME
        para1['BUS1KV'] = str(ba[0].KV)
        para1['BUS2NAME'] = ba[1].NAME
        para1['BUS2KV'] = str(ba[1].KV)
        para1['CKTID'] = str(e1.CID)
        para1['BRTYPE'] = str(vBRTYPE[type(e1)])
    #
    return para1
#
def __getSEA__(sp):
    # obj,fltApp,fltConn,deviceOpt,tiers,Z
    obj = sp.__paraInput__['obj']
    fltApp = sp.__paraInput__['fltApp'].upper()
    fltConn = sp.__paraInput__['fltConn'].upper()
    deviceOpt = sp.__paraInput__['deviceOpt']
    tiers = sp.__paraInput__['tiers']
    Z = sp.__paraInput__['Z']
    #
    para1 = dict()
    para1['hnd'] = obj.__hnd__
    #
    para1['fltOpt'] = (c_double*64)(0)
    try:
        para1['fltOpt'][1] = float(fltApp[:fltApp.index('%')])
    except:
        pass
    #
    para1['fltOpt'][0] = __OLXOBJ_fltConnSEA__[fltConn]
    para1['fltOpt'][2] = Z[0]
    para1['fltOpt'][3] = Z[1]
    #
    para1['runOpt'] = (c_int*7)(0)
    for i in range(7):
        para1['runOpt'][i]=deviceOpt[i]
    #
    para1['tiers'] = tiers
    return para1
#
def __getSEA_EX__(sp):
    # time,fltConn,Z
    time = sp.__paraInput__['time']
    fltConn = sp.__paraInput__['fltConn'].upper()
    Z = sp.__paraInput__['Z']
    #
    para1 = dict()
    para1['time'] = time
    para1['Z'] = Z
    #
    para1['fltOpt'] = __OLXOBJ_fltConnSEA__[fltConn]
    return para1
#
def __checkClassical__(sp):
    obj = sp.__paraInput__['obj']
    fltApp = sp.__paraInput__['fltApp']
    fltConn = sp.__paraInput__['fltConn']
    Z = sp.__paraInput__['Z']
    outage = sp.__paraInput__['outage']
    #
    se = '\nSPEC_FLT.Classical(obj,fltApp,fltConn,Z,outage)'
    # check obj
    if type(obj) not in {BUS,RLYGROUP,TERMINAL}:
        se+= '\nobj : Classical fault object'
        se+= '\n\tRequired          : BUS or RLYGROUP or TERMINAL'
        se+= '\n\tFound (ValueError): '+ __getSobjLst__(obj)
        raise ValueError(se)
    # check runOpt: (str) simulation option code
    vrunOpt = ['BUS','CLOSE-IN','CLOSE-IN-EO','REMOTE-BUS','LINE-END']#,'xx%','xx%-EO'
    if type(fltApp)==str:
        fltApp = fltApp.upper().replace(' ','')
        flag = fltApp in vrunOpt
        if not flag:
            try:
                percent = float(fltApp[:fltApp.index('%')])
                flag = percent>=0 and percent<=100
            except:
                flag = False
            if not(fltApp.endswith('%') or fltApp.endswith('%-WEO')):
                flag = False
    else:
        flag = False
    if not flag:
        se+= '\nfltApp : (str) Fault application'
        vrunOpt.extend(['xx%','xx%-EO'])
        se+= '\n\tRequired      (str): ' +str(vrunOpt) + ' (0<=xx<=100)'
        se+= '\n\t' +__getErrValue__(str,fltApp)
        raise ValueError(se)
    if (type(obj)==BUS and fltApp !='BUS') or (type(obj)!=BUS and fltApp =='BUS'):
        se+= '\n\tobj : Classical fault object'
        se+= '\n\tfltApp : (str) Fault application'
        se+= "\n\tobj=%s, fltApp=%s"%(type(obj).__name__,fltApp)
        raise ValueError(se)
    if fltApp not in vrunOpt:
        if type(obj.EQUIPMENT)!=LINE:
            se+= '\n\tobj : Classical fault object'
            se+= '\n\tfltApp : (str) Fault application'
            se+= "\n\tobj(EQUIPMENT)=%s, fltApp=%s"%(type(obj.EQUIPMENT).__name__,fltApp)
            raise ValueError(se)
    #check fltConn
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        se+= '\nfltConn : Fault connection code'
        se+= "\n\tRequired     (str) :" + str(__OLXOBJ_fltConn__[:7])[:-1]
        se+= ',\n\t                     ' + str(__OLXOBJ_fltConn__[7:])[1:]
        se+= '\n\t' +__getErrValue__(str,fltConn)
        raise ValueError(se)
    #check Z
    if not (Z is None or (type(Z)==list and len(Z)==2 and type(Z[0]) in {float,int} and type(Z[1]) in {float,int})):
        se+= '\n\tZ : [R,X] Fault impedances in Ohm'
        se+= '\n\tRequired           : [float,float]'
        se+= '\n\t' +__getErrValue__(list,Z)
        raise ValueError(se)
    #
    if not(outage is None or type(outage)==OUTAGE):
        se+= '\n\toutage : Outage option'
        se+= '\n\tRequired           : None or (OUTAGE)'
        se+= '\n\t' +__getErrValue__(OUTAGE,outage)
        raise ValueError(se)
#
def __checkSimultaneous__(sp):
    obj = sp.__paraInput__['obj']
    fltApp = sp.__paraInput__['fltApp']
    fltConn = sp.__paraInput__['fltConn']
    Z = sp.__paraInput__['Z']
    #
    se = '\nSPEC_FLT.Simultaneous(obj,fltApp,fltConn,Z)'
    dfltApp  =['BUS','CLOSE-IN','BUS2BUS','LINE-END','OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN']
    dfltApp1 =['BUS','CLOSE-IN','BUS2BUS','LINE-END','OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN','xx%']
    flag = type(fltApp)==str
    if flag:
        fltApp1 = fltApp.upper()
        flag = fltApp1 in dfltApp
        if not flag:
            try:
                percent = float(fltApp1[:fltApp1.index('%')])
                flag = percent>=0 and percent<=100
            except:
                flag = False
            if not fltApp1.endswith('%'):
                flag = False
    if not flag:
        se+= '\n  fltApp : (str) Fault application code'
        se+= '\n\tRequired : (str) in '+str(dfltApp1) +' (0<=xx<=100)'
        se+= '\n\t' +__getErrValue__(str,fltApp)
        raise ValueError(se)
    #
    se += ", with fltApp='%s'"%fltApp
    # fltApp = Bus ---------------------------
    if fltApp1=='BUS':
        if type(obj)!=BUS:
            se+= '\n  obj : Simultaneous fault object'
            se+= '\n\tRequired           : (BUS)'
            se+= '\n\t'+__getErrValue__(BUS,obj)
            raise ValueError(se)
        __check_currFileIdx__(obj)
    # fltApp = Close-In,Line-End ---------------------------
    if fltApp1 in {'CLOSE-IN','LINE-END'}:
        if type(obj) not in {TERMINAL,RLYGROUP}:
            se+= '\n  obj : Simultaneous fault object'
            se+= '\n\tRequired          : TERMINAL or RLYGROUP'
            se+= '\n\t'+__getErrValue__(TERMINAL,obj)
            raise ValueError(se)
        __check_currFileIdx__(obj)
    # fltApp = xx% ---------------------------
    if fltApp1.find('%')>0:
        if type(obj) not in {TERMINAL,RLYGROUP} or type(obj.EQUIPMENT)!=LINE:#
            se+= '\n  obj : Simultaneous fault object'
            se+= '\n\tRequired          : TERMINAL or RLYGROUP (on a LINE)'
            if type(obj) not in {TERMINAL,RLYGROUP}:
                se+= '\n\t'+__getErrValue__(TERMINAL,obj)
            else:
                se+= '\n\tFound (ValueError) :TERMINAL or RLYGROUP (on a '+type(obj.EQUIPMENT).__name__
            raise ValueError(se)
        __check_currFileIdx__(obj)
    #
    if fltApp1 in {'BUS','CLOSE-IN','LINE-END'} or fltApp1.find('%')>0:
        if type(fltConn)!=str or fltConn.upper() not in __OLXOBJ_fltConn__:
            se+= '\n  fltConn  : (str) Fault connection code'
            se+= '\n\tRequired : (str) in '+str(__OLXOBJ_fltConn__[:7])[:-1]
            se+= ',\n\t                     ' +str(__OLXOBJ_fltConn__[7:])[1:]
            se+= '\n\t' +__getErrValue__(str,fltConn)
            raise ValueError(se)
        #
        if Z is not None:
            flag = type(Z)==list and len(Z)==8
            if flag:
                for z1 in Z:
                    if type(z1) not in {float,int}:
                        flag =False
                        break
            if not flag:
                se+= '\n  Z : Fault impedances'
                se+= '\n\tRequired [float]*8 : [RA,XA,RB,XB,RC,XC,RG,XG]'
                se+= '\n\t' +__getErrValue__(list,Z)
                raise ValueError(se)
    # fltApp1 = BUS2BUS ---------------------------
    if fltApp1=='BUS2BUS':
        if type(obj)!=list or len(obj)!=2 or type(obj[0])!=BUS or type(obj[1])!=BUS:
            se+= '\n  obj : Simultaneous fault object'
            se+= '\n\tRequired           : [BUS,BUS]'
            se+= '\n\t' +__getErrValue__(list,obj)
            raise ValueError(se)
        #
        __check_currFileIdx__(obj[0])
        __check_currFileIdx__(obj[1])
        vd = ['AA','BB','CC','AB','AC','BC','BA','CA','CB']
        if type(fltConn)!=str or fltConn.upper() not in vd:
            se+= '\n  fltConn  : (str) Fault connection code'
            se+= '\n\tRequired : (str) in '+str(vd)
            se+= '\n\t' +__getErrValue__(str,fltConn)
            raise ValueError(se)
        #
        if Z is not None:
            if type(Z)!=list or len(Z)!=2 or type(Z[0]) not in {float,int} or type(Z[1]) not in {float,int}:
                se+= '\n  Z : Fault impedances'
                se+= '\n\tRequired [float]*2 : [R,X]'
                se+= '\n\t' +__getErrValue__(list,Z)
                raise ValueError(se)
    # fltApp ='OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN'----------------
    if fltApp1 in {'OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN'}:
        if type(obj) not in {TERMINAL,RLYGROUP}:
            se+= '\n  obj : Simultaneous fault object'
            se+= '\n\tRequired          : TERMINAL or RLYGROUP'
            se+= '\n\t'+__getErrValue__(TERMINAL,obj)
            raise ValueError(se)
        __check_currFileIdx__(obj)
        if fltApp1=='1P-OPEN':
            vd = ['A','B','C']
            if type(fltConn)!=str or fltConn.upper() not in vd:
                se+= '\n  fltConn  : Fault connection code'
                se+= '\n\tRequired : (str) in '+str(vd)
                se+= '\n\t' +__getErrValue__(str,fltConn)
                raise ValueError(se)
        if fltApp1=='2P-OPEN':
            vd = ['AB','AC','BC','BA','CA','CB']
            if type(fltConn)!=str or fltConn.upper() not in vd:
                se += '\n  fltConn  : Fault connection code'
                se += '\n\tRequired : (str) in '+str(vd)
                se += '\n\t' +__getErrValue__(str,fltConn)
                raise ValueError(se)
#
def __checkSEA__(sp):
    obj = sp.__paraInput__['obj']
    fltApp = sp.__paraInput__['fltApp']
    fltConn = sp.__paraInput__['fltConn']
    deviceOpt = sp.__paraInput__['deviceOpt']
    tiers = sp.__paraInput__['tiers']
    Z = sp.__paraInput__['Z']
    #
    se = '\nSPEC_FLT.SEA(obj,fltApp,fltConn,deviceOpt,tiers,Z)'
    if type(obj) not in {BUS,RLYGROUP,TERMINAL}:
        se+= '\nobj : Stepped-Event Analysis object'
        se+= '\n\tRequired          : BUS or RLYGROUP or TERMINAL'
        se += '\n\t' +__getErrValue__(TERMINAL,obj)
        raise ValueError(se)
    if type(obj)==BUS:
        if type(fltApp)!=str or fltApp.upper()!='BUS':
            se+= 'with obj=BUS'
            se+= '\n  fltApp : (str) Fault application code'
            se+= "\n\tRequired : 'BUS'"
            se+= '\n\t' +__getErrValue__(str,fltApp)
            raise ValueError(se)
    else:
        flag = type(fltApp)==str
        if flag:
            flag = fltApp.upper()=='CLOSE-IN'
            if not flag:
                try:
                    percent = float(fltApp[:fltApp.index('%')])
                    flag = percent>=0 and percent<=100
                except:
                    flag = False
                if not fltApp.endswith('%'):
                    flag = False
        if not flag:
            se+= ' with obj=RLYGROUP/TERMINAL'
            se+= '\n  fltApp : (str) Fault application code'
            se+= "\n\tRequired : 'CLOSE-IN' or 'xx%' (0<=xx<=100)"
            se+= '\n\t' +__getErrValue__(str,fltApp)
            raise ValueError(se)
    #
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        se+= '\nfltConn : Fault connection code'
        se+= "\n\tRequired       :(str) in " + str(__OLXOBJ_fltConn__[:7])
        se+= ",\n\t                          " + str(__OLXOBJ_fltConn__[7:])
        se+= '\n\t' +__getErrValue__(str,fltConn)
        raise ValueError(se)
    #
    flag = type(deviceOpt)==list and len(deviceOpt)==7
    if flag:
        for i in range(7):
            flag = flag and type(deviceOpt[i])==int and (deviceOpt[i] in {0,1})
    if not flag:
        se+= "\ndeviceOpt : Simulation options flags"
        se+= "\n\tRequired          : [int]*7 (1 - set; 0 - reset)"
        se+= "\n\t\tdeviceOpt[0] = 0/1 Consider OCGnd operations"
        se+= "\n\t\tdeviceOpt[1] = 0/1 Consider OCPh operations"
        se+= "\n\t\tdeviceOpt[2] = 0/1 Consider DSGnd operations"
        se+= "\n\t\tdeviceOpt[3] = 0/1 Consider DSPh operations"
        se+= "\n\t\tdeviceOpt[4] = 0/1 Consider Protection scheme operations"
        se+= "\n\t\tdeviceOpt[5] = 0/1 Consider Voltage relay operations"
        se+= "\n\t\tdeviceOpt[6] = 0/1 Consider Differential relay operations"
        se+= '\n\t' +__getErrValue__(list,deviceOpt)
        se+= '\n\t                     ' + str(deviceOpt)
        raise ValueError(se)
    if type(tiers)!=int or tiers<2:
        se+= '\ntiers : Study extent. Take into account protective devices located within this number of tiers only'
        se+= "\n\tRequired     (int): >=2"
        se+= '\n\t' +__getErrValue__(int,tiers)
        raise ValueError(se)
    #
    if type(Z)!=list or len(Z)!=2 or type(Z[0]) not in {float,int} or type(Z[1]) not in {float,int}:
        se+= '\n\tZ : [R,X] Fault impedances in Ohm'
        se+= '\n\tRequired      : [float,float]'
        se+= '\n\t' +__getErrValue__(list,Z)
        raise ValueError(se)
#
def __checkSEA_EX__(sp):
    time = sp.__paraInput__['time']
    fltConn = sp.__paraInput__['fltConn']
    Z = sp.__paraInput__['Z']
    #
    se = '\nSPEC_FLT.SEA_EX(time,fltConn,Z)'
    if type(time) not in {int,float} or time<=0:
        se+= '\ttime [s]: time of Additional Event'
        se+= "\n\tRequired : >0"
        se+= '\n\t' +__getErrValue__(float,time)
        raise ValueError(se)
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        se+= '\nfltConn : Fault connection code'
        se+= "\n\tRequired       :(str) in " + str(__OLXOBJ_fltConn__[:7])[:-1]
        se+= ",\n\t                          " + str(__OLXOBJ_fltConn__[7:])[1:]
        se+= '\n\t' +__getErrValue__(str,fltConn)
        raise ValueError(se)
    #
    if type(Z)!=list or len(Z)!=2 or type(Z[0]) not in {float,int} or type(Z[1]) not in {float,int}:
        se+= '\n\tZ : [R,X] Fault impedances in Ohm'
        se+= '\n\tRequired      : [float,float]'
        se+= '\n\t' +__getErrValue__(list,Z)
        raise ValueError(se)
#
def __getSEA_Result__(index):
    dTime = c_double(0)
    dCurrent = c_double(0) # Highest phase fault current magnitude at this step
    nUserEvent = c_int(0) # flag showing is this is an user-defined event
    szEventDesc = create_string_buffer(b'\000' * 512 * 4)     # 4*512 bytes buffer for event description
    szFaultDesc = create_string_buffer(b'\000' * 512 * 50)    # 50*512 bytes buffer for fault description
    nStep = OlxAPI.GetSteppedEvent(c_int(index),byref(dTime),byref(dCurrent),byref(nUserEvent),szEventDesc,szFaultDesc)
    if index==0: # step = 0 to get total number of steps
        return nStep-1
    res = dict()
    res['time'] = dTime.value
    res['currentMax'] = dCurrent.value
    res['EventFlag'] = nUserEvent.value
    res['EventDesc'] = OlxAPI.decode(cast(szEventDesc, c_char_p).value)
    res['FaultDesc'] = OlxAPI.decode(cast(szFaultDesc, c_char_p).value)
    res['breaker'] = [] # @TO ADD
    res['device'] = []  # @TO ADD
    return res
#
def __runSimulate__(specFlt,clearPrev):
    global __INDEX_SIMUL__,__COUNT_FAULT__,__TYPEF_SIMUL__,FltSimResult
    #
    if clearPrev==1 or (type(specFlt)==SPEC_FLT and specFlt.__type__=='SEA') or \
       (type(specFlt)==list and specFlt[0].__type__=='SEA') :
        __INDEX_SIMUL__ +=1
        FltSimResult.clear()
    #
    if type(specFlt)==SPEC_FLT:
        specFlt = [specFlt]
    # check data
    for sp1 in specFlt:
        sp1.checkData()
    # run Classical
    if specFlt[0].__type__=='Classical':
        __TYPEF_SIMUL__ = 'Classical'
        para = specFlt[0].getData()
        if OlxAPIConst.OLXAPI_FAILURE == OlxAPI.DoFault(para['hnd'],para['fltConn'],para['fltOpt'],para['outageOpt'],para['outageLst'],para['R'],para['X'],c_int(clearPrev)):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.PickFault(c_int(-1),c_int(0)): # pick last fault
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("OLCase.simulateFault('Classical')-Classical: Operation executed successfully.")
        #
        des = OlxAPI.FaultDescription(0)
        id1 = des.index('.')
        __COUNT_FAULT__ = int(des[:id1])
    # run SEA
    elif specFlt[0].__type__=='SEA':
        __TYPEF_SIMUL__ = 'SEA'
        para = specFlt[0].getData()
        sn ='SEA-Single'
        for i in range(1, len(specFlt)):
            sn = 'SEA-Multiple'
            para1 = specFlt[i].getData()
            k = 4*i
            para['fltOpt'][k] = para1['fltOpt']
            para['fltOpt'][k+1] = para1['time']
            para['fltOpt'][k+2] = para1['Z'][0]
            para['fltOpt'][k+3] = para1['Z'][1]
        #
        if OlxAPIConst.OLXAPI_FAILURE == OlxAPI.DoSteppedEvent(para['hnd'],para['fltOpt'],para['runOpt'],para['tiers']):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("OLCase.simulateFault('%s'): Operation executed successfully."%sn)
        #
        __COUNT_FAULT__ = __getSEA_Result__(0)
    else:
        # run simultaneous
        __TYPEF_SIMUL__ = 'Simultaneous'
        import xml.etree.ElementTree as ET
        data = ET.Element('SIMULATEFAULT')
        data.set('CLEARPREV',str(clearPrev))
        data1 = ET.SubElement(data, 'FAULT')
        for i in range(len(specFlt)):
            sp1 = specFlt[i]
            para1 = sp1.getData()
            data2 = ET.SubElement(data1, 'FLTSPEC')
            data2.set('FLTDESC','specFlt: '+str(i+1))
            for k,v in para1.items():
                data2.set(k,v)
        #
        sInput = ET.tostring(data).decode('UTF-8')
        #print(sInput)
        if OlxAPIConst.OLXAPI_FAILURE==OlxAPI.Run1LPFCommand(sInput):
            raise OlxAPI.OlxAPIException(OlxAPI.ErrorString())
        print("OLCase.simulateFault('Simultaneous'): Operation executed successfully.")
        OlxAPI.PickFault(c_int(-1),c_int(0)) # pick last fault
        des = OlxAPI.FaultDescription(0)
        id1 = des.index('.')
        __COUNT_FAULT__ = int(des[:id1])
    # finish ---------------------------------------------
    for sp1 in specFlt:
        print('\t'+sp1.toString())
    for i in range(len(FltSimResult),__COUNT_FAULT__):
        FltSimResult.append( RESULT_FLT(i+1) )
#
def __errorsPara__(ob,sPara):
    sname = type(ob).__name__
    se = '\nAll methods for %s: \n\t'%sname
    for vi in ob.__allMethods__:
        se+=vi +'(),'
    se = se[:-1]+'\nAll attributes for %s:'%sname
    for a1 in ob.getAttributes():
        se+='\n\t' + a1.ljust(15)+ ' : '
        try:
            se+=__OLXOBJ_PARA__[sname][a1][1]
        except:
            if type(ob) in __OLXOBJ_EQUIPMENTO__:
                if a1=='BUS':
                    se+='[BUS] List of Buses of ' +sname
                elif a1=='RLYGROUP':
                    se+='[RLYGROUP] List of RLYGROUPs of ' +sname
                elif a1=='TERMINAL':
                    se+='[TERMINAL] List of TERMINALs of ' +sname
        #
        if type(ob)==RECLSR:
            if a1.startswith('PH_') :
                se+= __OLXOBJ_PARA__[sname][a1[3:]][2]
            elif a1.startswith('GR_') :
                se+= __OLXOBJ_PARA__[sname][a1[3:]][3]
        #
        if a1 in ob.__paraUDF__:
            se+='(str) %s User-Defined Field'%sname
        elif a1=='GUID':
            se+='(str) %s GUID'%sname
        elif a1=='TAGS':
            se+='(str) %s TAGS'%sname
        elif a1=='JRNL':
            se+='(str) %s Creation and modification log records'%sname
        elif a1=='MEMO':
            se+='(str) %s Memo'%sname
        elif a1=='HANDLE' :
            se+= '(int) %s handle'%sname
    #
    if type(ob) in {RLYOCG,RLYOCP,RLYDSG,RLYDSP}:
        aset = ob.getSetting()
        if sPara in aset.keys():
            se+= "\ntry to call %s.getSetting('%s')"%(sname,sPara)
        else:
            se+='\nAll sPara for %s.getSetting(sPara):\n'%sname +str(list(aset.keys()))
    #
    se+= "\nAttributeError:\n\t%s object has no attribute/method '%s'" %(sname,ob.__spara__)
    raise AttributeError (se)
#
def __getSubStation__(b1,setEqui,setEquiHnd,xsmall,lsmall):
    setBus = []
    for t1 in b1.TERMINAL:
        e1 = t1.EQUIPMENT
        eh1 = e1.__hnd__
        if eh1 not in setEquiHnd:
            if type(e1) not in {LINE,DCLINE2}:
                b23 = t1.BUS23
                setBus.extend(b23)
                setEqui.append(e1)
                setEquiHnd.add(eh1)
                #
                for bi in b23:
                    b2e = __getSubStation__(bi,setEqui,setEquiHnd,xsmall,lsmall)
                    setBus.extend(b2e)
            # check small LINE
            elif type(e1)==LINE and xsmall is not None and lsmall is not None:
                kv1 = t1.BUS1.KV
                zbase = kv1**2/OLCase.BASEMVA
                r1,x1 = e1.R*zbase,e1.X*zbase #Ohm
                ln = e1.LN
                if ln>0.0:
                    ln *= __convert_LengthUnit__('km',e1.UNIT.lower())
                if r1<xsmall and x1<xsmall and ln<lsmall:
                    b2 = t1.BUS23[0]
                    setBus.append(b2)
                    setEqui.append(e1)
                    setEquiHnd.add(eh1)
                    b2e = __getSubStation__(b2,setEqui,setEquiHnd,xsmall,lsmall)
                    setBus.extend(b2e)
    return setBus
#
def __convert_LengthUnit__(len_unit_to,len_unit_from):
    """  convert length unit from [len_unit_from] => [len_unit_to]  """
    dictC = {'ft_km':0.0003048,'kt_km':1.852,'mi_km':1.609344,'m_km':0.001,'km_ft':3280.839895,'km_kt':0.539956803,'km_mi':0.6213711922,'km_m':1000,'km_km':1}
    if len_unit_to == len_unit_from:
        return 1
    try:
        return dictC[len_unit_from+'_km'] * dictC['km_'+len_unit_to]
    except:
        return 1
#