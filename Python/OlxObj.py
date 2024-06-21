"""
ASPEN OLX OBJECT LIBRARY
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "3.5.17"

import uuid
import cmath
import math
import os
import _collections_abc
import xml.etree.ElementTree as ET
from OlxAPIConst import HND_SYS, HND_SC, OLXAPI_OK, OLXAPI_FAILURE, TC_BRANCH
from ctypes import cast, c_int, c_double, c_char, byref, pointer, c_char_p, c_void_p, create_string_buffer, POINTER
from OlxAPI import GetData, SetData, encode3, decode, ErrorString, EquipmentType, GetBusEquipment, FindObj1LPF
import OlxAPIConst
import OlxAPI


def getVersion():
    """ (str) OlxObj version. """
    return __version__

def load_olxapi(dllPath, verbose=True):
    """ Initialize the OlxAPI module.

    Args:
        dllPath (string) : Full path name of the disk folder where the olxapi.dll is located
        verbose (bool)   : (obsolete) replace by OlxObj.setVerbose()

    return: [ OlxAPI engine version, Build Number]

    Remark: Successfull initialization is required before the OlxObj library objects and functions can be utilized.
    """
    OlxAPI.InitOlxAPI(dllPath, __OLXOBJ_VERBOSE__)
    v = OlxAPI.Version()
    va = v.split('.')
    if int(va[0]) <= 15 and int(va[1]) < 5:
        raise Exception('Incorrect OlxAPI.dll version: found '+v+va[0]+'.'+va[1]+'. Required V15.5 or higher.')
    if __OLXOBJ_VERBOSE__ and verbose:
        print('OlxObj Version:'+getVersion())
    return [v, str(OlxAPI.BuildNumber())]


def setVerbose(verbose, intern=False):
    """ set Option to print OlxAPI module infos to the consol.

    Args:
        verbose: (0/1) yes/no option.
    """
    if intern:
        global __OLXOBJ_VERBOSE1__
        __OLXOBJ_VERBOSE1__ = verbose
    else:
        global __OLXOBJ_VERBOSE__
        __OLXOBJ_VERBOSE__ = verbose


def setDefaulAddOBJ(val):
    """ Set Option default parameter of Object when addOBJ.

    Args:
        val: (0/1) yes/no option.
    """
    global __OLXOBJ_NEWADD_DEFAULT__
    __OLXOBJ_NEWADD_DEFAULT__ = val


def toString(v, nRound=5, mode='rectangle'):
    """ Convert Object/value to String.

    Args:
        v: value or object
        nRound: rounding
        mode: 'rectangle' or 'polar' form
    """
    if v is None:
        return 'None'
    t = type(v)
    if t == str:
        if "'" in v:
            return ('"'+v+'"').replace('\n', ' ')
        return ("'"+v+"'").replace('\n', ' ')
    if t == int:
        return str(v)
    if t == float:
        v1 = abs(v)
        if v1 > 1.0:
            s1 = str(round(v, nRound))
            return s1[:-2] if s1.endswith('.0') else s1
        elif v1 < 1e-8:
            return '0'
        s1 = '%.'+str(nRound)+'g'
        return s1 % v
    if t == complex:
        if mode=='rectangle':
            if v.imag >= 0:
                return '('+toString(v.real, nRound)+' +'+toString(v.imag, nRound)+'j)'
            return '('+toString(v.real, nRound)+' '+toString(v.imag, nRound)+'j)'
        elif mode=='polar':
            return toString(abs(v),nRound) +'@'+ toString(cmath.phase(v)*180/cmath.pi,nRound)
        else:
            raise Exception ('Error value of mode=%s not in [rectangle,polar]'%mode)
    try:
        return v.toString()
    except:
        pass
    if t in {list, tuple, set}:
        s1 = ''
        for v1 in v:
            try:
                s1 += toString(v1.toString())+', '
            except:
                s1 += toString(v1, nRound, mode)+', '
        if v:
            s1 = s1[:-2]
        if t == list:
            return '['+s1+']'
        elif t == tuple:
            return '('+s1+')'
        else:
            return '{'+s1+'}'
    if t == dict:
        s1 = ''
        for k1, v1 in v.items():
            s1 += toString(k1)+':'
            try:
                s1 += toString(v1.toString())+', '
            except:
                s1 += toString(v1, nRound, mode)+', '
        if s1:
            s1 = s1[:-2]
        return '{'+s1+'}'
    return str(v)


class DATAABSTRACT:
    """ Abstract Class of Object. """

    def __init__(self, hnd):
        """ Constructor by handle (hnd). """
        super().__setattr__('__ob__', type(self).__name__)
        super().__setattr__('__hnd__', hnd)
        super().__setattr__('__currFileIdx__', __CURRENT_FILE_IDX__)
        #
        paramEx = dict()
        va2 = set()
        flag1, flag2 = True, False
        for v1 in dir(self):
            if v1.startswith('__') and v1.endswith('__'):
                flag1 = False
            elif not flag1:
                flag2 = True
            if flag2:
                va2.add(v1)
        #
        if type(self) == TERMINAL:
            va2 = va2 - {'delete', 'changeData', 'postData'}
        #
        va1 = list(__OLXOBJ_PARA__[self.__ob__].keys())
        va1.append('HANDLE')
        va1.sort()
        paramEx['allAttributes'] = va1
        va2.discard('init')
        va2 = list(va2)
        va2.sort()
        paramEx['allMethods'] = va2
        super().__setattr__('__paramEx__', paramEx)

    @property
    def GUID(self):
        """ (str) Globally Unique IDentifier of Object. """
        return self.getData('GUID')

    @property
    def HANDLE(self):
        """ (int) Handle of Object. """
        return self.getData('HANDLE')

    @property
    def JRNL(self):
        """ (str) Creation and modification log records. """
        return self.getData('JRNL')

    @property
    def KEYSTR(self):
        """ (str) Key of Object. """
        return self.getData('KEYSTR')

    @property
    def MEMO(self):
        """ (str) Memo of Object (Line breaks must be included in the memo string as escape character). """
        return self.getData('MEMO')

    @property
    def PARAMSTR(self):
        """ (str) Parameters in JSON Form of Object. """
        return self.getData('PARAMSTR')

    @property
    def PYTHONSTR(self):
        """ (str) String in python of Object. """
        return self.getData('PYTHONSTR')

    @property
    def TAGS(self):
        """ (str) Tags of Object. """
        return self.getData('TAGS')

    def changeData(self, sParam, value=None):
        """ Change Data of Object.

        Samples:
            b1 = OLCase.findOBJ("[BUS] 28 'ARIZONA' 132 kV")
            b1.changeData('NAME','new CALIFORNIA')                 # change 1 param
            b1.changeData(['NAME','AREANO'],['new CALIFORNIA',3])  # change list of param
            b1.changeData({'NAME':'new CALIFORNIA','AREANO':3})    # change by a dict
        """
        if self.__ob__ == 'TERMINAL':
            raise Exception('TERMINAL cannot be changed/updated data')
        #
        global messError
        if (type(sParam) == list and type(value) == list) or (type(sParam) == tuple and type(value) == tuple):
            if len(sParam) != len(value):
                messError = '\n%s.changeData(sParam,value)' % self.__ob__
                messError += '\n\ttype(sParam)    : %s len=%i ' % (type(sParam).__name__, len(sParam))+toString(sParam)
                messError += '\n\ttype(value)     : %s len=%i ' % (type(value).__name__, len(value))+toString(value)
                messError += '\n\tValueError len(sParam)!=len(value)'
                raise ValueError(messError)
            #
            for i in range(len(sParam)):
                self.changeData(sParam[i], value[i])
            return
        #
        if type(sParam) == dict and value is None:
            for k, v in sParam.items():
                self.changeData(k, v)
            return
        #
        if __check_currFileIdx__(self):
            raise Exception(messError)
        #
        hnd = self.__hnd__
        if type(sParam) != str:
            messError = '\n%s.changeData(sParam,value)' % self.__ob__
            messError += '\n\ttype(sParam)    : str'
            messError += '\n\t'+__getErrValue__(str, sParam)
            raise ValueError(messError)
        sparam0 = sParam
        sParam = sParam.upper()
        #
        if sParam in {'GUID', 'JRNL', 'HANDLE', 'PARAMSTR', 'SETTINGSTR', 'KEYSTR', 'PYTHONSTR'} or (sParam in {'LOGICVARNAME', 'LOGICSTR'} and self.__ob__ == 'SCHEME'):
            raise Exception('\n%s.%s cannot be changed' % (self.__ob__, sparam0))
        try:
            paramCode = __OLXOBJ_PARA__[self.__ob__][sParam][0]
            t0 = __OLXOBJ_PARA__[self.__ob__][sParam][2]
        except:
            paramCode, t0 = 0, -1
        if self.__ob__ in {'RLYDSG', 'RLYDSP'} and sParam == 'DSTYPE':
            vold = self.DSTYPE
            if value == vold:
                return
            t0 = 0

        testD = True
        if self.__ob__ == 'RLYD' and t0 == 'RLYD' and (value is None or value == 'N/A'):
            testD = False
        #
        if t0 == 0:
            messError = '\n%s.%s cannot be changed (%s is read only)' % (self.__ob__, sparam0, sparam0)
            raise Exception(messError)
        if ((paramCode > 0 or sParam in {'MEMO', 'TAGS', 'GR_MEMO', 'PH_MEMO'}) and testD) or (self.__ob__=='MULINE' and t0=='b1b2id'):
            messError = __check00__(t0, value)
            if messError:
                messError = '\n%s.%s = ' % (self.__ob__, sparam0)+toString(value)+'\n\t'+sparam0.ljust(18)+' : '+__OLXOBJ_PARA__[self.__ob__][sParam][1]+messError
                raise TypeError(messError)
        if self.__ob__ == 'MULINE':
            if sParam in {'R','X','FROM1','TO1','FROM2','TO2'}:
                self.__paramEx__['MULINEVAL'][sParam] = value
                return
            if sParam in {'ORIENTLINE1', 'ORIENTLINE2'}:
                self.__paramEx__['MULINEVAL'][sParam] = [BUS(value[0]), BUS(value[1]), value[2]]
                return
        if self.__ob__ == 'RECLSR':
            if sParam[:3] == 'GR_':
                hnd += 1
            if sParam == 'ID':
                value += '_P'
        #
        if sParam == 'POLAR' and self.__ob__ in {'RLYOCG', 'RLYOCP'}:
            value = __convert2Int__(value)
            if self.__ob__ == 'RLYOCG' and value not in {0, 1, 2, 3}:
                messError = '\n%s.%s = ' % (self.__ob__, sparam0)+toString(value)+'\n\t'+sparam0.ljust(18)
                messError += ' : '+ __OLXOBJ_PARA__[self.__ob__][sParam][1]
                messError += '\n\tRequired POLAR     : 0,1,2,3 (OC ground relay)'
                messError += '\n\tFound (ValueError) :  '+toString(value)
                raise ValueError(messError)
            if self.__ob__ == 'RLYOCP' and value not in {0, 2}:
                messError = '\n%s.%s = ' % (self.__ob__, sparam0)+toString(value)+'\n\t'+sparam0.ljust(18)
                messError += ' : '+__OLXOBJ_PARA__[self.__ob__][sParam][1]
                messError += '\n\tRequired POLAR     : 0,2 (OC phase relay)'
                messError += '\n\tFound (ValueError) :  '+toString(value)
                raise ValueError(messError)
        #
        if paramCode > 0:
            if t0 == 'RLYGROUPn' and self.__ob__ == 'RLYGROUP':
                ra, _ = __getRLYGROUPn__(value)
                value = [v1.__hnd__ for v1 in ra]
                value.append(0)
            elif t0 in {'RLYGROUP', 'BUS', 'RLYD'}:
                value = OLCase.findOBJ(t0, value) if testD else 0
            elif t0 == 'BUS0':
                value = OLCase.findOBJ('BUS', value)
            elif t0 == 'equip10':
                ra, _ = __getequip10__(value)
                value = [v1.__hnd__ for v1 in ra]
            elif self.__ob__ == 'RLYV' and sParam == 'SGLONLY':
                value = __bin2int__(value)
            #
            if value is None and sParam == 'LTCCTRL' and self.__ob__ in {'XFMR', 'XFMR3'}:
                self.LTCSIDE = 0
                return
            #
            val1 = __setValue__(hnd, paramCode, value)
            #
            try:
                if OLXAPI_OK == SetData(c_int(hnd), c_int(paramCode), byref(val1)):
                    if sParam == 'POLAR' and self.__ob__ in {'RLYOCG', 'RLYOCP'}:
                        self.__paramEx__['POLAR'] = value
                        __getSettingName__(self)
                        self.postData()
                    if sParam == 'EQUATION' and self.__ob__ == 'SCHEME':
                        self.postData()
                    return
            except:
                pass
            messError = '\n'+ErrorString()+'\nCheck %s.%s' % (self.__ob__, sparam0)
            if sParam == 'EQUATION' and self.__ob__ == 'SCHEME':
                messError += '\ntry with SCHEME.setLogic(logic)'
            raise Exception(messError)
        #
        if sParam == 'MEMO' and self.__ob__ not in {'GENUNIT', 'SHUNTUNIT', 'LOADUNIT'}:
            if OLXAPI_FAILURE == OlxAPI.SetObjMemo(hnd, decode(value)):
                messError = ErrorString()
                if messError == 'SetObjMemo failure: Invalid Device Type':
                    raise Exception('\nCheck %s.%s' % (self.__ob__, sparam0))
                else:
                    raise Exception(self.toString()+'\n'+messError)
            return
        #
        if sParam in {'GR_MEMO', 'PH_MEMO'} and self.__ob__ == 'RECLSR':
            if OLXAPI_FAILURE == OlxAPI.SetObjMemo(hnd, decode(value)):
                messError = ErrorString()
                if messError == 'SetObjMemo failure: Invalid Device Type':
                    raise Exception('\nCheck %s.%s' % (self.__ob__, sparam0))
                else:
                    raise Exception(self.toString()+'\n'+messError)
            return
        # Tags
        if sParam == 'TAGS':
            if OLXAPI_FAILURE == OlxAPI.SetObjTags(hnd, value):
                messError = self.toString()+'\n'+ErrorString()
                raise Exception(messError)
            if self.__ob__ == 'RECLSR':
                if OLXAPI_FAILURE == OlxAPI.SetObjTags(hnd+1, value):
                    messError = self.toString()+'\n'+ErrorString()
                    raise Exception(messError)
            return

        # UDF
        if self.__ob__ in OLCase.__UDF__.keys() and sParam in OLCase.__UDF__[self.__ob__]:
            messError = __check00__('str', value)
            if messError:
                messError = '\n%s.%s = ' % (self.__ob__, sParam)+toString(value)+'\n\t'\
                    +sParam.ljust(18)+' : '+self.__ob__+' User-Defined Field'+messError
                raise TypeError(messError)
            if OLXAPI_OK == OlxAPI.SetObjUDF(hnd, sParam, value):
                return
            messError = '\nError in %s.%s = ' % (self.__ob__, sparam0)+str(value)
            raise Exception(messError)
        #
        __errorsParam__(self, sParam, sparam0, 'changeData')
        raise AttributeError(messError)

    def copyFrom(self, o2):
        """ Copy data (o2) => this Object. """
        if type(o2) != type(self):
            raise Exception('Error type Object in copy Object')
        #
        param = eval(o2.paramstr)
        if self.__ob__ == 'SCHEME':
            logicvar = o2.getLogicVar()
            self.resetScheme(param, logicvar)
        else:
            vignore = {'GUID', 'HANDLE'}
            if self.__ob__ == 'RLYGROUP':
                vignore.add('OPFLAG')
            if self.__ob__ == 'BUS':
                vignore.update({'NO', 'NAME', 'KV'})
            #
            for vi in vignore:
                if vi in param.keys():
                    param.pop(vi)
            self.changeData(param)
            #
            if self.__ob__ in __OLXOBJ_RELAY3__:
                self.changeSetting(eval(o2.settingstr))
            self.postData()
        if __OLXOBJ_VERBOSE__:
            print('\nCopy data from:'+o2.toString()+' => '+self.toString())

    def delete(self):
        """ Delete Object. """
        if type(self) == TERMINAL:
            raise Exception('TERMINAL cannot be deleted')
        if __check_currFileIdx__(self):
            raise Exception(messError)
        s1 = self.toString()
        if OLXAPI_FAILURE == OlxAPI.DeleteEquipment(self.__hnd__):
            raise Exception(ErrorString())
        super().__setattr__('__hnd__', -self.__hnd__)
        if __OLXOBJ_VERBOSE__:
            print('\nDelete:'+s1)

    def equals(self, o2):
        """ (bool) Comparison of 2 Objects (this Object vs o2 Object). """
        try:
            return self.__hnd__ == o2.__hnd__
        except:
            return False

    def getAttributes(self):
        """ [str] List of All attributes of Object. """
        if __check_currFileIdx__(self):
            raise Exception(messError)
        res = self.__paramEx__['allAttributes'][:]
        if self.__ob__ in OLCase.__UDF__.keys():
            res.extend(OLCase.__UDF__[self.__ob__])
        return res

    def getData(self, sParam=None):
        """ Retrieve data of Object.

        Args:
            sParam = None or str or list(str)

        Samples:
            b1.getData('KV')           # b1.KV
            b1.getData(['kV','NAME'])  # {'kV':b1.KV, 'NAME':b1.NAME}
            b1.getData('')             # print in console all sParam
        """
        if sParam is None:
            param1 = {'HANDLE', 'PARAMSTR', 'SETTINGSTR', 'PYTHONSTR', 'KEYSTR', 'LOGICSTR','JRNL'}
            if self.__ob__ == 'SCHEME':
                param1.update({'EQUATION', 'LOGICVARNAME'})
            elif self.__ob__ == 'MULINE':
                param1.update({'ORIENTLINE1', 'ORIENTLINE2'})
            param = list(set(self.getAttributes()) - param1)
            param.sort()
            return self.getData(param)
        #
        if type(sParam) in __OLXOBJ_LISTT__:
            return {s1: self.getData(s1) for s1 in sParam}
        #
        if type(sParam) != str:
            se = '\n%s.getData(sParam)' % self.__ob__
            se += '\n\tRequired sParam     : None or str or list/set/tuple/dict_keys of str'
            se += '\n\t'+__getErrValue__(str, sParam)
            raise TypeError(se)
        #
        if __check_currFileIdx__(self):
            raise Exception(messError)
        hnd = self.__hnd__
        sparam0 = sParam
        sParam = sParam.upper()
        if sParam == 'HANDLE':
            return hnd
        if self.__ob__ == 'MULINE' and sParam in {'ORIENTLINE1', 'ORIENTLINE2'}:
            return self.__paramEx__['MULINEVAL'][sParam]
        #
        if sParam == 'PARAMSTR' and self.__ob__ != 'TERMINAL':
            sr = "{'GUID': "+toString(self.GUID)+','
            scomp2 = 1
            if self.__ob__ == 'SERIESRC' and self.SCOMP == 2:
                self.SCOMP = 1
                self.postData()
                scomp2 = 2
            #
            param = []
            for k in self.getAttributes():
                if k not in __OLXOBJ_PARA__[self.__ob__].keys() or __OLXOBJ_PARA__[self.__ob__][k][2] not in {0,'strk','b1b2id'}:
                    param.append(k)
            param.sort()
            va = self.getData(param)
            if self.__ob__ == 'SCHEME':
                va.pop('EQUATION')
            if self.__ob__ in {'RLYOCG', 'RLYOCP'}:
                va.pop('CT')
            #
            for k, v in va.items():
                if k == 'ID':
                    if self.__ob__ not in __OLXOBJ_RELAY2__:
                        sr += toString(k)+':'+toString(v)+', '
                elif k == 'SCOMP':
                    sr += "'SCOMP': "+str(scomp2)+', '
                elif k in {'MEMO', 'GR_MEMO', 'PH_MEMO'}:
                    sr += toString(k)+': '+toString(encode3(v))+', '
                elif k not in {'GUID', 'HANDLE'}:
                    try:
                        sr += toString(k)+': '+toString(v.toString())+', '
                    except:
                        sr += toString(k)+': '+toString(v)+', '
            if scomp2 == 2:
                self.SCOMP = 2
                self.postData()
            return sr[:-2]+'}'
        #
        if sParam == 'KEYSTR':
            if self.__ob__ == 'BUS' or self.__ob__ in __OLXOBJ_BUS1__:
                b1 = self.BUS
                return '['+toString(b1.NAME)+', '+toString(b1.KV)+']'
            elif self.__ob__ == 'XFMR3':
                b1, b2, b3 = self.BUS1, self.BUS2, self.BUS3
                return '[['+toString(b1.NAME)+', '+toString(b1.KV)+'], ['+toString(b2.NAME)+', '+toString(b2.KV)+\
                        '], ['+toString(b3.NAME)+', '+toString(b3.KV)+'], '+toString(self.CID)+']'
            elif self.__ob__ in __OLXOBJ_EQUIPMENT__:
                b1, b2 = self.BUS1, self.BUS2
                return '[['+toString(b1.NAME)+', '+toString(b1.KV)+'], ['+toString(b2.NAME)+', '+toString(b2.KV)+'], '+toString(self.CID)+']'
            elif self.__ob__ in __OLXOBJ_BUS2__:
                b1 = self.BUS
                return '['+toString(b1.NAME)+', '+toString(b1.KV)+', '+toString(self.CID)+']'
            elif self.__ob__ in __OLXOBJ_RELAY__:
                t1 = self.RLYGROUP.TERMINAL
                ba = t1.BUS
                return '[['+toString(ba[0].NAME)+', '+toString(ba[0].KV)+'], ['+toString(ba[1].NAME)+', '+\
                       toString(ba[1].KV)+'], '+toString(t1.CID)+', '+toString(t1.EQUIPMENT.BRCODE)+', '+toString(self.ID)+']'
            elif self.__ob__ == 'MULINE':
                l1, l2 = self.LINE1, self.LINE2
                sres = '[[['+toString(l1.bus1.NAME)+', '+toString(l1.bus1.KV)+'], ['+toString(l1.bus2.NAME)+', '+toString(l1.bus2.KV)+'], '+toString(l1.CID)+'],'
                sres += '[['+toString(l2.bus1.NAME)+', '+toString(l2.bus1.KV)+'], ['+toString(l2.bus2.NAME)+', '+toString(l2.bus2.KV)+'], '+toString(l2.CID)+']]'
                return sres
            elif self.__ob__ == 'RLYGROUP':
                b1, b2, t1 = self.BUS[0], self.BUS[1], self.TERMINAL
                return '[['+toString(b1.NAME)+', '+toString(b1.KV)+'], ['+toString(b2.NAME)+', '+toString(b2.KV)+'], '+toString(t1.CID)+', '+toString(t1.EQUIPMENT.BRCODE)+']'
            elif self.__ob__ == 'BREAKER':
                b1 = self.BUS
                return '['+toString(b1.NAME)+', '+toString(b1.KV)+', '+toString(self.NAME)+']'
            elif self.__ob__ == 'TERMINAL':
                b1, b2, e1 = self.BUS1, self.BUS2, self.EQUIPMENT
                return '[['+toString(b1.NAME)+', '+toString(b1.KV)+'], ['+toString(b2.NAME)+', '+toString(b2.KV)+'], '+toString(e1.CID)+', '+toString(e1.BRCODE)+']'
            raise Exception('error KEY')
        #
        if sParam == 'TOSTRING':
            return self.toString()
        if sParam == 'PYTHONSTR' and self.__ob__ != 'TERMINAL':
            sres = 'OlxObj.OLCase.addOBJ('+toString(self.__ob__)+','
            sres += self.KEYSTR+','+self.PARAMSTR
            if self.__ob__ in __OLXOBJ_RELAY3__:
                sres += ','+self.SETTINGSTR
            if self.__ob__ == 'SCHEME':
                sres += ','+self.LOGICSTR
            return sres+')'
        if self.__ob__ == 'RLYGROUP':
            if sParam == 'TERMINAL':
                return TERMINAL(hnd=__getDatai__(hnd, OlxAPIConst.RG_nBranchHnd))
            if sParam == 'EQUIPMENT':
                return self.TERMINAL.EQUIPMENT
            if sParam in __OLXOBJ_RELAY__:
                return [OLCase.toOBJ(h1) for h1 in __getRLYGROUP_OBJ__(self, sParam)]
        #
        if sParam == 'SETTINGSTR' and self.__ob__ in __OLXOBJ_RELAY3__:
            sr = '{'
            for k, v in self.getSetting().items():
                try:
                    sr += toString(k)+': '+toString(v.toString())+', '
                except:
                    sr += toString(k)+': '+toString(v)+', '
            return sr[:-2]+' }' if sr!='{' else '{}'
        if sParam == 'ATTRIBUTES':
            return self.getAttributes()
        if sParam == 'SETTINGNAME' and self.__ob__ in __OLXOBJ_RELAY3__:
            return self.__paramEx__['SettingName']
        if self.__ob__ == 'SCHEME' and sParam == 'LOGICSTR':
            va = self.getLogic()
            sr = '{'
            for k, v in va.items():
                try:
                    sr += toString(k)+': ['+toString(v[0])+', '+toString(v[1].toString())+'], '
                except:
                    sr += toString(k)+': '+toString(v)+', '
            return sr[:-2]+'}'
        #
        if self.__ob__ == 'RECLSR' and sParam[:3] == 'GR_':
            hnd += 1
        #
        if self.__ob__ == 'RLYV' and sParam == 'SGLONLY':
            paramCode = __OLXOBJ_PARA__['RLYV']['SGLONLY'][0]
            res = __getData__(hnd, paramCode)
            return __int2bin4__(res)
        #
        if self.__ob__ == 'LINE' and sParam == 'MULINE':
            paramCode = __OLXOBJ_PARA__['LINE']['MULINE'][0]
            return OLCase.toOBJ(__getDatai_Array__(hnd, paramCode))
        #
        try:
            paramCode = __OLXOBJ_PARA__[self.__ob__][sParam][0]
            if paramCode != 0:
                if sParam == 'LTCCTRL' and self.__ob__ in {'XFMR', 'XFMR3'}:
                    pc = __OLXOBJ_PARA__[self.__ob__]['LTCSIDE'][0]
                    side = __getData__(hnd, pc)
                    if side == 0:
                        return None
                res = __getData__(hnd, paramCode)
                return __getOBJ__(res, sParam=sParam)
        except:
            er = ErrorString()
            if er == 'GetParam failure: Invalid Device Handle':
                raise Exception(er)
        #
        if sParam == 'JRNL':
            return OlxAPI.GetObjJournalRecord(hnd)
        if sParam == 'GUID':
            return OlxAPI.GetObjGUID(hnd)
        if sParam == 'TAGS':
            return OlxAPI.GetObjTags(hnd)
        if sParam in {'MEMO', 'GR_MEMO', 'PH_MEMO'} and sParam in self.__paramEx__['allAttributes']:
            s1 = OlxAPI.GetObjMemo(hnd)
            return s1
        #
        if self.__ob__ in OLCase.__UDF__.keys() and sParam in OLCase.__UDF__[self.__ob__]:
            fval = create_string_buffer(b'\000'*OlxAPIConst.MXUDF)
            if OLXAPI_OK == OlxAPI.GetObjUDF(hnd, sParam, fval):
                return decode(fval.value)
            else:
                raise Exception('\nUDF Exception:'+ErrorString())
        #
        if sParam not in self.getAttributes():
            __errorsParam__(self, sParam, sparam0, 'getData')
            raise AttributeError(messError)
        #
        return getattr(self, sParam)

    def getDataList(self, sParam=None):
        """ Retrieves data of Object in List form [].

        Args:
            sParam = None or str or list/tuple(str).

        Samples:
            b1.getDataList('KV')           => [b1.KV]
            b1.getDataList(['kV','NAME'])  => [b1.KV, b1.NAME]
            b1.getDataList(['BUS','LINE']) => [bus, line1,line2]
            b1.getDataList('xxx')          => print help in console.
        """
        va = self.getData(sParam)
        if type(va)!=dict:
            return [va]
        if sParam==None:
            return list(va.values())
        res = []
        for k in sParam:
            k1 = k.upper()
            if type(va[k1])==list:
                res.extend(va[k1])
            else:
                res.append(va[k1])
        return res

    def locate1LObj(self, opt=0):
        """ Locate an object on the 1-line diagram

        Args:
            opt (0/1): If the object is not visible show nearest visible one
                         1 – Set; 0 - Reset

        return: None
        """
        if OLXAPI_OK == OlxAPI.Locate1LObj(self.HANDLE, 1 if opt else 0):
            print('Found: ' + toString(self) )
        else:
            print('Not found: ' + toString(self) )

    def isInList(self, la):
        """ (bool) Check if Object is in List/Set/tuple of Objects. """
        if __check_currFileIdx__(self):
            raise Exception(messError)
        sa = set()
        for l1 in la:
            if type(l1) in __OLXOBJ_LISTT__:
                if self.isInList(l1):
                    return True
            else:
                if type(self) == type(l1):
                    sa.add(l1.__hnd__)
                    if __check_currFileIdx__(l1):
                        raise Exception(messError)
        return self.__hnd__ in sa

    def postData(self):
        """ Perform validation and update Object data in the Network. """
        if self.__ob__ == 'TERMINAL':
            raise Exception('TERMINAL cannot be changed/updated data')
        if self.__ob__ == 'MULINE':
            __MU_setValue__(self)
        if OLXAPI_OK != OlxAPI.PostData(c_int(self.__hnd__)):
            messError = '\n'+self.toString()+'\n'+ErrorString()
            raise Exception(messError)
        if self.__ob__ == 'RECLSR':
            if OLXAPI_OK != OlxAPI.PostData(c_int(self.__hnd__+1)):
                messError = '\n'+self.toString()+'\n'+ErrorString()
                raise Exception(messError)

    def toString(self, option=0):
        """ (str) Text description/composed of Object.

        return:
            option = 0  => (str) text composed for Object
                    BUS   : name and kV
                    EQUIPMENT: Bus, Bus2, (Bus3), Circuit ID and type of a branch
                    RELAY : relay type, name and branch location
            option = 1  => (str) text description of Object
        """
        if __check_currFileIdx__(self):
            raise Exception(messError)
        if option == 0:
            s1 = OlxAPI.PrintObj1LPF(self.__hnd__)
            if self.__ob__ != 'TERMINAL':
                return s1
            id1 = s1.find('[')
            id2 = s1.find(']')
            if id1 >= 0 and id2 >= 0:
                s1 = s1[:id1+1]+'TERMINAL'+s1[id2:]
                tc1 = __getDatai__(self.__hnd__, OlxAPIConst.BR_nType)
                s1 += __OLXOBJ_CONST1__[tc1][2]
            return s1
        if option == 1:
            if self.__ob__ == 'BUS':
                return OlxAPI.FullBusName(self.__hnd__)
            elif self.__ob__ in __OLXOBJ_EQUIPMENT__:
                return OlxAPI.FullBranchName(self.__hnd__)
            elif self.__ob__ in __OLXOBJ_RELAY2__:
                return OlxAPI.FullRelayName(self.__hnd__)
            s1 = OlxAPI.PrintObj1LPF(self.__hnd__)
            id2 = s1.find(']')
            return s1[id2+1:]
        raise ValueError('\n'+self.__ob__+'.toString(option) with option = 0 or 1\n\tFound option = '+str(option))

    def __getattr__(self, name):
        return self.getData(name)

    def __setattr__(self, sParam, value):
        self.changeData(sParam, value)

    def init(val1=None, hnd=None):
        """ Obsolete : no longer supported. """
        messError = '\nThe OBJ.init(..) can no longer be used by the OlxObj module version 2.1.4 or above.'
        messError += "\nYou must use OlxObj.OLCase.findOBJ('OBJ',...) instead."
        raise Exception(messError)


class RELAYABSTRACT(DATAABSTRACT):
    """ Abstract Class of Relay Object. """

    def __init__(self, hnd):
        """ Constructor by handle-hnd. """
        super().__init__(hnd)

    @property
    def ASSETID(self):
        """ (str) AssetID of Relay. """
        return self.getData('ASSETID')

    @property
    def BUS1(self):
        """ (BUS) Bus Local of Relay. """
        return self.RLYGROUP.BUS1

    @property
    def BUS(self):
        """ [BUS] List of Buses of Relay: BUS[0] : Bus Local ; BUS[1],(BUS[2]): Bus Opposite(s). """
        return self.RLYGROUP.BUS

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) LINE,XFMR,... that Relay located on. """
        return self.RLYGROUP.EQUIPMENT

    @property
    def FLAG(self):
        """ (int) In service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def ID(self):
        """ (str) Relay ID. """
        return self.getData('ID')

    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Relay Group that Relay located on. """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)

    def computeRelayTime(self, current, voltage, preVoltage):
        """ Computes operating time at given currents and voltages.

        Args:
            current    = [complex]*5 : relay current in [A] of  phase A, B, C,
                                        and if applicable, Io currents in neutral of transformer windings P and S
            voltage    = [complex]*3 : relay voltage in, line to neutral, in kV of phase A, B and C
            preVoltage = (complex)   : relay pre-fault positive sequence voltage magnitude in kV, line to neutral

        return:
            opTime   = (float): relay operating time in seconds
            opDevice = (str)  : relay operation code:
                        NOP  No operation
                        ZGn  Ground distance zone n tripped
                        ZPn  Phase distance zone n tripped
                        Ix   Overcurrent relay operating quantity: Ia, Ib, Ic, Io, I2, 3Io, 3I2
        """
        res = __computeRelayTime__(self, current, voltage, preVoltage)
        if messError:
            raise Exception(messError)
        return res


class RELAY3ABSTRACT(RELAYABSTRACT):
    """ Abstract Class of Relay Object with setting (RLYOCG,RLYOCP,RLYDSG,RLYDSP). """

    def __init__(self, hnd):
        """ Constructor by handle (hnd). """
        super().__init__(hnd)

    @property
    def SETTINGNAME(self):
        """ [str] List of Settings name. """
        return self.getData('SETTINGNAME')

    @property
    def SETTINGSTR(self):
        """ (str) Setting in JSON Form. """
        return self.getData('SETTINGSTR')

    def changeSetting(self, sSetting, value=None):
        """ change Setting with (sSetting,value) of Relay.

        Samples:
            rl.changeSetting('PT ratio','101:1')
            rl.postData()                # Perform validation and update Object data.
            rl.changeSetting('xx','1')   # help (message error) in Console.
        """
        if type(sSetting) == dict and value is None:
            for s1, v1 in sSetting.items():
                self.changeSetting(s1, v1)
            return
        #
        global messError
        if __check_currFileIdx__(self):
            raise Exception(messError)
        #
        if type(sSetting) != str:
            messError = '\n%s.changeSetting(sSetting=%s,value=%s) ' % (self.__ob__, toString(sSetting), toString(value))
            messError += '\n\tRequired (sSetting) : str'
            messError += '\n\tFound             : '+type(sSetting).__name__+'  '+toString(sSetting)
            raise Exception(messError)
        sSetting1 = sSetting.upper()
        #
        if sSetting1 not in self.__paramEx__['SettingNameU']:
            messError = '\nAll sSetting for %s.changeSetting(sSetting,value):' % self.__ob__
            if self.__ob__ in {'RLYDSG', 'RLYDSP'}:
                messError += ' (DSTYPE = '+toString(self.DSTYPE)+')\n'
            else:
                messError += ' (POLAR = '+toString(self.POLAR)+')\n'
            messError += toString(self.__paramEx__['SettingName'])
            messError += '\n\nAttributeError: %s.changeSetting(' % self.__ob__+toString(sSetting)
            messError += ',%s)\n\t' % toString(value)+toString(sSetting)+': not Found'
            raise AttributeError(messError)
        #
        if type(value) != str:
            messError = '\n%s.changeSetting(sSetting=%s,value=%s) ' % (self.__ob__, toString(sSetting), toString(value))
            messError += '\n\tRequired (value): str'
            messError += '\n\tFound           : '+type(value).__name__+'  '+toString(value)
            raise Exception(messError)
        #
        __changeSetting__(self, sSetting1, value)
        if messError:
            messError = '\n%s.changeSetting(sSetting=%s,value=%s) ' % (self.__ob__, toString(sSetting), toString(value))+messError
            raise Exception(messError)

    def getSetting(self, sSetting=None):
        """ (dict) Get Settings of Relay.

        Samples:
            rl.getSetting('PT ratio')
            rl.getSetting('xxx')           => help (message error) in Console
        """
        global messError
        if __check_currFileIdx__(self):
            raise Exception(messError)
        if sSetting is None:
            sSetting = self.__paramEx__['SettingName']
        if type(sSetting) in __OLXOBJ_LISTT__:
            return {sp1: self.getSetting(sp1) for sp1 in sSetting}
        #
        if type(sSetting) != str:
            messError = '\n%s.getSetting(sSetting=%s) ' % (self.__ob__, toString(sSetting))
            messError += '\n\tRequired sSetting : str'
            messError += '\n\tFound           : '+type(sSetting).__name__+'  '+toString(sSetting)
            raise Exception(messError)
        sSetting1 = sSetting.upper()
        if sSetting1 not in self.__paramEx__['SettingNameU']:
            messError = '\nAll sSetting for %s.getSetting(sSetting):' % self.__ob__
            if self.__ob__ in {'RLYDSG', 'RLYDSP'}:
                messError += ' (DSTYPE = '+toString(self.DSTYPE)+')\n'
            else:
                messError += ' (POLAR = '+toString(self.POLAR)+')\n'
            messError += toString(self.__paramEx__['SettingName'])
            messError += '\n\nAttributeError:\n\t%s.getSetting(' % self.__ob__
            messError += toString(sSetting)+')\n\t'+toString(sSetting)+': not Found'
            raise AttributeError(messError)
        #
        if self.__ob__ in {'RLYDSG', 'RLYDSP'}:
            valF = __getData__(self.__hnd__, __OLXOBJ_SETTING__[self.__ob__][2])
        else:
            valF = __getSettingOC__(self)
        valD = dict(zip(self.__paramEx__['SettingNameU'], valF[:len(
            self.__paramEx__['SettingNameU'])]))
        return valD[sSetting1]


class OUTAGE:
    def __init__(self, option, G=0):
        """  Construct equipment outage contingencies in Fault Simulation.

        Args:
            -option: (str) Branch outage option code
                'SINGLE'     : one at a time
                'SINGLE-GND' : (LINE only) one at a time with ends grounded
                'DOUBLE'     : two at a time
                'ALL'        : all at once
            -G: Admittance Outage line grounding (mho) (option=SINGLE-GND)
        """
        super().__setattr__('__outageLst__', [])
        super().__setattr__('__option__', option)
        super().__setattr__('__G__', __convert2Float__(G))
        if __check1_Outage__(self):
            raise ValueError(messError)

    def add_outageLst(self, la):
        """ Add to Outage list (la: list of Object to add). """
        se = '\nOUTAGE.add_outageLst(la)'
        if type(la) == list:
            for l1 in la:
                if type(l1) not in {LINE, XFMR, XFMR3, SERIESRC, SHIFTER, SWITCH,TERMINAL}:
                    se += '\n\tRequired: (list)[(str/obj)] in [LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH]'
                    se += '\n\t'+__getErrValue__(list, la)
                    raise ValueError(se)
                if __check_currFileIdx__(l1):
                    raise Exception(messError)
                if not l1.isInList(self.outageLst):
                    self.__outageLst__.append(l1)
            return
        if type(la) not in {LINE, XFMR, XFMR3, SERIESRC, SHIFTER, SWITCH}:
            se += '\n\tRequired: (list)[(str/obj)] in [LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH]'
            se += '\n\t'+__getErrValue__(None, la)
            raise ValueError(se)
        if __check_currFileIdx__(la):
            raise Exception(messError)
        if not la.isInList(self.outageLst):
            self.__outageLst__.append(la)

    def build_outageLst(self, obj, tiers, wantedTypes):
        """ Return list of neighboring branches.

        Args:
            -obj: BUS or TERMINAL or RLYGROUP or EQUIPMENT (LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH)
            -tiers: Number of tiers-neighboring (must be positive)
            -wantedTypes [(str/obj)]: Branch type to consider
                ['L','X','T','S','P','DC','W'] or
                ['LINE','XFMR','XFMR3','SERIESRC','SHIFTER','SWITCH']

        Samples:
            o1.build_outageLst(obj=t1,tiers=2,wantedTypes=['XFMR'])
        """
        wantedTypes1 = []
        if type(wantedTypes) == list:
            for w1 in wantedTypes:
                if type(w1) == str:
                    wantedTypes1.append(w1.upper())
                else:
                    try:
                        wantedTypes1.append(w1.__name__)
                    except:
                        wantedTypes1.append(w1)
        else:
            wantedTypes1 = wantedTypes
        tiers = __convert2Int__(tiers)
        if __check1_Outage__(self):
            raise ValueError(messError)
        if __check2_Outage__(self, obj, tiers, wantedTypes1):
            raise ValueError(messError)
        #
        if len(wantedTypes1) == 0:
            vw = 1+2+4+8+16
        else:
            wa = set([w1.upper() for w1 in wantedTypes1])
            vw = sum([ __OLXOBJ_EQUIPMENTB5__[w1] for w1 in wa])
        #
        wantedT = c_int(vw)
        listLen = c_int(0)
        hnd = obj.__hnd__
        OlxAPI.MakeOutageList(hnd, c_int(tiers-1), wantedT, None, pointer(listLen))
        branchList = (c_int*(5+listLen.value))(0)
        OlxAPI.MakeOutageList(hnd, c_int(tiers-1), wantedT, branchList, pointer(listLen))
        for i in range(listLen.value):
            r1 = __getOBJ__(branchList[i])
            if not r1.isInList(self.outageLst):
                self.__outageLst__.append(r1)

    def reset_outageLst(self):
        """ Remove all Outage List. """
        self.__outageLst__.clear()

    def toString(self):
        """ (str) Text description/composed of OUTAGE. """
        if __check1_Outage__(self):
            raise ValueError(messError)
        #
        res = '{option='+str(self.option)
        if self.option == 'SINGLE-GND':
            res += ',G='+str(self.G)
        res += ',List=['
        for o1 in self.outageLst:
            res += o1.__ob__+','
        res = res[:-1]+'](len=%i)}' % len(self.outageLst)
        return res

    def __getattr__(self, sParam):
        sParam0 = sParam
        if type(sParam) == str:
            sParam = sParam.upper()
            if sParam == 'OUTAGELST':
                return self.__outageLst__
            if sParam == 'OPTION':
                return self.__option__
            if sParam == 'G':
                return self.__G__
            if sParam == 'TOSTRING':
                return self.toString()
        messError = __errorOutage__(sParam0)
        raise ValueError(messError)

    def __setattr__(self, sParam, value):
        sParam0 = sParam
        if type(sParam) == str:
            sParam = sParam.upper()
            if sParam == 'OPTION':
                return super().__setattr__('__option__', value)
            if sParam == 'OUTAGELST':
                return super().__setattr__('__outageLst__', value)
            if sParam == 'G':
                return super().__setattr__('__G__', __convert2Float__(value))
        #
        messError = __errorOutage__(sParam0)
        raise ValueError(messError)


class RESULT_FLT:
    def __init__(self, index):
        """ Constructor of a Fault Solution case.

        Args:
            index (int): case index.
        """
        if __COUNT_FAULT__ == 0:
            raise Exception('\nNo Fault Simulation result buffer is empty.')
        if type(index) == RESULT_FLT:
            index = index.__index__
        if type(index) != int or index <= 0 or index > __COUNT_FAULT__:
            se = '\nRESULT_FLT(index)'
            se += '\n  Case index'
            se += '\n\tRequired           : (int) 1-%i' % __COUNT_FAULT__
            se += '\n\t'+__getErrValue__(int, index)
            raise ValueError(se)
        self.__index__ = index
        self.__tiers__ = __TIERS_FAULT__
        self.__index_simul__ = __INDEX_SIMUL__
        if __TYPEF_SIMUL__ == 'SEA':
            self.__SEAResult__ = __getSEA_Result__(index)
            self.__fdes__ = str(index)+'. '+self.__SEAResult__['FaultDesc']
            self.__SEAResult__.pop('FaultDesc')
        else:
            self.__fdes__ = OlxAPI.FaultDescriptionEx(index, 0)

    @property
    def CONVERGED(self):
        """ (bool) Convergent Simulation FLAG (True/False). """
        if __checkFault__(self):
            raise Exception(messError)
        r = __getDatai1__(HND_SC, OlxAPIConst.FT_vFlag, 1, True)
        if r==0:
            return True
        if r==OlxAPIConst.FTFLAG_ITERFAILURE:
            print("Warning: Iterative solution did not converge")
        elif r==OlxAPIConst.FTFLAG_PREFAULTITERFAILURE:
            print("Warning: Iterative solution did not converge in pre-fault")
        return False

    @property
    def FAULTDESCRIPTION(self):
        """ (str) Fault description string. """
        if __checkFault__(self):
            raise Exception(messError)
        return self.__fdes__

    @property
    def MVA(self):
        """ (float) Short circuit MVA. """
        if __pickFault__(self):
            raise Exception(messError)
        return __getDatad__(HND_SC, OlxAPIConst.FT_dMVA, excep=True)

    @property
    def SEARES(self):
        """ (dict) SEA Results : Stepped-Event Analysis.

        return:

            SEARes['time']      : (float) SEA event time stamp [s].
            SEARes['currentMax']: (float) Highest phase Fault current magnitude at this step.
            SEARes['EventFlag'] : (0/1)   flag of user-defined event.
            SEARes['EventDesc'] : (str)   Event Description.
            SEARes['breaker']   : []      List of Breakers that operated in the current SEA event.
            SEARes['device']    : []      List of Devices that tripped in the current SEA event.
        """
        if __checkFault__(self):
            raise Exception(messError)
        try:
            return self.__SEAResult__
        except:
            raise Exception('\nSEARES is only available for SEA (Stepped-Event Analysis)')

    @property
    def TIERS(self):
        """ (int) Number of tiers in this solution scope. """
        if __checkFault__(self):
            raise Exception(messError)
        return self.__tiers__

    @property
    def THEVENIN(self):
        """ [complex]*3 Thevenin equivalent impedance [Zp,Zn,Z0]. """
        if __pickFault__(self):
            raise Exception(messError)
        ra = []
        for c1 in [OlxAPIConst.FT_dRPt, OlxAPIConst.FT_dXPt, OlxAPIConst.FT_dRNt, OlxAPIConst.FT_dXNt, OlxAPIConst.FT_dRZt, OlxAPIConst.FT_dXZt]:
            ra.append(__getDatad__(HND_SC, c1, excep=True))
        return [complex(ra[0], ra[1]), complex(ra[2], ra[3]), complex(ra[4], ra[5])]

    @property
    def XR_RATIO(self):
        """ [float]*4 Short circuit RATIO [X/R , ANSI X/R, R0/X1 , X0/X1].

            ANSI X/R = None if uncheck box 'Compute ANSI x/r ratio' in File\\Preferences\\X/R.
        """
        if __pickFault__(self):
            raise Exception(messError)
        ra = []
        for c1 in [OlxAPIConst.FT_dXPt, OlxAPIConst.FT_dRZt, OlxAPIConst.FT_dXZt, OlxAPIConst.FT_dXR]:
            ra.append(__getDatad__(HND_SC, c1, excep=True))
        aisi = __getDatad__(HND_SC, OlxAPIConst.FT_dXRANSI, excep=False)
        res = [ra[3], aisi, ra[1]/ra[0], ra[2]/ra[0]]
        return res

    def current(self, obj=None):
        """ [complex] Retrieve ABC Phase post Fault current on a Object or at the fault.

        Args:
            obj: Object None,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,GEN,GENUNIT,GENW3,GENW4,CCGEN
                         LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL

        return: [complex]

            -obj=None
                    => [IA,IB,IC] total Fault current.
            -obj=GEN,GENUNIT,GENW3,GENW4,CCGEN,LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
                    => [IA,IB,IC]
            -obj=LINE,DCLINE2,SERIESRC,SWITCH
                    => [IA1,IB1,IC1, IA2,IB2,IC2] current from Bus1, Bus2.
            -obj=XFMR,SHIFTER
                    => [IA1,IB1,IC1,IN1, IA2,IB2,IC2,IN2]
                      with IN1,IN2 are Neutral current of winding on Bus1, Bus2.
            -obj=XFMR3
                    => [IA1,IB1,IC1,IN1, IA2,IB2,IC2,IN2, IA3,IB3,IC3,Id3]
                       with IN1,IN2 are Neutral current of winding on Bus1, Bus2.
                            Id3 circulating current on Bus 3 (3-W Transformer only: Delta).
        """
        if __pickFault__(self):
            raise Exception(messError)
        if obj is not None and type(obj) not in __OLXOBJ_IFLT__:
            se = '\nRESULT_FLT.current(obj) unsupported Object'
            se += '\n\tSupported Object: None'
            for vi in __OLXOBJ_IFLT__:
                se += ','+vi.__name__
            se += '\n\tFound : '+obj.__ob__
            raise ValueError(se)
        return __currentFault__(obj, style=3)

    def currentSeq(self, obj=None):
        """ [complex] Retrieve 012 Sequence post Fault current on a Object or at the Fault.

        Args:
            obj: Object None,XFMR,XFMR3,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,GEN,GENUNIT,GENW3,GENW4,CCGEN
                         LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL

        return: [complex]

            -obj=None
                    => [I0,I1,I2] total Fault current.
            -obj=GEN,GENUNIT,GENW3,GENW4,CCGEN,LOAD,LOADUNIT,SHUNT,SHUNTUNIT,TERMINAL
                    => [I0,I1,I2]
            -obj=LINE,DCLINE2,SERIESRC,SWITCH
                    => [I0_1,I1_1,I2_1 , I0_2,I1_2,I2_2] current from Bus1, Bus2.
            -obj=XFMR,SHIFTER
                    => [I0_1,I1_1,I2_1,IN1 , I0_2,I1_2,I2_2,IN2]
                      with IN1,IN2 are Neutral current of winding on Bus1, Bus2.
            -obj=XFMR3
                    => [I0_1,I1_1,I2_1,IN1 , I0_2,I1_2,I2_2,IN2 , I0_3,I1_3,I2_3,Id3]
                       with IN1,IN2 are Neutral current of winding on Bus1, Bus2.
                            Id3 circulating current on Bus 3 (3-W Transformer only: Delta).
        """
        if __pickFault__(self):
            raise Exception(messError)
        if obj is not None and type(obj) not in __OLXOBJ_IFLT__:
            se = '\nRESULT_FLT.currentSeq(obj) unsupported Object'
            se += '\n\tSupported Object: None'
            for vi in __OLXOBJ_IFLT__:
                se += ','+vi.__name__
            se += '\n\tFound : '+obj.__ob__
            raise ValueError(se)
        return __currentFault__(obj, style=1)

    def optime(self, obj, mult, signalonly):
        """ get Operating time of a protective device or logic scheme in a fault.

        Args:
            obj        : a protective device or logic scheme.
            mult       : Relay current multiplying factor.
            signalonly : [int] Consider relay element signal-only flag 1 - Yes; 0 - No

        return: (time,code) = (float,str)

            - time: relay operating time in [s].
            - code: relay operation code.
                    NOP  No operation.
                    ZGn  Ground distance zone n tripped.
                    ZPn  Phase distance zone n tripped.
                    TOC=value Time overcurrent element operating quantity in secondary amps.
                    IOC=value Instantaneous overcurrent element operating quantity in secondary amps.
        """
        if __pickFault__(self):
            raise Exception(messError)
        if type(obj) not in {FUSE, RLYOCG, RLYOCP, RLYDSG, RLYDSP, RECLSR, RLYD, RLYV, SCHEME} or\
                type(mult) not in {float, int} or signalonly not in {0, 1}:
            se = '\nRESULT_FLT.optime(obj,mult,signalonly)'
            se += '\n\t             Required  Found'
            se += '\n\tobj          RELAY     '+obj.__ob__.ljust(8)
            se += ' with RELAY=FUSE,RLYOCG,RLYOCP,RLYDSG,RLYDSP,RECLSR,RLYD,RLYV'
            se += '\n\tmult         float     '+str(mult).ljust(8)
            se += ' Relay current multiplying factor'
            se += '\n\tsignalonly   0,1       '+str(signalonly).ljust(8)
            se += ' Consider relay element signal-only flag'
            raise ValueError(se)
        #
        sx = create_string_buffer(b'\000'*128)
        triptime = c_double(0)
        if OLXAPI_OK != OlxAPI.GetRelayTime(obj.__hnd__, c_double(mult), c_int(signalonly), byref(triptime), sx):
            raise Exception(ErrorString())
        return triptime.value, decode(sx.value)

    def preVoltage(self, obj):
        """ Retrieves pre-fault voltage positive sequence, line to neutral (kV) of a Object.

        Args:
            obj: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3

        return: [complex]
        """
        if __pickFault__(self):
            raise Exception(messError)
        if type(obj) not in {BUS, XFMR, XFMR3, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
            se = '\nRESULT_FLT.preVoltage(obj) unsupported Object'
            se += '\n\tSupported Object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'
            se += '\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        return __preVoltage__(obj, 1)

    def preVoltagePU(self, obj):
        """ Retrieve pre-fault voltage positive sequence, line to neutral (pu) of a Object.

        Args:
            obj: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3

        return: [complex]
        """
        if __pickFault__(self):
            raise Exception(messError)
        if type(obj) not in {BUS, XFMR, XFMR3, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
            se = '\nRESULT_FLT.preVoltagePU(obj) unsupported Object'
            se += '\n\tSupported Object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'
            se += '\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        return __preVoltage__(obj, 2)

    def setScope(self, tiers=9, buses=[]):
        """ Set scope for compute solution results.

        Args:
            tiers (int): number of tiers around faulted bus to compute solution results.
        """
        if __checkFault__(self):
            raise Exception(messError)
        if type(tiers) != int or tiers < 0:
            se = '\nRESULT_FLT.setScope()\n\ttiers: an integer>=0 is required (got type %s) tiers=' % type(tiers).__name__+str(tiers)
            raise TypeError(se)
        self.__tiers__ = tiers

    def voltage(self, obj):
        """ Retrieve ABC PHASE post-fault voltage of a Object (or of connected BUSES of Object).

        Args:
            obj : (Object) BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3,TERMINAL

        return: [complex]

            -obj=BUS   => [VA,VB,VC]
            -obj=XFMR3 => [VA1,VB1,VC1 , VA2,VB2,VC2 , VA3,VB3,VC3]
            -obj=XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,TERMINAL => [VA1,VB1,VC1 , VA2,VB2,VC2]
        """
        if __pickFault__(self):
            raise Exception(messError)
        if type(obj) not in {BUS, XFMR, XFMR3, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH, TERMINAL}:
            se = '\nRESULT_FLT.voltage(obj) unsupported Object'
            se += '\n\tSupported Object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3,TERMINAL'
            se += '\n\tFound : '+type(obj).__name__
            raise ValueError(se)
        return __voltageFault__(obj, style=3)

    def voltageSeq(self, obj):
        """ Retrieve 012 SEQUENCE post-fault voltage of a Object (or of connected BUSES of Object).

        Args:
            obj   : (Object) BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3

        return: [complex]

            -obj=BUS   => [V0,V1,V2]
            -obj=XFMR3 => [V0_1,V1_1,V2_1 , V0_2,V1_2,V2_2 , V0_3,V1_3,V2_3]
            -obj=XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH => [V0_1,V1_1,V2_1 , V0_2,V1_2,V2_2]
        """
        if __pickFault__(self):
            raise Exception(messError)
        #
        if type(obj) not in {BUS, XFMR, XFMR3, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
            se = '\nRESULT_FLT.voltageSeq(obj) unsupported Object'
            se += '\n\tSupported Object: BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'
            se += '\n\tFound : '+type(obj).__name__
            raise AttributeError(se)
        return __voltageFault__(obj, style=1)

    def __getattr__(self, sParam):
        sParam0 = sParam
        if type(sParam) == str:
            sParam = sParam.upper()
            if sParam == 'FAULTDESCRIPTION':
                return self.FAULTDESCRIPTION
            if sParam == 'MVA':
                return self.MVA
            if sParam == 'SEARES':
                return self.SEARES
            if sParam == 'TIERS':
                return self.TIERS
            if sParam == 'THEVENIN':
                return self.THEVENIN
            if sParam == 'XR_RATIO':
                return self.XR_RATIO
            if sParam == 'CONVERGED':
                return self.CONVERGED
        #
        ma = [['current', '[complex] Retrieve ABC PHASE post fault current on a object or at the fault'],
              ['currentSeq', '[complex] Retrieve 012 SEQUENCE post fault current on a object or at the fault'],
              ['optime', '(time,code) Operating time of a protective device or logic scheme in a fault'],
              ['preVoltage', '[complex] Retrieves pre-fault voltage positive sequence, line to neutral (kV)'+'\n'.ljust(25)+\
                'of a BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'],
              ['preVoltagePU', '[complex] Retrieves pre-fault voltage positive sequence, line to neutral (PU)'+'\n'.ljust(25)+\
                'of a BUS,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH,XFMR3'],
              ['setScope', '(None) set scope for NETWORK access'],
              ['voltage', '[complex] Retrieve ABC PHASE post-fault voltage of a BUS or of connected BUSES' +
               '\n'.ljust(25)+'of a LINE,XFMR,XFMR3,SWITCH,SHIFTER,SERIESRC'],
              ['voltageSeq', '[complex] Retrieve 012 SEQUENCE post-fault voltage of a BUS or of connected BUSES'+\
                '\n'.ljust(25)+'of a LINE,XFMR,XFMR3,SWITCH,SHIFTER,SERIESRC']]
        messError = '\nAll methods for RESULT_FLT:'
        for m1 in ma:
            messError += '\n\t'+(m1[0]+'()').ljust(20)+m1[1]
        messError += '\nAll attributes for RESULT_FLT:'
        aa = [['FaultDescription', '(str) Fault description string'],
              ['MVA', '(float) Short circuit MVA'],
              ['SEARes', '(dict) SEA Results (Stepped-Event Analysis)'],
              ['tiers', '(int) Number of tiers in this solution result scope'],
              ['Thevenin', '[complex]*3 Thevenin equivalent impedance [Zp,Zn,Z0]'],
              ['XR_ratio', '[float]*4 Short circuit RATIO [X/R, ANSI X/R, R0/X1, X0/X1] '+'\n'.ljust(25)+'ANSI X/R = None if uncheck box Compute ANSI x/r ratio'],
              ['converged','(bool) Convergent Simulation FLAG (True/False).']]
        for m1 in aa:
            messError += '\n\t'+m1[0].ljust(20)+m1[1]
        messError += '\nAttributeError: '+toString(sParam0)
        raise AttributeError(messError)


FltSimResult = []


class SPEC_FLT:
    """ Object define Fault Specfication (Classical Fault,Simultaneous Fault,Stepped-Event Analysis). """

    def __init__(self, param, typ):
        """ Constructor by (param,type) - used in static method. """
        super().__setattr__('__paramInput__', param)
        super().__setattr__('__type__', typ)
        self.checkData()

    def Classical(obj, fltApp, fltConn, Z=None, outage=None):
        """ Define a Specfication for Classical Fault.

        Args:
            -obj: Classical Fault Object.
                BUS or RLYGROUP or TERMINAL
            -fltApp: (str) Fault application
                'Bus'
                'Close-in'
                'Close-in-EO' (EO: End Opened)
                'Remote-bus'
                'Line-end'
                'xx%'         (xx:Intermediate Fault location % 0<=xx<=100)
                'xx%-EO'
            -fltConn: (str) Fault connection code
                '3LG'
                '2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA'
                '1LG:A' ,'1LG:B' ,'1LG:C'
                'LL:BC' ,'LL:CA' ,'LL:AB' ,'LL:CB' ,'LL:AC' ,'LL:BA'
            -Z : [R,X] Fault Impedance (Ohm)
            -outage : (OUTAGE) outage option.

        Samples:
            fs1 = SPEC_FLT.Classical(obj=b1, fltApp='BUS', fltConn='2LG:AB', Z=[0.1,0.2])
            fs1 = SPEC_FLT.Classical(obj=b1, fltApp='xx', fltConn='xx',Z='xx')           => help (message error) in Console.
        """
        #
        param = {'obj': obj, 'fltApp': fltApp, 'fltConn': fltConn, 'Z': __convert2Float__(Z), 'outage': outage}
        sp = SPEC_FLT(param, 'Classical')
        if sp.checkData():
            raise ValueError(messError)
        return sp

    def SEA_EX(time, fltConn, Z):
        """ Define a Stepped-Event Analysis with Additional Event.

        Args:
            - time: [s] time of Additional Event.
            - fltConn: (str) Fault connection code.
                '3LG'
                '2LG:BC','2LG:CA','2LG:AB','2LG:CB','2LG:AC','2LG:BA'
                '1LG:A' ,'1LG:B' ,'1LG:C'
                'LL:BC' ,'LL:CA' ,'LL:AB' ,'LL:CB' ,'LL:AC' ,'LL:BA'
            - Z: [R,X]: Fault Impedance (Ohm)

        Samples:
            fs2 = SPEC_FLT.SEA_EX(time=0.01,fltConn='2LG:AB',Z=[0.01,0.02])
            fs2 = SPEC_FLT.SEA_EX(time='xx',fltConn='xx',Z='xx')  => help (message error) in Console.
        """
        param = {'time': __convert2Float__(time), 'fltConn': fltConn, 'Z': __convert2Float__(Z)}
        sp = SPEC_FLT(param, 'SEA_EX')
        if sp.checkData():
            raise ValueError(messError)
        return sp

    def SEA(obj, fltApp, fltConn, deviceOpt, tiers, Z=None):
        """ Define a Stepped-Event Analysis.

        Args:
            - obj: Stepped-Event Analysis Object.
                BUS or RLYGROUP or TERMINAL
            - fltApp: (str) Fault application code.
                'Bus'
                'Close-In'
                'xx%' (xx: Intermediate Fault location %)
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

        Samples:
            fs1 = SPEC_FLT.SEA(obj=t1,fltApp='15%',fltConn='1LG:A',deviceOpt=[1,1,1,1,1,1,1],tiers=5,Z=[0,0])
            fs1 = SPEC_FLT.SEA(obj='xx',fltApp='xx',fltConn='xx',deviceOpt='xx',tiers='xx',Z='xx')   => help (message error) in Console.
        """
        param = {'obj': obj, 'fltApp': fltApp, 'fltConn': fltConn, 'deviceOpt': __convert2Int__(deviceOpt),\
                 'tiers': __convert2Int__(tiers), 'Z': __convert2Float__(Z)}
        sp = SPEC_FLT(param, 'SEA')
        if sp.checkData():
            raise ValueError(messError)
        return sp

    def Simultaneous(obj, fltApp, fltConn, Z=None):
        """ Define a Specfication for Simultaneous Fault:

        Args:
            -obj: Simultaneous Fault Object.
                BUS       : fltApp='Bus'
                [BUS,BUS] : fltApp='Bus2Bus'
                RLYGROUP or TERMINAL (on a LINE)      : fltApp='xx%' (with xx: Intermediate Fault location)
                RLYGROUP or TERMINAL : fltApp='Close-In','Line-End','Outage','1P-Open','2P-Open','3P-Open'
            -fltApp: (str) Fault application code.
                'Bus'
                'Close-In'
                'Bus2Bus'
                'Line-End'
                'xx%'           (xx:Intermediate Fault location %)
                'Outage'
                '1P-Open','2P-Open','3P-Open'
            -fltConn: (str) Fault connection code.
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

        Samples:
            fs2 = SPEC_FLT.Simultaneous(obj=t1,fltApp='15%',fltConn='3LG',Z=[0.1,0.2,0.5,0,0,0,0,0.5])
            fs2 = SPEC_FLT.Simultaneous(obj='xx',fltApp='xx',fltConn='xx',Z='xx') => help (message error) in Console.
        """
        param = {'obj': obj, 'fltApp': fltApp, 'fltConn': fltConn, 'Z': __convert2Float__(Z)}
        sp = SPEC_FLT(param, 'Simultaneous')
        if sp.checkData():
            raise ValueError(messError)
        return sp

    def checkData(self):
        """ Check Data of Fault Specfication Object (SPEC_FLT). """
        if self.__type__ == 'Classical':
            return __checkClassical__(self)
        if self.__type__ == 'Simultaneous':
            return __checkSimultaneous__(self)
        if self.__type__ == 'SEA':
            return __checkSEA__(self)
        if self.__type__ == 'SEA_EX':
            return __checkSEA_EX__(self)

    def getData(self):
        """ Get Data of Fault Specfication Object (SPEC_FLT). """
        if self.__type__ == 'Classical':
            param1 = __getClassical__(self)
        elif self.__type__ == 'Simultaneous':
            param1 = __getSimultaneous__(self)
        elif self.__type__ == 'SEA':
            param1 = __getSEA__(self)
        elif self.__type__ == 'SEA_EX':
            param1 = __getSEA_EX__(self)
        else:
            raise Exception('Error type of SPEC_FLT')
        return param1

    def toString(self):
        """ (str) Text description/composed of SPEC_FLT. """
        try:
            obj = self.__paramInput__['obj']
            sp = 'obj='+type(obj).__name__+', '
        except:
            sp = ''
        for k, v in self.__paramInput__.items():
            if k not in {'obj', 'outage'}:
                if k == 'Z' and v is None:
                    sp += k+'=[0]*, '
                else:
                    sp += k+'='+str(v).upper()+', '
            if k == 'outage' and v is not None:
                sp += 'outage='+v.toString()+', '
        return self.__type__+': '+sp[:-2]

    def __setattr__(self, sParam, value):
        sParam0 = sParam
        for k in self.__paramInput__.keys():
            if k.upper() == sParam.upper():
                if k=='Z':
                    self.__paramInput__[k] = __convert2Float__(value)
                else:
                    self.__paramInput__[k] = value
                sParam0 = k
                break
        if __errrorsFaultSpec__(self, sParam0):
            raise AttributeError(messError)

    def __getattr__(self, sParam):
        for k in self.__paramInput__.keys():
            if k.upper() == sParam.upper():
                return self.__paramInput__[k]
        if __errrorsFaultSpec__(self, sParam):
            raise AttributeError(messError)


__K_INI_NETWORK__ = 0
__CURRENT_FILE_IDX__ = 0
__INDEX_FAULT__ = -1
__COUNT_FAULT__ = 0
__TIERS_FAULT__ = 9
__INDEX_SIMUL__ = 1
__TYPEF_SIMUL__ = ''


class NETWORK:
    def __init__(self, val1=None, val2=None):
        """ OLCase Network Constructor. """
        if val1 is not None or val2 is not None:
            se = '\nThe NETWORK class can no longer be instantiated outside the OlxObj module version 2.1.4 or above.'
            se += '\nYou must use the global Object OLCase instead.'
            raise Exception(se)
        global __K_INI_NETWORK__
        __K_INI_NETWORK__ += 1
        if __K_INI_NETWORK__ > 1:
            raise Exception('\nCan not init class NETWORK(), try to use OlxObj.OLCase.')
        #
        va1, va2 = [], []
        flag1, flag2 = True, False
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
        paramEx = {'allAttributes': va1, 'allMethods': va2, 'allBus': None}
        paramEx['scope'] = {'isFullNetWork': True}
        super().__setattr__('__paramEx__', paramEx)

    @property
    def AREA(self):
        """ [AREA] List of Areas Object in Network. """
        return [AREA(a1) for a1 in self.AREANO]

    @property
    def AREANO(self):
        """ [int] List of Areas Number in Network. """
        if self.__paramEx__['allBus'] is None:
            self.__paramEx__['allBus'] = self.getData('BUS')
        va = [b1.getData('AREANO') for b1 in self.__paramEx__['allBus']]
        return sorted(list(set(va)))

    @property
    def BASEMVA(self):
        """ (float) System MVA base. """
        return self.getData('BASEMVA')

    @property
    def BREAKER(self):
        """ [BREAKER] List of Breakers Rating in Network. """
        return [BREAKER(hnd=b1.__hnd__) for b1 in self.getData('BREAKER')]

    @property
    def BUS(self):
        """ [BUS] List of Buses in Network. """
        return [BUS(hnd=b1.__hnd__) for b1 in self.getData('BUS')]

    @property
    def CCGEN(self):
        """ [CCGEN] List of Voltage Controlled Current Sources in Network. """
        return [CCGEN(hnd=b1.__hnd__) for b1 in self.getData('CCGEN')]

    @property
    def COMMENT(self):
        """ (str) File Comments of OLR file. """
        return self.getData('COMMENT')

    @property
    def DCLINE2(self):
        """ [DCLINE2] List of DC Lines in Network. """
        return [DCLINE2(hnd=b1.__hnd__) for b1 in self.getData('DCLINE2')]

    @property
    def FUSE(self):
        """ [FUSE] List of Fuses in Network. """
        return [FUSE(hnd=b1.__hnd__) for b1 in self.getData('FUSE')]

    @property
    def GEN(self):
        """ [GEN] List of Generators in Network. """
        return [GEN(hnd=b1.__hnd__) for b1 in self.getData('GEN')]

    @property
    def GENUNIT(self):
        """ [GENUNIT] List of Generator Units in Network. """
        return [GENUNIT(hnd=b1.__hnd__) for b1 in self.getData('GENUNIT')]

    @property
    def GENW3(self):
        """ [GENW3] List of Type-3 Wind Plants in Network. """
        return [GENW3(hnd=b1.__hnd__) for b1 in self.getData('GENW3')]

    @property
    def GENW4(self):
        """ [GENW4] List of Converter-Interfaced Resources in Network. """
        return [GENW4(hnd=b1.__hnd__) for b1 in self.getData('GENW4')]

    @property
    def KV(self):
        """ [float] List of kV Nominal in Network. """
        if self.__paramEx__['allBus'] is None:
            self.__paramEx__['allBus'] = self.getData('BUS')
        va = [round(b1.getData('KV'), 5) for b1 in self.__paramEx__['allBus']]
        return sorted(list(set(va)), reverse=True)

    @property
    def LINE(self):
        """ [LINE] List of AC Transmission Lines in Network. """
        return [LINE(hnd=b1.__hnd__) for b1 in self.getData('LINE')]

    @property
    def LOAD(self):
        """ [LOAD] List of Loads in Network. """
        return [LOAD(hnd=b1.__hnd__) for b1 in self.getData('LOAD')]

    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List of Load Units in Network. """
        return [LOADUNIT(hnd=b1.__hnd__) for b1 in self.getData('LOADUNIT')]

    @property
    def MULINE(self):
        """ [MULINE] List of Mutual Coupling Pairs in Network. """
        return [MULINE(hnd=b1.__hnd__) for b1 in self.getData('MULINE')]

    @property
    def OBJCOUNT(self):
        """ (dict) Number of Object (by dict) in Network. """
        return self.getData('OBJCOUNT')

    @property
    def OLRFILE(self):
        """ (str) OLR file path. """
        return self.getData('OLRFILE')

    @property
    def RECLSR(self):
        """ [RECLSR] List of Reclosers in Network. """
        return [RECLSR(hnd=b1.__hnd__) for b1 in self.getData('RECLSR')]

    @property
    def RLYD(self):
        """ [RLYD] List of Differential Relays in Network. """
        return [RLYD(hnd=b1.__hnd__) for b1 in self.getData('RLYD')]

    @property
    def RLYDS(self):
        """ [RLYDSG+RLYDSP] List of Distance Relays (Ground+Phase) in Network. """
        return self.getData('RLYDS')

    @property
    def RLYDSG(self):
        """ [RLYDSG] List of Distance Ground Relays in Network. """
        return [RLYDSG(hnd=b1.__hnd__) for b1 in self.getData('RLYDSG')]

    @property
    def RLYDSP(self):
        """ [RLYDSP] List of Distance Phase Relays in Network. """
        return [RLYDSP(hnd=b1.__hnd__) for b1 in self.getData('RLYDSP')]

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of Relay Groups in Network. """
        return [RLYGROUP(hnd=b1.__hnd__) for b1 in self.getData('RLYGROUP')]

    @property
    def RLYOC(self):
        """ [RLYOCG+RLYOCP] List of OverCurrent Relays (Ground+Phase) in Network. """
        return self.getData('RLYOC')

    @property
    def RLYOCG(self):
        """ [RLYOCG] List of OverCurrent Ground Relays in Network. """
        return [RLYOCG(hnd=b1.__hnd__) for b1 in self.getData('RLYOCG')]

    @property
    def RLYOCP(self):
        """ [RLYOCP] List of OverCurrent Phase Relays in Network. """
        return [RLYOCP(hnd=b1.__hnd__) for b1 in self.getData('RLYOCP')]

    @property
    def RLYV(self):
        """ [RLYV] List of Voltage Relays in Network. """
        return [RLYV(hnd=b1.__hnd__) for b1 in self.getData('RLYV')]

    @property
    def SCHEME(self):
        """ [SCHEME] List of Logic Schemes in Network. """
        return [SCHEME(hnd=b1.__hnd__) for b1 in self.getData('SCHEME')]

    @property
    def SERIESRC(self):
        """ [SERIESRC] List of Series capacitor/reactors in Network. """
        return [SERIESRC(hnd=b1.__hnd__) for b1 in self.getData('SERIESRC')]

    @property
    def SHIFTER(self):
        """ [SHIFTER] List of Phase Shifters in Network. """
        return [SHIFTER(hnd=b1.__hnd__) for b1 in self.getData('SHIFTER')]

    @property
    def SHUNT(self):
        """ [SHUNT] List of Shunts in Network. """
        return [SHUNT(hnd=b1.__hnd__) for b1 in self.getData('SHUNT')]

    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List of Shunt Units in Network. """
        return [SHUNTUNIT(hnd=b1.__hnd__) for b1 in self.getData('SHUNTUNIT')]

    @property
    def SVD(self):
        """ [SVD] List of Switched Shunts in Network. """
        return [SVD(hnd=b1.__hnd__) for b1 in self.getData('SVD')]

    @property
    def SWITCH(self):
        """ [SWITCH] List of Switches in Network. """
        return [SWITCH(hnd=b1.__hnd__) for b1 in self.getData('SWITCH')]

    @property
    def XFMR(self):
        """ [XFMR] List of 2-Windings Transformers in Network. """
        return [XFMR(hnd=b1.__hnd__) for b1 in self.getData('XFMR')]

    @property
    def XFMR3(self):
        """ [XFMR3] List of 3-Windings Transformers in Network. """
        return [XFMR3(hnd=b1.__hnd__) for b1 in self.getData('XFMR3')]

    @property
    def ZCORRECT(self):
        """ [ZCORRECT] List of Impedance Correction Tables in Network. """
        return [ZCORRECT(hnd=b1.__hnd__) for b1 in self.getData('ZCORRECT')]

    @property
    def ZONE(self):
        """ [ZONE] List of Zones Object in Network. """
        return [ZONE(a1) for a1 in self.ZONENO]

    @property
    def ZONENO(self):
        """ [int] List of Zones Number in Network. """
        if self.__paramEx__['allBus'] is None:
            self.__paramEx__['allBus'] = self.getData('BUS')
        va = [b1.getData('ZONENO') for b1 in self.__paramEx__['allBus']]
        return sorted(list(set(va)))

    def addOBJ(self, ob, key, param={}, setting={}):
        """ add/modify Object to OLCase (OLR file).

        Args:
            ob    : (str) Object 'BUS','LINE',...
            key   : Key define of Object (BUS = ['fieldale1',132],...)
            param : Parameters of Object.

        Samples:
            param = {'AREANO':2,'LOCATION':'CLAYTOR','MEMO':'m1','MIDPOINT':0,'NO':0,'SLACK':1,'SPCX':15,'SPCY':25,'SUBGRP':10,'TAGS':'tag1;tag2;','TAP':0,'ZONENO':2}
            obn   = OLCase.addOBJ('BUS',['fieldale1',132],param)

            param = {'GUID':'{48c97fdd-5fab-4cbd-8601-18b5b50856e6}','BACKUP':[],'INTRPTIME':0.5,'MEMO':b'grtg','PRIMARY':[],'RECLSRTIME':[2.11,1,0.5,0]}
            obn   = OLCase.addOBJ('RLYGROUP',[['NEVADA',132],['OHIO',132],'1','L'],param)
        """
        if __check_currFileIdx1__():
            raise Exception(messError)
        setVerbose(0, 1)
        ob = __updateSTR1__(ob)
        if ob in {'GEN','GENUNIT','GENW3','GENW4','CCGEN','LOAD','LOADUNIT','SHUNT','SHUNTUNIT','SVD'} and type(key)==list:
            if len(key)==2 and type(key[0])==list:
                k1 = key[0]
            elif len(key)==3:
                k1 = key[:2]
            else:
                k1 = key
            try:
                if None==self.findOBJ('BUS',k1):
                    self.addOBJ('BUS',k1)
            except:
                pass
            setVerbose(0, 1)
        if __check_ob_key__('\nOlxObj.OLCase.addOBJ', ob, key, param, setting):
            raise ValueError(messError)
        param1 = {k.upper(): param[k] for k in param.keys()}
        if 'GUID' not in param.keys() or not param['GUID']:
            param1['GUID'] = __getGUID__('')
        if ob == 'RLYGROUP' and 'RECLSRTIME' in param1.keys():
            try:
                if param1['RECLSRTIME'][0] == 0.0:
                    param1.pop('RECLSRTIME')
            except:
                pass
        if ob in {'RLYOCG', 'RLYOCP'} and 'CT' in param1.keys() and 'CTSTR' in param1.keys():
            param1.pop('CT')
        #
        o1 = OLCase.findOBJ(ob, key)
        if o1 != None:
            param1.pop('GUID')
            if ob=='MULINE':
                param1['ORIENTLINE1'] = key[0]
                param1['ORIENTLINE2'] = key[1]
            __updateOBJNew__(o1, param1, setting)
            if messError:
                raise Exception(messError)
            setVerbose(1, 1)
            return o1
        #
        if ob == 'BUS':
            v0 = [['NAME', key[0], 'str'], ['KV', key[1], 'float']]
            if 'NAME' in param1.keys():
                param1.pop('NAME')
        elif ob in __OLXOBJ_BUS1__:
            b1 = self.findOBJ('BUS', key)
            v0 = [['BUS', b1.HANDLE, 'int']]
        elif ob in __OLXOBJ_BUS2__:
            ob1 = ob[:-4]
            if len(key) == 3:
                g1 = self.findOBJ(ob1, key[:-1])
            else:
                g1 = self.findOBJ(ob1, key[0])
            #
            if g1 is None:
                if len(key) == 3:
                    g1 = self.addOBJ(ob1, key[:-1])
                else:
                    b1 = self.findOBJ('BUS', key[0])
                    g1 = self.addOBJ(ob1, b1)
            v0 = [[ob1, g1.HANDLE, 'int'], ['CID', key[-1], 'str']]
        elif ob == 'XFMR3':
            b1 = self.findOBJ('BUS', key[0])
            b2 = self.findOBJ('BUS', key[1])
            b3 = self.findOBJ('BUS', key[2])
            v0 = [['BUS1', b1.HANDLE, 'int'], ['BUS2', b2.HANDLE, 'int'],
                  ['BUS3', b3.HANDLE, 'int'], ['CID', key[3], 'str']]
        elif ob in __OLXOBJ_EQUIPMENT__:
            b1 = self.findOBJ('BUS', key[0])
            b2 = self.findOBJ('BUS', key[1])
            v0 = [['BUS1', b1.HANDLE, 'int'], ['BUS2', b2.HANDLE, 'int'], ['CID', key[2], 'str']]
        elif ob == 'MULINE':
            l1 = self.findOBJ('LINE', key[0])
            l2 = self.findOBJ('LINE', key[1])
            v0 = [['LINE1', l1.HANDLE, 'int'], ['LINE2', l2.HANDLE, 'int']]
        elif ob == 'RLYGROUP':
            t1 = TERMINAL(key)
            v0 = [t1.HANDLE, 0]
        elif ob in __OLXOBJ_RELAY__:
            t1 = TERMINAL(key[:-1])
            v0 = [t1.HANDLE, 0]
            param1['ID'] = key[4]
        elif ob == 'BREAKER':
            b1 = BUS(key[0]) if len(key) <= 2 else BUS(key[:-1])
            v0 = [['NAME', key[-1], 'str'], ['BUS', b1.HANDLE, 'int']]
        else:
            raise Exception('\nNot yet API for %s please contact: support@aspeninc.com' % ob)
        #
        setVerbose(1, 1)
        if ob == 'MULINE':
            param2  = {'FROM1': [0]*5, 'FROM2': [0]*5, 'TO1': [100, 0, 0, 0, 0], 'TO2': [100, 0, 0, 0, 0], 'X': [1e-6, 0, 0, 0, 0]}
            param3 = {'ORIENTLINE1' : key[0], 'ORIENTLINE2' : key[1]}
            for k1,v1 in param1.items():
                if k1 in {'R','X','FROM1','TO1','FROM2','TO2','MEMO'}:
                    param3[k1] = v1
                else:
                    param2[k1] = v1
            o1 = __addOBJ__(ob, v0, param2, setting)
            o1.changeData(param3)
            o1.postData()
        elif ob == 'RECLSR':
            o1 = __addOBJ_RECLSR__(v0, param1)
        elif ob == 'SCHEME':
            o1 = __addOBJ_SCHEME__(v0, param1, setting)
        else:
            o1 = __addOBJ__(ob, v0, param1, setting)
        if o1 == None:
            if ob == 'SHIFTER' and 'ZCORRECTNO' in param1.keys():
                param1.pop('ZCORRECTNO')
                o1 = __addOBJ__(ob, v0, param1, setting)
            if o1 == None:
                raise Exception(messError)
        return o1

    def addUDFTemplate(self, obj, valudf):
        """ Add UDF Template of Object.

        Args:
            obj    : 'BUS','LINE'...
            valudf : [(str)Field name, (str)Display Label, (int)Order]

        Samples:
            OLCase.addUDFTemplate('BUS',['BUSUDF2', 'busudf2', 7])
        """
        messError = '\nOLCase.addUDFTemplate(obj,valudf)'
        if type(obj) != str or obj.upper() not in __OLXOBJ_LISTUDF__:
            messError = '\nOBJUDF: '+str(__OLXOBJ_LISTUDF__)+'\n'+messError
            messError += '\n\tRequired obj :  (str) in OBJUDF'
            messError += '\n\tFound        :  (%s) ' % type(obj).__name__+toString(obj)
            raise ValueError(messError)
        #
        if type(valudf) != list or len(valudf) != 3:
            messError += '\n\tRequired valudf :  [(str)Field name, (str)Display Label, (int)Order]'
            if type(valudf) == list:
                messError += '\n\tFound           :  (%s) len=%i' % (type(valudf).__name__, len(valudf))
            else:
                messError += '\n\tFound           :  (%s) ' % type(valudf).__name__+toString(valudf)
            raise ValueError(messError)
        #
        if type(valudf[0]) != str or type(valudf[1]) != str or type(valudf[2]) != int:
            messError += '\n\tRequired valudf :  [(str)Field name, (str)Display Label, (int)Order]'
            messError += '\n\tFound           :  [%s , %s , %s]' % (type(valudf[0]).__name__, type(valudf[1]).__name__, type(valudf[2]).__name__)
            raise ValueError(messError)
        #
        obj = obj.upper()
        char10K = c_char*10000
        buff = char10K()
        if OLXAPI_OK != GetData(HND_SYS, OlxAPIConst.SY_sUDFTemplate, buff):
            raise Exception(ErrorString())
        val = decode(buff.value)+'\n'+obj+'\t'+str(valudf[2])+'\t'+valudf[0]+'\t'+valudf[1]
        #
        buff = cast(pointer(c_char_p(encode3(val))), c_void_p)
        if OLXAPI_OK != SetData(HND_SYS, OlxAPIConst.SY_sUDFTemplate, buff):
            raise Exception(ErrorString())
        #
        if OLXAPI_OK != OlxAPI.PostData(HND_SYS):
            raise Exception(ErrorString())
        #
        self.__UDF__[obj].append(valudf[0].upper())

    def applyScope(self, areaNum=None, zoneNum=None, optionTie=0, kV=[], verbose=True):
        """ Apply Scope for Network access.

        Args:
            areaNum    : [int]/int/str: area Number [1,2] or 3 or '1,3,4-10'
            zoneNum    : [int]/int/str: zone Number [1,2] or 3 or '1,3,4-10'
            optionTie  : 0-strictly in areaNum/zoneNum
                         1-with tie
                         2-only tie
            kV       []: [kVmin,kVmax]
            verbose  : (obsolete) replace by OlxObj.setVerbose()

        return:
            (str) infos of Network Access Scope.
        """
        se = '\nOLCase.applyScope(areaNum,zoneNum,optionTie,kV)'
        areaNum = __getAREAZONE__(areaNum, 'areaNum', se)
        if messError:
            raise ValueError(messError)
        zoneNum = __getAREAZONE__(zoneNum, 'zoneNum', se)
        if messError:
            raise ValueError(messError)
        #
        if kV and (not __checkListType__(kV, float, 2) or kV[0] > kV[1]):
            se += '\n\tRequired kV       : [kVmin,kVmax] with kVmin<=kVMax'
            se += '\n\tFound (ValueError): '+str(kV)
            raise ValueError(se)
        if optionTie not in {0, 1, 2}:
            se += '\n\tRequired optionTie : 0/1/2 0- in selected areas/zones; 1- with tie ; 2- only tie'
            se += '\n\tFound (ValueError) : '+str(optionTie)
            raise ValueError(se)
        #
        self.__scope__ = {'areaNum': areaNum, 'zoneNum': zoneNum, 'optionTie': optionTie,
                          'kV': kV, 'isFullNetWork': not (areaNum or zoneNum or kV)}
        #
        s1 = 'NetWork Access Scope: '
        if self.__scope__['isFullNetWork']:
            s1 += 'Full Network'
        else:
            s1 += '\n\tareaNum   : '+str(areaNum)
            s1 += '\n\tzoneNum   : '+str(zoneNum)
            s1 += '\n\toptionTie : '+str(optionTie)+' (0- in selected areas/zones; 1- with tie ; 2- only tie)'
            s1 += '\n\tkV        : '+str(kV)
        if __OLXOBJ_VERBOSE__:
            print(s1)
        return s1
    #
    def busPicker(self, title = 'Pick Buses'):
        """ Displays the bus selection dialog.

        return:
            [BUS]: List of Buses selected
        """
        Opt = (c_int)(0)
        BusList = (c_int*100)(0)
        if OLXAPI_OK == OlxAPI.BusPicker(c_char_p(encode3(title)), BusList, Opt):
            res = []
            for b1 in BusList:
                if b1>0:
                    print('You selected:', OlxAPI.PrintObj1LPF(b1))
                    res.append(b1)
            return self.toOBJ(res)
        print('You pressed Cancel')
        return []


    def checkInit(self, title=''):
        try:
            OlxAPI.__checkInit__(1)
        except Exception as err :
            if title:
                import tkinter as tk
                import tkinter.messagebox as tkm
                root = tk.Tk()
                root.attributes('-topmost', True)
                root.withdraw()
                tkm.showerror(title, str(err))
            raise Exception(err)

    def close(self):
        """ Close the network data file that had been loaded previously return 0 if OK. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        if OLXAPI_OK == OlxAPI.CloseDataFile():
            global __CURRENT_FILE_IDX__, __INDEX_SIMUL__
            if __CURRENT_FILE_IDX__ > 0:
                __CURRENT_FILE_IDX__ *= -1
            if __INDEX_SIMUL__ > 0:
                __INDEX_SIMUL__ *= -1
            return 0
        return 1

    def create_Network(self, baseMVA):
        """ Create a new Network with baseMVA. """
        OlxAPI.CreateNetwork(baseMVA)
        global __CURRENT_FILE_IDX__, __INDEX_SIMUL__, __COUNT_FAULT__, FltSimResult
        __CURRENT_FILE_IDX__ = abs(__CURRENT_FILE_IDX__)+1
        __INDEX_SIMUL__ += 1
        __COUNT_FAULT__ = 0
        FltSimResult.clear()
        self.__setattr__('__currFileIdx__', __CURRENT_FILE_IDX__)
        self.__setattr__('__scope__', {'isFullNetWork': True})
        self.__setattr__('__BUS__', None)

    def findOBJ(self, ob, key=None):
        """ Find Object by key.

        Args:
            ob  : 'BUS','LINE',....'RLYD','SCHEME',...
            key : Key define of Object (BUS = ['fieldale1',132],...)

        Samples:

            b1 = OLCase.findOBJ('BUS',['arizona',132])
            b1 = OLCase.findOBJ("[BUS] 28 'ARIZONA' 132 kV")              #STR
            b1 = OLCase.findOBJ('{1bdfa3eb-2992-40cf-8d12-eb7f9f484126}') #GUID

            l1 = OLCase.findOBJ('LINE',"[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") # by STR
            l1 = OLCase.findOBJ("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")

            dg1 = OLCase.findOBJ("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")
            dg1 = OLCase.findOBJ('RLYDSG',[6,['REUSENS',132],'1','L','NV_Reusen G1'])
        """
        if __check_currFileIdx1__():
            raise Exception(messError)
        if key is None and type(ob) == str:
            hnd = __findObj1LPF__(ob).value
            return self.toOBJ(hnd) if hnd>0 else None
        #
        ob = __updateSTR1__(ob)
        if ob == '':
            for ob1 in __OLXOBJ_LIST__:
                try:
                    o1 = self.findOBJ(ob1, key)
                    if o1 != None:
                        return o1
                except:
                    pass
            return None
        ka = {'DEVICE':['RLYD','RLYV'], 'RLYDS': ['RLYDSG', 'RLYDSP'], 'RLYOC': ['RLYOCG', 'RLYOCP']}
        if type(ob) == str and ob in ka.keys():
            try:
                r1 = self.findOBJ(ka[ob][0], key)
                if r1!=None:
                    return r1
                ob = ka[ob][1]
            except:
                raise Exception(messError)
        if __check_ob_key__('\nOlxObj.OLCase.findOBJ', ob, key):
            raise ValueError(messError)
        #
        try:
            if ob == 'BUS':
                return BUS(key)
            if ob == 'LINE':
                return LINE(key)
            if ob == 'SERIESRC':
                return SERIESRC(key)
            if ob == 'SWITCH':
                return SWITCH(key)
            if ob == 'DCLINE2':
                return DCLINE2(key)
            if ob == 'XFMR':
                return XFMR(key)
            if ob == 'XFMR3':
                return XFMR3(key)
            if ob == 'SHIFTER':
                return SHIFTER(key)
            if ob == 'GEN':
                return GEN(key)
            if ob == 'GENUNIT':
                return GENUNIT(key)
            if ob == 'GENW3':
                return GENW3(key)
            if ob == 'GENW4':
                return GENW4(key)
            if ob == 'CCGEN':
                return CCGEN(key)
            if ob == 'LOAD':
                return LOAD(key)
            if ob == 'LOADUNIT':
                return LOADUNIT(key)
            if ob == 'SHUNT':
                return SHUNT(key)
            if ob == 'SHUNTUNIT':
                return SHUNTUNIT(key)
            if ob == 'SVD':
                return SVD(key)
            if ob == 'RLYGROUP':
                return RLYGROUP(key)
            if ob == 'RLYOCG':
                return RLYOCG(key)
            if ob == 'RLYOCP':
                return RLYOCP(key)
            if ob == 'RLYDSG':
                return RLYDSG(key)
            if ob == 'RLYDSP':
                return RLYDSP(key)
            if ob == 'RLYD':
                return RLYD(key)
            if ob == 'RLYV':
                return RLYV(key)
            if ob == 'FUSE':
                return FUSE(key)
            if ob == 'RECLSR':
                return RECLSR(key)
            if ob == 'BREAKER':
                return BREAKER(key)
            if ob == 'TERMINAL':
                return TERMINAL(key)
            if ob == 'MULINE':
                return MULINE(key)
            if ob == 'SCHEME':
                return SCHEME(key)
        except Exception as err:
            if __OLXOBJ_VERBOSE__ and __OLXOBJ_VERBOSE1__:
                if not str(err).endswith(': Not Found'):
                    print(str(err))
            return None

    def findOBJByTag(self, tags, sObj=None):
        """ Find list of Object of sObj type that has the given tags.

        Args:
            - tags : Tag string
            - sObj : [] list of name of Object type.
                   : [] or None      => all Object types.
                   : ['LINE','XFMR'] => object LINE and XFMR.

        return: List of Objects
        """
        if type(tags) != str:
            se = '\nOLCase.findOBJByTag(tags)' if sObj is None else '\nOLCase.findOBJByTag(tags,sObj)'
            se += '\n\tRequired tags      : str'
            se += '\n\t'+__getErrValue__(str, tags)
            raise ValueError(se)
        if sObj is not None:
            flag = False
            if type(sObj) == list:
                for s1 in sObj:
                    if s1 not in __OLXOBJ_LIST__:
                        flag = True
            elif type(sObj) == str:
                if sObj not in __OLXOBJ_LIST__:
                    flag = True
            else:
                flag = True
            if flag:
                se = '\nOLCase.findOBJByTag(tags,sObj)'
                se += '\n\tRequired sObj      : []/str in '+str(__OLXOBJ_LIST__)
                se += '\n\n\tFound (ValueError) : '+str(sObj)
                raise ValueError(se)
        res = []
        equType = 0#c_int(0)
        equHnd = (c_int*1)(0)
        while OLXAPI_OK == OlxAPI.FindEquipmentByTag(str(tags), equType, equHnd):
            r1 = self.toOBJ(equHnd[0])
            if sObj is None or sObj == [] or type(r1).__name__ in sObj:
                if __isInScope__(r1, self.__scope__):
                    res.append(r1)
        return res

    def findBREAKER(self, key):
        """ find BREAKER (None if not found).

        Samples:
            OLCase.findBREAKER('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}')
            OLCase.findBREAKER("[BREAKER]  1E82A@ 6 'NEVADA' 132 kV")
        """
        try:
            return BREAKER(key)
        except:
            __errorNotFound__('BREAKER')

    def findBUS(self, val1=None, val2=None):
        """ find BUS (None if not found).

        Samples:
            OLCase.findBUS('_{220879BF-1297-4314-B10F-53765B11E02C}') # GUID
            OLCase.findBUS("[BUS] 28 'ARIZONA' 132 kV")               # STR
            OLCase.findBUS(28)                                        # Bus Number
            OLCase.findBUS('arizona',132)                             # name,kV
        """
        try:
            return BUS(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('BUS')

    def findCCGEN(self, val1=None, val2=None):
        """ find CCGEN (Voltage controlled current source) (None if not found).

        Samples:
            OLCase.findCCGEN('_{39238FAA-5D98-4F17-A669-09E6BD2E20A0}') #GUID
            OLCase.findCCGEN("[CCGENUNIT] 'BUS5' 13 kV")                #STR
            OLCase.findCCGEN('GLEN LYN',132)                            #Name,kV
        """
        try:
            return CCGEN(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('CCGEN')

    def findDCLINE2(self, val1=None, val2=None, val3=None):
        """ find DCLINE2 (DC Transmission Line) (None if not found).

        Samples:
            OLCase.findDCLINE2('_{2BB0E750-FC57-497E-B878-081D74319DE5}')          #GUID
            OLCase.findDCLINE2("[DCLINE2] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") #STR
            OLCase.findDCLINE2(b1,b2,'1')                                          #bus1,bus2,CID
            OLCase.findDCLINE2(['CLAYTOR',132],['NEVADA',132],'1')
        """
        try:
            return DCLINE2(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('DCLINE2')

    def findFUSE(self, val1):
        """ find FUSE (None if not found).

        Samples:
            OLCase.findFUSE('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}')                    #GUID
            OLCase.findFUSE("[FUSE]  NV Fuse@6 'NEVADA' 132 kV-4 'TENNESSEE' 132 kV 1 P") #STR
        """
        try:
            return FUSE(val1)
        except:
            __errorNotFound__('FUSE')

    def findGEN(self, val1=None, val2=None):
        """ find Generator (None if not found).

        Samples:
            OLCase.findGEN('_{82BC751C-7988-49CF-874A-57C9CEB1ECA3}') #GUID
            OLCase.findGEN("[GENERATOR] 2 'CLAYTOR' 132 kV")          #STR
            OLCase.findGEN('claytor',132)                             #Name,kV
        """
        try:
            return GEN(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('GEN')

    def findGENUNIT(self, val1, val2=None, val3=None):
        """ find Generator Unit (None if not found).

        Samples:
            OLCase.findGENUNIT('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}') #GUID
            OLCase.findGENUNIT("[GENUNIT]  1@2 'CLAYTOR' 132 kV")         #STR
            OLCase.findGENUNIT(b1,'1')                                    #b1 = BUS/GEN
            OLCase.findGENUNIT('claytor',132,'1')
        """
        try:
            if val2 is None:
                return GENUNIT(val1)
            elif val3 is None:
                return GENUNIT([val1, val2])
            else:
                return GENUNIT([val1, val2, val3])
        except:
            __errorNotFound__('GENUNIT')

    def findGENW3(self, val1=None, val2=None):
        """ find GENW3 Type-3 Wind Plant (None if not found).

        Samples:
            OLCase.findGENW3('_{F74E7DEC-C835-406A-B62C-B4DEA4D7C764}') #GUID
            OLCase.findGENW3("[GENW3] 'BUS3' 33 kV")                    #STR
            OLCase.findGENW3('CLAYTOR',132)                             #Name,kV
        """
        try:
            return GENW3(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('GENW3')

    def findGENW4(self, val1=None, val2=None):
        """ find GENW4 Converter-Interfaced Resource (None if not found).

        Samples:
            OLCase.findGENW4('_{9E828F29-0CAF-4992-90AA-5A04D5F01463}') #GUID
            OLCase.findGENW4("[GENW4] 'BUS3' 33 kV")                    #STR
            OLCase.findGENW4('CLAYTOR',132)                             #Name,kV
        """
        try:
            return GENW4(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('GENW4')

    def findLINE(self, val1=None, val2=None, val3=None):
        """ find AC Transmission Line (None if not found).

        Samples:
            OLCase.findLINE('_{7BFBA30C-6476-4A55-BAB6-4DF3A9B9443E}')        #GUID
            OLCase.findLINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")  #STR
            OLCase.findLINE(b1,b2,'1')                                        #Bus1,Bus2,CID
            OLCase.findLINE(['CLAYTOR',132],['NEVADA',1],'1')
        """
        try:
            return LINE(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('LINE')

    def findLOAD(self, val1=None, val2=None):
        """ find LOAD (None if not found).

        Samples:
            OLCase.findLOAD('_{5B7BD740-AAA7-4A34-98BC-CC34B795299D}') #GUID
            OLCase.findLOAD("[LOAD] 17 'WASHINGTON' 33 kV")            #STR
            OLCase.findLOAD('WASHINGTON',33)                           #Name,kV
        """
        try:
            return LOAD(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('LOAD')

    def findMULINE(self, val1):
        """ find MULINE-Mutual Coupling pair (None if not found).

        Samples:
            OLCase.findMULINE("{f1aba445-7598-4c1c-82a5-b908f34b5692}") #GUID
            OLCase.findMULINE("[MUPAIR] 6 'NEVADA' 132 kV-'TAP BUS' 132 kV 1|8 'REUSENS' 132 kV-28 'ARIZONA' 132 kV 1")
            OLCase.findMULINE([[2,5,'1'],[2,6,'1']])                    #Line1,Line2
        """
        try:
            return MULINE(val1)
        except:
            __errorNotFound__('MULINE')

    def findLOADUNIT(self, val1, val2=None, val3=None):
        """ find LOADUNIT-Load Unit (None if not found).

        Samples:
            OLCase.findLOADUNIT('_{68A5C3B4-29FC-41DA-86DD-9412E271E186}') #GUID
            OLCase.findLOADUNIT("[LOADUNIT]  1@17 'WASHINGTON' 33 kV")     #STR
            OLCase.findLOADUNIT(b1,'1')                                    #b1 = BUS/LOAD
            OLCase.findLOADUNIT('claytor',132,'1')
        """
        try:
            if val2 is None:
                return LOADUNIT(val1)
            elif val3 is None:
                return LOADUNIT([val1, val2])
            else:
                return LOADUNIT([val1, val2, val3])
        except:
            __errorNotFound__('LOADUNIT')

    def findRECLSR(self, key):
        """ find Recloser (None if not found).

        Samples:
            OLCase.findRECLSR('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}')              #GUID
            OLCase.findRECLSR("[RECLSRP]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            OLCase.findRECLSR([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RECLSR(key)
        except:
            __errorNotFound__('RECLSR')

    def findRLYD(self, key):
        """ find RLYD-Differential Relay (None if not found).

        Samples:
            OLCase.findRLYD('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}')                 #GUID
            OLCase.findRLYD("[DEVICEDIFF]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            OLCase.findRLYD([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYD(key)
        except:
            __errorNotFound__('RLYD')

    def findRLYDSG(self, key):
        """ find RLYDSG -Distance Ground Relay (None if not found).

        Samples:
            OLCase.findRLYDSG("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                          #GUID
            OLCase.findRLYDSG("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")  #STR
            OLCase.findRLYDSG([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYDSG(key)
        except:
            __errorNotFound__('RLYDSG')

    def findRLYDSP(self, key):
        """ find RLYDSP Distance Phase Relay (None if not found).

        Samples:
            OLCase.findRLYDSP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                    #GUID
            OLCase.findRLYDSP("[DSRLYP]  GCXTEST@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L") #STR
            OLCase.findRLYDSP([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYDSP(key)
        except:
            __errorNotFound__('RLYDSP')

    def findRLYGROUP(self, key):
        """ find Relay Group (None if not found).

        Samples:
            OLCase.findRLYGROUP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")               #GUID
            OLCase.findRLYGROUP("[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            OLCase.findRLYGROUP([bus1,bus2,CID,BRCODE])
        """
        try:
            return RLYGROUP(key)
        except:
            __errorNotFound__('RLYGROUP')

    def findRLYOCG(self, key):
        """ find RLYOCG-Overcurrent Ground Relay (None if not found).

        Samples:
            OLCase.findRLYOCG("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                  #GUID
            OLCase.findRLYOCG("[OCRLYG]  NV-G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            OLCase.findRLYOCG([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYOCG(key)
        except:
            __errorNotFound__('RLYOCG')

    def findRLYOCP(self, key):
        """ find RLYOCP-Overcurrent Phase Relay (None if not found).

        Samples:
            OLCase.findRLYOCP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                   #GUID
            OLCase.findRLYOCP("[OCRLYP]  NV-P1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L")  #STR
            OLCase.findRLYOCP([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYOCP(key)
        except:
            __errorNotFound__('RLYOCP')

    def findRLYV(self, key):
        """ find RLYV-Voltage Relay (None if not found).

        Samples:
            OLCase.findRLYV("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")               #GUID
            OLCase.findRLYV("[DEVICEVR]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            OLCase.findRLYV([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return RLYV(key)
        except:
            __errorNotFound__('RLYV')

    def findSCHEME(self, key):
        """ find SCHEME-Logic Scheme (None if not found).

        Samples:
            OLCase.findSCHEME("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                            #GUID
            OLCase.findSCHEME("[PILOT]  272-POTT-SEL421G@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            OLCase.findSCHEME([bus1,bus2,CID,BRCODE,ID])
        """
        try:
            return SCHEME(key)
        except:
            __errorNotFound__('SCHEME')

    def findSERIESRC(self, val1=None, val2=None, val3=None):
        """ find SERIESRC-Series reactors/capacitor (None if not found).

        Samples:
            OLCase.findSERIESRC('_{7BFBA30C-6476-4A55-BAB6-4DF3A9B9443E}')           #GUID
            OLCase.findSERIESRC("[SERIESRC] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") #STR
            OLCase.findSERIESRC(['CLAYTOR',132],['NEVADA',132],'2')                  #bus1,bus2,CID
        """
        try:
            return SERIESRC(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('SERIESRC')

    def findSHIFTER(self, val1=None, val2=None, val3=None):
        """ find SHIFTER-Phase Shifter (None if not found).

        Samples:
            OLCase.findSHIFTER('_{F915233C-20AB-43C5-A7C4-3BF29F0F890F}')           #GUID
            OLCase.findSHIFTER("[SHIFTER] 'NEVADA PST' 132 kV-6 'NEVADA' 132 kV 1") #STR
            OLCase.findSHIFTER(['NEVADA',132],['NEVADA PST',132],'1')               #bus1,bus2,CID
        """

        try:
            return SHIFTER(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('SHIFTER')

    def findSHUNT(self, val1=None, val2=None):
        """ find SHUNT (None if not found).

        Samples:
            OLCase.findSHUNT('_{5C9526DC-D64A-4E92-AC4D-9D56AE380A05}') #GUID
            OLCase.findSHUNT("[SHUNT] 21 'IOWA' 33 kV")                 #STR
            OLCase.findSHUNT('OHIO',132)                                #Name,kV
        """
        try:
            return SHUNT(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('SHUNT')

    def findSHUNTUNIT(self, val1, val2=None, val3=None):
        """ find SHUNTUNIT-Shunt Unit (None if not found).

        Samples:
            OLCase.findSHUNTUNIT('_{5C9526DC-D64A-4E92-AC4D-9D56AE380A05}') #GUID
            OLCase.findSHUNTUNIT("[CAPUNIT]  1@21 'IOWA' 33 kV")            #STR
            OLCase.findSHUNTUNIT('claytor',132,'1')                         #Name,kV,CID
        """
        try:
            if val2 is None:
                return SHUNTUNIT(val1)
            elif val3 is None:
                return SHUNTUNIT([val1, val2])
            else:
                return SHUNTUNIT([val1, val2, val3])
        except:
            __errorNotFound__('SHUNTUNIT')

    def findSVD(self, val1=None, val2=None):
        """ find SVD-Switched Shunt (None if not found).

        Samples:
            OLCase.findSVD('_{7BFBA30C-6476-4A55-BAB6-4DF3A9B9443E}') #GUID
            OLCase.findSVD("[SVD] 3 'TEXAS' 132 kV")                  #STR
            OLCase.findSVD('Ohio',132)                                #Name,kV
        """
        try:
            return SVD(val1 if val2 is None else [val1, val2])
        except:
            __errorNotFound__('SVD')

    def findSWITCH(self, val1=None, val2=None, val3=None):
        """ find SWITCH (None if not found).

        Samples:
            OLCase.findSWITCH('_{7BFBA30C-6476-4A55-BAB6-4DF3A9B9443E}')          #GUID
            OLCase.findSWITCH("[SWITCH] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1")  #STR
            OLCase.findSWITCH(['CLAYTOR',132],['NEVADA',132],'1')                 #bus1,bus2,CID
        """
        try:
            return SWITCH(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('SWITCH')

    def findTERMINAL(self, val1=None, val2=None, val3=None, val4=None):
        """ find TERMINAL (None if not found).

        Samples:
            OLCase.findTERMINAL(b1,l1)
            OLCase.findTERMINAL(b1,b2,'1','LINE')
                BUS = BUS | int | str | [str float/int]
                CID (str) circuit ID
                BRCODE   in ['X','T','P','L','DC','S','W','XFMR3','XFMR','SHIFTER','LINE','DCLINE2','SERIESRC','SWITCH']
                OBJ (Object) in [XFMR3,XFMR,SHIFTER,LINE,DCLINE2,SERIESRC,SWITCH]
        """
        if val2 is None and val3 is None and val4 is None:
            try:
                return TERMINAL(val1)
            except:
                if not messError.endswith(': Not Found'):
                    raise Exception(messError.replace('.TERMINAL', '.OLCase.findTERMINAL'))
                return None
        #
        if val3 is None and val4 is None:
            try:
                b1 = BUS(val1)
                if type(val2) == str:
                    val2 = OLCase.findOBJ(val2)
                for t1 in val2.TERMINAL:
                    if t1.BUS1.__hnd__ == b1.__hnd__:
                        return t1
            except:
                return None
        try:
            return TERMINAL([val1, val2, val3, val4])
        except:
            __errorNotFound__('TERMINAL')

    def findXFMR(self, val1=None, val2=None, val3=None):
        """ find XFMR 2-windings transformer (None if not found).

        Samples:
            OLCase.findXFMR('_{3C556AB9-563B-45C5-AAA5-38061571FC46}')              #GUID
            OLCase.findXFMR("[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1") #STR
            OLCase.findXFMR('VERMONT',132,'VERMONT',33,'1')                         #bus1,bus2,CID
        """
        try:
            return XFMR(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('XFMR')

    def findXFMR3(self, val1=None, val2=None, val3=None):
        """ find XFMR3 3-windings transformer (None if not found).

        Samples:
            OLCase.findXFMR3('_{9006257F-CD51-4EC8-AE8D-24451B972323}')                                 #GUID
            OLCase.findXFMR3("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-'DOT BUS' 13.8 kV 1") #STR
            OLCase.findXFMR3(['NEVADA',132],['NEVADA1',33],'1')                                         #bus1,bus2,CID
        """
        try:
            return XFMR3(val1 if val2 is None else [val1, val2, val3])
        except:
            __errorNotFound__('XFMR3')

    def getAttributes(self):
        """ [str] List of All attributes. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return self.__paramEx__['allAttributes']

    def getData(self, sParam=None):
        """ Retrieve data of Network by name of parameter.

        Args:
            sParam: None or (str) name of parameter.
        """
        if __check_currFileIdx1__():
            raise Exception(messError)
        if sParam is None:
            return {s1: self.getData(s1) for s1 in self.getAttributes()}
        if type(sParam) in {list, set, tuple}:
            res = []
            for sp1 in sParam:
                v1 = self.getData(sp1)
                if v1 is not None:
                    if type(v1) == list:
                        res.extend(v1)
                    else:
                        res.append(v1)
            return res
        #
        if type(sParam) != str:
            se = '\nin call OLCase.getData(sParam)'
            se += '\n\tRequire  sParam: None or str or list(str)'
            se += '\n\t'+__getErrValue__(str, sParam)
            raise TypeError(se)
        #
        sparam0 = sParam
        sParam = sParam.upper()
        try:
            sParam = __OLXOBJ_STR_KEYS1__[sParam]
        except:
            pass
        if sParam == 'BASEMVA':
            return __getData__(HND_SYS, OlxAPIConst.SY_dBaseMVA)
        if sParam == 'OLRFILE':
            return OlxAPI.GetOlrFileName()
        if sParam == 'COMMENT':
            return __getData__(HND_SYS, OlxAPIConst.SY_sFComment)
        if sParam == 'RLYOC':
            res = self.getData('RLYOCG')
            res.extend(self.getData('RLYOCP'))
            return res
        if sParam == 'RLYDS':
            res = self.getData('RLYDSG')
            res.extend(self.getData('RLYDSP'))
            return res
        if sParam == 'OBJCOUNT':
            res = {}
            for sParam in __OLXOBJ_PARA__['SYS'].keys():
                paramCode = __OLXOBJ_PARA__['SYS'][sParam][0]
                res[sParam] = __getData__(HND_SYS, paramCode)
            return res
        if sParam in __OLXOBJ_CONST__.keys():
            return __getEquipment__(sParam, self.__scope__)
        #
        if sParam not in self.getAttributes():
            se = '\nAll methods of OLCase: '
            for sp in self.__paramEx__['allMethods']:
                se += sp+'(), '
            se = se[:-1]+'\n\nAll attributes of OLCase:'
            for sp in self.__paramEx__['allAttributes']:
                if sp in __OLXOBJ_CONST__.keys():
                    se += '\n'+sp.ljust(15)+' : '+__OLXOBJ_CONST__[sp][2]+' in NETWORK'
                elif sp in __OLXOBJ_PARA__['SYS2'].keys():
                    se += '\n'+sp.ljust(15)+' : '+__OLXOBJ_PARA__['SYS2'][sp][1]
                elif sp=='OLRFILE':
                    se += '\n'+sp.ljust(15)+' : (str) OLR file path'
                else:
                    se += '\n'+sp.ljust(15)
            se += "\nAttributeError : OLCase has no attribute/method '%s'" % sparam0
            raise AttributeError(se)
        return getattr(self, sParam)

    def getOBJSelected(self):
        """ (list) of Selected Object(s) in the 1-line diagram """
        res = []
        hnd = c_int(0)
        while OLXAPI_OK == OlxAPI.GetEquipment(OlxAPIConst.TC_SELECTED, byref(hnd)):
            res.append(self.toOBJ(hnd.value))
        return res

    def getSystemParams(self):
        """ (dict) get System Parameters. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        char10K = c_char*10000
        buff = char10K()
        if OLXAPI_OK != GetData(HND_SYS, OlxAPIConst.SY_vParams, buff):
            raise Exception(ErrorString())
        sysParams = decode(buff.value).split('\t')
        k, res = 0, dict()
        while True:
            if sysParams[k] == '':
                break
            key = sysParams[k]
            val = sysParams[k+1]
            res[key] = val
            k += 2
        return res

    def getUDFTemplate(self):
        """ get UDF Template of all Object.

        return: {'BUS': ['busudf1', 'busudf2'], 'GEN': ['gudf1', 'gudf2'],... }
        """
        char10K = c_char*10000
        buff = char10K()
        if OLXAPI_OK != GetData(HND_SYS, OlxAPIConst.SY_sUDFTemplate, buff):
            raise Exception(ErrorString())
        val = decode(buff.value).split('\n')
        res = {k1: [] for k1 in ['BUS', 'GEN', 'GENUNIT', 'GENW3', 'GENW4', 'CCGEN', 'XFMR', 'XFMR3', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH', 'LOAD',\
                      'LOADUNIT', 'SHUNT', 'SHUNTUNIT', 'SVD', 'BREAKER', 'RLYGROUP', 'RLYOC', 'FUSE', 'RLYDS', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME', 'PROJECT']}
        for v1 in val:
            try:
                va = v1.split('\t')
                res[va[0]].append([va[2], va[3], int(va[1])])
            except:
                pass
        return res

    def open(self, olrFile, readonly=False, verbose=True, olxpath=''):
        """ Read ASPEN OLR data file from disk.

        Args:
            olrFile  : Full path name of ASPEN OLR file.
            readonly : open in read-only mode. 1-true; 0-false
            verbose  : (obsolete) replace by OlxObj.setVerbose()
            olxpath  : Full path name of the disk folder where the olxapi.dll is located.
        """
        verbose = verbose and OlxAPI.ASPENOlxAPIDLL is None
        load_olxapi(olxpath, verbose)
        try:
            OlxAPI.CloseDataFile()
        except:
            pass
        #
        if type(olrFile) != str:
            raise TypeError("File not found: "+str(olrFile))
        olrFile = os.path.abspath(olrFile)
        #
        if not os.path.isfile(olrFile):
            raise ValueError("File not found: "+olrFile)
        #
        ext1 = (os.path.splitext(olrFile)[1]).upper()
        if ext1 not in ['.OLR', '.OLX']:
            raise ValueError('\nError file format (.OLR .OLX) : '+olrFile)
        #
        if OlxAPI.ASPENOlxAPIDLL is None:
            raise Exception('\nNot yet init OlxAPI, try to load_olxapi(dllPath) before to open a NetWork')
        val = OlxAPI.LoadDataFile(olrFile, True if readonly else False)
        if OLXAPI_FAILURE == val:
            raise Exception(ErrorString())
        #
        self.__initEmbed__()

    def resetScope(self):
        """ reset Scope: All Network Access. """
        self.__scope__ = {'isFullNetWork': True}
        if __OLXOBJ_VERBOSE__:
            print('resetScope(): All NETWORK Access')

    def run1LPFCommand(self, cmdParams):
        """ Run OneLiner command.

        Args:
            cmdParams : XML command

        Samples:

            cmdParams = '\
                <SIMULATEFAULT CLEARPREV="1" COUNT="1" TIERSOUT="30" SHOWTYPE="1">\
                    <FAULT FLTCOUNT="1" DESC="  1. Bus Fault on:          19 DELAWARE         33.  kV 1LG Type=A " PREFV="0" GENZTYPE="0" IGNORE="0" GENILIMIT="0" VCCS="1" GENW3="1" GENW4="1" MOV="1" MOVITERF="0.41" MINZMUPAIR="0" MVASTYLE="0">\
                        <FLTSPEC FLTSTR="&apos; &apos; 0 2 0 0 0 0 0 0 0 0 0 0 &apos;DELAWARE&apos; 33 "/>\
                    </FAULT>\
                </SIMULATEFAULT>'
            OLCase.run1LPFCommand(cmdParams)
        """
        if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(cmdParams):
            raise Exception(ErrorString())
        #
        if decode(cmdParams).strip().startswith('<SIMULATEFAULT'):
            if __OLXOBJ_VERBOSE__:
                print('OLCase.OLCase.run1LPFCommand(%s) :\n\tOperation executed successfully.' % (' '.join(decode(cmdParams).split())))
            __runSimulateFin__(len(FltSimResult))


    def save(self, fileNew='', verbose=True):
        """ Save ASPEN OLR data file.

        Args:
            fileNew : name path of new file
                if fileNew=='' (default) => save current file
            verbose  : (obsolete) replace by OlxObj.setVerbose()
        """
        if __check_currFileIdx1__():
            raise Exception(messError)
        fn = os.path.abspath(fileNew if fileNew else OlxAPI.ASPENOLRFILE)
        if OLXAPI_FAILURE == OlxAPI.SaveDataFile(fn):
            raise Exception(ErrorString())
        #
        if OLXAPI_FAILURE == OlxAPI.LoadDataFile(fn, 0, 0):
            raise Exception(ErrorString())
        if __OLXOBJ_VERBOSE__:
            print('File saved successfully: '+fn)

    def setVerbose(self, verbose, intern=False):
        """ (0/1) Option to print OlxAPI module infos to the consol. """
        setVerbose(verbose, intern)

    def simulateFault(self, specFlt, clearPrev=0, verbose=True):
        """ Run simulation of Fault.

        Args:
            specFlt : Classical/Simultaneous/Stepped-Event Analysis specfication(s).
                    SPEC_FLT: 'Classical' or 'Simultaneous' or 'SEA' (Stepped-Event Analysis)
                    [SPEC_FLT] : Simultaneous
                    [sp_0,sp_i,...] sp_0=SPEC_FLT('SEA') sp_i=SPEC_FLT('SEA_EX')
            clearPrev : (0/1) clear previous result flag.
            verbose   :  (obsolete) replace by OlxObj.setVerbose()

        Samples:
            fs1 = SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='2LG:AB',Z=[0.1,0.2])
            OLCase.simulateFault(fs1,0)

            fs1 = SPEC_FLT.Simultaneous(obj=rg,fltApp='CLOSE-IN',fltConn='3LG')
            fs2 = SPEC_FLT.Simultaneous(obj=t1,fltApp='15%',fltConn='3LG',Z=[0.1,0.2,0.5,0,0,0,0,0.5])
            OLCase.simulateFault([fs1],1)
            OLCase.simulateFault([fs1,fs2],0)
        """
        if __runSimulate__(specFlt, clearPrev):
            raise ValueError(messError)

    def setSystemParams(self, val):
        """ change System Parameters.

        Samples:
            OLCase.setSystemParams({'fSmallX': '0.0002','bSimulateCCGen': '0'})
        """
        messError = '\nOLCase.setSystemParams(val)'
        if type(val) != dict:
            messError += '\n\tRequired val :  dict'
            messError += '\n\tFound        :  '+type(val).__name__
            raise ValueError(messError)
        #
        for k, v in val.items():
            if type(k) != str:
                messError += '\n\tval          : {' +toString(k)+' : '+toString(v)+'}'
                messError += '\n\tRequired key :  str'
                messError += '\n\tFound        :  '+type(k).__name__+' (%s)' % toString(k)
                raise ValueError(__mesParamSys__()+'\n'+messError)
            elif k not in __OLXOBJ_PARASYS__.keys():
                messError += '\n\tval          : {'+toString(k)+' : '+toString(v)+'}'
                messError += '\n\tRequired key : str in list '
                messError += '\n\tFound        : '+toString(k)
                raise ValueError(__mesParamSys__()+'\n'+messError)
            #
            vd = __OLXOBJ_PARASYS__[k]
            if vd:
                if vd[0] == '(float)':
                    v1 = __convert2Float__(v)
                    if type(v1)!=float:
                        messError += '\n\tval          : {'+toString(k)+' : '+toString(v)+'}'
                        messError += '\n\tRequired key : (float)'
                        messError += '\n\tFound        : '+toString(v)
                        raise ValueError(messError)
                elif vd[0] == '(int)':
                    v1 = __convert2Int__(v)
                    if type(v1)!=int:
                        messError += '\n\tval          : {'+toString(k)+' : '+toString(v)+'}'
                        messError += '\n\tRequired key : (int)'
                        messError += '\n\tFound        : '+toString(v)
                        raise ValueError(messError)
                elif type(vd[0]) == list:
                    try:
                        v1 = __convert2Int__(v)
                        if v1 not in vd[0]:
                            v1 = None
                    except:
                        v1 = None
                    if v1 is None:
                        messError += '\n\tval          : {'+toString(k)+' : '+toString(v)+'}'
                        messError += '\n\tRequired key : (int) in '+str(vd[0])
                        messError += '\n\tFound        : '+str(v)
                        raise ValueError(messError)
        #
        for k, v in val.items():
            params = __voidArray__()
            params[0] = cast(pointer(c_char_p(encode3(k+'\t'+str(v)))), c_void_p)
            params[1] = cast(pointer(c_char_p(encode3(""))), c_void_p)  # Terminator
            if OLXAPI_OK != SetData(HND_SYS, OlxAPIConst.SY_vParams, params):
                raise Exception(ErrorString())
        #
        if OLXAPI_OK != OlxAPI.PostData(HND_SYS):
            raise Exception(ErrorString())

    def tapLineTool(self, t0, tapSCAP=False, verbose=False):
        """ Find main sections of Line and sum impedance(Z0,Z1) and Length.
            All taps are ignored. Close switches are included.
            Branches out of service are ignored.

        Args:
            t0      : start TERMINAL/RLYGROUP/LINE/BUS/SWITCH/SERIESRC
                        Exception if:
                            1st Bus (of t0) is a TAP Bus.
                            t0 located on XFMR,XFMR3,SHIFTER
            tapSCAP : if True continue even if no tapbus (with SERIESRC)
            verbose : (obsolete) replace by OlxObj.setVerbose().

        return:
            res['mainLine']   = [[]] List of all TERMINALs in the main-line of method 1,2,3 in Help 8.9.
                                        METHOD 1: Enter the same name in the Name field of the line’s main segments.
                                        METHOD 2: Include these three characters [T] (or [t]) in the tap lines name.
                                        METHOD 3: Give the tap lines circuit IDs that contain letter T or t.
            res['remoteBus']  = []   List of remote BUSES on the mainLine: two for 2-terminal line and >2 for 3-terminal line.
            res['remoteRLG']  = []   List of all RLYGROUPs at the remote end on the mainLine.
            res['localBus']   = []   List of local buses on the mainLine.
            res['localRLG']   = []   List of all relay groups at the local end on the mainLine.
            res['allPath']    = [[]] List of all TERMINALs of all path branch.
            res['Z1']         = []   List of positive sequence Impedance of the mainLine.
            res['Z0']         = []   List of zero sequence Impedance of the mainLine.
            res['Length']     = []   List of sum length of the mainLine.
        """
        import OlxAPILib
        ty0 = type(t0)
        se = '\nOLCase.tapLineTool(t0)'
        if ty0 not in {TERMINAL, RLYGROUP, LINE, SWITCH, SERIESRC, BUS}:
            se += '\n\tRequired (t0) : TERMINAL,RLYGROUP,LINE,SWITCH,SERIESRC,BUS'
            se += '\n\tFound         : '+toString(t0)
            raise Exception(se+'\n\nUnable to continue.')
        if ty0 == BUS:
            res = {'mainLine': [], 'remoteBus': [], 'localBus': [], 'localRLG': [], 'remoteRLG': [], 'Z1': [], 'Z0': [], 'Length': [], 'allPath': []}
            for t1 in t0.TERMINAL:
                if t1.FLAG == 1 and type(t1.EQUIPMENT) in {LINE, SWITCH, SERIESRC}:
                    r1 = self.tapLineTool(t1)
                    for k, v in r1.items():
                        res[k].extend(v)
            return res
        #
        if ty0 in {LINE, SWITCH, SERIESRC}:
            ta = t0.TERMINAL
            b2 = ta[1].BUS1
            if b2.TAP == 0:
                t0 = ta[1]
            else:
                t0 = ta[0]
        elif ty0 == RLYGROUP:
            t0 = t0.TERMINAL
        e0 = t0.EQUIPMENT
        if type(e0) not in {LINE, SWITCH, SERIESRC}:
            se += '\n\tRequired (t0 located on) : LINE,SWITCH,SERIESRC'
            se += '\n\tFound         : '+toString(e0)
            raise Exception(se+'\n\nUnable to continue.')
        if t0.FLAG != 1:
            se += '\n\tt0 is out-of-service : '+toString(t0)
            raise Exception(se+'\n\nUnable to continue.')
        #
        ra = OlxAPILib.tapLineTool(t0.HANDLE, tapSCAP=tapSCAP, prt=__OLXOBJ_VERBOSE__)
        res = dict()
        res['mainLine'] = self.toOBJ(ra['mainLineHnd'])
        res['allPath'] = self.toOBJ(ra['allPathHnd'])
        res['localBus'] = self.toOBJ(ra['localBusHnd'])
        res['remoteBus'] = self.toOBJ(ra['remoteBusHnd'])
        res['remoteRLG'] = self.toOBJ(ra['remoteRLGHnd'])
        res['localRLG'] = self.toOBJ(ra['localRLGHnd'])
        res['Z1'] = ra['Z1']
        res['Z0'] = ra['Z0']
        res['Length'] = ra['Length']
        return res

    def toOBJ(self, handle):
        """ Convert handle => OBJECT (None if not found).

        Samples: (with 4,5 are handle of BUS)
            4       => BUS
            [4,5]   => [BUS, BUS]
            [[4,5]] => [[BUS, BUS]]
        """
        if type(handle) == list:
            return [self.toOBJ(h1) for h1 in handle]
        try:
            tc1 = EquipmentType(handle)
            return __getOBJ__(handle, tc=tc1)
        except:
            return None

    def __getattr__(self, name):  # internal
        return self.getData(name)

    def __initEmbed__(self):  # internal
        """ internal embeded mode. """
        if OlxAPI.ASPENOlxAPIDLL is None:
            return
        olrFile = OlxAPI.GetOlrFileName()
        self.__setattr__('olrFile', olrFile)
        #
        global __CURRENT_FILE_IDX__, __INDEX_SIMUL__, __COUNT_FAULT__, FltSimResult
        if olrFile:
            __CURRENT_FILE_IDX__ = abs(__CURRENT_FILE_IDX__)+1
        __INDEX_SIMUL__ += 1
        __COUNT_FAULT__ = 0
        FltSimResult.clear()
        #
        self.__setattr__('__currFileIdx__', __CURRENT_FILE_IDX__)
        self.__setattr__('__scope__', {'isFullNetWork': True})
        self.__setattr__('__BUS__', None)
        # UDF
        udfn = dict()
        try:
            udf = self.getUDFTemplate()
        except:
            udf = dict()
        for k, v in udf.items():
            udfn[k] = [v1[0] for v1 in v]
        for v1 in [['RLYDSG', 'RLYDS'], ['RLYDSP', 'RLYDS'], ['RLYOCG', 'RLYOC'], ['RLYOCP', 'RLYOC']]:
            try:
                udfn[v1[0]] = udfn[v1[1]]
            except:
                pass
        self.__setattr__('__UDF__', udfn)


OLCase = NETWORK()
OLCase.__initEmbed__()


class AREA:
    """ AREA Object. """

    def __init__(self, no):
        """ AREA Constructor by Area Number. """
        super().__setattr__('__no__', no)
        super().__setattr__('__currFileIdx__', __CURRENT_FILE_IDX__)

    @property
    def EXPORT(self):
        """ (float) Export MW. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return 0

    @property
    def GEN(self):
        """ (float) Generation MW. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return 0

    @property
    def LOAD(self):
        """ (float) Load MW. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return 0

    @property
    def NAME(self):
        """ (str) Name. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return OlxAPI.GetAreaName(self.__no__)

    @property
    def NO(self):
        """ (int) Number. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return self.__no__

    @property
    def SLACKBUS(self):
        """ (BUS) Slack Bus. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return None

    def equals(self, o2):
        """ (bool) Comparison of 2 AREA Object. """
        if type(o2) == AREA:
            if __check_currFileIdx1__():
                raise Exception(messError)
            return self.__no__ == o2.__no__
        return False

    def toString(self, option=0):
        """ (str) Text description/composed of Area Object. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        if option == 0:
            return "[AREA] %i '%s'" % (self.__no__, self.NAME)
        return '%i %s' % (self.__no__, self.NAME)


class BREAKER(DATAABSTRACT):
    """ BREAKER Rating Object. """

    def __init__(self, key=None, hnd=None):
        """ BREAKER Rating constructor (Exception if not found).

        Samples:
            BREAKER('{369fce04-353b-4c81-8e8e-9c4d97784206}') #GUID
            BREAKER("[BREAKER]  1E82A@ 6 'NEVADA' 132 kV")    #STR
            BREAKER(['NEVADA',132,'1E82A'])                   #BUS,Breaker ID
            BREAKER([6,'1E82A'])
        """
        hnd = __initOBJ__(BREAKER, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) Bus of Breaker. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def CPT1(self):
        """ (float) Contact parting time for group 1 (cycles). """
        return self.getData('CPT1')

    @property
    def CPT2(self):
        """ (float) Contact parting time for group 2 (cycles). """
        return self.getData('CPT2')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1-true; 2-false. """
        return self.getData('FLAG')

    @property
    def G1OUTAGES(self):
        """ [EQUIPMENT]*10 Protected equipment group 1 List of additional outage. """
        return self.getData('G1OUTAGES')

    @property
    def G2OUTAGES(self):
        """ [EQUIPMENT]*10 Protected equipment group 2 List of additional outage. """
        return self.getData('G2OUTAGES')

    @property
    def GROUPTYPE1(self):
        """ (int) Group 1 interrupting current: 1-max current; 0-group current. """
        return self.getData('GROUPTYPE1')

    @property
    def GROUPTYPE2(self):
        """ (int) Group 2 interrupting current: 1-max current; 0-group current. """
        return self.getData('GROUPTYPE2')

    @property
    def INTRATING(self):
        """ (float) Interrupting rating. """
        return self.getData('INTRATING')

    @property
    def INTRTIME(self):
        """ (float) Interrupting time (cycles). """
        return self.getData('INTRTIME')

    @property
    def K(self):
        """ (float) kV range factor. """
        return self.getData('K')

    @property
    def MRATING(self):
        """ (float) Momentary rating. """
        return self.getData('MRATING')

    @property
    def NACD(self):
        """ (float) No-ac-decay ratio. """
        return self.getData('NACD')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def NODERATE(self):
        """ (int) Do not derate in reclosing operation FLAG: 1-true; 0-false. """
        return self.getData('NODERATE')

    @property
    def OBJLST1(self):
        """ [EQUIPMENT]*10 Protected equipment group 1 List of equipment. """
        return self.getData('OBJLST1')

    @property
    def OBJLST2(self):
        """ [EQUIPMENT]*10 Protected equipment group 2 List of equipment. """
        return self.getData('OBJLST2')

    @property
    def OPKV(self):
        """ (float) Operating kV. """
        return self.getData('OPKV')

    @property
    def OPS1(self):
        """ (int) Total operations for group 1. """
        return self.getData('OPS1')

    @property
    def OPS2(self):
        """ (int) Total operations for group 2. """
        return self.getData('OPS2')

    @property
    def RATEDKV(self):
        """ (float) Max design kV. """
        return self.getData('RATEDKV')

    @property
    def RATINGTYPE(self):
        """ (int) Rating type: 0- symmetrical current basis;1- total current basis; 2- IEC. """
        return self.getData('RATINGTYPE')

    @property
    def RECLOSE1(self):
        """ [float]*3 Reclosing intervals for group 1 (s). """
        return self.getData('RECLOSE1')

    @property
    def RECLOSE2(self):
        """ [float]*3 Reclosing intervals for group 2 (s). """
        return self.getData('RECLOSE2')


class BUS(DATAABSTRACT):
    """ BUS Object. """

    def __init__(self, key=None, hnd=None):
        """ BUS constructor (Exception if not found).

        Samples:
            BUS('_{220879BF-1297-4314-B10F-53765B11E02C}') #GUID
            BUS("[BUS] 28 'ARIZONA' 132 kV")               #STR
            BUS(28)                                        #Bus Number
            BUS(['arizona',132])                           #Name,kV
        """
        hnd = __initOBJ__(BUS, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def ANGLEP(self):
        """ (float) Voltage angle (degree) from a power flow solution. """
        return self.getData('ANGLEP')

    @property
    def AREANO(self):
        """ (int) Area Number. """
        return self.getData('AREANO')

    @property
    def AREA(self):
        """ (AREA) Area Object. """
        return AREA(self.getData('AREANO'))

    @property
    def BREAKER(self):
        """ [BREAKER] List of Breakers Rating connected to BUS. """
        return [BREAKER(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'BREAKER')]

    @property
    def BUS(self):
        """ return BUS-self. """
        return self

    @property
    def CCGEN(self):
        """ (CCGEN) Voltage Controlled Current Sources connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'CCGEN')
        try:
            return CCGEN(hnd=g[0])
        except:
            return None

    @property
    def DCLINE2(self):
        """ [DCLINE2] List of DC Lines connected to BUS. """
        return [DCLINE2(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'DCLINE2')]

    @property
    def GEN(self):
        """ (GEN) Generator connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'GEN')
        try:
            return GEN(hnd=g[0])
        except:
            return None

    @property
    def GENUNIT(self):
        """ [GENUNIT] List of Generator Units connected to BUS. """
        return [GENUNIT(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'GENUNIT')]

    @property
    def GENW3(self):
        """ (GENW3) Type-3 Wind Plants connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'GENW3')
        try:
            return GENW3(hnd=g[0])
        except:
            return None

    @property
    def GENW4(self):
        """ (GENW4) Converter-Interfaced Resources connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'GENW4')
        try:
            return GENW4(hnd=g[0])
        except:
            return None

    @property
    def KV(self):
        """ (float) Voltage nominal. """
        return self.getData('KV')

    @property
    def KVP(self):
        """ (float) Voltage magnitude (kV) from a power flow solution. """
        return self.getData('KVP')

    @property
    def LINE(self):
        """ [LINE] List of AC Transmission Lines connected to BUS. """
        return [LINE(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'LINE')]

    @property
    def LOAD(self):
        """ (LOAD) Load connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'LOAD')
        try:
            return LOAD(hnd=g[0])
        except:
            return None

    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List of Load Units connected to BUS. """
        return [LOADUNIT(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'LOADUNIT')]

    @property
    def LOCATION(self):
        """ (str) Location. """
        return self.getData('LOCATION')

    @property
    def MEMO(self):
        """ (str) Memo. """
        return self.getData('MEMO')

    @property
    def MIDPOINT(self):
        """ (int) 3-winding Transformer mid-point FLAG: 0-no; 1-yes. """
        return self.getData('MIDPOINT')

    @property
    def NAME(self):
        """ (str) Bus Name. """
        return self.getData('NAME')

    @property
    def NO(self):
        """ (int) Bus Number. """
        return self.getData('NO')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of Relay Groups connected to BUS. """
        return [RLYGROUP(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'RLYGROUP')]

    @property
    def SERIESRC(self):
        """ [SERIESRC] List of Series capacitor/reactors connected to BUS. """
        return [SERIESRC(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'SERIESRC')]

    @property
    def SHIFTER(self):
        """ [SHIFTER] List of Phase Shifters connected to BUS. """
        return [SHIFTER(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'SHIFTER')]

    @property
    def SHUNT(self):
        """ (SHUNT) Shunt connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'SHUNT')
        try:
            return SHUNT(hnd=g[0])
        except:
            return None

    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List of Shunt Units connected to BUS. """
        return [SHUNTUNIT(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'SHUNTUNIT')]

    @property
    def SLACK(self):
        """ (int) System slack bus FLAG: 1-yes; 0-no. """
        return self.getData('SLACK')

    @property
    def SPCX(self):
        """ (float) Bus state plane coordinate - X. """
        return self.getData('SPCX')

    @property
    def SPCY(self):
        """ (float) Bus state plane coordinate - Y. """
        return self.getData('SPCY')

    @property
    def SUBGRP(self):
        """ (int) Bus Substation group. """
        return self.getData('SUBGRP')

    @property
    def SVD(self):
        """ (SVD) Switched Shunt connected to BUS (None if not found). """
        g = __getBusEquipmentHnd__(self, 'SVD')
        try:
            return SVD(hnd=g[0])
        except:
            return None

    @property
    def SWITCH(self):
        """ [SWITCH] List of Switches connected to BUS. """
        return [SWITCH(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'SWITCH')]

    @property
    def TAP(self):
        """ (int) Tap bus FLAG: 0-No; 1-Tap bus; 3-Tap bus of 3-terminal line. """
        return self.getData('TAP')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs connected to BUS. """
        return [TERMINAL(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'TERMINAL')]

    @property
    def VISIBLE(self):
        """ (int) Bus hide FLAG: 1-Visible; 2-Hidden; 0-Not yet placed. """
        return self.getData('VISIBLE')

    @property
    def XFMR(self):
        """ [XFMR] List of 2-Windings Transformers connected to BUS. """
        return [XFMR(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'XFMR')]

    @property
    def XFMR3(self):
        """ [XFMR3] List of 3-Windings Transformers connected to BUS. """
        return [XFMR3(hnd=h1) for h1 in __getBusEquipmentHnd__(self, 'XFMR3')]

    @property
    def ZONE(self):
        """ (ZONE) Zone Object. """
        return ZONE(self.getData('ZONENO'))

    @property
    def ZONENO(self):
        """ (int) Zone Number. """
        return self.getData('ZONENO')
    #

    def findBusNeibor(self, tiers, ignoreTapBus=1):
        """ Find list of Buses Neibor to BUS.

        Args:
            tiers        : (int) Number of tiers around this BUS to find
            ignoreTapBus : 0/1 option ignore Tap Bus

        return:
            list of Buses
        """
        if __check_currFileIdx__(self):
            raise Exception(messError)
        res = [self]
        bs1, bs2 = set(), set()
        bs2.add(self.__hnd__)
        for i in range(tiers):
            n1 = len(res)
            res = __findBusNeibor__(res, bs1, bs2, ignoreTapBus)
            if n1 == len(res):
                break
        return res

    def terminalTo(self, b2, sObj=None, CID=None):
        """ [TERMINAL] List of TERMINALs from BUS to BUS b2.

        Samples:
            b1.terminalTo(b2)          # list all TERMINALs from BUS b1 to BUS b2
            b1.terminalTo(b2,sObj)     # list all TERMINALs from BUS b1 to BUS b2 of sObj
            b1.terminalTo(b2,sObj,CID) # (TERMINAL) from BUS b1 to BUS b2 of sObj and CID
        """
        if type(b2) != BUS:
            se = '\nBUS.terminalTo(b2) ValueError of b2'
            se += '\n\tRequired           : (BUS)'
            se += '\n\t'+__getErrValue__(BUS, b2)
            raise ValueError(se)
        if (sObj is not None) and (type(sObj) != str or sObj.upper() not in __OLXOBJ_EQUIPMENTL__):
            se = '\nBUS.terminalTo(b2,sObj) ValueError of sObj'
            se += "\n\tRequired           : (str) in ['XFMR3','XFMR','SHIFTER','LINE','DCLINE2','SERIESRC','SWITCH']"
            se += '\n\t'+__getErrValue__(str, sObj)
            raise ValueError(se)
        if (CID is not None) and type(CID) != str:
            se = '\nBUS.terminalTo(b2,sObj,CID) ValueError of CID'
            se += '\n\tRequired           : (str)'
            se += '\n\t'+__getErrValue__(str, CID)
            raise ValueError(se)
        if __check_currFileIdx__(self):
            raise Exception(messError)
        if __check_currFileIdx__(b2):
            raise Exception(messError)
        #
        if sObj != None and CID != None:
            return OLCase.findTERMINAL(self, b2, sObj.upper(), CID)
        #
        hnd1, hnd2 = self.__hnd__, b2.__hnd__
        val1 = c_int(0)
        res = []
        while OLXAPI_OK == GetBusEquipment(hnd1, c_int(TC_BRANCH), byref(val1)):
            val2 = __getDatai__(val1.value, OlxAPIConst.BR_nBus2Hnd)
            val3 = __getDatai__(val1.value, OlxAPIConst.BR_nBus3Hnd)
            t1 = TERMINAL(hnd=val1.value)
            e1 = t1.EQUIPMENT
            if val2 == hnd2 or val3 == hnd2:
                if (sObj is None or type(e1).__name__ == sObj.upper()) and (CID is None or e1.CID == CID):
                    res.append(t1)
        return res


class CCGEN(DATAABSTRACT):
    """ CCGEN Voltage controlled current source Object. """

    def __init__(self, key=None, hnd=None):
        """ CCGEN constructor (Exception if not found).

        Samples:
            CCGEN('_{39238FAA-5D98-4F17-A669-09E6BD2E20A0}') #GUID
            CCGEN("[CCGENUNIT] 'BUS5' 13 kV")                #STR
            CCGEN(['CLAYTOR',132])                           #Name,kV
        """
        hnd = __initOBJ__(CCGEN, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def A(self):
        """ [float]*10 Angle. """
        return self.getData('A')

    @property
    def BLOCKPHASE(self):
        """ (int) Number block on phase. """
        return self.getData('BLOCKPHASE')

    @property
    def BUS(self):
        """ (BUS) BUS that CCGEN located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1-true; 2-false. """
        return self.getData('FLAG')

    @property
    def I(self):
        """ [float]*10 Current. """
        return self.getData('I')

    @property
    def MVARATE(self):
        """ (float) MVA rating. """
        return self.getData('MVARATE')

    @property
    def V(self):
        """ [float]*10 Voltage. """
        return self.getData('V')

    @property
    def VLOC(self):
        """ (int) Voltage measurement location FLAG: 0-Device terminal; 1-Network side of transformer. """
        return self.getData('VLOC')

    @property
    def VMAXMUL(self):
        """ (float) Maximum voltage limit, in pu. """
        return self.getData('VMAXMUL')

    @property
    def VMIN(self):
        """ (float) Minimum voltage limit, in pu. """
        return self.getData('VMIN')


class DCLINE2(DATAABSTRACT):
    """ DC Transmission Line Object. """

    def __init__(self, key=None, hnd=None):
        """ DC Transmission Line constructor (Exception if not found).

        Samples:
            DCLINE2('_{2BB0E750-FC57-497E-B878-081D74319DE5}')       #GUID
            DCLINE2("[DCLINE2] 5 'FIELDALE' 132 kV-'BUS2' 132 kV 1") #STR
            DCLINE2([['FIELDALE',132],['BUS2',132],'1'])             #bus1,bus2,CID
        """
        hnd = __initOBJ__(DCLINE2, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def ANGMAX(self):
        """ [float]*2 Angle max. """
        return self.getData('ANGMAX')

    @property
    def ANGMIN(self):
        """ [float]*2 Angle min. """
        return self.getData('ANGMIN')

    @property
    def BRCODE(self):
        """ (str) BRCODE = 'DC'. """
        return 'DC'

    @property
    def BRIDGE(self):
        """ [int]*2 No. of bridges. """
        return self.getData('BRIDGE')

    @property
    def BUS(self):
        """ [BUS] List of Buses of DC Line. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def CID(self):
        """ (string) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (string) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (string) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (DCLINE2) return DCLINE2-self. """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def LINER(self):
        """ (float) Resistance, in Ohm. """
        return self.getData('LINER')

    @property
    def MARGIN(self):
        """ (float) Margin. """
        return self.getData('MARGIN')

    @property
    def MODE(self):
        """ (int) Control mode. """
        return self.getData('MODE')

    @property
    def MVA(self):
        """ [float]*2 MVA Rating. """
        return self.getData('MVA')

    @property
    def NAME(self):
        """ (string) DC Line Name. """
        return self.getData('NAME')

    @property
    def NOMKV(self):
        """ [float]*2 Nominal kV on DC side. """
        return self.getData('NOMKV')

    @property
    def R(self):
        """ [float]*2 Transformer Resistance, in pu. """
        return self.getData('R')

    @property
    def TAP(self):
        """ [float]*2 Tap. """
        return self.getData('TAP')

    @property
    def TAPMAX(self):
        """ [float]*2 Tap max. """
        return self.getData('TAPMAX')

    @property
    def TAPMIN(self):
        """ [float]*2 Tap min. """
        return self.getData('TAPMIN')

    @property
    def TAPSTEP(self):
        """ [float]*2 Tap step size. """
        return self.getData('TAPSTEP')

    @property
    def TARGET(self):
        """ (float) Target MW. """
        return self.getData('TARGET')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def TIE(self):
        """ (int) Meteted FLAG: 1- at BUS1; 2-at BUS2; 0-line is in a single area. """
        return self.getData('TIE')

    @property
    def VSCHED(self):
        """ (float) Schedule DC volts, in pu. """
        return self.getData('VSCHED')

    @property
    def X(self):
        """ [float]*2 Transformer Reactance X, in pu. """
        return self.getData('X')


class FUSE(RELAYABSTRACT):
    """ FUSE Object. """

    def __init__(self, key=None, hnd=None):
        """ FUSE constructor (Exception if not found).

        Samples:
            FUSE('{d67118b8-bb8c-4f37-9765-bf3c49cb2cba}')
            FUSE("[FUSE]  NV Fuse@6 'NEVADA' 132 kV-4 'TENNESSEE' 132 kV 1 P")
            FUSE([b1,b2,CID,BRCODE,ID])
                b1      BUS|str|int|[str,f_i] with f_i: float or int
                b2      BUS|str|int|[str,f_i] with f_i: float or int
                CID     (str) circuit ID
                BRCODE  (str) in ['X','T','P','L','DC','S','W']
                ID      (str) Relay ID
        """
        hnd = __initOBJ__(FUSE, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def FUSECURDIV(self):
        """ (float) Current divider. """
        return self.getData('FUSECURDIV')

    @property
    def FUSECVE(self):
        """ (int) Compute time using FLAG: 1- minimum melt; 2- Total clear. """
        return self.getData('FUSECVE')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def RATING(self):
        """ (float) Rating. """
        return self.getData('RATING')

    @property
    def TIMEMULT(self):
        """ (float) Time multiplier. """
        return self.getData('TIMEMULT')

    @property
    def TYPE(self):
        """ (str) Type. """
        return self.getData('TYPE')


class GEN(DATAABSTRACT):
    """ Generator Object. """

    def __init__(self, key=None, hnd=None):
        """ Generator constructor (Exception if not found).

        Samples:
            GEN("_{82BC751C-7988-49CF-874A-57C9CEB1ECA3}") #GUID
            GEN("[GENERATOR] 2 'CLAYTOR' 132 kV")          #STR
            GEN(['CLAYTOR',132])                           #Name,kV
        """
        hnd = __initOBJ__(GEN, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that Generator located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def CNTBUS(self):
        """ (BUS) Controlled Bus. """
        return BUS(hnd=self.getData('CNTBUS').__hnd__)

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def GENUNIT(self):
        """ [GENUNIT] List of Generator Units. """
        return [GENUNIT(hnd=g1.__hnd__) for g1 in self.BUS.GENUNIT]

    @property
    def ILIMIT1(self):
        """ (float) Current limit 1. """
        return self.getData('ILIMIT1')

    @property
    def ILIMIT2(self):
        """ (float) Current limit 2. """
        return self.getData('ILIMIT2')

    @property
    def MVARATE(self):
        """ (float) Rating MVA. """
        return sum([g1.MVARATE for g1 in self.BUS.GENUNIT])

    @property
    def PGEN(self):
        """ (float) MW (load flow solution). """
        return self.getData('PGEN')

    @property
    def PMAX(self):
        """ (float) Max MW. """
        return sum([g1.PMAX for g1 in self.BUS.GENUNIT])

    @property
    def PMIN(self):
        """ (float) Min MW. """
        return sum([g1.PMIN for g1 in self.BUS.GENUNIT])

    @property
    def QMAX(self):
        """ (float) Max MVAR. """
        return sum([g1.QMAX for g1 in self.BUS.GENUNIT])

    @property
    def QMIN(self):
        """ (float) Min MVAR. """
        return sum([g1.QMIN for g1 in self.BUS.GENUNIT])

    @property
    def QGEN(self):
        """ (float) MVAR (load flow solution). """
        return self.getData('QGEN')

    @property
    def REFANGLE(self):
        """ (float) Reference angle. """
        return self.getData('REFANGLE')

    @property
    def REFV(self):
        """ (float) Internal voltage source magnitude, in pu. """
        return self.getData('REFV')

    @property
    def REG(self):
        """ (int) Regulation FLAG: 1- PQ; 0- PV. """
        return self.getData('REG')

    @property
    def SCHEDP(self):
        """ (float) Scheduled P. """
        return self.getData('SCHEDP')

    @property
    def SCHEDQ(self):
        """ (float) Scheduled Q. """
        return self.getData('SCHEDQ')

    @property
    def SCHEDV(self):
        """ (float) Scheduled V. """
        return self.getData('SCHEDV')


class GENUNIT(DATAABSTRACT):
    """ Generator Unit Object. """

    def __init__(self, key=None, hnd=None):
        """ Generator Unit constructor (Exception if not found).

        Samples:
            GENUNIT('_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}') #GUID
            GENUNIT("[GENUNIT]  1@2 'CLAYTOR' 132 kV")         #STR
            GENUNIT(['CLAYTOR',132,'1'])                       #Name,kV,CID
        """
        hnd = __initOBJ__(GENUNIT, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that Generator Unit located on. """
        return BUS(hnd=self.GEN.BUS.__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def GEN(self):
        """ (GEN) Generator that Generator Unit located on. """
        return GEN(hnd=self.getData('GEN').__hnd__)

    @property
    def MVARATE(self):
        """ (float) Rating MVA. """
        return self.getData('MVARATE')

    @property
    def PMAX(self):
        """ (float) Max MW. """
        return self.getData('PMAX')

    @property
    def PMIN(self):
        """ (float) Min MW. """
        return self.getData('PMIN')

    @property
    def QMAX(self):
        """ (float) Max MVAR. """
        return self.getData('QMAX')

    @property
    def QMIN(self):
        """ (float) Min MVAR. """
        return self.getData('QMIN')

    @property
    def R(self):
        """ [float]*5 Resistances [subtransient, synchronous, transient, negative sequence, zero sequence]. """
        return self.getData('R')

    @property
    def RG(self):
        """ (float) Grounding Resistance, in Ohm (do not multiply by 3). """
        return self.getData('RG')

    @property
    def SCHEDP(self):
        """ (float) Scheduled P. """
        return self.getData('SCHEDP')

    @property
    def SCHEDQ(self):
        """ (float) Scheduled Q. """
        return self.getData('SCHEDQ')

    @property
    def X(self):
        """ [float]*5 Reactances [subtransient, synchronous, transient, negative sequence, zero sequence]. """
        return self.getData('X')

    @property
    def XG(self):
        """ (float) Grounding Reactance Xg, in Ohm (do not multiply by 3). """
        return self.getData('XG')


class GENW3(DATAABSTRACT):
    """ Type-3 Wind Plant Object. """

    def __init__(self, key=None, hnd=None):
        """ Type-3 Wind Plant constructor (Exception if not found).

        Samples:
            GENW3('_{F74E7DEC-C835-406A-B62C-B4DEA4D7C764}') #GUID
            GENW3("[GENW3] 4 'TENNESSEE' 132 kV")            #STR
            GENW3(['TENNESSEE', 132])                        #Name,kV
        """
        hnd = __initOBJ__(GENW3, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that GENW3 located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def CBAR(self):
        """ (int) Crowbarred FLAG: 1-yes; 0-no. """
        return self.getData('CBAR')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def FLRZ(self):
        """ (float) Filter X, in pu. """
        return self.getData('FLRZ')

    @property
    def IMAXG(self):
        """ (float) Grid side limit, in pu. """
        return self.getData('IMAXG')

    @property
    def IMAXR(self):
        """ (float) Rotor side limit, in pu. """
        return self.getData('IMAXR')

    @property
    def LLR(self):
        """ (float) Rotor leakage L, in pu. """
        return self.getData('LLR')

    @property
    def LLS(self):
        """ (float) Stator leakage L, in pu. """
        return self.getData('LLS')

    @property
    def LM(self):
        """ (float) Mutual L, in pu. """
        return self.getData('LM')

    @property
    def MVA(self):
        """ (float) MVA Unit rated. """
        return self.getData('MVA')

    @property
    def MW(self):
        """ (float) Unit MW generation. """
        return self.getData('MW')

    @property
    def MWR(self):
        """ (float) MW unit rated. """
        return self.getData('MWR')

    @property
    def RR(self):
        """ (float) Rotor Resistance, in pu. """
        return self.getData('RR')

    @property
    def RS(self):
        """ (float) Stator Resistance, in pu. """
        return self.getData('RS')

    @property
    def SLIP(self):
        """ (float) Slip at rated kW. """
        return self.getData('SLIP')

    @property
    def UNITS(self):
        """ (int) Number of Units. """
        return self.getData('UNITS')

    @property
    def VMAX(self):
        """ (float) Maximum voltage limit, in pu. """
        return self.getData('VMAX')

    @property
    def VMIN(self):
        """ (float) Minimum voltage limit, in pu. """
        return self.getData('VMIN')


class GENW4(DATAABSTRACT):
    """ GENW4 Converter-Interfaced Resource Object. """

    def __init__(self, key=None, hnd=None):
        """ GENW4 constructor (Exception if not found).

        Samples:
            GENW4('_{9E828F29-0CAF-4992-90AA-5A04D5F01463}') #GUID
            GENW4("[GENW4] 4 'TENNESSEE' 132 kV")            #STR
            GENW4(['TENNESSEE', 132])                        #Name,kV
        """
        hnd = __initOBJ__(GENW4, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that GENW4 located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def CTRLMETHOD(self):
        """ (int) Control method. """
        return self.getData('CTRLMETHOD')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def MAXI(self):
        """ (float) Max current. """
        return self.getData('MAXI')

    @property
    def MAXILOW(self):
        """ (float) Max current reduce to. """
        return self.getData('MAXILOW')

    @property
    def MVA(self):
        """ (float) Unit MVA rating. """
        return self.getData('MVA')

    @property
    def MVAR(self):
        """ (float) Unit MVAR. """
        return self.getData('MVAR')

    @property
    def MW(self):
        """ (float) Unit MW generation or consumption. """
        return self.getData('MW')

    @property
    def SLOPE(self):
        """ (float) Slope of +seq. """
        return self.getData('SLOPE')

    @property
    def SLOPENEG(self):
        """ (float) Slope of -seq. """
        return self.getData('SLOPENEG')

    @property
    def UNITS(self):
        """ (int) Number of Units. """
        return self.getData('UNITS')

    @property
    def VLOW(self):
        """ (float) When +seq(pu)>. """
        return self.getData('VLOW')

    @property
    def VMAX(self):
        """ (float) Maximum voltage limit. """
        return self.getData('VMAX')

    @property
    def VMIN(self):
        """ (float) Minimum voltage limit. """
        return self.getData('VMIN')


class LINE(DATAABSTRACT):
    """ AC Transmission Line Object. """

    def __init__(self, key=None, hnd=None):
        """ AC Transmission Line constructor (Exception if not found).

        Samples:
            LINE('_{D87FBEC5-C4F9-47FE-96D3-8F24CCB101FF}')       #GUID
            LINE("[LINE] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") #STR
            LINE([['CLAYTOR', 132],['NEVADA', 132],'1'])          #bus1,bus2,CID
        """
        hnd = __initOBJ__(LINE, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def B1(self):
        """ (float) +seq Susceptance B1 (at Bus1), in pu. """
        return self.getData('B1')

    @property
    def B10(self):
        """ (float) zero Susceptance B10 (at Bus1), in pu. """
        return self.getData('B10')

    @property
    def B2(self):
        """ (float) +seq Susceptance B2 (at Bus2), in pu. """
        return self.getData('B2')

    @property
    def B20(self):
        """ (float) zero Susceptance B20 (at Bus2), in pu. """
        return self.getData('B20')

    @property
    def BRCODE(self):
        """ (str) BRCODE='L'. """
        return 'L'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (LINE) return LINE-self. """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def G1(self):
        """ (float) +seq Conductance G1 (at Bus1), in pu. """
        return self.getData('G1')

    @property
    def G10(self):
        """ (float) zero Conductance G10 (at Bus1), in pu. """
        return self.getData('G10')

    @property
    def G2(self):
        """ (float) +seq Conductance G2 (at Bus2), in pu. """
        return self.getData('G2')

    @property
    def G20(self):
        """ (float) zero Conductance G20 (at Bus2), in pu. """
        return self.getData('G20')

    @property
    def I2T(self):
        """ (float) I^2T rating in ampere^2 Sec. """
        return self.getData('I2T')

    @property
    def LN(self):
        """ (float) Line length. """
        return self.getData('LN')

    @property
    def MULINE(self):
        """ [MULINE] List of Mutual Coupling Pairs of LINE. """
        return self.getData('MULINE')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def R(self):
        """ (float) +seq Resistance R, in pu. """
        return self.getData('R')

    @property
    def R0(self):
        """ (float) zero Resistance Ro, in pu. """
        return self.getData('R0')

    @property
    def RATG(self):
        """ [float]*4 Ratings, in [A]. """
        return self.getData('RATG')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2]

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group 1, at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group 2, at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs of LINE. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def TIE(self):
        """ (int) Meteted FLAG: 1- at Bus1; 2-at Bus2; 0-line is in a single area. """
        return self.getData('TIE')

    @property
    def TYPE(self):
        """ (str) Type. """
        return self.getData('TYPE')

    @property
    def UNIT(self):
        """ (str) Length unit in [ft,kt,mi,m,km]. """
        return self.getData('UNIT')

    @property
    def X(self):
        """ (float) +seq Reactance X, in pu. """
        return self.getData('X')

    @property
    def X0(self):
        """ (float) zero Reactance Xo, in pu. """
        return self.getData('X0')


class LOAD(DATAABSTRACT):
    """ LOAD Object. """

    def __init__(self, key=None, hnd=None):
        """ LOAD constructor (Exception if not found).

        Samples:
            LOAD('_{5B7BD740-AAA7-4A34-98BC-CC34B795299D}') #GUID
            LOAD("[LOAD] 17 'WASHINGTON' 33 kV")            #STR
            LOAD(['WASHINGTON', 33])                        #Name,kV
        """
        hnd = __initOBJ__(LOAD, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that LOAD located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def LOADUNIT(self):
        """ [LOADUNIT] List of Load Units. """
        return [LOADUNIT(hnd=ti.__hnd__) for ti in self.BUS.LOADUNIT]

    @property
    def P(self):
        """ (float) Total load MW (load flow solution). """
        return self.getData('P')

    @property
    def Q(self):
        """ (float) Total load MVAR (load flow solution). """
        return self.getData('Q')

    @property
    def UNGROUNDED(self):
        """ (int) UnGrounded FLAG: 1-UnGrounded ; 0-Grounded. """
        return self.getData('UNGROUNDED')


class LOADUNIT(DATAABSTRACT):
    """ Load Unit Object. """

    def __init__(self, key=None, hnd=None):
        """ Load Unit constructor (Exception if not found).

        Samples:
            LOADUNIT('_{68A5C3B4-29FC-41DA-86DD-9412E271E186}') #GUID
            LOADUNIT("[LOADUNIT]  1@17 'WASHINGTON' 33 kV")     #STR
            LOADUNIT(['WASHINGTON', 33,'1'])                    #Name,kV,CID
        """
        hnd = __initOBJ__(LOADUNIT, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that LOADUNIT located on. """
        return BUS(hnd=self.LOAD.BUS.__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1-active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def LOAD(self):
        """ (LOAD) LOAD that Load Unit located on. """
        return LOAD(hnd=self.getData('LOAD').__hnd__)

    @property
    def MVAR(self):
        """ [float]*3 MVARs: [const. P, const. I, const. Z]. """
        return self.getData('MVAR')

    @property
    def MW(self):
        """ [float]*3 MWs: [const. P, const. I, const. Z]. """
        return self.getData('MW')

    @property
    def P(self):
        """ (float) MW (load flow solution). """
        return self.getData('P')

    @property
    def Q(self):
        """ (float) MVAR (load flow solution). """
        return self.getData('Q')


class MULINE(DATAABSTRACT):
    """ Mutual Coupling Pair Object. """

    def __init__(self, key=None, hnd=None):
        """ Mutual Coupling Pair constructor (Exception if not found).

        Samples:
            MULINE('{03d127b7-2363-4642-b70e-0fb6e9636813}')   #GUID
            MULINE("[MUPAIR] 1 'GLEN LYN' 132 kV-2 'CLAYTOR' 132 kV 1|1 'GLEN LYN' 132 kV-2 'CLAYTOR' 132 kV 2") #STR
            MULINE([l1,l2])     #l1,l2 = STR/GUID/LINE
        """
        hnd = __initOBJ__(MULINE, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)
        l1,l2 = self.LINE1, self.LINE2
        self.__paramEx__['MULINEVAL'] = {'ORIENTLINE1': [l1.BUS1, l1.BUS2, l1.CID], 'ORIENTLINE2': [l2.BUS1, l2.BUS2, l2.CID]}

    @property
    def FROM1(self):
        """ [float]*5 From percent (Line1). """
        return self.getData('FROM1')

    @property
    def FROM2(self):
        """ [float]*5 From percent (Line2). """
        return self.getData('FROM2')

    @property
    def LINE1(self):
        """ (LINE) Line1 of MULINE. """
        try:
            return LINE(hnd=self.getData('LINE1').__hnd__)
        except:
            return None

    @property
    def LINE2(self):
        """ (LINE) Line2 of MULINE. """
        try:
            return LINE(hnd=self.getData('LINE2').__hnd__)
        except:
            return None

    @property
    def R(self):
        """ [float]*5 Resistance R, in pu. """
        return self.getData('R')

    @property
    def TO1(self):
        """ [float]*5 To percent 1 (Line1). """
        return self.getData('TO1')

    @property
    def TO2(self):
        """ [float]*5 To percent 2 (Line2). """
        return self.getData('TO2')

    @property
    def X(self):
        """ [float]*5 Mutual Reactance X, in pu. """
        return self.getData('X')


class RECLSR(RELAYABSTRACT):
    """ Recloser Object. """

    def __init__(self, key=None, hnd=None):
        """ Recloser constructor (Exception if not found).

        Samples:
            RECLSR('{9e8f6488-6bb0-4e3a-9f4f-ac878d182f35}')               #GUID
            RECLSR("[RECLSRP]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            RECLSR([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RECLSR, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BYTADD(self):
        """ (int) Time adder modifies. """
        return self.getData('BYTADD')

    @property
    def BYTMULT(self):
        """ (int) Time multiplier modifies. """
        return self.getData('BYTMULT')

    @property
    def FASTOPS(self):
        """ (int) Number of fast operations. """
        return self.getData('FASTOPS')

    @property
    def INTRPTIME(self):
        """ (float) Interrupting time (s). """
        return self.getData('INTRPTIME')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def RATING(self):
        """ (float) Rating (Rated momentary amps). """
        return self.getData('RATING')

    @property
    def RECLOSE1(self):
        """ (float) Reclosing interval 1. """
        return self.getData('RECLOSE1')

    @property
    def RECLOSE2(self):
        """ (float) Reclosing interval 2. """
        return self.getData('RECLOSE2')

    @property
    def RECLOSE3(self):
        """ (float) Reclosing interval 3. """
        return self.getData('RECLOSE3')

    @property
    def TOTALOPS(self):
        """ (int) Total operations to locked out. """
        return self.getData('TOTALOPS')

    @property
    def GR_BYFAST(self):
        """ (str) Recloser-Ground Curve in use FLAG: 1- Fast curve; 0- Slow curve. """
        return self.getData('GR_BYFAST')

    @property
    def GR_FASTTYPE(self):
        """ (str) Recloser-Ground fast curve. """
        return self.getData('GR_FASTTYPE')

    @property
    def GR_FLAG(self):
        """ (int) Recloser-Ground In service FLAG: 1- active; 2- out-of-service. """
        return self.getData('GR_FLAG')

    @property
    def GR_INST(self):
        """ (float) Recloser-Ground high current trip. """
        return self.getData('GR_INST')

    @property
    def GR_INSTDELAY(self):
        """ (float) Recloser-Ground high current trip delay. """
        return self.getData('GR_INSTDELAY')

    @property
    def GR_MINTRIPF(self):
        """ (float) Recloser-Ground fast curve pickup. """
        return self.getData('GR_MINTRIPF')

    @property
    def GR_MINTIMEF(self):
        """ (float) Recloser-Ground fast curve minimum time. """
        return self.getData('GR_MINTIMEF')

    @property
    def GR_MINTIMES(self):
        """ (float) Recloser-Ground slow curve minimum time. """
        return self.getData('GR_MINTIMES')

    @property
    def GR_MINTRIPS(self):
        """ (float) Recloser-Ground slow curve pickup. """
        return self.getData('GR_MINTRIPS')

    @property
    def GR_SLOWTYPE(self):
        """ (str) Recloser-Ground slow curve """
        return self.getData('GR_SLOWTYPE')

    @property
    def GR_TIMEADDF(self):
        """ (float) Recloser-Ground fast curve time adder. """
        return self.getData('GR_TIMEADDF')

    @property
    def GR_TIMEADDS(self):
        """ (float) Recloser-Ground slow curve time adder. """
        return self.getData('GR_TIMEADDS')

    @property
    def GR_TIMEMULTF(self):
        """ (float) Recloser-Ground fast curve time multiplier. """
        return self.getData('GR_TIMEMULTF')

    @property
    def GR_TIMEMULTS(self):
        """ (float) Recloser-Ground slow curve time multiplier. """
        return self.getData('GR_TIMEMULTS')

    @property
    def PH_BYFAST(self):
        """ (str) Recloser-Phase Curve in use FLAG: 1- Fast curve; 0- Slow curve. """
        return self.getData('PH_BYFAST')

    @property
    def PH_FASTTYPE(self):
        """ (str) Recloser-Phase fast curve. """
        return self.getData('PH_FASTTYPE')

    @property
    def PH_FLAG(self):
        """ (int) Recloser-Phase In service FLAG: 1- active; 2- out-of-service. """
        return self.getData('PH_FLAG')

    @property
    def PH_INST(self):
        """ (float) Recloser-Phase high current trip. """
        return self.getData('PH_INST')

    @property
    def PH_INSTDELAY(self):
        """ (float) Recloser-Phase high current trip delay. """
        return self.getData('PH_INSTDELAY')

    @property
    def PH_MINTIMEF(self):
        """ (float) Recloser-Phase fast curve minimum time. """
        return self.getData('PH_MINTIMEF')

    @property
    def PH_MINTIMES(self):
        """ (float) Recloser-Phase slow curve minimum time. """
        return self.getData('PH_MINTIMES')

    @property
    def PH_MINTRIPF(self):
        """ (float) Recloser-Phase fast curve pickup. """
        return self.getData('PH_MINTRIPF')

    @property
    def PH_MINTRIPS(self):
        """ (float) Recloser-Phase slow curve pickup. """
        return self.getData('PH_MINTRIPS')

    @property
    def PH_SLOWTYPE(self):
        """ (str) Recloser-Phase slow curve ."""
        return self.getData('PH_SLOWTYPE')

    @property
    def PH_TIMEADDF(self):
        """ (float) Recloser-Phase fast curve time adder. """
        return self.getData('PH_TIMEADDF')

    @property
    def PH_TIMEADDS(self):
        """ (float) Recloser-Phase slow curve time adder. """
        return self.getData('PH_TIMEADDS')

    @property
    def PH_TIMEMULTS(self):
        """ (float) Recloser-Phase slow curve time multiplier. """
        return self.getData('PH_TIMEMULTS')

    @property
    def PH_TIMEMULTF(self):
        """ (float) Recloser-Phase fast curve time multiplier. """
        return self.getData('PH_TIMEMULTF')


class RLYD(RELAYABSTRACT):
    """ Differential Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Differential Relay constructor (Exception if not found).

        Samples:
            RLYD("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                 #GUID
            RLYD("[DEVICEDIFF]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            RLYD([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYD, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def CTGRP1(self):
        """ (RLYGROUP) Local current input 1. """
        return self.getData('CTGRP1')

    @property
    def CTR1(self):
        """ (float) CTR1. """
        return self.getData('CTR1')

    @property
    def IMIN3I0(self):
        """ (float) Minimum enable differential current (3I0). """
        return self.getData('IMIN3I0')

    @property
    def IMIN3I2(self):
        """ (float) Minimum enable differential current (3I2). """
        return self.getData('IMIN3I2')

    @property
    def IMINPH(self):
        """ (float) Minimum enable differential current (phase). """
        return self.getData('IMINPH')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def RMTE1(self):
        """ (EQUIPMENT) Remote device 1. """
        return self.getData('RMTE1')

    @property
    def RMTE2(self):
        """ (EQUIPMENT) Remote device 2. """
        return self.getData('RMTE2')

    @property
    def SGLONLY(self):
        """ (int) Signal only: 1-true; 0-false. """
        return self.getData('SGLONLY')

    @property
    def TLCCV3I0(self):
        """ (str) Tapped load coordination curve (I0). """
        return self.getData('TLCCV3I0')

    @property
    def TLCCVI2(self):
        """ (str) Tapped load coordination curve (I2). """
        return self.getData('TLCCVI2')

    @property
    def TLCCVPH(self):
        """ (str) Tapped load coordination curve (phase). """
        return self.getData('TLCCVPH')

    @property
    def TLCTD3I0(self):
        """ (float) Tapped load coordination delay (I0). """
        return self.getData('TLCTD3I0')

    @property
    def TLCTDI2(self):
        """ (float) Tapped load coordination delay (I2). """
        return self.getData('TLCTDI2')

    @property
    def TLCTDPH(self):
        """ (float) Tapped load coordination delay (phase). """
        return self.getData('TLCTDPH')


class RLYDSG(RELAY3ABSTRACT):
    """ Distance Ground Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Distance Ground Relay constructor (Exception if not found).

        Samples:
            RLYDSG("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                         #GUID
            RLYDSG("[DSRLYG]  NV_Reusen G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            RLYDSG([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYDSG, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)
        __getSettingName__(self)

    @property
    def DSTYPE(self):
        """ (str) Type name (7SA511,SEL411G,..). """
        return self.getData('DSTYPE')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def SNLZONE(self):
        """ (int) Signal-only zone. """
        return self.getData('SNLZONE')

    @property
    def STARTZ2FLAG(self):
        """ (int) Start Z2 timer on forward Z3 or Z4 pickup FLAG. """
        return self.getData('STARTZ2FLAG')

    @property
    def TYPE(self):
        """ (str) Type (ID2). """
        return self.getData('TYPE')

    @property
    def VTBUS(self):
        """ (BUS) VT at Bus. """
        return self.getData('VTBUS')

    @property
    def Z2OCTYPE(self):
        """ (str) Zone 2 OC supervision type name. """
        return self.getData('Z2OCTYPE')

    @property
    def Z2OCPICKUP(self):
        """ (float) Z2 OC supervision Pickup(A). """
        return self.getData('Z2OCPICKUP')

    @property
    def Z2OCTD(self):
        """ (float) Z2 OC supervision time dial. """
        return self.getData('Z2OCTD')


class RLYDSP(RELAY3ABSTRACT):
    """ Distance Phase Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Distance Phase Relay Object constructor (Exception if not found).

        Samples:
            RLYDSP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                    #GUID
            RLYDSP("[DSRLYP]  GCXTEST@6 'NEVADA' 132 kV-2 'CLAYTOR' 132 kV 1 L") #STR
            RLYDSP([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYDSP, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)
        __getSettingName__(self)

    @property
    def DSTYPE(self):
        """ (str) Type name (7SA511,SEL411G,..). """
        return self.getData('DSTYPE')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def SNLZONE(self):
        """ (int) Signal-only zone. """
        return self.getData('SNLZONE')

    @property
    def STARTZ2FLAG(self):
        """ (int) Start Z2 timer on forward Z3 or Z4 pickup FLAG. """
        return self.getData('STARTZ2FLAG')

    @property
    def TYPE(self):
        """ (str) Type (ID2). """
        return self.getData('TYPE')

    @property
    def VTBUS(self):
        """ (BUS) VT at Bus. """
        return self.getData('VTBUS')

    @property
    def Z2OCTYPE(self):
        """ (str) Zone 2 OC supervision type name. """
        return self.getData('Z2OCTYPE')

    @property
    def Z2OCPICKUP(self):
        """ (float) Z2 OC supervision Pickup(A). """
        return self.getData('Z2OCPICKUP')

    @property
    def Z2OCTD(self):
        """ (float) Z2 OC supervision time dial. """
        return self.getData('Z2OCTD')


class RLYGROUP(DATAABSTRACT):
    """ Relay Group Object. """

    def __init__(self, key=None, hnd=None):
        """ Relay Group constructor (Exception if not found).

        Samples:
            RLYGROUP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")               #GUID
            RLYGROUP("[RELAYGROUP] 6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            RLYGROUP([b1,b2,CID,BRCODE])
                        b1      BUS|str|int|[str,float/int]
                        b2      BUS|str|int|[str,float/int]
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
        """
        hnd = __initOBJ__(RLYGROUP, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BACKUP(self):
        """ [RLYGROUP] List of RLYGROUPs-Downstream (this RLYGROUP is backups for). """
        return [RLYGROUP(hnd=h1.__hnd__) for h1 in self.getData('BACKUP')]

    @property
    def BUS1(self):
        """ (BUS) Bus Local of RLYGROUP. """
        return self.TERMINAL.BUS1

    @property
    def BUS(self):
        """ [BUS] List of Buses of RLYGROUP:  BUS[0] - Bus Local ; BUS[1],(BUS[2]) - Bus Opposite(s). """
        return [BUS(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'BUS')]

    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that RLYGROUP located on. """
        return self.getData('EQUIPMENT')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def FUSE(self):
        """ [FUSE] List of Fuses of RLYGROUP. """
        return [FUSE(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'FUSE')]

    @property
    def INTRPTIME(self):
        """ (float) Interrupting time (cycles). """
        return self.getData('INTRPTIME')

    @property
    def LOGICRECL(self):
        """ (SCHEME) Reclose logic scheme of RLYGROUP (None if not found). """
        try:
            return SCHEME(hnd=self.getData('LOGICRECL').__hnd__)
        except:
            return None

    @property
    def LOGICTRIP(self):
        """ (SCHEME) Trip logic scheme of RLYGROUP (None if not found). """
        try:
            return SCHEME(hnd=self.getData('LOGICTRIP').__hnd__)
        except:
            return None

    @property
    def NOTE(self):
        """ (str) Annotation. """
        return self.getData('NOTE')

    @property
    def OPFLAG(self):
        """ (int) Total operations. """
        return self.getData('OPFLAG')

    @property
    def PRIMARY(self):
        """ [RLYGROUP] List of RLYGROUPs-Upstream (backups for this RLYGROUP). """
        return [RLYGROUP(hnd=h1.__hnd__) for h1 in self.getData('PRIMARY')]

    @property
    def RECLSR(self):
        """ [RECLSR] List of Reclosers of RLYGROUP. """
        return [RECLSR(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RECLSR')]

    @property
    def RECLSRTIME(self):
        """ [float]*4 Reclosing intervals. """
        return self.getData('RECLSRTIME')

    @property
    def RELAY(self):
        """ [RELAY] List of All Relays of RLYGROUP. """
        return __getRLYGROUP_RLY__(self)

    @property
    def RLYD(self):
        """ [RLYD] List of Differential Relays of RLYGROUP. """
        return [RLYD(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYD')]

    @property
    def RLYDS(self):
        """ [RLYDSG+RLYDSP] List of Distance Relays (Ground+Phase) of RLYGROUP. """
        res = self.RLYDSG
        res.extend(self.RLYDSP)
        return res

    @property
    def RLYDSG(self):
        """ [RLYDSG] List of Distance Ground Relays of RLYGROUP. """
        return [RLYDSG(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYDSG')]

    @property
    def RLYDSP(self):
        """ [RLYDSP] List of Distance Phase Relays of RLYGROUP. """
        return [RLYDSP(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYDSP')]

    @property
    def RLYOC(self):
        """ [RLYOCG+RLYOCP] List of Overcurrent Relays (Ground+Phase) of RLYGROUP. """
        res = self.RLYOCG
        res.extend(self.RLYOCP)
        return res

    @property
    def RLYOCG(self):
        """ [RLYOCG] List of Overcurrent Ground Relays of RLYGROUP. """
        return [RLYOCG(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYOCG')]

    @property
    def RLYOCP(self):
        """ [RLYOCP] List of Overcurrent Phase Relays of RLYGROUP. """
        return [RLYOCP(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYOCP')]

    @property
    def RLYV(self):
        """ [RLYV] List of Voltage Relays of RLYGROUP. """
        return [RLYV(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'RLYV')]

    @property
    def SCHEME(self):
        """ [SCHEME] List of Logic Schemes of RLYGROUP. """
        return [SCHEME(hnd=h1) for h1 in __getRLYGROUP_OBJ__(self, 'SCHEME')]

    @property
    def TERMINAL(self):
        """ (TERMINAL) TERMINAL of RLYGROUP. """
        return TERMINAL(hnd=self.getData('TERMINAL').__hnd__)

    def addBACKUP(self, r1):
        """ add BACKUP to this RLYGROUP

        Args:
            r1 (str/RLYGROUP): RLYGROUP to Add

        return: None

        Remarks: do Nothing if r1 (RLYGROUP) not found
        """
        return __addCOORDPAIR__(self, 'BACKUP', r1)

    def addPRIMARY(self, r1):
        """ add PRIMARY to this RLYGROUP

        Args:
            r1 (str/RLYGROUP): RLYGROUP to Add

        return: None

        Remarks: do Nothing if r1 (RLYGROUP) not found
        """
        return __addCOORDPAIR__(self, 'PRIMARY', r1)

    def removeBACKUP(self, r1):
        """ remove BACKUP from this RLYGROUP

        Args:
            r1 : str/RLYGROUP  RLYGROUP to remove

        return: None

        Remarks: do Nothing if r1 (RLYGROUP) not found or r1 not in current BACKUP
        """
        return __removeCOORDPAIR__(self, 'BACKUP', r1)

    def removePRIMARY(self, r1):
        """ remove PRIMARY from this RLYGROUP

        Args:
            r1 : str/RLYGROUP  RLYGROUP to remove

        return: None

        Remarks: do Nothing if r1 (RLYGROUP) not found or r1 not in current PRIMARY
        """
        return __removeCOORDPAIR__(self, 'PRIMARY', r1)


class RLYOCG(RELAY3ABSTRACT):
    """ Overcurrent Ground Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Overcurrent Ground Relay constructor (Exception if not found).

        Samples:
            RLYOCG("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                  #GUID
            RLYOCG("[OCRLYG]  NV-G1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            RLYOCG([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYOCG, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)
        self.__paramEx__['POLAR'] = self.POLAR
        __getSettingName__(self)

    @property
    def ASYM(self):
        """ (int) Sensitive to DC offset: 1-true; 0-false. """
        return self.getData('ASYM')

    @property
    def CT(self):
        """ (float) CT ratio. """
        return self.getData('CT')

    @property
    def CTLOC(self):
        """ (int) CT Location. """
        return self.getData('CTLOC')

    @property
    def CTSTR(self):
        """ (str) String CT ratio. """
        return self.getData('CTSTR')

    @property
    def DTDIR(self):
        """ (int) Inst. Directional FLAG: 0=None; 1-Fwd.; 2-Rev. """
        return self.getData('DTDIR')

    @property
    def DTPICKUP(self):
        """ [float]*5 Pickups Sec.A. """
        return self.getData('DTPICKUP')

    @property
    def DTDELAY(self):
        """ [float]*5 Delays seconds. """
        return self.getData('DTDELAY')

    @property
    def DTTIMEADD(self):
        """ (float) Time adder for INST/DTx. """
        return self.getData('DTTIMEADD')

    @property
    def DTTIMEMULT(self):
        """ (float) Time multiplier for INST/DTx. """
        return self.getData('DTTIMEMULT')

    @property
    def FLATINST(self):
        """ (int) Flat definite time delay FLAG: 1-true; 0-false. """
        return self.getData('FLATINST')

    @property
    def INSTSETTING(self):
        """ (float) Instantaneous setting. """
        return self.getData('INSTSETTING')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def MINTIME(self):
        """ (float) Minimum trip time. """
        return self.getData('MINTIME')

    @property
    def OCDIR(self):
        """ (int) Directional FLAG: 0-None; 1-Fwd.; 2-Rev. """
        return self.getData('OCDIR')

    @property
    def OPI(self):
        """ (int) Operate On: 0-3I0; 1-3I2; 2-I0; 3-I2. """
        return self.getData('OPI')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def POLAR(self):
        """ (int) Polar option: 0-'Vo,Io polarized',1:'V2,I2 polarized',2:'SEL V2 polarized',3:'SEL V2 polarized'. """
        return self.getData('POLAR')

    @property
    def PICKUPTAP(self):
        """ (float) Pickup (A). """
        return self.getData('PICKUPTAP')

    @property
    def SGNL(self):
        """ (int) Signal only: Example 1-INST; 2-OC; 4-DT1;... """
        return self.getData('SGNL')

    @property
    def TAPTYPE(self):
        """ (str) Tap Type. """
        return self.getData('TAPTYPE')

    @property
    def TDIAL(self):
        """ (float) Time dial. """
        return self.getData('TDIAL')

    @property
    def TIMEADD(self):
        """ (float) Time adder. """
        return self.getData('TIMEADD')

    @property
    def TIMEMULT(self):
        """ (float) Time multiplier. """
        return self.getData('TIMEMULT')

    @property
    def TIMERESET(self):
        """ (float) Reset time. """
        return self.getData('TIMERESET')

    @property
    def TYPE(self):
        """ (str) Type. """
        return self.getData('TYPE')


class RLYOCP(RELAY3ABSTRACT):
    """ Overcurrent Phase Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Overcurrent Phase Relay constructor (Exception if not found).

        Samples:
            RLYOCP("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                  #GUID
            RLYOCP("[OCRLYP]  NV-P1@6 'NEVADA' 132 kV-8 'REUSENS' 132 kV 1 L") #STR
            RLYOCP([b1,b2,CID,BRCODE,ID])
                    b1      BUS|str|int|[str,f_i] with f_i: float or int
                    b2      BUS|str|int|[str,f_i] with f_i: float or int
                    CID     (str) circuit ID
                    BRCODE  (str) in ['X','T','P','L','DC','S','W']
                    ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYOCP, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)
        self.__paramEx__['POLAR'] = self.POLAR
        __getSettingName__(self)

    @property
    def ASYM(self):
        """ (int) Sensitive to DC offset: 1-true; 0-false. """
        return self.getData('ASYM')

    @property
    def CT(self):
        """ (float) CT ratio. """
        return self.getData('CT')

    @property
    def CTCONNECT(self):
        """ (int) CT connection: 0- Wye; 1-Delta. """
        return self.getData('CTCONNECT')

    @property
    def CTSTR(self):
        """ (str) String CT ratio. """
        return self.getData('CTSTR')

    @property
    def DTDIR(self):
        """ (int) Inst. Directional flag: 0=None;1=Fwd.;2=Rev. """
        return self.getData('DTDIR')

    @property
    def DTTIMEADD(self):
        """ (float) Time adder for INST/DTx. """
        return self.getData('DTTIMEADD')

    @property
    def DTTIMEMULT(self):
        """ (float) Time multiplier for INST/DTx. """
        return self.getData('DTTIMEMULT')

    @property
    def DTPICKUP(self):
        """ [float]*5 Pickup. """
        return self.getData('DTPICKUP')

    @property
    def DTDELAY(self):
        """ [float]*5 Delay. """
        return self.getData('DTDELAY')

    @property
    def FLATINST(self):
        """ (int) Flat delay: 1-true; 0-false. """
        return self.getData('FLATINST')

    @property
    def INSTSETTING(self):
        """ (float) Instantaneous setting. """
        return self.getData('INSTSETTING')

    @property
    def LIBNAME(self):
        """ (str) Library. """
        return self.getData('LIBNAME')

    @property
    def MINTIME(self):
        """ (float) Minimum trip time. """
        return self.getData('MINTIME')

    @property
    def OCDIR(self):
        """ (int) Directional FLAG: 0=None;1=Fwd.;2=Rev. """
        return self.getData('OCDIR')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def POLAR(self):
        """ (int) Polar option: 0-'Cross-V polarized'; 2-'SEL V2 polarized'. """
        return self.getData('POLAR')

    @property
    def PICKUPTAP(self):
        """ (float) Tap. """
        return self.getData('PICKUPTAP')

    @property
    def SGNL(self):
        """ (int) Signal only: Example 1-INST; 2-OC; 4-DT1;... """
        return self.getData('SGNL')

    @property
    def TAPTYPE(self):
        """ (str) Tap Type. """
        return self.getData('TAPTYPE')

    @property
    def TDIAL(self):
        """ (float) Time dial. """
        return self.getData('TDIAL')

    @property
    def TIMEADD(self):
        """ (float) Time adder. """
        return self.getData('TIMEADD')

    @property
    def TIMEMULT(self):
        """ (float) Time multiplier. """
        return self.getData('TIMEMULT')

    @property
    def TIMERESET(self):
        """ (float) Reset time. """
        return self.getData('TIMERESET')

    @property
    def TYPE(self):
        """ (str) Type. """
        return self.getData('TYPE')

    @property
    def VOLTCONTROL(self):
        """ (int) Voltage controlled or restrained. """
        return self.getData('VOLTCONTROL')

    @property
    def VOLTPERCENT(self):
        """ (float) Voltage controlled or restrained percentage. """
        return self.getData('VOLTPERCENT')


class RLYV(RELAYABSTRACT):
    """ Voltage Relay Object. """

    def __init__(self, key=None, hnd=None):
        """ Voltage Relay constructor (Exception if not found).

        Samples:
            RLYV("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")               #GUID
            RLYV("[DEVICEVR]  @2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            RLYV([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(RLYV, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def OPQTY(self):
        """ (int) Operate on voltage option: 1-Phase-to-Neutral; 2- Phase-to-Phase; 3-3V0;4-V1;5-V2;6-VA;7-VB;8-VC;9-VBC;10-VAB;11-VCA. """
        return self.getData('OPQTY')

    @property
    def OVCVR(self):
        """ (str) Over-voltage element curve. """
        return self.getData('OVCVR')

    @property
    def OVDELAYTD(self):
        """ (float) Over-voltage delay. """
        return self.getData('OVDELAYTD')

    @property
    def OVINST(self):
        """ (float) Over-voltage instant pickup (V). """
        return self.getData('OVINST')

    @property
    def OVPICKUP(self):
        """ (float) Over-voltage pickup (V). """
        return self.getData('OVPICKUP')

    @property
    def PACKAGE(self):
        """ (int) Package option. """
        return self.getData('PACKAGE')

    @property
    def PTR(self):
        """ (float) PT ratio. """
        return self.getData('PTR')

    @property
    def SGLONLY(self):
        """ [int]*4 Signal only: 0-No check; 1-Check for [Over-voltage Instant, Over-voltage Delay, Under-voltage Instant, Under-voltage Delay]. """
        return self.getData('SGLONLY')

    @property
    def UVCVR(self):
        """ (str) Under-voltage element curve. """
        return self.getData('UVCVR')

    @property
    def UVDELAYTD(self):
        """ (float) Under-voltage delay. """
        return self.getData('UVDELAYTD')

    @property
    def UVINST(self):
        """ (float) Under-voltage instant pickup (V). """
        return self.getData('UVINST')

    @property
    def UVPICKUP(self):
        """ (float) Under-voltage pickup (V). """
        return self.getData('UVPICKUP')


class SCHEME(DATAABSTRACT):
    """ Logic Scheme Object. """

    def __init__(self, key=None, hnd=None):
        """ Logic Scheme constructor (Exception if not found).

        Samples:
            SCHEME("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}")                            #GUID
            SCHEME("[PILOT]  272-POTT-SEL421G@2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L") #STR
            SCHEME([b1,b2,CID,BRCODE,ID])
                        b1      BUS|str|int|[str,f_i] with f_i: float or int
                        b2      BUS|str|int|[str,f_i] with f_i: float or int
                        CID     (str) circuit ID
                        BRCODE  (str) in ['X','T','P','L','DC','S','W']
                        ID      (str) Relay ID
        """
        hnd = __initOBJ__(SCHEME, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def ASSETID(self):
        """ (str) Asset ID. """
        return self.getData('ASSETID')

    @property
    def BUS1(self):
        """ (BUS) Bus Local of SCHEME. """
        return self.RLYGROUP.BUS1

    @property
    def BUS(self):
        """ [BUS] List of Buses of SCHEME. """
        return self.RLYGROUP.BUS

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that SCHEME located on. """
        return self.RLYGROUP.EQUIPMENT

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def ID(self):
        """ (str) ID. """
        return self.getData('ID')

    @property
    def TYPE(self):
        """ (str) Type. """
        return self.getData('TYPE')

    @property
    def EQUATION(self):
        """ (str) Equation. """
        return self.getData('EQUATION')

    @property
    def RLYGROUP(self):
        """ (RLYGROUP) Relay Group that SCHEME located on. """
        return RLYGROUP(hnd=self.getData('RLYGROUP').__hnd__)

    @property
    def SGLONLY(self):
        """ (int) Signal only. """
        return self.getData('SIGNALONLY')

    @property
    def LOGICVARNAME(self):
        """ [str] List of Logic Var names. """
        return self.getData('LOGICVARNAME')

    @property
    def LOGICSTR(self):
        """ (str) Logic in JSON Form. """
        return self.getData('LOGICSTR')

    def getLogic(self, nameVar=None):
        """ (str) get Logic of SCHEME.

        Args:
            nameVar : name of var  (if=None get all Logic +EQUATION of SCHEME)

        Samples:
            ls1.getLogic()                 => {'RU_NEAR':['INST/DT PICKUP', OlxObj.RLYOCG],'TS':0.4,...}
            ls1.getLogic('RU_NEAR')        => ['INST/DT PICKUP', OlxObj.RLYOCG]
            ls1.getLogic(['RU_NEAR','TS']) => {'RU_NEAR':['INST/DT PICKUP', OlxObj.RLYOCG],'TS':0.4}
        """
        res = __scheme_getLogic__(self, nameVar)
        if messError:
            raise Exception(messError)
        return res

    def changeLogic(self, nameVar, value):
        """ changeLogic of SCHEME
            postData() is not necessary

        Args:

            nameVar : (str) logic var name
            value   : float or [str, str/obj] new values of logic var

        Samples:
            ls1.changeLogic('TS',0.4)
            ls1.changeLogic('RU_NEAR',['INST/DT TRIP',"[OCRLYG]  FL-G1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"])
        """
        __scheme_changeLogic__(self, nameVar, value)
        if messError:
            raise Exception(messError)

    def setLogic(self,logic):
        """ set Logic of SCHEME (all logic + EQUATION)
            postData() is not necessary

        Args:
            logic : all logic + EQUATION

        Samples:
            logic = {'EQUATION': 'RU_NEAR * (RU_FAR @ TS + RO_NEAR)',
                'RU_NEAR': ['INST/DT PICKUP', "[OCRLYG]  FL-G1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"],
                'RU_FAR': ['OPEN OP. #1', "[TERMINAL] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1 L"],
                'TS': 0.42,
                'RO_NEAR': ['OV PICKUP', "[DEVICEVR]  rlv1@5 'FIELDALE' 132 kV-2 'CLAYTOR' 132 kV 1 L"]}
            ls1.setLogic(logic)
        """
        __scheme_setLogic__(self, logic)
        if messError:
            raise Exception(messError)


class SERIESRC(DATAABSTRACT):
    """ Series reactors/capacitor Object. """

    def __init__(self, key=None, hnd=None):
        """ Series reactors/capacitor constructor (Exception if not found).

        Samples:
            SERIESRC('_{D87FBEC5-C4F9-47FE-96D3-8F24CCB101FF}')           #GUID
            SERIESRC("[SERIESRC] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") #STR
            SERIESRC([['CLAYTOR', 132], ['NEVADA', 132],'1'])             #bus1,bus2,CID
        """
        hnd = __initOBJ__(SERIESRC, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BRCODE(self):
        """ (str) BRCODE='S'. """
        return 'S'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (SERIESRC) return SERIESRC-self. """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service; 3- bypassed. """
        return self.getData('FLAG')

    @property
    def IPR(self):
        """ (float) Protective level current. """
        return self.getData('IPR')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def R(self):
        """ (float) Resistance R, in pu. """
        return self.getData('R')

    @property
    def SCOMP(self):
        """ (int) Bypassed FLAG: 1- no bypassed; 2-yes bypassed. """
        return self.getData('SCOMP')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2]

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def X(self):
        """ (float) Reactance X, in pu. """
        return self.getData('X')


class SHIFTER(DATAABSTRACT):
    """ Phase Shifter Object. """

    def __init__(self, key=None, hnd=None):
        """ Phase Shifter constructor (Exception if not found).

        Samples:
            SHIFTER('_{F915233C-20AB-43C5-A7C4-3BF29F0F890F}')           #GUID
            SHIFTER("[SHIFTER] 'NEVADA PST' 132 kV-6 'NEVADA' 132 kV 1") #STR
            SHIFTER([['NEVADA PST', 1],['NEVADA', 1],'1'])               #bus1,bus2,CID
        """
        hnd = __initOBJ__(SHIFTER, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def ANGMAX(self):
        """ (float) Shift Angle max. """
        return self.getData('ANGMAX')

    @property
    def ANGMIN(self):
        """ (float) Shift Angle min. """
        return self.getData('ANGMIN')

    @property
    def BASEMVA(self):
        """ (float) BaseMVA (for per-unit quantities). """
        return self.getData('BASEMVA')

    @property
    def BN1(self):
        """ (float) -seq Susceptance Bn1 (at Bus1), in pu. """
        return self.getData('BN1')

    @property
    def BN2(self):
        """ (float) -seq Susceptance Bn2 (at Bus2), in pu. """
        return self.getData('BN2')

    @property
    def BP1(self):
        """ (float) +seq Susceptance Bp1 (at Bus1), in pu. """
        return self.getData('BP1')

    @property
    def BP2(self):
        """ (float) +seq Susceptance Bp2 (at Bus2), in pu. """
        return self.getData('BP2')

    @property
    def BRCODE(self):
        """ (str) BRCODE = 'P'. """
        return 'P'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def BZ1(self):
        """ (float) zero Susceptance Bz1 (at Bus1), in pu. """
        return self.getData('BZ1')

    @property
    def BZ2(self):
        """ (float) zero Susceptance Bz2 (at Bus2), in pu. """
        return self.getData('BZ2')

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def CNTL(self):
        """ (int) Control mode: 0-Fixed; 1-automatically control real power flow. """
        return self.getData('CNTL')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (SHIFTER) return SHIFTER-self """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def GN1(self):
        """ (float) -seq Conductance Gn1 (at Bus1), in pu. """
        return self.getData('GN1')

    @property
    def GN2(self):
        """ (float) -seq Conductance Gn2 (at Bus2), in pu. """
        return self.getData('GN2')

    @property
    def GP1(self):
        """ (float) +seq Conductance Gp1 (at Bus1), in pu. """
        return self.getData('GP1')

    @property
    def GP2(self):
        """ (float) +seq Conductance Gp2 (at Bus2), in pu. """
        return self.getData('GP2')

    @property
    def GZ1(self):
        """ (float) zero Conductance Gz1 (at Bus1), in pu. """
        return self.getData('GZ1')

    @property
    def GZ2(self):
        """ (float) zero Conductance Gz2 (at Bus2), in pu. """
        return self.getData('GZ2')

    @property
    def MVA1(self):
        """ (float) Rating MVA1. """
        return self.getData('MVA1')

    @property
    def MVA2(self):
        """ (float) Rating MVA2. """
        return self.getData('MVA2')

    @property
    def MVA3(self):
        """ (float) Rating MVA3. """
        return self.getData('MVA3')

    @property
    def MWMAX(self):
        """ (float) MW max. """
        return self.getData('MWMAX')

    @property
    def MWMIN(self):
        """ (float) MW min. """
        return self.getData('MWMIN')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2]

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group 1 at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group 2 at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def RN(self):
        """ (float) -seq Resistance Rn, in pu. """
        return self.getData('RN')

    @property
    def RP(self):
        """ (float) +seq Resistance Rp, in pu. """
        return self.getData('RP')

    @property
    def RZ(self):
        """ (float) zero Resistance Rz, in pu. """
        return self.getData('RZ')

    @property
    def SHIFTANGLE(self):
        """ (float) Shift Angle. """
        return self.getData('SHIFTANGLE')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def XN(self):
        """ (float) -seq Reactance Xn, in pu. """
        return self.getData('XN')

    @property
    def XP(self):
        """ (float) +seq Reactance Xp, in pu. """
        return self.getData('XP')

    @property
    def XZ(self):
        """ (float) zero Reactance Xz, in pu. """
        return self.getData('XZ')

    @property
    def ZCORRECTNO(self):
        """ (int) Correct Table Number. """
        return self.getData('ZCORRECTNO')


class SHUNT(DATAABSTRACT):
    """ Shunt Object. """

    def __init__(self, key=None, hnd=None):
        """ Shunt constructor (Exception if not found).

        Samples:
            SHUNT('_{5C9526DC-D64A-4E92-AC4D-9D56AE380A05}') #GUID
            SHUNT("[SHUNT] 21 'IOWA' 33 kV")                 #STR
            SHUNT(['IOWA', 33])                              #Name,kV
        """
        hnd = __initOBJ__(SHUNT, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ (BUS) BUS that Shunt located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def SHUNTUNIT(self):
        """ [SHUNTUNIT] List of Shunt Units in SHUNT. """
        return [SHUNTUNIT(hnd=ti.__hnd__) for ti in self.BUS.SHUNTUNIT]


class SHUNTUNIT(DATAABSTRACT):
    """ Shunt Unit Object. """

    def __init__(self, key=None, hnd=None):
        """ Shunt Unit constructor (Exception if not found).

        Samples:
            SHUNTUNIT('_{5C9526DC-D64A-4E92-AC4D-9D56AE380A05}') #GUID
            SHUNTUNIT("[CAPUNIT]  1@21 'IOWA' 33 kV")            #STR
            SHUNTUNIT(['IOWA', 33,'1'])                          #Name,kV,CID
        """
        hnd = __initOBJ__(SHUNTUNIT, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def B(self):
        """ (float) +seq succeptance, in pu. """
        return self.getData('B')

    @property
    def B0(self):
        """ (float) zero succeptance, in pu. """
        return self.getData('B0')

    @property
    def BUS(self):
        """ (BUS) BUS that Shunt Unit located on. """
        return BUS(hnd=self.SHUNT.BUS.__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def G(self):
        """ (float) +seq conductance, in pu. """
        return self.getData('G')

    @property
    def G0(self):
        """ (float) zero conductance, in pu. """
        return self.getData('G0')

    @property
    def SHUNT(self):
        """ (SHUNT) Shunt that Shunt Unit located on. """
        return SHUNT(hnd=self.getData('SHUNT').__hnd__)

    @property
    def TX3(self):
        """ (int) 3-winding transformer FLAG: 1-true; 0-false. """
        return self.getData('TX3')


class SVD(DATAABSTRACT):
    """ Switched Shunt Object. """

    def __init__(self, key=None, hnd=None):
        """ Switched Shunt constructor (Exception if not found).

        Samples:
            SVD('_{7BFBA30C-6476-4A55-BAB6-4DF3A9B9443E}') #GUID
            SVD("[SVD] 3 'TEXAS' 132 kV")                  #STR
            SVD(['TEXAS', 132])                            #Name,kV
        """
        hnd = __initOBJ__(SVD, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def B(self):
        """ [float]*8 +seq susceptance, in pu. """
        return self.getData('B')

    @property
    def B0(self):
        """ [float]*8 zero susceptance, in pu. """
        return self.getData('B0')

    @property
    def BUS(self):
        """ (BUS) BUS that SVD located on. """
        return BUS(hnd=self.getData('BUS').__hnd__)

    @property
    def B_USE(self):
        """ (float) Susceptance in use, in pu. """
        return self.getData('B_USE')

    @property
    def CNTBUS(self):
        """ (BUS) Controled Bus. """
        return BUS(hnd=self.getData('CNTBUS').__hnd__)

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def MODE(self):
        """ (int) Control mode: 0-Fixed; 1-Discrete; 2-Continous. """
        return self.getData('MODE')

    @property
    def STEP(self):
        """ [int]*8 Number of step. """
        return self.getData('STEP')

    @property
    def VMAX(self):
        """ (float) Max V. """
        return self.getData('VMAX')

    @property
    def VMIN(self):
        """ (float) Min V. """
        return self.getData('VMIN')


class SWITCH(DATAABSTRACT):
    """ Switch Object. """

    def __init__(self, key=None, hnd=None):
        """ SWITCH constructor (Exception if not found).

        Samples:
            SWITCH('_{D87FBEC5-C4F9-47FE-96D3-8F24CCB101FF}')         #GUID
            SWITCH("[SWITCH] 2 'CLAYTOR' 132 kV-6 'NEVADA' 132 kV 1") #STR
            SWITCH([['CLAYTOR', 132],['NEVADA', 132],'1'])            #bus1,bus2,CID
        """
        hnd = __initOBJ__(SWITCH, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BRCODE(self):
        """ (str) BRCODE = 'W'. """
        return 'W'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def DEFAULT(self):
        """ (int) Default position FLAG: 1- normaly open; 2- normaly close; 0-Not defined. """
        return self.getData('DEFAULT')

    @property
    def EQUIPMENT(self):
        """ (SWITCH) return SWITCH-self. """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def RATING(self):
        """ (float) Current Rating. """
        return self.getData('RATING')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2]

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def STAT(self):
        """ (int) Position FLAG: 7- close; 0- open. """
        return self.getData('STAT')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]


class TERMINAL(DATAABSTRACT):
    """ Terminal Object. """

    def __init__(self, key=None, hnd=None):
        """ TERMINAL constructor (Exception if not found).

        Samples:
            TERMINAL([b1,b2,CID,BRCODE])
                b1      BUS|str|int|[str,float/int]
                b2      BUS|str|int|[str,float/int]
                CID     (str) circuit ID
                BRCODE  (str) in ['X','T','P','L','DC','S','W','XFMR3','XFMR','SHIFTER','LINE','DCLINE2','SERIESRC','SWITCH']
        """
        hnd = __initOBJ__(TERMINAL, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [BUS(hnd=bi.__hnd__) for bi in __getTERMINAL_OBJ__(self, 'BUS')]

    @property
    def BUS1(self):
        """ (BUS) Bus Local. """
        return BUS(hnd=__getDatai__(self.__hnd__, OlxAPIConst.BR_nBus1Hnd))

    @property
    def BUS2(self):
        """ (BUS) 1st Bus opposite. """
        return BUS(hnd=__getDatai__(self.__hnd__, OlxAPIConst.BR_nBus2Hnd))

    @property
    def BUS3(self):
        """ (BUS) 2nd Bus opposite (None if not on 3-windings Transformers). """
        h = __getDatai__(self.__hnd__, OlxAPIConst.BR_nBus3Hnd)
        if h is None:
            return None
        return BUS(hnd=h)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.EQUIPMENT.CID

    @property
    def EQUIPMENT(self):
        """ (EQUIPMENT) EQUIPMENT that TERMINAL located on. """
        return __getTERMINAL_OBJ__(self, 'EQUIPMENT')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return __getTERMINAL_OBJ__(self, 'FLAG')

    @property
    def OPPOSITE(self):
        """ [TERMINAL] List of TERMINALs that opposite on the EQUIPMENT. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __getTERMINAL_OBJ__(self, 'OPPOSITE')]

    @property
    def REMOTE(self):
        """ [TERMINAL] List of TERMINALs remote to TERMINAL.

            All taps are ignored.
            Close switches are included.
            Out of service branches are ignored.
        """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __getTERMINAL_OBJ__(self, 'REMOTE')]

    @property
    def RLYGROUP(self):
        """ [RLYGROUP]*3 List of RLYGROUPs that attached to TERMINAL.

            RLYGROUP[0] : local RLYGROUP,   = None if not found
            RLYGROUP[1] : opposite RLYGROUP = None if not found
            RLYGROUP[2] : opposite RLYGROUP = None if not found or not on 3-windings Transformer
        """
        res = []
        for ri in __getTERMINAL_OBJ__(self, 'RLYGROUP'):
            if ri is not None:
                res.append(RLYGROUP(hnd=ri.__hnd__))
            else:
                res.append(None)
        return res

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Local RLYGROUP (None if not found). """
        return __getTERMINAL_OBJ__(self, 'RLYGROUP1')

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) 1st opposite RLYGROUP (None if not found). """
        return __getTERMINAL_OBJ__(self, 'RLYGROUP2')

    @property
    def RLYGROUP3(self):
        """ (RLYGROUP) 2nd opposite RLYGROUP (None if not found). """
        return __getTERMINAL_OBJ__(self, 'RLYGROUP3')


class XFMR(DATAABSTRACT):
    """ 2-Windings Transformer Object. """

    def __init__(self, key=None, hnd=None):
        """ 2-Windings Transformer constructor (Exception if not found).

        Samples:
            XFMR("_{4D49357D-E376-4776-8618-D47CA9490EC3}")              #GUID
            XFMR("[XFORMER] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV 1") #STR
            XFMR([['NEVADA', 132],['NEW HAMPSHR', 132],'1'])             #bus1,bus2,CID
        """
        hnd = __initOBJ__(XFMR, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def AUTOX(self):
        """ (int) Auto transformer FLAG: 1-true ; 0-false. """
        return self.getData('AUTOX')

    @property
    def B(self):
        """ (float) +seq susceptance B, in pu. """
        return self.getData('B')

    @property
    def B0(self):
        """ (float) zero susceptance Bo, in pu. """
        return self.getData('B0')

    @property
    def B1(self):
        """ (float) +seq susceptance B1 (at Bus1), in pu. """
        return self.getData('B1')

    @property
    def B10(self):
        """ (float) zero susceptance B10 (at Bus1), in pu. """
        return self.getData('B10')

    @property
    def B2(self):
        """ (float) +seq susceptance B2 (at Bus2), in pu. """
        return self.getData('B2')

    @property
    def B20(self):
        """ (float) zero susceptance B20 (at Bus2), in pu. """
        return self.getData('B20')

    @property
    def BASEMVA(self):
        """ (float) Base MVA (for per-unit quantities). """
        return self.getData('BASEMVA')

    @property
    def BRCODE(self):
        """ (str) BRCODE = 'T'. """
        return 'T'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def CONFIGP(self):
        """ (str) Primary winding config. """
        return self.getData('CONFIGP')

    @property
    def CONFIGS(self):
        """ (str) Secondary winding config. """
        return self.getData('CONFIGS')

    @property
    def CONFIGST(self):
        """ (str) Secondary winding config in test. """
        return self.getData('CONFIGST')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (XFMR) return XFMR-self """
        return self

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def G1(self):
        """ (float) +seq conductance G1 (at Bus1), in pu. """
        return self.getData('G1')

    @property
    def G10(self):
        """ (float) zero conductance G10 (at Bus1), in pu. """
        return self.getData('G10')

    @property
    def G2(self):
        """ (float) +seq conductance G2 (at Bus2), in pu. """
        return self.getData('G2')

    @property
    def G20(self):
        """ (float) zero conductance G20 (at Bus2), in pu. """
        return self.getData('G20')

    @property
    def GANGED(self):
        """ (int) LTC tag ganged FLAG: 0-False; 1-True. """
        return self.getData('GANGED')

    @property
    def LTCCENTER(self):
        """ (float) LTC center tap. """
        return self.getData('LTCCENTER')

    @property
    def LTCCTRL(self):
        """ (BUS) BUS whose voltage magnitude is to be regulated by the LTC. """
        return self.getData('LTCCTRL')

    @property
    def LTCSIDE(self):
        """ (int) LTC side: 1; 2; 0. """
        return self.getData('LTCSIDE')

    @property
    def LTCSTEP(self):
        """ (float) LTC step size. """
        return self.getData('LTCSTEP')

    @property
    def LTCTYPE(self):
        """ (int) LTC type: 1- control voltage; 2- control MVAR. """
        return self.getData('LTCTYPE')

    @property
    def MAXTAP(self):
        """ (float) LTC max tap. """
        return self.getData('MAXTAP')

    @property
    def MAXVW(self):
        """ (float) LTC min controlled quantity limit. """
        return self.getData('MAXVW')

    @property
    def MINTAP(self):
        """ (float) LTC min tap. """
        return self.getData('MINTAP')

    @property
    def MINVW(self):
        """ (float) LTC max controlled quantity limit. """
        return self.getData('MINVW')

    @property
    def MVA1(self):
        """ (float) Rating MVA1. """
        return self.getData('MVA1')

    @property
    def MVA2(self):
        """ (float) Rating MVA2. """
        return self.getData('MVA2')

    @property
    def MVA3(self):
        """ (float) Rating MVA3. """
        return self.getData('MVA3')

    @property
    def MVA(self):
        """ [float]*3 Ratings [MVA1,MVA2,MVA3]. """
        return [self.getData('MVA1'), self.getData('MVA2'), self.getData('MVA3')]

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def PRIORITY(self):
        """ (int) LTC adjustment priority: 0-Normal; 1-Medieum; 2-High. """
        return self.getData('PRIORITY')

    @property
    def PRITAP(self):
        """ (float) Primary Tap. """
        return self.getData('PRITAP')

    @property
    def R(self):
        """ (float) +seq Resistance R, in pu. """
        return self.getData('R')

    @property
    def R0(self):
        """ (float) zero Resistance Ro, in pu. """
        return self.getData('R0')

    @property
    def RG1(self):
        """ (float) Grounding Resistance Rg1, in Ohm. """
        return self.getData('RG1')

    @property
    def RG2(self):
        """ (float) Grounding Resistance Rg2, in Ohm. """
        return self.getData('RG2')

    @property
    def RGN(self):
        """ (float) Grounding Resistance Rgn, in Ohm. """
        return self.getData('RGN')

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2]

    @property
    def SECTAP(self):
        """ (float) Secondary Tap. """
        return self.getData('SECTAP')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def TIE(self):
        """ (int) Metered FLAG: 1-at Bus1; 2-at Bus2; 0-XFMR in a single area. """
        return self.getData('TIE')

    @property
    def X(self):
        """ (float) +seq Reactance X, in pu. """
        return self.getData('X')

    @property
    def X0(self):
        """ (float) zero Reactance X0, in pu. """
        return self.getData('X0')

    @property
    def XG1(self):
        """ (float) Grounding Reactance Xg1, in Ohm. """
        return self.getData('XG1')

    @property
    def XG2(self):
        """ (float) Grounding Reactance Xg2, in Ohm. """
        return self.getData('XG2')

    @property
    def XGN(self):
        """ (float) Grounding Reactance Xgm, in Ohm. """
        return self.getData('XGN')


class XFMR3(DATAABSTRACT):
    """ 3-Windings Transformer Object. """

    def __init__(self, key=None, hnd=None):
        """ 3-Windings Transformer constructor (Exception if not found).

        Samples:
            XFMR3("_{9006257F-CD51-4EC8-AE8D-24451B972323}")  #GUID
            XFMR3("[XFORMER3] 6 'NEVADA' 132 kV-10 'NEW HAMPSHR' 33 kV-'DOT BUS' 13.8 kV 1") #STR
            XFMR3([['NEVADA', 132],['NEW HAMPSHR', 132],'1'])  #bus1,bus2,CID
        """
        hnd = __initOBJ__(XFMR3, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)

    @property
    def AUTOX(self):
        """ (int) Auto transformer FLAG: 1-true;0-false. """
        return self.getData('AUTOX')

    @property
    def B(self):
        """ (float) +seq susceptance B, in pu. """
        return self.getData('B')

    @property
    def B0(self):
        """ (float) zero susceptance B0, in pu. """
        return self.getData('B0')

    @property
    def BASEMVA(self):
        """ (float) Base MVA (for per-unit quantities). """
        return self.getData('BASEMVA')

    @property
    def BRCODE(self):
        """ (str) BRCODE = 'X'. """
        return 'X'

    @property
    def BUS(self):
        """ [BUS] List of Buses. """
        return [self.BUS1, self.BUS2, self.BUS3]

    @property
    def BUS1(self):
        """ (BUS) Bus1. """
        return BUS(hnd=self.getData('BUS1').__hnd__)

    @property
    def BUS2(self):
        """ (BUS) Bus2. """
        return BUS(hnd=self.getData('BUS2').__hnd__)

    @property
    def BUS3(self):
        """ (BUS) Bus3. """
        return BUS(hnd=self.getData('BUS3').__hnd__)

    @property
    def CID(self):
        """ (str) Circuit ID. """
        return self.getData('CID')

    @property
    def CONFIGP(self):
        """ (str) Primary winding config. """
        return self.getData('CONFIGP')

    @property
    def CONFIGS(self):
        """ (str) Secondary winding config. """
        return self.getData('CONFIGS')

    @property
    def CONFIGST(self):
        """ (str) Secondary winding config in test. """
        return self.getData('CONFIGST')

    @property
    def CONFIGT(self):
        """ (str) Tertiary winding config. """
        return self.getData('CONFIGT')

    @property
    def CONFIGTT(self):
        """ (str) Tertiary winding config in test. """
        return self.getData('CONFIGTT')

    @property
    def DATEOFF(self):
        """ (str) Out of service date. """
        return self.getData('DATEOFF')

    @property
    def DATEON(self):
        """ (str) In service date. """
        return self.getData('DATEON')

    @property
    def EQUIPMENT(self):
        """ (XFMR3) return XFMR3-self """
        return self

    @property
    def FICTBUSNO(self):
        """ (int) Fiction Bus Number. """
        return self.getData('FICTBUSNO')

    @property
    def FLAG(self):
        """ (int) In-service FLAG: 1- active; 2- out-of-service. """
        return self.getData('FLAG')

    @property
    def GANGED(self):
        """ (int) LTC tag ganged FLAG: 0-False; 1-True. """
        return self.getData('GANGED')

    @property
    def LTCCENTER(self):
        """ (float) LTC center Tap. """
        return self.getData('LTCCENTER')

    @property
    def LTCCTRL(self):
        """ (BUS) BUS whose voltage magnitude is to be regulated by the LTC. """
        return self.getData('LTCCTRL')

    @property
    def LTCSIDE(self):
        """ (int) LTC side: 1; 2; 0. """
        return self.getData('LTCSIDE')

    @property
    def LTCSTEP(self):
        """ (float) LTC step size. """
        return self.getData('LTCSTEP')

    @property
    def LTCTYPE(self):
        """ (int) LTC type: 1- control voltage; 2- control MVAR. """
        return self.getData('LTCTYPE')

    @property
    def MAXTAP(self):
        """ (float) LTC max Tap. """
        return self.getData('MAXTAP')

    @property
    def MAXVW(self):
        """ (float) LTC min controlled quantity limit. """
        return self.getData('MAXVW')

    @property
    def MINTAP(self):
        """ (float) LTC min Tap. """
        return self.getData('MINTAP')

    @property
    def MINVW(self):
        """ (float) LTC controlled quantity limit. """
        return self.getData('MINVW')

    @property
    def MVA(self):
        """ [float]*3 Ratings [MVA1,MVA2,MVA3]. """
        return [self.getData('MVA1'), self.getData('MVA2'), self.getData('MVA3')]

    @property
    def MVA1(self):
        """ (float) Rating MVA1. """
        return self.getData('MVA1')

    @property
    def MVA2(self):
        """ (float) Rating MVA2. """
        return self.getData('MVA2')

    @property
    def MVA3(self):
        """ (float) Rating MVA3. """
        return self.getData('MVA3')

    @property
    def NAME(self):
        """ (str) Name. """
        return self.getData('NAME')

    @property
    def PRIORITY(self):
        """ (int) LTC adjustment priority: 0-Normal; 1-Medieum; 2-High. """
        return self.getData('PRIORITY')

    @property
    def PRITAP(self):
        """ (float) Primary Tap. """
        return self.getData('PRITAP')

    @property
    def RG1(self):
        """ (float) Grounding Resistance Rg1, in Ohm. """
        return self.getData('RG1')

    @property
    def RG2(self):
        """ (float) Grounding Resistance Rg2, in Ohm. """
        return self.getData('RG2')

    @property
    def RG3(self):
        """ (float) Grounding Resistance Rg3, in Ohm. """
        return self.getData('RG3')

    @property
    def RGN(self):
        """ (float) Grounding Resistance Rgn, in Ohm. """
        return self.getData('RGN')

    @property
    def RLYGROUP(self):
        """ [RLYGROUP] List of RLYGROUPs. """
        return [self.RLYGROUP1, self.RLYGROUP2, self.RLYGROUP3]

    @property
    def RLYGROUP1(self):
        """ (RLYGROUP) Relay Group at Bus1 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP1').__hnd__)
        except:
            return None

    @property
    def RLYGROUP2(self):
        """ (RLYGROUP) Relay Group at Bus2 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP2').__hnd__)
        except:
            return None

    @property
    def RLYGROUP3(self):
        """ (RLYGROUP) Relay Group at Bus3 (None if not found). """
        try:
            return RLYGROUP(hnd=self.getData('RLYGROUP3').__hnd__)
        except:
            return None

    @property
    def RMG0(self):
        """ (float) RMG0. """
        return self.getData('RMG0')

    @property
    def RPM0(self):
        """ (float) RPM0. """
        return self.getData('RPM0')

    @property
    def RPS(self):
        """ (float) +seq Resistance Rps, in pu. """
        return self.getData('RPS')

    @property
    def RPS0(self):
        """ (float) zero Resistance R0ps, in pu. """
        return self.getData('RPS0')

    @property
    def RPT(self):
        """ (float) +seq Resistance Rpt, in pu. """
        return self.getData('RPT')

    @property
    def RPT0(self):
        """ (float) zero Resistance R0pt, in pu. """
        return self.getData('RPT0')

    @property
    def RSM0(self):
        """ (float) RSM0 """
        return self.getData('RSM0')

    @property
    def RST(self):
        """ (float) +seq Resistance Rst, in pu. """
        return self.getData('RST')

    @property
    def RST0(self):
        """ (float) zero Resistance R0st, in pu. """
        return self.getData('RST0')

    @property
    def SECTAP(self):
        """ (float) Secondary Tap. """
        return self.getData('SECTAP')

    @property
    def TERMINAL(self):
        """ [TERMINAL] List of TERMINALs. """
        return [TERMINAL(hnd=ti.__hnd__) for ti in __get_OBJTERMINAL__(self)]

    @property
    def TERTAP(self):
        """ (float) Tertiary Tap. """
        return self.getData('TERTAP')

    @property
    def XG1(self):
        """ (float) Grounding Reactance Xg1, in Ohm. """
        return self.getData('XG1')

    @property
    def XG2(self):
        """ (float) Grounding Reactance Xg2, in Ohm. """
        return self.getData('XG2')

    @property
    def XG3(self):
        """ (float) Grounding Reactance Xg3, in Ohm. """
        return self.getData('XG3')

    @property
    def XGN(self):
        """ (float) Grounding Reactance Xgn, in Ohm. """
        return self.getData('XGN')

    @property
    def XMG0(self):
        """ (float) XMG0. """
        return self.getData('XMG0')

    @property
    def XPM0(self):
        """ (float) XPM0. """
        return self.getData('XPM0')

    @property
    def XPS(self):
        """ (float) +seq Reactance Xps, in pu. """
        return self.getData('XPS')

    @property
    def XPS0(self):
        """ (float) zero Reactance X0ps, in pu. """
        return self.getData('XPS0')

    @property
    def XPT(self):
        """ (float) +seq Reactance Xpt, in pu. """
        return self.getData('XPT')

    @property
    def XPT0(self):
        """ (float) zero Reactance X0pt, in pu. """
        return self.getData('XPT0')

    @property
    def XSM0(self):
        """ (float) XSM0. """
        return self.getData('XSM0')

    @property
    def XST(self):
        """ (float) +seq Reactance Xst, in pu. """
        return self.getData('XST')

    @property
    def XST0(self):
        """ (float) zero Reactance X0st, in pu. """
        return self.getData('XST0')

    @property
    def Z0METHOD(self):
        """ (int) Z0 method: 1-Short circuit impedance; 2-Classical T model impedance. """
        return self.getData('Z0METHOD')


class ZCORRECT(DATAABSTRACT):
    """ Impedance Correction Table Object. """

    def __init__(self, key=None, hnd=None):
        """ Impedance Correction Table constructor (Exception if not found).

        Samples:
            ZCORRECT("_{D378E9B7-22A9-49E0-8B49-E3974E63CCF3}") #GUID
        """
        hnd = __initOBJ__(ZCORRECT, key, hnd)
        if hnd == -1:
            raise Exception(messError)
        super().__init__(hnd)


class ZONE:
    """ ZONE Object. """

    def __init__(self, no):
        """ Constructor by Zone number. """
        super().__setattr__('__no__', no)
        super().__setattr__('__currFileIdx__', __CURRENT_FILE_IDX__)

    @property
    def NAME(self):
        """ (str) Name. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return OlxAPI.GetZoneName(self.__no__)

    @property
    def NO(self):
        """ (int) Number. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        return self.__no__

    def equals(self, o2):
        """ (bool) Comparison of 2 Zones Object. """
        if type(o2) == ZONE:
            if __check_currFileIdx1__():
                raise Exception(messError)
            return self.__no__ == o2.__no__
        return False

    def toString(self, option=0):
        """ (str) Text description/composed of Zone Object. """
        if __check_currFileIdx1__():
            raise Exception(messError)
        if option == 0:
            return "[ZONE] %i '%s'" % (self.__no__, self.NAME)
        return '%i %s' % (self.__no__, self.NAME)


def __addCOORDPAIR__(o1, val, r1):
    try:
        rg1 = RLYGROUP(r1)
    except Exception as err:
        if __OLXOBJ_VERBOSE__:
            print(str(err))
        rg1 = None
    #
    if rg1 != None:
        ra = o1.getData(val)
        if not rg1.isInList(ra):
            ra.append(rg1)
            o1.changeData(val, ra)
            o1.postData()


def __addLogicVar__(o1, nameVar, value):
    if __check_currFileIdx__(o1):
        return
    o1.__paramEx__['SCHEME_FLAG'] = 0
    global messError
    if type(nameVar) != str:
        messError = '\nSCHEME.addLogicVar(nameVar=%s,value) ' % toString(nameVar)
        messError += '\n\tRequired nameVar : str'
        messError += '\n\tFound            : '+type(nameVar).__name__+'  '+toString(nameVar)
        return
    #
    nameVar0 = nameVar
    nameVar = nameVar.upper()
    #
    if not (type(value) in __fi or (type(value) == list and len(value) == 2)):
        messError = '\nSCHEME.addLogicVar(nameVar=%s,value) ' % toString(nameVar0)
        messError += '\n\tRequired value  : float/int or []*2'
        messError += '\n\tFound           : '+type(value).__name__
        if type(value) == list:
            messError += ' (len=%i)' % len(value)
        messError += ' '+toString(value)
        return
    #
    if nameVar in o1.__paramEx__['SCHEME_NAME']:
        messError = '\nSCHEME.addLogicVar(nameVar=%s,value) ' % toString(nameVar0)
        messError += '\n\tnameVar already exists : '+toString(nameVar)
        return
    #
    if type(value) != list:
        o1.__paramEx__['SCHEME_NAME'].append(nameVar)
        o1.__paramEx__['SCHEME_TYPE'][nameVar] = OlxAPIConst.LVT_CONST
        o1.__paramEx__['SCHEME_VAL'][nameVar] = value
    else:
        if value[0] not in __OLXOBJ_SCHEME__.keys():
            messError = '\nAll value for logic var : '+toString(list(__OLXOBJ_SCHEME__.keys()))
            messError += '\n\nSCHEME.addLogicVar(nameVar=%s,value=%s)' % (toString(nameVar0), toString(value))
            messError += '\nValueError : '+toString(value[0])
            return
        ox = None
        if type(value[1]) == str:
            ox = OLCase.findOBJ(value[1])
        else:
            for ob1 in __OLXOBJ_SCHEME_OB__:
                try:
                    ox = OLCase.findOBJ(ob1, value[1])
                    if ox != None:
                        break
                except:
                    pass
        if ox is None:
            messError = '\nSCHEME.addLogicVar(nameVar=%s,value=%s)' % (toString(nameVar0), toString(value))
            messError += '\nObject not found : '+toString(value[1])
            return
        #
        o1.__paramEx__['SCHEME_NAME'].insert(0, nameVar)
        # o1.__paramEx__['SCHEME_NAME'].append(nameVar)
        o1.__paramEx__['SCHEME_TYPE'][nameVar] = __OLXOBJ_SCHEME__[value[0]]
        o1.__paramEx__['SCHEME_VAL'][nameVar] = ox.__hnd__


def __addOBJ_SCHEME__(v0, param, logics):
    global messError
    messError = ''
    setting = {}
    for k,v in logics.items():
        try:
            setting[k.upper()] = v
        except:
            setting[k] = v
    #
    se = "\nOlxObj.OLCase.addOBJ('SCHEME', key, param, logic)"
    tokens = __intArray__()
    params = __voidArray__()
    #
    setVerbose(0, 1)
    #
    param1, param2 = {}, {}
    for k, v in param.items():
        if k in __OLXOBJ_PARA__['SCHEME'].keys() and __OLXOBJ_PARA__['SCHEME'][k][2]:
            param1[k] = v
        else:
            param2[k] = v  # for UDF
    #
    if 'EQUATION' in setting.keys():
        param1['EQUATION'] = setting['EQUATION']
        setting.pop('EQUATION')

    #
    s1, s2 = __check_param_setting__(None, 'SCHEME', param1, setting)
    if s2:
        messError = s1+se+s2
        return None
    #
    ii = 0
    for k, v1 in {'MEMO': OlxAPIConst.OBJ_sMemo, 'TAGS': OlxAPIConst.OBJ_sTags, 'GUID': OlxAPIConst.OBJ_sGUID}.items():
        if k in param1.keys():
            val = c_char_p(encode3(param1[k]))
            params[ii] = cast(pointer(val), c_void_p)
            tokens[ii] = v1
            ii += 1
            param1.pop(k)
    #
    for k, v in param1.items():
        if v != None:
            typ0 = __OLXOBJ_PARA__['SCHEME'][k][2]
            val, se1 = __addOBJ_getVal__(typ0, v)
            #
            if se1 == '' and val != None:
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = __OLXOBJ_PARA__['SCHEME'][k][0]
                ii += 1


    #
    setting1 = dict()
    for k, v in setting.items():
        if type(v) == list:
            setting1[k] = v
    for k, v in setting.items():
        if type(v) != list:
            setting1[k] = v
    #
    signalName = __voidArray__()
    signalType = __intArray__()
    signalVar = __doubleArray1__()
    i1 = -1
    for k1, v1 in setting1.items():
        i1 += 1
        signalName[i1] = cast(pointer(c_char_p(encode3(k1))), c_void_p)
        v1i = __convert2Float__(v1)
        if type(v1i)==float:
            signalType[i1] = c_int(OlxAPIConst.LVT_CONST)
            signalVar[i1] = c_double(v1i)
        elif type(v1) == list and len(v1) == 2 and type(v1[0]) == str:
            v1u = v1[0].upper()
            if v1u not in __OLXOBJ_SCHEME__.keys():
                messError = '\nKEYS SCHEME : '+toString(list(__OLXOBJ_SCHEME__.keys()))+'\n'
                messError += se+'\n\tlogic : {'+toString(k1)+':'+toString(v1)+"}\n\tKey error : "+toString(v1[0])
                return None
            t0 = __OLXOBJ_SCHEME__[v1u]
            signalType[i1] = c_int(t0)
            for ob1 in __OLXOBJ_SCHEME_OB__:
                try:
                    o1 = OLCase.findOBJ(ob1, v1[1])
                    if o1 != None:
                        break
                except:
                    pass
            #
            if o1 is None:
                messError = se+'\n\tlogic : {'+toString(k1)+':'+toString(v1)+"\n\tObject not found : "+toString(v1[1])
                return None
            signalVar[i1] = c_double(o1.__hnd__)
        else:
            messError = se+'\n\tlogic : {'+toString(k1)+':'+toString(v1)+'}'
            messError += '\n\tRequired: int/float or [str,str/obj]'
            messError += '\n\tFound   : '+type(v1).__name__+'  '+toString(v1)
            return None
    #
    signalType[i1+1] = c_int(0)
    signalName[i1+1] = cast(pointer(c_char_p(encode3(''))), c_void_p)
    params[ii] = cast(pointer(signalName), c_void_p)
    tokens[ii] = OlxAPIConst.LS_vsSignalName
    ii += 1
    params[ii] = cast(pointer(signalType), c_void_p)
    tokens[ii] = OlxAPIConst.LS_vnSignalType
    ii += 1
    params[ii] = cast(pointer(signalVar), c_void_p)
    tokens[ii] = OlxAPIConst.LS_vdSignalVar
    ii += 1
    #
    tokens[ii] = 0
    setVerbose(1, 1)
    hnd = OlxAPI.AddDevice(c_int(OlxAPIConst.TC_SCHEME), c_int(v0[0]), c_int(v0[1]), tokens, params)
    if hnd <= 0:
        messError = se+'\n'+ErrorString()
        return None
    #
    o1 = __getOBJ__(hnd, tc=OlxAPIConst.TC_SCHEME)
    #
    if param2:
        o1 = __updateOBJNew__(o1, param2, {}, verbose=False)
    #
    if __OLXOBJ_VERBOSE__ and o1:
        print("\nOlxObj.OLCase.addOBJ('SCHEME')"+" (new) : "+o1.toString())
    return o1


def __addOBJ_RECLSR__(v0, param):
    global messError
    messError = ''
    se = "\nOlxObj.OLCase.addOBJ('RECLSR', key, param)"
    tokens = __intArray__()
    params = __voidArray__()
    #
    param0, paramp, paramg, param2 = {}, {}, {}, {}
    for k, v in param.items():
        if k in __OLXOBJ_PARA__['RECLSR'].keys() and __OLXOBJ_PARA__['RECLSR'][k][2]:
            if k.startswith('PH_'):
                paramp[k] = v
            elif k.startswith('GR_'):
                paramg[k] = v
            else:
                param0[k] = v
        else:
            param2[k] = v  # for UDF
    #
    ii = 0
    if paramp:
        paramp.update(param0)
        if 'TAGS' in paramp.keys():
            paramg['TAGS'] = paramp['TAGS']
        s1, s2 = __check_param_setting__(None, 'RECLSR', paramp, {})
        if s2:
            messError = s1+se+s2
            return None
        #
        for k, v1 in {'PH_MEMO': OlxAPIConst.OBJ_sMemo, 'TAGS': OlxAPIConst.OBJ_sTags, 'GUID': OlxAPIConst.OBJ_sGUID}.items():
            if k in paramp.keys():
                val = c_char_p(encode3(paramp[k]))
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = v1
                ii += 1
                #
                paramp.pop(k)
        #
        for k, v in paramp.items():
            if v != None:
                typ0 = __OLXOBJ_PARA__['RECLSR'][k][2]
                v1 = __convert2Type2__(v, typ0)
                if v1 == None:
                    messError = se+' error'
                    return None
                if typ0 == 'str':
                    val = c_char_p(encode3(v1))
                elif typ0 == 'float':
                    val = c_double(v1)
                elif typ0 == 'int':
                    val = c_int(v1)
                else:
                    messError = se+'\n\tunsupported : '+typ0
                    return None
                #
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = __OLXOBJ_PARA__['RECLSR'][k][0]
                ii += 1
        #
        tokens[ii] = 0
        hnd = OlxAPI.AddDevice(c_int(OlxAPIConst.TC_RECLSRP), c_int(v0[0]), c_int(v0[1]), tokens, params)
        if hnd <= 0:
            messError = se+'\n'+ErrorString()
            return None
    else:
        paramg.update(param0)
    #
    if paramg:
        paramg['ID'] = param['ID']
        s1, s2 = __check_param_setting__(None, 'RECLSR', paramg, {})
        if s2:
            messError = s1+se+s2
            return None
        ii = 0
        tokens = __intArray__()
        params = __voidArray__()
        for k, v1 in {'GR_MEMO': OlxAPIConst.OBJ_sMemo, 'TAGS': OlxAPIConst.OBJ_sTags, 'GUID': OlxAPIConst.OBJ_sGUID, 'ID': OlxAPIConst.CG_sID}.items():
            if k in paramg.keys():
                val = c_char_p(encode3(paramg[k]))
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = v1
                ii += 1
                paramg.pop(k)
        #
        for k, v in paramg.items():
            if v != None:
                typ0 = __OLXOBJ_PARA__['RECLSR'][k][2]
                v1 = __convert2Type2__(v, typ0)
                if v1 == None:
                    messError = se+' error'
                    return None
                if typ0 == 'str':
                    val = c_char_p(encode3(v1))
                elif typ0 == 'float':
                    val = c_double(v1)
                elif typ0 == 'int':
                    val = c_int(v1)
                else:
                    messError = se+'\n\tunsupported : '+typ0
                    return None
                #
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = __OLXOBJ_PARA__['RECLSR'][k][0]
                ii += 1
        #
        tokens[ii] = 0
        hnd = OlxAPI.AddDevice(c_int(OlxAPIConst.TC_RECLSRG), c_int(v0[0]), c_int(v0[1]), tokens, params)
        if hnd <= 0:
            messError = se+'\n'+ErrorString()
            return None
    o1 = None
    try:
        t1 = TERMINAL(hnd=v0[0]).RLYGROUP1
        for r1 in t1.RECLSR:
            if r1.ID == param['ID']:
                o1 = r1
                break
    except:
        messError = se+': ERROR'
        return None
    #
    if param2:
        o1 = __updateOBJNew__(o1, param2, {}, verbose=False)
    if __OLXOBJ_VERBOSE1__ and __OLXOBJ_VERBOSE__ and o1:
        print("\nOlxObj.OLCase.addOBJ('RECLSR')"+" (new) : "+o1.toString())
    return o1


def __addOBJ_getVal__(typ0, v):
    se = ''
    if typ0 == 'str' or typ0 == 'strk':
        val = c_char_p(encode3(v))
    elif typ0 == 'float':
        val = c_double(float(v))
    elif typ0 == 'int':
        val = c_int(int(v))
    elif typ0 == 'RLYGROUP':
        o1 = OLCase.findOBJ(typ0, v)
        val = c_int(o1.__hnd__)
    elif typ0 == 'RLYD':
        try:
            o1 = OLCase.findOBJ(typ0, v)
            val = c_int(o1.__hnd__)
        except:
            val = None
    elif typ0 in {'BUS0', 'BUS'}:
        o1 = OLCase.findOBJ('BUS', v)
        val = c_int(o1.__hnd__)
    elif typ0 == 'int401':
        v = __bin2int__(v)
        val = c_int(v)
    elif typ0 in {'float2', 'float3', 'float4', 'float5', 'float8', 'float10'}:
        val = __doubleArray1__()
        for i in range(int(typ0[5:])):
            val[i] = float(v[i])
    elif typ0 in {'int2', 'int8'}:
        val = __intArray__()
        for i in range(int(typ0[3:])):
            val[i] = __convert2Int__(v[i])
    elif typ0 == 'equip10':
        val = __intArray__()
        va, _ = __getequip10__(v)
        for i in range(len(va)):
            val[i] = va[i].__hnd__
    elif typ0 == 'RLYGROUPn':
        try:
            val = __intArray__()
            va, _ = __getRLYGROUPn__(v)
            for i in range(len(va)):
                val[i] = va[i].__hnd__
        except:
            val = None
    else:
        val,se = None,'unsupported %s please contact: support@aspeninc.com'%str(typ0)
    return val, se


def __addOBJ_getVal1__(typ0, v):
    se = ''
    if typ0 == 'str':
        val = c_char_p(encode3(v))
    elif typ0 == 'float':
        val = c_double(v)
    elif typ0 == 'int':
        val = c_int(v)
    else:
        se = 'unsupported %s please contact: support@aspeninc.com' % typ0
    return val, se


def __addOBJ__(ob, v0, param, setting):
    global messError
    messError = ''
    se = "\nOlxObj.OLCase.addOBJ('%s'" % ob
    se += ', key, param' if param else ''
    se += ', setting)' if setting else ')'
    tokens = __intArray__()
    params = __voidArray__()
    tc = __OLXOBJ_CONST__[ob][0]
    #
    flagSlack = False
    setVerbose(0, 1)
    if ob in {'GENW3', 'GENW4', 'CCGEN', 'GEN'}:
        b1 = BUS(hnd=v0[0][1])
        for v1 in ['GENW3', 'GENW4', 'CCGEN', 'GEN']:
            o1 = OLCase.findOBJ(v1, b1)
            if o1 != None:
                messError = se+":failure\nThere's already a generator (%s) on this BUS" % v1
                messError += '\n\t'+o1.toString()
                return None
    elif ob == 'GENUNIT':
        b1 = BUS(hnd=v0[0][1])
        for v1 in ['GENW3', 'GENW4', 'CCGEN']:
            o1 = OLCase.findOBJ(v1, b1)
            if o1 != None:
                messError = se+":failure\nThere's already a generator (%s) on this BUS" % v1
                messError += '\n\t'+o1.toString()
                return None
    elif ob == 'BUS':
        if 'SLACK' in param.keys() and int(param['SLACK']) > 0:
            param['SLACK'] = 0
            flagSlack = True
    #
    param1, param2 = {}, {}
    for k, v in param.items():
        if k in __OLXOBJ_PARA__[ob].keys() and __OLXOBJ_PARA__[ob][k][2]:
            param1[k] = v
        else:
            param2[k] = v  # for UDF
    #
    s1, s2 = __check_param_setting__(None, ob, param1, setting)
    if s2:
        messError = s1+se+s2
        return None
    #
    ii = 0
    if ob not in __OLXOBJ_RELAY2__:
        for i in range(len(v0)):
            vi = v0[i][1]
            val, se1 = __addOBJ_getVal1__(v0[i][2], vi)
            if se1:
                messError = se+se1
                return None
            if val != None:
                params[i] = cast(pointer(val), c_void_p)
                tokens[i] = __OLXOBJ_PARA__[ob][v0[i][0]][0]
        ii = len(v0)
    #
    __addOBJ_default__(ob, param1, setting, v0)
    #
    for k, v1 in {'MEMO': OlxAPIConst.OBJ_sMemo, 'TAGS': OlxAPIConst.OBJ_sTags, 'GUID': OlxAPIConst.OBJ_sGUID}.items():
        if k in param1.keys():
            val = c_char_p(encode3(param1[k]))
            params[ii] = cast(pointer(val), c_void_p)
            tokens[ii] = v1
            ii += 1
            param1.pop(k)
    #
    if ob in {'XFMR', 'XFMR3'}:
        if 'LTCCTRL' in param1.keys() and param1['LTCCTRL'] is None:
            param1['LTCSIDE'] = 0
        try:
            side = int(param1['LTCSIDE'])
        except:
            side = 0
        if 'LTCSIDE' in param1.keys() and side > 0 and 'LTCCTRL' not in param1.keys():
            if side == 1:
                param1['LTCCTRL'] = BUS(hnd=v0[0][1])
            elif side == 2:
                param1['LTCCTRL'] = BUS(hnd=v0[1][1])
            elif side == 3 and ob == 'XFMR3':
                param1['LTCCTRL'] = BUS(hnd=v0[2][1])
    #
    for k, v in param1.items():
        if v != None:
            typ0 = __OLXOBJ_PARA__[ob][k][2]
            val, se1 = __addOBJ_getVal__(typ0, v)
            if se1:
                messError = se+se1
                return None
            #
            if val != None:
                params[ii] = cast(pointer(val), c_void_p)
                tokens[ii] = __OLXOBJ_PARA__[ob][k][0]
                ii += 1
    #
    if ob in {'RLYDSG', 'RLYDSP'}:
        sSetting = ''
        for k, v in setting.items():
            sSetting += k+'\t'+str(v)+'\t'
        if sSetting:
            params[ii] = cast(pointer(c_char_p(encode3(sSetting))), c_void_p)
            tokens[ii] = __OLXOBJ_SETTING__[ob][3]
            ii += 1
    elif ob in {'RLYOCG', 'RLYOCP'}:
        val = __doubleArray__()
        polar = param['POLAR']
        na = __OLXOBJ_OCSETTING__[polar][0]
        ida = __OLXOBJ_OCSETTING__[polar][1]
        for k, v in setting.items():
            for i in range(len(na)):
                if k.upper() == na[i].upper():
                    val[ida[i]] = float(v)
                    break
        params[ii] = cast(pointer(val), c_void_p)
        tokens[ii] = __OLXOBJ_SETTING__[ob][4]
        ii += 1
    #
    tokens[ii] = 0
    setVerbose(1, 1)
    if ob in __OLXOBJ_RELAY2__:
        hnd = OlxAPI.AddDevice(c_int(tc), c_int(v0[0]), c_int(v0[1]), tokens, params)
    else:
        hnd = OlxAPI.AddEquipment(c_int(tc), tokens, params)
    #
    if hnd <= 0:
        messError = se+'\n'+ErrorString()
        return None
    #
    o1 = __getOBJ__(hnd, tc=tc)
    #
    if param2:
        o1 = __updateOBJNew__(o1, param2, {}, verbose=False)
    #
    if flagSlack and o1:
        g = o1.GEN
        if g != None and g.REG == 0 and g.ILIMIT1 <= 0 and g.ILIMIT2 <= 0:
            o1.changeData('SLACK', 1)
    #
    if __OLXOBJ_VERBOSE1__ and __OLXOBJ_VERBOSE__ and o1:
        print("\nOlxObj.OLCase.addOBJ('"+ob+"')"+" (new) : "+o1.toString())
    return o1


def __addOBJ_default__(ob, param1, setting1, v0):
    if __OLXOBJ_NEWADD_DEFAULT__:
        kp = set(param1.keys())
        if ob in {'LINE', 'XFMR', 'SERIESRC'}:
            if 'R' not in kp and 'X' not in kp:
                param1['X'] = 0.1
            #
            if ob != 'SERIESRC' and 'R0' not in kp and 'X0' not in kp:
                param1['X0'] = 0.1
            #
            if ob == 'XFMR':
                if 'PRITAP' not in kp:
                    param1['PRITAP'] = BUS(hnd=v0[0][1]).kV
                if 'SECTAP' not in param1.keys():
                    param1['SECTAP'] = BUS(hnd=v0[1][1]).kV
        elif ob == 'XFMR3':
            if 'RPS' not in kp and 'XPS' not in kp:
                param1['XPS'] = 0.1
            if 'RPT' not in kp and 'XPT' not in kp:
                param1['XPT'] = 0.1
            if 'RST' not in kp and 'XST' not in kp:
                param1['XST'] = 0.1
            if 'RPS0' not in kp and 'XPS0' not in kp:
                param1['XPS0'] = 0.1
            if 'RPT0' not in kp and 'XPT0' not in kp:
                param1['XPT0'] = 0.1
            if 'RST0' not in kp and 'XST0' not in kp:
                param1['XST0'] = 0.1
            if 'PRITAP' not in kp:
                param1['PRITAP'] = BUS(hnd=v0[0][1]).kV
            if 'SECTAP' not in param1.keys():
                param1['SECTAP'] = BUS(hnd=v0[1][1]).kV
            if 'TERTAP' not in param1.keys():
                param1['TERTAP'] = BUS(hnd=v0[2][1]).kV
        elif ob == 'SHIFTER':
            if 'RP' not in kp and 'XP' not in kp:
                param1['XP'] = 0.1
            if 'RN' not in kp and 'XN' not in kp:
                param1['XN'] = 0.1
            if 'RZ' not in kp and 'XZ' not in kp:
                param1['XZ'] = 0.1
        elif ob == 'GENW3':
            if 'MVA' not in kp:
                param1['MVA'] = 0.01
            if 'MWR' not in kp:
                param1['MWR'] = 0.01
        elif ob == 'GENW4':
            if 'MVA' not in kp:
                param1['MVA'] = 0.01
        elif ob == 'CCGEN':
            if 'I' not in kp:
                param1['I'] = [0.001, 0.001, 0, 0, 0, 0, 0, 0, 0, 0]
            if 'V' not in kp:
                param1['V'] = [1.0, 1.1, 0, 0, 0, 0, 0, 0, 0, 0]
        elif ob == 'GENUNIT':
            if 'R' not in kp and 'X' not in kp:
                param1['X'] = [0.1, 0.1, 0.1, 0.1, 0.1]
        elif ob in {'RLYOCG', 'RLYOCP'}:
            if 'POLAR' not in kp:
                param1['POLAR'] = 0


def __bin2int__(value):
    sv = ''
    for v1 in value:
        sv = str(v1)+sv
    return int(sv, 2)


def __brIsInScope__(ba, areaNum, zoneNum, optionTie, kVa):
    if optionTie == 0:  # 0-strictly in areaNum/zoneNum
        for i in range(len(ba)):
            if not __busIsInScope__(ba[i], areaNum, zoneNum, kVa[i]):
                return False
        return True
    elif optionTie == 1:  # 1- with tie
        for i in range(len(ba)):
            if __busIsInScope__(ba[i], areaNum, zoneNum, kVa[i]):
                return True
        return False
    elif optionTie == 2:  # 2- only tie
        haveT, haveF = False, False
        for i in range(len(ba)):
            if __busIsInScope__(ba[i], areaNum, zoneNum, kVa[i]):
                haveT = True
            else:
                haveF = True
        return haveT and haveF
    else:
        raise ValueError('Error value of optionTie = [0,1,2]')


def __busIsInScope__(b1, areaNum, zoneNum, kV):
    """ check if BUS is in scope:
        areaNum []: List of area Number
        zoneNum []: List of zone Number
        kV      []: [kVmin,kVmax] """
    if areaNum:
        if b1.AREANO not in areaNum:
            return False
    if zoneNum:
        if b1.ZONENO not in zoneNum:
            return False
    if kV:
        kv1 = b1.KV
        if kv1 < kV[0] or kv1 > kV[1]:
            return False
    return True


def __changeSetting__(o1, sSetting, value):
    global messError
    messError = ''
    if o1.__ob__ in {'RLYDSG', 'RLYDSP'}:
        val = sSetting+'\t'+value
        val = c_char_p(encode3(val))
        if OLXAPI_OK != SetData(c_int(o1.__hnd__), c_int(__OLXOBJ_SETTING__[o1.__ob__][3]), byref(val)):
            messError = '\n'+ErrorString()
        return
    #
    polar = o1.__paramEx__['POLAR']
    name = __OLXOBJ_OCSETTING__[polar][0]
    ida = __OLXOBJ_OCSETTING__[polar][1]
    for i in range(len(ida)):
        if sSetting == name[i].upper():
            id1 = ida[i]
            break
    #
    valo = [0]*9
    try:
        valo[id1] = float(value)
    except:
        messError = '\nValueError : '+toString(value)
        return
    #
    paracode = __OLXOBJ_SETTING__[o1.__ob__][4]
    #
    val1 = __setValue__(o1.__hnd__, paracode, valo)
    if OLXAPI_OK != SetData(c_int(o1.__hnd__), c_int(0x10000*(id1+1)+paracode), byref(val1)):
        messError = '\n'+ErrorString()


def __checkFault__(sf):
    global messError
    messError = ''
    if sf.__index_simul__ != __INDEX_SIMUL__:
        messError = '\nAll previous RESULT_FLT are not available'
        messError += '\n\t   by a SEA simulation or Classical/Simultaneous with clearPrev=1 has been done'
        messError += '\n\tor by OLCase.close()'
        messError += '\n\tor by OLCase.open() to open a other OLR file'
    return messError


def __check00__(typc, key):
    se = ''
    if typc == 'typeNAMEKV':  # [NAME,kV]
        if not (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi):
            se += '\n\tRequired           : [str,float/int]'
            if key is None:
                se += '\n\tFound (ValueError) : None'
            else:
                if type(key) == list:
                    se += '\n\tFound (ValueError) : (len=%i) ' % len(key)+__getStr__([key])[1]+'  '+toString(key)
                else:
                    se += '\n\tFound (ValueError) : '+ __getStr__([key])[1]+'  '+toString(key)
        return se
    #
    elif typc == 'typeBUS0':  # None|BUS|str|int|[str,f_i]
        if key != None and (not (type(key) in {int, str, BUS} or (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi))):
            se += '\n\tRequired           : None | [str,float/int] or str or int or BUS'
            se += '\n\tFound (ValueError) : '+__getStr__(key)[1]+'  '+toString(key)
        if type(key) == BUS and __check_currFileIdx__(key):
            se = messError
        return se
    #
    elif typc == 'typeBUS1':  # BUS|str|int|[str,f_i]
        if not (type(key) in {int, str, BUS} or (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi)):
            se += '\n\tRequired           : [str,float/int] or str or int or BUS'
            if type(key) == list:
                se += '\n\tFound (ValueError) : (len=%i)' % len(key) +                     ' '+__getStr__([key])[1]+'  '+toString(key)
            else:
                se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    #
    elif (typc == 'str' or typc == 'strk'):
        if type(key) not in {str, bytes}:
            se += '\n\tRequired           : str'
            se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    #
    elif typc == 'int':
        k1 = __convert2Int__(key)
        if type(k1) != int:
            se += '\n\tRequired           : int'
            se += '\n\tFound (ValueError) : '+__getStr__([k1])[1]+'  '+toString(k1)
        return se
    elif typc == 'int2':
        return __testI__(key, 2)
    elif typc == 'int401':
        key = __convert2Int__(key)
        if type(key) != list or len(key) != 4 or key[0] not in {0, 1} or key[1] not in {0, 1} or key[2] not in {0, 1} or key[3] not in {0, 1}:
            se = '\n\tRequired           : [int 0/1]*4'
            se += '\n\tFound (ValueError) : '+toString(key)
            return se
        return ''
    elif typc == 'int8':
        return __testI__(key, 8)
    elif typc == 'float':
        k1 = __convert2Float__(key)
        if type(k1) not in __fi:
            se += '\n\tRequired           : int/float'
            se += '\n\tFound (ValueError) : '+__getStr__([k1])[1]+'  '+toString(k1)
        return se
    elif typc == 'float2':
        return __testFI__(key, 2)
    elif typc == 'float3':
        return __testFI__(key, 3)
    elif typc == 'float4':
        return __testFI__(key, 4)
    elif typc == 'float5':
        return __testFI__(key, 5)
    elif typc == 'float8':
        return __testFI__(key, 8)
    elif typc == 'float10':
        return __testFI__(key, 10)
    elif typc == 'typeBRCODE':
        if type(key) != str or key.upper() not in {'X', 'T', 'P', 'L', 'DC', 'S', 'W'}:
            se += "\n\tRequired           : str in ['X','T','P','L','DC','S','W']"
            se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    elif typc == 'equip10':
        if type(key) != list or len(key) > 10:
            se += '\n\tRequired           : list of 10 equipments'
            se += '\n\tFound (ValueError) : '+__getStr__(key)[1]+'  '+toString(key)
            return se
        #
        _, se = __getequip10__(key)
        return se
    elif typc == 'RLYGROUPn':
        t = type(key) == list
        if t:
            for key1 in key:
                if not (type(key1) == RLYGROUP or (type(key1) == str and key1.startswith('[RELAYGROUP] '))):
                    t = False
                    break
        if not t:
            se += '\n\tRequired           : list of RLYGROUP or str of RLYGROUP '
            se += '\n\tFound (ValueError) : '+ __getStr__(key)[1]+'  '+toString(key)
            return se
        return se
    if typc == 'RLYD' and key is None:
        return se
    elif typc in {'RLYGROUP', 'BUS', 'RLYD'}:
        try:
            v1 = OLCase.findOBJ(typc, key)
            if v1 is None:
                se = '\n\t%s(' % typc+toString(key)+') : not Found'
        except Exception as err:
            se = str(err)
        return se
    elif typc == 'BUS0':
        if key is None:
            return ''
        return __check00__('BUS', key)
    elif typc == 'b1b2id':
        try:
            BUS(key[0]), BUS(key[1])
            key[2] + ''
            if len(key)==3:
                return ''
        except:
            pass
        se += '\n\tRequired           : [BUS1, BUS2, CID] BUS = [name, kV] or busNumber or str'
        se += '\n\tFound (ValueError) : '+ __getStr__(key)[1]+'  '+toString(key)
        return se
    else:
        se = 'unsupport typc : '+typc
    return se


def __getequip10__(key):
    setVerbose(0, 1)
    res = []
    for i in range(len(key)):
        o1 = None
        for ob in ['XFMR', 'XFMR3', 'LINE', 'SERIESRC', 'SWITCH', 'SHIFTER', 'GENUNIT', 'DCLINE2', 'GEN', 'LOAD', 'LOADUNIT', 'SHUNT', 'SHUNTUNIT', 'BUS', 'SVD']:
            try:
                o1 = OLCase.findOBJ(ob, key[i])
                if o1 != None:
                    break
            except:
                pass
        if o1 is None:
            return [], '\nObject not Found [%i] : ' % i+toString(key[i])
        res.append(o1)
    setVerbose(1, 1)
    return res, ''


def __getRLYGROUPn__(key):
    setVerbose(0, 1)
    res = []
    for i in range(len(key)):
        try:
            o1 = OLCase.findOBJ('RLYGROUP', key[i])
        except:
            o1 = None
        if o1 is None:
            return [], '\nRLYGROUP not Found [%i] : ' % i+toString(key[i])
        res.append(o1)
    setVerbose(1, 1)
    return res, ''


def __check1_Outage__(ot):
    global messError
    messError = '\nOUTAGE(option,G)'
    option = ot.OPTION
    if type(option) != str or option.upper() not in {'SINGLE', 'SINGLE-GND', 'DOUBLE', 'ALL'}:
        messError += '\n\toption   : (str) Branch outage option code'
        messError += '\n\tRequired : (str)'
        messError += '\n\t        SINGLE     : one at a time'
        messError += '\n\t        SINGLE-GND : (LINE only) one at a time with ends grounded'
        messError += '\n\t        DOUBLE     : two at a time'
        messError += '\n\t        ALL        : all at once'
        messError += '\n\t'+__getErrValue__(int, option)
        return True
    option = option.upper()
    G = ot.G
    if (type(G) not in __fi or G < 0) and option == 'SINGLE-GND':
        messError += '\n  G : Admittance Outage line grounding (mho) (option=SINGLE-GND)'
        messError += '\n\tRequired: >0'
        messError += '\n\t'+__getErrValue__(float, G)
        return True
    return False


def __check2_Outage__(ot, obj, tiers, wantedTypes):
    global messError
    messError = '\nOUTAGE.build_outageLst(obj,tiers,wantedTypes)'
    if obj is None or type(obj) not in {BUS, RLYGROUP, LINE, XFMR, XFMR3, SERIESRC, SHIFTER, SWITCH, TERMINAL}:
        messError += '\n\tobj                : Object for Outage option'
        messError += '\n\tRequired           : BUS or TERMINAL or RLYGROUP or EQUIPMENT (LINE,XFMR,XFMR3,SERIESRC,SHIFTER,SWITCH)'
        messError += '\n\t'+__getErrValue__('', obj)
        return True
    if __check_currFileIdx__(obj):
        return True
    messError = '\nOUTAGE.build_outageLst(obj,tiers,wantedTypes)'
    if tiers is None or type(tiers) != int or tiers < 0:
        messError += '\n\ttiers              : Number of tiers-neighboring'
        messError += '\n\tRequired           : (int)>0'
        messError += '\n\t'+__getErrValue__(int, tiers)
        return True
    flag = type(wantedTypes) == list
    if flag:
        for w1 in wantedTypes:
            if type(w1) != str or w1.upper() not in __OLXOBJ_EQUIPMENTB4__:
                flag = False
    if not flag or wantedTypes is None:
        messError += '\n\twantedTypes        : Branch type to consider'
        messError += "\n\tRequired           : (list)[(str/obj)] in "+str(__OLXOBJ_EQUIPMENTB2__)[1:-1].replace(' ', '')
        messError += "\n\t                                    or in "+str(__OLXOBJ_EQUIPMENTB3__)[1:-1].replace(' ', '')
        messError += '\n\tFound (ValueError) : '+str(wantedTypes)
        return True
    return False


def __check01__(typc, key):
    se = ''
    if typc == 'typeNAMEKV':  # [NAME,kV]
        __convert2Type1__(key,1)
        if not (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi):
            se += '\n\tRequired           : [str,float/int]'
            if key is None:
                se += '\n\tFound (ValueError) : None'
            else:
                if type(key) == list:
                    se += '\n\tFound (ValueError) : (len=%i) ' % len(key)+__getStr__([key])[1]+'  '+toString(key)
                else:
                    se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    #
    elif typc == 'typeBUS0':  # None|BUS|str|int|[str,f_i]
        __convert2Type1__(key,1)
        if key != None and (not (type(key) in {int, str, BUS} or (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi))):
            se += '\n\tRequired           : None | [str,float/int] or str or int or BUS'
            se += '\n\tFound (ValueError) : '+__getStr__(key)[1]+'  '+toString(key)
        if type(key) == BUS and __check_currFileIdx__(key):
            se = messError
        return se
    #
    elif typc == 'typeBUS1':  # BUS|str|int|[str,f_i]
        __convert2Type1__(key,1)
        if not (type(key) in {int, str, BUS} or (type(key) == list and len(key) == 2 and type(key[0]) == str and type(key[1]) in __fi)):
            se += '\n\tRequired           : [str,float/int] or str or int or BUS'
            if type(key) == list:
                se += '\n\tFound (ValueError) : (len=%i)' % len(key)+' '+__getStr__([key])[1]+'  '+toString(key)
            else:
                se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    #
    elif typc == 'typeSTR':  # str
        if type(key) not in {str, int}:
            se += '\n\tRequired           : str'
            se += '\n\tFound (ValueError) : '+ __getStr__([key])[1]+'  '+toString(key)
        return se
    elif typc == 'typeBRCODE':
        if type(key) != str or key.upper() not in __OLXOBJ_EQUIPMENTB4__:
            se = '\n\tRequired           : str in '+str(__OLXOBJ_EQUIPMENTB2__)[1:-1].replace(' ', '')
            se += '\n\t                      or in '+str(__OLXOBJ_EQUIPMENTB3__)[1:-1].replace(' ', '')
            se += '\n\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+toString(key)
        return se
    return se


def __check02__(ob, key, fn):  # check for [b1,b2,CID] or STR
    if type(key) in {str, ob} and 'addOBJ' not in fn:
        return ''
    if type(key) == list and len(key) != 3:
        if 'addOBJ' not in fn:
            se = '\n\tRequired           : str or [b1,b2,CID]'
        else:
            se = '\n\tRequired           : [b1,b2,CID] '
        se += '\n\tFound (ValueError) : (len=%i)' % len(key)+' '+toString(key)
        return se
    if type(key) != list:
        if 'addOBJ' not in fn:
            se = '\n\tRequired           : str or [b1,b2,CID]'
        else:
            se = '\n\tRequired           : [b1,b2,CID]'
        se += '\n\tFound (ValueError) : (%s)' % type(key).__name__
        return se
    #
    for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeSTR', 'CID']]:
        __convert2Type1__(key,2,str)
        s1 = __check01__(v[1], key[v[0]])
        if s1:
            return '\n\t'+v[2]+' = '+toString(key[v[0]])+s1
    return ''


def __check021__(ob, key, fn):# check for XFMR3: [b1,b2,CID]/[b1,b2,b3,CID] or STR, no addOBJ
    if type(key) in {str, ob} and 'addOBJ' not in fn:
        return ''
    #
    if type(key) == list and not (len(key) == 3 or len(key) == 4) and 'addOBJ' not in fn:
        se = '\n\tRequired           : str or [b1,b2,CID] or [b1,b2,b3,CID]'
        se += '\n\tFound (ValueError) : (len=%i)' % len(key)+' '+toString(key)
        return se
    #
    if type(key) == list and (len(key) != 4) and 'addOBJ' in fn:
        se = '\n\tRequired           : [b1,b2,b3,CID] '
        se += '\n\tFound (ValueError) : (len=%i)' % len(key)+' '+toString(key)
        return se
    #
    if type(key) != list:
        if 'addOBJ' not in fn:
            se = '\n\tRequired           : str or [b1,b2,CID] or [b1,b2,b3,CID]'
        else:
            se = '\n\tRequired           : [b1,b2,b3,CID] '
        se += '\n\tFound (ValueError) : (%s)' % type(key).__name__
        return se
    #
    if len(key) == 3:
        __convert2Type1__(key,2,str)
        for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeSTR', 'CID']]:
            s1 = __check01__(v[1], key[v[0]])
            if s1:
                return '\n\t'+v[2]+' = '+toString(key[v[0]])+s1
    if len(key) == 4:
        __convert2Type1__(key,3,str)
        for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeBUS1', 'b3'], [3, 'typeSTR', 'CID']]:
            s1 = __check01__(v[1], key[v[0]])
            if s1:
                return '\n\t'+v[2]+' = '+toString(key[v[0]])+s1
    return ''


def __check03__(ob, key, fn):# check for [b1,CID] for GENUNIT,LOADUNIT,SHUNTUNIT
    if 'addOBJ' not in fn:
        if type(key) == str:
            return ''
    if type(key) != list:
        se = '\n\tRequired           : [b1,CID] with b1 = [str,float/int] or str or int or BUS or '+ob[:-4]
        se += '\n\tFound (ValueError) : (%s)' % type(key).__name__
        return se
    if len(key) not in [2, 3]:
        se = '\n\tRequired           : [b1,CID] with b1 = [str,float/int] or str or int or BUS or '+ob[:-4]
        se += '\n\tFound (ValueError) : (list) len=%i' % len(key)
        return se
    __convert2Type1__(key,-1,str)
    if type(key[-1]) != str:
        se = '\n\tRequired               : [b1,CID] with CID=str'
        se += '\n\tFound (ValueError) CID : (%s)' % type(key[-1]).__name__
        return se
    if len(key) == 3:
        __convert2Type1__(key, 1)
        if (type(key[0]) != str or type(key[1]) not in __fi):
            se = '\n\tRequired           : [b1,CID] with b1 = [str,float/int] or str or int or BUS or '+ob[:-4]
            se += '\n\tFound (ValueError) : '+ __getStr__([key])[1]+'  '+__getStr__([key])[0]
            return se
    if len(key) == 2:
        va = {int, str, BUS, __OLXOBJ_OBJECT__[ob[:-4]]}
        __convert2Type1__(key[0],1)
        if type(key[0]) in va or (type(key[0]) == list and len(key[0]) == 2 and type(key[0][0]) == str and type(key[0][1]) in __fi):
            return ''
        se = '\n\tRequired             : [b1,CID] with b1 = [str,float/int] or str or int or BUS or '+ob[:-4]
        se += '\n\tFound (ValueError)b1 : '+ __getStr__([key[0]])[1]+'  '+__getStr__([key[0]])[0]
        return se
    return ''


def __check04__(typc, key):  # check for MULINE
    se = ''
    if typc == 'typeLINE1':
        if type(key) in {str, LINE}:
            return se
        if not (type(key) == list and len(key) == 3):
            se += '\n\tRequired           : str or LINE or [bus1,bus2,CID]'
            se += '\n\tFound (ValueError) : '+str(type(key).__name__)+'  '+toString(key)
            return se
        for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeSTR', 'CID']]:
            s1 = __check01__(v[1], key[v[0]])
            if s1:
                return ' = [b1,b2,CID]\n\t'+v[2]+' = '+toString(key[v[0]])+s1
    return se


def __check05__(ob, key, fn):  # check for [b1,name] for BREAKER
    if 'addOBJ' not in fn:
        if type(key) == str:
            return ''
    if type(key) != list:
        se = '\n\tRequired           : [b1,BreakerName] with b1 = [str,float/int] or str or int or BUS'
        se += '\n\tFound (ValueError) : (%s)' % type(key).__name__
        return se
    if len(key) not in [2, 3]:
        se = '\n\tRequired           : [b1,BreakerName] with b1 = [str,float/int] or str or int or BUS'
        se += '\n\tFound (ValueError) : (list) len=%i' % len(key)
        return se

    if type(key[-1]) != str:
        se = '\n\tRequired                       : [b1,BreakerName] with BreakerName=str'
        se += '\n\tFound (ValueError) BreakerName : (%s)' % type(key[-1]).__name__
        return se
    if len(key) == 3:
        __convert2Type1__(key, 1)
        if (type(key[0]) != str or type(key[1]) not in __fi):
            se = '\n\tRequired           : [b1,CID] with b1 = [str,float/int] or str or int or BUS'
            se += '\n\tFound (ValueError) : '+__getStr__([key])[1]+'  '+__getStr__([key])[0]
            return se
    if len(key) == 2:
        __convert2Type1__(key[0], 1)
        va = {int, str, BUS}
        if type(key[0]) in va or (type(key[0]) == list and len(key[0]) == 2 and type(key[0][0]) == str and type(key[0][1]) in __fi):
            return ''
        se = '\n\tRequired             : [b1,CID] with b1 = [str,float/int] or str or int or BUS'
        se += '\n\tFound (ValueError)b1 : '+__getStr__([key[0]])[1]+'  '+__getStr__([key[0]])[0]
        return se
    return ''


def __check_ob_key__(fn, ob, key, param={}, setting={}):
    global messError
    messError = ''
    if 'addOBJ' in fn:
        se = fn+"('%s'" % ob+", key"
        se += ', param' if param else ''
        if ob == 'SCHEME':
            se += ', logic)'
        else:
            se += ', setting)' if setting else ')'
    elif 'findOBJ' in fn:
        se = fn+"('%s', key)" % ob
    else:
        se = fn+"%s(key)" % ob
    if 'OLCase.addOBJ' in fn:
        if type(ob) != str or ob.upper() not in __OLXOBJ_LIST__:
            messError = se+"\n\tValueError with ob = "+toString(ob)+' not in\n'+toString(__OLXOBJ_LIST__[:-1])
            return True
    else:
        if type(ob) != str or ob.upper() not in __OLXOBJ_OBJECT__.keys():
            messError = se+"\n\tValueError with ob = "+toString(ob)+' not in\n'
            messError += toString(list(__OLXOBJ_OBJECT__.keys()))
            return True
    if type(param) != dict:
        messError = se+'\n\tValueError param = '+toString(param)
        messError += '\n\tRequired           : (dict) '
        messError += '\n\tFound (ValueError) : (%s)' % type(param).__name__
        return True

    if type(setting) != dict:
        if ob == 'SCHEME':
            messError = se+'\n\tValueError logicvar = '+toString(setting)
        else:
            messError = se+'\n\tValueError setting = '+toString(setting)
        messError += '\n\tRequired           : (dict) '
        messError += '\n\tFound (ValueError) : (%s)' % type(setting).__name__
        return True

    # check key
    if setting and (ob not in __OLXOBJ_RELAY3__ and ob != 'SCHEME'):
        messError = se+'\n\tValueError: setting works only with RLYOCG,RLYOCP,RLYDSG,RLYDSP,SCHEME'
        return True
    if type(key) in __OLXOBJ_OBJECT__.values() and type(key).__name__ == ob:
        return False
    #
    if ob == 'BUS':
        if 'addOBJ' in fn:
            s1 = __check01__('typeNAMEKV', key)
        else:
            s1 = __check01__('typeBUS1', key)
        if s1:
            messError = se+'\n\twith key = '+toString(key)+s1
            return True
    #
    if ob in __OLXOBJ_BUS1__:
        if 'addOBJ' in fn:
            try:
                BUS(key)
            except Exception as err:
                messError = se+'\n\twith key = '+toString(key)+'\n\t'+str(err)
                return True
        else:
            s1 = __check01__('typeBUS1', key)
            if s1:
                messError = se+'\n\twith key = '+toString(key)+s1
                return True
    #
    if ob in __OLXOBJ_BUS2__:
        s1 = __check03__(ob, key, fn)
        if s1:
            messError = se+' \n\twith key = '+str(key)+s1
            return True
        #
        if 'addOBJ' in fn:
            try:
                if len(key) == 3:
                    g1 = OLCase.findOBJ(ob[:-4], key[:-1])
                else:
                    g1 = OLCase.findOBJ(ob[:-4], key[0])
            except:
                g1 = None
            #
            if g1 is None:
                try:
                    g1 = OLCase.findOBJ('BUS', key[:-1]) if len(key) == 3 else OLCase.findOBJ('BUS', key[0])
                except:
                    g1 = None
            if g1 is None:
                messError = se+' \n\twith key = '+str(key)+messError
                return True
    #
    if ob == 'BREAKER':
        s1 = __check05__(ob, key, fn)
        if s1:
            messError = se+' \n\twith key = '+str(key)+s1
            return True
        if 'addOBJ' in fn:
            g1 = OLCase.findOBJ('BUS', key[0]) if len(key) <= 2 else OLCase.findOBJ('BUS', key[:-1])
            if g1 is None:
                messError = se+' \n\twith key = '+str(key)+messError
                return True
        return False
    #
    if ob == 'XFMR3':
        s1 = __check021__(__OLXOBJ_OBJECT__[ob], key, fn)
        if s1:
            if 'addOBJ' in fn:
                messError = se+' \n\twith key = [b1,b2,b3,CID] = '+toString(key)+s1
            else:
                messError = se+' \n\twith key = [b1,b2,CID] or [b1,b2,b3,CID] = '+toString(key)+s1
            return True

        if 'addOBJ' in fn:
            try:
                BUS(key[0]), BUS(key[1]), BUS(key[2])
            except Exception as err:
                messError = se+'\n\twith key = '+toString(key)+'\n\t'+str(err)
                return True
    #
    elif ob in __OLXOBJ_EQUIPMENT__:
        s1 = __check02__(__OLXOBJ_OBJECT__[ob], key, fn)
        if s1:
            messError = se+' \n\twith key = [b1,b2,CID] = '+toString(key)+s1
            return True
        #
        if 'addOBJ' in fn:
            try:
                b1, b2 = BUS(key[0]), BUS(key[1])
            except Exception as err:
                messError = se+'\n\twith key = '+toString(key)+'\n\t'+str(err)
                return True
            #
            if ob in {'LINE', 'SWITCH', 'SERIESRC'} and b1.kV != b2.kV:
                messError = se+'\n\twith key = '+toString(key)
                messError += '\n\t2 buses with different nominal voltage:'
                messError += '\n\t\tb1 :'+toString(b1)
                messError += '\n\t\tb2 :'+toString(b2)
                return True
    #
    elif ob in {'RLYGROUP', 'TERMINAL'}:
        if 'addOBJ' not in fn:
            if type(key) == str:
                return False
        #
        flag = type(key) == list and len(key) != 4
        if not flag:
            flag = type(key) != list
        if flag:
            messError = se+' key = '+toString(key)
            if 'addOBJ' in fn:
                messError += '\n\tRequired           : [b1,b2,CID,BRCODE]'
            else:
                messError += '\n\tRequired           : str or [b1,b2,CID,BRCODE]'
            messError += '\n\tFound (ValueError) : len='+str(len(key))+' '+toString(key)
            return True
        #
        for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeSTR', 'CID'], [3, 'typeBRCODE', 'BRCODE']]:
            s1 = __check01__(v[1], key[v[0]])
            if s1:
                messError = se+' key = [b1,b2,CID,BRCODE]\n\t'+v[2]+' = '+toString(key[v[0]])+s1
                return True
        #
        if 'addOBJ' in fn and ob == 'RLYGROUP':
            try:
                TERMINAL(key)
            except Exception as err:
                messError = se+' key = '+toString(key)+'\t'+str(err)
                return True
    #
    elif ob in __OLXOBJ_RELAY__:
        if type(key) in {str, ob} and 'addOBJ' not in fn:
            return False
        if type(key) == list and len(key) != 5:
            messError = se+'\n\tRequired key       : len=5 [b1,b2,CID,BRCODE,ID] or str'
            messError += '\n\tFound (ValueError) : len='+str(len(key))+' '+toString(key)
            return True
        if type(key) != list:
            messError = se+'\n\tRequired key       : len=5 [b1,b2,CID,BRCODE,ID] or str'
            messError += '\n\tFound (ValueError) : '+str(type(key).__name__)+'  '+toString(key)
            return True
        #
        if type(key[2]) in {int, str, bytes}:
            key[2] = str(key[2])
        for v in [[0, 'typeBUS1', 'b1'], [1, 'typeBUS1', 'b2'], [2, 'typeSTR', 'CID'], [3, 'typeBRCODE', 'BRCODE'], [4, 'typeSTR', 'ID']]:
            s1 = __check01__(v[1], key[v[0]])
            if s1:
                messError = se +'\n\tkey                : len=5 [b1,b2,CID,BRCODE,ID] = '+toString(key)
                messError += '\n\t'+v[2]+s1
                return True
        #
        if 'addOBJ' in fn:
            try:
                TERMINAL(key[:-1])
            except Exception as err:
                messError = se+'\nkey = '+toString(key)+str(err)
                return True
    #
    elif ob == 'MULINE':
        if type(key) in {str, ob} and 'addOBJ' not in fn:
            return False
        if type(key) == list and len(key) != 2:
            messError = se+'\n\tRequired key       : len=2 [l1,l2]'
            messError += '\n\tFound (ValueError) : len='+str(len(key))+' '+toString(key)
            return True
        if type(key) != list:
            messError = se+'\n\tRequired key       : len=2 [l1,l2]'
            messError += '\n\tFound (ValueError) : '+str(type(key).__name__)+'  '+toString(key)
            return True
        #
        for v in [[0, 'typeLINE1', 'l1'], [1, 'typeLINE1', 'l2']]:
            s1 = __check04__(v[1], key[v[0]])
            if s1:
                messError = se+'\n\tkey                : [l1,l2] = '+toString(key)
                messError += '\n\t'+v[2]+s1
                return True
        #
        if 'addOBJ' in fn:
            try:
                LINE(key[0]), LINE(key[1])
            except Exception as err:
                messError = se+'\n\twith key = '+toString(key)+'\n\t'+str(err)
                return True
    messError = ''
    return False


def __check_param_setting__(o1, ob, param, setting):
    try:
        vUDF = OLCase.__UDF__[o1.__ob__]
    except:
        vUDF = []
    #
    for k, v in param.items():
        if not (k in __OLXOBJ_PARA__[ob].keys() and __OLXOBJ_PARA__[ob][k][2] or k in vUDF):
            s1 = "\nAll parameters for add '%s' object"%ob
            for k1, v1 in __OLXOBJ_PARA__[ob].items():
                if v1[2]:
                    s1 += '\n\t'+k1.ljust(15)+':'+v1[1]
            for u1 in vUDF:
                s1 += '\n\t'+u1.ljust(15)+':(str) '+ob+' User-Defined Field'
            #
            se = '\n\tparam : {'+toString(k)+':'+toString(v)+'}'
            se += '\nKey Error '+toString(k)
            return s1, se
        #
        if k in vUDF:
            if type(v) != str:
                se = '\n\tparam              : {'+toString(k)+':'+toString(v)+'}  for %s User-Defined Field' % ob
                se += '\n\tRequired           : (str) '
                se += '\n\tFound (ValueError) : (%s) ' % type(v).__name__+toString(v)
                return '', se
        else:
            t0 = __OLXOBJ_PARA__[ob][k][2]
            se = __check00__(t0, v)
            if not (o1 is None and ob == 'RLYD' and k in {'RMTE1', 'RMTE2'}) and se:
                se = '\n\tparam              : {'+toString(k)+':'+toString(v)+'}\t'+__OLXOBJ_PARA__[ob][k][1]+se
                return '', se
    #
    if o1 != None and o1.__ob__ in {'RLYDSG', 'RLYDSP'}:
        for k, v in setting.items():
            if k not in o1.__paramEx__['SettingNameU']:
                s1 = '\nAll sSetting for %s.changeSetting(sSetting,value):' % o1.__ob__+'(DSTYPE='+toString(o1.DSTYPE)+')\n'
                s1 += toString(o1.__paramEx__['SettingName'])+'\n'
                se = '\n\tsetting : {'+toString(k)+':'+toString(v)+'}'
                se += '\nValueError : '+toString(k)
                return s1, se
    #
    if ob in __OLXOBJ_RELAY3__:
        for k, v in setting.items():
            if type(v) != str:
                se = '\n\tsetting : {'+toString(k)+':'+toString(v)+'}'
                se += '\nRequired    : (str)'
                se += '\nFound       : ('+type(v).__name__+') '+str(v)
                return '', se
    #
    if ob in {'RLYOCG', 'RLYOCP'}:
        for vr1 in [['Z2F Fwd threshold','Z2F: Fwd threshold'], ['Z0F Fwd threshold', 'Z0F: Fwd threshold']]:
            if vr1[0] in setting.keys() and vr1[1] not in setting.keys():
                setting[vr1[1]] = setting[vr1[0]]
                setting.pop(vr1[0])

        if 'POLAR' in param.keys():
            polar = param['POLAR']
            if ob == 'RLYOCG' and polar not in {0, 1, 2, 3, 255}:
                se = '\n\tparam : {'+toString('POLAR')+':'+toString(polar)+'}'
                se += '\n\tRequired POLAR     : 0,1,2,3 (OC ground relay)'
                se += '\n\tFound (ValueError) :  '+toString(polar)
                return '', se
            if ob == 'RLYOCP' and polar not in {0, 2, 255}:
                se = '\n\tparam : {'+toString('POLAR')+':'+toString(polar)+'}'
                se += '\n\tRequired POLAR     : 0,2 (OC phase relay)'
                se += '\n\tFound (ValueError) :  '+toString(polar)
                return '', se
            #
            name = __OLXOBJ_OCSETTING__[polar][0]
            nameU = [name[i].upper() for i in range(len(name))]
            for k, v in setting.items():
                if k.upper() not in nameU:
                    s1 = '\nAll sSetting for %s:' % ob+'(POLAR='+toString(polar)+')\n'
                    s1 += toString(name)+'\n'
                    se = '\n\tsetting : {'+toString(k)+':'+toString(v)+'}'
                    se += '\n\t           '+toString(k)+' : not found'
                    return s1, se
    return '', ''


def __checkClassical__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = sp.__paramInput__['fltApp']
    fltConn = sp.__paramInput__['fltConn']
    Z = sp.__paramInput__['Z']
    outage = sp.__paramInput__['outage']
    #
    global messError
    messError = '\nSPEC_FLT.Classical(obj, fltApp, fltConn, Z, outage)'
    #
    if type(obj) not in {BUS, RLYGROUP, TERMINAL}:
        messError += '\n                  obj : Classical Fault Object'
        messError += '\n\tRequired          : BUS or RLYGROUP or TERMINAL'
        messError += '\n\tFound (ValueError): '+__getStr__(obj)[1]+'  ('+__getStr__(obj)[0]+')'
        return True
    if __check_currFileIdx__(obj):
        return True
    #
    vrunOpt = ['BUS', 'CLOSE-IN', 'CLOSE-IN-EO', 'REMOTE-BUS', 'LINE-END']
    if type(fltApp) == str:
        fltApp = fltApp.upper().replace(' ', '')
        flag = fltApp in vrunOpt
        if not flag:
            try:
                percent = float(fltApp[:fltApp.index('%')])
                flag = percent >= 0 and percent <= 100
            except:
                flag = False
            if not (fltApp.endswith('%') or fltApp.endswith('%-EO')):
                flag = False
    else:
        flag = False
    #
    if not flag:
        messError += '\n                fltApp : (str) Fault application'
        vrunOpt.extend(['xx%', 'xx%-EO'])
        messError += '\n\tRequired     (str) : '+str(vrunOpt)[1:-1].replace(' ', '')+' (0<=xx<=100)'
        messError += '\n\t'+__getErrValue__(str, fltApp)
        return True
    if (type(obj) == BUS and fltApp != 'BUS') or (type(obj) != BUS and fltApp == 'BUS'):
        messError += '\n\tobj    : Classical Fault Object'
        messError += '\n\tfltApp : (str) Fault application'
        messError += '\nFound\n\tobj    : %s\n\tfltApp : %s' % (type(obj).__name__, fltApp)
        return True
    if fltApp not in vrunOpt:
        if type(obj.EQUIPMENT) != LINE:
            messError += '\n\tobj    : Classical Fault Object'
            messError += '\n\tfltApp : (str) Fault application'
            messError += "\n\tFound\n\tobj(EQUIPMENT): %s\n\tfltApp : %s" % (type(obj.EQUIPMENT).__name__, fltApp)
            return True
    # check fltConn
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        messError += '\n               fltConn : Fault connection code'
        messError += '\n\tRequired     (str) : '+str(__OLXOBJ_fltConn__[:7])[1:-1].replace(' ', '')
        messError += ',\n\t                     '+str(__OLXOBJ_fltConn__[7:])[1:-1].replace(' ', '')
        messError += '\n\t'+__getErrValue__(str, fltConn)
        return True
    # check Z
    if not (Z is None or (type(Z) == list and len(Z) == 2 and type(Z[0]) in __fi and type(Z[1]) in __fi)):
        messError += '\n\t                 Z : [R,X] Fault impedances, in Ohm'
        messError += '\n\tRequired           : [float,float]'
        messError += '\n\tFound (ValueError) : '+__getStr__([Z])[1]+'  '+__getStr__([Z])[0]
        return True
    #
    if not (outage is None or type(outage) == OUTAGE):
        messError += '\n\t            outage : Outage option'
        messError += '\n\tRequired           : None or (OUTAGE)'
        messError += '\n\t'+__getErrValue__(OUTAGE, outage)
        return True
    return False


def __checkListType__(val, typ, leng=None, valref=None):
    if type(val) != list:
        return False
    if leng != None and len(val) != leng:
        return False
    for v1 in val:
        if typ == float:
            if type(v1) not in __fi:
                return False
        else:
            if type(v1) != typ:
                return False
        if valref is not None:
            if v1 not in valref:
                return False
    return True


def __checkSimultaneous__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = sp.__paramInput__['fltApp']
    fltConn = sp.__paramInput__['fltConn']
    Z = sp.__paramInput__['Z']
    #
    global messError
    messError = '\nSPEC_FLT.Simultaneous(obj,fltApp,fltConn,Z)'
    dfltApp = ['BUS', 'CLOSE-IN', 'BUS2BUS', 'LINE-END', 'OUTAGE', '1P-OPEN', '2P-OPEN', '3P-OPEN']
    dfltApp1 = ['BUS', 'CLOSE-IN', 'BUS2BUS', 'LINE-END', 'OUTAGE', '1P-OPEN', '2P-OPEN', '3P-OPEN', 'xx%']
    flag = type(fltApp) == str
    if flag:
        fltApp1 = fltApp.upper()
        flag = fltApp1 in dfltApp
        if not flag:
            try:
                percent = float(fltApp1[:fltApp1.index('%')])
                flag = percent >= 0 and percent <= 100
            except:
                flag = False
            if not fltApp1.endswith('%'):
                flag = False
    if not flag:
        messError += '\n\t            fltApp : (str) Fault application code'
        messError += '\n\tRequired           : (str) in '+str(dfltApp1)+' (0<=xx<=100)'
        messError += '\n\t'+__getErrValue__(str, fltApp)
        return True
    #
    messError += ", with fltApp='%s'" % fltApp
    #
    if fltApp1 == 'BUS':
        if type(obj) != BUS:
            messError += '\n\tobj                : Simultaneous Fault Object'
            messError += '\n\tRequired           : (BUS)'
            messError += '\n\t'+__getErrValue__(BUS, obj)
            return True
        if __check_currFileIdx__(obj):
            return True
    #
    if fltApp1 in {'CLOSE-IN', 'LINE-END'}:
        if type(obj) not in {TERMINAL, RLYGROUP}:
            messError += '\n\t               obj : Simultaneous Fault Object'
            messError += '\n\tRequired           : TERMINAL or RLYGROUP'
            messError += '\n\t'+__getErrValue__(TERMINAL, obj)
            return True
        if __check_currFileIdx__(obj):
            return True
    #
    if fltApp1.find('%') > 0:
        if type(obj) not in {TERMINAL, RLYGROUP} or type(obj.EQUIPMENT) != LINE:
            messError += '\n\t               obj : Simultaneous Fault Object'
            messError += '\n\tRequired           : TERMINAL or RLYGROUP (on a LINE)'
            if type(obj) not in {TERMINAL, RLYGROUP}:
                messError += '\n\t'+__getErrValue__(TERMINAL, obj)
            else:
                messError += '\n\tFound (ValueError) :TERMINAL or RLYGROUP (on a '+type(obj.EQUIPMENT).__name__
            return True
        if __check_currFileIdx__(obj):
            return True
    #
    if fltApp1 in {'BUS', 'CLOSE-IN', 'LINE-END'} or fltApp1.find('%') > 0:
        if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
            messError += '\n\t           fltConn : (str) Fault connection code'
            messError += '\n\tRequired           : (str) in '+str(__OLXOBJ_fltConn__[:7])[:-1]
            messError += ',\n\t                     '+str(__OLXOBJ_fltConn__[7:])[1:]
            messError += '\n\t'+__getErrValue__(str, fltConn)
            return True
        #
        if Z != None:
            flag = type(Z) == list and len(Z) == 8
            if flag:
                for z1 in Z:
                    if type(z1) not in __fi:
                        flag = False
                        break
            if not flag:
                messError += '\n  Z : Fault impedances'
                messError += '\n\tRequired [float]*8 : [RA,XA,RB,XB,RC,XC,RG,XG]'
                messError += '\n\t'+__getErrValue__(list, Z)
                return True
    #
    if fltApp1 == 'BUS2BUS':
        if type(obj) != list or len(obj) != 2 or type(obj[0]) != BUS or type(obj[1]) != BUS:
            messError += '\n\tobj                : Simultaneous Fault Object'
            messError += '\n\tRequired           : [BUS,BUS]'
            messError += '\n\t'+__getErrValue__(list, obj)
            return True
        #
        if __check_currFileIdx__(obj[0]):
            return True
        if __check_currFileIdx__(obj[1]):
            return True
        vd = ['AA', 'BB', 'CC', 'AB', 'AC', 'BC', 'BA', 'CA', 'CB']
        if type(fltConn) != str or fltConn.upper() not in vd:
            messError += '\n\t           fltConn : (str) Fault connection code'
            messError += '\n\tRequired           : (str) in '+str(vd)
            messError += '\n\t'+__getErrValue__(str, fltConn)
            return True
        #
        if Z != None:
            if type(Z) != list or len(Z) != 2 or type(Z[0]) not in __fi or type(Z[1]) not in __fi:
                messError += '\n\t                 Z : Fault impedances'
                messError += '\n\tRequired [float]*2 : [R,X]'
                messError += '\n\t'+__getErrValue__(list, Z)
                return True
    # fltApp
    if fltApp1 in {'OUTAGE', '1P-OPEN', '2P-OPEN', '3P-OPEN'}:
        if type(obj) not in {TERMINAL, RLYGROUP}:
            messError += '\n\tobj                : Simultaneous Fault Object'
            messError += '\n\tRequired           : TERMINAL or RLYGROUP'
            messError += '\n\t'+__getErrValue__(TERMINAL, obj)
            return True
        if __check_currFileIdx__(obj):
            return True
        if fltApp1 == '1P-OPEN':
            vd = ['A', 'B', 'C']
            if type(fltConn) != str or fltConn.upper() not in vd:
                messError += '\n\t           fltConn : Fault connection code'
                messError += '\n\tRequired           : (str) in '+str(vd)
                messError += '\n\t'+__getErrValue__(str, fltConn)
                return True
        if fltApp1 == '2P-OPEN':
            vd = ['AB', 'AC', 'BC', 'BA', 'CA', 'CB']
            if type(fltConn) != str or fltConn.upper() not in vd:
                messError += '\n\t           fltConn : Fault connection code'
                messError += '\n\tRequired           : (str) in '+str(vd)
                messError += '\n\t'+__getErrValue__(str, fltConn)
                return True
    return False


def __checkSEA__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = sp.__paramInput__['fltApp']
    fltConn = sp.__paramInput__['fltConn']
    deviceOpt = sp.__paramInput__['deviceOpt']
    tiers = sp.__paramInput__['tiers']
    Z = sp.__paramInput__['Z']
    #
    global messError
    messError = '\nSPEC_FLT.SEA(obj,fltApp,fltConn,deviceOpt,tiers,Z)'
    if type(obj) not in {BUS, RLYGROUP, TERMINAL}:
        messError += '\n\t               obj : Stepped-Event Analysis Object'
        messError += '\n\tRequired           : BUS or RLYGROUP or TERMINAL'
        messError += '\n\t'+__getErrValue__(TERMINAL, obj)
        return True
    if type(obj) == BUS:
        if type(fltApp) != str or fltApp.upper() != 'BUS':
            messError += 'with obj=BUS'
            messError += '\n\t            fltApp : (str) Fault application code'
            messError += "\n\tRequired           : 'BUS'"
            messError += '\n\t'+__getErrValue__(str, fltApp)
            return True
    else:
        flag = type(fltApp) == str
        if flag:
            flag = fltApp.upper() == 'CLOSE-IN'
            if not flag:
                try:
                    percent = float(fltApp[:fltApp.index('%')])
                    flag = percent >= 0 and percent <= 100
                except:
                    flag = False
                if not fltApp.endswith('%'):
                    flag = False
        if not flag:
            messError += ' with obj=RLYGROUP/TERMINAL'
            messError += '\n\t            fltApp : (str) Fault application code'
            messError += "\n\tRequired           : 'CLOSE-IN' or 'xx%' (0<=xx<=100)"
            messError += '\n\t'+__getErrValue__(str, fltApp)
            return True
    #
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        messError += '\n\t           fltConn : Fault connection code'
        messError += '\n\tRequired           : (str) in '+str(__OLXOBJ_fltConn__[:7])
        messError += ',\n\t                          '+str(__OLXOBJ_fltConn__[7:])
        messError += '\n\t'+__getErrValue__(str, fltConn)
        return True
    #
    flag = type(deviceOpt) == list and len(deviceOpt) == 7
    if flag:
        for i in range(7):
            flag = flag and type(deviceOpt[i]) == int and (deviceOpt[i] in {0, 1})
    if not flag:
        messError += '\n\t         deviceOpt : Simulation options flags'
        messError += '\n\tRequired           : [int]*7 (1 - set; 0 - reset)'
        messError += '\n\t\tdeviceOpt[0] = 0/1 Consider OCGnd operations'
        messError += '\n\t\tdeviceOpt[1] = 0/1 Consider OCPh operations'
        messError += '\n\t\tdeviceOpt[2] = 0/1 Consider DSGnd operations'
        messError += '\n\t\tdeviceOpt[3] = 0/1 Consider DSPh operations'
        messError += '\n\t\tdeviceOpt[4] = 0/1 Consider Protection scheme operations'
        messError += '\n\t\tdeviceOpt[5] = 0/1 Consider Voltage relay operations'
        messError += '\n\t\tdeviceOpt[6] = 0/1 Consider Differential relay operations'
        messError += '\n\t'+__getErrValue__(list, deviceOpt)
        messError += '\n\t                     '+str(deviceOpt)
        return True
    if type(tiers) != int or tiers < 2:
        messError += '\n\t             tiers : Study extent. Take into account protective devices located within this number of tiers only'
        messError += '\n\tRequired      (int): >=2'
        messError += '\n\t'+__getErrValue__(int, tiers)
        return True
    #
    if type(Z) != list or len(Z) != 2 or type(Z[0]) not in __fi or type(Z[1]) not in __fi:
        messError += '\n\t                 Z : [R,X] Fault impedances in Ohm'
        messError += '\n\tRequired           : [float,float]'
        messError += '\n\t'+__getErrValue__(list, Z)
        return True
    return False


def __checkSEA_EX__(sp):
    time = sp.__paramInput__['time']
    fltConn = sp.__paramInput__['fltConn']
    Z = sp.__paramInput__['Z']
    #
    global messError
    messError = '\nSPEC_FLT.SEA_EX(time,fltConn,Z)'
    if type(time) not in __fi or time <= 0:
        messError += '\n\t          time [s] : time of Additional Event'
        messError += '\n\tRequired           : >0'
        messError += '\n\t'+__getErrValue__(float, time)
        return True
    if type(fltConn) != str or fltConn.upper() not in __OLXOBJ_fltConn__:
        messError += '\n\t           fltConn : Fault connection code'
        messError += '\n\tRequired           : (str) in '+str(__OLXOBJ_fltConn__[:7])[:-1]
        messError += ',\n\t                          '+str(__OLXOBJ_fltConn__[7:])[1:]
        messError += '\n\t'+__getErrValue__(str, fltConn)
        return True
    #
    if type(Z) != list or len(Z) != 2 or type(Z[0]) not in __fi or type(Z[1]) not in __fi:
        messError += '\n\t                 Z : [R,X] Fault impedances, in Ohm'
        messError += '\n\tRequired           : [float,float]'
        messError += '\n\t'+__getErrValue__(list, Z)
        return True
    return False


def __check_currFileIdx1__():
    global messError
    if __CURRENT_FILE_IDX__ == 0:
        messError = '\nASPEN OLR file is not yet opened'
        return True
    if __CURRENT_FILE_IDX__ < 0:
        messError = '\nASPEN OLR file is closed'
        return True
    return False


def __check_currFileIdx__(ob, op=0):
    global messError
    if __CURRENT_FILE_IDX__ == 0:
        messError = '\nASPEN OLR file is not yet opened'
        if op > 0:
            raise Exception(messError)
        return True
    if __CURRENT_FILE_IDX__ != ob.__currFileIdx__:
        messError = '\nASPEN OLR related to Object (%s) is closed' % type(ob).__name__
        if op:
            raise Exception(messError)
        return True
    if ob.__hnd__ < 0:
        messError = '\nObject has been deleted'
        if op:
            raise Exception(messError)
        return True
    return False


def __complexe2MagnitudePhaseDeg__(v):
    return abs(v), cmath.phase(v)*180/math.pi


def __computeRelayTime__(rl, current, voltage, preVoltage):
    """ Computes operating time for a fuse, recloser, an overcurrent relay (phase or ground),
           or a distance relay (phase or ground) at given currents and voltages """
    if __check_currFileIdx__(rl):
        raise Exception(messError)
    #
    opTime = c_double(0)
    opDevice = create_string_buffer(b'\000'*128)
    #
    Imag = (c_double*5)(0)
    Iang = (c_double*5)(0)
    for i in range(len(current)):
        Imag[i], Iang[i] = __complexe2MagnitudePhaseDeg__(current[i])
    #
    Vmag = (c_double*3)(0)
    Vang = (c_double*3)(0)
    for i in range(3):
        Vmag[i], Vang[i] = __complexe2MagnitudePhaseDeg__(voltage[i])
    #
    VpreMag, VpreAng = __complexe2MagnitudePhaseDeg__(preVoltage)
    #
    if OLXAPI_OK != OlxAPI.ComputeRelayTime(rl.__hnd__, Imag, Iang, Vmag, Vang, c_double(VpreMag), c_double(VpreAng), pointer(opTime), pointer(opDevice)):
        raise Exception(ErrorString())
    #
    return opTime.value, decode(opDevice.value)


def __convert2Float__(val):
    try:
        return float(val)
    except:
        pass
    if type(val) == list:
        res = []
        for v1 in val:
            try:
                res.append(float(v1))
            except:
                res.append(v1)
        return res
    return val

def __convert2Type1__(val, k, typ=float):
    if type(val)==list:
        try:
            val[k] = typ(val[k])
        except:
            pass


def __convert2Type2__(val,typ):
    if typ == 'str':
        return str(val)
    if typ == 'int':
        if type(val) in {str, int}:
            return int(val)
    if typ == 'float':
        if type(val) in {str, int, float}:
            return float(val)


def __convert2Int__(val):
    if type(val) == int:
        return val
    if type(val) == str:
        try:
            return int(val)
        except:
            pass

    if type(val) == list:
        res = []
        for v1 in val:
            if type(v1) in {int,str}:
                try:
                    res.append(int(v1))
                except:
                    res.append(v1)
            else:
                res.append(v1)
        return res
    return val


def __currentFault__(obj, style):
    hnd = HND_SC if obj is None else obj.__hnd__
    vd12Mag = (c_double*12)(0)
    vd12Ang = (c_double*12)(0)
    #
    if OLXAPI_FAILURE == OlxAPI.GetSCCurrent(hnd, vd12Mag, vd12Ang, c_int(style)):
        raise Exception(ErrorString())
    #
    if obj is None or type(obj) == TERMINAL:
        return __resultComplex__(vd12Mag[:3], vd12Ang[:3])
    if type(obj) == XFMR3:
        return __resultComplex__(vd12Mag[:12], vd12Ang[:12])
    if type(obj) in {XFMR, SHIFTER}:
        return __resultComplex__(vd12Mag[:8], vd12Ang[:8])
    if type(obj) in {LINE, DCLINE2, SERIESRC, SWITCH}:
        return __resultComplex__([vd12Mag[0], vd12Mag[1], vd12Mag[2], vd12Mag[4], vd12Mag[5], vd12Mag[6]], [vd12Ang[0], vd12Ang[1], vd12Ang[2], vd12Ang[4], vd12Ang[5], vd12Ang[6]])
    return __resultComplex__(vd12Mag[:4], vd12Ang[:4])

def __errorNotFound__(ob):
    global messError
    if not messError.endswith(': Not Found'):
        messError = messError.replace('.'+ob, '.OLCase.find'+ob)
        raise Exception(messError)

def __errorOutage__(sParam0):
    ma = [['add_outageLst', 'add branches to outage List'],
          ['build_outageLst', 'build list branches outage'],
          ['reset_outageLst', 'reset list branches outage =[]'],
          ['toString', '(str) text description/composed of OUTAGE ']]
    messError = '\nAll methods of OUTAGE:'
    for m1 in ma:
        messError += '\n\t'+(m1[0]+'()').ljust(20)+m1[1]
    messError += '\nAll attributes of OUTAGE:'
    aa = [['outageLst', 'List branches outage'],
          ['option', 'Branch outage option code : SINGLE or SINGLE-GND or DOUBLE or ALL'],
          ['G', 'Admittance Outage line grounding (mho)']]
    for m1 in aa:
        messError += '\n\t'+m1[0].ljust(20)+m1[1]
    messError += '\nValueError '+toString(sParam0)
    return messError


def __errorsParam__(o1, sParam, sParam0, sfunc):
    global messError
    ob = o1.__ob__
    messError = '\n\nAll methods for %s: \n\t' % ob
    for vi in o1.__paramEx__['allMethods']:
        messError += vi+'(), '
    messError = messError[:-1]+'\n\nAll attributes for %s of %s:' % (sfunc, ob)
    for a1 in o1.__paramEx__['allAttributes']:
        try:
            if sfunc == 'getData' or (__OLXOBJ_PARA__[ob][a1][2] != 0 and a1 != 'GUID'):
                messError += '\n\t'+a1.ljust(15)+' : '
                messError += __OLXOBJ_PARA__[ob][a1][1]
        except:
            pass
        #
        if a1 == 'HANDLE' and sfunc != 'changeData':
            messError += '(int) %s handle' % ob
    #
    if ob in OLCase.__UDF__.keys():
        for a1 in OLCase.__UDF__[ob]:
            messError += '\n\t'+a1.ljust(15)+' : '+'(str) %s User-Defined Field' % ob
    #
    if ob in {'RLYDSG', 'RLYDSP'}:
        se = '\nAll sSetting for %s.getSetting(sSetting):' % ob+'(DSTYPE='+toString(o1.DSTYPE)+')\n'
        se += toString(o1.__paramEx__['SettingName'])
        messError = se+messError+"\nAttributeError:\n\t%s Object has no attribute/method '%s'" % (ob, sParam0)
        for i in range(len(o1.__paramEx__['SettingName'])):
            if sParam == o1.__paramEx__['SettingNameU'][i]:
                messError += "\n\ntry with %s.getSetting('%s')" % (ob, o1.__paramEx__['SettingName'][i])
                break
    else:
        messError += "\nAttributeError:\n\t%s Object has no attribute/method '%s'" % (ob, sParam0)


def __errrorsFaultSpec__(spec, sParam):
    global messError
    messError = ''
    if spec.__type__ == 'Classical':
        if sParam not in {'obj', 'fltApp', 'fltConn', 'Z', 'outage'}:
            messError = '\nSPEC_FLT.Classical(obj,runOpt,fltConn,Z,fParam,outageOpt,outageLst)'
            messError += '\nAll attributes available:'
            messError += '\n\tobj       : (obj) Classical Fault Object: BUS or RLYGROUP or TERMINAL'
            messError += "\n\tfltApp    : (str) Fault application: 'Bus','Close-in','Close-in-EO','Remote-bus','Line-end','xx%','xx%-EO'"
            messError += '\n\tfltConn   : (str) Fault connection code: '+str(__OLXOBJ_fltConn__[:7])[1:-1].replace(' ', '')
            messError += ',\n                                             '+str(__OLXOBJ_fltConn__[7:])[1:-1].replace(' ', '')
            messError += '\n\tZ         : ([R,X]) Fault impedances in Ohm'
            messError += '\n\toutage    : (OUTAGE) outage option'
            messError += '\nAttributeError: '+toString(sParam)
            return True
    elif spec.__type__ == 'Simultaneous':
        if sParam not in {'obj', 'fltApp', 'fltConn', 'Z'}:
            messError = '\nSPEC_FLT.Simultaneous(obj,fltApp,fltConn,Z)'
            messError += '\n  All attributes available:'
            messError += '\n\tobj       : (obj) Simultaneous Fault Object: BUS,[BUS,BUS],RLYGROUP,TERMINAL'
            messError += "\n\tfltApp    : (str) Simulation option code: 'BUS','CLOSE-IN','BUS2BUS','LINE-END','OUTAGE','1P-OPEN','2P-OPEN','3P-OPEN','xx%'"
            messError += '\n\tfltConn   : (str) Fault connection code:'+str(__OLXOBJ_fltConn__[:7])[1:-1].replace(' ', '')
            messError += ',\n                                            '+str(__OLXOBJ_fltConn__[7:])[1:-1].replace(' ', '')
            messError += '\n\tZ         : ([]) Fault impedances, in Ohm'
            messError += '\nAttributeError: '+toString(sParam)
            return True
    elif spec.__type__ == 'SEA':
        if sParam not in {'obj', 'runOpt', 'fltConn', 'Z', 'fParam', 'addParam'}:
            messError = '\nSPEC_FLT.SEA(obj,fltConn,runOpt,tiers,Z,fParam,addParam)'
            messError += '\n  All attributes available:'
            messError += '\n\tobj       : (obj) SEA Object'
            messError += '\n\tfltConn   : (str) Fault connection code'
            messError += '\n\trunOpt    : ([]) Simulation option falg'
            messError += '\n\tZ         : ([]) Fault impedances in Ohm'
            messError += '\n\tfParam    : (float) Intermediate Fault location in %'
            messError += '\n\taddParam  : ([]) Additionnal event for Multiple User-Defined Event'
            messError += '\nAttributeError: '+toString(sParam)
            return True
    return False


def __findTerminalHnd__(b1, b2, CID, BRCODE):
    ob = __OLXOBJ_EQUIPMENTB1__[BRCODE]
    val1 = c_int(0)
    hnd1, hnd2 = b1.__hnd__, b2.__hnd__
    while OLXAPI_OK == GetBusEquipment(hnd1, c_int(TC_BRANCH), byref(val1)):
        val2 = __getDatai__(val1.value, OlxAPIConst.BR_nBus2Hnd)
        val3 = __getDatai__(val1.value, OlxAPIConst.BR_nBus3Hnd)
        if val2 == hnd2 or val3 == hnd2:
            e1 = __getDatai__(val1.value, OlxAPIConst.BR_nHandle)
            tc1 = EquipmentType(e1)
            o1 = __getOBJ__(e1, tc=tc1)
            if type(o1) == ob:
                if o1.CID == CID:
                    return val1.value
    return 0


def __findObj1LPF__(ob):
    hnd = c_int(0)
    try:
        v1 = FindObj1LPF(ob, hnd)
    except:
        v1 = OLXAPI_FAILURE
    #
    if OLXAPI_FAILURE == v1:
        ob1 = __updateSTR__(ob)
        if ob1.startswith('[OCRLY]') or ob1.startswith('[DSRLY]'):
            ob2 = ob1[:6]+'G'+ob1[6:]
            ob3 = ob1[:6]+'P'+ob1[6:]
            if OLXAPI_FAILURE == FindObj1LPF(ob2, hnd):
                if OLXAPI_FAILURE == FindObj1LPF(ob3, hnd):
                    return c_int(-1)
            return hnd
        if ob1.startswith('[DEVICE]'):
            ob2 = ob1[:7]+'DIFF'+ob1[7:]
            ob3 = ob1[:7]+'VR'+ob1[7:]
            if OLXAPI_FAILURE == FindObj1LPF(ob2, hnd):
                if OLXAPI_FAILURE == FindObj1LPF(ob3, hnd):
                    return c_int(-1)
            return hnd
        try:
            v1 = FindObj1LPF(ob1, hnd)
        except:
            return c_int(-1)
    return hnd


def __getAREAZONE__(a1, az, se):
    global messError
    se1 = se+"\n\tError format (int) '1,2,4-10' of "+az+'\n\tFound (ValueError): '+str(a1)
    messError = ''
    if type(a1) == list:
        a1.sort()
        for ai in a1:
            if type(ai) != int:
                messError = se1
        return a1
    if a1 is None or (type(a1) == str and a1.strip() == ''):
        return []
    #
    va = str(a1).split(',')
    ra = set()
    try:
        for vi in va:
            va1 = vi.split('-')
            if len(va1) == 1:
                ra.add(int(va1[0]))
            elif len(va1) > 2:
                raise Exception(se)
            else:
                va1 = [int(v1) for v1 in va1]
                va1.sort()
                for i in range(va1[0], va1[1]+1):
                    ra.add(i)
    except:
        messError = se1
    ra = list(ra)
    ra.sort()
    return ra


def __getBusEquipment__(bus, sType):
    """ Retrieves all equipment of a given type sType that is attached to bus
        - bus : BUS
        - sType : in OLXOBJ_BUS """
    res = []
    if type(sType) == list:
        for sType1 in sType:
            res.extend(__getBusEquipment__(bus, sType1))
        return res
    #
    if sType not in __OLXOBJ_BUS__:
        se = '\nString parameter available for __getBusEquipment__(bus,sType):\n'
        se += "\nNot found: '%s'" % sType
        raise Exception(se)
    #
    __check_currFileIdx__(bus, 1)
    hnd = bus.__hnd__
    val1, val2 = c_int(0), c_int(0)
    if sType == 'TERMINAL':
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
            o1 = __getOBJ__(val1.value, tc=TC_BRANCH)
            res.append(o1)
        return res
    #
    if sType == 'RLYGROUP':
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
            if OLXAPI_OK == GetData(val1.value, c_int(OlxAPIConst.BR_nRlyGrp1Hnd), byref(val2)):
                rg1 = RLYGROUP(hnd=val2.value)
                res.append(rg1)
        return res
    #
    tc = __OLXOBJ_CONST__[sType][0]
    if sType in __OLXOBJ_EQUIPMENT__:
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
            e1 = __getDatai__(val1.value, OlxAPIConst.BR_nHandle)
            tc1 = EquipmentType(e1)
            if tc1 == tc:
                o1 = __getOBJ__(e1, tc=tc)
                res.append(o1)
        return res
    else:
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(tc), byref(val1)):
            o1 = __getOBJ__(val1.value, tc=tc)
            res.append(o1)
            if sType in __OLXOBJ_BUS1__:
                break
    return res


def __getBusEquipmentHnd__(bus, sType):
    """ Retrieves all equipment hnd of a given type sType that is attached to bus
        - bus : BUS
        - sType : in __OLXOBJ_BUS__ """
    if sType not in __OLXOBJ_BUS__:
        se = '\nString parameter available for __getBusEquipmentHnd__(bus,sType):\n'
        se += "\nNot found: '%s'" % sType
        raise Exception(se)
    #
    res = []
    if type(bus) == BUS:
        hnd = bus.__hnd__
        __check_currFileIdx__(bus, 1)
    else:
        hnd = bus
    val1, val2 = c_int(0), c_int(0)
    if sType == 'TERMINAL':
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
            res.append(val1.value)
    #
    elif sType == 'RLYGROUP':
        while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
            if OLXAPI_OK == GetData(val1.value, c_int(OlxAPIConst.BR_nRlyGrp1Hnd), byref(val2)):
                res.append(val2.value)
    else:
        tc = __OLXOBJ_CONST__[sType][0]
        if sType in __OLXOBJ_EQUIPMENT__:
            while OLXAPI_OK == GetBusEquipment(hnd, c_int(TC_BRANCH), byref(val1)):
                e1 = __getDatai__(val1.value, OlxAPIConst.BR_nHandle)
                tc1 = EquipmentType(e1)
                if tc1 == tc:
                    res.append(e1)
        else:
            while OLXAPI_OK == GetBusEquipment(hnd, c_int(tc), byref(val1)):
                res.append(val1.value)
                if sType in __OLXOBJ_BUS1__:
                    break
    return res


def __getClassical__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = (sp.__paramInput__['fltApp']).upper().replace(' ', '')
    fltConn = (sp.__paramInput__['fltConn']).upper().replace(' ', '')
    Z = sp.__paramInput__['Z']
    outage = sp.__paramInput__['outage']
    #
    param1 = dict()
    if Z is None:
        Z = [0, 0]
    #
    if __check_currFileIdx__(obj):
        return None
    param1['hnd'] = obj.__hnd__
    param1['R'] = c_double(Z[0])
    param1['X'] = c_double(Z[1])
    #
    param1['fltConn'] = (c_int*4)(0)
    v1 = __OLXOBJ_fltConnCLS__[fltConn]
    param1['fltConn'][v1[0]] = v1[1]
    #
    param1['outageLst'] = (c_int*100)(0)
    param1['outageOpt'] = (c_int*4)(0)
    fltOpt = (c_double*15)(0)
    #
    flagOutage = False
    if outage is not None:
        option = outage.option.upper()
        if __check1_Outage__(outage):
            return None
        la = outage.outageLst
        for i in range(len(la)):
            if __check_currFileIdx__(la[i]):
                return None
            param1['outageLst'][i] = la[i].__hnd__
            flagOutage = True
        if flagOutage:
            if option in {'SINGLE', 'SINGLE-GND'}:
                param1['outageOpt'][0] = 1
            elif option == 'DOUBLE':
                param1['outageOpt'][1] = 1
            elif option == 'ALL':
                param1['outageOpt'][2] = 1
            else:
                raise Exception('Error value of OUTAGE.option')
        if option == 'SINGLE-GND':
            fltOpt[14] = outage.G
    #
    dfltApp = {'BUS': 0, 'CLOSE-IN': 0, 'CLOSE-IN-EO': 2, 'REMOTE-BUS': 4, 'LINE-END': 6}
    if fltApp in dfltApp.keys():
        vs = dfltApp[fltApp]
        if flagOutage:
            fltOpt[vs+1] = 1
        else:
            fltOpt[vs] = 1
    else:
        percent = float(fltApp[:fltApp.index('%')])
        if 'EO' in fltApp:
            if flagOutage:
                fltOpt[11] = percent
            else:
                fltOpt[10] = percent
        else:
            if flagOutage:
                fltOpt[9] = percent
            else:
                fltOpt[8] = percent
    param1['fltOpt'] = fltOpt
    return param1


def __getDatai__(hnd, paramCode, excep=False):
    val1 = c_int(0)
    if OLXAPI_FAILURE == GetData(hnd, c_int(paramCode), byref(val1)):
        if excep:
            raise Exception(ErrorString())
        return None
    return val1.value

def __getDatai1__(hnd, paramCode,inputCode, excep=False):
    buff = __intArray1__()
    buff[0] = c_int(inputCode)
    if OLXAPI_FAILURE == GetData(hnd, c_int(paramCode), cast(pointer(buff), c_void_p)):
        if excep:
            raise Exception(ErrorString())
        return None
    return int(buff[1])

def __getDatad__(hnd, paramCode, excep=False):
    val1 = c_double(0)
    if OLXAPI_FAILURE == GetData(hnd, c_int(paramCode), byref(val1)):
        if excep:
            raise Exception(ErrorString())
        return None
    return val1.value


def __getDatai_Array__(hnd, paramCode):
    res = []
    val1 = c_int(0)
    while OLXAPI_OK == GetData(hnd, c_int(paramCode), byref(val1)):
        res.append(val1.value)
    return res


def __getData__(hnd, paramCode):
    """ Find data of element (line/bus/...)
    Args :
        hnd     : ([] or INT) handle element
        paramCode: ([] or INT) code data (BUS_,LN_,...)
    Returns:
        data    : SCALAR or []  or [][] """
    if type(hnd) != list and type(paramCode) != list:
        vt = paramCode//100
        if vt == OlxAPIConst.VT_STRING:
            val1 = create_string_buffer(b'\000'*128)
        elif vt in [0, OlxAPIConst.VT_INTEGER]:
            val1 = c_int(0)
        elif vt == OlxAPIConst.VT_DOUBLE:
            val1 = c_double(0)
        elif vt in [OlxAPIConst.VT_ARRAYDOUBLE, OlxAPIConst.VT_ARRAYSTRING, OlxAPIConst.VT_ARRAYINT]:
            val1 = create_string_buffer(b'\000'*10*1024)
        else:
            raise Exception('Error of paramCode')
        #
        if OLXAPI_FAILURE == GetData(hnd, c_int(paramCode), byref(val1)):
            er = ErrorString()
            if er == 'GetParam failure: Invalid Device Handle':
                raise Exception(er)
            return None
        #
        return __getValue__(hnd, paramCode, val1)
    #
    res = []
    if type(hnd) == list:
        if type(paramCode) != list:
            return [__getData__(hnd1, paramCode) for hnd1 in hnd]
        else:
            for hnd1 in hnd:
                r1 = [__getData__(hnd1, p1) for p1 in paramCode]
                res.append(r1)
    else:
        for p1 in paramCode:
            res.append(__getData__(hnd, p1))
    return res


def __getEquipment__(sType, scope=None):
    """ return : OBJ in system with scope (areaNum,zoneNum,kV = [kVmin kVmax]) """
    if type(sType) == list:
        return [__getEquipment__(sType1, scope) for sType1 in sType]
    #
    try:
        tc = __OLXOBJ_CONST__[sType][0]
    except:
        se = '\nString parameter available for __getEquipment__(str):\n'
        se += str(__OLXOBJ_LIST__)+'\n'
        se += "\nNot found: '%s'" % sType
        raise Exception(se)
    #
    res = []
    hnd = c_int(0)
    while OLXAPI_OK == OlxAPI.GetEquipment(c_int(tc), byref(hnd)):
        r = __getOBJ__(hnd.value, tc=tc)
        if __isInScope__(r, scope):
            res.append(r)
    return res


def __getGUID__(GUID):
    if not GUID:
        return '_{'+str(uuid.uuid4()).upper()+'}'
    return GUID


def __getErrValue__(st, obj):
    if type(obj) == st:
        if st == str:
            return "Found (ValueError) : '%s'" % obj
        if st == list:
            return 'Found (ValueError) : '+__getSobjLst__(obj)+' len=%i' % len(obj)+' '+toString(obj)
        return 'Found (ValueError) : '+toString(obj)
    else:
        se = 'Found (ValueError) : (%s) ' % type(obj).__name__
        if type(obj) == list:
            se += 'len='+str(len(obj))+' ['
            for o1 in obj:
                se += type(o1).__name__+','
            if len(obj) > 0:
                se = se[:-1]
            se += ']'
        else:
            se += toString(obj)
        return se


def __getOBJ__(hnd, tc=None, sParam=None):
    if hnd is None:
        return None
    if tc != None:
        if tc == OlxAPIConst.TC_BUS:
            return BUS(hnd=hnd)
        elif tc == OlxAPIConst.TC_GEN:
            return GEN(hnd=hnd)
        elif tc == OlxAPIConst.TC_GENUNIT:
            return GENUNIT(hnd=hnd)
        elif tc == OlxAPIConst.TC_GENW3:
            return GENW3(hnd=hnd)
        elif tc == OlxAPIConst.TC_GENW4:
            return GENW4(hnd=hnd)
        elif tc == OlxAPIConst.TC_CCGEN:
            return CCGEN(hnd=hnd)
        elif tc == OlxAPIConst.TC_XFMR:
            return XFMR(hnd=hnd)
        elif tc == OlxAPIConst.TC_XFMR3:
            return XFMR3(hnd=hnd)
        elif tc == OlxAPIConst.TC_PS:
            return SHIFTER(hnd=hnd)
        elif tc == OlxAPIConst.TC_LINE:
            return LINE(hnd=hnd)
        elif tc == OlxAPIConst.TC_DCLINE2:
            return DCLINE2(hnd=hnd)
        elif tc == OlxAPIConst.TC_SCAP:
            return SERIESRC(hnd=hnd)
        elif tc == OlxAPIConst.TC_SWITCH:
            return SWITCH(hnd=hnd)
        elif tc == OlxAPIConst.TC_MU:
            return MULINE(hnd=hnd)
        elif tc == OlxAPIConst.TC_LOAD:
            return LOAD(hnd=hnd)
        elif tc == OlxAPIConst.TC_LOADUNIT:
            return LOADUNIT(hnd=hnd)
        elif tc == OlxAPIConst.TC_SHUNT:
            return SHUNT(hnd=hnd)
        elif tc == OlxAPIConst.TC_SHUNTUNIT:
            return SHUNTUNIT(hnd=hnd)
        elif tc == OlxAPIConst.TC_SVD:
            return SVD(hnd=hnd)
        elif tc == OlxAPIConst.TC_BREAKER:
            return BREAKER(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYGROUP:
            return RLYGROUP(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYOCG:
            return RLYOCG(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYOCP:
            return RLYOCP(hnd=hnd)
        elif tc == OlxAPIConst.TC_FUSE:
            return FUSE(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYDSG:
            return RLYDSG(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYDSP:
            return RLYDSP(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYD:
            return RLYD(hnd=hnd)
        elif tc == OlxAPIConst.TC_RLYV:
            return RLYV(hnd=hnd)
        elif tc == OlxAPIConst.TC_RECLSR:
            return RECLSR(hnd=hnd)
        elif tc == OlxAPIConst.TC_RECLSRG:
            return RECLSR(hnd=hnd-1)
        elif tc == TC_BRANCH:
            return TERMINAL(hnd=hnd)
        elif tc == OlxAPIConst.TC_SCHEME:
            return SCHEME(hnd=hnd)
        else:
            raise Exception('Error TC:'+str(tc))
    else:
        if type(sParam) != list:
            if sParam not in __OLXOBJ_oHND__ and sParam != None:
                return hnd
            #
            if type(hnd) != list:
                if hnd <= 0:
                    return None
                tc1 = EquipmentType(hnd)
                return __getOBJ__(hnd, tc=tc1)
            #
            res = []
            for hnd1 in hnd:
                if hnd1 > 0:
                    tc1 = EquipmentType(hnd1)
                    res.append(__getOBJ__(hnd1, tc=tc1))
            return res
        else:
            res = []
            for i in range(len(sParam)):
                res.append(__getOBJ__(hnd[i], sParam=sParam[i]))
            return res


def __getRLYGROUP_OBJ__(rg, sType):
    """ Retrieves all Object (RLY+BUS+SCHEME) that is attached to RLYGROUP
        - rg : RLYGROUP
        - sType : OLXOBJ_RELAY+BUS """
    __check_currFileIdx__(rg, 1)
    val1, val2 = c_int(0), c_int(0)
    hnd = rg.__hnd__
    res = []
    if sType == 'BUS':
        if OLXAPI_OK == GetData(hnd, c_int(OlxAPIConst.RG_nBranchHnd), byref(val1)):
            if OLXAPI_OK == GetData(val1.value, c_int(OlxAPIConst.BR_nBus1Hnd), byref(val2)):
                res.append(val2.value)
            else:
                raise Exception('\nBRANCH not found for :'+rg.toString())
            if OLXAPI_OK == GetData(val1.value, c_int(OlxAPIConst.BR_nBus2Hnd), byref(val2)):
                res.append(val2.value)
            if OLXAPI_OK == GetData(val1.value, c_int(OlxAPIConst.BR_nBus3Hnd), byref(val2)):
                res.append(val2.value)
        else:
            raise Exception('\nBRANCH not found for :'+rg.toString())
    elif sType == 'SCHEME':
        while OLXAPI_OK == OlxAPI.GetLogicScheme(hnd, byref(val1)):
            res.append(val1.value)
    else:
        while OLXAPI_OK == OlxAPI.GetRelay(hnd, byref(val1)):
            tc1 = EquipmentType(val1.value)
            if type(sType) in __OLXOBJ_LIST__:
                for sType1 in sType:
                    if tc1 == __OLXOBJ_CONST__[sType1][0]:
                        res.append(val1.value)
            else:
                if tc1 == __OLXOBJ_CONST__[sType][0]:
                    res.append(val1.value)
    return res


def __getRLYGROUP_RLY__(rg):
    """ Retrieves all RELAY Object that is attached to RLYGROUP
        - rg : RLYGROUP """
    val1 = c_int(0)
    __check_currFileIdx__(rg, 1)
    hnd = rg.__hnd__
    res = []
    while OLXAPI_OK == OlxAPI.GetRelay(hnd, byref(val1)):
        tc1 = EquipmentType(val1.value)
        if tc1 != OlxAPIConst.TC_RECLSRG:
            r1 = __getOBJ__(val1.value, tc1)
            res.append(r1)
    return res


def __getSettingName__(o1):
    """ [str] All setting name of Relay """
    if o1.__ob__ in {'RLYDSG', 'RLYDSP'}:
        if 'SettingName' in o1.__paramEx__.keys():
            return
        nC = __getDatai__(o1.__hnd__, __OLXOBJ_SETTING__[o1.__ob__][0])
        label = __getData__(o1.__hnd__, __OLXOBJ_SETTING__[o1.__ob__][1])
        o1.__paramEx__['SettingName'] = label[:nC]
        o1.__paramEx__['SettingNameU'] = [label[i].upper() for i in range(nC)]
        return
    #
    try:
        v0 = __OLXOBJ_OCSETTING__[o1.__paramEx__['POLAR']][0]
    except:
        v0 = []
    o1.__paramEx__['SettingName'] = v0
    o1.__paramEx__['SettingNameU'] = [v0[i].upper() for i in range(len(v0))]


def __getSettingOC__(o1):
    vid = __OLXOBJ_OCSETTING__[o1.__paramEx__['POLAR']][1]
    paracode = __OLXOBJ_SETTING__[o1.__ob__][4]
    vr = __getData__(o1.__hnd__, paracode)
    return [toString(vr[i]) for i in vid]


def __getSEA__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = sp.__paramInput__['fltApp'].upper()
    fltConn = sp.__paramInput__['fltConn'].upper()
    deviceOpt = sp.__paramInput__['deviceOpt']
    tiers = sp.__paramInput__['tiers']
    Z = sp.__paramInput__['Z']
    #
    param1 = {'hnd': obj.__hnd__}
    param1['fltOpt'] = (c_double*64)(0)
    try:
        param1['fltOpt'][1] = float(fltApp[:fltApp.index('%')])
    except:
        pass
    #
    param1['fltOpt'][0] = __OLXOBJ_fltConnSEA__[fltConn]
    param1['fltOpt'][2] = Z[0]
    param1['fltOpt'][3] = Z[1]
    #
    param1['runOpt'] = (c_int*7)(0)
    for i in range(7):
        param1['runOpt'][i] = deviceOpt[i]
    #
    param1['tiers'] = tiers
    return param1


def __getSEA_EX__(sp):
    time = sp.__paramInput__['time']
    fltConn = sp.__paramInput__['fltConn'].upper()
    Z = sp.__paramInput__['Z']
    param1 = {'time': time, 'Z': Z}
    param1['fltOpt'] = __OLXOBJ_fltConnSEA__[fltConn]
    return param1


def __getSEA_Result__(index):
    dTime = c_double(0)
    # Highest phase fault current magnitude at this step
    dCurrent = c_double(0)
    nUserEvent = c_int(0)  # flag showing is this is an user-defined event
    # 4*512 bytes buffer for event description
    szEventDesc = create_string_buffer(b'\000'*512*4)
    # 50*512 bytes buffer for fault description
    szFaultDesc = create_string_buffer(b'\000'*512*50)
    nStep = OlxAPI.GetSteppedEvent(c_int(index), byref(dTime), byref(
        dCurrent), byref(nUserEvent), szEventDesc, szFaultDesc)
    if index == 0:  # step = 0 to get total number of steps
        return nStep-1
    res = dict()
    res['time'] = dTime.value
    res['currentMax'] = dCurrent.value
    res['EventFlag'] = nUserEvent.value
    res['EventDesc'] = decode(cast(szEventDesc, c_char_p).value)
    res['FaultDesc'] = decode(cast(szFaultDesc, c_char_p).value)
    res['breaker'] = []  # please contact: support@aspeninc.com
    res['device'] = []  # please contact: support@aspeninc.com
    return res


def __getSobjLst__(obj):
    if type(obj) != list:
        s1 = type(obj).__name__
    else:
        s1 = '['
        for o1 in obj:
            s1 += type(o1).__name__+','
        if len(s1) > 1:
            s1 = s1[:-1]
        s1 += ']'
    return s1


def __getStr__(va):
    sr1, sr2 = '', ''
    if type(va) != list:
        va = [va]
    for v1 in va:
        if v1 != None:
            if type(v1) == str:
                sr1 += '"'+v1+'", '
            else:
                try:
                    sr1 += v1.toString()+', '
                except:
                    sr1 += str(v1)+", "
            #
            if type(v1) == list:
                sr2 += '['
                for vi in v1:
                    sr2 += type(vi).__name__+', '
                sr2 = sr2[:-1]+'], '
            else:
                sr2 += type(v1).__name__+', '
    return sr1[:-2], sr2[:-2]


def __getSimultaneous__(sp):
    obj = sp.__paramInput__['obj']
    fltApp = sp.__paramInput__['fltApp'].upper()
    fltConn = sp.__paramInput__['fltConn'].upper()
    Z = sp.__paramInput__['Z']
    #
    Z1 = [0]*8
    try:
        for i in range(len(Z)):
            Z1[i] = Z[i]
    except:
        pass
    #
    vBRTYPE = {LINE: 1, SERIESRC: 1, XFMR: 2, SHIFTER: 3, SWITCH: 7, XFMR3: 10}
    #
    param1 = dict()
    param1['FRA'] = str(Z1[0])
    param1['FXA'] = str(Z1[1])
    param1['FRB'] = str(Z1[2])
    param1['FXB'] = str(Z1[3])
    param1['FRC'] = str(Z1[4])
    param1['FXC'] = str(Z1[5])
    param1['FRG'] = str(Z1[6])
    param1['FXG'] = str(Z1[7])
    param1['FPARAM'] = '0'
    param1['FLTCONN'] = '0'
    param1['PHASETYPE'] = '0'
    #
    dfltApp = {'BUS': 0, 'CLOSE-IN': 1, 'BUS2BUS': 2, 'LINE-END': 3, 'OUTAGE': 5, '1P-OPEN': 6, '2P-OPEN': 7, '3P-OPEN': 8}
    if fltApp in dfltApp.keys():
        param1['FLTAPPL'] = str(dfltApp[fltApp])
    else:
        param1['FLTAPPL'] = '4'
        param1['FPARAM'] = fltApp[:fltApp.index('%')]
    #
    if fltApp == '1P-OPEN':
        dfltConn = {'A': 0, 'B': 1, 'C': 2}
        param1['PHASETYPE'] = str(dfltConn[fltConn])
    elif fltApp == '2P-OPEN':
        dfltConn = {'AB': 0, 'AB': 0, 'BC': 1, 'CB': 1, 'AC': 2, 'CA': 2}
        param1['PHASETYPE'] = str(dfltConn[fltConn])
    elif fltApp in {'BUS', 'CLOSE-IN', 'LINE-END'}:
        v1 = __OLXOBJ_fltConnSIM__[fltConn]
        param1['FLTCONN'] = str(v1[0])
        param1['PHASETYPE'] = str(v1[1])
    elif fltApp == 'BUS2BUS':
        dfltConn = {'AA': [0, 0], 'BB': [0, 3], 'CC': [0, 5], 'AB': [0, 1], 'AC': [0, 2], 'BC': [0, 4], 'BA': [0, 1], 'CA': [0, 2], 'CB': [0, 4]}
        param1['PHASETYPE'] = str(dfltConn[fltConn][1])
    #
    if fltApp == 'BUS':
        param1['BUS1NAME'] = obj.NAME
        param1['BUS1KV'] = str(obj.KV)
    elif fltApp in {'OUTAGE', '1P-OPEN', '2P-OPEN', '3P-OPEN'}:
        if type(obj) in {TERMINAL, RLYGROUP}:
            obj = obj.EQUIPMENT
        bus1, bus2 = obj.BUS1, obj.BUS2
        param1['BUS1NAME'] = bus1.NAME
        param1['BUS1KV'] = str(bus1.KV)
        param1['BUS2NAME'] = bus2.NAME
        param1['BUS2KV'] = str(bus2.KV)
        param1['CKTID'] = str(obj.CID)
        param1['BRTYPE'] = str(vBRTYPE[type(obj)])
    elif fltApp == 'BUS2BUS':
        if fltConn in {'BA', 'CA', 'CB'}:
            b1, b2 = obj[1], obj[0]
        else:
            b1, b2 = obj[0], obj[1]
        param1['BUS1NAME'] = b1.NAME
        param1['BUS1KV'] = str(b1.KV)
        param1['BUS2NAME'] = b2.NAME
        param1['BUS2KV'] = str(b2.KV)
    elif fltApp in {'CLOSE-IN', 'LINE-END'} or '%' in fltApp:
        ba = obj.BUS
        e1 = obj.EQUIPMENT
        param1['BUS1NAME'] = ba[0].NAME
        param1['BUS1KV'] = str(ba[0].KV)
        param1['BUS2NAME'] = ba[1].NAME
        param1['BUS2KV'] = str(ba[1].KV)
        param1['CKTID'] = str(e1.CID)
        param1['BRTYPE'] = str(vBRTYPE[type(e1)])
    return param1


def __getTERMINAL_OBJ__(ob, sType):
    __check_currFileIdx__(ob, 1)
    hnd = ob.__hnd__
    if sType == 'BUS':
        res = []
        h1 = __getDatai__(hnd, OlxAPIConst.BR_nBus1Hnd)
        res.append(BUS(hnd=h1))
        #
        h2 = __getDatai__(hnd, OlxAPIConst.BR_nBus2Hnd)
        res.append(BUS(hnd=h2))
        h3 = __getDatai__(hnd, OlxAPIConst.BR_nBus3Hnd)
        if h3 is not None:
            res.append(BUS(hnd=h3))
        return res
    if sType == 'EQUIPMENT':
        e1 = __getDatai__(hnd, OlxAPIConst.BR_nHandle)
        tc1 = EquipmentType(e1)
        return __getOBJ__(e1, tc1)
    if sType == 'RLYGROUP':
        res = []
        for c1 in [OlxAPIConst.BR_nRlyGrp1Hnd, OlxAPIConst.BR_nRlyGrp2Hnd, OlxAPIConst.BR_nRlyGrp3Hnd]:
            r1 = __getDatai__(hnd, c1)
            try:
                res.append(RLYGROUP(hnd=r1))
            except:
                res.append(None)
        return res
    if sType in {'RLYGROUP1', 'RLYGROUP2', 'RLYGROUP3'}:
        try:
            vd = {'RLYGROUP1': OlxAPIConst.BR_nRlyGrp1Hnd,'RLYGROUP2': OlxAPIConst.BR_nRlyGrp2Hnd, 'RLYGROUP3': OlxAPIConst.BR_nRlyGrp3Hnd}
            r1 = __getDatai__(hnd, vd[sType])
            return RLYGROUP(hnd=r1)
        except:
            return None
    if sType == 'FLAG':
        return __getDatai__(hnd, OlxAPIConst.BR_nInService)
    if sType == 'OPPOSITE':
        res = []
        for t1 in ob.EQUIPMENT.TERMINAL:
            if t1.__hnd__ != hnd:
                res.append(t1)
        return res
    if sType == 'REMOTE':
        import OlxAPILib
        ba, _, _, _ = OlxAPILib.getRemoteTerminals(hnd, [])
        return [TERMINAL(hnd=b1) for b1 in ba]


def __getValue_i__(buf, count, stop=False):
    array = []
    val = cast(buf, POINTER(c_int*count)).contents
    for ii in range(count):
        if stop and val[ii] == 0:
            break
        array.append(val[ii])
    return array


def __getValue__(hnd, tokenV, buf):
    """ Convert GetData binary data buffer into Python Object of correct type """
    vt = tokenV//100
    if vt == OlxAPIConst.VT_STRING:
        return decode(buf.value)
    elif vt in [OlxAPIConst.VT_DOUBLE, OlxAPIConst.VT_INTEGER]:
        return buf.value
    #
    tc = EquipmentType(hnd)
    if tc == OlxAPIConst.TC_BREAKER and tokenV in {OlxAPIConst.BK_vnG1DevHnd, OlxAPIConst.BK_vnG2DevHnd, OlxAPIConst.BK_vnG1OutageHnd, OlxAPIConst.BK_vnG2OutageHnd}:
        count = OlxAPIConst.MXSBKF
        return __getValue_i__(buf, count, True)
    if tc == OlxAPIConst.TC_RLYGROUP and tokenV in {OlxAPIConst.RG_vnPrimaryGroup, OlxAPIConst.RG_vnBackupGroup}:
        count = OlxAPIConst.MXSBKF
        return __getValue_i__(buf, count, True)
    #
    if tc == OlxAPIConst.TC_SVD and tokenV == OlxAPIConst.SV_vnNoStep:
        count = 8
        return __getValue_i__(buf, count)
    #
    if tc in {OlxAPIConst.TC_RLYDSP, OlxAPIConst.TC_RLYDSG} and tokenV in {OlxAPIConst.DP_vParamLabels, OlxAPIConst.DG_vParamLabels, OlxAPIConst.DG_vParams, OlxAPIConst.DP_vParams}:
        return decode(cast(buf, c_char_p).value).split("\t")
    if tc == OlxAPIConst.TC_SCHEME:
        if tokenV == OlxAPIConst.LS_vsSignalName:
            va = decode(cast(buf, c_char_p).value).split("\t")
            res = []
            for v1 in va:
                if v1:
                    res.append(v1)
            return res
        if tokenV == OlxAPIConst.LS_vnSignalType:
            count = 8
            return __getValue_i__(buf, count)
        if tokenV == OlxAPIConst.LS_vdSignalVar:
            count = 8
    elif tc == OlxAPIConst.TC_DCLINE2 and tokenV == OlxAPIConst.DC_vnBridges:
        count = 2
        return __getValue_i__(buf, count)
    #
    elif tc == OlxAPIConst.TC_GENUNIT and tokenV in {OlxAPIConst.GU_vdR, OlxAPIConst.GU_vdX}:
        count = 5
    elif tc == OlxAPIConst.TC_LOADUNIT and tokenV in {OlxAPIConst.LU_vdMW, OlxAPIConst.LU_vdMVAR}:
        count = 3
    elif tc == OlxAPIConst.TC_SVD and tokenV in {OlxAPIConst.SV_vdBinc, OlxAPIConst.SV_vdB0inc}:
        count = 8
    elif tc == OlxAPIConst.TC_LINE and tokenV == OlxAPIConst.LN_vdRating:
        count = 4
    elif tc == OlxAPIConst.TC_RLYGROUP and tokenV == OlxAPIConst.RG_vdRecloseInt:
        count = 4
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV == OlxAPIConst.OG_vdDirSetting:
        count = 8
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV == OlxAPIConst.OG_vdDirSettingV15:
        count = 9
    elif tc == OlxAPIConst.TC_RLYOCG and tokenV in {OlxAPIConst.OG_vdDTPickup, OlxAPIConst.OG_vdDTDelay}:
        count = 5
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV == OlxAPIConst.OP_vdDirSetting:
        count = 8
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV == OlxAPIConst.OP_vdDirSettingV15:
        count = 9
    elif tc == OlxAPIConst.TC_RLYOCP and tokenV in {OlxAPIConst.OP_vdDTPickup, OlxAPIConst.OP_vdDTDelay}:
        count = 5
    elif tc == OlxAPIConst.TC_RLYDSG and tokenV == OlxAPIConst.DG_vdParams:
        count = OlxAPIConst.MXDSPARAMS
    elif tc == OlxAPIConst.TC_RLYDSG and tokenV in {OlxAPIConst.DG_vdDelay, OlxAPIConst.DG_vdReach, OlxAPIConst.DG_vdReach1}:
        count = OlxAPIConst.MXZONE
    elif tc == OlxAPIConst.TC_RLYDSP and tokenV == OlxAPIConst.DP_vParams:
        count = OlxAPIConst.MXDSPARAMS
    elif tc == OlxAPIConst.TC_RLYDSP and tokenV in {OlxAPIConst.DP_vdDelay, OlxAPIConst.DP_vdReach, OlxAPIConst.DP_vdReach1}:
        count = OlxAPIConst.MXZONE
    elif tc == OlxAPIConst.TC_CCGEN and tokenV in {OlxAPIConst.CC_vdV, OlxAPIConst.CC_vdI, OlxAPIConst.CC_vdAng}:
        count = OlxAPIConst.MAXCCV
    elif tc == OlxAPIConst.TC_BREAKER and tokenV in {OlxAPIConst.BK_vdRecloseInt1, OlxAPIConst.BK_vdRecloseInt2}:
        count = 3
    elif tc == OlxAPIConst.TC_MU and tokenV in [OlxAPIConst.MU_vdX, OlxAPIConst.MU_vdR, OlxAPIConst.MU_vdFrom1, OlxAPIConst.MU_vdFrom2, OlxAPIConst.MU_vdTo1, OlxAPIConst.MU_vdTo2]:
        count = 5
    # and tokenV in {OlxAPIConst.DC_vdAngleMax,OlxAPIConst.DC_vdAngleMin}:
    elif tc == OlxAPIConst.TC_DCLINE2:
        count = 2
    else:
        count = OlxAPIConst.MXDSPARAMS
    val = cast(buf, POINTER(c_double*count)).contents
    return [v for v in val]

#
def __comp__(b00,b01, i1,i2,i3,i4, j1,j2,j3,j4):
    return  b00[i1].__hnd__ == b01[j1].__hnd__ and b00[i2].__hnd__ == b01[j2].__hnd__\
        and b00[i3].__hnd__ == b01[j3].__hnd__ and b00[i4].__hnd__ == b01[j4].__hnd__

def __get_OBJTERMINAL__(ob):  # return list TERMINAL of Object
    res = []
    for b1 in ob.BUS:
        for t1 in b1.TERMINAL:
            if t1.EQUIPMENT.__hnd__ == ob.__hnd__:
                res.append(t1)
    return res


def __initOBJ__(sg, key, hnd):
    global messError
    if __check_currFileIdx1__():
        return -1
    messError = ''
    #
    ob = sg.__name__
    if key != None and hnd != None:
        messError = '\nOlxObj.'+ob+'()takes 1 argument but 2 were given'
        return -1
    #
    if key != None:
        if __check_ob_key__('\nOlxObj.', ob, key):
            return -1
        #
        t1 = type(key)
        if t1 == sg:
            if __check_currFileIdx__(key):
                return -1
            return key.__hnd__
        elif t1 == str:
            hnd = __findObj1LPF__(key)
            if hnd.value<0:
                messError = '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
                return -1
            tc1 = EquipmentType(hnd)
            if tc1 == OlxAPIConst.TC_BUS and ob in __OLXOBJ_BUS1__:
                try:
                    b1 = BUS(hnd=hnd.value)
                    hnd = b1.getData(ob).__hnd__
                except:
                    messError = '\nOlxObj.'+ob + '('+toString(key)+') : Not Found'
                    return -1
                tc1 = EquipmentType(hnd)
            else:
                hnd = hnd.value
            tc = __OLXOBJ_CONST__[ob][0]
            if tc == OlxAPIConst.TC_RECLSR and tc1 == OlxAPIConst.TC_RECLSRG:
                return hnd-1
            if tc != tc1:
                messError = '\nOlxObj.'+ob+'('+toString(key)+')'
                messError += '\n\tSearch: '+ob
                messError += '\n\tFound : '+__OLXOBJ_CONST1__[tc1][0]
                return -1
            return hnd
        #
        if ob == 'BUS' or ob in __OLXOBJ_BUS1__:
            if t1 == int:
                hnd = OlxAPI.FindBusNo(key)
            elif t1 == BUS:
                hnd = key.__hnd__
            else:
                hnd = OlxAPI.FindBus(key[0], key[1])
            if sg != BUS and hnd > 0:
                try:
                    hnd = BUS(hnd=hnd).getData(ob).__hnd__
                except:
                    messError = '\nOlxObj.'+ob + '('+toString(key)+') : Not Found'
                    hnd = -1
            if hnd <= 0:
                messError = '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
                return -1
            return hnd
        #
        if ob in __OLXOBJ_BUS2__:
            if type(key) in {LOADUNIT, SHUNTUNIT, GENUNIT} and type(key).__name__ == ob:
                hnd = key.__hnd__
            elif ob == 'GENUNIT':
                if type(key[0]) == GEN:
                    g = GEN(hnd=key[0].__hnd__)
                else:
                    try:
                        b = BUS(key[0]) if len(key) == 2 else BUS(key[:2])
                        g = b.GEN
                    except:
                        g = None
                if g != None:
                    for g1 in g.GENUNIT:
                        if g1.CID == key[-1]:
                            hnd = g1.__hnd__
            elif ob == 'SHUNTUNIT':
                if type(key[0]) == SHUNT:
                    g = SHUNT(hnd=key[0].__hnd__)
                else:
                    try:
                        b = BUS(key[0]) if len(key) == 2 else BUS(key[:2])
                        g = b.SHUNT
                    except:
                        g = None
                if g != None:
                    for g1 in g.SHUNTUNIT:
                        if g1.CID == key[-1]:
                            hnd = g1.__hnd__
            elif ob == 'LOADUNIT':
                if type(key[0]) == LOAD:
                    g = LOAD(hnd=key[0].__hnd__)
                else:
                    try:
                        b = BUS(key[0]) if len(key) == 2 else BUS(key[:2])
                        g = b.LOAD
                    except:
                        g = None
                if g != None:
                    for g1 in g.LOADUNIT:
                        if g1.CID == key[-1]:
                            hnd = g1.__hnd__
        #
        elif ob in __OLXOBJ_EQUIPMENT__:
            b3 = None
            try:
                b1, b2 = BUS(key[0]), BUS(key[1])
                if len(key) == 4:
                    b3 = BUS(key[2])
            except Exception as err:
                messError = '\nOlxObj.'+ob+'(key)\n\tkey= '+toString(key)+str(err)
                return -1
            #
            for t1 in b1.TERMINAL:
                if b2.isInList(t1.BUS[1:]):
                    if b3 is None or b3.isInList(t1.BUS[1:]):
                        e1 = t1.EQUIPMENT
                        if type(e1) == sg and e1.CID == key[-1]:
                            return e1.__hnd__
            messError = '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
            return -1
        #
        elif ob == 'TERMINAL':
            try:
                b1, b2 = BUS(key[0]), BUS(key[1])
                hnd = __findTerminalHnd__(b1, b2, key[2], key[3].upper())
            except Exception as err:
                messError = '\nOlxObj.'+ob +'('+toString(key)+') : Not Found'+str(err)
                return -1
            if hnd == 0:
                messError = '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
                return -1
            return hnd
        #
        elif ob == 'RLYGROUP':
            try:
                b1, b2 = BUS(key[0]), BUS(key[1])
                hnd = __findTerminalHnd__(b1, b2, key[2], key[3])
            except Exception as err:
                messError = '\nOlxObj.'+ob + '('+toString(key)+') : Not Found\n'+str(err)
                return -1
            if hnd > 0:
                try:
                    hnd = TERMINAL(hnd=hnd).RLYGROUP1.__hnd__
                except:
                    hnd = 0
            if hnd <= 0:
                messError = '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
                return -1
        #
        elif ob in __OLXOBJ_RELAY__:
            try:
                b1, b2 = BUS(key[0]), BUS(key[1])
                rg = RLYGROUP([b1, b2, key[2], key[3]])
                ra = rg.getData(ob)
                for r1 in ra:
                    if r1.ID == key[-1]:
                        return r1.__hnd__
                #
                if ob == 'RECLSR' and key[-1].endswith('_P'):
                    for r1 in ra:
                        if r1.ID == key[-1][:-2]:
                            return r1.__hnd__
            except Exception as err:
                messError = str(err)+'\nOlxObj.'+ob + '('+toString(key)+') : Not Found'
                return -1
        #
        elif ob == 'MULINE':
            try:
                l1, l2 = LINE(key[0]), LINE(key[1])
            except Exception as err:
                messError = str(err)+'\nOlxObj.'+ob + '('+toString(key)+') : Not Found'
                return -1
            #
            for m1 in l1.MULINE:
                li1, li2 = m1.LINE1, m1.LINE2
                if (l1.__hnd__ == li1.__hnd__ and l2.__hnd__ == li2.__hnd__) or (l2.__hnd__ == li1.__hnd__ and l1.__hnd__ == li2.__hnd__):
                    return m1.__hnd__
        #
        elif ob == 'BREAKER':
            try:
                b1 = BUS(key[0]) if len(key) == 2 else BUS(key[:-1])
                for bk1 in b1.BREAKER:
                    if bk1.NAME == key[-1]:
                        return bk1.__hnd__
            except Exception as err:
                messError = str(err)
                hnd = None
    if hnd is None:
        messError += '\nOlxObj.'+ob+'('+toString(key)+') : Not Found'
        hnd = -1
    return hnd


def __int2bin__(n):
    return __int2bin__(n >> 1)+[n & 1] if n > 1 else [1]


def __int2bin4__(n):
    va = [0] if n == 0 else __int2bin__(n)
    va.reverse()
    le = len(va)
    for i in range(le, 4):
        va.append(0)
    return va


def __isInScope__(ob, scope):
    """ check if Object is in scope """
    if scope is None or scope['isFullNetWork']:
        return True
    #
    typ = type(ob)
    if typ == BUS:
        return __busIsInScope__(ob, scope['areaNum'], scope['zoneNum'], scope['kV'])
    elif typ == XFMR:
        b1, b2 = ob.BUS1, ob.BUS2
        kv1, kv2 = b1.KV, b2.KV
        if kv1 == kv2:
            return __brIsInScope__([b1, b2], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], scope['kV']])
        if kv1 > kv2:
            return __brIsInScope__([b1, b2], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], []])
        return __brIsInScope__([b1, b2], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [[], scope['kV']])
    elif typ == XFMR3:
        b1, b2, b3 = ob.BUS1, ob.BUS2, ob.BUS3
        kv1, kv2, kv3 = b1.KV, b2.KV, b3.KV
        if kv1 >= kv2 and kv1 >= kv3:
            return __brIsInScope__([b1, b2, b3], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], [], []])
        elif kv2 >= kv1 and kv2 >= kv3:
            return __brIsInScope__([b1, b2, b3], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [[], scope['kV'], []])
        return __brIsInScope__([b1, b2, b3], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [[], [], scope['kV']])
    elif typ == MULINE:
        return __brIsInScope__([ob.LINE1.BUS1, ob.LINE1.BUS2], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], scope['kV']]) or \
               __brIsInScope__([ob.LINE2.BUS1, ob.LINE2.BUS2], scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], scope['kV']])
    elif typ in __OLXOBJ_EQUIPMENTO__:
        return __brIsInScope__(ob.BUS, scope['areaNum'], scope['zoneNum'], scope['optionTie'], [scope['kV'], scope['kV']])
    else:
        ba = ob.BUS
        if type(ba) == list:
            return __busIsInScope__(ob.BUS[0], scope['areaNum'], scope['zoneNum'], scope['kV'])
        return __busIsInScope__(ob.BUS, scope['areaNum'], scope['zoneNum'], scope['kV'])


def __mesParamSys__():
    mes = 'All System Parameters:'
    for k, v in __OLXOBJ_PARASYS__.items():
        mes += '\n'+k.ljust(20)+": "
        if v:
            if type(v[0]) == str:
                mes += (v[0].ljust(8)+v[1]).ljust(40)+v[2]
            else:
                mes += (str(v[0])[1:-1].replace(' ','').ljust(8)+v[1]).ljust(40)+v[2]
    return mes


def __pickFault__(sf):
    global messError, __INDEX_FAULT__
    messError = ''
    if __checkFault__(sf):
        return messError
    #if __INDEX_FAULT__ != sf.__index__:
    if OLXAPI_FAILURE == OlxAPI.PickFault(c_int(sf.__index__), c_int(sf.__tiers__)):
        messError = '\nError PickFault index=%i, with index available: 1-' % sf.__index__+str(__COUNT_FAULT__)
    __INDEX_FAULT__ = sf.__index__
    return messError


def __preVoltage__(obj, style):  # style 1:kV; 2:PU
    vdOut1 = (c_double*3)(0)
    vdOut2 = (c_double*3)(0)
    if OLXAPI_FAILURE == OlxAPI.GetPSCVoltage(obj.__hnd__, vdOut1, vdOut2, c_int(style)):
        raise Exception(ErrorString())
    #
    if type(obj) in {XFMR, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
        return __resultComplex__(vdOut1[:2], vdOut2[:2], 2)
    if type(obj) == XFMR3:
        return __resultComplex__(vdOut1, vdOut2, 2)
    return __resultComplex__(vdOut1[:1], vdOut2[:1], 2)


def __removeCOORDPAIR__(o1, val, r1):
    try:
        rg1 = RLYGROUP(r1)
    except Exception as err:
        if __OLXOBJ_VERBOSE__:
            print(str(err))
        rg1 = None
    #
    if rg1 != None:
        ra = []
        ra0 = o1.getData(val)
        for ri in ra0:
            if ri.__hnd__ != rg1.__hnd__:
                ra.append(ri)
        if len(ra) != len(ra0):
            o1.changeData(val, ra)
            o1.postData()
        #
        val2 = 'BACKUP' if val == 'PRIMARY' else 'PRIMARY'
        ra = []
        ra0 = rg1.getData(val2)
        for ri in ra0:
            if ri.__hnd__ != o1.__hnd__:
                ra.append(ri)
        if len(ra) != len(ra0):
            rg1.changeData(val2, ra)
            rg1.postData()


def __resultComplex__(v1, v2, style=1):# style=1: real, imag   style=2: magnitude, angle in degree
    res = []
    if style == 1:
        return [complex(v1[i], v2[i]) for i in range(len(v1))]
    if style == 2:
        res = []
        for i in range(len(v1)):
            angle = v2[i] * math.pi/180
            res.append(complex(v1[i]*math.cos(angle), v1[i]*math.sin(angle)))
        return res
    raise Exception('__resultComplex__ error style')


def __runSimulate__(specFlt, clearPrev):
    global messError
    messError = '\nOLCase.SimulateFault(specFlt,clearPrev)'
    if type(clearPrev) != int or clearPrev not in {0, 1}:
        messError += '\n\t        clearPrev : clear previous result flag'
        messError += '\n\tRequired          : (int) 0 or 1 '
        if type(clearPrev) == int:
            messError += '\n\tFound (ValueError): %i' % clearPrev
        else:
            messError += '\n\tFound (ValueError): (%s)' % type(clearPrev).__name__+' '+str(clearPrev)
        return True
    #
    if type(specFlt) != SPEC_FLT or specFlt.__type__ not in {'Classical', 'Simultaneous', 'SEA'}:
        flag = type(specFlt) == list and len(specFlt) > 0
        if flag:
            flag1 = True
            for s1 in specFlt:
                if type(s1) != SPEC_FLT or s1.__type__ != 'Simultaneous':
                    flag1 = False
            flag2 = type(specFlt[0]) == SPEC_FLT and specFlt[0].__type__ == 'SEA'
            for i in range(1, len(specFlt)):
                if type(specFlt[i]) != SPEC_FLT or specFlt[i].__type__ != 'SEA_EX':
                    flag2 = False
            flag = flag1 or flag2
        #
        if not flag:
            messError += '\n\t           specFlt : Fault Specfication'
            messError += '\n\tRequired           : '
            messError += "\n\t\t   SPEC_FLT ('Classical','Simultaneous','SEA')"
            messError += "\n\t\tor [SPEC_FLT] ('Simultaneous')"
            messError += "\n\t\tor [sp_0,sp_i,...] sp_0=SPEC_FLT('SEA') sp_i=SPEC_FLT('SEA_EX')"
            messError += '\n\tFound (ValueError) : '
            if type(specFlt) != list:
                messError += type(specFlt).__name__
            else:
                messError += '['
                for sp1 in specFlt:
                    if type(sp1) == SPEC_FLT:
                        messError += "SPEC_FLT('%s')" % sp1.__type__+','
                    else:
                        messError += type(sp1).__name__+','
                messError = messError[:-1]+']'
            return True
    #
    global __INDEX_SIMUL__, __TYPEF_SIMUL__, FltSimResult
    #
    if clearPrev == 1 or (type(specFlt) == SPEC_FLT and specFlt.__type__ == 'SEA') or (type(specFlt) == list and specFlt[0].__type__ == 'SEA'):
        __INDEX_SIMUL__ = abs(__INDEX_SIMUL__)+1
        FltSimResult.clear()
    #
    if type(specFlt) == SPEC_FLT:
        specFlt = [specFlt]
    # check data
    for sp1 in specFlt:
        if sp1.checkData():
            return True
    # run Classical
    if specFlt[0].__type__ == 'Classical':
        __TYPEF_SIMUL__ = 'Classical'
        param = specFlt[0].getData()
        if param is None:
            return True
        if OLXAPI_FAILURE == OlxAPI.DoFault(param['hnd'], param['fltConn'], param['fltOpt'], param['outageOpt'], param['outageLst'], param['R'], param['X'], c_int(clearPrev)):
            raise Exception(ErrorString())
        if __OLXOBJ_VERBOSE__:
            print("OLCase.simulateFault('Classical') : Operation executed successfully.")
    # run SEA
    elif specFlt[0].__type__ == 'SEA':
        __TYPEF_SIMUL__ = 'SEA'
        param = specFlt[0].getData()
        sn = 'SEA-Single'
        for i in range(1, len(specFlt)):
            sn = 'SEA-Multiple'
            param1 = specFlt[i].getData()
            k = 4*i
            param['fltOpt'][k] = param1['fltOpt']
            param['fltOpt'][k+1] = param1['time']
            param['fltOpt'][k+2] = param1['Z'][0]
            param['fltOpt'][k+3] = param1['Z'][1]
        #
        if OLXAPI_FAILURE == OlxAPI.DoSteppedEvent(param['hnd'], param['fltOpt'], param['runOpt'], param['tiers']):
            raise Exception(ErrorString())
        if __OLXOBJ_VERBOSE__:
            print("OLCase.simulateFault('%s') : Operation executed successfully." % sn)
    else:
        __TYPEF_SIMUL__ = 'Simultaneous'
        data = ET.Element('SIMULATEFAULT')
        data.set('CLEARPREV', str(clearPrev))
        data1 = ET.SubElement(data, 'FAULT')
        for i in range(len(specFlt)):
            sp1 = specFlt[i]
            param1 = sp1.getData()
            data2 = ET.SubElement(data1, 'FLTSPEC')
            data2.set('FLTDESC', 'specFlt: '+str(i+1))
            for k, v in param1.items():
                data2.set(k, v)
        #
        sInput = decode(ET.tostring(data))
        #
        if OLXAPI_FAILURE == OlxAPI.Run1LPFCommand(sInput):
            raise Exception(ErrorString())
        if __OLXOBJ_VERBOSE__:
            print(
                "OLCase.simulateFault('Simultaneous') : Operation executed successfully.")
    if __OLXOBJ_VERBOSE__:
        for sp1 in specFlt:
            print('\t'+sp1.toString())
    __runSimulateFin__(len(FltSimResult))
    return False


def __runSimulateFin__(n0):
    global __COUNT_FAULT__, FltSimResult
    __COUNT_FAULT__ = __getDatai__(HND_SC, OlxAPIConst.FT_nNOfaults, True)
    for i in range(n0, __COUNT_FAULT__):
        FltSimResult.append(RESULT_FLT(i+1))  # ini result fault


def __scheme_getLogic__(o1, nameVar):
    global messError
    if __check_currFileIdx__(o1):
        return
    #
    if type(nameVar)==str:
        nameVar = nameVar.upper()
    #
    if nameVar == 'EQUATION':
        return o1.getData('EQUATION')
    #
    namea, typa, vala = __scheme_get0__(o1)
    #
    if nameVar == None:
        res = {'EQUATION': o1.getData('EQUATION')}
        for nv1 in namea:
            res[nv1] = __scheme_get1__(nv1,typa,vala)
        return res

    if nameVar in namea:
        return __scheme_get1__(nameVar,typa,vala)

    nv00 = ''
    if type(nameVar) in __OLXOBJ_LISTT__:
        res = dict()
        for nv1 in nameVar:
            if type(nv1) == str:
                nv1u = nv1.upper()
                if nv1u in namea:
                    res[nv1u] = __scheme_get1__(nv1u,typa,vala)
                else:
                    nv00 = nv1
            else:
                nv00 = nv1
        if nv00 == '':
            return res
    else:
        nv00 = nameVar
    messError = '\nSCHEME.getLogic(%s)'%toString(nameVar)
    messError += '\n\tnameVar : str or None or list(tuple,set) in '+toString(namea)+" or 'EQUATION'"

def __scheme_get0__(o1):
    namea = __getData__(o1.__hnd__, OlxAPIConst.LS_vsSignalName)
    typa = dict(zip(namea, __getData__(o1.__hnd__, OlxAPIConst.LS_vnSignalType)))
    vala = dict(zip(namea, __getData__(o1.__hnd__, OlxAPIConst.LS_vdSignalVar)))
    return namea, typa, vala

def __scheme_get1__(nv1,typa,vala):
    t1 = int(typa[nv1])
    v1 = vala[nv1]
    if t1 == OlxAPIConst.LVT_CONST:
        return v1

    v1 = __getOBJ__( int(v1) )
    for k, v in __OLXOBJ_SCHEME__.items():
        if v == t1:
            return [k, v1]
    return [str(t1), v1]

def __scheme_changeLogic__(o1, nameVar, value):
    if __check_currFileIdx__(o1):
        return
    global messError
    messError = ''

    if type(nameVar) != str:
        messError = "\nSCHEME.changeLogic(nameVar=%s,value) " % toString(nameVar)
        messError += "\n\tRequired nameVar : str"
        messError += "\n\tFound            : "+type(nameVar).__name__+'  '+toString(nameVar)
        return

    nameVar = nameVar.upper()

    if nameVar == 'EQUATION':
        try:
            o1.changeData(nameVar, value)
        except Exception as er:
            messError = str(er).replace('SCHEME.EQUATION =',"changeLogic('EQUATION', %s)"%str(value))
            if messError.endswith('Check SCHEME.EQUATION'):
                messError = messError.replace('Check SCHEME.EQUATION','')
        return

    namea = __getData__(o1.__hnd__, OlxAPIConst.LS_vsSignalName)
    typa = dict(zip(namea, __getData__( o1.__hnd__, OlxAPIConst.LS_vnSignalType)))
    vala = dict(zip(namea, __getData__( o1.__hnd__, OlxAPIConst.LS_vdSignalVar)))

    if nameVar not in namea:
        messError = '\nSCHEME.changeLogic(%s , %s)' % (str(nameVar) , str(value))
        messError += "\n\tName Var ('%s') not found in : %s" % (nameVar, str(namea)) +" or 'EQUATION'"
        return

    if typa[nameVar] == OlxAPIConst.LVT_CONST:
        try:
            vala[nameVar] = float(value)
            __scheme_setData__(o1, namea, typa, vala)
        except:
            messError = "\n\tLogicVar %s: must be a float/int \n\tFound: "%toString(nameVar)+toString(value)
            return
        o1.postData()
        return
    #
    if type(value) != list or len(value) != 2:
        messError = '\n\nSCHEME.changeLogic(%s , %s)' % (str(nameVar) , str(value))
        messError += "\n\tLogicVar %s: must be a [str, str/object]\n\t"%toString(nameVar)+__getErrValue__(list,value)
        return

    if value[0] != None:
        try:
            typa[nameVar] = __OLXOBJ_SCHEME__[value[0].upper()]
        except:
            messError = '\nAll value for logic var : '+toString(list(__OLXOBJ_SCHEME__.keys()))
            messError += '\nSCHEME.changeLogic(%s , %s)' % (str(nameVar) , str(value))
            messError += '\n\n\tValueError : '+toString(value[0])
            return

    if value[1] != None:
        ox = None
        if type(value[1]) == str:
            ox = OLCase.findOBJ(value[1])
        else:
            for ob1 in __OLXOBJ_SCHEME_OB__:
                try:
                    ox = OLCase.findOBJ(ob1,value[1])
                    if ox!=None:
                        break
                except:
                    pass
        if ox == None:
            messError = '\nSCHEME.changeLogic(%s , %s)' % (str(nameVar) , str(value))
            messError += '\n\tObject not found : '+toString(value[1])
            return

        vala[nameVar] = ox.__hnd__
    __scheme_setData__(o1, namea, typa, vala)

    try:
        o1.postData()
    except Exception as err:
        messError = '\nSCHEME.changeLogic(%s , %s)' % (str(nameVar) , str(value))
        messError += str(err)

def __scheme_setLogic__(o1, logic):
    if __check_currFileIdx__(o1):
        return
    global messError
    messError = ''

    if type(logic) != dict:
        messError = 'SCHEME.setLogic(logic)'
        messError += "\n\n\tlogic %s: must be a dict\n\n"%toString(logic)+__getErrValue__(dict,logic)
        return

    la,leq = {},''
    for k, v in logic.items():
        try:
            la[k.upper()] = v
        except:
            messError = '\nSCHEME.setLogic(logic)'
            messError += "\n\t%s : %s"%(toString(k), toString(v))
            messError += "\n\tlogic : must be a dict with string keys"
            messError += '\n\t'+__getErrValue__(str,k)
            return

    if 'EQUATION' in la.keys():
        if type(la['EQUATION'])!=str:
            messError = '\nSCHEME.setLogic(logic)'
            messError += "\n\t'EQUATION' : %s"%toString(la['EQUATION'])
            messError += "\n\t'EQUATION' : must be a String"
            messError += '\n\t'+__getErrValue__(str,la['EQUATION'])
            return

        val1 = c_char_p(encode3(la['EQUATION']))
        if OLXAPI_OK != SetData(c_int(o1.__hnd__), c_int(OlxAPIConst.LS_sEquation), byref(val1)):
            messError = '\nSCHEME.setLogic(logic)'
            messError += "\n\t'EQUATION' : %s"%toString(la['EQUATION'])
            messError += ErrorString()
            return

        leq = la['EQUATION']
        la.pop('EQUATION')

    namea, typa, vala = list(la.keys()), dict(), dict()
    namea.sort()

    for k,v in la.items():
        try:
            v1 = float(v)
        except:
            v1 = v
        #
        if type(v1)==float:
            typa[k] = OlxAPIConst.LVT_CONST
            vala[k] = v1
        elif type(v1) == list and len(v1) == 2:
            try:
                typa[k] = __OLXOBJ_SCHEME__[v1[0].upper()]
            except:
                messError = '\nSCHEME.setLogic(logic)'
                messError += '\nAll value for logic var : '+toString(list(__OLXOBJ_SCHEME__.keys()))
                messError += "\n\n\t%s : %s"%(toString(k), toString(v))
                messError += '\n\tValueError : '+toString(v1[0])
                return
            #
            ox = None
            if type(v1[1]) == str:
                ox = OLCase.findOBJ(v1[1])
            else:
                for ob1 in __OLXOBJ_SCHEME_OB__:
                    try:
                        ox = OLCase.findOBJ(ob1, v1[1])
                        if ox != None:
                            break
                    except:
                        pass
            if ox == None:
                messError = '\nSCHEME.setLogic(logic)'
                messError += '\n\t%s : %s'%(toString(k), toString(v))
                messError += '\n\tObject not found : '+toString(v1[1])
                return
            vala[k] = ox.__hnd__
        else:
            messError = '\nSCHEME.setLogic(logic)'
            messError += '\n\t%s : %s'%(toString(k), toString(v))
            messError += '\n\tlogic var %s: must be int/float or [str,str/obj]\n\tFound: '%toString(k)+toString(v1)
            return

    __scheme_setData__(o1, namea, typa, vala)
    try:
        o1.postData()
    except Exception as err:
        messError = '\nSCHEME.setLogic(logic)'
        if leq:
            messError += "\n\t'EQUATION' : %s"%toString(leq)
        messError += str(err)

def __scheme_setData__(o1, namea, typa, vala):
    global messError
    messError = ''
    #
    nn = len(namea)
    signalName = __voidArray__()
    for i in range(nn):
        signalName[i] = cast(pointer(c_char_p(encode3(namea[i]))), c_void_p)
    #
    signalName[nn] = cast(pointer(c_char_p(encode3(''))), c_void_p)
    if OLXAPI_OK != SetData(o1.__hnd__, c_int(OlxAPIConst.LS_vsSignalName), pointer(signalName)):
        messError = ErrorString()
        return
    #
    signalVar = __doubleArray__()
    signalType = __intArray__()
    for i in range(nn):
        signalType[i] = c_int(typa[namea[i]])
        signalVar[i] = c_double(vala[namea[i]])
    #
    signalType[nn] = c_int(0)  # terminator
    if OLXAPI_OK != SetData(o1.__hnd__, c_int(OlxAPIConst.LS_vnSignalType), pointer(signalType)):
        messError = ErrorString()
        return
    if OLXAPI_OK != SetData(o1.__hnd__, c_int(OlxAPIConst.LS_vdSignalVar), pointer(signalVar)):
        messError = ErrorString()
        return
    #
    try:
        o1.postData()
    except Exception as err:
        messError = str(err)


def __setValue__(hnd, paramCode, value):
    vt = paramCode//100
    if vt == OlxAPIConst.VT_STRING:
        return c_char_p(encode3(value))
    elif vt in [0, OlxAPIConst.VT_INTEGER]:
        try:
            return c_int(value.__hnd__)
        except:
            return c_int(int(value))
    elif vt == OlxAPIConst.VT_DOUBLE:
        return c_double(float(value))
    elif vt == OlxAPIConst.VT_ARRAYINT:
        return (c_int*len(value))(*value)
    elif vt in {OlxAPIConst.VT_ARRAYDOUBLE, OlxAPIConst.VT_ARRAYSTRING}:
        return (c_double*len(value))(*value)
    raise Exception("Error of paramCode")


def __findBusNeibor__(b1a, bs1, bs2, ignoreTapBus):
    res = []
    res.extend(b1a)
    for b1 in b1a:
        h1 = b1.__hnd__
        if h1 not in bs1:
            bs1.add(h1)
            for t1 in b1.TERMINAL:
                b23 = [t1.BUS2]
                b3 = t1.BUS3
                if b3:
                    b23.append(b3)
                #
                for bi in b23:
                    hi = bi.__hnd__
                    if hi not in bs2:
                        if ignoreTapBus and bi.TAP > 0:
                            res.extend(__findBusNeibor__([bi], bs1, bs2, 1))
                        else:
                            res.append(bi)
                        bs2.add(hi)
    return res


def __testFI__(key, n):
    flag = type(key) == list and len(key) == n
    if flag:
        for i in range(n):
            key[i] = OlxAPIConst.OLXAPI_DFF if key[i] is None else key[i]
            #
            try:
                key[i] = float(key[i])
            except:
                pass
            #
            if type(key[i]) != float:
                flag = False
                break
    if not flag:
        se = '\n\tRequired           : [float/int]*'+str(n)
        se += '\n\t'+__getErrValue__(float, key)
        return se
    return ''


def __testI__(key, n):
    flag = type(key) == list and len(key) == n
    if flag:
        for i in range(n):
            key[i] = OlxAPIConst.OLXAPI_DFI if key[i] is None else key[i]
            key[i] = __convert2Int__(key[i])
            if type(key[i]) != int:
                flag = False
                break
    if not flag:
        se = '\n\tRequired           : [int]*'+str(n)
        se += '\n\t'+__getErrValue__(int, key)
        return se
    return ''


def __updateOBJNew__(o1, param, setting={}, verbose=True):
    global messError
    messError = ''
    se = "\nOLCase.addOBJ('%s', key, param" % o1.__ob__
    if o1.__ob__ == 'SCHEME':
        se += ', logic)'
    else:
        se += ', setting)' if o1.__ob__ in __OLXOBJ_RELAY3__ else ')'
    #
    s1, s2 = __check_param_setting__(o1, o1.__ob__, param, setting)
    if s2:
        messError = s1+se+s2
        return
    #
    if o1.__ob__ in {'RLYDSG', 'RLYDSP'} and 'DSTYPE' in param.keys():
        if o1.DSTYPE != param['DSTYPE']:
            messError = se+' : already exists\n%s.DSTYPE cannot be changed ' % o1.__ob__+toString(o1.DSTYPE)+'=>'+toString(param['DSTYPE'])
            return
    #
    for k, v in param.items():
        flag = v is None
        if k in __OLXOBJ_PARA__[o1.__ob__].keys() and __OLXOBJ_PARA__[o1.__ob__][k][2] == 0:
            flag = True
        if o1.__ob__ == 'BUS' and k == 'SLACK' and int(v) > 0:
            g = o1.GEN
            if g != None and g.REG == 0 and g.ILIMIT1 <= 0 and g.ILIMIT2 <= 0:
                flag = False
            else:
                flag = True
        if o1.__ob__ in {'RLYDSG', 'RLYDSP'} and k == 'DSTYPE':
            flag = True
        #
        if not flag:
            try:
                o1.changeData(k, v)
            except Exception as err:
                messError = se+'\t '+str(err)
                return
    #
    if o1.__ob__ in __OLXOBJ_RELAY3__:
        for k, v in setting.items():
            try:
                o1.changeSetting(k, v)
            except Exception as err:
                messError = se+'\t '+str(err)
                return
    #
    if o1.__ob__ == 'SCHEME':
        try:
            o1.setLogic(setting)
        except Exception as err:
            messError = se+'\t '+str(err)
            return
    try:
        o1.postData()
    except:
        messError = se+'\n\t'+ErrorString()
        return
    if __OLXOBJ_VERBOSE__ and verbose:
        print("\nOLCase.addOBJ("+toString(o1.__ob__)+') (already exists) : '+o1.toString())
    return o1


def __updateSTR1__(s1):
    if type(s1)==str:
        s2 = s1.upper()
        try:
            return __OLXOBJ_STR_KEYS1__[s2]
        except:
            return s2
    return s1


def __updateSTR__(s1):
    if type(s1) != str:
        return s1
    s2 = s1.strip()
    try:
        k1,k2 = s2.index('['),s2.index(']')
    except:
        k1, k2 = -1, -1
    if k1<0 or k2<0:
        s2 = s2.upper()
        try:
            return __OLXOBJ_STR_KEYS__[s2]
        except:
            return s2

    key = s2[k1+1:k2].strip().upper()
    try:
        key = __OLXOBJ_STR_KEYS__[key]
    except:
        pass
    return '['+key+']'+s2[k2+1:]


def __voltageFault__(obj, style):
    vd9Mag = (c_double*9)(0)
    vd9Ang = (c_double*9)(0)
    if OLXAPI_FAILURE == OlxAPI.GetSCVoltage(obj.__hnd__, vd9Mag, vd9Ang, c_int(style)):
        raise Exception(ErrorString())
    #
    if type(obj) in {XFMR, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}:
        return __resultComplex__(vd9Mag[:6], vd9Ang[:6])
    if type(obj) == XFMR3:
        return __resultComplex__(vd9Mag[:9], vd9Ang[:9])
    return __resultComplex__(vd9Mag[:3], vd9Ang[:3])


__intArray__ = c_int*OlxAPIConst.MXOBJPARAMS
__doubleArray__ = c_double*OlxAPIConst.MXOBJPARAMS
__voidArray__ = c_void_p*OlxAPIConst.MXOBJPARAMS
__doubleArray1__ = c_double*20
__intArray1__ = c_int*50
__fi = {float, int}
__fis = {float, int, str}
__OLXOBJ_NEWADD_DEFAULT__ = 1
__OLXOBJ_VERBOSE__ = 1
__OLXOBJ_VERBOSE1__ = 1

__OLXOBJ_CONST__ = {
    'BUS': [OlxAPIConst.TC_BUS, 'BUS_', '[BUS] List of Buses'],
    'GEN': [OlxAPIConst.TC_GEN, 'GE_', '[GEN] List of Generators'],
    'GENUNIT': [OlxAPIConst.TC_GENUNIT, 'GU_', '[GENUNIT] List of Generator Units'],
    'GENW3': [OlxAPIConst.TC_GENW3, 'G3_', '[GENW3] List of Type-3 Wind Plants GENW3'],
    'GENW4': [OlxAPIConst.TC_GENW4, 'G4_', '[GENW4] List of Converter-Interfaced Resources GENW4'],
    'CCGEN': [OlxAPIConst.TC_CCGEN, 'CC_', '[CCGEN] List of Voltage Controlled Current Sources CCGEN'],
    'XFMR': [OlxAPIConst.TC_XFMR, 'XR_', '[XFMR] List of 2-Windings Transformers'],
    'XFMR3': [OlxAPIConst.TC_XFMR3, 'X3_', '[XFMR3] List of 3-Windings Transformers'],
    'SHIFTER': [OlxAPIConst.TC_PS, 'PS_', '[SHIFTER] List of Phase Shifters'],
    'LINE': [OlxAPIConst.TC_LINE, 'LN_', '[LINE] List of AC Transmission Lines'],
    'DCLINE2': [OlxAPIConst.TC_DCLINE2, 'DC_', '[DCLINE2] List of DC Lines'],
    'MULINE': [OlxAPIConst.TC_MU, 'MU_', '[MULINE] List of Mutual Coupling Pairs'],
    'SERIESRC': [OlxAPIConst.TC_SCAP, 'SC_', '[SERIESRC] List of Series capacitor/reactors'],
    'SWITCH': [OlxAPIConst.TC_SWITCH, 'SW_', '[SWITCH] List of Switches'],
    'LOAD': [OlxAPIConst.TC_LOAD, 'LD_', '[LOAD] List of Loads'],
    'LOADUNIT': [OlxAPIConst.TC_LOADUNIT, 'LU_', '[LOADUNIT] List of Load Units'],
    'SHUNT': [OlxAPIConst.TC_SHUNT, 'SH_', '[SHUNT] List of Shunts'],
    'SHUNTUNIT': [OlxAPIConst.TC_SHUNTUNIT, 'SU_', '[SHUNTUNIT] List of Shunt Units'],
    'SVD': [OlxAPIConst.TC_SVD, 'SV_', '[SVD] List of Switched Shunts'],
    'BREAKER': [OlxAPIConst.TC_BREAKER, 'BK_', '[BREAKER] List of Breakers Rating'],
    'RLYGROUP': [OlxAPIConst.TC_RLYGROUP, 'RG_', '[RLYGROUP] List of Relay Groups'],
    'RLYOCG': [OlxAPIConst.TC_RLYOCG, 'OG_', '[RLYOCG] List of Overcurrent Ground Relays'],
    'RLYOCP': [OlxAPIConst.TC_RLYOCP, 'OP_', '[RLYOCP] List of Overcurrent Phase Relays'],
    'RLYOC': [0, 'OP_', '[RLYOCG/RLYOCP] List of Overcurrent Relays (Ground+Phase)'],
    'FUSE': [OlxAPIConst.TC_FUSE, 'FS_', '[FUSE] List of Fuses'],
    'RLYDSG': [OlxAPIConst.TC_RLYDSG, 'DG_', '[RLYDSG] List of Distance Ground Relays'],
    'RLYDSP': [OlxAPIConst.TC_RLYDSP, 'DP_', '[RLYDSP] List of Distance Phase Relays'],
    'RLYDS': [0, 'DP_', '[RLYDSG/RLYDSP] List of Distance Relays (Ground+Phase)'],
    'RLYD': [OlxAPIConst.TC_RLYD, 'RD_', '[RLYD] List of Differential Relays'],
    'RLYV': [OlxAPIConst.TC_RLYV, 'RV_', '[RLYV] List of Voltage Relays'],
    'RECLSR': [OlxAPIConst.TC_RECLSRP, 'CP_', '[RECLSR] List of Reclosers'],
    'SCHEME': [OlxAPIConst.TC_SCHEME, 'LS_', '[SCHEME] List of Logic Schemes '],
    'ZCORRECT': [OlxAPIConst.TC_ZCORRECT, 'ZC_', '[ZCORRECT] List of Impedance Correction Tables'],
    'TERMINAL': [TC_BRANCH, 'BR_', '[TERMINAL] List of TERMINALs']}


__OLXOBJ_CONST1__ = {
    OlxAPIConst.TC_BUS: ['BUS', BUS],
    OlxAPIConst.TC_GEN: ['GEN', GEN],
    OlxAPIConst.TC_GENUNIT: ['GENUNIT', GENUNIT],
    OlxAPIConst.TC_GENW3: ['GENW3', GENW3],
    OlxAPIConst.TC_GENW4: ['GENW4', GENW4],
    OlxAPIConst.TC_CCGEN: ['CCGEN', CCGEN],
    OlxAPIConst.TC_XFMR: ['XFMR', XFMR, ' T'],
    OlxAPIConst.TC_XFMR3: ['XFMR3', XFMR3, ' X'],
    OlxAPIConst.TC_PS: ['SHIFTER', SHIFTER, ' P'],
    OlxAPIConst.TC_LINE: ['LINE', LINE, ' L'],
    OlxAPIConst.TC_DCLINE2: ['DCLINE2', DCLINE2, ' DC'],
    OlxAPIConst.TC_MU: ['MULINE', MULINE],
    OlxAPIConst.TC_SCAP: ['SERIESRC', SERIESRC, ' S'],
    OlxAPIConst.TC_SWITCH: ['SWITCH', SWITCH, ' W'],
    OlxAPIConst.TC_LOAD: ['LOAD', LOAD],
    OlxAPIConst.TC_LOADUNIT: ['LOADUNIT', LOADUNIT],
    OlxAPIConst.TC_SHUNT: ['SHUNT', SHUNT],
    OlxAPIConst.TC_SHUNTUNIT: ['SHUNTUNIT', SHUNTUNIT],
    OlxAPIConst.TC_SVD: ['SVD', SVD],
    OlxAPIConst.TC_BREAKER: ['BREAKER', BREAKER],
    OlxAPIConst.TC_RLYGROUP: ['RLYGROUP', RLYGROUP],
    OlxAPIConst.TC_RLYOCG: ['RLYOCG', RLYOCG],
    OlxAPIConst.TC_RLYOCP: ['RLYOCP', RLYOCP],
    OlxAPIConst.TC_FUSE: ['FUSE', FUSE],
    OlxAPIConst.TC_RLYDSG: ['RLYDSG', RLYDSG],
    OlxAPIConst.TC_RLYDSP: ['RLYDSP', RLYDSP],
    OlxAPIConst.TC_RLYD: ['RLYD', RLYD],
    OlxAPIConst.TC_RLYV: ['RLYV', RLYV],
    OlxAPIConst.TC_RECLSRP: ['RECLSR', RECLSR],
    OlxAPIConst.TC_SCHEME: ['SCHEME', SCHEME],
    OlxAPIConst.TC_ZCORRECT: ['ZCORRECT', ZCORRECT]}

__OLXOBJ_SETTING__ = {'RLYOCG': [OlxAPIConst.OG_nParamCount, OlxAPIConst.OG_vParamLabels, OlxAPIConst.OG_vParams, OlxAPIConst.OG_sParam, OlxAPIConst.OG_vdDirSettingV15],
                      'RLYOCP': [OlxAPIConst.OP_nParamCount, OlxAPIConst.OP_vParamLabels, OlxAPIConst.OP_vParams, OlxAPIConst.OP_sParam, OlxAPIConst.OP_vdDirSettingV15],
                      'RLYDSG': [OlxAPIConst.DG_nParamCount, OlxAPIConst.DG_vParamLabels, OlxAPIConst.DG_vParams, OlxAPIConst.DG_sParam],
                      'RLYDSP': [OlxAPIConst.DP_nParamCount, OlxAPIConst.DP_vParamLabels, OlxAPIConst.DP_vParams, OlxAPIConst.DP_sParam]}

__OLXOBJ_OCSETTING__ = {255: [[], []],
                        0: [['Characteristic angle', 'Forward pickup', 'Reverse pickup'], [1, 2, 3]],
                        1: [['Characteristic angle', 'Forward pickup', 'Reverse pickup'], [1, 2, 3]],
                        2: [['PTR: PT ratio', 'Z2F: Fwd threshold', 'Z2R: Rev threshold', 'a2: I2/I1 restr.factor', 'k2: i2/I0 restr.factor', 'Z1ANG: Line Z1 angle', '50QF: Fwd 3I2 pickup', '50QR: Rev 3I2 pickup'], [1, 2, 3, 4, 5, 6, 7, 8]],
                        3: [['PTR: PT ratio', 'Z0F: Fwd threshold', 'Z0R: Rev threshold', 'a0: I0/I1 restr.factor', 'Z0ANG: Line Z0 angle', '50GF: Fwd 3Io pickup', '50GR: Rev 3Io pickup'], [1, 2, 3, 4, 5, 6, 7]]}

__OLXOBJ_LISTT__ = {list, set, tuple, _collections_abc.dict_keys}
__OLXOBJ_GRAPHIC__ = {'BUS', 'GEN', 'GENW3', 'GENW4', 'CCGEN', 'LOAD', 'SHUNT', 'SVD', 'XFMR', 'SHIFTER', 'SERIESRC', 'SWITCH', 'DCLINE2', 'XFMR3', 'LINE'}
__OLXOBJ_EQUIPMENT__ = {'XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH'}
__OLXOBJ_EQUIPMENTB1__ = {'XFMR3': XFMR3, 'XFMR': XFMR, 'SHIFTER': SHIFTER, 'LINE': LINE, 'DCLINE2': DCLINE2, 'SERIESRC': SERIESRC, 'SWITCH': SWITCH,
                          'X': XFMR3, 'T': XFMR, 'P': SHIFTER, 'L': LINE, 'DC': DCLINE2, 'S': SERIESRC, 'W': SWITCH}

__OLXOBJ_EQUIPMENTB2__ = ['X', 'T', 'P', 'L', 'DC', 'S', 'W']
__OLXOBJ_EQUIPMENTB3__ = ['XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH']
__OLXOBJ_EQUIPMENTB4__ = {'X', 'T', 'P', 'L', 'DC', 'S', 'W', 'XFMR3', 'XFMR', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH'}
__OLXOBJ_EQUIPMENTB5__ = {'LINE': 1, 'L': 1, 'XFMR': 2, 'T': 2, 'XFMR3': 8, 'X': 3, 'SHIFTER': 4, 'P': 4, 'SWITCH': 16, 'W': 16}
__OLXOBJ_EQUIPMENTB6__ = {1: 'L', 2: 'T', 3:'X', 4: 'P', 16: 'W'}

__OLXOBJ_EQUIPMENTL__ = ['LINE', 'XFMR3', 'XFMR', 'SHIFTER', 'DCLINE2', 'SERIESRC', 'SWITCH']
__OLXOBJ_EQUIPMENTO__ = {XFMR3, XFMR, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH}

__OLXOBJ_RELAY__ = {'RLYOCG', 'RLYOCP', 'FUSE', 'RLYDSG', 'RLYDSP', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME'}
__OLXOBJ_RELAY2__ = {'RLYGROUP', 'RLYOCG', 'RLYOCP', 'FUSE', 'RLYDSG', 'RLYDSP', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME'}
__OLXOBJ_RELAY3__ = {'RLYOCG', 'RLYOCP', 'RLYDSG', 'RLYDSP'}  # relay with setting

__OLXOBJ_BUS__ = {'BREAKER', 'CCGEN', 'DCLINE2', 'GEN', 'GENW3', 'GENW4', 'LINE', 'LOAD', 'SHIFTER', 'RLYGROUP',
                  'SERIESRC', 'SHUNT', 'SVD', 'SWITCH', 'XFMR', 'XFMR3', 'LOADUNIT', 'SHUNTUNIT', 'GENUNIT', 'TERMINAL'}

__OLXOBJ_BUS1__ = {'LOAD', 'SHUNT', 'SVD', 'GEN', 'GENW3', 'GENW4', 'CCGEN'}
__OLXOBJ_BUS2__ = {'LOADUNIT', 'SHUNTUNIT', 'GENUNIT'}

__OLXOBJ_SCHEME_OB__ = ['RLYD', 'RLYV', 'RLYDSG', 'RLYDSP', 'RLYOCG', 'RLYOCP', 'SCHEME', 'TERMINAL']

__OLXOBJ_OBJECT__ = {'BUS': BUS, 'GEN': GEN, 'GENUNIT': GENUNIT, 'GENW3': GENW3, 'GENW4': GENW4, 'CCGEN': CCGEN, 'XFMR': XFMR, 'XFMR3': XFMR3, 'SHIFTER': SHIFTER,
                     'LINE': LINE, 'DCLINE2': DCLINE2, 'MULINE': MULINE, 'SERIESRC': SERIESRC, 'SWITCH': SWITCH, 'LOAD': LOAD, 'LOADUNIT': LOADUNIT, 'SHUNT': SHUNT, 'SHUNTUNIT': SHUNTUNIT,
                     'SVD': SVD, 'BREAKER': BREAKER, 'RLYGROUP': RLYGROUP, 'RLYOCG': RLYOCG, 'RLYOCP': RLYOCP, 'FUSE': FUSE, 'RLYDSG': RLYDSG,
                     'RLYDSP': RLYDSP, 'RLYD': RLYD, 'RLYV': RLYV, 'RECLSR': RECLSR, 'SCHEME': SCHEME, 'ZCORRECT': ZCORRECT, 'TERMINAL': TERMINAL}

__OLXOBJ_LIST__ = ['BUS', 'GEN', 'GENUNIT', 'GENW3', 'GENW4', 'CCGEN', 'XFMR', 'XFMR3', 'SHIFTER', 'LINE', 'DCLINE2', 'MULINE', 'SERIESRC', 'SWITCH',
                   'LOAD', 'LOADUNIT', 'SHUNT', 'SHUNTUNIT', 'SVD', 'BREAKER', 'RLYGROUP', 'RLYOCG', 'RLYOCP', 'FUSE', 'RLYDSG', 'RLYDSP', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME', 'ZCORRECT', 'TERMINAL']

__OLXOBJ_LISTUDF__ = ['BUS', 'GEN', 'GENUNIT', 'GENW3', 'GENW4', 'CCGEN', 'XFMR', 'XFMR3', 'SHIFTER', 'LINE', 'DCLINE2', 'SERIESRC', 'SWITCH', 'LOAD',
                      'LOADUNIT', 'SHUNT', 'SHUNTUNIT', 'SVD', 'BREAKER', 'RLYGROUP', 'RLYOC', 'FUSE', 'RLYDS', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME', 'PROJECT']
__OLXOBJ_IFLT__ = [XFMR, XFMR3, SHIFTER, LINE, DCLINE2, SERIESRC, SWITCH, GEN, GENUNIT, GENW3, GENW4, CCGEN, LOAD, LOADUNIT, SHUNT, SHUNTUNIT, TERMINAL]

__OLXOBJ_RLYSET__ = {}
__OLXOBJ_PARA__ = {}

__OLXOBJ_PARA__['SYS2'] = {
    'BASEMVA': [OlxAPIConst.SY_dBaseMVA, '(float) Network System MVA Base'],
    'COMMENT': [OlxAPIConst.SY_sFComment, '(str) File comments'],
    'OBJCOUNT': [0, '(dict) Number of Object in Network'],
    'AREA': [0, '[AREA] List of AREAs Object in Network'],
    'AREANO': [0, '[int] List of AREAs Number in Network'],
    'ZONE': [0, '[AREA] List of ZONEs Object in Network'],
    'ZONENO': [0, '[int] List of ZONEs Number in Network'],
    'KV': [0, '[float] List of kV Nominal in Network']}

__OLXOBJ_PARA__['SYS'] = {
    'DCLINE': [OlxAPIConst.SY_nNODCLine2, '(int) Number of DC Lines in Network'],
    'IED': [OlxAPIConst.SY_nNOIED, '(int) Number of IED in Network'],
    'AREA': [OlxAPIConst.SY_nNOarea, '(int) Number of Area in Network'],
    'BREAKER': [OlxAPIConst.SY_nNObreaker, '(int) Number of Breakers Rating in Network'],
    'BUS': [OlxAPIConst.SY_nNObus, '(int) Number of Buses in Network'],
    'CCGEN': [OlxAPIConst.SY_nNOccgen, '(int) Number of Voltage controlled current sources-CCGEN in Network'],
    'FUSE': [OlxAPIConst.SY_nNOfuse, '(int) Number of Fuses in Network'],
    'GEN': [OlxAPIConst.SY_nNOgen, '(int) Number of Generators in Network'],
    'GENUNIT': [OlxAPIConst.SY_nNOgenUnit, '(int) Number of Generator Units in Network'],
    'GENW3': [OlxAPIConst.SY_nNOgenW3, '(int) Number of Type-3 Wind Plant-GENW3 in Network'],
    'GENW4': [OlxAPIConst.SY_nNOgenW4, '(int) Number of Converter-Interfaced Resource-GENW4 in Network'],
    'LINE': [OlxAPIConst.SY_nNOline, '(int) Number of AC Transmission Lines in Network'],
    'LOAD': [OlxAPIConst.SY_nNOload, '(int) Number of Loads in Network'],
    'LOADUNIT': [OlxAPIConst.SY_nNOloadUnit, '(int) Number of Load Units in Network'],
    'LTC': [OlxAPIConst.SY_nNOltc, '(int) Number of Load Tap Changer of 2-Winding Transformers in Network'],
    'LTC3': [OlxAPIConst.SY_nNOltc3, '(int) Number of Load Tap Changer of 3-Winding Transformers in Network'],
    'MULINE': [OlxAPIConst.SY_nNOmuPair, '(int) Number of Mutual Coupling Pairs in Network'],
    'SHIFTER': [OlxAPIConst.SY_nNOps, '(int) Number of Phase Shifters in Network'],
    'RECLSR': [OlxAPIConst.SY_nNOrecloserP, '(int) Number of Reclosers in Network'],
    'RLYD': [OlxAPIConst.SY_nNOrlyD, '(int) Number of Differential Relays in Network'],
    'RLYDSG': [OlxAPIConst.SY_nNOrlyDSG, '(int) Number of Distance Ground Relays in Network'],
    'RLYDSP': [OlxAPIConst.SY_nNOrlyDSP, '(int) Number of Distance Phase Relays in Network'],
    'RLYGROUP': [OlxAPIConst.SY_nNOrlyGroup, '(int) Number of Relay Group in Network'],
    'RLYOCG': [OlxAPIConst.SY_nNOrlyOCG, '(int) Number of Overcurrent Ground Relays in Network'],
    'RLYOCP': [OlxAPIConst.SY_nNOrlyOCP, '(int) Number of Overcurrent Phase Relays in Network'],
    'RLYV': [OlxAPIConst.SY_nNOrlyV, '(int) Number of Voltage Relays in Network'],
    'SCHEME': [OlxAPIConst.SY_nNOscheme, '(int) Number of Logic Schemes in Network'],
    'SERIESRC': [OlxAPIConst.SY_nNOseriescap, '(int) Number of Series reactors/capacitors in Network'],
    'SHUNT': [OlxAPIConst.SY_nNOshunt, '(int) Number of Shunts in Network'],
    'SHUNTUNIT': [OlxAPIConst.SY_nNOshuntUnit, '(int) Number of Shunt Units in Network'],
    'SVD': [OlxAPIConst.SY_nNOsvd, '(int) Number of Switched Shunts in Network'],
    'SWITCH': [OlxAPIConst.SY_nNOswitch, '(int) Number of Switches in Network'],
    'XFMR': [OlxAPIConst.SY_nNOxfmr, '(int) Number of 2-Winding Transformers in Network'],
    'XFMR3': [OlxAPIConst.SY_nNOxfmr3, '(int) Number of 3-Winding Transformers in Network'],
    'ZCORRECT': [OlxAPIConst.SY_nNOzCorrect, '(int) Number of Impedance Correction Table in Network'],
    'ZONE': [OlxAPIConst.SY_nNOzone, '(int) Number of Zone in Network']}

__OLXOBJ_PARA__['BUS'] = {
    'ANGLEP': [OlxAPIConst.BUS_dAngleP, '(float) Voltage angle (degree) (from power flow solution)', 0],
    'KVP': [OlxAPIConst.BUS_dKVP, '(float) Voltage magnitude (kV) (from power flow solution)', 0],
    'KV': [OlxAPIConst.BUS_dKVnominal, '(float) Voltage nominal', 0],
    'SPCX': [OlxAPIConst.BUS_dSPCx, '(float) State plane coordinate - X', 'float'],
    'SPCY': [OlxAPIConst.BUS_dSPCy, '(float) State plane coordinate - Y', 'float'],
    'AREANO': [OlxAPIConst.BUS_nArea, '(int) Area Number', 'int'],
    'NO': [OlxAPIConst.BUS_nNumber, '(int) Number', 'int'],
    'SLACK': [OlxAPIConst.BUS_nSlack, '(int) System slack bus FLAG: 1-yes; 0-no', 'int'],
    'SUBGRP': [OlxAPIConst.BUS_nSubGroup, '(int) Substation group', 'int'],
    'TAP': [OlxAPIConst.BUS_nTapBus, '(int) Tap bus FLAG: 0-no; 1-tap bus; 3-tap bus of 3-terminal line', 'int'],
    'MIDPOINT': [OlxAPIConst.BUS_nXfmrMidPoint, '(int) 3-winding xfmr mid-point FLAG: 0-no; 1-yes', 'int'],
    'VISIBLE': [OlxAPIConst.BUS_nVisible, '(int) Hide FLAG: 1-visible; -2-hidden; 0-not yet placed', 0],
    'ZONENO': [OlxAPIConst.BUS_nZone, '(int) Zone Number', 'int'],
    'LOCATION': [OlxAPIConst.BUS_sLocation, '(str) Location', 'str'],
    'NAME': [OlxAPIConst.BUS_sName, '(str) Name', 'strk'],

    'PYTHONSTR': [0, '(str) String python of BUS', 0],
    'KEYSTR': [0, '(str) Key defined of BUS', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of BUS', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'AREA': [0, '(AREA) Area Object', 0],
    'ZONE': [0, '(ZONE) Zone Object', 0],
    'BUS': [0, '(BUS) BUS-self', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str'],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs connected to BUS', 0],
    'BREAKER': [0, '[BREAKER] List of Breakers Rating connected to BUS', 0],
    'CCGEN': [0, '(CCGEN) Voltage controlled current sources connected to BUS', 0],
    'DCLINE2': [0, '[DCLINE2] List of DC Lines connected to BUS', 0],
    'GEN': [0, '(GEN) Generator connected to BUS', 0],
    'GENUNIT': [0, '[GENUNIT] List of Generator Units connected to BUS', 0],
    'GENW3': [0, '(GENW3) Type-3 Wind Plants connected to BUS', 0],
    'GENW4': [0, '(GENW4) Converter-Interfaced Resource connected to BUS', 0],
    'LINE': [0, '[LINE] List of AC Transmission Lines connected to BUS', 0],
    'LOAD': [0, '(LOAD) Load connected to BUS', 0],
    'LOADUNIT': [0, '[LOADUNIT] List of Load Units connected to BUS', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of Relays Group connected to BUS', 0],
    'SERIESRC': [0, '[SERIESRC] List of Series reactors/capacitors connected to BUS', 0],
    'SHIFTER': [0, '[SHIFTER] List of Phase Shifters connected to BUS', 0],
    'SHUNT': [0, '(SHUNT) Shunt connected to BUS', 0],
    'SHUNTUNIT': [0, '[SHUNTUNIT] List of Shunt Units connected to BUS', 0],
    'SVD': [0, '(SVD) Switched Shunts connected to BUS', 0],
    'SWITCH': [0, '[SWITCH] List of Switches connected to BUS', 0],
    'XFMR': [0, '[XFMR] List of 2-Winding Transformers connected to BUS', 0],
    'XFMR3': [0, '[XFMR3] List of 3-Winding Transformers connected to BUS', 0]}


__OLXOBJ_PARA__['TERMINAL'] = {
    'BUS': [0, '[BUS] List of Buses of TERMINAL', 0],
    'BUS1': [0, '(BUS) Bus local of TERMINAL', 0],
    'BUS2': [0, '(BUS) 1st Bus opposite of TERMINAL', 0],
    'BUS3': [0, '(BUS) 2nd Bus opposite of TERMINAL', 0],
    'CID': [0, '(str) Circuit ID', 0],
    'KEYSTR': [0, '(str) Key defined of TERMINAL', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that attached TERMINAL', 0],
    'FLAG': [0, '(int) In-service FLAG: 1- active; 2- out-of-service', 0],
    'OPPOSITE': [0, '[TERMINAL] List of TERMINALs that opposite on the EQUIPMENT', 0],
    'REMOTE': [0, '[TERMINAL] List of TERMINALs remote to TERMINAL\
                                    \n\t                             All taps are ignored.\
                                    \n\t                             Close switches are included\
                                    \n\t                             Out of service branches are ignored', 0],
    'RLYGROUP1': [0, '(RLYGROUP) Local RLYGROUP (None if not found)', 0],
    'RLYGROUP2': [0, '(RLYGROUP) 1st opposite RLYGROUP on the EQUIPMENT (None if not found)', 0],
    'RLYGROUP3': [0, '(RLYGROUP) 2nd opposite RLYGROUP on the EQUIPMENT (None if not found)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUPs that attached to TERMINAL \
                                    \n\t                             RLYGROUP[0] : local RLYGROUP (None if not found)\
                                    \n\t                             RLYGROUP[i] : opposite RLYGROUP (None if not found)', 0]}

__OLXOBJ_PARA__['GEN'] = {
    'ILIMIT1': [OlxAPIConst.GE_dCurrLimit1, '(float) Current limit 1', 'float'],
    'ILIMIT2': [OlxAPIConst.GE_dCurrLimit2, '(float) Current limit 2', 'float'],
    'PGEN': [OlxAPIConst.GE_dPgen, '(float) MW (load flow solution)', 0],
    'QGEN': [OlxAPIConst.GE_dQgen, '(float) MVAR (load flow solution)', 0],
    'REFANGLE': [OlxAPIConst.GE_dRefAngle, '(float) Reference angle', 'float'],
    'SCHEDP': [OlxAPIConst.GE_dScheduledP, '(float) Scheduled P', 0],
    'SCHEDQ': [OlxAPIConst.GE_dScheduledQ, '(float) Scheduled Q', 0],
    'SCHEDV': [OlxAPIConst.GE_dScheduledV, '(float) Scheduled V', 'float'],
    'REFV': [OlxAPIConst.GE_dVSourcePU, '(float) Internal voltage source per unit magnitude', 'float'],
    'FLAG': [OlxAPIConst.GE_nActive, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'REG': [OlxAPIConst.GE_nFixedPQ, '(int) Regulation FLAG: 1- PQ; 0- PV', 'int'],
    'BUS': [OlxAPIConst.GE_nBusHnd, '(BUS) BUS that Generator connected to', 0],
    'CNTBUS': [OlxAPIConst.GE_nCtrlBusHnd, '(BUS) BUS that Generator control the voltage', 'BUS'],
    'MVARATE': [0, '(float) Rating MVA', 0],
    'PYTHONSTR': [0, '(str) String python of Generator', 0],
    'KEYSTR': [0, '(str) Key defined of Generator', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Generator', 0],
    'GENUNIT': [0, '[GENUNIT] List of Generator Units', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['GENUNIT'] = {
    'MVARATE': [OlxAPIConst.GU_dMVArating, '(float) Rating MVA', 'float'],
    'PMAX': [OlxAPIConst.GU_dPmax, '(float) Max MW', 'float'],
    'PMIN': [OlxAPIConst.GU_dPmin, '(float) Min MW', 'float'],
    'QMAX': [OlxAPIConst.GU_dQmax, '(float) Max MVAR', 'float'],
    'QMIN': [OlxAPIConst.GU_dQmin, '(float) Min MVAR', 'float'],
    'RG': [OlxAPIConst.GU_dRz, '(float) Grounding Resistance Rg, in Ohm (do not multiply by 3)', 'float'],
    'XG': [OlxAPIConst.GU_dXz, '(float) Grounding reactance Xg, in Ohm (do not multiply by 3)', 'float'],
    'SCHEDP': [OlxAPIConst.GU_dSchedP, '(float) Scheduled P', 'float'],
    'SCHEDQ': [OlxAPIConst.GU_dSchedQ, '(float) Scheduled Q', 'float'],
    'FLAG': [OlxAPIConst.GU_nOnline, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'CID': [OlxAPIConst.GU_sID, '(str) Circuit ID', 'strk'],
    'DATEOFF': [OlxAPIConst.GU_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.GU_sOnDate, '(str) In service date', 'str'],
    'R': [OlxAPIConst.GU_vdR, '[float]*5 Resistances, in pu: [subtransient, synchronous, transient, negative sequence, zero sequence]', 'float5'],
    'X': [OlxAPIConst.GU_vdX, '[float]*5 Reactances, in pu: [subtransient, synchronous, transient, negative sequence, zero sequence]', 'float5'],
    'GEN': [OlxAPIConst.GU_nGenHnd, '(GEN) Generator that Generator Unit located on', 0],
    'PYTHONSTR': [0, '(str) String python of Generator Unit', 0],
    'KEYSTR': [0, '(str) Key defined of Generator Unit', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Generator Unit', 0],
    'BUS': [0, '(BUS) BUS that Generator Unit connected to', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['GENW3'] = {
    'DATEON': [OlxAPIConst.G3_sOnDate, '(str) In service date', 'str'],
    'DATEOFF': [OlxAPIConst.G3_sOffDate, '(str) Out of service date', 'str'],
    'MVA': [OlxAPIConst.G3_dUnitRatedMVA, '(float) MVA unit rated', 'float'],
    'MWR': [OlxAPIConst.G3_dUnitRatedMW, '(float) MW unit rated', 'float'],
    'MW': [OlxAPIConst.G3_dUnitMW, '(float) Unit MW generation', 'float'],
    'IMAXR': [OlxAPIConst.G3_dImaxRsc, '(float) Rotor side limit, in pu', 'float'],
    'IMAXG': [OlxAPIConst.G3_dImaxGsc, '(float) Grid side limit, in pu', 'float'],
    'VMAX': [OlxAPIConst.G3_dMaxV, '(float) Maximum voltage limit, in pu', 'float'],
    'VMIN': [OlxAPIConst.G3_dMinV, '(float) Minimum voltage limit, in pu', 'float'],
    'RR': [OlxAPIConst.G3_dRr, '(float) Rotor resistance, in pu', 'float'],
    'LLR': [OlxAPIConst.G3_dLlr, '(float) Rotor leakage L, in pu', 'float'],
    'RS': [OlxAPIConst.G3_dRs, '(float) Stator resistance, in pu', 'float'],
    'LLS': [OlxAPIConst.G3_dLls, '(float) Stator leakage L, in pu', 'float'],
    'LM': [OlxAPIConst.G3_dLm, '(float) Mutual L, in pu', 'float'],
    'SLIP': [OlxAPIConst.G3_dSlipRated, '(float) Slip at rated kW', 'float'],
    'FLRZ': [OlxAPIConst.G3_dZfilter, '(float) Filter X, in pu', 'float'],
    'FLAG': [OlxAPIConst.G3_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'UNITS': [OlxAPIConst.G3_nNOUnits, '(int) Number of Units', 'int'],
    'CBAR': [OlxAPIConst.G3_nCrowbared, '(int) Crowbarred FLAG: 1-yes; 0-no', 'int'],
    'BUS': [OlxAPIConst.G3_nBusHnd, '(BUS) BUS that GENW3 connected to', 0],
    'PYTHONSTR': [0, '(str) String python of GENW3', 0],
    'KEYSTR': [0, '(str) Key defined of GENW3', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of GENW3', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['GENW4'] = {
    'DATEON': [OlxAPIConst.G4_sOnDate, '(str) In service date', 'str'],
    'DATEOFF': [OlxAPIConst.G4_sOffDate, '(str) Out of service date', 'str'],
    'MVA': [OlxAPIConst.G4_dUnitRatedMVA, '(float) Unit MVA rating', 'float'],
    'MW': [OlxAPIConst.G4_dUnitMW, '(float) Unit MW generation or consumption', 'float'],
    'VLOW': [OlxAPIConst.G4_dVlow, '(float) When +seq(pu)>', 'float'],
    'MAXI': [OlxAPIConst.G4_dMaxI, '(float) Max current', 'float'],
    'MAXILOW': [OlxAPIConst.G4_dMaxIlow, '(float) Max current reduce to', 'float'],
    'MVAR': [OlxAPIConst.G4_dUnitMVAR, '(float) Unit MVAR', 'float'],
    'SLOPE': [OlxAPIConst.G4_dSlopePos, '(float) Slope of +seq', 'float'],
    'SLOPENEG': [OlxAPIConst.G4_dSlopeNeg, '(float) Slope of -seq', 'float'],
    'VMAX': [OlxAPIConst.G4_dMaxV, '(float) Maximum voltage limit', 'float'],
    'VMIN': [OlxAPIConst.G4_dMinV, '(float) Minimum voltage limit', 'float'],
    'FLAG': [OlxAPIConst.G4_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'UNITS': [OlxAPIConst.G4_nNOUnits, '(int) Number of Units', 'int'],
    'CTRLMETHOD': [OlxAPIConst.G4_nControlMethod, '(int) Control method', 'int'],
    'BUS': [OlxAPIConst.G4_nBusHnd, '(BUS) BUS that GENW4 connected to', 0],
    'PYTHONSTR': [0, '(str) String python of GENW4', 0],
    'KEYSTR': [0, '(str) Key defined of GENW4', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of GENW4', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['CCGEN'] = {
    'MVARATE': [OlxAPIConst.CC_dMVArating, '(float) MVA rating', 'float'],
    'VMAXMUL': [OlxAPIConst.CC_dVmax, '(float) Maximum voltage limit in pu', 'float'],
    'VMIN': [OlxAPIConst.CC_dVmin, '(float) Minimum voltage limit in pu', 'float'],
    'BLOCKPHASE': [OlxAPIConst.CC_nBlockOnPhaseV, '(int) Number block on phase', 'int'],
    'FLAG': [OlxAPIConst.CC_nInService, '(int) In-service FLAG: 1-true; 2-false', 'int'],
    'VLOC': [OlxAPIConst.CC_nVloc, '(int) Voltage measurement location 0-Device terminal; 1-Network side of transformer', 'int'],
    'DATEOFF': [OlxAPIConst.CC_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.CC_sOnDate, '(str) In service date', 'str'],
    'A': [OlxAPIConst.CC_vdAng, '[float]*10 Angle', 'float10'],
    'I': [OlxAPIConst.CC_vdI, '[float]*10 Current', 'float10'],
    'V': [OlxAPIConst.CC_vdV, '[float]*10 Voltage', 'float10'],
    'BUS': [OlxAPIConst.CC_nBusHnd, '(BUS) BUS that CCGEN connected to', 0],
    'PYTHONSTR': [0, '(str) String python of CCGEN', 0],
    'KEYSTR': [0, '(str) Key defined of CCGEN', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of CCGEN', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['XFMR'] = {
    'B': [OlxAPIConst.XR_dB, '(float) +seq Susceptance B, in pu', 'float'],
    'B0': [OlxAPIConst.XR_dB0, '(float) zero Susceptance B0, in pu', 'float'],
    'B1': [OlxAPIConst.XR_dB1, '(float) +seq Susceptance B1 (at Bus1), in pu', 'float'],
    'B10': [OlxAPIConst.XR_dB10, '(float) zero Susceptance B10 (at Bus1), in pu', 'float'],
    'B2': [OlxAPIConst.XR_dB2, '(float) +seq Susceptance B2 (at Bus2), in pu', 'float'],
    'B20': [OlxAPIConst.XR_dB20, '(float) zero Susceptance B20 (at Bus2), in pu', 'float'],
    'BASEMVA': [OlxAPIConst.XR_dBaseMVA, '(float) Base MVA for per-unit quantities', 'float'],
    'EQUIPMENT': [0, '(XFMR) XFMR-self', 0],
    'G1': [OlxAPIConst.XR_dG1, '(float) +seq Conductance G1 (at Bus1), in pu', 'float'],
    'G10': [OlxAPIConst.XR_dG10, '(float) zero Conductance G10 (at Bus1), in pu', 'float'],
    'G2': [OlxAPIConst.XR_dG2, '(float) +seq Conductance G2 (at Bus2), in pu', 'float'],
    'G20': [OlxAPIConst.XR_dG20, '(float) zero Conductance G20 (at Bus2), in pu', 'float'],
    'LTCCENTER': [OlxAPIConst.XR_dLTCCenterTap, '(float) LTC center tap', 'float'],
    'LTCSTEP': [OlxAPIConst.XR_dLTCstep, '(float) LTC step size', 'float'],
    'MVA1': [OlxAPIConst.XR_dMVA1, '(float) Rating MVA1', 'float'],
    'MVA2': [OlxAPIConst.XR_dMVA2, '(float) Rating MVA2', 'float'],
    'MVA3': [OlxAPIConst.XR_dMVA3, '(float) Rating MVA3', 'float'],
    'MAXTAP': [OlxAPIConst.XR_dMaxTap, '(float) LTC max tap', 'float'],
    'MAXVW': [OlxAPIConst.XR_dMaxVW, '(float) LTC min controlled quantity limit', 'float'],
    'MINTAP': [OlxAPIConst.XR_dMinTap, '(float) LTC min tap', 'float'],
    'MINVW': [OlxAPIConst.XR_dMinVW, '(float) LTC max controlled quantity limit', 'float'],
    'PRITAP': [OlxAPIConst.XR_dPriTap, '(float) Primary tap', 'float'],
    'R': [OlxAPIConst.XR_dR, '(float) +seq Resistance R, in pu', 'float'],
    'R0': [OlxAPIConst.XR_dR0, '(float) zero Resistance R0, in pu', 'float'],
    'RG1': [OlxAPIConst.XR_dRG1, '(float) Grounding Resistance Rg1, in Ohm', 'float'],
    'RG2': [OlxAPIConst.XR_dRG2, '(float) Grounding Resistance Rg2, in Ohm', 'float'],
    'RGN': [OlxAPIConst.XR_dRGN, '(float) Grounding Resistance Rgn, in Ohm', 'float'],
    'SECTAP': [OlxAPIConst.XR_dSecTap, '(float) Secondary Tap', 'float'],
    'X': [OlxAPIConst.XR_dX, '(float) +seq Reactance X, in pu', 'float'],
    'X0': [OlxAPIConst.XR_dX0, '(float) zero Reactance X0, in pu', 'float'],
    'XG1': [OlxAPIConst.XR_dXG1, '(float) Grounding Reactance Xg1, in Ohm', 'float'],
    'XG2': [OlxAPIConst.XR_dXG2, '(float) Grounding Reactance Xg2, in Ohm', 'float'],
    'XGN': [OlxAPIConst.XR_dXGN, '(float) Grounding Reactance Xgn, in Ohm', 'float'],
    'AUTOX': [OlxAPIConst.XR_nAuto, '(int) Auto transformer FLAG:1-true;0-false', 'int'],
    'FLAG': [OlxAPIConst.XR_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'GANGED': [OlxAPIConst.XR_nLTCGanged, '(int) LTC tag ganged FLAG: 0-False; 1-True', 'int'],
    'PRIORITY': [OlxAPIConst.XR_nLTCPriority, '(int) LTC adjustment priority: 0-Normal; 1-Medieum; 2-High', 'int'],
    'LTCSIDE': [OlxAPIConst.XR_nLTCside, '(int) LTC side: 1; 2; 0', 'int'],
    'LTCTYPE': [OlxAPIConst.XR_nLTCtype, '(int) LTC type: 0- control voltage; 1- control MVAR', 'int'],
    'LTCCTRL': [OlxAPIConst.XR_nLTCCtrlBusHnd, '(BUS) Bus whose voltage magnitude is to be regulated by the LTC', 'BUS0'],
    'CONFIGP': [OlxAPIConst.XR_sCfgP, '(str) Primary winding config', 'str'],
    'CONFIGS': [OlxAPIConst.XR_sCfgS, '(str) Secondary winding config', 'str'],
    'CONFIGST': [OlxAPIConst.XR_sCfgST, '(str) Secondary winding config in test', 'str'],
    'CID': [OlxAPIConst.XR_sID, '(str) Circuit ID', 'strk'],
    'NAME': [OlxAPIConst.XR_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.XR_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.XR_sOnDate, '(str) In service date', 'str'],
    'METEREDEND': [OlxAPIConst.XR_nMeteredEnd, '(int) Metered FLAG: 1-at Bus1; 2 at Bus2; 0 XFMR in a single area', 0],# obsolete
    'TIE': [OlxAPIConst.XR_nMeteredEnd, '(int) Metered FLAG: 1-at Bus1; 2 at Bus2; 0 XFMR in a single area', 'int'],
    'BUS1': [OlxAPIConst.XR_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.XR_nBus2Hnd, '(BUS) Bus2', 0],
    'RLYGROUP1': [OlxAPIConst.XR_nRlyGr1Hnd, '(RLYGROUP) Relay Group at Bus1 (None if not found)', 0],
    'RLYGROUP2': [OlxAPIConst.XR_nRlyGr2Hnd, '(RLYGROUP) Relay Group at Bus2 (None if not found)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUPs [RLYGROUP1,RLYGROUP2]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'PYTHONSTR': [0, '(str) String python of XFMR', 0],
    'KEYSTR': [0, '(str) Key defined of XFMR', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of XFMR', 0],
    'MVA': [0, '[float]*3 Ratings [MVA1,MVA2,MVA3]', 0],
    'BRCODE': [0, "(str) BRCODE='T'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['XFMR3'] = {
    'B': [OlxAPIConst.X3_dB, '(float) +seq Susceptance B, in pu', 'float'],
    'B0': [OlxAPIConst.X3_dB0, '(float) zero Susceptance B0, in pu', 'float'],
    'BASEMVA': [OlxAPIConst.X3_dBaseMVA, '(float) Base MVA for per-unit quantities', 'float'],
    'EQUIPMENT': [0, '(XFMR3) XFMR3-self', 0],
    'LTCCENTER': [OlxAPIConst.X3_dLTCCenterTap, '(float) LTC center tap', 'float'],
    'LTCSTEP': [OlxAPIConst.X3_dLTCstep, '(float) LTC step size', 'float'],
    'LTCSIDE': [OlxAPIConst.X3_nLTCside, '(int) LTC side: 1; 2; 3; 0', 'int'],
    'LTCTYPE': [OlxAPIConst.X3_nLTCtype, '(int) LTC type: 0- control voltage; 1- control MVAR', 'int'],
    'LTCCTRL': [OlxAPIConst.X3_nLTCCtrlBusHnd, '(BUS) Bus whose voltage magnitude is to be regulated by the LTC', 'BUS0'],
    'MVA1': [OlxAPIConst.X3_dMVA1, '(float) Rating MVA1', 'float'],
    'MVA2': [OlxAPIConst.X3_dMVA2, '(float) Rating MVA2', 'float'],
    'MVA3': [OlxAPIConst.X3_dMVA3, '(float) Rating MVA3', 'float'],
    'MAXTAP': [OlxAPIConst.X3_dMaxTap, '(float) LTC max tap', 'float'],
    'MAXVW': [OlxAPIConst.X3_dMaxVW, '(float) LTC min controlled quantity limit', 'float'],
    'MINTAP': [OlxAPIConst.X3_dMinTap, '(float) LTC min tap', 'float'],
    'MINVW': [OlxAPIConst.X3_dMinVW, '(float) LTC controlled quantity limit', 'float'],
    'PRITAP': [OlxAPIConst.X3_dPriTap, '(float) Primary Tap', 'float'],
    'RPS0': [OlxAPIConst.X3_dR0ps, '(float) zero Resistance R0ps, in pu', 'float'],
    'RPT0': [OlxAPIConst.X3_dR0pt, '(float) zero Resistance R0pt, in pu', 'float'],
    'RST0': [OlxAPIConst.X3_dR0st, '(float) zero Resistance R0st, in pu', 'float'],
    'RG1': [OlxAPIConst.X3_dRG1, '(float) Grounding Resistance Rg1, in Ohm', 'float'],
    'RG2': [OlxAPIConst.X3_dRG2, '(float) Grounding Resistance Rg2, in Ohm', 'float'],
    'RG3': [OlxAPIConst.X3_dRG3, '(float) Grounding Resistance Rg3, in Ohm', 'float'],
    'RGN': [OlxAPIConst.X3_dRGN, '(float) Grounding Resistance Rgn, in Ohm', 'float'],
    'RPS': [OlxAPIConst.X3_dRps, '(float) +seq Resistance Rps, in pu', 'float'],
    'RPT': [OlxAPIConst.X3_dRpt, '(float) +seq Resistance Rpt, in pu', 'float'],
    'RST': [OlxAPIConst.X3_dRst, '(float) +seq Resistance Rst, in pu', 'float'],
    'SECTAP': [OlxAPIConst.X3_dSecTap, '(float) Secondary Tap', 'float'],
    'TERTAP': [OlxAPIConst.X3_dTerTap, '(float) Tertiary Tap', 'float'],
    'XPS0': [OlxAPIConst.X3_dX0ps, '(float) zero Reactance X0ps, in pu', 'float'],
    'XPT0': [OlxAPIConst.X3_dX0pt, '(float) zero Reactance X0pt, in pu', 'float'],
    'XST0': [OlxAPIConst.X3_dX0st, '(float) zero Reactance X0st, in pu', 'float'],
    'XG1': [OlxAPIConst.X3_dXG1, '(float) Grounding Reactance Xg1, in Ohm', 'float'],
    'XG2': [OlxAPIConst.X3_dXG2, '(float) Grounding Reactance Xg2, in Ohm', 'float'],
    'XG3': [OlxAPIConst.X3_dXG3, '(float) Grounding Reactance Xg3, in Ohm', 'float'],
    'XGN': [OlxAPIConst.X3_dXGN, '(float) Grounding Reactance Xgn, in Ohm', 'float'],
    'XPS': [OlxAPIConst.X3_dXps, '(float) +seq Reactance Xps, in pu', 'float'],
    'XPT': [OlxAPIConst.X3_dXpt, '(float) +seq Reactance Xpt, in pu', 'float'],
    'XST': [OlxAPIConst.X3_dXst, '(float) +seq Reactance Xst, in pu', 'float'],
    'RPM0': [OlxAPIConst.X3_dXst, '(float) zero Resistance RPM0, in pu', 'float'],
    'XPM0': [OlxAPIConst.X3_dXst, '(float) zero Reactance XPM0, in pu', 'float'],
    'RSM0': [OlxAPIConst.X3_dXst, '(float) zero Resistance RSM0, in pu', 'float'],
    'XSM0': [OlxAPIConst.X3_dXst, '(float) zero Reactance XSM0, in pu', 'float'],
    'RMG0': [OlxAPIConst.X3_dXst, '(float) zero Resistance RMG0, in pu', 'float'],
    'XMG0': [OlxAPIConst.X3_dXst, '(float) XMG0, in pu', 'float'],
    'AUTOX': [OlxAPIConst.X3_nAuto, '(int) Auto transformer FLAG: 1-true;0-false', 'int'],
    'FLAG': [OlxAPIConst.X3_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'Z0METHOD': [OlxAPIConst.X3_nInService, '(int) Z0 method: 1-Short circuit impedance; 2-Classical T model impedance', 'int'],
    'FICTBUSNO': [OlxAPIConst.X3_nFictBusNo, '(int) Fiction Bus Number', 'int'],
    'GANGED': [OlxAPIConst.X3_nLTCGanged, '(int) LTC Tag ganged FLAG: 0-False; 1-True', 'int'],
    'PRIORITY': [OlxAPIConst.X3_nLTCPriority, '(int) LTC adjustment priority: 0-Normal; 1-Medieum; 2-High', 'int'],
    'CONFIGP': [OlxAPIConst.X3_sCfgP, '(str) Primary winding config', 'str'],
    'CONFIGS': [OlxAPIConst.X3_sCfgS, '(str) Secondary winding config', 'str'],
    'CONFIGST': [OlxAPIConst.X3_sCfgST, '(str) Secondary winding config in test', 'str'],
    'CONFIGT': [OlxAPIConst.X3_sCfgT, '(str) Tertiary winding config', 'str'],
    'CONFIGTT': [OlxAPIConst.X3_sCfgTT, '(str) Tertiary winding config in test', 'str'],
    'CID': [OlxAPIConst.X3_sID, '(str) Circuit ID', 'strk'],
    'NAME': [OlxAPIConst.X3_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.X3_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.X3_sOnDate, '(str) In service date', 'str'],
    'BUS1': [OlxAPIConst.X3_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.X3_nBus2Hnd, '(BUS) Bus2', 0],
    'BUS3': [OlxAPIConst.X3_nBus3Hnd, '(BUS) Bus3', 0],
    'RLYGROUP1': [OlxAPIConst.X3_nRlyGr1Hnd, '(RLYGROUP) Relay Group 1 (at Bus1)', 0],
    'RLYGROUP2': [OlxAPIConst.X3_nRlyGr2Hnd, '(RLYGROUP) Relay Group 2 (at Bus2)', 0],
    'RLYGROUP3': [OlxAPIConst.X3_nRlyGr3Hnd, '(RLYGROUP) Relay Group 3 (at Bus3)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUPs [RLYGROUP1,RLYGROUP2,RLYGROUP3]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'PYTHONSTR': [0, '(str) String python of XFMR3', 0],
    'KEYSTR': [0, '(str) Key defined of XFMR3', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of XFMR3', 0],
    'BRCODE': [0, "(str) BRCODE='X'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [Bus1,Bus2,Bus3]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SHIFTER'] = {
    'SHIFTANGLE': [OlxAPIConst.PS_dAngle, '(float) Shift angle', 'float'],
    'ANGMAX': [OlxAPIConst.PS_dAngleMax, '(float) Shift angle max', 'float'],
    'ANGMIN': [OlxAPIConst.PS_dAngleMin, '(float) Shift angle min', 'float'],
    'BN1': [OlxAPIConst.PS_dBN1, '(float) -seq Susceptance Bn1 (at Bus1), in pu', 'float'],
    'BN2': [OlxAPIConst.PS_dBN2, '(float) -seq Susceptance Bn2 (at Bus2), in pu', 'float'],
    'BP1': [OlxAPIConst.PS_dBP1, '(float) +seq Susceptance Bp1 (at Bus1), in pu', 'float'],
    'BP2': [OlxAPIConst.PS_dBP2, '(float) +seq Susceptance Bp2 (at Bus2), in pu', 'float'],
    'BZ1': [OlxAPIConst.PS_dBZ1, '(float) zero Susceptance Bz1 (at Bus1), in pu', 'float'],
    'BZ2': [OlxAPIConst.PS_dBZ2, '(float) zero Susceptance Bz2 (at Bus2), in pu', 'float'],
    'BASEMVA': [OlxAPIConst.PS_dBaseMVA, '(float) BaseMVA (for per-unit quantities)', 'float'],
    'EQUIPMENT': [0, '(SHIFTER) SHIFTER-self', 0],
    'GN1': [OlxAPIConst.PS_dGN1, '(float) -seq Conductance Gn1 (at Bus1), in pu', 'float'],
    'GN2': [OlxAPIConst.PS_dGN2, '(float) -seq Conductance Gn2 (at Bus2), in pu', 'float'],
    'GP1': [OlxAPIConst.PS_dGP1, '(float) +seq Conductance Gp1 (at Bus1), in pu', 'float'],
    'GP2': [OlxAPIConst.PS_dGP2, '(float) +seq Conductance Gp2 (at Bus2), in pu', 'float'],
    'GZ1': [OlxAPIConst.PS_dGZ1, '(float) zero Conductance Gz1 (at Bus1), in pu', 'float'],
    'GZ2': [OlxAPIConst.PS_dGZ2, '(float) zero Conductance Gz2 (at Bus2), in pu', 'float'],
    'MVA1': [OlxAPIConst.PS_dMVA1, '(float) Rating MVA1', 'float'],
    'MVA2': [OlxAPIConst.PS_dMVA2, '(float) Rating MVA2', 'float'],
    'MVA3': [OlxAPIConst.PS_dMVA3, '(float) Rating MVA3', 'float'],
    'MWMAX': [OlxAPIConst.PS_dMWmax, '(float) MW max', 'float'],
    'MWMIN': [OlxAPIConst.PS_dMWmin, '(float) MW min', 'float'],
    'RN': [OlxAPIConst.PS_dRN, '(float) -seq Resistance Rn, in pu', 'float'],
    'RP': [OlxAPIConst.PS_dRP, '(float) +seq Resistance Rp, in pu', 'float'],
    'RZ': [OlxAPIConst.PS_dRZ, '(float) zero Resistance Rz, in pu', 'float'],
    'XN': [OlxAPIConst.PS_dXN, '(float) -seq Reactance Xn, in pu', 'float'],
    'XP': [OlxAPIConst.PS_dXP, '(float) +seq Reactance Xp, in pu', 'float'],
    'XZ': [OlxAPIConst.PS_dXZ, '(float) zero Reactance Xz, in pu', 'float'],
    'CNTL': [OlxAPIConst.PS_nControlMode, '(int) Control mode: 0-Fixed; 1-automatically control real power flow', 'int'],
    'FLAG': [OlxAPIConst.PS_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'ZCORRECTNO': [OlxAPIConst.PS_nZCorrectNO, '(int) Impedance Correct Table Number', 'int'],
    'CID': [OlxAPIConst.PS_sID, '(str) Circuit ID', 'strk'],
    'NAME': [OlxAPIConst.PS_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.PS_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.PS_sOnDate, '(str) In service date', 'str'],
    'BUS1': [OlxAPIConst.PS_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.PS_nBus2Hnd, '(BUS) Bus2', 0],
    'RLYGROUP1': [OlxAPIConst.PS_nRlyGr1Hnd, '(RLYGROUP) Relay Group 1 (at Bus1)', 0],
    'RLYGROUP2': [OlxAPIConst.PS_nRlyGr2Hnd, '(RLYGROUP) Relay Group 2 (at Bus2)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUP(s) [RLYGROUP1,RLYGROUP2]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'PYTHONSTR': [0, '(str) String python of Phase Shifter', 0],
    'KEYSTR': [0, '(str) Key defined of Phase Shifter', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Phase shifter', 0],
    'BRCODE': [0, "(str) BRCODE='P'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['LINE'] = {
    'B1': [OlxAPIConst.LN_dB1, '(float) +seq Susceptance B1 (at BUS1), in pu', 'float'],
    'B10': [OlxAPIConst.LN_dB10, '(float) zero Susceptance B10 (at BUS1), in pu', 'float'],
    'B2': [OlxAPIConst.LN_dB2, '(float) +seq Susceptance B2 (at BUS2), in pu', 'float'],
    'B20': [OlxAPIConst.LN_dB20, '(float) zero Susceptance B20 (at BUS2), in pu', 'float'],
    'G1': [OlxAPIConst.LN_dG1, '(float) +seq Conductance G1 (at BUS1), in pu', 'float'],
    'G10': [OlxAPIConst.LN_dG10, '(float) zero Conductance G10 (at BUS1), in pu', 'float'],
    'G2': [OlxAPIConst.LN_dG2, '(float) +seq Conductance G2 (at BUS2), in pu', 'float'],
    'G20': [OlxAPIConst.LN_dG20, '(float) zero Conductance G20 (at BUS2), in pu', 'float'],
    'LN': [OlxAPIConst.LN_dLength, '(float) Length', 'float'],
    'R': [OlxAPIConst.LN_dR, '(float) +seq Resistance R, in pu', 'float'],
    'R0': [OlxAPIConst.LN_dR0, '(float) zero Resistance Ro, in pu', 'float'],
    'X': [OlxAPIConst.LN_dX, '(float) +seq Reactance X, in pu', 'float'],
    'X0': [OlxAPIConst.LN_dX0, '(float) zero Reactance X0, in pu', 'float'],
    'FLAG': [OlxAPIConst.LN_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'METEREDEND': [OlxAPIConst.LN_nMeteredEnd, '(int) Meteted FLAG: 1- at Bus1; 2-at Bus2; 0-line is in a single area;', 0],# obsolete
    'TIE': [OlxAPIConst.LN_nMeteredEnd, '(int) Meteted FLAG: 1- at Bus1; 2-at Bus2; 0-line is in a single area;', 'int'],
    'I2T': [OlxAPIConst.LN_dI2T, '(float) I^2T rating in ampere^2 Sec.', 'float'],
    'CID': [OlxAPIConst.LN_sID, '(str) Circuit ID', 'strk'],
    'UNIT': [OlxAPIConst.LN_sLengthUnit, '(str) Length unit in [ft,kt,mi,m,km]', 'str'],
    'NAME': [OlxAPIConst.LN_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.LN_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.LN_sOnDate, '(str) In service date', 'str'],
    'TYPE': [OlxAPIConst.LN_sType, '(str) Type', 'str'],
    'RATG': [OlxAPIConst.LN_vdRating, '[float]*4 Ratings, in A', 'float4'],
    'BUS1': [OlxAPIConst.LN_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.LN_nBus2Hnd, '(BUS) Bus2', 0],
    'MULINE': [OlxAPIConst.LN_nMuPairHnd, '[MULINE] List of Mutual Coupling Pairs', 0],
    'RLYGROUP1': [OlxAPIConst.LN_nRlyGr1Hnd, '(RLYGROUP) Relay Group 1 (at Bus1)', 0],
    'RLYGROUP2': [OlxAPIConst.LN_nRlyGr2Hnd, '(RLYGROUP) Relay Group 2 (at Bus2)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUP(s) [RLYGROUP1,RLYGROUP2]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'EQUIPMENT': [0, '(LINE) LINE-self', 0],
    'PYTHONSTR': [0, '(str) String python of LINE', 0],
    'KEYSTR': [0, '(str) Key defined of LINE', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of LINE', 0],
    'BRCODE': [0, "(str) BRCODE='L'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['DCLINE2'] = {
    'NAME': [OlxAPIConst.DC_sName, '(string) Name', 'str'],
    'CID': [OlxAPIConst.DC_sID, '(string) Circuit ID', 'strk'],
    'DATEON': [OlxAPIConst.DC_sOnDate, '(string) In service date', 'str'],
    'DATEOFF': [OlxAPIConst.DC_sOffDate, '(string) Out of service date', 'str'],
    'TARGET': [OlxAPIConst.DC_dControlTarget, '(float) Target MW', 'float'],
    'MARGIN': [OlxAPIConst.DC_dControlMargin, '(float) Margin', 'float'],
    'VSCHED': [OlxAPIConst.DC_dVdcSched, '(float) Schedule DC volts in pu', 'float'],
    'LINER': [OlxAPIConst.DC_dDClineR, '(float) Resistance R, in Ohm', 'float'],
    'FLAG': [OlxAPIConst.DC_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'METEREDEND': [OlxAPIConst.DC_nMeteredEnd, '(int) Meteted FLAG: 1- at Bus1; 2-at Bus2; 0-line is in a single area;', 0],# obsolete
    'TIE': [OlxAPIConst.DC_nMeteredEnd, '(int) Meteted FLAG: 1- at Bus1; 2-at Bus2; 0-line is in a single area;', 'int'],
    'MODE': [OlxAPIConst.DC_nInService, '(int) Control mode', 'int'],
    'ANGMAX': [OlxAPIConst.DC_vdAngleMax, '[float]*2 Angle max', 'float2'],
    'ANGMIN': [OlxAPIConst.DC_vdAngleMin, '[float]*2 Angle min', 'float2'],
    'TAPMAX': [OlxAPIConst.DC_vdTapMax, '[float]*2 Tap max', 'float2'],
    'TAPMIN': [OlxAPIConst.DC_vdTapMax, '[float]*2 Tap min', 'float2'],
    'TAPSTEP': [OlxAPIConst.DC_vdTapStepSize, '[float]*2 Tap step size', 'float2'],
    'TAP': [OlxAPIConst.DC_vdTapRatio, '[float]*2 Tap', 'float2'],
    'MVA': [OlxAPIConst.DC_vdMVA, '[float]*2 MVA rating', 'float2'],
    'NOMKV': [OlxAPIConst.DC_vdNomKV, '[float]*2 Nomibal KV on DC side', 'float2'],
    'R': [OlxAPIConst.DC_vdXfmrR, '[float]*2 Transformer Resistance R, in pu', 'float2'],
    'X': [OlxAPIConst.DC_vdXfmrX, '[float]*2 Transformer Reactance X, in pu', 'float2'],
    'BRIDGE': [OlxAPIConst.DC_vnBridges, '[int]*2 No. of bridges', 'int2'],
    'BUS1': [OlxAPIConst.DC_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.DC_nBus2Hnd, '(BUS) Bus2', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'EQUIPMENT': [0, '(DCLINE2) DCLINE2-self', 0],
    'PYTHONSTR': [0, '(str) String python of DCLINE', 0],
    'KEYSTR': [0, '(str) Key defined of DCLINE', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of DCLINE', 0],
    'BRCODE': [0, "(str) BRCODE='D'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SERIESRC'] = {
    'IPR': [OlxAPIConst.SC_dIpr, '(float) Protective level current', 'float'],
    'SCOMP': [OlxAPIConst.SC_nSComp, '(int) Bypassed FLAG: 1- no bypassed; 2-yes bypassed', 'int'],
    'R': [OlxAPIConst.SC_dR, '(float) Resistance R, in pu', 'float'],
    'X': [OlxAPIConst.SC_dX, '(float) Reactance X, in pu', 'float'],
    'FLAG': [OlxAPIConst.SC_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service; 3- bypassed', 'int'],
    'CID': [OlxAPIConst.SC_sID, '(str) Circuit ID', 'strk'],
    'NAME': [OlxAPIConst.SC_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.SC_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.SC_sOnDate, '(str) In service date', 'str'],
    'BUS1': [OlxAPIConst.SC_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.SC_nBus2Hnd, '(BUS) Bus2', 0],
    'RLYGROUP1': [OlxAPIConst.SC_nRlyGr1Hnd, '(RLYGROUP) Relay Group 1 (at Bus1)', 0],
    'RLYGROUP2': [OlxAPIConst.SC_nRlyGr2Hnd, '(RLYGROUP) Relay Group 2 (at Bus2)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUP(s) [RLYGROUP1,RLYGROUP2]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'EQUIPMENT': [0, '(SERIESRC) SERIESRC-self', 0],
    'PYTHONSTR': [0, '(str) String python of SERIESRC', 0],
    'KEYSTR': [0, '(str) Key defined of SERIESRC', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of SERIESRC', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BRCODE': [0, "(str) BRCODE='S'", 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SWITCH'] = {
    'RATING': [OlxAPIConst.SW_dRating, '(float) Rating', 'float'],
    'DEFAULT': [OlxAPIConst.SW_nDefault, '(int) Default position FLAG: 1- normaly open; 2- normaly close; 0-Not defined', 'int'],
    'FLAG': [OlxAPIConst.SW_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'STAT': [OlxAPIConst.SW_nStatus, '(int) Position FLAG: 7- close; 0- open', 'int'],
    'CID': [OlxAPIConst.SW_sID, '(str) Circuit ID', 'strk'],
    'NAME': [OlxAPIConst.SW_sName, '(str) Name', 'str'],
    'DATEOFF': [OlxAPIConst.SW_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.SW_sOnDate, '(str) In service date', 'str'],
    'BUS1': [OlxAPIConst.SW_nBus1Hnd, '(BUS) Bus1', 0],
    'BUS2': [OlxAPIConst.SW_nBus2Hnd, '(BUS) Bus2', 0],
    'RLYGROUP1': [OlxAPIConst.SW_nRlyGrHnd1, '(RLYGROUP) Relay Group 1 (at Bus1)', 0],
    'RLYGROUP2': [OlxAPIConst.SW_nRlyGrHnd2, '(RLYGROUP) Relay Group 2 (at Bus2)', 0],
    'RLYGROUP': [0, '[RLYGROUP] List of RLYGROUP(s) [RLYGROUP1,RLYGROUP2]', 0],
    'TERMINAL': [0, '[TERMINAL] List of TERMINALs', 0],
    'EQUIPMENT': [0, '(SWITCH) SWITCH-self', 0],
    'PYTHONSTR': [0, '(str) String python of Switch', 0],
    'KEYSTR': [0, '(str) Key defined of Switch', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Switch', 0],
    'BRCODE': [0, "(str) BRCODE='W'", 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'BUS': [0, '[BUS] List of BUSes [BUS1,BUS2]', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['LOAD'] = {
    'P': [OlxAPIConst.LD_dPload, '(float) Total load MW (load flow solution)', 0],
    'Q': [OlxAPIConst.LD_dQload, '(float) Total load MVAR (load flow solution)', 0],
    'FLAG': [OlxAPIConst.LD_nActive, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'UNGROUNDED': [OlxAPIConst.LD_nUnGrounded, '(int) UnGrounded FLAG: 1-UnGrounded; 0-Grounded', 'int'],
    'BUS': [OlxAPIConst.LD_nBusHnd, '(BUS) BUS that Load connected to', 0],
    'PYTHONSTR': [0, '(str) String python of Load', 0],
    'KEYSTR': [0, '(str) Key defined of Load', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Load', 0],
    'LOADUNIT': [0, '[LOADUNIT] List of Load Unit of Load', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['LOADUNIT'] = {
    'P': [OlxAPIConst.LU_dPload, '(float) MW (load flow solution)', 0],
    'Q': [OlxAPIConst.LU_dQload, '(float) MVAR (load flow solution)', 0],
    'FLAG': [OlxAPIConst.LU_nOnline, '(int) In-service FLAG: 1-active; 2- out-of-service', 'int'],
    'CID': [OlxAPIConst.LU_sID, '(str) Circuit ID', 'strk'],
    'DATEOFF': [OlxAPIConst.LU_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.LU_sOnDate, '(str) In service date', 'str'],
    'MVAR': [OlxAPIConst.LU_vdMVAR, '[float]*3 MVARs: [const. P, const. I, const. Z]', 'float3'],
    'MW': [OlxAPIConst.LU_vdMW, '[float]*3 MWs: [const. P, const. I, const. Z]', 'float3'],
    'LOAD': [OlxAPIConst.LU_nLoadHnd, '(LOAD) Load that Load Unit located', 0],
    'PYTHONSTR': [0, '(str) String python of LOADUNIT', 0],
    'KEYSTR': [0, '(str) Key defined of LOADUNIT', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of LOADUNIT', 0],
    'BUS': [0, '(BUS) Bus that Load Unit is connected to', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SHUNT'] = {
    'FLAG': [OlxAPIConst.SH_nActive, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'BUS': [OlxAPIConst.SH_nBusHnd, '(BUS) Bus that Shunt connected to', 0],
    'SHUNTUNIT': [0, '[SHUNTUNIT] List of Shunt Units of Shunt', 0],
    'PYTHONSTR': [0, '(str) String python of Shunt', 0],
    'KEYSTR': [0, '(str) Key defined of Shunt', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Shunt', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SHUNTUNIT'] = {
    'B': [OlxAPIConst.SU_dB, '(float) +seq succeptance', 'float'],
    'B0': [OlxAPIConst.SU_dB0, '(float) zero succeptance', 'float'],
    'G': [OlxAPIConst.SU_dG, '(float) +seq conductance', 'float'],
    'G0': [OlxAPIConst.SU_dG0, '(float) zero conductance', 'float'],
    'TX3': [OlxAPIConst.SU_n3WX, '(int) 3-winding transformer FLAG: 1-true;0-false', 'int'],
    'FLAG': [OlxAPIConst.SU_nOnline, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'CID': [OlxAPIConst.SU_sID, '(str) Circuit ID', 'strk'],
    'DATEOFF': [OlxAPIConst.SU_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.SU_sOnDate, '(str) In service date', 'str'],
    'SHUNT': [OlxAPIConst.SU_nShuntHnd, '(SHUNT) Shunt that Shunt Unit located on', 0],
    'PYTHONSTR': [0, '(str) String python of SHUNTUNIT', 0],
    'KEYSTR': [0, '(str) Key defined of SHUNTUNIT', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of SHUNTUNIT', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SVD'] = {
    'B_USE': [OlxAPIConst.SV_dB, '(float) Susceptance in use', 'float'],
    'VMAX': [OlxAPIConst.SV_dVmax, '(float) Max V', 'float'],
    'VMIN': [OlxAPIConst.SV_dVmin, '(float) Min V', 'float'],
    'FLAG': [OlxAPIConst.SV_nActive, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'MODE': [OlxAPIConst.SV_nCtrlMode, '(int) Control mode: 0-Fixed; 1-Discrete; 2-Continous', 'int'],
    'DATEOFF': [OlxAPIConst.SV_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.SV_sOnDate, '(str) In service date', 'str'],
    'B0': [OlxAPIConst.SV_vdB0inc, '[float]*8  zero Susceptance', 'float8'],
    'B': [OlxAPIConst.SV_vdBinc, '[float]*8 +seq Susceptance', 'float8'],
    'STEP': [OlxAPIConst.SV_vnNoStep, '[int]*8 Number of step', 'int8'],
    'BUS': [OlxAPIConst.SV_nBusHnd, '(BUS) Bus that SVD connected to', 0],
    'CNTBUS': [OlxAPIConst.SV_nCtrlBusHnd, '(BUS) SVD controled Bus', 'BUS'],
    'PYTHONSTR': [0, '(str) String python of SVD', 0],
    'KEYSTR': [0, '(str) Key defined of SVD', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of SVD', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['BREAKER'] = {
    'CPT1': [OlxAPIConst.BK_dCPT1, '(float) Contact parting time for group 1 (cycles)', 'float'],
    'CPT2': [OlxAPIConst.BK_dCPT2, '(float) Contact parting time for group 2 (cycles)', 'float'],
    'INTRTIME': [OlxAPIConst.BK_dCycles, '(float) Interrupting time (cycles)', 'float'],
    'K': [OlxAPIConst.BK_dK, '(float) kV range factor', 'float'],
    'NACD': [OlxAPIConst.BK_dNACD, '(float) No-ac-decay ratio', 'float'],
    'OPKV': [OlxAPIConst.BK_dOperatingKV, '(float) Operating kV', 'float'],
    'RATEDKV': [OlxAPIConst.BK_dRatedKV, '(float) Max design kV', 'float'],
    'INTRATING': [OlxAPIConst.BK_dRating1, '(float) Interrupting Rating', 'float'],
    'MRATING': [OlxAPIConst.BK_dRating2, '(float) Momentary rating', 'float'],
    'NODERATE': [OlxAPIConst.BK_nDontDerate, '(int) Do not derate in reclosing operation flag: 1-true; 0-false;', 'int'],
    'FLAG': [OlxAPIConst.BK_nInService, '(int) In-service FLAG: 1-true; 2-false;', 'int'],
    'GROUPTYPE1': [OlxAPIConst.BK_nInterrupt1, '(int) Group 1 interrupting current: 1-max current; 0-group current', 'int'],
    'GROUPTYPE2': [OlxAPIConst.BK_nInterrupt2, '(int) Group 1 interrupting current: 1-max current; 0-group current', 'int'],
    'RATINGTYPE': [OlxAPIConst.BK_nRatingType, '(int) Rating type: 0- symmetrical current basis;1- total current basis; 2- IEC', 'int'],
    'OPS1': [OlxAPIConst.BK_nTotalOps1, '(int) Total operations for group 1', 'int'],
    'OPS2': [OlxAPIConst.BK_nTotalOps2, '(int) Total operations for group 2', 'int'],
    'NAME': [OlxAPIConst.BK_sID, '(str) Name (ID)', 'strk'],
    'RECLOSE1': [OlxAPIConst.BK_vdRecloseInt1, '[float]*3 Reclosing intervals for group 1 (s)', 'float3'],
    'RECLOSE2': [OlxAPIConst.BK_vdRecloseInt2, '[float]*3 Reclosing intervals for group 2 (s)', 'float3'],
    'BUS': [OlxAPIConst.BK_nBusHnd, '(BUS) Breaker Bus', 0],
    'OBJLST1': [OlxAPIConst.BK_vnG1DevHnd, '[EQUIPMENT]*10 Protected equipment group 1 List of equipment', 'equip10'],
    'G1OUTAGES': [OlxAPIConst.BK_vnG1OutageHnd, '[EQUIPMENT]*10 Protected equipment group 1 List of additional outage', 'equip10'],
    'OBJLST2': [OlxAPIConst.BK_vnG2DevHnd, '[EQUIPMENT]*10 Protected equipment group 2 List of equipment', 'equip10'],
    'G2OUTAGES': [OlxAPIConst.BK_vnG2OutageHnd, '[EQUIPMENT]*10 Protected equipment group 2 List of additional outage', 'equip10'],
    'PYTHONSTR': [0, '(str) String python of BREAKER', 0],
    'KEYSTR': [0, '(str) Key defined of BREAKER', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of BREAKER', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['MULINE'] = {
    'FROM1': [OlxAPIConst.MU_vdFrom1, '[float]*5 Line1 From percent', 'float5'],
    'FROM2': [OlxAPIConst.MU_vdFrom2, '[float]*5 Line2 From percent', 'float5'],
    'TO1': [OlxAPIConst.MU_vdTo1, '[float]*5 Line1 To percent', 'float5'],
    'TO2': [OlxAPIConst.MU_vdTo2, '[float]*5 Line2 To percent', 'float5'],
    'R': [OlxAPIConst.MU_vdR, '[float]*5 Resistane R, in pu', 'float5'],
    'X': [OlxAPIConst.MU_vdX, '[float]*5 Mutual Reactance X, in pu', 'float5'],
    'LINE1': [OlxAPIConst.MU_nHndLine1, '(LINE) Line 1', 0],
    'LINE2': [OlxAPIConst.MU_nHndLine2, '(LINE) Line 2', 0],
    'ORIENTLINE1': [0, '[BUS1, BUS2, CID] Line orient 1', 'b1b2id'],
    'ORIENTLINE2': [0, '[BUS1, BUS2, CID] Line orient 2', 'b1b2id'],
    'PYTHONSTR': [0, '(str) String python of MULINE', 0],
    'KEYSTR': [0, '(str) Key defined of MULINE', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of MULINE', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'MEMO': [0, '(str) MEMO', 'str'],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYGROUP'] = {
    'INTRPTIME': [OlxAPIConst.RG_dBreakerTime, '(float) Interrupting time (cycles)', 'float'],
    'FLAG': [OlxAPIConst.RG_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 0],
    'OPFLAG': [OlxAPIConst.RG_nOps, '(int) Total operations', 0],
    'NOTE': [OlxAPIConst.RG_sNote, '(str) Annotation', 0],
    'RECLSRTIME': [OlxAPIConst.RG_vdRecloseInt, '[float]*4 Reclosing intervals', 'float4'],
    'BACKUP': [OlxAPIConst.RG_vnBackupGroup, '[RLYGROUP] Back up Group', 'RLYGROUPn'],
    'EQUIPMENT': [OlxAPIConst.RG_nBranchHnd, '(EQUIPMENT) Equipment', 0],
    'PRIMARY': [OlxAPIConst.RG_vnPrimaryGroup, '[RLYGROUP] Primary Group', 'RLYGROUPn'],
    'LOGICRECL': [OlxAPIConst.RG_nReclLogicHnd, '(SCHEME) Reclose logic scheme', 0],
    'LOGICTRIP': [OlxAPIConst.RG_nTripLogicHnd, '(SCHEME) Trip logic scheme', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYGROUP', 0],
    'BUS': [0, '[BUS] List of Buses of RLYGROUP(BUS[0]: Bus Local; BUS[1],(BUS[2]): Bus Opposite(s)', 0],
    'FUSE': [0, '[FUSE] List of Fuses of RLYGROUP', 0],
    'RECLSR': [0, '[RECLSR] List of Reclosers of RLYGROUP', 0],
    'RLYD': [0, '[RLYD] List of Differential Relays of RLYGROUP', 0],
    'RLYDS': [0, '[RLYDSG+RLYDSP] List of Distance Relays (Ground+Phase) of RLYGROUP', 0],
    'RLYDSG': [0, '[RLYDSG] List of Distance Ground Relays of RLYGROUP', 0],
    'RLYDSP': [0, '[RLYDSP] List of Distance Phase Relays of RLYGROUP', 0],
    'RLYOC': [0, '[RLYOCG+RLYOCP] List of Overcurrent Relays (Ground+Phase) of RLYGROUP', 0],
    'RLYOCG': [0, '[RLYOCG] List of Overcurrent Ground Relays of RLYGROUP', 0],
    'RLYOCP': [0, '[RLYOCP] List of Overcurrent Phase Relays of RLYGROUP', 0],
    'RLYV': [0, '[RLYV] List of Voltage Relays of RLYGROUP', 0],
    'SCHEME': [0, '[SCHEME] List of Logic Schemes of RLYGROUP', 0],
    'PYTHONSTR': [0, '(str) String python of Relay Group', 0],
    'KEYSTR': [0, '(str) Key defined of Relay Group', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Relay Group', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYOCG'] = {
    'ID': [OlxAPIConst.OG_sID, '(str) ID', 'strk'],
    'ASSETID': [OlxAPIConst.OG_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.OG_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.OG_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.OG_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'CT': [OlxAPIConst.OG_dCT, '(float) CT ratio', 'float'],
    'CTSTR': [OlxAPIConst.OG_sCT, '(str) String CT ratio', 'str'],
    'CTLOC': [OlxAPIConst.OG_nCTLocation, '(int) CT Location', 'int'],
    'OPI': [OlxAPIConst.OG_nOperateOn, '(int) Operate On: 0-3I0; 1-3I2; 2-I0; 3-I2', 'int'],
    'ASYM': [OlxAPIConst.OG_nDCOffset, '(int) Sensitive to DC offset: 1-true; 0-false', 'int'],
    'FLATINST': [OlxAPIConst.OG_nFlatDelay, '(int) Flat definite time delay FLAG: 1-true; 0-false', 'int'],
    'SGNL': [OlxAPIConst.OG_nSignalOnly, '(int) Signal only: Example 1-INST; 2-OC; 4-DT1;...', 'int'],
    'OCDIR': [OlxAPIConst.OG_nDirectional, '(int) Directional FLAG: 0-None; 1-Fwd.; 2-Rev.', 'int'],
    'DTDIR': [OlxAPIConst.OG_nIDirectional, '(int) Inst. Directional FLAG: 0=None; 1-Fwd.; 2-Rev.', 'int'],
    'POLAR': [OlxAPIConst.OG_nPolar, "(int) Polar option 0-'Vo,Io polarized',1:'V2,I2 polarized',2:'SEL V2 polarized',3:'SEL V2 polarized'", 'int'],
    'PICKUPTAP': [OlxAPIConst.OG_dTap, '(float) Pickup (A)', 'float'],
    'TDIAL': [OlxAPIConst.OG_dTDial, '(float) Time dial', 'float'],
    'TIMEADD': [OlxAPIConst.OG_dTimeAdd, '(float) Time adder', 'float'],
    'TIMEMULT': [OlxAPIConst.OG_dTimeMult, '(float) Time multiplier', 'float'],
    'TIMERESET': [OlxAPIConst.OG_dResetTime, '(float) Reset time', 'float'],
    'DTTIMEADD': [OlxAPIConst.OG_dTimeAdd2, '(float) Time adder for INST/DTx', 'float'],
    'DTTIMEMULT': [OlxAPIConst.OG_dTimeMult2, '(float) Time multiplier for INST/DTx', 'float'],
    'INSTSETTING': [OlxAPIConst.OG_dInst, '(float) Instantaneous setting', 'float'],
    'DTPICKUP': [OlxAPIConst.OG_vdDTPickup, '[float]*5 Pickups Sec.A ', 'float5'],
    'DTDELAY': [OlxAPIConst.OG_vdDTDelay, '[float]*5 Delays seconds', 'float5'],
    'MINTIME': [OlxAPIConst.OG_dMinTripTime, '(float) Minimum trip time', 'float'],
    'TYPE': [OlxAPIConst.OG_sType, '(str) Type', 'str'],
    'TAPTYPE': [OlxAPIConst.OG_sTapType, '(str) Tap type', 'str'],
    'LIBNAME': [OlxAPIConst.OG_sLibrary, '(str) Library', 'str'],
    'PACKAGE': [OlxAPIConst.OG_nPackage, '(int) Package option', 'int'],
    'RLYGROUP': [OlxAPIConst.OG_nRlyGrHnd, '(RLYGROUP) Relay Group of RLYOCG', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYOCG located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYOCG', 0],
    'BUS': [0, '[BUS] List of Buses of RLYOCG', 0],
    'PYTHONSTR': [0, '(str) String python of RLYOCG', 0],
    'KEYSTR': [0, '(str) Key defined of RLYOCG', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYOCG', 0],
    'SETTINGSTR': [0, '(str) Settings in JSON Form of RLYOCG', 0],
    'SETTINGNAME': [0, '[str] List of Settings name', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYOCP'] = {
    'ID': [OlxAPIConst.OP_sID, '(str) ID', 'strk'],
    'ASSETID': [OlxAPIConst.OP_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.OP_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.OP_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.OP_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'CT': [OlxAPIConst.OP_dCT, '(float) CT ratio', 'float'],
    'CTSTR': [OlxAPIConst.OP_sCT, '(str) String CT ratio', 'str'],
    'CTCONNECT': [OlxAPIConst.OP_nByCTConnect, '(int) CT connection: 0- Wye; 1-Delta', 'int'],
    'ASYM': [OlxAPIConst.OP_nDCOffset, '(int) Sensitive to DC offset: 1-true; 0-false', 'int'],
    'FLATINST': [OlxAPIConst.OP_nFlatDelay, '(int) Flat delay: 1-true; 0-false', 'int'],
    'SGNL': [OlxAPIConst.OP_nSignalOnly, '(int) Signal only: 1-INST;2-OC; 4-DT1;...', 'int'],
    'OCDIR': [OlxAPIConst.OP_nDirectional, '(int) Directional FLAG: 0=None;1=Fwd.;2=Rev.', 'int'],
    'DTDIR': [OlxAPIConst.OP_nIDirectional, '(int) Inst. Directional FLAG: 0=None;1=Fwd.;2=Rev.', 'int'],
    'POLAR': [OlxAPIConst.OP_nPolar, "(int) Polar option 0:'Cross-V polarized'; 2:'SEL V2 polarized'", 'int'],
    'VOLTCONTROL': [OlxAPIConst.OP_nVoltControl, '(int) Voltage controlled or restrained', 'int'],
    'PICKUPTAP': [OlxAPIConst.OP_dTap, '(float) Tap Ampere', 'float'],
    'TDIAL': [OlxAPIConst.OP_dTDial, '(float) Time dial', 'float'],
    'TIMEADD': [OlxAPIConst.OP_dTimeAdd, '(float) Time adder', 'float'],
    'TIMEMULT': [OlxAPIConst.OP_dTimeMult, '(float) Time multiplier', 'float'],
    'TIMERESET': [OlxAPIConst.OP_dResetTime, '(float) Reset time', 'float'],
    'VOLTPERCENT': [OlxAPIConst.OP_dVCtrlRestPcnt, '(float) Voltage controlled or restrained percentage', 'float'],
    'DTTIMEADD': [OlxAPIConst.OP_dTimeAdd2, '(float) Time adder for INST/DTx', 'float'],
    'DTTIMEMULT': [OlxAPIConst.OP_dTimeMult2, '(float) Time multiplier for INST/DTx', 'float'],
    'INSTSETTING': [OlxAPIConst.OP_dInst, '(float) Instantaneous setting', 'float'],
    'DTPICKUP': [OlxAPIConst.OP_vdDTPickup, '[float]*5 Pickup', 'float5'],
    'DTDELAY': [OlxAPIConst.OP_vdDTDelay, '[float]*5 Delay', 'float5'],
    'MINTIME': [OlxAPIConst.OP_dMinTripTime, '(float) Minimum trip time', 'float'],
    'TYPE': [OlxAPIConst.OP_sType, '(str) Type', 'str'],
    'TAPTYPE': [OlxAPIConst.OP_sTapType, '(str) Tap type', 'str'],
    'LIBNAME': [OlxAPIConst.OP_sLibrary, '(str) Library', 'str'],
    'PACKAGE': [OlxAPIConst.OP_nPackage, '(int) Package option', 'int'],
    'RLYGROUP': [OlxAPIConst.OP_nRlyGrHnd, '(RLYGROUP) Relay Group of RLYOCP', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYOCP located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYOCP', 0],
    'BUS': [0, '[BUS] List of Buses of RLYOCP', 0],
    'PYTHONSTR': [0, '(str) String python of RLYOCP', 0],
    'KEYSTR': [0, '(str) Key defined of RLYOCP', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYOCP', 0],
    'SETTINGSTR': [0, '(str) Settings in JSON Form of RLYOCP', 0],
    'SETTINGNAME': [0, '[str] List of Settings name', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['FUSE'] = {
    'ID': [OlxAPIConst.FS_sID, '(str) Name(ID)', 'strk'],
    'ASSETID': [OlxAPIConst.FS_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.FS_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.FS_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.FS_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'FUSECURDIV': [OlxAPIConst.FS_dCT, '(float) Current divider', 'float'],
    'RATING': [OlxAPIConst.FS_dRating, '(float) Rating', 'float'],
    'TIMEMULT': [OlxAPIConst.FS_dTimeMult, '(float) Time multiplier', 'float'],
    'FUSECVE': [OlxAPIConst.FS_nCurve, '(int) Compute time using FLAG: 1- minimum melt; 2- Total clear', 'int'],
    'PACKAGE': [OlxAPIConst.FS_nPackage, '(int) Package option', 'int'],
    'LIBNAME': [OlxAPIConst.FS_sLibrary, '(str) Library', 'str'],
    'TYPE': [OlxAPIConst.FS_sType, '(str) Type', 'str'],
    'RLYGROUP': [OlxAPIConst.FS_nRlyGrHnd, '(RLYGROUP) Relay Group of Fuse', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that FUSE located on', 0],
    'BUS1': [0, '(BUS) Bus Local of Fuse', 0],
    'BUS': [0, '[BUS] List of Buses of Fuse', 0],
    'PYTHONSTR': [0, '(str) String python of Fuse', 0],
    'KEYSTR': [0, '(str) Key defined of Fuse', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of Fuse', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYDSG'] = {
    'ID': [OlxAPIConst.DG_sID, '(str) Name(ID)', 'strk'],
    'TYPE': [OlxAPIConst.DG_sType, '(str) Type (ID2)', 'str'],
    'ASSETID': [OlxAPIConst.DG_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.DG_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.DG_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.DG_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'RLYGROUP': [OlxAPIConst.DG_nRlyGrHnd, '(RLYGROUP) Relay group of RLYDSG', 0],
    'DSTYPE': [OlxAPIConst.DG_sDSType, '(str) Type name (7SA511,SEL411G,..)', 'str'],
    'Z2OCTYPE': [OlxAPIConst.DG_sZ2OCCurve, '(str) Zone 2 OC supervision type name', 'str'],
    'VTBUS': [OlxAPIConst.DG_nVTBus, '(BUS) VT at Bus', 'BUS'],
    'Z2OCPICKUP': [OlxAPIConst.DG_dZ2OCPickup, '(float) Z2 OC supervision Pickup(A)', 'float'],
    'Z2OCTD': [OlxAPIConst.DG_dZ2OCTD, '(float) Z2 OC supervision time dial', 'float'],
    'STARTZ2FLAG': [OlxAPIConst.DG_nStartZ2Flag, '(int) Start Z2 timer on forward Z3 or Z4 pickup FLAG', 'int'],
    'SNLZONE': [OlxAPIConst.DG_nSignalOnly, '(int) Signal-only zone', 'int'],
    'LIBNAME': [OlxAPIConst.DG_sLibrary, '(str) Library', 'str'],
    'PACKAGE': [OlxAPIConst.DG_nPackage, '(int) Package option', 'int'],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYDSG located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYDSG', 0],
    'BUS': [0, '[BUS] List of Buses of RLYDSG', 0],
    'PYTHONSTR': [0, '(str) String python of DS ground relay', 0],
    'KEYSTR': [0, '(str) Key defined of DS ground relay', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYDSG', 0],
    'SETTINGSTR': [0, '(str) Settings in JSON Form of RLYDSG', 0],
    'SETTINGNAME': [0, '[str] List of Settings name', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) DS ground relay MEMO', 'str'],
    'TAGS': [0, '(str) DS ground relay TAGS', 'str']}


__OLXOBJ_PARA__['RLYDSP'] = {
    'ID': [OlxAPIConst.DP_sID, '(str) ID', 'strk'],
    'TYPE': [OlxAPIConst.DP_sType, '(str) Type (ID2)', 'str'],
    'ASSETID': [OlxAPIConst.DP_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.DP_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.DP_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.DP_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'VTBUS': [OlxAPIConst.DP_nVTBus, '(BUS) VT at Bus', 'BUS'],
    'DSTYPE': [OlxAPIConst.DP_sDSType, '(str) Type name (7SA511,SEL411G,..)', 'str'],
    'Z2OCTYPE': [OlxAPIConst.DP_sZ2OCCurve, '(str) Zone 2 OC supervision type name', 'str'],
    'Z2OCPICKUP': [OlxAPIConst.DP_dZ2OCPickup, '(float) Z2 OC supervision Pickup(A)', 'float'],
    'Z2OCTD': [OlxAPIConst.DP_dZ2OCTD, '(float) Z2 OC supervision time dial', 'float'],
    'STARTZ2FLAG': [OlxAPIConst.DP_nStartZ2Flag, '(int) Start Z2 timer on forward Z3 or Z4 pickup FLAG', 'int'],
    'PACKAGE': [OlxAPIConst.DP_nPackage, '(int) Package option', 'int'],
    'SNLZONE': [OlxAPIConst.DP_nSignalOnly, '(int) Signal-only zone FLAG', 'int'],
    'LIBNAME': [OlxAPIConst.DP_sLibrary, '(str) Library', 'str'],
    'RLYGROUP': [OlxAPIConst.DP_nRlyGrHnd, '(RLYGROUP) Relay Group of RLYDSP', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYDSP located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYDSP', 0],
    'BUS': [0, '[BUS] List of Buses of RLYDSP', 0],
    'PYTHONSTR': [0, '(str) String python of RLYDSP', 0],
    'KEYSTR': [0, '(str) Key defined of RLYDSP', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYDSP', 0],
    'SETTINGSTR': [0, '(str) Settings in JSON Form of RLYDSP', 0],
    'SETTINGNAME': [0, '[str] List of Settings name', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYD'] = {
    'ID': [OlxAPIConst.RD_sID, '(str) ID (NAME)', 'strk'],
    'ASSETID': [OlxAPIConst.RD_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.RD_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.RD_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.RD_nInService, '(int) In service FLAG: 1- active; 2- out-of-service', 'int'],
    'PACKAGE': [OlxAPIConst.RD_nInService, '(int) Package option', 'int'],
    'RLYGROUP': [OlxAPIConst.RD_nRlyGrpHnd, '(RLYGROUP) Relay Group of RLYD', 0],
    'CTR1': [OlxAPIConst.RD_dCTR1, '(float) CTR1', 'float'],
    'IMIN3I0': [OlxAPIConst.RD_dPickup3I0, '(float) Minimum enable differential current (3I0)', 'float'],
    'IMIN3I2': [OlxAPIConst.RD_dPickup3I2, '(float) Minimum enable differential current (3I2)', 'float'],
    'IMINPH': [OlxAPIConst.RD_dPickupPh, '(float) Minimum enable differential current (phase)', 'float'],
    'TLCTD3I0': [OlxAPIConst.RD_dTLCTDDelayI0, '(float) Tapped load coordination delay (I0)', 'float'],
    'TLCTDI2': [OlxAPIConst.RD_dTLCTDDelayI2, '(float) Tapped load coordination delay (I2)', 'float'],
    'TLCTDPH': [OlxAPIConst.RD_dTLCTDDelayPh, '(float) Tapped load coordination delay (phase)', 'float'],
    'SGLONLY': [OlxAPIConst.RD_nSignalOnly, '(int) Signal only: 1-true; 2-false', 'int'],
    'TLCCV3I0': [OlxAPIConst.RD_sTLCCurveI0, '(str) Tapped load coordination curve (I0)', 0],
    'TLCCVI2': [OlxAPIConst.RD_sTLCCurveI2, '(str) Tapped load coordination curve (I2)', 0],
    'TLCCVPH': [OlxAPIConst.RD_sTLCCurvePh, '(str) Tapped load coordination curve (phase)', 0],
    'CTGRP1': [OlxAPIConst.RD_nLocalCTHnd1, '(RLYGROUP) Local current input 1', 'RLYGROUP'],
    'RMTE1': [OlxAPIConst.RD_nRmeDevHnd1, '(RLYD) Remote device 1', 'RLYD'],
    'RMTE2': [OlxAPIConst.RD_nRmeDevHnd2, '(RLYD) Remote device 2', 'RLYD'],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYD located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYD', 0],
    'BUS': [0, '[BUS] List of Buses of RLYD', 0],
    'PYTHONSTR': [0, '(str) String python of RLYD', 0],
    'KEYSTR': [0, '(str) Key defined of RLYD', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYD', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RLYV'] = {
    'ID': [OlxAPIConst.RV_sID, '(str) ID (NAME)', 'strk'],
    'ASSETID': [OlxAPIConst.RV_sAssetID, '(str) Asset ID', 'str'],
    'DATEOFF': [OlxAPIConst.RV_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.RV_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.RV_nInService, '(int) In service FLAG: 1- active; 2- out-of-service', 'int'],
    'PACKAGE': [OlxAPIConst.RV_nInService, '(int) Package option', 'int'],
    'PTR': [OlxAPIConst.RV_dCTR, '(float) PT ratio', 'float'],
    'OVINST': [OlxAPIConst.RV_dOVIPickup, '(float) Over-voltage instant pickup (V)', 'float'],
    'UVINST': [OlxAPIConst.RV_dUVIPickup, '(float) Under-voltage instant pickup (V)', 'float'],
    'OVPICKUP': [OlxAPIConst.RV_dOVTPickup, '(float) Over-voltage pickup (V)', 'float'],
    'UVPICKUP': [OlxAPIConst.RV_dUVTPickup, '(float) Under-voltage pickup (V)', 'float'],
    'UVDELAYTD': [OlxAPIConst.RV_dUVTDelay, '(float) Under-voltage delay', 'float'],
    'OVDELAYTD': [OlxAPIConst.RV_dOVTDelay, '(float) Over-voltage delay', 'float'],
    'SGLONLY': [OlxAPIConst.RV_nSignalOnly, '[int 0/1]*4 Signal only: 0-No check; 1-Check for [Over-voltage Instant, Over-voltage Delay, Under-voltage Instant, Under-voltage Delay]', 'int401'],
    'OPQTY': [OlxAPIConst.RV_nVoltOperate, '(int) Operate on voltage option: 1-Phase-to-Neutral; 2- Phase-to-Phase; 3-3V0;4-V1;5-V2;6-VA;7-VB;8-VC;9-VBC;10-VAB;11-VCA', 'int'],
    'OVCVR': [OlxAPIConst.RV_sOVCurve, '(str) Over-voltage element curve', 0],
    'UVCVR': [OlxAPIConst.RV_sUVCurve, '(str) Under-voltage element curve', 0],
    'RLYGROUP': [OlxAPIConst.RV_nRlyGrpHnd, '(RLYGROUP) Relay Group of RLYV', 0],
    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RLYV located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RLYV', 0],
    'BUS': [0, '[BUS] List of Buses of RLYV', 0],
    'PYTHONSTR': [0, '(str) String python of RLYV', 0],
    'KEYSTR': [0, '(str) Key defined of RLYV', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RLYV', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['RECLSR'] = {
    'ID': [OlxAPIConst.CP_sID, '(str) ID', 'strk'],
    'ASSETID': [OlxAPIConst.CP_sAssetID, '(str) AssetID', 'str'],
    'DATEOFF': [OlxAPIConst.CP_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.CP_sOnDate, '(str) In service date', 'str'],
    'FLAG': [OlxAPIConst.CP_nInService, '(int) In service FLAG: 1- active; 2- out-of-service', 'int'],
    'RLYGROUP': [OlxAPIConst.CP_nRlyGrHnd, '(RLYGROUP) Relay Group of RECLSR', 0],
    'TOTALOPS': [OlxAPIConst.CP_nTotalOps, '(int) Total operations to locked out', 'int'],
    'FASTOPS': [OlxAPIConst.CP_nFastOps, '(int) Number of fast operations', 'int'],
    'RECLOSE1': [OlxAPIConst.CP_dRecIntvl1, '(float) Reclosing interval 1', 'float'],
    'RECLOSE2': [OlxAPIConst.CP_dRecIntvl2, '(float) Reclosing interval 2', 'float'],
    'RECLOSE3': [OlxAPIConst.CP_dRecIntvl3, '(float) Reclosing interval 3', 'float'],
    'INTRPTIME': [OlxAPIConst.CP_dIntrTime, '(float) Interrupting time(s)', 'float'],
    'BYTADD': [OlxAPIConst.CP_nTAddAppl, '(int) Time adder modifies', 'int'],
    'BYTMULT': [OlxAPIConst.CP_nTMultAppl, '(int) Time multiplier modifies', 'int'],
    'LIBNAME': [OlxAPIConst.CP_sLibrary, '(str) Library', 'str'],
    'RATING': [OlxAPIConst.CP_dRating, '(float) Rating (Rated momentary amps)', 'float'],

    'PH_BYFAST': [OlxAPIConst.CP_nCurveInUse, '(int) Recloser-Phase Curve in use FLAG: 1- Fast curve; 0- Slow curve', 'int'],
    'PH_FLAG': [OlxAPIConst.CP_nInService, '(int) Recloser-Phase In service FLAG: 1- active; 2- out-of-service', 'int'],
    'PH_INST': [OlxAPIConst.CP_dHiAmps, '(float) Recloser-Phase high current trip', 'float'],
    'PH_INSTDELAY': [OlxAPIConst.CP_dHiAmpsDelay, '(float) Recloser-Phase high current trip delay', 'float'],
    'PH_MINTIMEF': [OlxAPIConst.CP_dMinTF, '(float) Recloser-Phase fast curve minimum time', 'float'],
    'PH_MINTIMES': [OlxAPIConst.CP_dMinTS, '(float) Recloser-Phase slow curve minimum time', 'float'],
    'PH_MINTRIPF': [OlxAPIConst.CP_dPickupF, '(float) Recloser-Phase fast curve pickup', 'float'],
    'PH_MINTRIPS': [OlxAPIConst.CP_dPickupS, '(float) Recloser-Phase slow curve pickup', 'float'],
    'PH_TIMEADDF': [OlxAPIConst.CP_dTimeAddF, '(float) Recloser-Phase fast curve time adder', 'float'],
    'PH_TIMEADDS': [OlxAPIConst.CP_dTimeAddS, '(float) Recloser-Phase slow curve time adder', 'float'],
    'PH_TIMEMULTF': [OlxAPIConst.CP_dTimeMultF, '(float) Recloser-Phase fast curve time multiplier', 'float'],
    'PH_TIMEMULTS': [OlxAPIConst.CP_dTimeMultS, '(float) Recloser-Phase slow curve time multiplier', 'float'],
    'PH_FASTTYPE': [OlxAPIConst.CP_sTypeFast, '(str) Recloser-Phase fast curve', 'str'],
    'PH_SLOWTYPE': [OlxAPIConst.CP_sTypeSlow, '(str) Recloser-Phase slow curve', 'str'],

    'GR_BYFAST': [OlxAPIConst.CP_nCurveInUse, '(int) Recloser-Ground Curve in use FLAG: 1- Fast curve; 0- Slow curve', 'int'],
    'GR_FLAG': [OlxAPIConst.CG_nInService, '(int) Recloser-Ground In service FLAG: 1- active; 2- out-of-service', 'int'],
    'GR_INST': [OlxAPIConst.CG_dHiAmps, '(float) Recloser-Ground high current trip', 'float'],
    'GR_INSTDELAY': [OlxAPIConst.CG_dHiAmpsDelay, '(float) Recloser-Ground high current trip delay', 'float'],
    'GR_MINTIMEF': [OlxAPIConst.CG_dMinTF, '(float) Recloser-Ground fast curve minimum time', 'float'],
    'GR_MINTIMES': [OlxAPIConst.CG_dMinTS, '(float) Recloser-Ground slow curve minimum time', 'float'],
    'GR_MINTRIPF': [OlxAPIConst.CG_dPickupF, '(float) Recloser-Ground fast curve pickup', 'float'],
    'GR_MINTRIPS': [OlxAPIConst.CG_dPickupS, '(float) Recloser-Ground slow curve pickup', 'float'],
    'GR_TIMEADDF': [OlxAPIConst.CG_dTimeAddF, '(float) Recloser-Ground fast curve time adder', 'float'],
    'GR_TIMEADDS': [OlxAPIConst.CG_dTimeAddS, '(float) Recloser-Ground slow curve time adder', 'float'],
    'GR_TIMEMULTF': [OlxAPIConst.CG_dTimeMultF, '(float) Recloser-Ground fast curve time multiplier', 'float'],
    'GR_TIMEMULTS': [OlxAPIConst.CG_dTimeMultS, '(float) Recloser-Ground slow curve time multiplier', 'float'],
    'GR_FASTTYPE': [OlxAPIConst.CG_sTypeFast, '(str) Recloser-Ground fast curve', 'str'],
    'GR_SLOWTYPE': [OlxAPIConst.CG_sTypeSlow, '(str) Recloser-Ground slow curve', 'str'],

    'EQUIPMENT': [0, '(EQUIPMENT) EQUIPMENT that RECLSR located on', 0],
    'BUS1': [0, '(BUS) Bus Local of RECLSR', 0],
    'BUS': [0, '[BUS] List of Buses of RECLSR', 0],
    'PYTHONSTR': [0, '(str) String python of RECLSR', 0],
    'KEYSTR': [0, '(str) Key defined of RECLSR', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of RECLSR', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'PH_MEMO': [0, '(str) Recloser-Phase MEMO', 'str'],
    'GR_MEMO': [0, '(str) Recloser-Ground MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_PARA__['SCHEME'] = {
    'DATEOFF': [OlxAPIConst.LS_sOffDate, '(str) Out of service date', 'str'],
    'DATEON': [OlxAPIConst.LS_sOnDate, '(str) In service date', 'str'],
    'ID': [OlxAPIConst.LS_sID, '(str) ID', 'strk'],
    'ASSETID': [OlxAPIConst.LS_sAssetID, '(str) Asset ID', 'str'],
    'TYPE': [OlxAPIConst.LS_sScheme, '(str) Type (POTT,PUTT..)', 'str'],
    'EQUATION': [OlxAPIConst.LS_sEquation, '(str) Equation', 'str'],
    'FLAG': [OlxAPIConst.LS_nInService, '(int) In-service FLAG: 1- active; 2- out-of-service', 'int'],
    'SIGNALONLY': [OlxAPIConst.LS_nSignalOnly, '(int) Signal only', 'int'],
    'BUS1': [0, '(BUS) Bus Local of SCHEME', 0],
    'BUS': [0, '[BUS] List of Buses of SCHEME', 0],
    'RLYGROUP': [OlxAPIConst.LS_nRlyGrpHnd, '(RLYGROUP) Relay Group of SCHEME', 0],
    'PYTHONSTR': [0, '(str) String python of SCHEME', 0],
    'KEYSTR': [0, '(str) Key defined of SCHEME', 0],
    'PARAMSTR': [0, '(str) Parameters in JSON Form of SCHEME', 0],
    'LOGICSTR': [0, '(str) Logic Var in JSON Form of SCHEME', 0],
    'LOGICVARNAME': [OlxAPIConst.LS_vsSignalName, '[str] List of Logic Var name', 0],
    'JRNL': [0, '(str) Creation and modification log records', 0],
    'GUID': [0, '(str) Globally Unique IDentifier', 'str'],
    'MEMO': [0, '(str) MEMO', 'str'],
    'TAGS': [0, '(str) TAGS', 'str']}


__OLXOBJ_SCHEME__ = {'INST/DT PICKUP': OlxAPIConst.LVT_OC_DT,
                     'INST/DT TRIP': OlxAPIConst.LVT_OC_DTTRIP,
                     'DT1 PICKUP': OlxAPIConst.LVT_OC_DT1,
                     'DT2 PICKUP': OlxAPIConst.LVT_OC_DT2,
                     'DT3 PICKUP': OlxAPIConst.LVT_OC_DT3,
                     'DT4 PICKUP': OlxAPIConst.LVT_OC_DT4,
                     'DT5 PICKUP': OlxAPIConst.LVT_OC_DT5,
                     'DT1 TRIP': OlxAPIConst.LVT_OC_DTTRIP1,
                     'DT2 TRIP': OlxAPIConst.LVT_OC_DTTRIP2,
                     'DT3 TRIP': OlxAPIConst.LVT_OC_DTTRIP3,
                     'DT4 TRIP': OlxAPIConst.LVT_OC_DTTRIP4,
                     'DT5 TRIP': OlxAPIConst.LVT_OC_DTTRIP5,
                     'TOC PICKUP': OlxAPIConst.LVT_OC_TOC,
                     'TOC TRIP': OlxAPIConst.LVT_OC_TOCTRIP,

                     'ZONE 1 PICKUP': OlxAPIConst.LVT_DS_ZONE1,
                     'ZONE 2 PICKUP': OlxAPIConst.LVT_DS_ZONE2,
                     'ZONE 3 PICKUP': OlxAPIConst.LVT_DS_ZONE3,
                     'ZONE 4 PICKUP': OlxAPIConst.LVT_DS_ZONE4,
                     'ZONE 5 PICKUP': OlxAPIConst.LVT_DS_ZONE5,
                     'ZONE 6 PICKUP': OlxAPIConst.LVT_DS_ZONE6,
                     'ZONE 1 TRIP': OlxAPIConst.LVT_DS_ZONETRIP1,
                     'ZONE 2 TRIP': OlxAPIConst.LVT_DS_ZONETRIP2,
                     'ZONE 3 TRIP': OlxAPIConst.LVT_DS_ZONETRIP3,
                     'ZONE 4 TRIP': OlxAPIConst.LVT_DS_ZONETRIP4,
                     'ZONE 5 TRIP': OlxAPIConst.LVT_DS_ZONETRIP5,
                     'ZONE 6 TRIP': OlxAPIConst.LVT_DS_ZONETRIP6,

                     'OV PICKUP': OlxAPIConst.LVT_OVT,
                     'OV INST.': OlxAPIConst.LVT_OVI,
                     'OV TRIP': OlxAPIConst.LVT_OVTTRIP,
                     'OV INST.TRIP': OlxAPIConst.LVT_OVITRIP,
                     'UV PICKUP': OlxAPIConst.LVT_UVT,
                     'UV INST.': OlxAPIConst.LVT_UVI,
                     'UV TRIP': OlxAPIConst.LVT_UVTTRIP,
                     'UV INST. TRIP': OlxAPIConst.LVT_UVITRIP,

                     'PICKUP': OlxAPIConst.LVT_DIFF,
                     'PICKUP PH': OlxAPIConst.LVT_DIFFPH,
                     'PICKUP 3I0': OlxAPIConst.LVT_DIFF3I0,
                     'PICKUP 3I2': OlxAPIConst.LVT_DIFF3I2,
                     'TRIP': OlxAPIConst.LVT_DIFFTRIP,

                     'OPEN OP.': OlxAPIConst.LVT_R_OP,
                     'OPEN OP. #1': OlxAPIConst.LVT_R_OP1,
                     'OPEN OP. #2': OlxAPIConst.LVT_R_OP2,
                     'OPEN OP. #3': OlxAPIConst.LVT_R_OP3,
                     'OPEN & LOCK': OlxAPIConst.LVT_R_OPL,
                     'RECLOSE OP.': OlxAPIConst.LVT_R_C,
                     'RECLOSE OP. #1': OlxAPIConst.LVT_R_C1,
                     'RECLOSE OP. #2': OlxAPIConst.LVT_R_C2,
                     'RECLOSE OP. #3': OlxAPIConst.LVT_R_C3,

                     'SCHEME': OlxAPIConst.LVT_SCHEME}
__OLXOBJ_SCHEME1__ = {v:k for k,v in __OLXOBJ_SCHEME__.items()}
__OLXOBJ_PARA__['ZCORRECT'] = {}
__OLXOBJ_PARA__['AREA'] = {}
__OLXOBJ_PARA__['ZONE'] = {}


__OLXOBJ_PARASYS__ = {'windowCenterX': ['(int)', 'Graphic window Center X', ''],
                      'windowCenterY': ['(int)', 'Graphic window Center Y', ''],

                      'nPrefaultV': [[0, 1, 2], 'Prefault voltage', "'File/Preferences.../Fault Simulation'"],
                      'fFlatBusV': ['(float)', 'V (pu)', ''],
                      'bIgnoreLoad': [[0, 1], 'Ignore in Short Circuits: Loads', ''],
                      'bIgnoreLineGB': [[0, 1], 'Ignore in Short Circuits: Transmission lin G+jB', ''],
                      'bIgnoreShunt': [[0, 1], 'Ignore in Short Circuits: Shunts with + seq values', ''],
                      'bIgnoreXfmrLineShunt': [[0, 1], 'Ignore in Short Circuits: Transformer line shunts', ''],
                      'nGenZType': [[0, 1, 2], 'Generator Impedance', ''],
                      'nIgnoreMOV': [[0, 1], 'MOV-Protected Series Capacitors : Iterate short circuit solutions', ''],
                      'fAccelFactor': ['(float)', 'Acceleration factor', ''],
                      'nFaultMVAstyle': [[0, 1], 'Define Fault MVA As Product of', ''],
                      'nGenILimitOption': [[0, 1, 2], 'Current Limited Generators', ''],
                      'dMuThreshold': ['(float)', 'Ignore Mutuals < This Threshold', ''],
                      'bSimulateCCGen': [[0, 1], 'Simulate voltage-controlled current sources', ''],
                      'bSimulateGenW4': [[0, 1], 'Simulate converter-interfaced resources', ''],
                      'bSimulateGenW3': [[0, 1], 'Simulate type-3 wind plants', ''],

                      'bSimulateANSIrx': [[0, 1], 'Compute ANSI x/r ratio', "'File/Preferences.../X/R'"],
                      'bUseZ2XR': [[0, 1], 'Assume Z2 equals Z1 for ANSI x/r calculatiom', ''],
                      'fSmallX': ['(float)', 'X-only calculation: If X is 0 use X=', ''],
                      'nSmallR': [[0, 1, 2], 'R-only calculation: If R is 0 use', ''],
                      'fSmallR': ['(float)', 'R-only calculation: Where Rc=', ''],
                      'fXR1': ['(float)', 'Typycal X/R ratio(g) for generators', ''],
                      'fXR2': ['(float)', 'Typycal X/R ratio(g) for transformers', ''],
                      'fXR3': ['(float)', 'Typycal X/R ratio(g) for reactors', ''],
                      'fXR4': ['(float)', 'Typycal X/R ratio(g) for all others', ''],

                      'baseMVA': ['(float)', 'System base MVA', "'File/Preferences.../Network'"],
                      'fHzBase': ['(float)', 'System frequency', ''],
                      'fSwitchR': ['(float)', 'Switch impedance R', ''],
                      'fSwitchX': ['(float)', 'Switch impedance X', ''],
                      'bIgnoreShift': [[0, 1], 'Ignore phase shift of transformer and phase shifters', ''],
                      'nCurRatingL': [[0, 1], 'Line rating are in unit of', ''],
                      'EquivalentID': ['(str)', 'Boundary equivalent equipment ID prefix', ''],
                      'Delimiter': ['(str)', 'String delimiter in text data file', ''],

                      'nIterMax': ['(int)', 'Max iterations', 'PFlow/Solve Power Flow'],
                      'fPTolerance': ['(float)', 'MW Tolerance', ''],
                      'fQTolerance': ['(float)', 'MVAR Tolerance', ''],
                      'fAdjustP': ['(float)', 'MW Auto Adjustment Threshold', ''],
                      'fAdjustQ': ['(float)', 'MVAR Auto Adjustment Threshold', ''],
                      'nSlack': ['(int)', 'System slack bus', ''],

                      'wPFoptions': [],
                      'nCurRatingX': [],
                      'nVoltStyle': [],
                      'nDataCompat': [],
                      'nReferenceBus': [],
                      'wCheckOnFileOpen': [],
                      'nEquivColor': [],
                      'nUseANSIxrIEC': [],
                      'nScaleI': [],
                      'nScaleIEC': [],
                      'wBkrFaults': [],
                      'uStudyDate': [],
                      'nPFMethod': [],
                      'nTagColor': [],
                      'bIgnoreInst': [],
                      'bRlyPickup1': [],
                      'bUseSmallB': [],
                      'bHorzText': [],
                      'bDoAC': [],
                      'bIgnore121kV': [],
                      'bIgnore121kVt': [],
                      'bIgnoreReclosing': [],
                      'bUseFlatPU': [],
                      'b1LG115Total': [],
                      'bShowAllFaults': [],
                      'b1LG115Sym': [],
                      'bIgnoreOOS': [],
                      'bUseANSIxr': [],

                      'bSEA_NoReverseZone': [],
                      'bSEA_NoPkgA': [],
                      'bSEA_NoPkgB': [],
                      'bSEA_NoReclosing': [],
                      'bSEA_TransientFault': [],
                      'b1LLayoutLocked': [],
                      'bRequireBusNo': [],
                      'bDoNotSaveJournal': [],
                      'cnst1': [],
                      'fOlrZoomFactor': [],
                      'fMinCTI': [],
                      'fMaxCTI': [],
                      'fCurPerc': [],
                      'fMinV': [],
                      'fMaxV': [],
                      'fBusVoltMult': [],
                      'fPickupCycles': [],
                      'fKVcutoff': [],
                      'fKVcutoffT': [],
                      'dFc1': [],
                      'dFc2': [],
                      'ColorTag': [],

                      'CaseGUID': ['(str)', 'CaseGUID', ''],
                      'BaseVersionGUID': ['(str)', 'BaseVersionGUID', ''],
                      'FILECOMMENTS': ['(str)', 'File comments', "'File/Info/File Comments...'"]}


__OLXOBJ_oHND__ = {'BUS', 'BUS1', 'BUS2', 'BUS3', 'CNTBUS', 'LTCCTRL', 'VTBUS', 'RLYGROUP', 'RLYGROUP1', 'RLYGROUP2', 'RLYGROUP3',
                   'CTGRP1', 'LOAD', 'SHUNT', 'GEN', 'LINE1', 'LINE2', 'EQUIPMENT', 'MULINE', 'PRIMARY', 'BACKUP', 'LOGICTRIP',
                   'LOGICRECL', 'OBJLST1', 'OBJLST2', 'G1DEVS', 'G2DEVS', 'G1OUTAGES', 'G2OUTAGES', 'RMTE1', 'RMTE2'}

__OLXOBJ_fltConn__ = ['3LG', '2LG:BC', '2LG:CA', '2LG:AB', '2LG:CB', '2LG:AC', '2LG:BA',
                      '1LG:A', '1LG:B', '1LG:C', 'LL:BC', 'LL:CA', 'LL:AB', 'LL:CB', 'LL:AC', 'LL:BA']

__OLXOBJ_fltConnSEA__ = {'3LG': 1,
                         '2LG:BC': 2, '2LG:CA': 3, '2LG:AB': 4, '2LG:CB': 2, '2LG:AC': 3, '2LG:BA': 4,
                         '1LG:A': 5, '1LG:B': 6, '1LG:C': 7,
                         'LL:BC': 8, 'LL:CA': 9, 'LL:AB': 10, 'LL:CB': 8, 'LL:AC': 9, 'LL:BA': 10}

__OLXOBJ_fltConnCLS__ = {'3LG': [0, 1],
                         '2LG:BC': [1, 1], '2LG:CA': [1, 2], '2LG:AB': [1, 3], '2LG:CB': [1, 1], '2LG:AC': [1, 2], '2LG:BA': [1, 3],
                         '1LG:A': [2, 1], '1LG:B': [2, 2], '1LG:C': [2, 3],
                         'LL:BC': [3, 1], 'LL:CA': [3, 2], 'LL:AB': [3, 3], 'LL:CB': [3, 1], 'LL:AC': [3, 2], 'LL:BA': [3, 3]}

__OLXOBJ_fltConnSIM__ = {'3LG': [0, 0],
                         '2LG:BC': [1, 1], '2LG:CA': [1, 2], '2LG:AB': [1, 0], '2LG:CB': [1, 1], '2LG:AC': [1, 2], '2LG:BA': [1, 0],
                         '1LG:A': [2, 0], '1LG:B': [2, 1], '1LG:C': [2, 2],
                         'LL:BC': [3, 1], 'LL:CA': [3, 2], 'LL:AB': [3, 0], 'LL:CB': [3, 1], 'LL:AC': [3, 2], 'LL:BA': [3, 0]}

__OLXOBJ_STR_KEYS__ = {'GEN': 'GENERATOR', 'CCGEN': 'CCGENUNIT', 'XFMR': 'XFORMER', 'XFMR3': 'XFORMER3', 'MULINE': 'MUPAIR',
                       'SHUNTUNIT': 'CAPUNIT', 'RLYGROUP': 'RELAYGROUP', 'RLYOCG': 'OCRLYG', 'RLYOCP': 'OCRLYP','RLYOC': 'OCRLY',
                       'RLYDSG': 'DSRLYG', 'RLYDSP': 'DSRLYP','RLYDS': 'DSRLY',
                       'RLYD': 'DEVICEDIFF', 'RLYV': 'DEVICEVR', 'SCHEME': 'PILOT', 'RECLSR': 'RECLSRP'}
__OLXOBJ_STR_KEYS1__ = {v:k for k,v in __OLXOBJ_STR_KEYS__.items()}


def __MU_getValue__(l1, l2, lx1, lx2, sParam, value):
    if sParam not in {'R','X','FROM1','TO1','FROM2','TO2'}:
        return sParam, value
    vd = {sParam: value}
    b1, b2, id1 = l1.BUS1, l1.BUS2, l1.CID
    b3, b4, id2 = l2.BUS1, l2.BUS2, l2.CID
    # case1
    if b1.__hnd__==lx1[0].__hnd__ and b2.__hnd__==lx1[1].__hnd__ and id1==lx1[2] and \
       b3.__hnd__==lx2[0].__hnd__ and b4.__hnd__==lx2[1].__hnd__ and id2==lx2[2]:
       res = __MU_convert__(vd, case=1)
    # case2
    elif b1.__hnd__==lx1[0].__hnd__ and b2.__hnd__==lx1[1].__hnd__ and id1==lx1[2] and \
       b3.__hnd__==lx2[1].__hnd__ and b4.__hnd__==lx2[0].__hnd__ and id2==lx2[2]:
       res = __MU_convert__(vd, case=2)
    # case3
    elif b1.__hnd__==lx1[1].__hnd__ and b2.__hnd__==lx1[0].__hnd__ and id1==lx1[2] and \
       b3.__hnd__==lx2[1].__hnd__ and b4.__hnd__==lx2[0].__hnd__ and id2==lx2[2]:
       res = __MU_convert__(vd, case=3)
    # case4
    elif b1.__hnd__==lx1[1].__hnd__ and b2.__hnd__==lx1[0].__hnd__ and id1==lx1[2] and \
       b3.__hnd__==lx2[0].__hnd__ and b4.__hnd__==lx2[1].__hnd__ and id2==lx2[2]:
       res = __MU_convert__(vd, case=4)
    # case5
    elif b1.__hnd__==lx2[0].__hnd__ and b2.__hnd__==lx2[1].__hnd__ and id1==lx2[2] and \
       b3.__hnd__==lx1[0].__hnd__ and b4.__hnd__==lx1[1].__hnd__ and id2==lx1[2]:
       res = __MU_convert__(vd, case=5)
    # case6 3 - 4  2 - 1
    elif b1.__hnd__==lx2[0].__hnd__ and b2.__hnd__==lx2[1].__hnd__ and id1==lx2[2] and \
       b3.__hnd__==lx1[1].__hnd__ and b4.__hnd__==lx1[0].__hnd__ and id2==lx1[2]:
       res = __MU_convert__(vd, case=60)
    # case7 4 - 3  1 - 2
    elif b1.__hnd__==lx2[1].__hnd__ and b2.__hnd__==lx2[0].__hnd__ and id1==lx2[2] and \
       b3.__hnd__==lx1[0].__hnd__ and b4.__hnd__==lx1[1].__hnd__ and id2==lx1[2]:
       res = __MU_convert__(vd, case=70)
    # case8
    elif b1.__hnd__==lx2[1].__hnd__ and b2.__hnd__==lx2[0].__hnd__ and id1==lx2[2] and \
       b3.__hnd__==lx1[1].__hnd__ and b4.__hnd__==lx1[0].__hnd__ and id2==lx1[2]:
       res = __MU_convert__(vd, case=8)
    else:
        raise Exception('Error case MutualCoupling')
    return res


def __MU_convert__(vd0=dict(),case=1,sMu=''):
    """
    Case1: 1 - 2  3 - 4        Case2:  1 - 2  3 - 4
           1 - 2  3 - 4                1 - 2  4 - 3
    Case3: 1 - 2  3 - 4        Case4:  1 - 2  3 - 4
           2 - 1  4 - 3                2 - 1  3 - 4
    Case5: 1 - 2  3 - 4        Case6:  1 - 2  3 - 4
           3 - 4  1 - 2                3 - 4  2 - 1
    Case7: 1 - 2  3 - 4        Case8:  1 - 2  3 - 4
           4 - 3  1 - 2                4 - 3  2 - 1
    """
    if sMu:
        va = [float(v1) for v1 in sMu.split()]
        vd = dict()
        nl = int(len(va)/6)
        vd['R'] = [va[6*i] for i in range(nl)]
        vd['X'] = [va[6*i+1] for i in range(nl)]
        vd['FROM1'] = [va[6*i+2] for i in range(nl)]
        vd['TO1'] = [va[6*i+3] for i in range(nl)]
        vd['FROM2'] = [va[6*i+4] for i in range(nl)]
        vd['TO2'] = [va[6*i+5] for i in range(nl)]
    else:
        vd = vd0.copy()
    vres, vres1 = dict(), dict()
    for k,v in vd.items():
        nl = len(v)
        break
    #
    if case==1:
        vres = vd.copy()
    elif case==2:
        if 'R' in vd.keys():
            vres['R'] = [-vi for vi in vd['R'][:nl]]
        if 'X' in vd.keys():
            vres['X'] = [-vi for vi in vd['X'][:nl]]
        if 'FROM1' in vd.keys():
            vres['FROM1'] = vd['FROM1'][:nl]
        if 'TO1' in vd.keys():
            vres['TO1'] = vd['TO1'][:nl]
        if 'TO2' in vd.keys():
            vres['FROM2'] = [100-vi for vi in vd['TO2'][:nl]]
        if 'FROM2' in vd.keys():
            vres['TO2'] = [100-vi for vi in vd['FROM2'][:nl]]
    elif case==3:
        if 'R' in vd.keys():
            vres['R'] = vd['R'][:nl]
        if 'X' in vd.keys():
            vres['X'] = vd['X'][:nl]
        if 'TO1' in vd.keys():
            vres['FROM1'] = [100-vd['TO1'][i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres['TO1'] = [100-vd['FROM1'][i] for i in range(nl)]
        if 'TO2' in vd.keys():
            vres['FROM2'] = [100-vd['TO2'][i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres['TO2'] = [100-vd['FROM2'][i] for i in range(nl)]
    elif case==4:
        if 'R' in vd.keys():
            vres['R'] = [-vd['R'][i] for i in range(nl)]
        if 'X' in vd.keys():
            vres['X'] = [-vd['X'][i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres['FROM1'] = [100-vd['TO1'][i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres['TO1'] = [100-vd['FROM1'][i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres['FROM2'] = [vd['FROM2'][i] for i in range(nl)]
        if 'TO2' in vd.keys():
            vres['TO2'] = [vd['TO2'][i] for i in range(nl)]
    elif case==5:
        if 'R' in vd.keys():
            vres['R'] = [vi for vi in vd['R'][:nl]]
        if 'X' in vd.keys():
            vres['X'] = [vi for vi in vd['X'][:nl]]
        if 'FROM2' in vd.keys():
            vres['FROM1'] = [vi for vi in vd['FROM2'][:nl]]
        if 'TO2' in vd.keys():
            vres['TO1'] = [vi for vi in vd['TO2'][:nl]]
        if 'FROM1' in vd.keys():
            vres['FROM2'] = [vi for vi in vd['FROM1'][:nl]]
        if 'TO1' in vd.keys():
            vres['TO2'] = [vi for vi in vd['TO1'][:nl]]
    elif case==6:
        if 'R' in vd.keys():
            vres['R'] = [-vd['R'][nl-1-i] for i in range(nl)]
        if 'X' in vd.keys():
            vres['X'] = [-vd['X'][nl-1-i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres1['FROM1'] = [vi for vi in vd['FROM2']]
            vres['TO1'] = [100-vres1['FROM1'][nl-1-i] for i in range(nl)]
        if 'TO2' in vd.keys():
            vres1['TO1'] = [vi for vi in vd['TO2']]
            vres['FROM1'] = [100-vres1['TO1'][nl-1-i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres1['FROM2'] = [100-vi for vi in vd['TO1']]
            vres['TO2'] = [100-vres1['FROM2'][nl-1-i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres1['TO2'] = [100-vi for vi in vd['FROM1']]
            vres['FROM2'] = [100-vres1['TO2'][nl-1-i] for i in range(nl)]
    elif case==60:
        if 'R' in vd.keys():
            vres['R'] = [-vd['R'][nl-1-i] for i in range(nl)]
        if 'X' in vd.keys():
            vres['X'] = [-vd['X'][nl-1-i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres['FROM1'] = [vd['FROM2'][nl-1-i] for i in range(nl)]
        if 'TO2' in vd.keys():
            vres['TO1'] = [vd['TO2'][nl-1-i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres['FROM2'] = [100-vd['TO1'][nl-1-i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres['TO2'] = [100-vd['FROM1'][nl-1-i] for i in range(nl)]
    elif case==7:
        if 'R' in vd.keys():
            vres['R'] = [-vi for vi in vd['R'][:nl]]
        if 'X' in vd.keys():
            vres['X'] = [-vi for vi in vd['X'][:nl]]
        if 'TO2' in vd.keys():
            vres1['FROM1'] = [100-vi for vi in vd['TO2']]
            vres['TO1'] = [100-vres1['FROM1'][i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres1['TO1'] = [100-vi for vi in vd['FROM2']]
            vres['FROM1'] = [100-vres1['TO1'][i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres1['FROM2'] = [vi for vi in vd['FROM1']]
            vres['TO2'] = [100-vres1['FROM2'][i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres1['TO2'] = [vi for vi in vd['TO1']]
            vres['FROM2'] = [100-vres1['TO2'][i] for i in range(nl)]
    elif case==70:
        if 'R' in vd.keys():
            vres['R'] = [-vi for vi in vd['R'][:nl]]
        if 'X' in vd.keys():
            vres['X'] = [-vi for vi in vd['X'][:nl]]
        if 'TO2' in vd.keys():
            vres['FROM1'] = [100-vd['TO2'][i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres['TO2'] = [vd['TO1'][i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres['FROM2'] = [vd['FROM1'][i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres['TO1'] = [100-vd['FROM2'][i] for i in range(nl)]
    elif case==8:
        if 'R' in vd.keys():
            vres['R'] = [vd['R'][i] for i in range(nl)]
        if 'X' in vd.keys():
            vres['X'] = [vd['X'][i] for i in range(nl)]
        if 'TO2' in vd.keys():
            vres['FROM1'] = [100-vd['TO2'][i] for i in range(nl)]
        if 'FROM2' in vd.keys():
            vres['TO1'] = [100-vd['FROM2'][i] for i in range(nl)]
        if 'TO1' in vd.keys():
            vres['FROM2'] = [100-vd['TO1'][i] for i in range(nl)]
        if 'FROM1' in vd.keys():
            vres['TO2'] = [100-vd['FROM1'][i] for i in range(nl)]
    #
    if sMu:
        sres, st = '', '  '
        if len(vres['FROM1'])>1 and vres['FROM1'][0]>vres['FROM1'][1]:
            for k in vres.keys():
                vres[k] = [vres[k][nl-1-i] for i in range(nl)]
        #
        for i in range(nl):
            sres += str(vres['R'][i]) if vres['R'][i] != 0 else '0.0'
            sres += st + str(vres['X'][i])+st+str(vres['FROM1'][i])+st+str(vres['TO1'][i])
            sres += st + str(vres['FROM2'][i])+st+str(vres['TO2'][i])+st
        return sres
    #
    if len(vres.keys())==1:
        for k,v in vres.items():
            return k,v
    return vres
#
def __MU_setValue__(o1):
    hnd, l1, l2 = o1.HANDLE, o1.LINE1, o1.LINE2
    lx1, lx2 = o1.__paramEx__['MULINEVAL']['ORIENTLINE1'], o1.__paramEx__['MULINEVAL']['ORIENTLINE2']
    try:
        xx = o1.__paramEx__['MULINEVAL']['X']
    except:
        xx = o1.X
    try:
        rr = o1.__paramEx__['MULINEVAL']['R']
    except:
        rr = o1.R
    try:
        for i in range(5):
            if xx[i] != 0.0 or rr[i] != 0.0:
                nl = i+1
    except:
        nl = 5
    vn = dict()
    for k, v in o1.__paramEx__['MULINEVAL'].items():
        if k not in {'ORIENTLINE1', 'ORIENTLINE2'} and v!=None:
            sParam,value = __MU_getValue__(l1, l2, lx1, lx2, k, v[:nl])
            vn[sParam] = value
    #
    flag = False
    if 'TO1' in vn.keys():
        if len(vn['TO1'])>1 and vn['TO1'][0]>vn['TO1'][1]:
            flag = True
    if 'FROM1' in vn.keys():
        if len(vn['FROM1'])>1 and vn['FROM1'][0]>vn['FROM1'][1]:
            flag = True
    if flag:
        for k in vn.keys():
            vn[k] = [vn[k][nl-1-i] for i in range(nl)]
    #add 00
    for k,v in vn.items():
        for i in range(nl,5):
            vn[k].append(0)
    #
    for sParam,value in vn.items():
        paramCode = __OLXOBJ_PARA__['MULINE'][sParam][0]
        val1 = __setValue__(hnd, paramCode, value)
        if OLXAPI_OK != SetData(c_int(hnd), c_int(paramCode), byref(val1)):
            messError = '\n'+ErrorString()+'\nCheck MULINE.%s' % (sParam)
            raise Exception(messError)
    for k in o1.__paramEx__['MULINEVAL'].keys():
        if k not in {'ORIENTLINE1', 'ORIENTLINE2'}:
            o1.__paramEx__['MULINEVAL'][k] = None
