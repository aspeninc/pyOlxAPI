"""
Purpose: tool internal ASPEN

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "2.1.5"

#
import os,sys,time, uuid
from tqdm import tqdm
from operator import setitem
from xml.etree import cElementTree as CET
import OlxAPI
import OlxAPILib
from OlxAPIConst import *
import AppUtils
import OlxObj
import tkinter as tk
import tkinter.messagebox as tkm
sTab = ['','\n    ','\n        ','\n            ','\n                ','\n                    ', '\n                        ']

dict_OBJ = {
#[            TXT                       ,CHF                   ,ACR ,COMP            ,COMPTxt]
'BUS'      : ['Bus'                     ,'BUS DATA'            ,''  ,'COMPBUSES'      ,'Buses'],
'GEN'      : ['Generator'               ,'BUS REGULATION'      ,''  ,'COMPGENS'       ,'Generators'],
'GENUNIT'  : ['Gen. unit'               ,'GENERATOR UNIT'      ,''  ,''               ,''],
'GENW3'    : ['Type-3 Wind Plant'       ,'TYPE-3 WIND GENERATOR','' ,''               ,''],
'GENW4'    : ['Converter-Interfaced Resource' ,'CONVERTER-INTERFACED RESOURCE'    ,''  ,'',''],
'CCGEN'    : ['Vol. Controlled Current Source','VOLTAGE CONTROLLED CURRENT SOURCE',''  ,'',''],
'LOAD'     : ['Load on bus'             ,'LOADEX'              ,''  ,'COMPLOADS'      ,'Loads'],
'LOADUNIT' : ['Load unit'               ,'LOAD UNIT'           ,''  ,''               ,''],
'SHUNT'    : ['Shunt on bus'            ,'SHUNTEX DATA'        ,''  ,'COMPSHUNTS'     ,'Shunts'],
'SHUNTUNIT': ['Shunt unit'              ,'SHUNT UNIT DATA'     ,''  ,''               ,''],
'SVD'      : ['Switched Shunt'          ,'SWITCHED SHUNT DATA' ,''  ,'COMPSVDS'       ,'Switched Shunt'],
'LINE'     : ['Trans. Line'             ,'TRANSMISSION LINE'   ,'L' ,'COMPLINES'      ,'Transmission lines'],
'SERIESRC' : ['Series Comp. Device'     ,'TRANSMISSION LINE'   ,'L' ,'',''],
'DCLINE2'  : ['DC Line'                 ,'DC LINE DATA'        ,'D' ,'COMPDCLINES'    ,'DC lines'],
'SHIFTER'  : ['Ph. Shifter'             ,'PHASE SHIFTER'       ,'P' ,'COMPSHIFTERS'   ,'Phase shifters'],
'SWITCH'   : ['Switch.'                 ,'SWITCH DATA'         ,'W' ,'COMPSWITCHES'   ,'Switches'],
'XFMR'     : ['2-W Transf.'             ,'2W TRANSFORMER'      ,'T' ,'COMPXFMRS'      ,'Transformers'],
'XFMR3'    : ['3-W Transf.'             ,'3W TRANSFORMER'      ,'X' ,''               ,''],
'LTC'      : ['Load Tap Changer 2-W Transf.','LOAD TAP CHANGER 2W TRANSFORMER','T' ,'',''],
'LTC3'     : ['Load Tap Changer 3-W Transf.','LOAD TAP CHANGER 3W TRANSFORMER','X' ,'',''],
'ZCORRECT' : ['Impedance correction table'  ,'IMPEDANCE CORRECTION'           ,'IC','',''],
'BREAKER'  : ['Breaker'                 ,'BREAKER'             ,''  ,'COMPBREAKERS'   ,'Breakers'],
'MULINE'   : ['Mutual'                  ,'MUTUAL'              ,'L' ,'COMPMUPAIRS'    ,'Mutual couplings'],
'AREA'     : ['Area No.'                ,'AREA'                ,''  ,'COMPAREAS'      ,'Areas and zones'],
'ZONE'     : ['Zone No.'                ,'ZONE'                ,''  ,''               ,''],
'RLYGROUP' : ['Relay group'             ,'RELAYGROUP'          ,''  ,'',''],
'RLYOC'    : ['Overcurrent Relay'       ,'OC GROUND RELAY'            ,''  ,'COMPRLYOC'      ,'Overcurrent relays and fuses'],
'FUSE'     : ['Fuse'                    ,'OC FUSE'             ,''  ,''               ,''],
'RLYDS'    : ['Distance Relay'          ,'DS RELAY'            ,''  ,'COMPRLYDS'      ,'Distance relays'],
'RLYD'     : ['Other protective devices','OTHER PROTECTIVE DEVICES','','COMPDEVICES'  ,'Other protective devices'],
'RLYV'     : ['Voltage Relay'           ,'VOLTAGE RELAY'       ,''  ,'',''],
'RECLSR'   : ['Recloser'                ,'RECLOSER'            ,''  ,'',''],
'SCHEME'   : ['Protection scheme'       ,'PROTECTION SCHEME'   ,''  ,'COMPSCHEMES'    ,'Protection schemes'],
'COORDPAIR': ['Coordination pair'       ,'COORDPAIR'           ,''  ,'COMPCOORDPAIRS' ,'Coordination pair']
}
OBJTYPELIST = list(dict_OBJ.keys())
ACTION_ADX = ['MODIFY','DELETE','ADD']
GRAPHICPR1 = {'GEN':'GN_', 'GENW3':'G3_', 'GENW4':'G4_', 'CCGEN':'GC_', 'LOAD':'LD_', 'SHUNT':'CP_','SVD':'CS_'}
#
def isShipVersion():
    pe = os.path.basename(sys.executable)
    return (pe not in ['python.exe','pythonw.exe'])
#
def pauseFinal(verbose):
    if verbose>0 and isShipVersion():
        print()
        os.system("pause")
#
def __graphic_obj(obj,nseg=1): ## noSegs,p1X,p1Y,p2X....,text1X,text1Y,text2X,text2Y
    if obj=='BUS':
        return ['BS_SIZE','BS_ANGLE','BS_X','BS_Y','BS_NAMEX','BS_NAMEY','BS_HIDEID']
    if obj=='LINE':
        na = ['LN_NOSEGS']
        for i in range(nseg+1):
            na.extend(['LN_P%iX'%(i+1),'LN_P%iY'%(i+1)])
        na.extend(['LN_TEXT1X','LN_TEXT1Y','LN_TEXT2X','LN_TEXT2Y'])
        return na
    #
    if obj in GRAPHICPR1.keys():
        p1 = GRAPHICPR1[obj]
        na = [p1+'XY',p1+'ANGLE',p1+'TEXTX',p1+'TEXTY']
        return na
#
def Olx2Olr_Graphic(olxFile,olrFile,olxpath=''):
    olxFile = os.path.abspath(olxFile)
    pi = [os.path.abspath(olxFile) , olrFile]
    OlrConverter(pi,olxpath)
    #
    aspen_olx = DataASPEN_OLX(olxFile)
    OlxObj.OLCase.open(olrFile,0)
    #
    for data in aspen_olx.xmlDict['OLXDBTABLE']:
        olxtable = data['OLXREC']
        if type(olxtable) != list:
            olxtable = [data['OLXREC']]
        for olxrec in olxtable:
            o1=OlxObj.OLCase.findOBJ(olxrec['@OBJGUID'])
            da = olxrec['DATAFIELD']
            if olxrec['@OBJTYPE']=='BUS': #size,angle,x,y,nameX,nameX,hideID
                value = []
                na = __graphic_obj('BUS')
                for n1 in na:
                    for d1 in da:
                        if d1['@NAME']==n1:
                            value.append( int(d1['@VALUE']) )
                            break
                o1.changeData('GRAPHIC',value)
            elif olxrec['@OBJTYPE'] in GRAPHICPR1.keys():# x|y,angle,textX,textY
                na = __graphic_obj(olxrec['@OBJTYPE'])
                value = []
                for n1 in na:
                    for d1 in da:
                        if d1['@NAME']==n1:
                            value.append( int(d1['@VALUE']) )
                            break
                #
                o1.changeData('GRAPHIC',value)
            elif olxrec['@OBJTYPE']=='LINE': # noSegs,p1X,p1Y,p2X....,text1X,text1Y,text2X,text2Y
                nseg = 0
                for d1 in da:
                    if d1['@NAME']=='LN_NOSEGS':
                        nseg = int(d1['@VALUE'])
                        value= [nseg]
                        break
                if nseg>0:
                    na = __graphic_obj('LINE',nseg)
                    value = []
                    for n1 in na:
                        for d1 in da:
                            if d1['@NAME']==n1:
                                value.append( int(d1['@VALUE']) )
                                break
                    o1.changeData('GRAPHIC',value)
    #
    OlxObj.OLCase.save()
    OlxObj.OLCase.close()
#
def Olr2Olx_Graphic(olrFile,olxFile,olxpath=''):
    olxFile = os.path.abspath(olxFile)
    pi = [os.path.abspath(olrFile) , olxFile]
    OlrConverter(pi,olxpath)
    #
    aspen_olx = DataASPEN_OLX(olxFile)
    OlxObj.OLCase.open(olrFile,1)
    #
    for data in aspen_olx.xmlDict['OLXDBTABLE']:
        olxtable = data['OLXREC']
        if type(olxtable) != list:
            olxtable = [data['OLXREC']]
        for olxrec in olxtable:
            o1= OlxObj.OLCase.findOBJ(olxrec['@OBJGUID'])
            if olxrec['@OBJTYPE']=='BUS':
                d1 = o1.getData('GRAPHIC')[:-1] #size,angle,x,y,nameX,nameX,hideID
                na = __graphic_obj('BUS')[:-1]
                for i in range(len(na)):
                    olxrec['DATAFIELD'].insert(0,{'@VALUE':str(d1[-i-1]),'@NAME':na[-i-1]})
                #olxrec['GFXDATA'] = {'@BS_SIZE':d1[0],'@BS_ANGLE':d1[1],'@BS_X':d1[2],'@BS_Y':d1[3],'@BS_NAMEX':d1[4],'@BS_NAMEY':d1[5]}
            #
            if olxrec['@OBJTYPE'] in GRAPHICPR1.keys():
                d1 = o1.getData('GRAPHIC') # x|y,angle,textX,textY
                na = __graphic_obj(olxrec['@OBJTYPE'])
                for i in range(len(na)):
                    olxrec['DATAFIELD'].insert(0,{'@VALUE':str(d1[-i-1]),'@NAME':na[-i-1]})
            #
            if olxrec['@OBJTYPE']=='LINE':
                d1 = o1.getData('GRAPHIC')
                na = __graphic_obj('LINE',d1[0])
                for i in range(len(na)):
                    olxrec['DATAFIELD'].insert(0,{'@VALUE':str(d1[-i-1]),'@NAME':na[-i-1]})
    #
    aspen_olx.export2XML(olxFile)
    print('File saved as:',olxFile)
#
def OlrConverter(pi,olxpath,pyFile='',sizeSubprocess = {'.OLR':10,'.OLX':200,'.DXT':20}):
    """
    Convert OLR <=>OLX,DXT
            2 OLR =>ADX

    Parameters
    ----------
    pi = [fi,fo] list input/ouput file
        fi : Filepath for input file
        fo : Filepath for out file

    samples:
        OlrConverter(pi=['haha.OLR','hoho.OLX']) : 'haha.OLR' => 'hoho.OLX'
        OlrConverter(pi=['haha.OLX','hoho.OLR']) : 'haha.OLX' => 'hoho.OLR'
        OlrConverter(pi=['haha1.OLR','haha2.OLR','hoho.ADX']) : compare ('haha1.OLR','haha2.OLR') => 'hoho.ADX'

    """
    t0 = time.time()
    if pyFile=='' or sizeSubprocess==None:
        return __OlrConverter(pi,olxpath)
    #
    ext1 = (os.path.splitext(pi[0])[1]).upper()
    fsz1 = os.path.getsize(pi[0])/1e6
    slog = '\tFile ' +ext1 + '           : %.2f MB'%fsz1

    subpro = False
    try:
        if fsz1>sizeSubprocess[ext1]:
            subpro= True
    except:
        pass
    #
    if len(pi)==3:
        slog = '\tFile ' +ext1 + ' [1]       : %.2f MB'%fsz1
        ext1 = (os.path.splitext(pi[1])[1]).upper()
        fsz1 = os.path.getsize(pi[1])/1e6
        slog += '\n\tFile ' +ext1 + ' [2]       : %.2f MB'%fsz1
        try:
            if fsz1>sizeSubprocess[ext1]:
                subpro= True
        except:
            pass
    #
    slog += '\t\tSubprocess : '+ str(subpro)
    import logging
    logging.info(slog)
    try:
        if not subpro:
            __OlrConverter(pi,olxpath)
        else:
            """
            run OLR convert OLR <=>OLX,DXT, ADX with subprocess
            """
            OlxAPI.UnloadOlxAPI()
            #
            pe = os.path.basename(sys.executable)
            if pe in ['python.exe','pythonw.exe']:
                args = [sys.executable,pyFile]
            else:
                args = [sys.executable] # pyFile.exe
            #
            #
            args.extend(['-olxpath',olxpath,'-pi'])
            args.extend(pi)
            #
            out,err,returncode = AppUtils.runSubprocess(args)
            if returncode>0:
                raise Exception(err)
    except:
        pass
    #
    if not os.path.isfile(pi[len(pi)-1]) and os.path.isfile('aspencim.exe'):
        args = ['aspencim.exe','-pi']
        args.extend(pi)
        out,err,returncode = AppUtils.runSubprocess(args)
        if returncode>0:
            raise Exception(err)
    #
    if not os.path.isfile(pi[len(pi)-1]):
        raise Exception("OneLiner data conversion failed.")
    #
    ext1 = (os.path.splitext(pi[len(pi)-1])[1]).upper()
    fsz1 = os.path.getsize(pi[len(pi)-1])/1e6
    slog = '\tFile size of ' +ext1 + ' generated : %.2f MB'%fsz1
    slog += '\t\ttime       : %0.2f s'%(time.time()-t0)
    logging.info(slog)

#
def __OlrConverter(pi,olxpath):
    """
    Convert OLR <=>OLX,DXT
            2 OLR =>ADX

    Parameters
    ----------
    pi = [fi,fo] list input/ouput file
        fi : Filepath for input file
        fo : Filepath for out file

    samples:
        OlrConverter(pi=['haha.OLR','hoho.OLX']) : 'haha.OLR' => 'hoho.OLX'
        OlrConverter(pi=['haha.OLX','hoho.OLR']) : 'haha.OLX' => 'hoho.OLR'
        OlrConverter(pi=['haha1.OLR','haha2.OLR','hoho.ADX']) : compare ('haha1.OLR','haha2.OLR') => 'hoho.ADX'

    """
    ext = []
    for p1 in pi:
        ext.append((os.path.splitext(p1)[1]).upper())
    #
    if len(pi)==2:
        if ext[0] not in['.OLR','.OLX','.DXT'] or ext[1] not in['.OLR','.OLX','.DXT']:
            raise Exception(" error format .OLR or .OLX or .DXT")
        #
        if ext[0] =='.OLR' and ext[1] in ['.OLX','.DXT']:
            OlxAPILib.open_olrFile(pi[0],dllPath = olxpath,prt=True) # read OLR file
            #
            if ext[1] =='.OLX':
                cmdParams = '<EXPORTNETWORK FORMAT="XML" XMLPATHNAME="' + pi[1] + '" />'
            else:
                cmdParams = '<EXPORTNETWORK FORMAT="DXT" DXTPATHNAME="' + pi[1] + '" />'
            #
            if not OlxAPILib.run1LPFCommand(cmdParams):
                raise Exception('ERROR convert OLR <=> OLX,DXT')
        elif ext[0] in ['.OLX','.DXT'] and ext[1] =='.OLR':
            OlxAPI.InitOlxAPI(olxpath,prt=True)
            #
            if not OlxAPI.LoadDataFile(pi[0],1):
                raise Exception('ERROR convert OLR <=> OLX,DXT')
            OlxAPILib.saveAsOlr(pi[1])
            OlxAPILib.open_olrFile(pi[1],dllPath = olxpath,readonly=1,prt=True)
    #
    elif len(pi)==3:
        if ext[0] != '.OLR' or ext[1] != '.OLR' or ext[2] != '.ADX':
            raise Exception("ERROR format for compare 2OLR => ADX")
        #
        OlxAPILib.open_olrFile(pi[0],dllPath = olxpath,readonly=1,prt=True)
        #
        inputParams = '<DIFFANDMERGE FILEPATHB="' +  pi[1] + '"' +' DIFFTYPE="3"'+ \
                        ' FILEPATHDIFF="' +  pi[2] + '" />'
        if not OlxAPILib.run1LPFCommand(inputParams):
            raise Exception("ERROR compare 2OLR => ADX")
    #
    return 0

class FilterOptions:
    def __init__(self,fCFG,PARSER_INPUTS):
        self.typ = []                  # ['BUS','LINE']
        self.action = []               # list of action ['MODIFY','DELETE','ADD']
        self.area = []                 # list of area number
        self.inside_area = True
        self.zone = []                 # list of zone number
        self.inside_zone = True
        self.tie = True                # include tie between selected items and the rest of the network
        self.ckid = []                 # ckid of tieline
        self.kv = []                   # [kvmin, kvmax]
        self.BNO = []                  # list of bus number
        self.OBJGUID = []              # list of objguid
        #
        self.__readOptionsConfig__(fCFG,PARSER_INPUTS)
        #print(self.optionsConfig)
        self.__getFilter__(self.optionsConfig)
    #
    def __readOptionsConfig__(self,fCFG,PARSER_INPUTS):
        import logging
        self.optionsConfig = {}
        if os.path.isfile(fCFG):
            logging.info('\nXML file with comparison options:\n"%s"'%fCFG)
            op = xml2Dict(fCFG,desc='',tagStop=None, encoding = 'iso8859-1')['ASPEN_CASECOMP_CONFIG']
            if 'ROW' in op.keys():
                try:
                    op = xml2Dict(fCFG,desc='',tagStop=None, encoding = 'iso8859-1')['ASPEN_CASECOMP_CONFIG']['ROW']
                    for op1 in op:
                        name = op1['@NAME']
                        val = op1['@VALUE']
                        self.optionsConfig[name] = val
                    return
                except:
                    PARSER_INPUTS.print_help()
                    raise Exception("XML file with comparison options : error format")
            else:
                try:
                    for k,v in op.items():
                        self.optionsConfig[k[1:]] = v
                    return
                except:
                    PARSER_INPUTS.print_help()
                    raise Exception("XML file with comparison options : error format")
        #
        elif not os.path.isfile(fCFG) and fCFG:
            PARSER_INPUTS.print_help()
            raise Exception('\nXML file with comparison options NOT FOUND:\n"%s"'%fCFG)
        else:
            logging.info('\nNot Found XML file with comparison options: use options by default')
        #
        for key,val in dict_OBJ.items():
            comp = val[3] #[            TXT                       ,CHF                   ,ACR ,COMP            ,COMPTxt]
            if comp!='':
                self.optionsConfig[comp]= '1'
        #
        self.optionsConfig['USEBUSNO']= '0'
        self.optionsConfig['COMPDEVICES']= '1'
        self.optionsConfig['COMPTIES']= '1'
        self.optionsConfig['ANAFASFORMAT']= '0'
        self.optionsConfig['COMPEXTENT']= '0'
        self.optionsConfig['KVRANGE']= '0-9999'
        self.optionsConfig['BOUNDARYID']= 'N'
        self.optionsConfig['AREAS']= '0-99999'
        self.optionsConfig['ZONES']= '0-99999'
    #
    def isAllNetwork(self):
        if self.typ or self.action or self.area or self.zone or self.kv or self.ckid or self.BNO or self.OBJGUID:
            return False
        return True
    #
    def printStatistic(self):
        if self.isAllNetwork():
            print('COMPARISON OPTIONS: ENTIRE NETWORK')
        else:
            print('COMPARISON OPTIONS:')
            print('OBJTYPE: ',  self.typ)
            print('ACTION: ' , self.action)
            print('AREA: ' , self.optionsConfig['AREAS'] , ' inside AREA: ', self.inside_area)
            print('ZONE: ' , self.optionsConfig['ZONES'] , ' inside ZONE: ' , self.inside_zone)
            print('TIE: ' , self.tie )
            print('CKID: ', self.ckid)
            print('KV: ' , self.kv)
            print('BNO: ' , self.BNO)
            print('OBJGUID: ' , self.OBJGUID)
        print()
    #
    def __getFilter__(self,optionsConfig):
        a = False
        for o1 in OBJTYPELIST:
            comp = dict_OBJ[o1][3]
            if comp!='' and  optionsConfig[comp]=='0':
                a = True
                break
        if a:
            # TYP
            if optionsConfig['COMPBUSES']=='1':
                self.typ.extend(['BUS'])
            if optionsConfig['COMPGENS']=='1':
                self.typ.extend(['GEN','GENUNIT','GENW3','GENW4','CCGEN'])
            if optionsConfig['COMPLOADS']=='1':
                self.typ.extend(['LOAD','LOADUNIT'])
            if optionsConfig['COMPSHUNTS']=='1':
                self.typ.extend(['SHUNT','SHUNTUNIT'])
            if optionsConfig['COMPSVDS']=='1':
                self.typ.extend(['SVD'])
            if optionsConfig['COMPLINES']=='1':
                self.typ.extend(['LINE','SERIESRC'])
            if optionsConfig['COMPDCLINES']=='1':
                self.typ.extend(['DCLINE2'])
            if optionsConfig['COMPSWITCHES']=='1':
                self.typ.extend(['SWITCH'])
            if optionsConfig['COMPSHIFTERS']=='1':
                self.typ.extend(['SHIFTER'])
            if optionsConfig['COMPCOORDPAIRS']=='1':
                self.typ.extend(['COORDPAIR'])
            if optionsConfig['COMPXFMRS']=='1':
                self.typ.extend(['XFMR','XFMR3','LTC','LTC3'])
            if optionsConfig['COMPSHIFTERS']=='1' or optionsConfig['COMPXFMRS']=='1' :
                self.typ.extend(['ZCORRECT'])
            if optionsConfig['COMPBREAKERS']=='1':
                self.typ.extend(['BREAKER'])
            if optionsConfig['COMPMUPAIRS']=='1':
                self.typ.extend(['MULINE'])
            if optionsConfig['COMPAREAS']=='1':
                self.typ.extend(['AREA','ZONE'])
            if optionsConfig['COMPRLYOC']=='1':
                self.typ.extend(['RLYOC','FUSE'])
            if optionsConfig['COMPRLYDS']=='1':
                self.typ.extend(['RLYDS'])
            if optionsConfig['COMPDEVICES']=='1':
                self.typ.extend(['RLYD','RECLSR','RLYV'])
            if optionsConfig['COMPSCHEMES']=='1':
                self.typ.extend(['SCHEME'])
        # AREA <ROW NAME="AREAS" VALUE="0-99999" CMT="Area range: '0-99999'*"/>
        # ZONE <ROW NAME="ZONES" VALUE="0-99999" CMT="Zone range: '0-99999'*"/>
        # '1'- Boundary equipment only; '2'- Inside of areas; '3'- Outside of areas; '4'- Inside zones; '5'- Outside of zones"
        if optionsConfig['COMPEXTENT'] in ['2','3','4','5']: #'0'- Entire network*
            if optionsConfig['AREAS']!='0-99999':
                self.area = self.__splitStr(optionsConfig['AREAS'])
                if optionsConfig['COMPEXTENT']=='3':
                    self.inside_area = False
                else:
                    self.inside_area = True
            if optionsConfig['ZONES']!='0-99999':
                self.zone = self.__splitStr(optionsConfig['ZONES'])
                if optionsConfig['COMPEXTENT']=='5':
                    self.inside_zone = False
                else:
                    self.inside_zone = True
        #
        if optionsConfig['COMPEXTENT']=='1':
            self.ckid.append(optionsConfig['BOUNDARYID'])
        # TIE
        if optionsConfig['COMPTIES'] =='0': #"Include ties lines in comparison:
            self.tie = False
        else:
            self.tie = True

        # KV
        kva = (optionsConfig['KVRANGE']).split('-')
        kv1 = float(kva[0])
        if len(kva)==1:
            self.kv = [kv1,kv1]
        else:
            kv2 = float(kva[1])
            if not(kv1<=0.0 and kv2>=1500):
                self.kv = [kv1,kv2]
    #
    def __splitStr(self,s1):
        res = []
        a1 = s1.split(',')
        for ai in a1:
            a2 =  ai.split('-')
            if len(a2)==1:
                res.append(int(a2[0]))
            else:
                for i in range(len(a2)-1):
                    for j in range(int(a2[i]),int(a2[i+1])+1):
                        res.append(j)
        #
        res = list(set(res))
        res.sort()
        return res
    #
    def runFilter(self,cr1):
        scope    = cr1['OBJSCOPE']
        objtype1 = cr1['@OBJTYPE']
        action1  = cr1['@ACTION']
        olnetid1 = cr1['@OLNETID']
        #
        if self.isAllNetwork():
            return True
        #
        if self.typ and (objtype1 not in self.typ):
            return False
        #
        if self.action and (action1 not in self.action):
            return False

        # special for area/zone
        if not (self.BNO or self.ckid or self.OBJGUID or self.kv):# if not filter by bno,ckid,kv
            if objtype1=='AREA':
                if self.area :
                    id1 = int(olnetid1)
                    if id1 in self.area:
                        return self.inside_area
                    return not self.inside_area
                elif not self.zone :
                    return True
            # zone
            if objtype1=='ZONE':
                if self.zone :
                    id1 = int(olnetid1)
                    if id1 in self.zone:
                        return self.inside_zone
                    return not self.inside_zone
                elif not self.area :
                    return True
        #
        if scope==None:
            return False
        #
        area,zone,bno,bkv,ckid = [],[],[],[],[]
        #
        for b1 in scope['SCOPEFIELD']:
            if b1['@NAME'] == 'AREA':
                area.append(int(b1['@VALUE']))
            elif b1['@NAME'] == 'ZONE':
                zone.append(int(b1['@VALUE']))
            elif b1['@NAME'] == 'BNO':
                bno.append(int(b1['@VALUE']))
            elif b1['@NAME'] == 'KV':
                bkv.append(float(b1['@VALUE']))
            elif b1['@NAME'] == 'CKTID':
                ckid.append(b1['@VALUE'])
            elif b1['@NAME'] == 'CKTID2':
                ckid.append(b1['@VALUE'])
        #CKID
        if self.ckid:
            test = False
            for c1 in ckid:
                if c1 in self.ckid:
                    test = True
            if not test:
                return False
        #AREA
        if self.area:
            if not area:
                return False
            #
            if self.inside_area:
                if self.tie:      # include tie between selected items and the rest of the network
                    test = False
                    for a1 in area:
                        if a1 in self.area:
                            test = True
                    if not test:
                        return False
                else: # 100% in area
                    for a1 in area:
                        if a1 not in self.area:
                            return False
            else: # outside area
                if self.tie:      # include tie between selected items and the rest of the network
                    test = False
                    for a1 in area:
                        if a1 not in self.area:
                            test = True
                    if not test:
                        return False
                else: # 100% out area
                    for a1 in area:
                        if a1 in self.area:
                            return False
        # ZONE
        if self.zone:
            if not zone:
                return False
            #
            if self.inside_zone:
                if self.tie:      # include tie between selected items and the rest of the network
                    test = False
                    for z1 in zone:
                        if z1 in self.zone:
                            test = True
                    if not test:
                        return False
                else: # 100% in zone
                    for a1 in zone:
                        if a1 not in self.zone:
                            return False
            else: # outside zone
                if self.tie:      # include tie between selected items and the rest of the network
                    test = False
                    for z1 in zone:
                        if z1 not in self.zone:
                            test = True
                    if not test:
                        return False
                else: # 100% out zone
                    for z1 in zone:
                        if z1 in self.zone:
                            return False
        # BNO
        if self.BNO:
            if not bno:
                return False
            #
            if self.tie: # include tie between selected items and the rest of the network
                test = False
                for b1 in bno:
                    if b1 in self.BNO:
                        test = True
                if not test:
                    return False
            else: #100% in
                for b1 in bno:
                    if b1 not in self.BNO:
                        return False
        # BKV
        if self.kv:
            if not bkv:
                return False
            #
            kvmax = max(bkv) # level kv of equipement is the kv max
            if kvmax<self.kv[0] or kvmax>self.kv[1]:
                return False
        #
        return True

class Data1ADX:
    def __init__(self,olnet, chf=None,action=None):
        self.objtype = olnet['@OBJTYPE']
        self.action = action
        #
        self.bn = []
        self.bkv = []
        self.bno = [0,0,0,0]
        #
        self.sid = None
        self.ckid = []
        #
        self.brcode = None
        self.elecode = None
        #
        self.changefield = []
        #
        self.__getData(olnet,chf)
    #
    def __getData(self,olnet,chf):
        #
        for b1 in olnet['OLNETFIELD']:
            if b1['@NAME'].startswith('TERMNAME'):
                self.bn.append(b1['@VALUE'])
            #
            elif b1['@NAME'].startswith('TERMKV'):
                self.bkv.append(b1['@VALUE'])
            #
            elif b1['@NAME'].startswith('TERMBNO'):
                self.bno.append(b1['@VALUE'])
            #
            elif b1['@NAME'].startswith('CKTID'):
                self.ckid.append(b1['@VALUE'])
            #
            elif b1['@NAME'] == 'SID':
                self.sid = b1['@VALUE']
            #
            elif b1['@NAME'] == 'BRCODE':
                self.brcode = b1['@VALUE']
            elif b1['@NAME'] == 'BRCODE2':
                self.brcode = [self.brcode, b1['@VALUE']]
            #
            elif b1['@NAME'] == 'ELECODE':
                self.elecode = b1['@VALUE']
        #
        if self.action=='MODIFY' and chf!=None :
            if type(chf)==list:
                self.changefield = chf
            else:
                self.changefield = [chf]
    #
    def getStringBus(self,lenBNO=6,lenBN=12):
        s1 = ''.ljust(lenBNO-len(str(self.bno[0]))) + str(self.bno[0])
        s1 += ' '+ self.bn[0].ljust(lenBN) + ' ' +str(self.bkv[0])
        if s1.endswith('.'):
            s1 = s1[:len(s1)-1]
        s1+='kV'
        return s1
    #
    def getStringBus_1(self,bn,bkv,bno,lenBNO=6,lenBN=12):
        s1 = ''.ljust(lenBNO-len(str(bno))) + str(bno)
        s1 += ' '+ bn.ljust(lenBN) + ' '
        s2 = '%.2f'%(float(bkv)) +'kV'
        s2 = ''.ljust(8-len(s2)) + s2
        return s1 +s2
    #
    def getStringBr2(self,lenBNO=6,lenBN=12,i0=0):
        s1  = self.getStringBus_1(self.bn[2*i0],self.bkv[0],self.bno[0],lenBNO=lenBNO,lenBN=lenBN)
        s1 +=' -'
        s1 += self.getStringBus_1(self.bn[2*i0+1],self.bkv[1],self.bno[1],lenBNO=lenBNO,lenBN=lenBN)
        s1 += ' ' +self.ckid[i0] + dict_OBJ[self.objtype][2]
        return s1
    #
    def getStringBr3(self,lenBNO=6,lenBN=12,lenL2=40):
        s1  = self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN)
        s1 +='\n'.ljust(lenL2) + ' -'
        s1 += self.getStringBus_1(self.bn[2],self.bkv[2],self.bno[2],lenBNO=lenBNO,lenBN=lenBN)
        s1 += ' ' +self.ckid[0] + dict_OBJ[self.objtype][2]
        return s1
    #
    def getStringMU(self,lenBNO=6,lenBN=12):
        s1  = self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN) +' ' +self.ckid[0] + dict_OBJ[self.objtype][2] +'\n            '
        #
        s1 += self.getStringBus_1(self.bn[2],self.bkv[2],self.bno[2],lenBNO=lenBNO,lenBN=lenBN)
        s1 +=' -'
        s1 += self.getStringBus_1(self.bn[3],self.bkv[3],self.bno[3],lenBNO=lenBNO,lenBN=lenBN)
        s1 += ' ' +self.ckid[1] + dict_OBJ[self.objtype][2]
        return s1
    #
    def getString(self,lenBNO=6,lenBN=12,lenL2=40):
        #
        if self.objtype in ['BUS', 'GEN','LOAD','SHUNT','SVD','BREAKER']:
            return self.getStringBus(lenBNO=lenBNO,lenBN=lenBN)
        elif self.objtype in ['GENUNIT','LOADUNIT','SHUNTUNIT','GENW3','GENW4','CCGEN']:
            if self.sid!=None:
                return ' "'+ self.sid +'"' + ' on bus' + self.getStringBus(lenBNO=lenBNO,lenBN=lenBN)
            return ' on bus' + self.getStringBus(lenBNO=lenBNO,lenBN=lenBN)
        elif self.objtype in ['LINE','SHIFTER','SWITCH','XFMR','LTC','SERIESRC','DCLINE2']:
            return self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN)
        elif self.objtype in ['XFMR3','LTC3']:
            return self.getStringBr3(lenBNO=lenBNO,lenBN=lenBN,lenL2=lenL2)
        elif self.objtype in ['MULINE']:
            return self.getStringMU(lenBNO=lenBNO,lenBN=lenBN)
        elif self.objtype in ['AREA','ZONE']:
            return ''.ljust(9-len(str(self.sid))) + str(self.sid)
        elif self.objtype in ['RLYOC','RLYDS','FUSE','RECLSR','RLYD','RLYV']:
            return ' ' + (self.sid).ljust(20) +' on\n  ' + self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN)+ ' ' +self.ckid[0] + dict_OBJ[self.objtype][2] + self.brcode
        elif self.objtype in ['RLYGROUP']:
            return self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN) + self.brcode
        elif self.objtype in ['ZCORRECT']:
            return ' "'+ self.sid +'"' +self.sid
        elif self.objtype =='COORDPAIR':
            return self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN) + self.brcode[0] +' : '+\
                    self.getStringBr2(lenBNO=lenBNO,lenBN=lenBN,i0=1) + self.brcode[1]
        return '__notsupportedyet__' #'SCHEME'
    #
    def getChange(self,lenS=32):
        s1 = ''
        for dc1 in self.changefield:
            s1 += '\n  '+(dc1['@LABEL']).ljust(lenS)
            try:
                s1 += (dc1['@VALUEA']).ljust(lenS)
                s1 += (dc1['@VALUEB']).ljust(lenS)
            except:
                try:
                    for olnet in dc1['VALUEA']['OLNET']:
                        a1 = Data1ADX(olnet)
                        s1 += '\n' + ''.ljust(lenS-3)+a1.getStringBr2()
                    #
                    for olnet in dc1['VALUEB']['OLNET']:
                        a1 = Data1ADX(olnet)
                        s1 += '\n' + ''.ljust(lenS*2-3)+a1.getStringBr2()
                except:
                    s1+='__notsupportedyet__'
        return s1
#
class DataASPEN_OLX:
    def __init__(self, fxml,tagStop=None,encoding ='iso8859-1'):
        self.xmlDict = xml2Dict(fxml, desc= 'Parsing ASPEN OLX', tagStop=tagStop,encoding =encoding)['ASPENOLXDB']
        # self.ERR = "Input file format required: ASPEN OLX"
    #
    def get_OLRVERSION(self):
        # return String
        return self.xmlDict['@OLRVERSION']

    def get_DATETIME(self):
        return self.xmlDict['@DATETIME']

    #
    def get_OBJCOUNT(self):
        # return dict{}
        return self.xmlDict['OBJCOUNT']
    #
    def get_SYSTEMPARAMS(self):
        # return dict{}
        return self.xmlDict['SYSTEMPARAMS']
    #
    def get_DATA_1(self,filterOptions=None):
        data1 = {}
        for o1 in OBJTYPELIST:
            data1[o1] = []

        return data1
    #

    def export2XML(self, folx):
        objcount = self.xmlDict['OBJCOUNT']

        s0 = ''
        try:
            for key, value in objcount.items():
                s0 += '%s="%s" ' %(str(key[1:]),str(value))
        except:
            pass

        s1 = ''
        try:
            systemparams = self.xmlDict['SYSTEMPARAMS']
            for key, value in systemparams.items():
                if key != "FILECOMMENTS":
                    s1 += '%s="%s" ' %(str(key[1:]),str(value))
                else:
                    s2 = value
        except:
            pass

        xmlDict = self.xmlDict
        fo = open(folx, "w", encoding="utf-8")
        fo.write("<?xml version='1.0'?>\n")
        fo.write('<ASPENOLXDB OLRVERSION="%s" DATETIME="%s">'%(self.xmlDict['@OLRVERSION'], self.xmlDict['@DATETIME']))
        fo.write(sTab[1]+'<OBJCOUNT %s/>'%(s0))
        try:
            udftemplate = self.xmlDict['UDFTEMPLATE']
            fo.write(sTab[1]+'<UDFTEMPLATE>')
            if type(udftemplate['OLRXOBJ']) == list:
                for olrxobj in udftemplate['OLRXOBJ']:
                    fo.write(sTab[2]+'<OLRXOBJ OBJTYPE="%s">'%(olrxobj['@OBJTYPE']))
                    for udf in olrxobj['UDFIELD']:
                        fo.write(sTab[3]+'<UDFIELD ROWNO="%s" FNAME="%s" LABEL="%s"/>'%(udf['ROWNO'],udf['FNAME'],udf['LABEL']))
                    fo.write(sTab[2]+'</OLRXOBJ>')
            elif type(udftemplate['OLRXOBJ']) == dict:
                olrxobj = udftemplate['OLRXOBJ']
                fo.write(sTab[2]+'<OLRXOBJ OBJTYPE="%s">'%(olrxobj['@OBJTYPE']))
                for udf in olrxobj['UDFIELD']:
                    fo.write(sTab[3]+'<UDFIELD ROWNO="%s" FNAME="%s" LABEL="%s"/>'%(udf['@ROWNO'],udf['@FNAME'],udf['@LABEL']))
                fo.write(sTab[2]+'</OLRXOBJ>')

            fo.write(sTab[1]+'</UDFTEMPLATE>')
        except:
            pass

        fo.write(sTab[1] + '<SYSTEMPARAMS %s>'%(s1))
        fo.write(sTab[2] + '<FILECOMMENTS> %s </FILECOMMENTS>'%(s2))
        fo.write(sTab[1] + '</SYSTEMPARAMS>')
        for data in xmlDict['OLXDBTABLE']:
            s1 = '<OLXDBTABLE NAME="%s" RECCOUNT="%s">'%(data['@NAME'], data['@RECCOUNT'])
            fo.write(sTab[1] +s1)
            olxtable = data['OLXREC']
            if type(olxtable) != list:
                olxtable = [data['OLXREC']]

            for olxrec in olxtable:
                fo.write(sTab[2] +'<OLXREC OBJTYPE="%s" OLNETID="%s" OBJGUID="%s">'%(olxrec['@OBJTYPE'],olxrec['@OLNETID'], olxrec['@OBJGUID']))
                fo.write(sTab[3] +'<OLNET>')
                for olnetfield in olxrec['OLNET']['OLNETFIELD']:
                    fo.write(sTab[4] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(olnetfield['@NAME'],olnetfield['@VALUE']))
                fo.write(sTab[3] +'</OLNET>')

                try:
                    if type(olxrec['DATAFIELD']) == dict:
                        df = olxrec['DATAFIELD']
                        if df['@NAME'] == "BK_OBJLST1" or df['@NAME'] == "BK_OBJLST2":
                                fo.write(sTab[3] +'<DATAFIELD NAME="%s">'%(df['@NAME']))
                                try:
                                    OLRelements = df['VALUE']['OLNET']
                                    if OLRelements != None:
                                        fo.write(sTab[4] +'<VALUE>')
                                        for elem in OLRelements:
                                            typeOLR = elem['@OBJTYPE']
                                            fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                            OBJLST1 = elem['OLNETFIELD']
                                            for obj1 in  OBJLST1:
                                                  fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                            fo.write(sTab[5] +'</OLNET>')
                                        fo.write(sTab[4] +'</VALUE>')
                                except:
                                    fo.write(sTab[4] +'<VALUE/>')
                                fo.write(sTab[3] +'</DATAFIELD>')
                        else:
                            fo.write(sTab[3] +'<DATAFIELD VALUE="%s" NAME="%s"/>'%(df['@VALUE'],df['@NAME']))
                    elif type(olxrec['DATAFIELD']) == list:
                        for df in olxrec['DATAFIELD']:
                            if df['@NAME'] == "BK_OBJLST1" or df['@NAME'] == "BK_OBJLST2":
                                fo.write(sTab[3] +'<DATAFIELD NAME="%s">'%(df['@NAME']))
                                try:
                                    OLRelements = df['VALUE']['OLNET']
                                    if type(OLRelements) != list:
                                        OLRelements = [df['VALUE']['OLNET']]
                                    if OLRelements != None:
                                        fo.write(sTab[4] +'<VALUE>')
                                        for elem in OLRelements:
                                            typeOLR = elem['@OBJTYPE']
                                            fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                            OBJLST1 = elem['OLNETFIELD']
                                            for obj1 in  OBJLST1:
                                                  fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                            fo.write(sTab[5] +'</OLNET>')
                                        fo.write(sTab[4] +'</VALUE>')
                                except:
                                    fo.write(sTab[4] +'<VALUE/>')
                                fo.write(sTab[3] +'</DATAFIELD>')
                            else:
                                fo.write(sTab[3] +'<DATAFIELD VALUE="%s" NAME="%s"/>'%(df['@VALUE'],df['@NAME']))

                except:
                    pass
                fo.write(sTab[2]+'</OLXREC>')
            fo.write(sTab[1]+'</OLXDBTABLE>')

        fo.write('\n</ASPENOLXDB>')
        fo.close()

        return fo

    def exportData2XML(self, data_olx, folx):
        objlist = []
        count = 0
        s0 = ''
        try:
            for key, value in data_olx.items():
                if len(value) > 0:
                    s0 += '%s="%s" ' %(str(key),str(len(value)))
                    if key!= 'COUNT':
                        objlist.append(key)
                    count += len(value)
        except:
            pass

        s1 = ''
        s2 = ''
        try:
            systemparams = self.xmlDict['SYSTEMPARAMS']
            for key, value in systemparams.items():
                if key != "FILECOMMENTS":
                    s1 += '%s="%s" ' %(str(key[1:]),str(value))
                else:
                    s2 = value
        except:
            pass



        xmlDict = self.xmlDict
        fo = open(folx, "w", encoding="utf-8")
        fo.write("<?xml version='1.0'?>\n")
        fo.write('<ASPENOLXDB OLRVERSION="%s" DATETIME="%s">'%(xmlDict['@OLRVERSION'], xmlDict['@DATETIME']))
        fo.write(sTab[1]+'<OBJCOUNT COUNT="%s" %s/>'%(count,s0))
        try:
            udftemplate = xmlDict['UDFTEMPLATE']
            fo.write(sTab[1]+'<UDFTEMPLATE>')
            if type(udftemplate['OLRXOBJ']) == list:
                for olrxobj in udftemplate['OLRXOBJ']:
                    fo.write(sTab[2]+'<OLRXOBJ OBJTYPE="%s">'%(olrxobj['@OBJTYPE']))
                    for udf in olrxobj['UDFIELD']:
                        fo.write(sTab[3]+'<UDFIELD ROWNO="%s" FNAME="%s" LABEL="%s"/>'%(udf['@ROWNO'],udf['@FNAME'],udf['@LABEL']))
                    fo.write(sTab[2]+'</OLRXOBJ>')
            elif type(udftemplate['OLRXOBJ']) == dict:
                olrxobj = udftemplate['OLRXOBJ']
                fo.write(sTab[2]+'<OLRXOBJ OBJTYPE="%s">'%(olrxobj['@OBJTYPE']))
                for udf in olrxobj['UDFIELD']:
                    fo.write(sTab[3]+'<UDFIELD ROWNO="%s" FNAME="%s" LABEL="%s"/>'%(udf['@ROWNO'],udf['@FNAME'],udf['@LABEL']))
                fo.write(sTab[2]+'</OLRXOBJ>')


            fo.write(sTab[1]+'</UDFTEMPLATE>')
        except:
            pass


        fo.write(sTab[1] + '<SYSTEMPARAMS %s>'%(s1))
        fo.write(sTab[2] + '<FILECOMMENTS> %s </FILECOMMENTS>'%(s2))
        fo.write(sTab[1] + '</SYSTEMPARAMS>')


        for data in xmlDict['OLXDBTABLE']:
            if data['@NAME'] not in objlist:
                continue
            dataobjlist = data_olx[data['@NAME']]
            ind = 0
            s1 = '<OLXDBTABLE NAME="%s" RECCOUNT="%s">'%(data['@NAME'], len(dataobjlist))
            fo.write(sTab[1] +s1)

            olxtable = data['OLXREC']
            if type(olxtable) != list:
                olxtable = [data['OLXREC']]
            for olxrec in olxtable:
                if ind == len(dataobjlist):
                    break
                dataobj = dataobjlist[ind]
                if dataobj['OBJGUID'] != olxrec['@OBJGUID']:
                    continue


                fo.write(sTab[2] +'<OLXREC OBJTYPE="%s" OLNETID="%s" OBJGUID="%s">'%(olxrec['@OBJTYPE'],olxrec['@OLNETID'], olxrec['@OBJGUID']))
                fo.write(sTab[3] +'<OLNET>')
                for olnetfield in olxrec['OLNET']['OLNETFIELD']:
                    fo.write(sTab[4] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(olnetfield['@NAME'],olnetfield['@VALUE']))
                fo.write(sTab[3] +'</OLNET>')

                try:
                    if type(olxrec['DATAFIELD']) == dict:
                        df = olxrec['DATAFIELD']
                        if df['@NAME'] == "BK_OBJLST1" or df['@NAME'] == "BK_OBJLST2":
                                fo.write(sTab[3] +'<DATAFIELD NAME="%s">'%(df['@NAME']))
                                try:
                                    OLRelements = dataobj[df['@NAME']]['OLNET']
                                    if OLRelements != None:
                                        fo.write(sTab[4] +'<VALUE>')
                                        for elem in OLRelements:
                                            typeOLR = elem['@OBJTYPE']
                                            fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                            OBJLST1 = elem['OLNETFIELD']
                                            for obj1 in  OBJLST1:
                                                  fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                            fo.write(sTab[5] +'</OLNET>')
                                        fo.write(sTab[4] +'</VALUE>')
                                except:
                                    fo.write(sTab[4] +'<VALUE/>')
                                fo.write(sTab[3] +'</DATAFIELD>')
                        else:
                            fo.write(sTab[3] +'<DATAFIELD VALUE="%s" NAME="%s"/>'%(dataobj[df['@NAME']],df['@NAME']))

                    elif type(olxrec['DATAFIELD']) == list:
                        for df in olxrec['DATAFIELD']:
                            if df['@NAME'] == "BK_OBJLST1" or df['@NAME'] == "BK_OBJLST2":
                                fo.write(sTab[3] +'<DATAFIELD NAME="%s">'%(df['@NAME']))
                                try:
                                    OLRelements = dataobj[df['@NAME']]['OLNET']
                                    if type(OLRelements) != list:
                                        OLRelements = [dataobj[df['@NAME']]['OLNET']]
                                    if OLRelements != None:
                                        fo.write(sTab[4] +'<VALUE>')
                                        for elem in OLRelements:
                                            typeOLR = elem['@OBJTYPE']
                                            fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                            OBJLST1 = elem['OLNETFIELD']
                                            for obj1 in  OBJLST1:
                                                  fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                            fo.write(sTab[5] +'</OLNET>')
                                        fo.write(sTab[4] +'</VALUE>')
                                except:
                                    fo.write(sTab[4] +'<VALUE/>')
                                fo.write(sTab[3] +'</DATAFIELD>')
                            else:
                                fo.write(sTab[3] +'<DATAFIELD VALUE="%s" NAME="%s"/>'%(dataobj[df['@NAME']],df['@NAME']))

                except:
                    pass
                fo.write(sTab[2]+'</OLXREC>')
                ind += 1
            fo.write(sTab[1]+'</OLXDBTABLE>')


        fo.write('\n</ASPENOLXDB>')
        fo.close()

        return fo


    def get_DATA(self,filterOptions=None,
                 typ=[],                # list of objtype ['BUS','LINE']
                 area=[], inside_area = True, # list of area number
                 zone=[], inside_zone = True, # list of zone number
                 tie=False,                   # include tie between selected items and the rest of the network
                 kv=[],                       # [kvmin, kvmax]
                 BNO=[],                      # list of bus number
                 OBJGUID=[],                 # list of objguid
                 ckid = [],
                 filter_data_out = False):
        """
        Get objects data from XML file (OLR objects)

        Parameters
        ----------
        filterOptions : object, optional
            Filter Option object to filter area, zone, kv, BNO.... The default is None.
        filter_data_out : boolean, optional
            True to get the data filtered, False to get full data. The default is False.

        Returns
        -------
        dict
            Full data from XML or filtered data in dictionary format, key = OBJTYPE, value = list of object parameters.
        dict
            Dictionary format, key = OBJTYPE, value = list of True/False value for each object.

        """

        checkrange = False
        if filterOptions != None:
            typ = filterOptions.typ
            area = filterOptions.area
            inside_area = filterOptions.inside_area
            zone = filterOptions.zone
            inside_zone = filterOptions.inside_zone
            tie = filterOptions.tie
            kv = filterOptions.kv
            BNO = filterOptions.BNO
            OBJGUID = filterOptions.OBJGUID
            ckid = filterOptions.ckid
            if len(filterOptions.kv) == 2:
                checkrange = True


        if not checkrange:
            arearange, zonerange, kvrange, BNOrange, CKTIDrange, checkrange =  get_filter_range(area, zone, kv, BNO, ckid)
        else:
            arearange, zonerange, kvrange, BNOrange, CKTIDrange, _ =  get_filter_range(area, zone, kv, BNO, ckid)


        olnetKV = ['TERMKV1', 'TERMKV2', 'TERMKV3','TERMKV4']
        olnetBNO = ['TERMBNO1','TERMBNO2','TERMBNO3','TERMBNO4']
        olnetTERMGUID = ['TERMGUID1','TERMGUID2','TERMGUID3','TERMGUID4']
        olnetCKTID = ['CKTID','CKTID2']
        olnetAREA = ['BS_AREANO','AR_NO']
        olnetZONE = ['BS_ZONENO','ZN_NO']
        dictBUS_AR = dict()
        dictBUS_ZN = dict()

        olxdbtables = self.xmlDict['OLXDBTABLE']
        self.data_olx = dict()
        self.data_filtered = dict()

        for olxtable in olxdbtables:
            objtype = olxtable['@NAME']
            objcount = int(olxtable['@RECCOUNT'])
            if objcount > 0:
                olxrecs = olxtable['OLXREC']
            if type(olxrecs) != list:
                olxrecs = [olxtable['OLXREC']]

            if (objtype in OBJTYPELIST) and (objcount > 0):
                data_objs = []
                data_filtered = []

                for olxrec in olxrecs:
                    data_elem = dict()
                    ARfilter = False
                    ZNfilter = False
                    KVfilter = False
                    BNOfilter = False
                    CKTIDfilter = False
                    ar = []
                    zn = []
                    bno = []
                    kvmax = []
                    cktid = []
                    area_BUS = []
                    zone_BUS = []

                    for obj in olxrec['OLNET']['OLNETFIELD'] :
                        data_elem[obj['@NAME']] = obj['@VALUE']

                    if olxrec.__contains__('DATAFIELD'):
                        for data in olxrec['DATAFIELD']:
                            if ('@NAME' and '@VALUE') in data.keys():
                                data_elem[data['@NAME']] = data['@VALUE']
                            elif ('@NAME' and 'VALUE') in data.keys():
                                data_elem[data['@NAME']] = data['VALUE']

                                # BUS should be first in data OLX
                                if data['@NAME'] == 'BS_AREANO':
                                    dictBUS_AR[data_elem['OBJGUID']] = int(data['@VALUE'])
                                if data['@NAME'] == 'BS_ZONENO':
                                    dictBUS_ZN[data_elem['OBJGUID']] = int(data['@VALUE'])

                    # LOOP for filter
                    for data in olxrec['OLNET']['OLNETFIELD']:
                        if data['@NAME'] in olnetKV:
                            if checkrange:
                                if (float(data['@VALUE']) >= kvrange[0]) and (float(data['@VALUE']) <= kvrange[1]):
                                    kvmax.append(True)
                                else:
                                    kvmax.append(False)
                            else:
                                if float(data['@VALUE']) in kvrange:
                                    kvmax.append(True)
                                else:
                                    kvmax.append(False)

                        if data['@NAME'] in olnetBNO:
                            if int(data['@VALUE']) in BNOrange:
                                bno.append(True)
                            else:
                                bno.append(False)

                        if data['@NAME'] in olnetCKTID:
                            if (data['@VALUE']) in CKTIDrange:
                                cktid.append(True)
                            else:
                                cktid.append(False)

                    if (objtype == 'BUS') or (objtype == 'ZONE') or (objtype == 'AREA'):
                        for data in olxrec['DATAFIELD']:
                            if data['@NAME'] in olnetAREA:
                                if (int(data['@VALUE']) in arearange) == inside_area:
                                    ar.append(True)
                                else:
                                    ar.append(False)

                            if data['@NAME'] in olnetZONE:
                               if (int(data['@VALUE']) in zonerange) == inside_zone:
                                   zn.append(True)
                               else:
                                   zn.append(False)

                    if (objtype != 'BUS') and (objtype != 'ZONE'):
                        for term in olnetTERMGUID:
                            if term in data_elem.keys():
                                if data_elem[term] in dictBUS_ZN.keys():
                                    zone_BUS.append(dictBUS_ZN[data_elem[term]])
                        for z1 in zone_BUS:
                            if (z1 in zonerange) == inside_zone:
                                zn.append(True)
                            else:
                                zn.append(False)

                    if (objtype != 'BUS') and (objtype != 'AREA'):
                        for term in olnetTERMGUID:
                            if term in data_elem.keys():
                                if data_elem[term] in dictBUS_AR.keys():
                                    area_BUS.append(dictBUS_AR[data_elem[term]])

                        for a1 in area_BUS:
                            if (a1 in arearange) == inside_area:
                                ar.append(True)
                            else:
                                ar.append(False)


                    if typ != [] and typ != None:
                        typ_res = objtype in typ
                    else:
                        typ_res = True

                    if OBJGUID != [] and OBJGUID != None:
                        objguid_res = data_elem['OBJGUID'] in OBJGUID
                    else:
                        objguid_res = True

                    if kvrange == [] or objtype == 'AREA' or objtype == 'ZONE' or objtype == 'ZCORRECT':
                        KVfilter = True
                    else:
                        try:
                            KVfilter = min(kvmax)
                        except:
                            KVfilter = True

                    if BNOrange == []:
                        BNOfilter = True
                    else:
                        BNOfilter = max(bno)

                    if CKTIDrange ==[]:
                        CKTIDfilter = True
                    else:
                        if len(cktid) > 0:
                            CKTIDfilter = max(cktid)

                    if arearange ==[] or objtype == 'ZONE' or objtype == 'ZCORRECT':
                        ARfilter = True
                    else:
                        if tie:
                            if len(ar) > 0:
                                ARfilter = max(ar)
                            else:
                                ARfilter = False
                        else:
                            if len(ar) > 0:
                                ARfilter = min(ar)
                            else:
                                ARfilter = False

                    if zonerange ==[] or objtype == 'AREA' or objtype == 'ZCORRECT':
                        ZNfilter = True
                    else:
                        if tie:
                            if len(zn) >0:
                                ZNfilter = max(zn)
                            else:
                                ZNfilter = False
                        else:
                            if len(zn) >0:
                                ZNfilter = min(zn)
                            else:
                                ZNfilter = False

                    result = ARfilter & ZNfilter & KVfilter & BNOfilter & CKTIDfilter & typ_res & objguid_res
                    data_filtered.append(result)
                    if filter_data_out:
                        if result:
                            data_objs.append(data_elem)
                    else:
                        data_objs.append(data_elem)

                self.data_olx[objtype] = data_objs
                self.data_filtered[objtype] =  data_filtered

        return self.data_olx, self.data_filtered
#
class DataASPEN_ADX:
    """
    read ASPEN ADX file
    """
    def __init__(self,fadx,filterOptions=None,encoding='utf-8',olxpath=''):
        self.fadx = fadx
        self.encoding = encoding
        self.filterOptions = filterOptions

        self.currentFileOpened = ''
        self.act0_RELAY = ''
        self.olxpath = olxpath
        self.valSum = {}
        #
        self.act0 = ''
        self.obj0 = ''
        self.olnets = ''

        self.xmlDict = None
    #
    def getAll_CHANGEREC(self):
        ADXdict = xml2Dict(self.fadx,desc = 'Parsing ASPEN ADX',tagStop=None,encoding=self.encoding)
        try:
            self.xmlDict = ADXdict['ASPENOLX']['OLXDIFF']
        except:
            raise Exception("\n\nInput file format required: ASPEN OLR DIFF XML")
        #
        try:
            if type(self.xmlDict['CHANGEREC']) !=list:
                self.xmlDict['CHANGEREC'] = [self.xmlDict['CHANGEREC']]
        except:
            self.xmlDict['CHANGEREC'] = []
    #
    def getValSum(self):
        source = open(self.fadx , 'r',encoding=self.encoding,errors = 'ignore')
        #
        parser = CET.XMLParser(encoding=self.encoding)
        # Start iterating over the Element Tree.
        context = CET.iterparse(source, events=('start', 'end'), parser = parser)
        for event, elem in context:
            if event == 'start' and elem.tag=='CHANGEREC':
                break
            if event == 'start' and ( elem.tag =='OLXDIFF' or elem.tag=='CHANGESTAT') :
                self.valSum.update(elem.attrib)
        source.close()
        #
        for objtype in OBJTYPELIST:
            self.valSum[objtype] = {}
            for action in ACTION_ADX:
                self.valSum[objtype][action] = 0
    #
    def getValFA(self):
        self.fA = self.valSum['FILEA']
        if not os.path.isfile(self.fA):
            self.fA = os.path.dirname(self.fadx) +'\\' + os.path.basename(self.fA)
        #
        if not os.path.isfile(self.fA):
            raise Exception ('\nADX: '+self.fadx +'\n\n'+'Path name of File A not found:\n'+self.valSum['FILEA'] +'\nor:\n'+self.fA)
        #
        self.fA_short = AppUtils.getShortNameFile(self.fA,27)
        self.valSumA =  getValSum(self.fA,self.olxpath)
    #
    def getValFB(self):
        self.fB = self.valSum['FILEB']
        if not os.path.isfile(self.fB):
            self.fB = os.path.dirname(self.fadx) +'\\' + os.path.basename(self.fB)
        #
        if not os.path.isfile(self.fB):
            raise Exception ('\nADX: '+self.fadx +'\n\n'+'Path name of File B not found:\n'+self.valSum['FILEB'] +'\nor:\n'+self.fB)
        #
        self.fB_short = AppUtils.getShortNameFile(self.fB,27)
        self.valSumB =  getValSum(self.fB,self.olxpath)

    #
    def run1CHANGEREC_txt(self,cr1):
        if self.ftxt=='':
            return
        objtype = cr1['@OBJTYPE']
        action = cr1['@ACTION']
        olnet = cr1['OLNET']
        self.valSum[objtype][action] += 1
        #
        try:
            chf = cr1['CHANGEFIELD']
        except:
            chf = None
        #
        sadd1 = action.ljust(8)+ dict_OBJ[objtype][0]
        if sadd1=='Relay group':
            sadd1 +=' on'
        elif sadd1.startswith('Load Tap Changer'):
            sadd1 +='\n           '
        #
        d1 = Data1ADX(olnet,chf,action)
        #
        s1 = sadd1+ d1.getString()

        if action=='ADD':
            s1 += '\n  is in [B]'+self.fB_short+' but not in [A]' + self.fA_short+'.'
        elif action=='DELETE':
            s1 += '\n  is in [A]'+self.fA_short+' but not in [B]' + self.fB_short+'.'
        elif action=='MODIFY':
            s1 +='\n                 the following parameter(s) are different:'
            s1 +='\n _Parameter_______________________[A]'+self.fA_short  +'  [B]'+self.fB_short
            s1 += d1.getChange()

        self.fftxt.write('\n\n'+s1)
    #
    def export_dxtrecs(self):
        if self.olnets=='':
            return
        #
        action = self.act0
        objtype = self.obj0
        #
        if objtype not in ['RLYOC', 'FUSE', 'RLYDS', 'RLYD', 'RLYV', 'RECLSR', 'SCHEME']:
            if objtype=='COORDPAIR':
                header = '\n[ADD COORDINATION PAIRS DATA]\n'
            else:
                header = '\n[{} {}]\n'.format(action, dict_OBJ[objtype][1] )
            self.act0_RELAY = ''
        else:
            header = ''
            if action!=self.act0_RELAY:
                header = '\n[%s RELAY]\n'%action
            #
            self.act0_RELAY = action
        #
        if action == 'DELETE':
            if self.currentFileOpened != self.fA:
                OlxAPILib.open_olrFile(self.fA,self.olxpath,prt=False)
                self.currentFileOpened = self.fA
        else:
            if self.currentFileOpened != self.fB:
                OlxAPILib.open_olrFile(self.fB,self.olxpath,prt=False)
                self.currentFileOpened = self.fB
        #
        if header:
            f = open(self.fchf, "a+")
            f.write(header)
            f.close()
        # #DXTRECS ANAFASRECS
        inputParams = '<EXPORTNETWORK ' \
                        'FORMAT="DXTRECS" APPENDOUTPUT ="YES" ' + \
                        'OUTPUTPATHNAME="' + self.fchf +'"' + ' >' + self.olnets +\
                        '</EXPORTNETWORK>'
        #
        OlxAPILib.run1LPFCommand(inputParams)
    #
    def run1CHANGEREC_chf(self,cr1):
        if self.fchf=='':
            return
        #
        objtype = cr1['@OBJTYPE']
        action = cr1['@ACTION']
        olnet = cr1['OLNET']
        #
        olnet_str = "<OLNET OBJTYPE="'"{}"'"> ".format(olnet['@OBJTYPE'])
        for olnetfield in olnet['OLNETFIELD'] :
            olnet_str += "<OLNETFIELD NAME="'"{}"'" VALUE="'"{}"'"/>".format(olnetfield['@NAME'], olnetfield['@VALUE'])
        olnet_str += "</OLNET>"

        #
        if (self.act0==action and self.obj0==objtype): # continue
            self.olnets += olnet_str
            if objtype=='COORDPAIR':
                self.olnets = olnet_str
                self.export_dxtrecs()
        else :
            self.export_dxtrecs()
            #
            self.olnets = olnet_str
            #
            self.act0 = action
            self.obj0 = objtype
    #
    def run1CHANGEREC(self,cr1):
        if not self.filterOptions.runFilter(cr1):
            return
        self.run1CHANGEREC_txt(cr1)
        #
        self.run1CHANGEREC_chf(cr1)
    #
    def runSuccesive(self,ftxt,fchf):
        if ftxt=='' and fchf=='':
            return
        #
        self.ftxt = ftxt
        self.fchf = fchf
        if ftxt:
            self.fftxt = open(ftxt, "a+")

        # get total CHANGEREC
        source = open(self.fadx , 'rb')
        contents = source.read()
        total = contents.count(b'<CHANGEREC ')
        source.close()
        #
        source = open(self.fadx , 'r',encoding=self.encoding,errors = 'ignore') #,errors = 'replace'
        #
        parser = CET.XMLParser(encoding=self.encoding)
        # Start iterating over the Element Tree.
        context = CET.iterparse(source, events=('start', 'end'), parser = parser)

        # This is the output dict.
        output = {}

        # Keeping track of the depth and position to store data in.
        current_position = []
        current_index = []
        #
        flag = False
        #
        with tqdm(total=total,unit=' CHANGEREC',ncols = 80,position=0,desc = 'Processing') as pbar:
            for event, elem in context:
                if event == 'start' and elem.tag=='CHANGEREC':
                    if output:
                        self.run1CHANGEREC(output['CHANGEREC'])
                        pbar.update(1)
                        output = {}
                    flag = True
                #
                elif event == 'end' and elem.tag=='OLXDIFF':
                    if output:
                        self.run1CHANGEREC(output['CHANGEREC'])
                        pbar.update(1)
                    break
                #
                if flag:
                    if event == 'start':# Start of new tag.
                        # Extract the current endpoint so add the new element to it.
                        tmp = output
                        for cp, ci in zip(current_position, current_index):
                            tmp = tmp[cp]
                            if ci:
                                tmp = tmp[ci]

                        this_tag_name = elem.tag
                        # If it is a previously unseen tag, create a new key and
                        # stick an empty dict there. Set index of this level to None.
                        if this_tag_name not in tmp:
                            tmp[this_tag_name] = {}
                            current_index.append(None)
                        else:
                            # The tag name already exists. This means that we have to change
                            # the value of this element's key to a list if this hasn't
                            # been done already and add an empty dict to the end of that
                            # list. If it already is a list, just add an new dict and update
                            # the current index.
                            if isinstance(tmp[this_tag_name], list):
                                current_index.append(len(tmp[this_tag_name]))
                                tmp[this_tag_name].append({})
                            else:
                                tmp[this_tag_name] = [tmp[this_tag_name], {}]
                                current_index.append(1)
                        # Set the position of the iteration to this element's tag name.
                        current_position.append(this_tag_name)
                    elif event == 'end': # End of a tag.
                        # Extract the current endpoint's parent so we can handle
                        # the endpoint's data by reference.
                        tmp = output
                        for cp, ci in zip(current_position[:-1], current_index[:-1]):
                            tmp = tmp[cp]
                            if ci:
                                tmp = tmp[ci]
                        cp = current_position[-1]
                        ci = current_index[-1]
                        # If this current endpoint is a dict in a list or not has
                        # implications on how to set data.
                        if ci:
                            setfcn = lambda x: setitem(tmp[cp], ci, x)
                            for attr_name, attr_value in elem.attrib.items():
                                tmp[cp][ci]["@{0}".format(attr_name)] = attr_value
                        else:
                            setfcn = lambda x: setitem(tmp, cp, x)
                            for attr_name, attr_value in elem.attrib.items():
                                tmp[cp]["@{0}".format(attr_name)] = attr_value

                        # If there is any text in the tag, add it here.
                        if elem.text and elem.text.strip():
                            setfcn({'#text': elem.text.strip()})

                        # Handle special cases:
                        # 1) when the tag only harbours text, replace the dict content with
                        #    that very text string.
                        # 2) when no text, attributes or children are present, content
                        #    is set to None
                        # These are detailed in reference [3] in README.
                        if ci:
                            nk = len(tmp[cp][ci].keys())
                            if nk == 1 and "#text" in tmp[cp][ci]:
                                tmp[cp][ci] = tmp[cp][ci]["#text"]
                            elif nk == 0:
                                tmp[cp][ci] = None
                        else:
                            nk = len(tmp[cp].keys())
                            if nk == 1 and "#text" in tmp[cp]:
                                tmp[cp] = tmp[cp]["#text"]
                            elif nk == 0:
                                tmp[cp] = None

                        # Remove the outermost position and index, since we just finished
                        # handling that element.
                        current_position.pop()
                        current_index.pop()

                        # Most important of all, release the element's memory allocations
                        # so we actually benefit from the iterative processing!
                        elem.clear()
                        # while elem.getprevious() is not None:
                        #     del elem.getparent()[0]
        #
        source.close()
        #
        self.export_dxtrecs()
        #
        if self.ftxt:
            self.fftxt.close()
        #
        OlxAPI.UnloadOlxAPI()


    def export2XML(self, fadx):
        info_dict = {}
        options = ['OLRVERSION','DATETIME','FILEA','FILEB',
                   'COMP_BUSES', 'COMP_AREAZONES', 'COMP_GENS', 'COMP_LOADS', 'COMP_SHUNTS', 'COMP_SVDS',
                   'COMP_LINES', 'COMP_SHIFTERS', 'COMP_XFMRS', 'COMP_BREAKERS', 'COMP_MUS', 'COMP_OCRLYS',
                   'COMP_DSRLYS', 'COMP_SWITCHES', 'COMP_DCLINE2S', 'COMP_SCHEMES', 'COMP_DEVICES',
                   'EXTENT', 'CORRELATE', 'TIES', 'AZRANGE', 'KVRANGE',
                    'BOUNDARYID', 'TAGS', 'CHANGECOUNT']

        for key in options:
            try:
                info_dict[key] = self.xmlDict['@'+key]
            except:
                if key == 'FILEA' or key == 'FILEB':
                    info_dict[key] = ""
                elif key == 'OLRVERSION':
                    info_dict[key] = "15.5"
                else:
                    pass

        s0 = ''
        for key, value in info_dict.items():
            s0 += '%s="%s" ' %(str(key),str(value))

        s1 = ''
        try:
            changestat = self.xmlDict['CHANGESTAT']
            for key, value in changestat.items():
                s1 += '%s="%s" ' %(str(key[1:]),str(value))
        except:
            pass

        xmlDict = self.xmlDict
        fo = open(fadx, "w", encoding="utf-8")
        fo.write("<?xml version='1.0'?>\n")
        fo.write('<ASPENOLX>')
        fo.write(sTab[1]+'<OLXDIFF %s>'%(s0))
        fo.write(sTab[2] + '<CHANGESTAT %s/>'%(s1))
        for changerec in xmlDict['CHANGEREC']:
            olnet = changerec['OLNET']

            s1 = '<CHANGEREC ACTION="%s" OBJTYPE="%s" OLNETID="%s" OBJGUID="%s">'%(changerec['@ACTION'],changerec['@OBJTYPE'],changerec['@OLNETID'], changerec['@OBJGUID'])
            fo.write(sTab[2] +s1)
            fo.write(sTab[3] +'<OLNET OBJTYPE="%s">'% changerec['@OBJTYPE'])
            for olnetfield in olnet['OLNETFIELD']:
                fo.write(sTab[4] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(olnetfield['@NAME'],olnetfield['@VALUE']))
            fo.write(sTab[3] +'</OLNET>')
            try:
                objscope = changerec['OBJSCOPE']
                fo.write(sTab[3] +'<OBJSCOPE>')
                for scope in objscope['SCOPEFIELD']:
                    fo.write(sTab[4] +'<SCOPEFIELD NAME="%s" VALUE="%s"/>'%(scope['@NAME'],scope['@VALUE']))
                fo.write(sTab[3] +'</OBJSCOPE>')
            except:
                pass
            try:
                if type(changerec['CHANGEFIELD']) == dict:
                    chf = changerec['CHANGEFIELD']
                    if chf['@NAME'] == "BK_OBJLST1" or chf['@NAME'] == "BK_OBJLST2":
                        if changerec['@ACTION'] == "MODIFY":
                            fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s">'%(chf['@LABEL'],chf['@NAME']))
                            OLRelements = chf['VALUEA']['OLNET']
                            if OLRelements != None:
                                fo.write(sTab[4] +'<VALUEA>')
                                for elem in OLRelements:
                                    typeOLR = elem['@OBJTYPE']
                                    fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                    OBJLST1 = elem['OLNETFIELD']
                                    for obj1 in  OBJLST1:
                                         fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                    fo.write(sTab[5] +'</OLNET>')
                                fo.write(sTab[4] +'</VALUEA>')
                            else:
                                fo.write(sTab[4] +'<VALUEA/>')

                            OLRelements = chf['VALUEB']['OLNET']
                            if OLRelements != None:
                                fo.write(sTab[4] +'<VALUEB>')
                                for elem in OLRelements:
                                    typeOLR = elem['@OBJTYPE']
                                    fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                    OBJLST2 = elem['OLNETFIELD']
                                    for obj2 in  OBJLST2:
                                         fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj2['@NAME'], obj2['@VALUE']))
                                    fo.write(sTab[5] +'</OLNET>')
                                fo.write(sTab[4] +'</VALUEB>')
                            else:
                                fo.write(sTab[4] +'<VALUEB/>')
                            fo.write(sTab[3] +'</CHANGEFIELD>')
                        elif changerec['@ACTION'] == "ADD":
                            fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s">'%(chf['@LABEL'],chf['@NAME']))
                            try:
                                OLRelements = chf['VALUE']['OLNET']
                                if OLRelements != None:
                                    fo.write(sTab[4] +'<VALUE>')
                                    for elem in OLRelements:
                                        typeOLR = elem['@OBJTYPE']
                                        fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                        OBJLST1 = elem['OLNETFIELD']
                                        for obj1 in  OBJLST1:
                                             fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                        fo.write(sTab[5] +'</OLNET>')
                                    fo.write(sTab[4] +'</VALUE>')
                            except:
                                fo.write(sTab[4] +'<VALUE/>')
                            fo.write(sTab[3] +'</CHANGEFIELD>')
                    else:
                        if changerec['@ACTION'] == "MODIFY":
                            fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s" VALUEA="%s" VALUEB="%s"/>'%(chf['@LABEL'],chf['@NAME'],chf['@VALUEA'],chf['@VALUEB']))
                        if changerec['@ACTION'] == "ADD":
                            fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s" VALUE="%s"/>'%(chf['@LABEL'],chf['@NAME'],chf['@VALUE']))

                elif type(changerec['CHANGEFIELD']) == list:
                    for chf in changerec['CHANGEFIELD']:
                        if chf['@NAME'] == "BK_OBJLST1" or chf['@NAME'] == "BK_OBJLST2":
                            if changerec['@ACTION'] == "MODIFY":
                                fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s">'%(chf['@LABEL'],chf['@NAME']))
                                OLRelements = chf['VALUEA']['OLNET']
                                if type(OLRelements) != list:
                                    OLRelements = [chf['VALUEA']['OLNET']]
                                if OLRelements != None:
                                    fo.write(sTab[4] +'<VALUEA>')
                                    for elem in OLRelements:
                                        typeOLR = elem['@OBJTYPE']
                                        fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                        OBJLST1 = elem['OLNETFIELD']
                                        for obj1 in  OBJLST1:
                                             fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                        fo.write(sTab[5] +'</OLNET>')
                                    fo.write(sTab[4] +'</VALUEA>')
                                else:
                                    fo.write(sTab[4] +'<VALUEA/>')

                                OLRelements = chf['VALUEB']['OLNET']
                                if type(OLRelements) != list:
                                    OLRelements = [chf['VALUEB']['OLNET']]
                                if OLRelements != None:
                                    fo.write(sTab[4] +'<VALUEB>')
                                    for elem in OLRelements:
                                        typeOLR = elem['@OBJTYPE']
                                        fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                        OBJLST2 = elem['OLNETFIELD']
                                        for obj2 in  OBJLST2:
                                             fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj2['@NAME'], obj2['@VALUE']))
                                        fo.write(sTab[5] +'</OLNET>')
                                    fo.write(sTab[4] +'</VALUEB>')
                                else:
                                    fo.write(sTab[4] +'<VALUEB/>')
                                fo.write(sTab[3] +'</CHANGEFIELD>')
                            elif changerec['@ACTION'] == "ADD":
                                fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s">'%(chf['@LABEL'],chf['@NAME']))
                                try:
                                    OLRelements = chf['VALUE']['OLNET']
                                    if type(OLRelements) != list:
                                        OLRelements = [chf['VALUE']['OLNET']]
                                    if OLRelements != None:
                                        fo.write(sTab[4] +'<VALUE>')
                                        for elem in OLRelements:
                                            typeOLR = elem['@OBJTYPE']
                                            fo.write(sTab[5] +'<OLNET OBJTYPE="%s">'%typeOLR)
                                            OBJLST1 = elem['OLNETFIELD']
                                            for obj1 in  OBJLST1:
                                                 fo.write(sTab[6] +'<OLNETFIELD NAME="%s" VALUE="%s"/>'%(obj1['@NAME'], obj1['@VALUE']))
                                            fo.write(sTab[5] +'</OLNET>')
                                        fo.write(sTab[4] +'</VALUE>')
                                except:
                                    fo.write(sTab[4] +'<VALUE/>')
                                fo.write(sTab[3] +'</CHANGEFIELD>')
                        else:
                            if changerec['@ACTION'] == "MODIFY":
                                fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s" VALUEA="%s" VALUEB="%s"/>'%(chf['@LABEL'],chf['@NAME'],chf['@VALUEA'],chf['@VALUEB']))
                            if changerec['@ACTION'] == "ADD":
                                fo.write(sTab[3] +'<CHANGEFIELD LABEL="%s" NAME="%s" VALUE="%s"/>'%(chf['@LABEL'],chf['@NAME'],chf['@VALUE']))
            except:
                pass

            fo.write(sTab[2]+'</CHANGEREC>')

        fo.write(sTab[1]+'</OLXDIFF>')
        fo.write('\n</ASPENOLX>')
        fo.close()

        return fo


    #
    def get_CHANGESTAT(self, action = None, objtype = None):
        """
        Get statistic of changes for object and action

        Parameters
        ----------
        action : string, optional
            Action in list: ADD, DELETE, MODIFY. The default is None.
        objtype : string, optional
            OLR object type. The default is None.

        Raises
        ------
        Exception
            Illegal input of action and object type.

        Returns
        -------
        dict
            Statistic of changes.

        """
        # return dict{}
        dictACTION = {'ADD':'ADD','DELETE':'DEL','MODIFY':'MOD'}

        if self.xmlDict == None:
            self.getAll_CHANGEREC()


        if  objtype in OBJTYPELIST:
            if action in ACTION_ADX:
                return self.xmlDict['CHANGESTAT']['@{}_{}'.format(dictACTION[action], objtype)]
            elif action == None:
                add = self.xmlDict['CHANGESTAT']['@ADD_{}'.format(objtype)]
                modify = self.xmlDict['CHANGESTAT']['@MOD_{}'.format(objtype)]
                delete = self.xmlDict['CHANGESTAT']['@DEL_{}'.format(objtype)]
                return {'OBJTYPE': objtype, 'CHANGECOUNT' :{'ADD':add, 'MOD':modify, 'DEL': delete}}
            else:
                raise Exception(" Action must be in {} and  object type in {}".format(ACTION_ADX, OBJTYPELIST))
        else:
            return self.xmlDict['CHANGESTAT']


    def get_OLX_DIFF_MODEL(self, olrmodel = None):
        """
        Get OLX DIFF MODEL

        Returns
        -------
        olx_diff_model : dict
            Dictionary of differences File A/ File B from ADX file.

        """

        if self.xmlDict == None:
            self.getAll_CHANGEREC()
        changerec = self.xmlDict['CHANGEREC']

        additions = {}
        deletions = {}
        modifications = {}
        reversemodifications = {}
        for cr in changerec:
            if cr['@ACTION'] == 'DELETE':
                data_obj = {}
                data_obj['OLNETID'] = cr['@OLNETID']
                olnetfield = cr['OLNET']['OLNETFIELD']
                if type(olnetfield) != list:
                    olnetfield = [cr['OLNET']['OLNETFIELD']]
                for oln in olnetfield:
                    try:
                        data_obj[oln['@NAME']] = oln['@VALUE']
                    except:
                        continue

                try:
                    scopefield = cr['OBJSCOPE']['SCOPEFIELD']
                    if type(scopefield) != list:
                        scopefield = [cr['OBJSCOPE']['SCOPEFIELD']]
                    for scope in scopefield:
                        try:
                            data_obj[scope['@NAME']] = scope['@VALUE']
                        except:
                            continue
                except:
                    pass

                data_obj = correctGUID(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel)
                # checkGUID(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel)
                if cr['@OBJTYPE'] == "LTC" or cr['@OBJTYPE'] == "LTC3":
                    deletions[genIDAspen()] = [cr['@OBJTYPE'],data_obj]
                else:
                    deletions[cr['@OBJGUID']] = [cr['@OBJTYPE'],data_obj]

            if cr['@ACTION'] == 'ADD':
                data_obj = {}
                data_obj['OLNETID'] = cr['@OLNETID']
                olnetfield = cr['OLNET']['OLNETFIELD']
                if type(olnetfield) != list:
                    olnetfield = [cr['OLNET']['OLNETFIELD']]
                for oln in olnetfield:
                    try:
                        data_obj[oln['@NAME']] = oln['@VALUE']
                    except:
                        continue

                try:
                    scopefield = cr['OBJSCOPE']['SCOPEFIELD']
                    if type(scopefield) != list:
                        scopefield = [cr['OBJSCOPE']['SCOPEFIELD']]
                    for scope in scopefield:
                        try:
                            data_obj[scope['@NAME']] = scope['@VALUE']
                        except:
                            continue
                except:
                    pass

                if type(cr['CHANGEFIELD']) != list:
                    changefield = [cr['CHANGEFIELD']]
                else:
                    changefield = cr['CHANGEFIELD']
                for cf in changefield:
                    try:
                        data_obj[cf['@NAME']] = cf['@VALUE']
                    except:
                        continue

                # dictGUID_FAB = getGUID_FA_FB(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel, dictGUID_FAB)
                data_obj = correctGUID(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel)
                # checkGUID(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel)
                if cr['@OBJTYPE'] == "LTC" or cr['@OBJTYPE'] == "LTC3":
                    additions[genIDAspen()] = [cr['@OBJTYPE'], data_obj]
                else:
                    additions[cr['@OBJGUID']] = [cr['@OBJTYPE'], data_obj]
                # print(cr['@OBJGUID'], cr['@OBJTYPE'], data_obj)
                # print('----------------------------------------------')
            if cr['@ACTION'] == 'MODIFY':
                data_objA = {}
                data_objA['OLNETID'] = cr['@OLNETID']
                data_objB = {}
                data_objB['OLNETID'] = cr['@OLNETID']

                olnetfield = cr['OLNET']['OLNETFIELD']
                if type(olnetfield) != list:
                    olnetfield = [cr['OLNET']['OLNETFIELD']]
                for oln in olnetfield:
                    try:
                        data_objA[oln['@NAME']] = oln['@VALUE']
                        data_objB[oln['@NAME']] = oln['@VALUE']
                    except:
                        continue

                try:
                    scopefield = cr['OBJSCOPE']['SCOPEFIELD']
                    if type(scopefield) != list:
                        scopefield = [cr['OBJSCOPE']['SCOPEFIELD']]
                    for scope in scopefield:
                        try:
                            data_objA[scope['@NAME']] = scope['@VALUE']
                            data_objB[scope['@NAME']] = scope['@VALUE']
                        except:
                            continue
                except:
                    pass

                if type(cr['CHANGEFIELD']) != list:
                    changefield = [cr['CHANGEFIELD']]
                else:
                    changefield = cr['CHANGEFIELD']
                for cf in changefield:
                    try:
                        data_objA[cf['@NAME']] = cf['@VALUEA']
                    except:
                        data_objA[cf['@NAME']] = cf['VALUEA']
                    try:
                        data_objB[cf['@NAME']] = cf['@VALUEB']
                    except:
                        data_objB[cf['@NAME']] = cf['VALUEB']

                data_objA = correctGUID(cr['@OBJTYPE'], cr['@ACTION'], data_objA, olrmodel)
                data_objB = correctGUID(cr['@OBJTYPE'], cr['@ACTION'], data_objB, olrmodel)
                # getGUID_FA_FB(cr['@OBJTYPE'], cr['@ACTION'], data_obj, olrmodel, dictGUID_FAB)
                # dictGUID_FAB = getGUID_FA_FB(cr['@OBJTYPE'], cr['@ACTION'], data_objB, olrmodel, dictGUID_FAB)
                # checkGUID(cr['@OBJTYPE'], cr['@ACTION'], data_objA, olrmodel)
                # checkGUID(cr['@OBJTYPE'], cr['@ACTION'], data_objB, olrmodel)
                if cr['@OBJTYPE'] == "LTC" or cr['@OBJTYPE'] == "LTC3":
                    guid = genIDAspen()
                    reversemodifications[guid] = [cr['@OBJTYPE'], data_objA]
                    modifications[guid] = [cr['@OBJTYPE'], data_objB]
                else:
                    reversemodifications[cr['@OBJGUID']] = [cr['@OBJTYPE'], data_objA]
                    modifications[cr['@OBJGUID']] = [cr['@OBJTYPE'], data_objB]

        olx_diff_model = {}
        olx_diff_model['ADD'] =  additions
        olx_diff_model['DELETE'] =  deletions
        olx_diff_model['FWDMODIFY'] = modifications
        olx_diff_model['REVMODIFY'] = reversemodifications

        return olx_diff_model

    #
    def get_CHANGEREC_1(self):
        #
        res = [False]*len(self.xmlDict['CHANGEREC'])
        #parse
        noXFMR = True
        for i in range(len(self.xmlDict['CHANGEREC'])):
            cr = self.xmlDict['CHANGEREC'][i]
            objtype1 = cr['@OBJTYPE']
            #
            if self.filterOptions.runFilter(cr):
                res[i]=True
                if objtype1 in ['XFMR','XFMR3']:
                    noXFMR = False

        # if no XFMR => no Zcorrect
        if noXFMR:
            for i in range(len(self.xmlDict['CHANGEREC'])):
                cr = self.xmlDict['CHANGEREC'][i]
                objtype1 = cr['@OBJTYPE']
                if objtype1 =='ZCORRECT':
                    res[i]=False
        return res

    def get_CHANGEREC(self, filterOptions=None,
                      typ=[],                # list of objtype ['BUS','LINE']
                      action=[],                   # list of action ['MODIFY','DELETE','ADD']
                      area=[], inside_area = True, # list of area number
                      zone=[], inside_zone = True, # list of zone number
                      tie=False,                   # include tie between selected items and the rest of the network
                      kv=[],                       # [kvmin, kvmax]
                      BNO=[],                      # list of bus number
                      OBJGUID=[],                 # list of objguid
                      ckid = []):

        """
        if val= [] => get all

        typ = ['BUS','LINE']
        action = ['ADD','DELETE','MODIFY']
        area
        zone
        kv =[0,9999]
        OBJGUID = ["{95492957-e939-4e82-9836-d18ba5942fb0}",...]
        BNO [0,99] busNumber
        """
        #
        # data list
        #     one value :  .OLNET, .CHANGEFIELD

        row = len(self.xmlDict['CHANGEREC'])
        result = [True] * row

        checkrange = False
        if filterOptions != None:
            typ = filterOptions.typ
            area = filterOptions.area
            inside_area = filterOptions.inside_area
            zone = filterOptions.zone
            inside_zone = filterOptions.inside_zone
            tie = filterOptions.tie
            kv = filterOptions.kv
            BNO = filterOptions.BNO
            OBJGUID = filterOptions.OBJGUID
            ckid = filterOptions.ckid
            if len(filterOptions.kv) == 2:
                checkrange = True


        if (action == None and typ == None and area == None and zone == None and kv == None and OBJGUID == None and BNO == None):
            return result

        if (action == [] and typ == [] and area == [] and zone == [] and kv == [] and OBJGUID == [] and BNO == [] and ckid == []):
            return result

        if not checkrange:
            arearange, zonerange, kvrange, BNOrange, CKTIDrange, checkrange =  get_filter_range(area, zone, kv, BNO, ckid)
        else:
            arearange, zonerange, kvrange, BNOrange, CKTIDrange, _ =  get_filter_range(area, zone, kv, BNO, ckid)


        result = []
        for cf in self.xmlDict['CHANGEREC']:
            if action != [] and action != None:
                act_res = cf['@ACTION'] in action
            else:
                act_res = True
            if typ != [] and typ != None:
                typ_res = cf['@OBJTYPE'] in typ
            else:
                typ_res = True
            if OBJGUID != [] and OBJGUID != None:
                objguid_res = cf['@OBJGUID'] in OBJGUID
            else:
                objguid_res = True

            objscope = cf['OBJSCOPE']
            ARfilter = False
            ZNfilter = False
            KVfilter = False
            BNOfilter = False
            CKTIDfilter = False
            ar = []
            zn = []
            bno = []
            kvmax = []
            cktid = []
            if objscope != None:
                for scope in objscope['SCOPEFIELD']:
                    if scope['@NAME'] == 'AREA':
                        if (int(scope['@VALUE']) in arearange) == inside_area:
                            ar.append(True)
                        else:
                            ar.append(False)
                    if scope['@NAME'] == 'ZONE':
                        if (int(scope['@VALUE']) in zonerange) == inside_zone:
                            zn.append(True)
                        else:
                            zn.append(False)
                    if scope['@NAME'] == 'CKTID':
                        if (scope['@VALUE']) in CKTIDrange:
                            cktid.append(True)
                        else:
                            cktid.append(False)
                    if scope['@NAME'] == 'KV':
                        if checkrange:
                            if (float(scope['@VALUE']) >= kvrange[0]) and (float(scope['@VALUE']) <= kvrange[1]):
                                kvmax.append(True)
                            else:
                                kvmax.append(False)
                        else:
                            if float(scope['@VALUE']) in kvrange:
                                kvmax.append(True)
                            else:
                                kvmax.append(False)
                    if scope['@NAME'] == 'BNO':
                        if int(scope['@VALUE']) in BNOrange:
                            bno.append(True)
                        else:
                            bno.append(False)
                    else:
                        BNOfilter = True

                if kvrange == []:
                    KVfilter = True
                else:
                    KVfilter = min(kvmax)

                if BNOrange == []:
                    BNOfilter = True
                else:
                    BNOfilter = max(bno)

                if CKTIDrange ==[]:
                    CKTIDfilter = True
                else:
                    if len(cktid) > 0:
                        CKTIDfilter = max(cktid)

                if arearange ==[]:
                    ARfilter = True
                else:
                    if tie:
                        if len(ar) > 0:
                            ARfilter = max(ar)
                        else:
                            ARfilter = False
                    else:
                        if len(ar) > 0:
                            ARfilter = min(ar)
                        else:
                            ARfilter = False

                if zonerange ==[]:
                    ZNfilter = True
                else:
                    if tie:
                        if len(zn) >0:
                            ZNfilter = max(zn)
                        else:
                            ZNfilter = False
                    else:
                        if len(zn) >0:
                            ZNfilter = min(zn)
                        else:
                            ZNfilter = False

                res = ARfilter & ZNfilter & KVfilter & BNOfilter & CKTIDfilter
                filter_scope = res
            else:
                filter_scope = False
                # if boundaryequipment:
                #     filter_scope = False
                # else:
                #     filter_scope = True


            filter_all = filter_scope & act_res & objguid_res & typ_res
            # if filter_scope:
            #     print(filter_scope, act_res, objguid_res, typ_res)

            result.append(filter_all)

        return result

#
class DataASPEN_CIMREF:
    """
    read ASPEN ASPENCIMREF file
    """
    def __init__(self, fREF, encoding = 'utf-8'):
        # self.fREF = fREF
        try:
            self.dictRef = xml2Dict(fREF, desc= 'Parsing ASPEN CIMREF', encoding = encoding)['ASPENCIMREF']
        except:
            self.dictRef = None

    def get_CIMREF(self):
        if self.dictRef == None:
            return None

        refdbtables = self.dictRef['REFDBTABLE']
        self.data_cimref = dict()
        sys_params = self.dictRef['SYSTEMPARAMS']
        if sys_params != None:
            try:
                for data in sys_params['OBJGUID']:
                    self.data_cimref[data['@NAME']] = data['@VALUE']
            except:
                pass


        for reftable in refdbtables:
            objtype = reftable['@NAME']
            objcount = int(reftable['@RECCOUNT'])
            if objcount > 0:
                try:
                    refrecs = reftable['REFREC']
                except:
                    refrecs = []
            if type(refrecs) != list:
                refrecs = [reftable['REFREC']]

            # if objtype in ['AREA','ZONE'] and objcount > 0: # by mRID
            #     data_objs = {}
            #     for refrec in refrecs:
            #         data_elem = dict()
            #         listrec = refrec['OBJGUID']
            #         if type(listrec) != list:
            #             listrec = [listrec]
            #         for data in listrec:
            #             data_elem[data['@NAME']] = data['@VALUE']
            #         #
            #         data_objs [refrec['@OLNETID']] = data_elem
            #     self.data_cimref[objtype] = data_objs
            #
            if (objtype in OBJTYPELIST) and (objcount > 0):
                data_objs = {}
                for refrec in refrecs:
                    data_elem = dict()
                    try:
                        listrec = refrec['OBJGUID']
                    except:
                        # listrec = refrec['@OBJGUID']
                        pass
                    if type(listrec) != list:
                        listrec = [listrec]
                    for data in listrec:
                        try:
                            data_elem[data['@NAME']] = data['@VALUE']
                        except:
                            pass
                            # print(refrec)

                    data_objs[refrec['@OBJGUID']] = (data_elem)
                self.data_cimref[objtype] = data_objs

        return self.data_cimref

def get_filter_range(area, zone, kv, BNO, ckid):

    arearange = []
    if (area != None) and (area != []) :
        for part in area:
            if type(part) == str and '-' in part:
                a, b = part.split('-')
                a, b = int(a), int(b)
                arearange.extend(range(a,b+1))
            else:
                arearange.append(int(part))
    zonerange = []
    if (zone != None) and (zone != []):
        for part in zone:
            if type(part) == str and '-' in part:
                a, b = part.split('-')
                a, b = int(a), int(b)
                zonerange.extend(range(a,b+1))
            else:
                zonerange.append(int(part))

    kvrange = []
    checkrangeKV = False
    if (kv != None) and (kv != []) :
        for part in kv:
            if type(part) == str and '-' in part:
                a, b = part.split('-')
                a, b = float(a), float(b)
                kvrange.append(a)
                kvrange.append(b)
                checkrangeKV = True
            else:
                kvrange.append(float(part))

    BNOrange = []
    if (BNO != None) and (BNO != []) :
        for part in BNO:
            if type(part) == str and '-' in part:
                a, b = part.split('-')
                a, b = int(a), int(b)
                BNOrange.extend(range(a,b+1))
            else:
                BNOrange.append(int(part))

    CKTIDrange = []
    if (ckid != None) and (ckid != []) :
        for part in ckid:
           CKTIDrange.append(part)



    return arearange, zonerange, kvrange, BNOrange, CKTIDrange, checkrangeKV

def xmlparse(source, filesize, desc, tagStop, encoding, **kwargs):
    """Parses a XML document into a dictionary.

    Details about how the XML is converted into this dictionary

    :param str,file-like source: Either the path to a XML document to parse
        or a file-like object (with `read` attribute) containing an XML.
    :param encoding
    :return: The parsed XML in dictionary representation.
    :rtype: dict

    """
    # This is the output dict.
    output = {}

    # Keeping track of the depth and position to store data in.
    current_position = []
    current_index = []
    parser1 = CET.XMLParser(encoding=encoding)
    # Start iterating over the Element Tree.
    context = CET.iterparse(source, events=('start', 'end'), parser = parser1, **kwargs) # cElementTree
    with tqdm(total = round(filesize/1E6, 2), desc = desc , ncols = 100, position=0, leave=(tagStop==None) and desc!='' , unit=' MB') as pbar:
        for event, elem in context:
            pbar.update(round(source.tell()/1E6, 2)-pbar.n)
            if tagStop!=None and elem.tag==tagStop:
                break
            if event == 'start':# Start of new tag.
                # Extract the current endpoint so add the new element to it.
                tmp = output
                for cp, ci in zip(current_position, current_index):
                    tmp = tmp[cp]
                    if ci:
                        tmp = tmp[ci]

                this_tag_name = elem.tag
                # If it is a previously unseen tag, create a new key and
                # stick an empty dict there. Set index of this level to None.
                if this_tag_name not in tmp:
                    tmp[this_tag_name] = {}
                    current_index.append(None)
                else:
                    # The tag name already exists. This means that we have to change
                    # the value of this element's key to a list if this hasn't
                    # been done already and add an empty dict to the end of that
                    # list. If it already is a list, just add an new dict and update
                    # the current index.
                    if isinstance(tmp[this_tag_name], list):
                        current_index.append(len(tmp[this_tag_name]))
                        tmp[this_tag_name].append({})
                    else:
                        tmp[this_tag_name] = [tmp[this_tag_name], {}]
                        current_index.append(1)
                # Set the position of the iteration to this element's tag name.
                current_position.append(this_tag_name)
            elif event == 'end': # End of a tag.
                # Extract the current endpoint's parent so we can handle
                # the endpoint's data by reference.
                tmp = output
                for cp, ci in zip(current_position[:-1], current_index[:-1]):
                    tmp = tmp[cp]
                    if ci:
                        tmp = tmp[ci]
                cp = current_position[-1]
                ci = current_index[-1]

                # If this current endpoint is a dict in a list or not has
                # implications on how to set data.
                if ci:
                    setfcn = lambda x: setitem(tmp[cp], ci, x)
                    for attr_name, attr_value in elem.attrib.items():
                        tmp[cp][ci]["@{0}".format(attr_name)] = attr_value
                else:
                    setfcn = lambda x: setitem(tmp, cp, x)
                    for attr_name, attr_value in elem.attrib.items():
                        tmp[cp]["@{0}".format(attr_name)] = attr_value

                # If there is any text in the tag, add it here.
                if elem.text and elem.text.strip():
                    setfcn({'#text': elem.text.strip()})

                # Handle special cases:
                # 1) when the tag only harbours text, replace the dict content with
                #    that very text string.
                # 2) when no text, attributes or children are present, content
                #    is set to None
                # These are detailed in reference [3] in README.
                if ci:
                    nk = len(tmp[cp][ci].keys())
                    if nk == 1 and "#text" in tmp[cp][ci]:
                        tmp[cp][ci] = tmp[cp][ci]["#text"]
                    elif nk == 0:
                        tmp[cp][ci] = None
                else:
                    nk = len(tmp[cp].keys())
                    if nk == 1 and "#text" in tmp[cp]:
                        tmp[cp] = tmp[cp]["#text"]
                    elif nk == 0:
                        tmp[cp] = None

                # Remove the outermost position and index, since we just finished
                # handling that element.
                current_position.pop()
                current_index.pop()

                # Most important of all, release the element's memory allocations
                # so we actually benefit from the iterative processing!
                elem.clear()
                # while elem.getprevious() is not None:
                #     del elem.getparent()[0]
    return output


def xmliter(source, tagname, encoding, **kwargs):
    """Iterates over a XML document and yields specified tags
    in dictionary form.
    This iteration method has a very small memory footprint as compared to
    the :py:meth:`xmlr.xmlparse` since it continuously discards all
    processed data. It is therefore useful for traversing structured
    XML where a known tag's members are desired.
    Details about how the XML is converted into this dictionary (json)
    representation is described in reference [3] in the README.
    :param str,file-like source: Either the path to a XML document to parse
        or a file-like object (with `read` attribute) containing an XML.
    :param str tagname: The name of the tag type to extract and iterate over.

    :return: The desired tags parsed XML in dictionary representation.
    :rtype: dict
    """


    output = None
    is_active = False

    # total_size = os.path.getsize(source)
    # fivepct = 5
    # print_flag = False

    # Keeping track of the depth and position to store data in.
    current_position = []
    current_index = []
    parser1 = CET.XMLParser(encoding=encoding)
    # Start iterating over the Element Tree.
    for event, elem in CET.iterparse(
            source, events=('start', 'end'), parser = parser1, **kwargs):
        if (event == 'start') and ((elem.tag == tagname) or is_active):
            # Start of new tag.
            if output is None:
                output = {}
                is_active = True

            # Extract the current endpoint so add the new element to it.
            tmp = output
            for cp, ci in zip(current_position, current_index):
                tmp = tmp[cp]
                if ci:
                    tmp = tmp[ci]

            this_tag_name = elem.tag
            # If it is a previously unseen tag, create a new key and
            # stick an empty dict there. Set index of this level to None.
            if this_tag_name not in tmp:
                tmp[this_tag_name] = {}
                current_index.append(None)
            else:
                # The tag name already exists. This means that we have to change
                # the value of this element's key to a list if this hasn't
                # been done already and add an empty dict to the end of that
                # list. If it already is a list, just add an new dict and update
                # the current index.
                if isinstance(tmp[this_tag_name], list):
                    current_index.append(len(tmp[this_tag_name]))
                    tmp[this_tag_name].append({})
                else:
                    tmp[this_tag_name] = [tmp[this_tag_name], {}]
                    current_index.append(1)

            # Set the position of the iteration to this element's tag name.
            current_position.append(this_tag_name)
        elif (event == 'end') and ((elem.tag == tagname) or is_active):
            # End of a tag.

            # Extract the current endpoint's parent so we can handle
            # the endpoint's data by reference.
            tmp = output
            for cp, ci in zip(current_position[:-1], current_index[:-1]):
                tmp = tmp[cp]
                if ci:
                    tmp = tmp[ci]
            cp = current_position[-1]
            ci = current_index[-1]

            # If this current endpoint is a dict in a list or not has
            # implications on how to set data.
            if ci:
                setfcn = lambda x: setitem(tmp[cp], ci, x)
                for attr_name, attr_value in elem.attrib.items():
                    tmp[cp][ci]["@{0}".format(attr_name)] = attr_value
            else:
                setfcn = lambda x: setitem(tmp, cp, x)
                for attr_name, attr_value in elem.attrib.items():
                    tmp[cp]["@{0}".format(attr_name)] = attr_value

            # If there is any text in the tag, add it here.
            if elem.text and elem.text.strip():
                setfcn({'#text': (elem.text.strip())})

            # Handle special cases:
            # 1) when the tag only harbours text, replace the dict content with
            #    that very text string.
            # 2) when no text, attributes or children are present, content
            #    is set to None
            # These are detailed in reference [3] in README.
            if ci:
                nk = len(tmp[cp][ci].keys())
                if nk == 1 and "#text" in tmp[cp][ci]:
                    tmp[cp][ci] = tmp[cp][ci]["#text"]
                elif nk == 0:
                    tmp[cp][ci] = None
            else:
                nk = len(tmp[cp].keys())
                if nk == 1 and "#text" in tmp[cp]:
                    tmp[cp] = tmp[cp]["#text"]
                elif nk == 0:
                    tmp[cp] = None

            if elem.tag == tagname:
                # End of our desired tag.
                # Finish up this document and yield it.
                current_position = []
                current_index = []
                is_active = False


                yield output.get(tagname)

                output = None
            else:
                # Remove the outermost position and index, since we just
                # finished handling that element.
                current_position.pop()
                current_index.pop()

            # Most important of all, release the element's memory
            # allocations so we actually benefit from the
            # iterative processing.
            elem.clear()

def xml2Dict(fxml,desc,tagStop=None, encoding = 'iso8859-1'):
    ## convert data XML => Dict
    filesize = os.path.getsize(fxml)
    # f = open(fxml, 'r',  encoding = encoding, errors = 'ignore')
    f = open(fxml, 'rb')
    try:
        xmlDict =  xmlparse(f, filesize, desc=desc,tagStop=tagStop, encoding= encoding)
    except Exception as e:
        print(e)
        raise Exception("Input file format required: ASPEN OLX/ADX/CIMREF")
    f.close()
    return xmlDict
#
def xml2Dict2(fi, desc='', tagStop=None, encoding='iso8859-1'):
    filesize = os.path.getsize(fi)
    source = open(fi,"rb")
    output = {}
    #
    current_position = []
    current_index = []
    parser1 = CET.XMLParser(encoding=encoding)
    #
    context = CET.iterparse(source, events=('start', 'end'), parser = parser1)
    context = iter(context)
    _, root = next(context)
    #
    k = 0
    with tqdm(total = round(filesize/1E6, 2), desc = desc , ncols = 100, position=0, leave=(tagStop is None) and desc , unit=' MB') as pbar:
        for event, elem in context:
            k+=1
            if k>10000:
                k = 0
                pbar.update(round(source.tell()/1E6, 2)-pbar.n)
            #
            if (tagStop is not None and elem.tag==tagStop):
                break
            if event == 'start':# Start of new tag.
                # Extract the current endpoint so add the new element to it.
                tmp = output
                for cp, ci in zip(current_position, current_index):
                    tmp = tmp[cp]
                    if ci:
                        tmp = tmp[ci]
                this_tag_name = elem.tag
                if this_tag_name not in tmp:
                    tmp[this_tag_name] = {}
                    current_index.append(None)
                else:
                    if isinstance(tmp[this_tag_name], list):
                        current_index.append(len(tmp[this_tag_name]))
                        tmp[this_tag_name].append({})
                    else:
                        tmp[this_tag_name] = [tmp[this_tag_name], {}]
                        current_index.append(1)
                # Set the position of the iteration to this element's tag name.
                current_position.append(this_tag_name)
            elif event == 'end': # End of a tag.
                tmp = output
                for cp, ci in zip(current_position[:-1], current_index[:-1]):
                    tmp = tmp[cp]
                    if ci:
                        tmp = tmp[ci]
                try:
                    cp = current_position[-1]
                    ci = current_index[-1]
                    if ci:
                        setfcn = lambda x: setitem(tmp[cp], ci, x)
                        for attr_name, attr_value in elem.attrib.items():
                            tmp[cp][ci]["@{0}".format(attr_name)] = attr_value
                    else:
                        setfcn = lambda x: setitem(tmp, cp, x)
                        for attr_name, attr_value in elem.attrib.items():
                            tmp[cp]["@{0}".format(attr_name)] = attr_value
                    # If there is any text in the tag, add it here.
                    if elem.text and elem.text.strip():
                        setfcn({'#text': elem.text.strip()})
                    if ci:
                        nk = len(tmp[cp][ci].keys())
                        if nk == 1 and "#text" in tmp[cp][ci]:
                            tmp[cp][ci] = tmp[cp][ci]["#text"]
                        elif nk == 0:
                            tmp[cp][ci] = None
                    else:
                        nk = len(tmp[cp].keys())
                        if nk == 1 and "#text" in tmp[cp]:
                            tmp[cp] = tmp[cp]["#text"]
                        elif nk == 0:
                            tmp[cp] = None
                    current_position.pop()
                    current_index.pop()
                    elem.clear()
                except:
                    pass
            root.clear()
        if k>0:
            pbar.update(round(source.tell()/1E6, 2)-pbar.n)
    source.close()
    return output

def tagxml2Dict(fxml, tag, encoding = 'iso8859-1'):
    f = open(fxml, 'rb')
    output = []
    for key in xmliter(f, tagname = tag, encoding= encoding):
        output.append(key)
        if len(output) >0:
            break
    f.close()
    return output[0]



def checkGUID(objtype, action, data, olrmodel):
    guidobjs = ['TERMGUID1','TERMGUID2','TERMGUID3','TERMGUID4','OBJGUID',
               'PGUID1','PGUID2']
    fullmodel = olrmodel.modelcim
    for guid in guidobjs:
        if guid in data:
            termguid = data[guid]
            try:
                obj = fullmodel[termguid]
                print('GUID={}, CLASS={}'.format(termguid,obj.__class__.__name__))
            except:
                print('Object OBJTYPE={}, ACTION={}, TYPE={}, GUID={}'.format(objtype, action, guid, termguid))
                continue

def correctGUID(objtype, action, data, olrmodel):
    """
    Correct GUID from File B and convert to GUID from File A

    Parameters
    ----------
    objtype : str
        Type of object.
    action : str
        Type of action: DELETE, ADD, MODIFY.
    data : dict
        Object data, contains GUID from file A and file B.
    olrmodel : Object
        OLR model of file A.

    Returns
    -------
    data : dict
        Modification of data to make data compatible with file A.

    """
    if objtype == 'BUS' and action =='ADD':
        return data

    for i in range(1,5):
        name = 'TERMNAME{}'.format(i)
        kv = 'TERMKV{}'.format(i)
        termguid = 'TERMGUID{}'.format(i)
        termbno = 'TERMBNO{}'.format(i)

        if termguid in data:
            busId = "'"+ data[name] + "' " + data[kv] + 'kV'
            try:
                node = olrmodel.busid_dict[busId]
                guidFA = node.UUID
                bnoFA = node.mRID
                if guidFA != data[termguid]:
                    data[termguid] = guidFA
                    if termbno in data:
                        if data[termbno] != bnoFA:
                            data[termbno] = bnoFA
                    if 'BS_NO' in data:
                        if data['BS_NO'] != bnoFA:
                            data['BS_NO'] = bnoFA
                    if 'BNO' in data:
                        if data['BNO'] != bnoFA:
                            data['BNO'] = bnoFA
            except:
                # print('OBJTYPE={} ACTION={} on new BUS with OLNETID={}  OBJGUID={}'.format(objtype, action, busId, data[termguid]))
                continue

    if objtype =='BUS':
        if 'OBJGUID' in data:
            if data['OBJGUID'] != data['TERMGUID1']:
                data['OBJGUID'] = data['TERMGUID1']


    return data

def getGUID_FA_FB(objtype, action, data, olrmodel, dictGUID_FAB):
    for i in range(1,4):
        name = 'TERMNAME{}'.format(i)
        kv = 'TERMKV{}'.format(i)
        termguid = 'TERMGUID{}'.format(i)

        if termguid in data:
            busId = "'"+ data[name] + "' " + data[kv] + 'kV'
            try:
                nodeFA = olrmodel.busid_dict[busId]
                guidFA = nodeFA.UUID
                if guidFA != data[termguid]:
                    if guidFA not in dictGUID_FAB:
                        dictGUID_FAB[guidFA] = data[termguid]
            except:
                 continue

    return dictGUID_FAB

#

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
            OlxAPILib.setIco(root,"","1LINER.ico")
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
            d = AppUtils.Gui_Select(parent=root,w1=25,xOK=120,yOK=150,xCom=40,yCom=70,data=data,select=select)
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
            self.bhnd = OlxAPILib.getEquipmentHandle(TC_BUS)
            self.kv   = OlxAPILib.getEquipmentData(self.bhnd,BUS_dKVnominal)
            self.name = OlxAPILib.getEquipmentData(self.bhnd,BUS_sName     )
            for i in range(len(self.name)):
                self.name[i]= (self.name[i].upper()).replace(" ","")
            self.num  = OlxAPILib.getEquipmentData(self.bhnd,BUS_nNumber)
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
        busa = OlxAPILib.getEquipmentHandle(TC_BUS)
        for b1 in busa:
            bra = OlxAPILib.getBusEquipmentData([b1],TC_BRANCH)[0]
            #
            ehand = OlxAPILib.getEquipmentData(bra,BR_nHandle)
            for i in range(len(ehand)):
                br1 = bra[i]
                e1 = ehand[i]
                if e1 not in setEhnd:
                    na1 = OlxAPILib.getDataByBranch(br1,"sName")
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
            br = OlxAPILib.getBusEquipmentData([b1],TC_BRANCH)[0]
            for br1 in br:
                if len(self.bus2)==0:
                    self.__test_sID(br1)
                else:
                    b2t = OlxAPILib.getEquipmentData([br1],BR_nBus2Hnd)[0]
                    if b2t in self.bus2:
                        self.__test_sID(br1)
        return self.__selectEnd()
    #
    def __test_sID(self,br1):
        if self.CktID!="":
            id1 = OlxAPILib.getDataByBranch(br1,"sID")
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
            AppUtils.setIco(root,"","1LINER.ico")
        #
        if len(self.bra)==1:
            if self.bra[0]==0:
                self.sFound = "No branch found!"
            else:
                self.sFound = "Branch found:\n\t" + OlxAPILib.fullBranchName(self.bra[0])
            #
            if self.gui ==1:
                root.withdraw()
                tkm.showinfo("Branch search:",self.sInput + "\n"+self.sFound)
                root.destroy()
            return self.bra[0]

        # multipe result
        data = []
        for bi in self.bra:
            data.append(OlxAPILib.fullBranchName(bi))
        #
        self.sFound = "Branch found:"
        for d1 in data:
            self.sFound += "\n\t"+ d1

        if self.gui==0:
            self.sSelected = "Selected:\n\t" + OlxAPILib.fullBranchName(self.bra[0])
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
            d = AppUtils.Gui_Select(parent=root,w1=50,xOK=180,yOK=150,xCom=40,yCom= 90,data=data,select=select)
            root.mainloop()
            i = data.index(select[0])
            self.sSelected = "Selected:\n\t" + OlxAPILib.fullBranchName(self.bra[i])
            return self.bra[i]
        except:
            self.sSelected = "Selected:\n\t" + OlxAPILib.fullBranchName(self.bra[0])
            return self.bra[0]
#
# Create a run marker file in this folder and keep it open during the
#        # entire Python execution session. OneLiner will pause execution
#        # as long as this file exists to wait for the result from Python
def markerStart(opath):
    if opath:
        AppUtils.deleteFile(opath+'\\success')
        AppUtils.deleteFile(opath+'\\cancel')
        return AppUtils.openFile(opath,'running')
    return None,None
#
def markerSucces(opath):
    if opath:
        AppUtils.createFile(opath,'success')
#
def markerStop(opath,FRUNNING,SRUNNING):
    if opath:
        if not AppUtils.isFile(opath,'success'):
            AppUtils.createFile(opath,'cancel')
        AppUtils.closeFile(FRUNNING)
        AppUtils.deleteFile(SRUNNING)
#
def getValSum(fi,olxpath):
    valSum = {}
    OlxAPILib.open_olrFile(fi,olxpath,prt=False)
    #
    valSum['BUS']     = OlxAPILib.getParaSys(SY_nNObus)
    valSum['GEN']     = OlxAPILib.getParaSys(SY_nNOgen)
    valSum['GENUNIT'] = OlxAPILib.getParaSys(SY_nNOgenUnit)
    valSum['GENW3']   = OlxAPILib.getParaSys(SY_nNOgenW3)
    valSum['GENW4']   = OlxAPILib.getParaSys(SY_nNOgenW4)
    valSum['LOAD']    = OlxAPILib.getParaSys(SY_nNOload)
    valSum['SHUNT']   = OlxAPILib.getParaSys(SY_nNOshunt)
    valSum['SHUNTUNIT']= OlxAPILib.getParaSys(SY_nNOshuntUnit)
    valSum['LINE']    = OlxAPILib.getParaSys(SY_nNOline)
    valSum['DCLINE2'] = OlxAPILib.getParaSys(SY_nNODCLine2)
    valSum['SERIESRC']= OlxAPILib.getParaSys(SY_nNOseriescap)
    valSum['XFMR']    = OlxAPILib.getParaSys(SY_nNOxfmr)
    valSum['XFMR3']   = OlxAPILib.getParaSys(SY_nNOxfmr3)
    valSum['SHIFTER'] = OlxAPILib.getParaSys(SY_nNOps)
    valSum['LTC']     = OlxAPILib.getParaSys(SY_nNOltc)
    try:
        valSum['LTC3']= OlxAPILib.getParaSys(SY_nNOltc3)
    except:
        valSum['LTC3'] = 0
    valSum['ZCORRECT']= OlxAPILib.getParaSys(SY_nNOzCorrect)
    valSum['MULINE']  = OlxAPILib.getParaSys(SY_nNOmuPair)
    valSum['SWITCH']  = OlxAPILib.getParaSys(SY_nNOswitch)
    valSum['LOADUNIT']= OlxAPILib.getParaSys(SY_nNOloadUnit)
    valSum['SVD']     = OlxAPILib.getParaSys(SY_nNOsvd)
    valSum['RLYGROUP']= OlxAPILib.getParaSys(SY_nNOrlyGroup)
    valSum['RLYOC']   = OlxAPILib.getParaSys(SY_nNOrlyOCP) + OlxAPILib.getParaSys(SY_nNOrlyOCG)
    valSum['RLYDS']   = OlxAPILib.getParaSys(SY_nNOrlyDSP) + OlxAPILib.getParaSys(SY_nNOrlyDSG)
    valSum['RLYD']    = OlxAPILib.getParaSys(SY_nNOrlyD)
    valSum['RLYV']    = OlxAPILib.getParaSys(SY_nNOrlyV)
    valSum['FUSE']    = OlxAPILib.getParaSys(SY_nNOfuse)
    valSum['RECLSR']  = OlxAPILib.getParaSys(SY_nNOrecloserG) + OlxAPILib.getParaSys(SY_nNOrecloserP)
    valSum['CCGEN']   = OlxAPILib.getParaSys(SY_nNOccgen)
    valSum['BREAKER'] = OlxAPILib.getParaSys(SY_nNObreaker)
    valSum['SCHEME']  = OlxAPILib.getParaSys(SY_nNOscheme)
    valSum['baseMVA'] = OlxAPILib.getParaSys(SY_dBaseMVA)
    valSum['AREA']    = OlxAPILib.getParaSys(SY_nNOarea)
    valSum['ZONE']    = OlxAPILib.getParaSys(SY_nNOzone)
    #
    return valSum
#
def compare2files(olxpath,f1,f2):
    fadx = AppUtils.get_file_out(fo='' , fi=f1 , subf='' , ad='_DIFF', ext='.ADX')
    #
    pi = [f1,f2,fadx]
    print("Generating ADX file....")
    OlrConverter(pi,olxpath) # with subprocess if big file
    try:
        adx = DataASPEN_ADX(fadx)
        adx.getValSum()
        CHANGECOUNT = int(adx.valSum['CHANGECOUNT'])
    except:
        CHANGECOUNT = -1
    return CHANGECOUNT,fadx
def getFileConfig(fCFG,path):
    if os.path.isfile(fCFG):
        return os.path.abspath(fCFG)
    f1 = os.path.join(path,fCFG)
    if os.path.isfile(f1):
        return f1
    return

def genIDAspen(): # generate UUID by Python uuid4
    # return "_"+str(uuid.uuid4())
    return '{%s}' %str(uuid.uuid4())
