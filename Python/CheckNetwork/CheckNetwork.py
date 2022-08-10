"""
Purpose: CheckNetwork anomalies
op-1: Check 2 Windings Transformer Config
    - Check phase shift of all 2-winding transformers with wye-delta connection
        When a transformer with high side lagging the low side is found,
        this function converts it to make the high side lead
    - Check 2 Windings Transformer Config DGD, convert =>GGG
op-2: Check GEN3,4 VCCS
    1. Under-rated transformers near GenW3, GenW4 or VCCS
    2. Lack of a wye-connected transformer winding in front of a  GenW3, GenW4 or VCCS (unless it is a simple STATCOM)
    3. Have significant prefault MW generation from GenW3, GenW4 or VCCS when there are no loads in the network
    4. GenW3, GenW4 or VCCS units that are  connected with short lines or switches
op-3: Check duplicate
    Load unit CID (duplicate in a bus)
    Gen unit CID (duplicate in a bus)
    Shunt unit CID (duplicate in a bus)
    BusNumber (duplicate and >0)
op-4: check Line
    Some line type field text contains leading and trailing blank spaces
    check RX

op-5: Check memo
    memo text contains leading and trailing blank spaces and tab characters

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.3"


# IMPORT -----------------------------------------------------------------------
import logging
logger = logging.getLogger(__name__)
import os,sys,csv,time,math
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
olxpath = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python'
#olxpathpy = os.path.split(PATH_FILE)[0]
# INPUTS cmdline ---------------------------------------------------------------
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'Check network anomalies'
PARSER_INPUTS.add_argument('-op' , help = '*(int) operation 0-Check ALL; 1-Check XFMR; 2-GEN34 VCCS; 3-Check Duplicate (CID, BusNumber); 4-Check Line; 5-Check Memo',default = 0, type=int,nargs='+', metavar='')
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-fo' , help = ' (str) Path name out file',default = "", type=str, metavar='')
PARSER_INPUTS.add_argument('-olxpath' , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
PARSER_INPUTS.add_argument('-demo', help = ' (int) demo [0-ignore, 1-run demo]', default = 0,type=int,metavar='')
#
PARSER_INPUTS.add_argument('-opath', help = ' (str) Path name output folder $ASPEN internal parameter$',default = '',type=str, metavar='')
#
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
try:
    import AppUtils
    from OlxObj import *
except:
    raise Exception('\nPlease check in folder olxpathpy:"%s"\n\tnot found OlxAPI Python wrapper OlxAPI.py and relevant libraries'%ARGVS.olxpathpy)

#
def run_demo():
    if ARGVS.demo==1:
        ft = PATH_FILE+ '\\sample\\CN01.OLR'
        if os.path.isfile(ft):
            ARGVS.fi = ft
        else:
            ARGVS.fi = AppUtils.getASPENFile(olxpath,'SAMPLE30.OLR')
        choice = AppUtils.ask_run_demo(PY_FILE,0,ARGVS.fi,'')
        if choice=="yes":
            ARGVS.op = [0]
            run()
    else:
        AppUtils.demo_notFound(PY_FILE,ARGVS.demo,[1])
#
def writeRes(res,fwriter,title):
    if res:
        fwriter.writerow(title)
        for r1 in res:
            fwriter.writerow(r1)
#
class CHECK_DUPLICATE:
    def __init__(self,fwriter):
        self.fwriter = fwriter
    #
    def __getNewCID(vi,k):
        try:
            i1 = int(vi[:1])
        except:
            i1 = 0
        return str(i1+k),k+1
    #
    def __changeCID(la,cid):
        sidOld, sidNew = '',''
        for c1 in cid:
            sidOld+= c1+ ' '
        #
        for i in range (1,len(cid)):
            vi = cid[i]
            if vi in cid[:i]:
                k = 1
                while True and k<89:
                    vn,k = CHECK_DUPLICATE.__getNewCID(vi,k)
                    try:
                        la[i].CID = vn
                        la[i].postData()
                        cid[i] = vn
                        break
                    except:
                        pass
        for c1 in cid:
            sidNew+= c1+ ' '
        return sidOld[:-1], sidNew[:-1]
    #
    def check_Loadunit(self):
        """
        Check: Buses contain multiple load units with the same ID
        """
        ec = 'E21'
        self.fwriter.writerow([ec,'Buses contain multiple load units with the same ID'])
        title = [ec,'BUS ID','PARAMETER','VALUE','CHANGED TO']
        la = OLCase.LOAD
        res = []
        for la1 in la:
            cid = []
            lau = la1.LOADUNIT
            for u1 in lau :
                cid.append(u1.CID)
            if len(cid)>1 and len(cid)!=len(set(cid)):
                sidOld, sidNew = CHECK_DUPLICATE.__changeCID(lau,cid)
                r1 = [ec,la1.BUS.toString(),'LOAD UNIT CID',sidOld,sidNew]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Buses contain multiple load units with the same ID',len(res)])
    #
    def check_Genunit(self):
        """
        Check: Buses contain multiple Gen units with the same ID
        """
        ec = 'E22'
        self.fwriter.writerow([ec,'Buses contain multiple Generator units with the same ID'])
        title = [ec,'BUS ID','PARAMETER','VALUE', 'CHANGED TO']
        ga = OLCase.GEN
        res = []
        for ga1 in ga:
            cid = []
            lau = ga1.GENUNIT
            for u1 in lau:
                cid.append(u1.CID)
            if len(cid)>1 and len(cid)!=len(set(cid)):
                sidOld, sidNew = CHECK_DUPLICATE.__changeCID(lau,cid)
                r1 = [ec,ga1.BUS.toString(),'GENERATOR UNIT CID',sidOld,sidNew]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Buses contain multiple Generator units with the same ID',len(res)])
    #
    def check_Shuntunit(self):
        """
        Check: Buses contain multiple Shunt units with the same ID
        """
        ec = 'E23'
        self.fwriter.writerow([ec,'Buses contain multiple Shunt units with the same ID'])
        title = [ec,'BUS ID','PARAMETER','VALUE', 'CHANGED TO']
        ga = OLCase.SHUNT
        res = []
        for ga1 in ga:
            cid = []
            lau = ga1.SHUNTUNIT
            for u1 in lau:
                cid.append(u1.CID)
            if len(cid)>1 and len(cid)!=len(set(cid)):
                sidOld, sidNew = CHECK_DUPLICATE.__changeCID(lau,cid)
                r1 = [ec,ga1.BUS.toString(),'SHUNT UNIT CID',sidOld,sidNew]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Buses contain multiple Shunt units with the same ID',len(res)])
    #
    def checkBusNumber(self):
        """
        Check: Buses have the same non-zero bus number
        """
        ec = 'E24'
        self.fwriter.writerow([ec,'Buses have the same non-zero bus number'])
        title = [ec,'BUS ID','PARAMETER','VALUE','CHANGED TO']
        ba = OLCase.BUS
        bn = []
        flag = []
        res = []
        for b1 in ba:
            bn1 = b1.NO
            bn.append(bn1)
            if bn1>0:
                flag.append(True)
            else:
                flag.append(False)
        #
        for i in range(len(ba)):
            if flag[i]:
                bn1 = bn[i]
                if bn1 in bn[i+1:]:
                    r1 = [ec,ba[i].toString()]
                    sn = str(bn1) + ' '
                    for j in range(i+1,len(ba)):
                        if bn[j]==bn1:
                            r1.append(ba[j].toString())
                            flag[j]=False
                            bn2 = CHECK_DUPLICATE.__changeBusNumber(ba[j],bn1)
                            sn +=str(bn2) + ' '
                    #
                    r1.append('BUS NUMBER')
                    r1.append(bn1)
                    r1.append(sn[:-1])
                    res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Buses have the same non-zero bus number',len(res)])
    #
    def __changeBusNumber(bus,bn):
        k = 0
        while True and k<1e5:
            try:
                k+=1
                bus.NO = bn+k
                bus.postData()
                return bn+k
            except:
                pass
        return 0
#
class CHECK_GENW34_VCCS:
    def __init__(self,fwriter):
        self.fwriter = fwriter
        self.ga = OLCase.GENW3
        self.ga.extend(OLCase.GENW4)
        self.ga.extend(OLCase.CCGEN)
    #
    def __getP_CCGEN(v,i,ang,kv,vc):
        k = 0
        for i1 in range(len(v)):
            if v[i1]>0:
                k=i1
            if vc>=v[i1]:
                break
        if k==0:
            v0 = v[0]
            i0 = i[0]
            a0 = ang[0]
        else:
            v0 = vc
            try:
                i0 = i[k] +( i[k-1]-i[k])/( v[k-1]-v[k]) *(v0-v[k])
                a0 = ang[k] +( ang[k-1]-ang[k])/( v[k-1]-v[k]) *(v0-v[k])
            except:
                i0 = i[k]
                a0 = ang[k]
        return v0*kv * i0 * math.cos(a0 *math.pi/180)/1e3
    #
    def check1(self):
        """
        Check: Under-rated transformers near GenW3, GenW4 or VCCS
        """
        ec = 'E11'
        self.fwriter.writerow([ec,'Check Under-rated transformers near GenW3+GenW4+VCCS'])
        title = [ec,'GEN ID','XFMR ID','MVA rating Gen','MVA StepUp Transf']
        res = []
        for gi in self.ga:
            b1 = gi.BUS
            if type(gi)==CCGEN:
                mva_i = gi.MVARATE
            else:
                mva_i = gi.UNITS*gi.MVA
            #
            x2a = b1.XFMR
            x3a = b1.XFMR3
            mvaX = 0
            xs = ''
            for x2i in x2a:
                ma = max(x2i.MVA1,x2i.MVA2,x2i.MVA3)
                if ma>0:
                    mvaX += ma
                else:
                    mbase = x2i.BASEMVA
                    if mbase!=100.0:
                        mvaX += ma
                    else:
                        xpu = x2i.X
                        if xpu>0.2:
                            mvaX += 100/(xpu/0.2)
                        elif xpu<0.05:
                            mvaX += 100/(0.05/xpu)
                        else:
                            mvaX +=100
                xs += x2i.toString()+' ; '
            for x3i in x3a:
                ma = max(x3i.MVA1,x3i.MVA2,x3i.MVA3)
                if ma>0:
                    mvaX += ma
                else:
                    mbase = x3i.BASEMVA
                    if mbase!=100.0:
                        mvaX += ma
                    else:
                        xpu = x3i.XPS
                        if xpu>0.2:
                            mvaX += 100/(xpu/0.2)
                        elif xpu<0.05:
                            mvaX += 100/(0.05/xpu)
                        else:
                            mvaX +=100
                #
                xs += x3i.toString()+' ; '
            xs = xs[:-2]
            #
            if len(x2a)==0 and len(x3a)==0:
                r1 = [ec, gi.toString(), 'None', str(round(mva_i,1)), str(round(mvaX,1))]
                res.append(r1)
            elif mvaX<mva_i:
                r1 = [ec, gi.toString(), xs, str(round(mva_i,1)), str(round(mvaX,1))]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Check Under-rated transformers near GenW3+GenW4+VCCS',len(res)])
    #
    def check2(self):
        """
        Lack of a wye-connected transformer winding in front of a GenW3+GenW4+VCCS (unless it is a simple STATCOM)
        """
        ec = 'E12'
        self.fwriter.writerow([ec,'Lack of a wye-connected transformer winding in front of a GenW3+GenW4+VCCS (unless it is a simple STATCOM)'])
        title = [ec,'GEN ID', 'XFMR ID','Connections StepUp Transf']
        #
        dictCon = {'GG':'YY/auto', 'GE':'Yd11','GD':'Yd1', 'DD':'dd','ZG':'zy11','ZX':'zy1','ZD':'zd0'}
        # cona = x2i.getData(['sCfgP', 'sCfgS', 'sCfgST', 'sCfg1', 'sCfg2', 'sCfg2T'])
        # 1,2 ['G', 'G', 'G', 'G', 'G', 'G']  YY/auto
        # 3   ['G', 'E', 'D', 'G', 'E', 'D']  Yd11
        # 4   ['G', 'D', 'D', 'G', 'D', 'D']  Yd1
        # 5   ['D', 'D', 'G', 'D', 'D', 'G']  dd
        # 6   ['Z', 'G', 'G', 'Z', 'G', 'G']  zy11
        # 7   ['Z', 'X', 'G', 'Z', 'X', 'G']  zy1
        # 8   ['Z', 'D', 'D', 'Z', 'D', 'D']  zd0
        res = []
        for gi in self.ga:
            b1 = gi.BUS
            x2a = b1.XFMR
            conX2 = ''
            x2s = ''
            flag = False
            for x2i in x2a:
                bus1 = x2i.BUS1
                bus2 = x2i.BUS2
                con1 = x2i.CONFIGP
                con2 = x2i.CONFIGS
                #
                try:
                    conX2 += dictCon[con1+con2] +' ; '
                except:
                    pass
                #
                if b1.equals(bus1):
                    if con1 not in['G']:
                        flag = True
                elif b1.equals(bus2) :
                    if con2 not in ['G','X']:
                        flag = True
                x2s += x2i.toString()+' ; '
            #
            x2s = x2s[:-2]
            conX2 = conX2[:-2]
            #
            if flag :
                r1 = [ec, gi.toString(), x2s,conX2 ]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Lack of a wye-connected transformer winding in front of a GenW3+GenW4+VCCS (unless it is a simple STATCOM)',len(res)])
    #
    def check3(self):
        """
        Have significant prefault MW generation from GenW3+GenW4+VCCS when there are no loads in the network
        """
        ec = 'E13'
        vc = 1.0
        self.fwriter.writerow([ec,'Have significant prefault MW generation from GenW3+GenW4+VCCS when there are no loads in the network'])
        title = [ec,'GEN ID','MW generation']
        #
        la = OLCase.LOADUNIT
        mwload = [0,0,0]
        for l1 in la:
            if l1.FLAG==1:
                l3 = l1.MW
                for i in range(3):
                    mwload[i] += l3[i]
        ma = sum(mwload)
        #
        mwgen = 0
        res = []
        for gi in self.ga:
            b1 = gi.BUS
            if type(gi)==CCGEN:
                v = gi.V
                i = gi.I
                a = gi.A
                kv = b1.KV
                m1 = CHECK_GENW34_VCCS.__getP_CCGEN(v,i,a,kv,vc)
            else:
                m1 = gi.UNITS*gi.MW
            mwgen +=m1
            #
            if m1>0.01:
                r1 = [ec,gi.toString(),str(round(m1,1))]
                res.append(r1)
        # finish
        nc =0
        if res and mwgen> ma:
            self.fwriter.writerow([ec, 'sumGen='+str(round(mwgen,1)),'sumLoad='+str(round(ma,1))])
            writeRes(res,self.fwriter,title)
            nc = len(res)
        self.fwriter.writerow(['COUNT','Have significant prefault MW generation from GenW3+GenW4+VCCS when there are no loads in the network',nc])
    #
    def __testFunc4(self,ba,i,listBus):
        for b1 in ba:
            for j in range(len(listBus)):
                if i!=j:
                    if b1.isInList(listBus[j]):
                        return True
        return False
    #
    def check4(self):
        """
        GenW3+GenW4+VCCS units that are connected with short lines or switches
        """
        ec = 'E14'
        self.fwriter.writerow([ec,'GenW3+GenW4+VCCS units that are connected with short lines or switches'])
        title = [ec, 'GEN ID']
        #
        xc = 0.002
        listBus = []
        for gi in self.ga:
            b1 = gi.BUS
            ba = [b1]
            x23 = b1.XFMR
            x23.extend(b1.XFMR)
            for xi in x23:
                b2 = xi.BUS
                for bi in b2:
                    if not bi.isInList(ba):
                        ba.append(bi)
            #
            sla0 = [] # sw+line
            flag = True
            k = 0
            while (flag and k<10):
                flag = False
                k+=1
                for bi in ba:
                    sla  = bi.SWITCH
                    sla.extend(bi.LINE)
                    for sl1 in sla:
                        if type(sl1)==SWITCH:
                            x1 = 0
                        else:
                            x1 = sl1.X
                        #
                        if x1<=xc:
                            if not sl1.isInList(sla0):
                                sla0.append(sl1)
                                bsl = sl1.BUS
                                for bsl1 in bsl:
                                    if not bsl1.isInList(ba):
                                        ba.append(bsl1)
                                        flag = True
            listBus.append(ba)
        #
        res = []
        for i in range(len(listBus)):
            ba = listBus[i]
            if self.__testFunc4(ba,i,listBus):
                r1 = [ec,self.ga[i].toString()]
                res.append(r1)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','GenW3+GenW4+VCCS units that are connected with short lines or switches',len(res)])
#
class CHECK_LINE:
    def __init__(self,fwriter):
        self.fwriter = fwriter
    #
    def checkLineType(self):
        """
        Some line type field text contains leading and trailing blank spaces
        """
        ec = 'E31'
        self.fwriter.writerow([ec,'Some line type field text contains leading and trailing blank spaces'])
        title = [ec,'LINE ID','PARAMETER','VALUE','CHANGED TO']
        la = OLCase.LINE
        res = []
        for l1 in la:
            lt = l1.TYPE
            if lt:
                lt1 = lt.strip()
                if lt!=lt1 :
                    r1 = [ec,l1.toString(),'LINE TYPE',lt,lt1]
                    res.append(r1)
                    l1.TYPE = lt1
                    l1.postData()
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Some line type field text contains leading and trailing blank spaces',len(res)])
    #
    def checkLineRX(self):
        res = []
        title = ''
        ec = 'E32'
        self.fwriter.writerow([ec,'Some line have RX=0 or R0X0=0'])
        title = [ec,'LINE ID','PARAMETER','VALUE','CHANGED TO']
        la = OLCase.LINE
        res = []
        for l1 in la:
            r1 = l1.R
            x1 = l1.X
            r0 = l1.R0
            x0 = l1.X0
            if (abs(r1) <1e-6 and abs(x1)<1e-6) or (abs(r0) <1e-6 and abs(x0)<1e-6):
                vi = [ec,l1.toString(),'R X R0 X0']
                vi.append( str(round(r1,6))+' '+str(round(x1,6))+' '+str(round(r0,6))+' '+str(round(x0,6)) )
                if abs(r1) <1e-6 and abs(x1)<1e-6:
                    r1 = 1e-6
                    x1 = 1e-6
                    l1.R = r1
                    l1.X = x1
                if abs(r0) <1e-6 and abs(x0)<1e-6:
                    r0 = 1e-6
                    x0 = 1e-6
                    l1.R0 = r0
                    l1.X0 = x0
                vi.append( str(round(r1,6))+' '+str(round(x1,6))+' '+str(round(r0,6))+' '+str(round(x0,6)) )
                res.append(vi)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Some line have R_X=0 or R0_X0=0',len(res)])
#
class CHECK_XFMR:
    def __init__(self,fwriter):
        self.fwriter = fwriter
    #
    def checkConfig(self):
        ec = 'E1'
        listCheck = {'DGD'}
        self.fwriter.writerow([ec,'Check 2 Windings Transformer Config'])
        title = [ec,'OBJ ID','PARAMETER','VALUE','CHANGED TO']
        res = []
        for x2 in OLCase.XFMR:
            c1 = x2.CONFIGP+x2.CONFIGS+x2.CONFIGST
            if c1 in listCheck:
                vi = ['E1',x2.toString().ljust(65),'CONFIG',c1 ,'GGG']
                res.append(vi)
                x2.CONFIGP = 'G'
                x2.CONFIGS = 'G'
                x2.CONFIGST = 'G'
                x2.postData()
        #
        for x2 in OLCase.XFMR:
            configA = x2.CONFIGP
            configB = x2.CONFIGS
            configC = x2.CONFIGST
            tapA = x2.PRITAP
            tapB = x2.SECTAP
            c1 = configA+configB+configC
            # low - hight  =>  low - hight
            #   G   - D             G  - E
            if (tapA<tapB) and configA=="G" and configB=="D":
                x2.CONFIGS = 'E'
                x2.postData()
                vi = ['E1',x2.toString().ljust(65),'CONFIG',c1 ,configA+x2.CONFIGS+configC]
                res.append(vi)
            # low - small  =>  low - small
            #  G    - E              G - D
            if (tapA>tapB) and configA=="G" and configB=="E":
                x2.CONFIGS = 'D'
                x2.postData()
                vi = ['E1',x2.toString().ljust(65),'CONFIG',c1 ,configA+x2.CONFIGS+configC]
                res.append(vi)
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Check 2 Windings Transformer Config',len(res)])
#
class CHECK_MEMO:
    def __init__(self,fwriter):
        self.fwriter = fwriter
    #
    def __getNewMemo(m1):
        mx = m1.replace('\n',' ')
        mx = mx.replace('\r',' ')
        mx = mx.replace('\\t',' ')
        ma = mx.split()
        mn = ''
        for mi in ma:
            mn+=mi+ ' '
        return mn
    #
    def check(self):
        """
        Memo text contains leading and trailing blank spaces and tab characters
        """
        ec = 'E41'
        self.fwriter.writerow([ec,'Memo text contains leading and trailing blank spaces and tab characters'])
        title = [ec,'OBJ ID','PARAMETER','VALUE','CHANGED TO']
        res = []
        va = OLCase.getData()
        for key,val in va.items():
            if type(val)==list and key not in {'RLYOC','RLYDS','RECLSR'}:
                for v1 in val:
                    m1 = v1.MEMO
                    m2 = m1.strip()
                    if m1!=m2 or '\t' in m1:
                        r1 = [ec,v1.toString(),'MEMO']
                        mn = CHECK_MEMO.__getNewMemo(m1)
                        r1.append(m1)
                        r1.append(mn)
                        v1.MEMO = mn
                        v1.postData()
                        res.append(r1)
        #
        for key,val in va.items():
            if key =='RECLSR':
                for v1 in val:
                    ma = v1.MEMO
                    for m1 in ma:
                        m2 = m1.strip()
                        if m1!=m2 or '\t' in m1:
                            mn = [CHECK_MEMO.__getNewMemo(ma[0]), CHECK_MEMO.__getNewMemo(ma[1])]
                            v1.MEMO = mn
                            v1.postData()
                            r1 = [ec,v1.toString(),'MEMO']
                            for mi in ma:
                                mx = mi.replace('\n',' ')
                                mx = mx.replace('\r',' ')
                                r1.append(mx)
                            #
                            for mi in mn:
                                r1.append(mi)
                            res.append(r1)
                            break
        #
        writeRes(res,self.fwriter,title)
        self.fwriter.writerow(['COUNT','Memo text contains leading and trailing blank spaces and tab characters',len(res)])
#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return False
    #
    if type(ARGVS.op)!=list:
        op = [ARGVS.op]
    else:
        op = ARGVS.op
    vd = {0:'Check All',1:'-Check XFMR',2:'-GEN34 VCCS',3:'-Check Duplicate (CID, BusNumber)',4:'-Check Line',5:'-Check Memo'}
    for op1 in op:
        if op1 not in vd.keys():
            se ='ValueError of op='+str(op1)
            for k,v in vd.items():
                se+='\n\top='+str(k)+':'+str(v)
            raise Exception(se)
        print('\nCheck: '+str(vd[op1]))
    #
    load_olxapi(ARGVS.olxpath)
    OLCase.open(ARGVS.fi,1)
    fr = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='' , ad='_CheckNetwork_Report', ext='.CSV')
    f1 = open(fr, "w")
    fwriter = csv.writer(f1,quotechar="'", lineterminator="\n")
    fwriter.writerow(['','CheckNetwork version=%s Date:%s'%(__version__,time.asctime())])
    fwriter.writerow(['','OLR file:%s'% os.path.abspath(ARGVS.fi)])
    fwriter.writerow([])
    fwriter.writerow(['CCODE','Details'])
    #
    if (0 in op) or (1 in op):
        cx = CHECK_XFMR(fwriter)
        cx.checkConfig()
    #
    if (0 in op) or (2 in op):
        cx = CHECK_GENW34_VCCS(fwriter)
        cx.check1()
        cx.check2()
        cx.check3()
        cx.check4()
    #
    if (0 in op) or (3 in op):
        cx = CHECK_DUPLICATE(fwriter)
        cx.check_Loadunit()
        cx.check_Genunit()
        cx.check_Shuntunit()
        cx.checkBusNumber()
    #
    if (0 in op) or (4 in op):
        cx = CHECK_LINE(fwriter)
        cx.checkLineRX()
        cx.checkLineType()
    #
    if (0 in op) or (5 in op):
        cx = CHECK_MEMO(fwriter)
        cx.check()
    fo = AppUtils.get_file_out(fo=ARGVS.fo , fi=ARGVS.fi , subf='' , ad='_CheckNetwork', ext='.OLR')
    OLCase.save(fo)
    OLCase.close()
    f1.close()
    print('\nChecking report:',fr)
    print('Updated network:',fo)
    # AppUtils.launch_excel(fr)
    AppUtils.launch_notepad(fr)
    AppUtils.launch_OneLiner(fo,olxpath=ARGVS.olxpath)

#
def main():
    AppUtils.logger2File(PY_FILE,version = __version__)
    #
    global FRUNNING,SRUNNING
    FRUNNING,SRUNNING = AppUtils.markerStart(ARGVS.opath)
    if ARGVS.demo>0:
        return run_demo()
    run()
#
if __name__ == '__main__':
##    ARGVS.op = [0]
##    ARGVS.fi = 'sample\\CN01.OLR'
##    ARGVS.fi = 'sample\\XFMRConn.OLR'
    # ARGVS.fi = 'sample\\ieee9_w4.OLR'
    # ARGVS.fi = 'sample\\CN02.OLR'
    # ARGVS.fi = 'sample\\BIGTEST1_0Relay.OLR'
    # ARGVS.fi = 'sample\\BIGTEST2_0Relay.OLR'
    # ARGVS.fi = 'sample\\BIGTEST3_0Relay.OLR'
##    ARGVS.fi = 'sample\\BIGTEST4_0Relay.OLR'
##    ARGVS.fi = 'sample\\ECIM_5tierNew_anon_iniCN_deleteSw.OLR'
##    ARGVS.fi = 'sample\\S01.OLR'
##    ARGVS.fi = 'sample\\problem_x2.olr'
##    ARGVS.demo = 1
    try:
        main()
        AppUtils.markerSucces(ARGVS.opath)
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)

