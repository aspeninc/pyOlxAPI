"""
Purpose: PRC-023-4 Transmission Relay Loadability Check
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2022, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.7"

import os,sys
import csv,time,math
import logging
logger = logging.getLogger(__name__)
#
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
# default olxpath/olxpathpy
olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python'
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'PRC-023-4 Transmission Relay Loadability Check'
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-fCFG',help = '*(str) XML config file path (default=...\\PRC023_CONFIG.XML)',default = 'PRC023_CONFIG.XML',type=str, metavar='')
PARSER_INPUTS.add_argument('-olxpath'   , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located,',default = olxpathpy,type=str,metavar='')
PARSER_INPUTS.add_argument('-opath' , help = ' (str) Path name output folder',default = "", type=str, metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
import OlxObj
import AppUtils
import Substation
#
nprec = 2 # nRound in output
## value default, example for Line thermal rating Default = 'C', all acceptable value: ['A','B','C','D']
configDefault = {'Line thermal rating'  :['C',['A','B','C','D']],\
                 'Line 15-minute rating':['D',['A','B','C','D']],\
                 'Transformer nameplate rating':['MVA2',['MVA1','MVA2','MVA3']],\
                 'Transformer emergency rating':['MVA3',['MVA1','MVA2','MVA3']],\
                 'Series capacitor emergency rating':['RATING',[]],\
                 'TAGS':['PRC-023',[]],'AREA':['',[]],'ZONE':['',[]]}
#
kr = 0
setLhandle = set()
#
def run1line(l1,config):
    global kr,setLhandle
    if type(l1)!=OlxObj.TERMINAL:
        if l1.HANDLE in setLhandle:
            return dict()
    #
    res = OlxObj.OLCase.tapLineTool(l1,tapSCAP=True)
    mainLine = res['mainLine']
    allPath = res['allPath']
    remoteBus = res['remoteBus']
    z10 = res['Z1'][0].imag
    #
    for a1 in allPath:
        for t1 in a1:
            setLhandle.add(t1.EQUIPMENT.HANDLE)
    if len(mainLine)>1:
        return dict() # ignore for multi main line
    #
    ea = [] # list all EQUIPMENT (LINE/SWITCH/SERIESRC)
    for t1 in mainLine[0]:
        ea.append(t1.EQUIPMENT)
    #
    z1 = 0 # line inductive reactance
    haveSeriesrc = False
    for e1 in ea:
        if type(e1)!=OlxObj.SWITCH:
            x1 = e1.X
            if x1>0: # reduced by the capacitive reactance
                z1+=x1
            else:
                haveSeriesrc = True
    if abs(z1)<=1e-4:
        return dict() # ignore for short line
    #
    kr+=1
    b1 = mainLine[0][0].BUS[0]
    b2 = mainLine[0][0].BUS[1]
    b3 = remoteBus[0]
    kV = b1.kV
    cid = mainLine[0][0].CID
    sum1 = [kr,kV,b1.NO,b1.NAME,b2.NO,b2.NAME,cid,b3.NO,b3.NAME]
    #
    ad = [[kr,'EQUIPMENT']]
    #
    kname = ['SC' if type(e1).__name__=='SERIESRC' else type(e1).__name__  for e1 in ea]
    #
    for v1 in ['LINE','SWITCH','SC']:
        k = sum(x==v1 for x in kname)
        if k>1:
            i1=1
            for i in range(len(kname)):
                if kname[i]==v1:
                    kname[i]+=' '+str(i1)
                    i1+=1
    print('Checking:',mainLine[0][0].toString()[11:])
    for i in range(len(ea)):
        t1 = mainLine[0][i]
        #
        s1 = t1.toString()
        r1 = [kr,'',kname[i] + ': ' + s1[s1.find(']')+1:] + ' ' + ea[i].GUID ]
        ad.append(r1)
    #
    ad2 = []
    r11,r12 = 1e9,1e9
    #
    a1 = [[kr,'R1.1: Line Thermal rating (Rating %s)'%config['Line thermal rating']]]
    a2 = [[kr,'R1.2: Line 15-minute Rating (Rating %s)'%config['Line 15-minute rating']]]

    for i in range(len(ea)):
        if kname[i].startswith('LINE'):
            rateg = ea[i].RATG
            r1,r2 = rateg[2],rateg[3]
            a1.append( [kr,'',kname[i]+' (amps)',r1] )
            a2.append( [kr,'',kname[i]+' (amps)',r2] )
            if r1>0:
                r11 = min(r11,r1)
            if r2>0:
                r12 = min(r12,r2)
    r11 = round(r11*1.5,nprec) if r11<1e9 else 'N/A'
    r12 = round(r12*1.15,nprec) if r12<1e9 else 'N/A'
    ad2.extend(a1)
    ad2.append([kr,'','FACILITY R1.1=min(RATING)*1.5',r11])
    ad2.extend(a2)
    ad2.append([kr,'','FACILITY R1.2=min(RATING)*1.15',r12])
    #
    ad2.append([kr,'R1.3: Maximum Theoretical Power Transfer Limit'])
    #
    ad2.append([kr,'','Xl line reactance (pu)',round(z10,6)])
    #
    pmax1,pmax10,Xs,Xr,pmax2,pmax20 = theoreticalPowerTransferCapability(ea,kV,b1,b3,z1,z10)
    imax1 = pmax1/math.sqrt(3)/kV*1000 # i=S/kV/sqrt(3)*1000
    imax10 = pmax10/math.sqrt(3)/kV*1000
    imax2 = pmax2/math.sqrt(3)/kV*1000 # i=S/kV/sqrt(3)*1000
    imax20 = pmax20/math.sqrt(3)/kV*1000
    #
    ad2.append([kr,'','Pmax R1.3.1 (MW)',round(pmax10,nprec)])
    ad2.append([kr,'','Imax R1.3.1 (amps)',round(imax10,nprec)])
    if Xs<1e9:
        ad2.append([kr,'','Xs @'+b1.toString() +'(Ohm)',round(Xs,nprec)])
    else:
        ad2.append([kr,'','Xs @'+b1.toString() +'(Ohm)','infinite'])
    if Xr<1e9:
        ad2.append([kr,'','Xr @'+b3.toString()+'(Ohm)',round(Xr,nprec)])
    else:
        ad2.append([kr,'','Xr @'+b3.toString()+'(Ohm)','infinite'])
    ad2.append([kr,'','Pmax R1.3.2 (MW)',round(pmax20,nprec)])
    ad2.append([kr,'','Imax R1.3.2 (amps)',round(imax20,nprec)])
    if Xs<1e9 and Xr<1e9:
        r13 = min(imax10,imax20)
    else:
        r13 = imax10
    #
    r13 = round(r13*1.15,nprec)
    ad2.append([kr,'','FACILITY R1.3=min(Imax R1.3.1 , Imax R1.3.2)*1.15',r13])
    #
    r14 = 'N/A'
    ad2.append([kr,'R1.4: Series-compensated Line'])
    if haveSeriesrc:
        r14 = 1e9
        for i in range(len(ea)):
            if kname[i].startswith('SC'):
                try:
                    ri = float(ea[i].RATING)
                except:
                    ri = 0
                if ri>0:
                    r14 = min(r14,ri)
                    ad2.append( [kr,'',kname[i]+' RATING (amps)',ri] )
                else:
                    ad2.append( [kr,'',kname[i]+' RATING (amps)','N/A'] )
        #
        ad2.append([kr,'','Xl full line inductive reactance (pu)',round(z1,6)])
        ad2.append([kr,'','Pmax R1.4.2.1 (MW)',round(pmax1,nprec)])
        ad2.append([kr,'','Imax R1.4.2.1 (amps)',round(imax1,nprec)])
        #
        ad2.append([kr,'','Pmax R1.4.2.2 (MW)',round(pmax2,nprec)])
        ad2.append([kr,'','Imax R1.4.2.2 (amps)',round(imax2,nprec)])
        #
        if Xs<1e9 and Xr<1e9:
            r142 = min(imax1,imax2)
        else:
            r142 = imax1
        #
        r14 = max(r14,r142) if r14<1e9 else r142
        r14 = round(r14 *1.15,nprec)
        ad2.append([kr,'','FACILITY R1.4=MAX(MIN(SC RATING),MIN(Imax R1.4.2.1 , R1.4.2.2))*1.15',r14])
    else:
        ad2.append([kr,'','FACILITY R1.4',r14])
    #
    ad2.append([kr,'R1.5: Weak Source Systems'])
    # Ifault 3LG EOL @ 'BUS0' 220 kV
    ic,ti = weakSource(mainLine[0][0],ea[0])
    ad2.append([kr,'','IFault 3LG EOL @'+ti.toString()[11:] + ' (amps)',round(ic,nprec)])
    r15 = round(ic*1.7,nprec)
    if r15==0:
        r15 ='N/A'
    ad2.append([kr,'','FACILITY R1.5=IFault*1.7 (amps)',r15])
    #-------------------
    r16 = 'N/A'
    ad2.append([kr,'R1.6: Not used'])
    ad2.append([kr,'','FACILITY R1.6',r16])
    #
    MVAmaxG,ga1 = rate_generator(b1)
    r17 = 'N/A'
    ad2.append([kr,'R1.7: Generation Remote to Load'])
    if len(ga1)>0:
        for i in range(len(ga1)):
            s1 = ga1[i].toString()
            ad.append([kr,'','GENERATOR '+ str(i+1) + ': ' + s1[s1.find(']')+1:] + ' ' + ga1[i].GUID])
        imaxG = MVAmaxG/math.sqrt(3)/kV*1000 # i=S/kV/sqrt(3)*1000
        r17 = round(1.15*imaxG,nprec)
        ad2.append([kr,'','MVA max (MVA Rate *2) @'+b1.toString() +' (MVA)',MVAmaxG])
        ad2.append([kr,'','Imax (amps)',round(imaxG,nprec)])
        ad2.append([kr,'','FACILITY R1.7=Imax*1.15 (amps)',r17])
    else:
        ad2.append([kr,'','FACILITY R1.7',r17])
    #
    r18 = 'N/A'
    ad2.append([kr,'R1.8: Bulk system-end of transmission lines, Load remote to the system'])
    ad2.append([kr,'','FACILITY R1.8 ',r18])
    r19 = 'N/A'
    ad2.append([kr,'R1.9: Load-end of transmission lines, Load remote to the bulk system'])
    ad2.append([kr,'','FACILITY R1.9 ',r19])
    # TRANSFORMER
    ta = b3.TERMINAL
    x = None
    if len(ta)==2:
        for ti in ta:
            e1 = ti.EQUIPMENT
            if type(e1) in {OlxObj.XFMR,OlxObj.XFMR3}:
                x = e1
    r110 = 'N/A'
    ad2.extend([[kr,'R1.10: Transformer Overcurrent Protection']])
    if x :
        s1 = x.toString()
        ad.append([kr,'','TRANSFORMER: ' + s1[s1.find(']')+1:] + ' ' + x.GUID])
        mva2 = x.MVA2
        mva3 = x.MVA3
        im2 = mva2 /math.sqrt(3)/kV *1000
        im3 = mva3 /math.sqrt(3)/kV *1000
        ad2.append([kr,'','MVA nameplate rating (%s)'%config['Transformer nameplate rating'],mva2])
        ad2.append([kr,'','Imax1 (nameplate rating as amps)',round(im2,nprec)])
        ad2.append([kr,'','MVA emergency rating (%s)'%config['Transformer emergency rating'],mva3])
        ad2.append([kr,'','Imax2 (Emergency Rating as amps)',round(im3,nprec)])
        r110 = 0
        if im2>0:
            r110 = max(im2*1.5,r110)
        if im3>0:
            r110 = max(im3*1.15,r110)
        #
        r110 =round(r110,nprec) if r110>0 else 'N/A'
        ad2.append([kr,'','FACILITY R1.10 = MAX(Imax1*1.5,Imax2*1.15) (amps)',r110])
    else:
        ad2.append([kr,'','FACILITY R1.10',r110])
    #
    ad.extend(ad2)
    sum1.extend([r11,r12,r13,r14,r15,r16,r17,r18,r19,r110])
    # for the opposite
    t2 = mainLine[0][-1].OPPOSITE[0]
    #


    return {'sum1':sum1,'details1':ad,'topposite':t2}
#
def rate_generator(b1):
    """ Set transmission line relays applied at the load center terminal, remote from generation
        stations, so they do not operate at or below 115% of the maximum current flow from the load
        to the generation source under any system configuration
        mvaMax = 2 * MVA rate

        R1.6 — Generation Remote to Load
            page 14-16/36 'pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf'
    """
    sub1 = Substation.substation(b1,leng=0.0)
    ga1 = sub1.GEN
    if len(ga1)==0:
        return 0,[]
    #
    for t1 in sub1.TERMINAL_EXT:
        if t1.FLAG==1:#in service
            res = OlxObj.OLCase.tapLineTool(t1,tapSCAP=True)
            remoteBus = res['remoteBus']
            for b5 in remoteBus:
                sub5 = Substation.substation(b5,leng=0.0)
                if sub5.GEN:
                    return 0,[]
    #
    mr1 = 0
    for g1 in ga1:
        mr1 += g1.MVARATE
    # page 15/36 'pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf'
    return mr1*2,ga1
#
def theoreticalPowerTransferCapability(ea,kV,b1,b3,z1,z10):
    """
    return:
    iCapab1:
        An infinite source (zero source impedance) with a 1.00 per unit bus voltage at each end of the line.
        using full line inductive reactance
        page 7/36: 'pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf'

    iCapab2:
        An impedance at each end of the line, which reflects the actual system source impedance
        with a 1.05 per unit voltage behind each source impedance
        page 10/36: 'pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf'

    For R1.5 page 13/36 'pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf'
    ic1, ic2: three-phase fault current at 2 ends of the line (remove the line under study)
    """
    pmax10 = 1.0/z10 * OlxObj.OLCase.BASEMVA
    pmax1 = 1.0/z1 * OlxObj.OLCase.BASEMVA
    # Remove the line or lines under study (parallel lines need to be removed prior to doing the fault study)
    o1 = OlxObj.OUTAGE(option='ALL',G=0)
    for e1 in ea:
        if type(e1).__name__=='LINE':
            o1.add_outageLst(e1)
    if len(o1.outageLst)==0:
        o1=None
    #
    fs1 = OlxObj.SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='3LG',Z=[0,0],outage=o1)
    OlxObj.OLCase.simulateFault(fs1,1,verbose=False)
    xS = OlxObj.FltSimResult[0].Thevenin[0].imag
    if abs(xS)>10000: # line end without gen, ic=0
        xS=1e9
    #
    fs2 = OlxObj.SPEC_FLT.Classical(obj=b3,fltApp='BUS',fltConn='3LG',Z=[0,0],outage=o1)
    OlxObj.OLCase.simulateFault(fs2,1,verbose=False)
    xR = OlxObj.FltSimResult[0].Thevenin[0].imag
    if abs(xR)>10000: # line end without gen, i=0
        xR=1e9
    #
    x1Ohm = z1 * kV*kV/OlxObj.OLCase.BASEMVA
    x1Ohm0 = z10 * kV*kV/OlxObj.OLCase.BASEMVA
    #
    pmax20 = (1.05*kV)**2/(xS+xR+x1Ohm0)
    pmax2 = (1.05*kV)**2/(xS+xR+x1Ohm)
    #
    return pmax1,pmax10,xS,xR,pmax2,pmax20
#
def weakSource(t1,e1):
    """
    In some cases, the maximum line end three-phase fault current is small relative to the thermal
    loadability of the conductor. Such cases exist due to some combination of weak sources, long
    lines, and the topology of the transmission system
    return
        ic: is the line-end three-phase fault current magnitude obtained from a short circuit
            study, reflecting sub-transient generator reactances
    (page 13/36 pa_Stand_PRC0231RD_Relay_Loadability_Reference_Doc_Clean_Final_2008July3.pdf)
    """
    ta = e1.TERMINAL
    if ta[0].__hnd__== t1.__hnd__:
        ti = ta[1]
    else:
        ti = ta[0]
    fs = OlxObj.SPEC_FLT.Classical(obj=ti,fltApp='Line-end',fltConn='3LG',Z=[0,0])
    OlxObj.OLCase.simulateFault(fs,1,verbose=False)
    ic = abs(OlxObj.FltSimResult[0].current()[0])
    return ic,ti
#
def writeCSV(fn,ha,ar):
    f1 = open(fn, "w")
    fwriter = csv.writer(f1,quotechar='"', lineterminator="\n")
    for a1 in ha:
        fwriter.writerow(a1)
    for a1 in ar:
        fwriter.writerow(a1)
    f1.close()
#
def readConfigFile(fCFG):
    from xml.etree.cElementTree import iterparse
    config = dict()
    if os.path.isfile(fCFG):
        source = open(fCFG,"rb")
        ipa = iterparse(source, ("start","end"))
        for (event, elem) in ipa:
            if event == "start" and elem.tag=="ROW":
                name = elem.attrib['NAME']
                val = elem.attrib['VALUE']
                config[name]=val
            if event == "end" and elem.tag=="PRC023_CONFIG":
                break
        source.close()
    else:
        for k,v in configDefault.items():
            config[k]= v[0]
    se = '\nFormat of XML config file:\n'+fCFG
    for k,v in configDefault.items():
        if k not in config.keys():
            se+= '\n\tkeys not found: '+k
            raise Exception(se)
        if v:
            if v[1] and (config[k] not in v[1]):
                se +='\n\t'+k+': '+str(v[1])+'\n\tFound: '+config[k]
                raise Exception(se)
    return config
#
def main():
    AppUtils.logger2File(PY_FILE,version = __version__)
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return False
    #
    ARGVS.fi = os.path.abspath(ARGVS.fi)
    ARGVS.fCFG = os.path.abspath(ARGVS.fCFG)
    #
    config = readConfigFile(ARGVS.fCFG)
    ## load api
    v = OlxObj.load_olxapi(ARGVS.olxpath)
    ## open file
    OlxObj.OLCase.open(ARGVS.fi,1)
    ## scope
    try:
        OlxObj.OLCase.applyScope(areaNum = config['AREA'],zoneNum = config['ZONE'],kV =[200,1000],verbose=False)
    except Exception as err:
        raise Exception('\nInvalid data n config file: '+ARGVS.fCFG+str(err))
    #
    la = OlxObj.OLCase.LINE
    # for line 0-200kV with Tags=config['TAGS']
    if config['TAGS']:
        OlxObj.OLCase.applyScope(areaNum = config['AREA'],zoneNum = config['ZONE'],kV =[0,200],verbose=False)
        ltags = OlxObj.OLCase.findOBJByTag(config['TAGS'],['LINE'])
        la.extend(ltags)
    #
    ar0 = [] # summary
    ard = [] # details
    for l1 in la:
        if l1.FLAG==1:
            r1 = run1line(l1,config)
            if r1:
                ar0.append(r1['sum1'])
                ard.extend(r1['details1'])
                t2 = r1['topposite']
                #
                r1 = run1line(t2,config)
                ar0.append(r1['sum1'])
                ard.extend(r1['details1'])
                #break
    #
    header = [  ['PRC-023 FACILITY LOADABILITY LIMITS'],\
        ['ASPEN OneLiner V%s Build:%s'%(v[0],v[1])],\
        ['Date',time.asctime()],\
        ['OLR File',ARGVS.fi],\
        ['Config File',(ARGVS.fCFG if os.path.isfile(ARGVS.fCFG) else 'None')],\
        ['CONFIGURATIONS:']]
    for k,v in config.items():
        header.append(['',k+' : '+v])
    header.append([])
    header.append(['FACILITY','SUMMARY'])
    header.append(['NO','KV','BUSNO1','BUSNAME1','BUSNO2','BUSNAME2','CKTID','REMOTEBUSNO','REMOTEBUSNAME','R1.1','R1.2','R1.3','R1.4','R1.5','R1.6','R1.7','R1.8','R1.9','R1.10','R1.11','R1.12','R1.13'])
    #
    opath = ARGVS.opath if ARGVS.opath else os.path.dirname(ARGVS.fi)+'\\PRC023_Results'
    #
    fsum = opath+'\\PRC023_1Summary_'+os.path.basename(ARGVS.fi)
    fdet = opath+'\\PRC023_2Details_'+os.path.basename(ARGVS.fi)
    #
    fsum = AppUtils.get_file_out(fo=fsum , fi='' , subf='' , ad='', ext='.CSV')
    fdet = AppUtils.get_file_out(fo=fdet , fi='' , subf='' , ad='', ext='.CSV')
    #
    writeCSV(fsum,header,ar0) # file summary
    header[-2]= ['FACILITY','DETAILS']
    writeCSV(fdet,header[:-1],ard) #file details
    #
    print('CSV summary report:',fsum)
    print('CSV details report:',fdet)
#
if __name__ == '__main__':
    #ARGVS.fi = 'sample\\LS1.OLR'
    #ARGVS.fi = 'sample\\PRC023.OLR'
    main()
    logging.shutdown()


