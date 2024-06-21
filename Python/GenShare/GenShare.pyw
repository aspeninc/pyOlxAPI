"""
Purpose:
  Generator fault contribution calculator
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2023, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__email__     = "support@aspeninc.com"
__status__    = "in Dev"
__version__   = "1.1.4"

import sys,os,time,csv
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
olxpath = PATH_LIB+'\\1lpfV15'
olxpathpy = PATH_LIB
#
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = 'ASPEN Generator Cluster Fault Contribution Calculator'
PARSER_INPUTS.add_argument('-fb'  , help = ' (str) Breaker bus list file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-fi'  , help = ' (str) OneLiner file path', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-tier', help = ' (int) Number of tiers around selected Buses', default = 2,type=int,metavar='')
PARSER_INPUTS.add_argument('-pcnt', help = ' (float) Pcnt duty threshold (with fb)', default = 90.0,type=float,metavar='')
PARSER_INPUTS.add_argument('-olxpath' , help = ' (str) Full pathname of the folder, where the ASPEN olxapi.dll is located',default = olxpath,type=str,metavar='')
PARSER_INPUTS.add_argument('-olxpathpy' , help = ' (str) Full pathname of the folder where the OlxAPI Python wrapper OlxAPI.py and relevant libraries are located',default = olxpathpy,type=str,metavar='')
PARSER_INPUTS.add_argument('-fo'  , help = ' (str) Output csv file path', default = '',type=str,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]
sys.path.insert(0,ARGVS.olxpathpy)
from OlxObj import *
import AppUtils
#
import threading
import tkinter as tk
import tkinter.filedialog as tkf
import tkinter.messagebox as tkm
import logging
logger = logging.getLogger(__name__)
fz1 = ("Arial", 13)
fz1b = ("Arial bold", 13)
#
def runFault(b1):
    fs1 = SPEC_FLT.Classical(obj=b1,fltApp='BUS',fltConn='3LG')
    OLCase.simulateFault(fs1,1)
    i3p = FltSimResult[0].current()[0]
    fs1.fltConn = '1LG:A'
    OLCase.simulateFault(fs1,1)
    i1p = FltSimResult[0].current()[0]
    return i3p,i1p
#
class PG():
    def __init__(self):
        self.popup = tk.Toplevel(bg='SystemWindow')# SystemWindow sky blue
        _window_w = 308
        _window_h = 95
        sw = self.popup.winfo_screenwidth()
        sh = self.popup.winfo_screenheight()
        self._window_x = int(sw/2-_window_w/2)
        self._window_y = int(sh/2-_window_h/2)
        self.popup.geometry('{w}x{h}+{x}+{y}'.format(w=_window_w,h=_window_h,x=self._window_x,y=self._window_y))
        #self.popup.wm_title("ASPEN")
        self.popup.overrideredirect(True)
        #
        self.percent = [tk.DoubleVar(),tk.DoubleVar()]
        #
        fr1 = tk.Frame(self.popup,bg='SystemWindow')
        fr1.grid(row=0,column=0, padx=1, pady=1)
        fr2 = tk.Frame(self.popup,bg='SystemWindow')
        fr2.grid(row=1,column=0, padx=1, pady=1)
        lb = tk.Label(fr1, text="Running",font=fz1b,bg='SystemWindow')
        lb.grid(row=0,column=0, padx=1, pady=1,sticky='NS')
        #
        tk.Label(fr2, text="Find Neibor ",font=fz1,bg='SystemWindow').grid(row=0,column=0,pady=2)
        tk.Label(fr2, text="Run Fault   ",font=fz1,bg='SystemWindow').grid(row=1,column=0,pady=2)
        #
        self.style=[0]*2
        self.style[0] = tk.ttk.Style(self.popup)
        self.style[0].layout('text.Horizontal.TProgressbar0',[('Horizontal.Progressbar.trough',{'children': [('Horizontal.Progressbar.pbar',\
                        {'side': 'left', 'sticky': 'ns'})],'sticky': 'nswe'}),('Horizontal.Progressbar.label', {'sticky': 'nswe'})])
        self.style[0].configure('text.Horizontal.TProgressbar0', text='', anchor='center',font=fz1)
        #
        self.style[1] = tk.ttk.Style(self.popup)
        self.style[1].layout('text.Horizontal.TProgressbar1',[('Horizontal.Progressbar.trough',{'children': [('Horizontal.Progressbar.pbar',\
                        {'side': 'left', 'sticky': 'ns'})],'sticky': 'nswe'}),('Horizontal.Progressbar.label', {'sticky': 'nswe'})])
        self.style[1].configure('text.Horizontal.TProgressbar1', text='', anchor='center',font=fz1)
        #
        self.pbar = [0]*2
        self.pbar[0] = tk.ttk.Progressbar(fr2,style='text.Horizontal.TProgressbar0',orient=tk.HORIZONTAL,mode='determinate',\
                        maximum=100, variable=self.percent[0],length=200)
        self.pbar[0].grid(row=0, column=1,padx=5,pady=5)
        #
        self.pbar[1] = tk.ttk.Progressbar(fr2,style='text.Horizontal.TProgressbar1',orient=tk.HORIZONTAL,mode='determinate',maximum=100, variable=self.percent[1],length=200)
        self.pbar[1].grid(row=1, column=1,padx=5,pady=5)
        #
        self.popup.pack_slaves()
        self.popup.bind('<Button-1>',self.clickwin)
        self.popup.bind('<B1-Motion>',self.dragwin)
    #
    def setVal(self,ip,percent):
        self.popup.update()
        self.style[ip].configure('text.Horizontal.TProgressbar'+str(ip),text=str(percent)+" %")
        self.percent[ip].set(percent)
    #
    def flush(self):
        return
    #
    def dragwin(self,event):
        delta_x = self.popup.winfo_pointerx() - self._offsetx
        delta_y = self.popup.winfo_pointery() - self._offsety
        x = self._window_x + delta_x
        y = self._window_y + delta_y
        self.popup.geometry("+{x}+{y}".format(x=x, y=y))
        self._offsetx = self.popup.winfo_pointerx()
        self._offsety = self.popup.winfo_pointery()
        self._window_x = x
        self._window_y = y
    #
    def clickwin(self,event):
        self._offsetx = self.popup.winfo_pointerx()
        self._offsety = self.popup.winfo_pointery()
#
class MainGUI(tk.Frame):
    def __init__(self,master):
        tk.Frame.__init__(self, master=master)
        self.master.resizable(0,0)# fixed size
        self._offsetx = 0
        self._offsety = 0
        self._window_w = 875
        self._window_h = 300
        self.sw = self.master.winfo_screenwidth()
        self.sh = self.master.winfo_screenheight()
        self._window_x = int(self.sw/2-self._window_w/2)
        self._window_y = int(self.sh/2-self._window_h/2)
        self.master.geometry("{0}x{1}+{2}+{3}".format(self._window_w,self._window_h,self._window_x,self._window_y))
        self.master.wm_title(PARSER_INPUTS.usage+' V%s'%__version__[:3])
        #
        try:
            icon = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\Python\\1LINER.ICO'
            if not os.path.isfile(icon):
                icon = AppUtils.get_opath('')+'\\1LINER.ICO'
                if not os.path.isfile(icon):
                    import icoextract
                    icoextract.IconExtractor(sys.executable).export_icon(icon)
            self.master.wm_iconbitmap(icon)
        except:
            pass
        # registry
        self.reg = AppUtils.WIN_REGISTRY(path = "SOFTWARE\ASPEN\OneLiner\GenShare",keyUser="",nmax =10)
        self.initGUI()
        #self.master.attributes("-topmost", True)
        #self.master.overrideredirect(True)
        #self.master.bind('<Button-1>',self.clickwin)
        #self.master.bind('<B1-Motion>',self.dragwin)
        AppUtils.remove_button(self.master)
    #
    def dragwin(self,event):
        delta_x = self.master.winfo_pointerx() - self._offsetx
        delta_y = self.master.winfo_pointery() - self._offsety
        x = self._window_x + delta_x
        y = self._window_y + delta_y
        self.master.geometry("+{x}+{y}".format(x=x, y=y))
        self._offsetx = self.master.winfo_pointerx()
        self._offsety = self.master.winfo_pointery()
        self._window_x = x
        self._window_y = y
    #
    def clickwin(self,event):
        self._offsetx = self.master.winfo_pointerx()
        self._offsety = self.master.winfo_pointery()
    #
    def write(self, txt):
        self.text1.insert(tk.INSERT,txt)
        #
    def close_buttons(self):
        logging.shutdown()
        self.saveRegistry()
        try:
            self.popup.destroy()
        except:
            pass
        self.master.destroy()
    #
    def clearConsol(self):
        self.text1.delete(1.0,tk.END)
    #
    def help(self):
        mes = "Purpose: compute generator contribution to fault in OneLiner network"
        mes+= "\n\nMethodology: superposition theorem"
        mes+= "\n\nInputs:"
        mes+= "\n    OneLiner OLR network"
        mes+= "\n    Breaker rating csv report:"
        mes+= "\n        breaker list"
        mes+= "\n        Worst-case faults list (from MAX_SC_CASE) with branch outage details (from FAULT DESCRIPTION TABLE) for each breaker"
        mes+= "\n\nReport: csv file"
        AppUtils.gui_info(PARSER_INPUTS.usage+' V%s'%__version__[:3],mes)
    #
    def saveRegistry(self):
        self.reg.updateValue('OutFile',self.ipOut.get())
        self.reg.updateValue('OLRFile',self.ipOLR.get())
        self.reg.updateValue('BKFile',self.ipBK.get())
        self.reg.updateValue('OutFile',self.ipOut.get())
        self.reg.updateValue('Pcnt',self.ipPcnt.get())
        self.reg.updateValue('Tier',self.ipTier.get())
    #
    def initGUI(self):
        self.ipOLR = tk.StringVar()
        self.ipOLR.set(self.reg.getValue('OLRFile').replace('/','\\'))
        #
        self.ipBK = tk.StringVar()
        self.ipBK.set(self.reg.getValue('BKFile').replace('/','\\'))
        #
        self.ipOut = tk.StringVar()
        self.ipOut.set(self.reg.getValue('OutFile').replace('/','\\'))
        #
        self.ipTier = tk.StringVar()
        self.ipTier.set(self.reg.getValue('Tier','5'))
        #
        self.ipPcnt = tk.StringVar()
        self.ipPcnt.set(self.reg.getValue('Pcnt','90'))
        #
        fr0 = tk.LabelFrame(self.master, text = "",bg='SystemWindow')
        fr0.grid(row=0,column=0, padx=1, pady=1,sticky="W")
##        fr1 = tk.Frame(fr0,bg='SystemWindow')
##        fr1.grid(row=0,column=0, padx=0, pady=10,sticky="W")
##        ti = tk.Label(fr1,width =60,anchor='w',bg='SystemWindow', text=PARSER_INPUTS.usage+' V%s'%__version__[:3],font=fz1b)
##        ti.grid(row=0,column=0, padx=0, pady=0,sticky="W")
        #
        fr10 = tk.Frame(fr0,bg='SystemWindow')
        fr10.grid(row=1,column=0, padx=0, pady=0,sticky="W")

        fr2 = tk.Frame(fr10,bg='SystemWindow')
        fr2.grid(row=1,column=0, padx=0, pady=15,sticky="WN")
        ip1 = tk.Label(fr2, text="Breaker bus list file:",font=fz1,bg='SystemWindow')
        ip1.grid(row=0, column=0, sticky='W', padx=5, pady=0)
        ip2 = tk.Label(fr2, text="",font=fz1,bg='SystemWindow')
        ip2.grid(row=1, column=0, sticky='W', padx=5, pady=10)
        #
        ip3 = tk.Label(fr2, text="OneLiner network:  ",font=fz1,bg='SystemWindow')
        ip3.grid(row=2, column=0, sticky='W', padx=5, pady=15)
        ip4 = tk.Label(fr2, text="",font=fz1,bg='SystemWindow')
        ip4.grid(row=3, column=0, sticky='W', padx=5, pady=10)
        ip5 = tk.Label(fr2, text="Output report (.csv):",font=fz1,bg='SystemWindow')
        ip5.grid(row=4, column=0, sticky='W', padx=5, pady=5)
        #
        fr3 = tk.Frame(fr10,bg='SystemWindow')
        fr3.grid(row=1,column=1, padx=0, pady=15,sticky="WN")
        ip11 = tk.Entry(fr3,width= 70,textvariable=self.ipBK,font=fz1,bd=2,bg='SystemWindow')
        ip11.grid(row=0, column=0, sticky="W",padx=5, pady=0)
        ip12 = tk.Button(fr3, text="...",width= 6,relief= tk.GROOVE,command=self.selectBKFile)
        ip12.grid(row=0, column=1, sticky='W', padx=5, pady=0)
        #
        fr4 = tk.Frame(fr3,bg='SystemWindow')
        fr4.grid(row=1,column=0, padx=0, pady=5,sticky="E")
        #
        ip13 = tk.Label(fr4, text="Select only buses with breakers having percent duty of at least:",font=fz1,bg='SystemWindow')
        ip13.grid(row=0, column=0, sticky='E', padx=5, pady=10)
        ip13 = tk.Entry(fr4,width= 5,textvariable=self.ipPcnt,font=fz1,bd=2)
        ip13.grid(row=0, column=1, sticky="W",padx=5, pady=5)
        #
        ip14 = tk.Entry(fr3,width= 70,textvariable=self.ipOLR,font=fz1,bd=2,bg='SystemWindow')
        ip14.grid(row=2, column=0, sticky="W",padx=5, pady=5)
        ip15 = tk.Button(fr3, text="...",width= 6,relief= tk.GROOVE,command=self.selectOLRFile)
        ip15.grid(row=2, column=1, sticky='W', padx=5, pady=5)
        #
        fr5 = tk.Frame(fr3,bg='SystemWindow')
        fr5.grid(row=3,column=0, padx=0, pady=5,sticky="E")
        ip9 = tk.Label(fr5, text="Report contribution from all generators within",font=fz1,bg='SystemWindow')
        ip9.grid(row=0, column=0, sticky='E', padx=5, pady=10)
        ip10 = tk.Entry(fr5,width= 5,textvariable=self.ipTier,font=fz1,bd=2)
        ip10.grid(row=0, column=1, sticky="W",padx=5, pady=5)
        ip11 = tk.Label(fr5, text="tiers from each breaker bus",font=fz1,bg='SystemWindow')
        ip11.grid(row=0, column=2, sticky='E', padx=5, pady=5)
        #
        ip13 = tk.Entry(fr3,width= 70,textvariable=self.ipOut,font=fz1,bd=2,bg='SystemWindow')
        ip13.grid(row=4, column=0, sticky="W",padx=5, pady=5)
        ip14 = tk.Button(fr3, text="...",width= 6,relief= tk.GROOVE,command=self.selectOutFile)
        ip14.grid(row=4, column=1, sticky='W', padx=5, pady=5)
        #
        fr8 = tk.Frame(fr0,bg='SystemWindow')
        fr8.grid(row=5,column=0, padx=50, pady=11,sticky="W")
        #
        ip15 = tk.Label(fr8, text="",font=fz1,bg='SystemWindow')
        ip15.grid(row=0, column=1, sticky='W', padx=80, pady=5)

        self.runBt = tk.Button(fr8, text="OK",font=fz1,width=20, command=self.simulate)
        self.runBt.grid(row=0,column=2, padx=25, pady=0)
        #
        cancelBt = tk.Button(fr8, text="Cancel",width= 10,font=fz1, command=self.close_buttons)
        cancelBt.grid(row=0,column=3, padx=25, pady=0)
        self.helpBt = tk.Button(fr8, text="Help",width= 10,font=fz1, command=self.help)
        self.helpBt.grid(row=0,column=4, padx=25, pady=0)
    #
    def reset(self):
        self.progress.stop()
        self.pgFrame.grid(row=4,column=10, padx=10, pady=0,sticky='W')
        self.btFrame.grid(row=4,column=0, padx=10, pady=0)
    #
    def simulate(self):
        self.runBt['state']='disabled'
        self.helpBt['state']='disabled'
        try:
            pbar = PG()
            pbar.setVal(0,0)
            if os.path.isfile('olxapi.dll'):
                olxpath = os.getcwd()
            elif os.path.isfile(PATH_LIB + '\\1lpfV15\\olxapi.dll'):
                olxpath = PATH_LIB + '\\1lpfV15'
            else:
                olxpath = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
            if not os.path.isfile(olxpath+'\\olxapi.dll'):
                raise Exception('Failed to locate OlxAPI.dll')
            #
            fb = self.ipBK.get()
            fi = self.ipOLR.get()
            fo = self.ipOut.get()
            tier = self.ipTier.get()
            pcnt = self.ipPcnt.get()
            #
            gs = GENSHARE(pbar)
            gs.setOlxPath(olxpath)
            gs.setInput(fb,fi,pcnt,tier,fo)
            gs.findNeibor()
            gs.runFault()
            #
            gs.saveResult()
            self.ipOut.set(gs.fo)
            #
            mes = "Calculation completed for %i faults of %i buses in the list.\n\nReport is in: %s.\n\nDo you want to open it in the spreadsheet program?"%(gs.nf,len(gs.lstBus),gs.fo)
            choice = AppUtils.gui_askquestion(PARSER_INPUTS.usage+' V%s'%__version__[:3], mes)
            if choice=='yes':
                os.system('start "excel.exe" "{0}"'.format(gs.fo))
            self.close_buttons()
        except Exception as er:
            se=str(er).strip()
            if se.endswith("command: application has been destroyed"):
                logging.error('Cancel by user')
            else:
                logging.error(se)
                AppUtils.gui_error(PARSER_INPUTS.usage+' V%s'%__version__[:3],se)
            self.runBt['state']='active'
            self.helpBt['state']='active'
            logging.shutdown()
        #
        try:
            pbar.popup.destroy()
        except:
            pass
    #
    def selectOLRFile(self):
        # OLR file
        v1 = tkf.askopenfilename(filetypes=[("OneLiner Files", "*.OLR *.olr")],title='Select OneLiner file')
        if v1!='':
            v1 = v1.replace('/','\\')
            self.ipOLR.set(v1)
            self.reg.updateValue('OLRFile',v1)
    #
    def selectBKFile(self):
        # Breaker rating report file
        v1 = tkf.askopenfilename(filetypes=[('Breaker list file','*.CSV *.csv'),("Breaker List Files", "*.CSV *.csv")],title='Select breaker list file')
        if v1!='':
            v1 = v1.replace('/','\\')
            self.ipBK.set(v1)
            self.reg.updateValue('BKFile',v1)
    #
    def selectOutFile(self):
        # Out file
        v1 = tkf.asksaveasfile(defaultextension=".csv",filetypes=(("Report file", "*.CSV *.csv"),("Report File", "*.CSV *.csv")),title='Save As Output report')
        try:
            v1 = v1.name.replace('/','\\')
            self.ipOut.set(v1)
            self.reg.updateValue('OutFile',v1)
        except:
            pass
    #
    def flush(self):
        return
class GENSHARE:
    def __init__(self,pbar=None):
        self.t0 = time.time()
        self.p0 = 0
        self.pbar = pbar
    #
    def setOlxPath(self,olxpath):
        self.ver = load_olxapi(olxpath)
    #
    def setInput(self,fb,fi,pcnt,tier,fo):
        #
        if fb.upper()==fo.upper():
            fo = fo[:-4]+'__GenShare.csv'
        self.fo = AppUtils.get_file_out(fo=fo , fi=fi , subf='res' , ad='', ext='.csv')
        #
        flog = AppUtils.get_file_out(fo=os.path.split(self.fo)[0]+'\\GenShare.log', fi=fi , subf='res' , ad='', ext='.log')
        AppUtils.logger2File(PY_FILE,flog = flog,version = __version__)
        logging.info('\n'+PARSER_INPUTS.usage+' V%s'%__version__[:3])
        #
        if not os.path.isfile(fb):
            raise Exception('Breaker bus list file not found: '+fb)
        if not os.path.isfile(fi):
            raise Exception('OneLiner file not found: '+fi)
        try:
            valtier = int(tier)
        except:
            valtier = 0
        if valtier<=0:
            raise Exception('Cluster size (tiers) must be an integer positive, found : '+str(tier))
        try:
            pcntVal = float(pcnt)
        except:
            pcntVal = 0
        if pcntVal<=0:
            raise Exception('Pcnt duty threshold be a float positive, found : '+str(pcnt))
        #
        self.fb = os.path.abspath(fb)
        self.fi = os.path.abspath(fi)
        self.tier = valtier
        self.pcnt = pcntVal
        self.busNotFound = []
        #
        logging.info('\tBreaker list file   : '+self.fb)
        logging.info('\tOneLiner file       : '+self.fi)
        logging.info('\tPcnt duty threshold : '+toString(self.pcnt)+' %')
        logging.info('\tCluster size (tiers): %i\n'%self.tier)
        #
        OLCase.open(fi,1)
        # set PreFault Voltage =>From linear Network solution
        OLCase.setSystemParams({'nPrefaultV': '0'})
        #
        self.lstBus = []
        with open(fb, mode= 'r') as f:
            k1,k2 =0,0
            reader = csv.reader(f, delimiter=',')
            ki=0
            kf=0
            for row in reader:
                ki+=1
                if len(row)>2 and row[0]=='' and row[1]=='' or len(row)==0:
                    break
                if k1>0 and k2>0:
                    try:
                        sn1 = row[1]
                        id1 = sn1.rindex(' ')
                        va = '[BUS] '+"'"+sn1[:id1]+"'"+sn1[id1:]
                        b1 = OLCase.findBUS(va)
                    except:
                        b1 = None
                    if b1==None:
                        self.busNotFound.append('Line number %i : '%ki+str(row[:2]))
                    else:
                        kf+=1
                        try:
                            v1 = float(row[k1])
                        except:
                            v1 = 0
                        #
                        try:
                            v2 = float(row[k2])
                        except:
                            v2 = 0
                        #
                        if self.pcnt<v1 or self.pcnt<v2:
                            if not b1.isInList(self.lstBus):
                                self.lstBus.append(b1)
                #
                if len(row)>2 and row[0]=='BUS_NO' and row[1]=='BUS':
                    for i in range(len(row)):
                        if row[i]=='DUTY_P':
                            k1=i
                        if row[i]=='M_DUTY_P':
                            k2=i
        #for b1 in self.lstBus:
        #    print(b1.toString())
        OLCase.setVerbose(0)
        logging.info(str(kf))
        logging.info('Found %i buses'%len(self.lstBus))
        if k1==0 or k2==0:
            raise Exception('Invalid breaker list format')
        elif k1>0 and kf==0:
            raise Exception('No breaker bus data can be found in the input file')
        elif kf>0 and len(self.lstBus)==0:
            raise Exception('No breaker bus data with Pcnt duty threshold>%s can be found in the input file'%toString(self.pcnt))
    #
    def updateLabel(self,k,nf,ip,s0=''):
        if self.pbar is not None:
            p = int(k/nf*100)
            if p!=self.p0:
                self.p0=p
                self.pbar.setVal(ip,p)
        else:
            p = int(k/nf*10)
            if p!=self.p0:
                self.p0=p
                logging.info(s0+(str(k)+'/'+str(nf)).ljust(8)+str(int(k/nf*100)).ljust(3)+'%')
    #
    def findNeibor(self,pg=None):
        t0 = time.time()
        self.genNeibor = []
        self.busWithGen = dict()
        p0,k=0,0
        nf = len(self.lstBus)
        logging.info('\nFinding Neibor (tier=%i)'%self.tier)
        for b1 in self.lstBus:
            k+=1
            self.updateLabel(k,nf,0,'Find Neibor : ')
            #
            bnei = b1.findBusNeibor(self.tier) # can be improve with new API
            #
            gn1=[]
            for b1i in bnei:
                g1 = b1i.GEN
                if g1 and g1.FLAG==1:
                    gn1.append(g1)
                    try:
                        self.busWithGen[g1.__hnd__].append(b1)
                    except:
                        self.busWithGen[g1.__hnd__]=[b1]
            self.genNeibor.append(gn1) # gen neibor
            if gn1:
                logging.info('\tFound %i GENERATOR '%len(gn1)+b1.toString())
        self.tnei= time.time()-t0
    #
    def runFault(self,var=[]):
        t0 = time.time()
        self.resa = dict()
        #
        self.nf = len(self.lstBus)
        for h1,ba1 in self.busWithGen.items():
            self.nf+=len(ba1)
        #
        logging.info('\nRunning Fault')
        #
        k=0
        for b1 in self.lstBus:
            k+=1
            self.updateLabel(k,self.nf,1,'Run Fault : ')
            #
            self.resa[str(b1.__hnd__)] = runFault(b1)
        #
        for h1,ba1 in self.busWithGen.items():
            g1 = GEN(hnd=h1)
            logging.info('\t%i buses with '%len(ba1)+g1.toString())
            vo = g1.REFV
            g1.REFV = 0
            g1.postData()
            s1 = str(h1)
            for b1 in ba1:
                k+=1
                self.updateLabel(k,self.nf,1,'Run Fault : ')
                #
                self.resa[s1+'_'+str(b1.__hnd__)] = runFault(b1)
            g1.REFV = vo
            g1.postData()
        self.trun = time.time()-t0
    #
    def saveResult(self):
        ## write result to CSV file
        ares = [[PARSER_INPUTS.usage+' V%s'%__version__[:3]]]
        ares.append(['OlxAPI version',self.ver[0]+' V'+self.ver[1]])
        ares.append(['Date',time.asctime()])
        ares.append(['Breaker list file',self.fb])
        ares.append(['Pcnt duty threshold',toString(self.pcnt)+'%'])
        ares.append(['OneLiner file',self.fi])
        ares.append(['Cluster size (tiers)',self.tier])
        ares.append([])
        ares.append(['','','__3LG, (amps)__','','','','__1LG, phase A (amps)__'])
        ares.append(['Breaker bus','Generator','IA_real','IA_imag','IA_mag','','IA_real','IA_imag','IA_mag'])
        #
        for i in range(len(self.lstBus)):
            b1 = self.lstBus[i]
            h1 = str(b1.__hnd__)
            i3a = self.resa[h1][0]
            i1a = self.resa[h1][1]
            ares.append( [b1.NAME+' '+toString(b1.KV)+'kV','',i3a.real,i3a.imag,abs(i3a),'',i1a.real,i1a.imag,abs(i1a)] )
            #
            i3c,i1c = i3a,i1a
            for g1 in self.genNeibor[i]:
                hgi = str(g1.__hnd__)+'_'+h1
                i3 = self.resa[hgi][0]
                i1 = self.resa[hgi][1]
                i3r = i3a-i3
                i1r = i1a-i1
                i3c-= i3r
                i1c-= i1r
                #
                ares.append( ['',g1.toString(),i3r.real,i3r.imag,abs(i3r),'',i1r.real,i1r.imag,abs(i1r)] )
            #
            ares.append( ['','Others',toString(i3c.real),toString(i3c.imag),toString(abs(i3c)),'',toString(i1c.real),toString(i1c.imag),toString(abs(i1c))] )
        #
        if self.busNotFound:
            ares.append(['Bus not found:'])
            for bs1 in self.busNotFound:
                ares.append(['',bs1])
        #
        AppUtils.save2CSV(self.fo,ares,',')
        logging.info('\nFile saved as  : '+self.fo)
        logging.info('Find neibor[s] : %.1f'%(self.tnei))
        logging.info('Run Fault  [s] : %.1f'%(self.trun))
        logging.info("All        [s] : %.1f"%(time.time()-self.t0))
        #
        if self.busNotFound:
            logging.info('\nBus not found:')
            for bs1 in self.busNotFound:
                logging.info('\t'+bs1)
        #
        logging.shutdown()
#
def run0GUI(): # with fb
    try:
        gs = GENSHARE()
        gs.setOlxPath(ARGVS.olxpath)
        gs.setInput(ARGVS.fb,ARGVS.fi,ARGVS.pcnt,ARGVS.tier,ARGVS.fo)
        # neibor
        gs.findNeibor()
        # run Fault
        gs.runFault()
        # saving result
        gs.saveResult()
    except Exception as er:
        logging.error('\nException:\n'+str(er))
    logging.shutdown()

# "C:\Program Files (x86)\ASPEN\Python38-32\python.exe" GenShare.pyw
if __name__ == '__main__':
##    ARGVS.fb = 'LS1_breakerRating.CSV'
##    ARGVS.fi = 'LS1.OLR'
##    #
##    ARGVS.fb = 'C:\\Users\\phuon\\Desktop\\bug\\BreakerList_PowerCoRelay_3.CSV'
##    ARGVS.fi = 'C:\\Users\\phuon\\Desktop\\bug\\PowerCoRelay_3x.olr'
##    #
##    ARGVS.pcnt = 60
##    ARGVS.tier = 5
    # run cmd
    if ARGVS.fb:
        run0GUI()
    else: # GUI
        root = tk.Tk()
        feedback = MainGUI(root)
        root.mainloop()

