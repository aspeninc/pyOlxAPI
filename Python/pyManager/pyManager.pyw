"""
OneLiner GUI tool for launching Python OlxAPI apps through the
OneLiner menu command Tools | OlxAPI App Laucher.


Note: Full file path of this Python program must be listed in OneLiner App manager
      setting in the Tools | User-defined command | Setup dialog box
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "no"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.3"

# IMPORT -----------------------------------------------------------------------
import logging
import sys,os
#
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import time
import pathlib
import importlib.util
from AppUtils import *
#
chekPythonVersion(PY_FILE)
#
import tkinter as tk
import tkinter.filedialog as tkf
from tkinter import ttk
import re

#
# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = iniInput(usage = "GUI for show and run python in a directory")
#
PARSER_INPUTS.add_argument('-fi' ,  help = '*(str) OLR file path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-tpath',help = ' (str) Output folder for PowerScript command file', default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-pk' ,  help = ' (str) Selected object in the 1-Liner diagram', default = [],nargs='+',metavar='')
ARGVS = parseInput(PARSER_INPUTS)
#
ARGSR,NARGS,TARGS,VERBOSE = None,[],{},1
#
def chekPathIsNotWM(path):
    if not os.path.isdir(path):
        return False
    #
    pathSysStart = ["C:\\Intel","C:\\MSOCache","C:\\PerfLogs","C:\\Recovery","C:\\System Volume Information","C:\\Windows","C:\\SYSTEM","C:\\ProgramData","C:\Documents and Settings",\
                "C:\\Program Files\\Microsoft","C:\\Program Files\\Windows","C:\\Program Files\\Intel","C:\\Program Files\\CodeMeter",\
                "C:\\Program Files (x86)\\Microsoft","C:\\Program Files (x86)\\Windows","C:\\Program Files (x86)\\Intel","C:\\Program Files (x86)\\CodeMeter",\
                "C:\\Users\\All Users\\Microsoft\\","C:\\Users\\All Users\\Package Cache\\","C:\\Users\\Public\\Roaming\\","C:\\Users\\All Users\\Windows"]
    for p1 in pathSysStart:
        if path.startswith(p1):
            return False
    #
    pathSysIn = ["\\$RECYCLE.BIN"]
    for p1 in pathSysIn:
        if p1 in path:
            return False
    #
    return True
#
def isPyFile(f1):
    if os.path.isfile(f1) and (f1.upper().endswith(".PY") or f1.upper().endswith(".PYW")):
        return True
    return False
#
def checkPyManagerVisu(f1):
    try:
        a1 = read_File_text(f1)
    except:
        return False
    for i in range(len(a1)):
        if i>100:
            return False
        if a1[i].startswith("__pyManager__"):
            try:
                yn = a1[i].split("=")[1]
                yn = yn.replace(" ","")
                #
                try:
                    yn = float(yn)
                    if yn>0:
                        return True
                except:
                    yn1 = (yn[1:len(yn)-1]).upper()
                    if yn1 =="YES" or yn1 == 'TRUE' or yn=='True':
                        return True
            except:
                pass
    return False
#
def getPyFileShow(alpy):
    res = []
    for p1 in alpy:
        if checkPyManagerVisu(p1):
            res.append(p1)
    return res
#
def pathHavePy(path):
    # not check .py in sub-folder
    if not chekPathIsNotWM(path):
        return False
    #
    try:
        for p in os.listdir(path):
            p1 = os.path.join(path,p)
            if isPyFile(p1) or chekPathIsNotWM(p1):
                return True
    except:
        pass
    return False
#
def get_category(alpy):
    res = []
    for p1 in alpy:
        a1 = read_File_text(p1)
        c1 = 'Common'
        for i in range(len(a1)):
            if i>100:
                break
            s1 = str(a1[i]).replace(" ","")
            if s1.startswith('__category__='):
                as1 = s1.split('=')
                c1 = as1[1]
                c1 = str(c1.split('#')[0])
                try:
                    c1 = c1[1:len(c1)-1]
                except:
                    pass
                break
        res.append(c1)
    return res

def corectOut(out):
    res = out
    idx1 = out.find('\nusage:')
    if idx1>=0:
        res = 'Description:' +out[idx1+6:]
    #
    res = res.replace('optional arguments:','Arguments:    *(Required)')
    res = res.replace('  -h, --help ','')
    res = res.replace(' show this help message and exit','')
    res = res.replace('-ut     ','')
    res = res.replace('(int) unit test [0-ignore, 1-unit test]','')
    res = deleteLineBlank(res)
    res = res.replace('Arguments:    *(Required)','\nArguments:    *(Required)')
    return res
#
def getInputRequired(out):
    res = []
    an = out.split('\n')
    test = False
    for a1 in an:
        if test:
            a2 = a1.split()
            try:
                if a2[1].startswith('*'):
                    res.append(a2[0])
                elif a2[1]=='[' and a2[2]=='...]' and a2[3].startswith('*'):
                    res.append(a2[0])
            except:
                pass
        #
        if a1.startswith('Arguments     * Required,-h --help:'):
            test = True
    return res
#
def getInputs(out):
    res = []
    an = out.split('\n')
    test = False
    for a1 in an:
        if len(a1)>6 and a1.strip().startswith('-'):
            a2 = a1.split()
            res.append(a2[0][1:])
        #
        if a1.startswith('Arguments     * Required,-h --help:'):
            test = True
    return res
#
def getAllPyFile(path,rer):
    res = []
    if not chekPathIsNotWM(path):
        return res
    try:
        for p in os.listdir(path):
            p1 = os.path.join(path,p)
            if isPyFile(p1):
                res.append(p1)
            elif rer and os.path.isdir(p1):
                for p2 in os.listdir(p1):
                    p3 = os.path.join(p1,p2)
                    if isPyFile(p3):
                        res.append(p3)
                    elif rer and os.path.isdir(p3):
                        r1 = getAllPyFile(p3,False)
                        res.extend(r1)
    except:
        pass
    return res
#
class MainGUI(tk.Frame):
    def __init__(self, master):
##        master.attributes('-topmost', True)
        self.splitter = tk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.master = master
        sw = master.winfo_screenwidth()
        sh = master.winfo_screenheight()
        self.master.resizable(0,0)# fixed size
        w = min(1000,sw)
        h = min(650,sh)
        master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2),int(sh/2-h/2)))
        master.wm_title("Python OlxAPI Apps")
##        pathico = os.path.join(os.path.dirname(sys.executable) ,"DLLs")
##        setIco(master,pathico,"pyc.ico")
        setIco_1Liner(master)
        remove_button(self.master)
        self.currentPy = ""
        self.nodeFavorite = []
        self.nodeRecent = []
        self.nodes = dict()
        self.pathFile = dict()
        #
        self.reg1 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\recents"  ,keyUser="",nmax =6)
        self.reg2 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\dir"      ,keyUser="",nmax =1)
        self.reg3 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\favorites",keyUser="",nmax =20)
        self.initGUI()
    #
    def initGUI(self):
        # left-side
        frame_left = tk.Frame(self.splitter)
        self.tree = ttk.Treeview(frame_left, show='tree')
        ysb1 = ttk.Scrollbar(frame_left, orient='vertical'  , command=self.tree.yview)
        xsb1 = ttk.Scrollbar(frame_left, orient='horizontal', command=self.tree.xview)
        # left-side widget layout
        self.tree.grid(row=0, column=0,padx=0,pady=0, sticky='NSEW')
        ysb1.grid(row=0, column=1, sticky='ns')
        xsb1.grid(row=1, column=0, sticky='ew')
        # setup
        self.tree.configure(yscrollcommand=lambda f, l:self.autoscroll(ysb1,f,l), xscrollcommand=lambda f, l:self.autoscroll(xsb1,f,l))
        self.tree.configure(yscrollcommand=ysb1.set, xscrollcommand=xsb1.set)
        self.tree.column("#0",minwidth=300, stretch=True)
        #
        frame_l1 = tk.Frame(frame_left)
        frame_l1.grid(row=2, column=0,padx=0, pady=14)
        #
        self.bt_dir = tk.Button(frame_l1, text="Change directory",width = 27,command=self.open_dir)
        frame_l11 = tk.Frame(frame_l1)

        self.bt_add = tk.Button(frame_l11, text="Add favorite",width = 12,command=self.addFavorite)
        self.bt_rmv = tk.Button(frame_l11, text="Remove favorite",width = 12,command=self.removeFavorite)
        self.bt_dir.grid(row=0, column=1,padx=5,pady=5)
        frame_l11.grid(row=1, column=1)
        self.bt_add.grid(row=1, column=1,padx=5,pady=5)
        self.bt_rmv.grid(row=1, column=3,padx=5,pady=5)

        #--------------------------------------------------------------------------RIGHT
        frame_right = tk.Frame(self.splitter)
        frame_r1 = tk.Frame(frame_right)
        frame_r1.grid(row=0, column=0,sticky='',pady=5,padx=5)#,
        self.text1 = tk.Text(frame_r1,wrap = tk.NONE,width=500,height=22)#
        # yScroll
        ysb2 = ttk.Scrollbar(frame_r1, orient='vertical'  , command=self.text1.yview)
        xsb2 = ttk.Scrollbar(frame_r1, orient='horizontal', command=self.text1.xview)
        ysb2.grid(row=1, column=1, sticky='ns')
        xsb2.grid(row=2, column=0, sticky='ew')
        self.text1.configure(yscrollcommand=lambda f, l:self.autoscroll(ysb2,f,l), xscrollcommand=lambda f, l:self.autoscroll(xsb2,f,l))
        self.text1.configure(yscrollcommand=ysb2.set, xscrollcommand=xsb2.set)
        self.text1.grid(row=1, column=0, sticky='ns')
        frame_r1.columnconfigure(0, weight=1)
        frame_r1.rowconfigure(0, weight=1)
        frame_r1.pack(fill=tk.BOTH, expand=True)

        # ----------------------------------------------------------------------------
        frame_r2 = tk.Frame(frame_right)
        #
        arg = tk.Label(frame_r2, text="Arguments")
        arg.grid(row=0, column=0, sticky='nw')

        self.text2 = tk.Text(frame_r2,wrap = tk.NONE,width=500,height=8)
        # yScroll
        ysb3 = ttk.Scrollbar(frame_r2, orient='vertical'  , command=self.text2.yview)
        xsb3 = ttk.Scrollbar(frame_r2, orient='horizontal', command=self.text2.xview)
        ysb3.grid(row=1, column=1, sticky='ns')
        xsb3.grid(row=2, column=0, sticky='ew')
        self.text2.configure(yscrollcommand=lambda f, l:self.autoscroll(ysb3,f,l), xscrollcommand=lambda f, l:self.autoscroll(xsb3,f,l))
        self.text2.configure(yscrollcommand=ysb3.set, xscrollcommand=xsb3.set)
        self.text2.grid(row=1, column=0, sticky='ns')
        frame_r2.columnconfigure(0, weight=1)
        frame_r2.rowconfigure(1, weight=1)
        frame_r2.pack(fill=tk.BOTH, expand=True)

        # ----------------------------------------------------------------------------
        frame_r3 = tk.Frame(frame_right)
        frame_r3.columnconfigure(1, weight=1)
        frame_r3.rowconfigure(1, weight=1)
        frame_r3.pack(fill=tk.BOTH,expand=True)
        # button launch
        self.bt_launch = tk.Button(frame_r3, text="Launch",width =12,command=self.launch)
        self.bt_launch.grid(row=0, column=1,padx = 0,pady=0)
        #
        frame_r4 = tk.Frame(frame_right)
        frame_r4.columnconfigure(1, weight=1)
        frame_r4.rowconfigure(5, weight=1)
        #
        #button edit
        self.bt_edit = tk.Button(frame_r4 , text = 'Edit in IDE',width =12,command=self.editPyIDE)
        self.bt_edit.grid(row=1, column=0,padx = 20,pady=35)

        #button exit
        bt_exit = tk.Button(frame_r4 , text = 'Exit',width =12,command=self.exit)
        bt_exit.grid(row=1, column=3,padx =0,pady=30)
        #
        self.bt_help = tk.Button(frame_r4 , text = 'Help',width =12,command=self.getHelp)
        self.bt_help.grid(row=1, column=5,padx = 30,pady=30)
        frame_r4.pack(fill=tk.BOTH,expand=True)
        #
        frame_left.columnconfigure(0, weight=1,minsize=210)
        frame_left.rowconfigure(0, weight=1,minsize=210)
        #
        frame_right.columnconfigure(0, weight=1)
        frame_right.rowconfigure(1, weight=1)
        frame_right.pack(fill=tk.BOTH, expand=True)

        #
        # overall layout
        self.splitter.add(frame_left)
        self.splitter.add(frame_right)
        self.splitter.pack(fill=tk.BOTH, expand=True)
        #
        self.dirPy = self.reg2.getValue0()
        #
        self.dirPy = os.path.abspath(self.dirPy)
        self.reg2.appendValue(self.dirPy)
        #
        self.builderTree()
        #
        self.setSTButton('disabled')
    #
    def flush(self):
        pass
    #
    def write(self, txt):
        self.textc.insert(tk.INSERT,txt)
    #
    def clearConsol(self):
        self.textc.delete(1.0,tk.END)
    #
    def getHelp(self):
        gui_info("INFO","@To Add")
    #
    def resetTree(self):
        x = self.tree.get_children()
        for item in x: ## Changing all children from root item
            self.tree.delete(item)
    #
    def builderTree(self):
        self.insert_nodeFavorite('')
        self.insert_nodeRecent('')
        self.insert_node1('', self.dirPy, self.dirPy)
        self.tree.bind('<<TreeviewSelect>>', self.open_node) # simple click
    #
    def autoscroll(self, sbar, first, last):
        """Hide and show scrollbar as needed."""
        first, last = float(first), float(last)
        if first <= 0 and last >= 1:
            sbar.grid_remove()
        else:
            sbar.grid()
        sbar.set(first, last)
    #
    def insert_node1(self, parent, text, abspath):
        alpy = getAllPyFile(abspath,True)
        alpy = getPyFileShow(alpy)
        if len(alpy)==0:
            node = self.tree.insert(parent, 'end', text=text, open=True)
            self.pathFile[node] = abspath
            node1 = self.tree.insert(node, 'end', text="None", open=True)
            self.pathFile[node1] = abspath
            return
        categ = get_category(alpy)
        categF = list(set(categ))
        categF.sort(reverse=True)
        #
        node = self.tree.insert(parent, 'end', text=text, open=True)
        self.pathFile[node] = abspath
        for c1 in categF:
            node1 = self.tree.insert(node, 'end', text=c1, open=True)
            self.pathFile[node1] = abspath
            for i in range(len(alpy)):
                ci = categ[i]
                p1 = alpy[i]
                if ci==c1:
                    path1,py1 = os.path.split(os.path.abspath(p1))
                    nodei = self.tree.insert(node1, 'end', text=py1, open=False)
                    self.pathFile[nodei] = p1
    #
    def insert_node(self, parent, text, abspath):
        hp1 = pathHavePy(abspath)
        hp2 = isPyFile(abspath) and checkPyManagerVisu(abspath)
        if hp1 or hp2 :
            node = self.tree.insert(parent, 'end', text=text, open=False)
            self.pathFile[node] = abspath
        #
        if hp1:
            self.nodes[node] = abspath
            self.tree.insert(node, 'end',open=False)
    #
    def insert_node0(self, parent, text, abspath):
        node = self.tree.insert(parent, 'end', text=text, open=True)
        self.pathFile[node] = abspath
        #
        if pathHavePy(abspath):
            for p in os.listdir(abspath):
                p1 = os.path.join(abspath, p)
                if isPyFile(p1) and checkPyManagerVisu(p1):
                    node1 = self.tree.insert(node, 'end', text=p, open=False)
                    self.pathFile[node1] = p1
                elif pathHavePy(p1):
                    self.insert_node(node, p, p1)
    #
    def insert_nodeRecent(self, node0):
        if node0=='':
            node = self.tree.insert('', 'end', text="Recent", open=True)
            self.pathFile[node] = "Recent"
        else:
            node = node0
        #
        recent = self.reg1.getAllValue()
        self.nodeRecent = [node]
        #
        for r1 in recent:
            if isPyFile(r1) and checkPyManagerVisu(r1):
                r1 = str(pathlib.Path(r1).resolve())
                path1,py1 = os.path.split(os.path.abspath(r1))
                node1 = self.tree.insert(node, 'end', text=py1, open=False)
                self.pathFile[node1] = r1
                self.nodeRecent.append(node1)
    #
    def insert_nodeFavorite(self, node0):
        if node0=='':
            node = self.tree.insert('', 'end', text="Favorite", open=True)
            self.pathFile[node] = "Favorite"
        else:
            node = node0
        #
        fav = self.reg3.getAllValue()
        self.nodeFavorite = [node]
        #
        for f1 in fav:
            if isPyFile(f1) and checkPyManagerVisu(f1):
                f1 = str(pathlib.Path(f1).resolve())
                path1,py1 = os.path.split(os.path.abspath(f1))
                node1 = self.tree.insert(node, 'end', text=py1, open=False)
                self.pathFile[node1] = f1
                self.nodeFavorite.append(node1)
    #
    def open_node(self, event):
        node = self.tree.focus()
        abspath = self.nodes.pop(node, None)
        if abspath:#path
            self.setSTButton('disabled')
            self.tree.delete(self.tree.get_children(node))
            if pathHavePy(abspath):
                for p in os.listdir(abspath):
                    p1 = os.path.join(abspath, p)
                    self.insert_node(node, p, p1)
        else:# File
            v1 = self.pathFile[node]
            if os.path.isfile(v1):
                self.currentPy = os.path.abspath(v1)
                self.text1.configure(state='normal')
                self.setSTButton('active')
                self.showPy()
                #
                if not node in self.nodeFavorite:
                    self.bt_rmv['state']='disabled'
                #
            else:
                self.setSTButton('disabled')
    #
    def setSTButton(self,stt):
        self.bt_add['state']=stt
        self.bt_rmv['state']=stt
        self.bt_launch['state']=stt
        self.bt_edit['state']=stt
    #
    def exit1(self):
        self.master.destroy()
    #
    def exit(self):
        global ARGSR
        ARGSR = None
        self.master.destroy()
    #
    def open_dir(self):
        """Open a directory."""
        self.dirPy = tkf.askdirectory()
        if self.dirPy!="":
            self.dirPy = os.path.abspath(self.dirPy)
        #
            self.resetTree()
            self.builderTree()
            self.reg2.appendValue(self.dirPy)
    #
    def addFavorite(self):
        if self.reg3.appendValue(self.currentPy):
            for i in range(1,len(self.nodeFavorite)):
                self.tree.delete(self.nodeFavorite[i])
            #
            self.insert_nodeFavorite(self.nodeFavorite[0])
    #
    def updateRecents(self):
        if self.reg1.appendValue(self.currentPy):
            for i in range(1,len(self.nodeRecent)):
                self.tree.delete(self.nodeRecent[i])
            #
            self.insert_nodeRecent(self.nodeRecent[0])
    #
    def removeFavorite(self):
        self.reg3.deleteValue(self.currentPy)
        for i in range(1,len(self.nodeFavorite)):
            self.tree.delete(self.nodeFavorite[i])
            #
        self.insert_nodeFavorite(self.nodeFavorite[0])
    #
    def getVersion(self):
        for i in range(len(self.currentPy_as)):
            if self.currentPy_as[i].startswith("__version__"):
                return self.currentPy_as[i]
            if i>100:
                break
        return "__version__ = Unknown"
    #
    def showPy(self):
        #
        global NARGS,TARGS,IREQUIRED,VERBOSE
        IREQUIRED = []
        try:
            filehandle = open(self.currentPy, 'a' )
            filehandle.close()
            self.bt_edit['state']='active'
        except IOError:
            self.bt_edit['state']='disabled'
        #
        self.currentPy_as,se = read_File_text_0(self.currentPy)
        self.text1.delete('1.0', tk.END)
        self.text2.delete('1.0', tk.END)
        self.text1.insert(tk.END, self.currentPy.replace('\\','/')+"\n")
        self.text1.insert(tk.END, self.getVersion()+"\n")
        #
        o1 = None
        if haveParseInput(self.currentPy_as):
            out,err,returncode = runSubprocess_getHelp(self.currentPy)
            if returncode!=0:
                self.text1.insert(tk.END, err.strip())
                return
            o1 = out.replace(" ","")
        #
        if o1==None:
            self.text1.insert(tk.END, se)
            return
        #
        IREQUIRED = getInputRequired (out)
        IInternal = getInputs(out)
        #
        # out = corectOut(out)
        self.text1.insert(tk.END, out)
        self.text1.configure(state='disabled')
        #
        try:
            path = os.path.dirname(self.currentPy)
            if path not in sys.path:
                os.environ['PATH'] = path + ";" + os.environ['PATH']
                sys.path.insert(0, path)
            #
            spec = importlib.util.spec_from_file_location('', self.currentPy)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            vala = module.ARGVS.__dict__
            valb = dict(sorted(vala.items(), key=lambda item: item[0]))
            #
            NARGS = ['-h','-verbose']
            TARGS = {'-h':str,'-verbose':int}
            for key, val in valb.items():
                s1 = '-'+key
                NARGS.append(s1)
                TARGS[s1]= type(val)
                if key in IInternal:
                    s1 +=' '
                    #
                    try:
                        val0 = ARGVS.__dict__[key]
                    except:
                        val0 = val
                    if type(val)==list:
                        for v1 in val0:
                            s1 += "\"" + v1 + "\"" + "  "
                    else:
                        if type(val)==str :
                            s1 += "\"" + str(val0) + "\""
                        else:
                            s1 += str(val0)
                    #
                    s1 +="\n"
                    s1 = s1.replace('\\','/')
                    self.text2.insert(tk.END,s1)
        except:
            pass
    #
    def editPyIDE(self):
        app = os.path.dirname(sys.executable)+"\\Lib\\idlelib\\idle.bat"
        args = [app,self.currentPy]
        runSubprocess_noWait(args) #run subprocess without waiting
    #
    def errorInputArgument(self,v1,se):
        global ARGSR
        ARGSR = None
        sMain= se.ljust(30) +'\n'+ v1
        gui_info(sTitle="ERROR Input Argument",sMain=sMain)
        logging.error('\n'+sMain)
    #
    def launch(self):
        global ARGSR,IREQUIRED,VERBOSE
        #
        logging.info('run : '+self.currentPy)
        #
        pathe = os.path.dirname(sys.executable)
        if (self.currentPy.upper()).endswith('.PYW'):
            ARGSR = [os.path.join(pathe,'pythonw.exe')]
        else:
            ARGSR = [os.path.join(pathe,'python.exe')]
        #
        ARGSR.append(self.currentPy)
        #
        a1 = (self.text2.get(1.0, tk.END)).split("\n")
        #
        errInputRequired = []
        demo1 = 0
        #
        for i in range(len(a1)):
            v1 = (a1[i]).strip()
            #
            if v1!='':
                k1 = str(v1).find(" ")
                if k1>0:
                    u1 = v1[:k1]
                    u2 = v1[k1+1:]
                else:
                    u1 = v1
                    u2 =''
                    #
                if u1 not in NARGS:
                    self.errorInputArgument(v1,'Input not found:')
                    return
                #
                t1 = TARGS[u1]
                u2 = u2.strip()
                u2 = u2.strip('"')
                u2 = u2.strip("'")
                #
                if u1=='-demo' and u2:
                    if u2 !='0':
                        demo1 = 1
                #
                if u1 in IREQUIRED and u2=='':
                    errInputRequired.append(v1)
                #
                if u2:
                    #
                    try:
                        if checkSpecialCharacter(u2):
                            self.errorInputArgument(v1,'Special Character found:')
                            return
                    except:
                        pass
                    #
                    try:
                        val = t1(u2)
                    except:
                        try:
                            if t1 in {int,float}:
                                u2a = re.split('\s|;|,|"',u2)
                            else:
                                u2a = re.split(';|,|"',u2)
                            for ui in u2a:
                                val = t1(ui)
                        except:
                            self.errorInputArgument(v1,'Type (' +str(t1.__name__) + ') not found:')
                            return
##                    if t1!=list:
##                        k2 = u2.find('"')
##                        if k2>=0:
##                            self.errorInputArgument(v1,'Error input:')
##                            return
##                        #
##                        if u2:
##                            ARGSR.append(u1)
##                            ARGSR.append(u2)
##                    else:
                    if u2:
                        if u1=='-verbose':
                            try:
                                VERBOSE = int(u2)
                            except:
                                pass
                        else:
                            ARGSR.append(u1)
                            #
                            if t1 in {int,float}:
                                ua = re.split('\s|;|,|"',u2)
                            else:
                                ua = re.split(';|,|"',u2)
                            for ui in ua:
                                ui2 = ui.strip()
                                if ui2:
                                    ARGSR.append(ui2)
        #
        if demo1==0 and errInputRequired:
            sv1 = ''
            for e1 in errInputRequired:
                sv1 +=e1+'\n'
            #
            self.errorInputArgument(sv1,'* Required input missing')
            return
        #
        self.updateRecents()
        self.exit1()
    #
    def run(self):
        self.frame_r4.grid_forget()
        self.pgFrame.grid(row=0,column=0, padx=0, pady=0)
        #
        a1,r1 = self.bt_add['state'],self.bt_rmv['state']
        self.stop_b['state']='active'
        self.bt_add['state']='disabled'
        self.bt_rmv['state']='disabled'
        self.bt_dir['state']='disabled'
        #
        self.var1.set("Running")
        self.progress['value'] = 0
        self.text2.configure(state='disabled')
        #
        self.progress.start()
        self.t = TraceThread(target=self.runPy)
        self.t.start()

        #BUTTON STATUS
        self.text2.configure(state='normal')
        self.bt_dir['state']='active'
        self.bt_add['state']= a1
        self.bt_rmv['state']= r1

    def finish(self):
        self.progress.stop()
        self.pgFrame.grid_forget()
        self.frame_r4.grid(row=0, column=0,sticky='',pady=5)
    #
    def stop_progressbar(self):
        self.textc.insert(tk.INSERT,"Stop by user")
        self.finish()
        self.t.killed = True
#
def checkSpecialCharacter(s):
    s1 = s.replace("\\",'')
    b= str(s1.encode('utf8'))
    b = b.replace("\\",'')
    b = b[2:len(b)-1]
    return b!=s1
#
def haveParseInput(ar1):
    for i in range(len(ar1)):
        a01 = ar1[i].replace(' ','')
        if a01.startswith('ARGVS=AppUtils.parseInput(PARSER_INPUTS') or\
           a01.startswith('ARGVS=parseInput(PARSER_INPUTS') or a01.startswith('PARSER_INPUTS.add_argument('):
            return True
        if i>200:
            return False
    return False
#
def createRunFile(tpath,args,verb):
    s  = "import sys,os\n"
    s += "PATH_FILE = os.path.split(os.path.abspath(__file__))[0]\n"
    if os.path.isfile(os.path.join(PATH_FILE,'AppUtils.py')):
        plb = PATH_FILE
    else:
        plb = PATH_LIB
    plb = plb.replace('\\','/')
    s += "PATH_LIB = "+ '"'+plb+'"\n'
    #
    #s += 'os.environ["PATH"] = PATH_LIB + ";" + os.environ["PATH"]\n'
    s += 'sys.path.insert(0, PATH_LIB)\n'
    s += 'import AppUtils\n'
    #
    s += "command = ["
    for a1 in args:
        a1 = a1.replace('\\','/')
        s+='"'+a1 + '"' +","
    s = s[:len(s)-1]
    s+=']\n'
    s+= 'print("Run : "+command[1])\n'
    s+= 'print()\n'
    s+= '#\n'
    s+='AppUtils.runCommand(command,PATH_FILE,%i)\n'%verb
    #
    sfile = os.path.join(tpath,'run.py')
    saveString2File(sfile,s)
    return sfile
#
def runFinal():
    global ARGSR
    if ARGSR!=None:
        s1 = "\ncmd:\n"
        for a1 in ARGSR:
            s1 += '"'+a1 + '" '
        #
        logging.info(s1)
        #
        sfile = createRunFile(ARGVS.tpath,ARGSR,VERBOSE)
        argn = [ARGSR[0],sfile]
        #
        runSubprocess_noWait(argn)
    else:
        #cancel
        logging.info('exit by user')
    #
def main():
    if ARGVS.tpath=='':
        ARGVS.tpath = get_tpath()
    #
    FRUNNING,SRUNNING = openFile(ARGVS.tpath,'running')
    #-----------------------------------
    root = tk.Tk()
    app = MainGUI(root)
    root.mainloop()
    #
    runFinal()
    #-----------------------------------
    closeFile(FRUNNING)
    deleteFile(SRUNNING)
    #
    if not isFile(ARGVS.tpath,'success'):
        createFile(ARGVS.tpath,'cancel')

# "C:\\Program Files (x86)\\ASPEN\\Python38-32\\python.exe" "pyManager.pyw"
# "C:\\Program Files (x86)\\ASPEN\\Python38-32\\python.exe" pyManager.py -h
if __name__ == '__main__':
    logger2File(PY_FILE)
    main()
    logging.shutdown()









