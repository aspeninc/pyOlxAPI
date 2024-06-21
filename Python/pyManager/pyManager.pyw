"""
OneLiner GUI tool for launching Python OlxAPI apps through the
OneLiner menu command Tools | OlxAPI App Laucher.


Note: Full file path of this Python program must be listed in OneLiner App manager
      setting in the Tools | User-defined command | Setup dialog box
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "no"
__email__     = "support@aspeninc.com"
__status__    = "In Development"
__version__   = "2.2.2"

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
import pathlib
import importlib.util
from AppUtils import *
#
chekPythonVersion(PY_FILE)
import ctypes
#
import tkinter as tk
import tkinter.filedialog as tkf
from tkinter import ttk
##import tkinter.font
import idlelib.colorizer as ic
import idlelib.percolator as ip

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # if your windows version >= 8.1
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware() # win 8.0 or less
    except:
        pass

# INPUTS cmdline ---------------------------------------------------------------
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage = "GUI for show and run python in a directory"
PARSER_INPUTS.add_argument('-python' ,  help = '*(str) python path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-opath' ,  help = '*(str) out file path' , default = '',type=str,metavar='')
PARSER_INPUTS.add_argument('-szwindow' , help = ' (str) size of window',default = '' ,type=str,metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]
#
ARGSR,NARGS,TARGS,VERBOSE = None,[],{},1
txtfont = ("MS Sans Serif", 8)
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
##    res = res.replace('-ut     ','')
##    res = res.replace('(int) unit test [0-ignore, 1-unit test]','')
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
        if a1.strip()=='-h, --help  show this help message and exit':
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
            if a2[0][1:].strip():
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

def fix_scaling(root):
    """Scale fonts on HiDPI displays."""
    import tkinter.font
    scaling = float(root.tk.call('tk', 'scaling'))
    if scaling > 1.4:
        for name in tkinter.font.names(root):
            font = tkinter.font.Font(root=root, name=name, exists=True)
            size = int(font['size'])
            if size < 0:
                font['size'] = round(-0.75*size)

def fixwordbreaks(root):
    # On Windows, tcl/tk breaks 'words' only on spaces, as in Command Prompt.
    # We want Motif style everywhere. See #21474, msg218992 and followup.
    tk = root.tk
    tk.call('tcl_wordBreakAfter', 'a b', 0) # make sure word.tcl is loaded
    tk.call('set', 'tcl_wordchars', r'\w')
    tk.call('set', 'tcl_nonwordchars', r'\W')


#
class MainGUI(tk.Frame):
    def __init__(self, master):
        #master.attributes('-topmost', True)
        offX,offY,sw,sh = calOffsetMonitors(master)
        #
        self.splitter = tk.PanedWindow(master, orient=tk.HORIZONTAL)
        self.master = master


        if ARGVS.szwindow=='':
            szwindow = '20 60 15 60'
        else:
            szwindow = ARGVS.szwindow
        va = szwindow.split()
        x0 = float(va[0])/100+offX/sw
        x1 = float(va[1])/100
        y0 = float(va[2])/100+offY/sh
        y1 = float(va[3])/100
        self.scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        self.btw = [12,12,25]
        szt1,szt2 = 8,8
##        self.txt_font1 = tkinter.font.Font( family = "MS Sans Serif", size = szt1)
##        self.txt_font = tkinter.font.Font( family = "MS Sans Serif", size = szt2)
##        self.txt_fontb = tkinter.font.Font( family = "MS Sans Serif",weight='bold', size = szt2)
        master.geometry("{0}x{1}+{2}+{3}".format(int(x1*sw),int(y1*sh) + 40,int(x0*sw),int(y0*sh)))

        master.wm_title("Python OlxAPI Apps")
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure((0,1), weight=1, uniform=1)

        self.adapt_size = int(36*(sh)/(1080/y1)/self.scale_factor*100)

        setIco_1Liner(master)#,png='1Liner.png'
        hide_minimize_maximize(self.master)
        self.outFile = ""
        self.currentPy = ""
        self.nodeFavorite = []
        self.nodeRecent = []
        self.nodes = dict()
        self.pathFile = dict()
        #
        self.reg1 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\recents"  ,keyUser="",nmax =10)
        self.reg2 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\dir"      ,keyUser="",nmax =1)
        self.reg3 = WIN_REGISTRY(path = "SOFTWARE\\ASPEN\\OneLiner\\PyManager\\favorites",keyUser="",nmax =20)
        self.reg1.deleteInvalidFile()
        self.reg3.deleteInvalidFile()
        self.initGUI()
        self.setSTButton('active')
    #
    def initGUI(self):
        self.currPy1 = tk.StringVar()
        self.currentPy = self.reg1.getValue0()
        # left-side
        frame_left = tk.Frame(self.splitter)
        self.tree = ttk.Treeview(frame_left, show='tree')
        ysb1 = ttk.Scrollbar(frame_left, orient='vertical'  , command=self.tree.yview)
        # left-side widget layout
        self.tree.grid(row=0, column=0,padx=5,pady=0, sticky='NSEW')
        ysb1.grid(row=0, column=1, sticky='ns')
        # setup
        self.tree.configure(yscrollcommand=ysb1.set)#, xscrollcommand=xsb1.set)
        self.tree.column("#0",minwidth=10, stretch=True)
        self.tree.grid_columnconfigure(0, weight=1)
        self.tree.grid_rowconfigure((0,1), weight=1, uniform=1)
        style = ttk.Style()
        rowheight_adap = 20 +int((self.scale_factor-100.)/100*10)
        style.configure("Treeview", highlightthickness=0, bd=0, font=txtfont, rowheight=rowheight_adap)


        frame_l1 = tk.Frame(frame_left)
        frame_l1.grid(row=2, column=0,padx=0, pady=5)
        #
        self.bt_dir = tk.Button(frame_l1, text="Change directory",width = self.btw[2],command=self.open_dir, font = txtfont)
        frame_l11 = tk.Frame(frame_l1)

        self.bt_add = tk.Button(frame_l11, text="Add favorite",width = self.btw[0],command=self.addFavorite, font = txtfont)
        self.bt_rmv = tk.Button(frame_l11, text="Remove",width = self.btw[0],command=self.removeFavorite, font = txtfont)
        self.bt_dir.grid(row=0, column=1,padx=5,pady=5)
        frame_l11.grid(row=1, column=1)
        self.bt_add.grid(row=1, column=1,padx=1,pady=5)
        self.bt_rmv.grid(row=1, column=2,padx=1,pady=5)


        #--------------------------------------------------------------------------RIGHT
        frame_right = tk.Frame(self.splitter)

         # ----------------------------------------------------------------------------
        frame_r1 = tk.Frame(frame_right)
        #
        arg = tk.Label(frame_r1, text="Path file ", font = txtfont)
        arg.grid(row=0, column=0, columnspan=2, sticky='w', pady=3)

        cupy1 = tk.Entry(frame_r1, width= 70, textvariable=self.currPy1,bd=2,bg='SystemWindow',font = txtfont)
        cupy1.grid(row=1, column=0, sticky='nsew',padx=0, pady=3)

        btselectpy = tk.Button(frame_r1, text=" ... ", width= 2,relief= tk.RAISED,command=self.selectPYFile, font = txtfont)
        btselectpy.grid(row=1, column=1, sticky='nsew', padx=5, pady=3)

        bt_newPy = tk.Button(frame_r1 , text = 'New',width= 4,relief= tk.RAISED, command=self.newPyIDE, font=txtfont)
        bt_newPy.grid(row=1, column=2, sticky='nsew', padx=5, pady=3)

        frame_r1.columnconfigure(0, weight=24)
        frame_r1.columnconfigure(1, weight=2)
        frame_r1.columnconfigure(2, weight=4)
        frame_r1.rowconfigure(1, weight=1)
        frame_r1.pack(fill=tk.BOTH, expand=False, pady=5)


        frame_r2 = tk.Frame(frame_right)
        # frame_r2.grid(row=0, column=0,sticky='',pady=5,padx=0)#,

        self.text1 = tk.Text(frame_r2, wrap = tk.NONE, width=600,height=self.adapt_size)#
        # yScroll
        ysb2 = ttk.Scrollbar(frame_r2, orient='vertical'  , command=self.text1.yview)
        xsb2 = ttk.Scrollbar(frame_r2, orient='horizontal', command=self.text1.xview)
        ysb2.grid(row=0, column=1, sticky='ns')
        xsb2.grid(row=1, column=0, sticky='ew')
        self.text1.configure(yscrollcommand=lambda f, l:self.autoscroll(ysb2,f,l), xscrollcommand=lambda f, l:self.autoscroll(xsb2,f,l))
        self.text1.configure(yscrollcommand=ysb2.set, xscrollcommand=xsb2.set)
        self.text1.grid(row=0, column=0, sticky='ns')
        cdg = ic.ColorDelegator()
        ip.Percolator(self.text1).insertfilter(cdg)
        # Parsed the Font object
        # to the Text widget using .configure( ) method.
        self.text1.configure(font = txtfont)

        frame_r2.columnconfigure(0, weight=1, uniform='1',minsize=30)
        frame_r2.rowconfigure(0, weight=1, uniform='1',minsize=30)

        frame_r2.pack(fill=tk.BOTH, expand=True)

        frame_r3 = tk.Frame(frame_right)

        # arg = tk.Label(frame_r3, text="Args  ", font = self.txt_font)
        # arg.grid(row=0, column=0, sticky='nw', pady=3)

        btlaunch = tk.Button(frame_r3, text="Launch",width= self.btw[0],borderwidth=2,relief= tk.RAISED,command=self.launch,font = txtfont)#background='red',font='sans 10 bold',
        btlaunch.grid(row=0, column=0, sticky='ns', padx=40, pady=5)

        btlaunch_cmd = tk.Button(frame_r3 , text = "Launch with Command line \n Parameters", width = 2* self.btw[0],command=self.showCmdLineWindow, font = txtfont)
        btlaunch_cmd.grid(row=0, column=1, sticky='ns', padx=40, pady=5)

        btedit2 = tk.Button(frame_r3, text="Edit in IDE",width= self.btw[0],relief= tk.RAISED, command=self.editPyIDE, font = txtfont)
        btedit2.grid(row=0, column=2, sticky='ns', padx=40, pady=5)



        # bt_exit = tk.Button(frame_r3 , text = 'Exit',width =self.btw[0],command=self.exit,font = self.txt_font)
        # bt_exit.grid(row=0, column=4, sticky='nw', padx=10, pady=3)


        frame_r3.columnconfigure((0,2), weight=2, uniform='1')
        frame_r3.columnconfigure(1, weight=4, uniform='1')
        frame_r3.rowconfigure(1, weight=1)
        frame_r3.pack(fill=tk.BOTH, expand=False, pady=5)

        self.cmdparwin = tk.Toplevel(self.master)
        self.cmdparwin.title("Arguments")
        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        w = min(600,sw)
        h = min(200,sh)
        self.cmdparwin.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2),int(sh/2-h/2)))
        self.cmdparwin.transient(self.master)
##        setIco_1Liner(self.cmdparwin)
        self.cmdparwin.wm_protocol ("WM_DELETE_WINDOW", self.cmdparwin.withdraw)
        self.cmdparwin.withdraw()


        # self.cmdlineparam.withdraw()
        frame_cmd = tk.Frame(self.cmdparwin)

        self.text2 = tk.Text(frame_cmd, wrap = tk.NONE, width=600)
        # yScroll
        ysb3 = ttk.Scrollbar(frame_cmd, orient='vertical'  , command=self.text2.yview)
        xsb3 = ttk.Scrollbar(frame_cmd, orient='horizontal', command=self.text2.xview)
        ysb3.grid(row=0, column=1, sticky='ns')
        xsb3.grid(row=1, column=0, sticky='ew')
        self.text2.configure(yscrollcommand=lambda f, l:self.autoscroll(ysb3,f,l), xscrollcommand=lambda f, l:self.autoscroll(xsb3,f,l))
        self.text2.configure(yscrollcommand=ysb3.set, xscrollcommand=xsb3.set)
        self.text2.configure(font = txtfont)
        self.text2.grid(row=0, column=0, sticky='ns')

        frame_btt = tk.Frame(self.cmdparwin)
        btrun_cmd = tk.Button(frame_btt , text = "Run", width =self.btw[0],command=self.launchcmd, font = txtfont)
        btrun_cmd.pack(pady = 5)

        frame_btt.pack()
        frame_cmd.pack()

        #
        frame_left.columnconfigure(0, weight=1,minsize=2)#210
        frame_left.rowconfigure(0, weight=1,minsize=2)
        #
        frame_right.columnconfigure(0, weight=1,minsize=600)
        frame_right.rowconfigure(1, weight=1,minsize=600)
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
        try:
            self.showPy()
        except:
            pass
        try:
            self.tree.selection_set([self.nodeRecent[1]])
        except:
            pass
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

    def search(self):
        query = self.currPy1
        selections = []
        for child in self.tree.get_children():
            if query.lower() in self.tree.item(child)['values'].lower():   # compare strings in  lower cases.
                # print(tree.item(child)['values'])
                selections.append(child)
        # print('search completed')
        self.tree.selection_set(selections)

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
            if isPyFile(r1):# and checkPyManagerVisu(r1):
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
            if isPyFile(f1):# and checkPyManagerVisu(f1):
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
            if node:
                v1 = self.pathFile[node]
                if os.path.isfile(v1):
                    self.currentPy = os.path.abspath(v1)
                    self.text1.configure(state='normal')
                    self.setSTButton('active')
                    self.showPy()
                    #
                else:
                    self.setSTButton('disabled')
    #
    def setSTButton(self,stt):
        self.bt_add['state']=stt
        self.bt_rmv['state']=stt
    #
    def exit1(self):
        logging.shutdown()
        self.master.destroy()
        self.master.quit()
    #
    def exit(self):
        global ARGSR
        ARGSR = None
        self.master.destroy()
        self.master.quit()

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
        node = self.tree.focus()
        if node in self.nodeFavorite:
            for ki in range(len(self.nodeFavorite)):
                if node==self.nodeFavorite[ki]:
                    break
            self.reg3.deleteValue(self.currentPy)
            for i in range(1,len(self.nodeFavorite)):
                self.tree.delete(self.nodeFavorite[i])
            self.insert_nodeFavorite(self.nodeFavorite[0])

            try:
                nc = self.nodeFavorite[max(1,ki-1)]
                self.tree.selection_set([nc])
                v1 = self.pathFile[nc]
                self.currentPy = v1
                self.showPy()
            except:
                pass
        #
        elif node in self.nodeRecent:
            for ki in range(len(self.nodeRecent)):
                if node==self.nodeRecent[ki]:
                    break
            #
            self.reg1.deleteValue(self.currentPy)
            for i in range(1,len(self.nodeRecent)):
                self.tree.delete(self.nodeRecent[i])
            self.insert_nodeRecent(self.nodeRecent[0])
            try:
                nc = self.nodeRecent[max(1,ki-1)]
                self.tree.selection_set([nc])
                v1 = self.pathFile[nc]
                self.currentPy = v1
                self.showPy()
            except:
                pass
    #
    def getVersion(self):
        for i in range(len(self.currentPy_as)):
            if self.currentPy_as[i].startswith("__version__"):
                return self.currentPy_as[i]
            if i>100:
                break
        return "__version__ = Unknown"
    #
    def selectPYFile(self):
        v1 = tkf.askopenfilename(filetypes=[("Python Files", "*.py *.pyw")],title='Select Python File')
        if v1!='':
            self.currentPy = v1
            self.updateRecents()
            self.showPy()
            self.tree.selection_set([self.nodeRecent[1]])
    #
    def showPy(self):
        global NARGS,TARGS,IREQUIRED,DFARGS
        IREQUIRED = []
        try:
            filehandle = open(self.currentPy, 'a' )
            filehandle.close()
        except IOError:
            pass
        #
        self.cmdparwin.title("Arguments: " +os.path.split(self.currentPy)[1])
        self.currentPy_as,se = read_File_text_0(self.currentPy)

        self.text1.configure(state='normal')
        self.text1.delete('1.0', tk.END)
        self.text2.delete('1.0', tk.END)

        #self.text1.insert(tk.END, self.currentPy.replace('\\','/')+"\n")
        #self.text1.insert(tk.END, self.getVersion()+"\n")
        self.currPy1.set(self.currentPy.replace('\\','/'))
        #
        o1 = None
        if haveParseInput(self.currentPy_as):
            try:
                if isEmbedded():
                    cwd = os.path.dirname(OlxAPI.__file__)
                else:
                    cwd = PATH_LIB
                #
                fo1 = get_opath('')+'\\'+os.path.basename(self.currentPy)
                fo1 = get_file_out(fo=fo1, fi='' , subf='' , ad='' , ext='.py')
                saveString2File(fo1,"import os,sys;sys.path.insert(0, os.getcwd()); "+ se)
                out,err,returncode = runSubprocess_getHelp(fo1, pythonPath=ARGVS.python,cwd=cwd,timeout=0.5)
                if returncode!=0:
                    self.text1.insert(tk.END, err.strip().replace(fo1,self.currentPy))
                    return
                o1 = out.replace(" ","")
            except:
                o1 = None
        #
        if o1==None:
            self.text1.insert(tk.END, se)
            self.text1.configure(state='disabled')
            self.text1.bind("<Key>", lambda event: "break")
            return
        #
        IREQUIRED = getInputRequired(out)
        IInternal = getInputs(out)
        #
        out = corectOut(out)
        self.text1.insert(tk.END, out)
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
            NARGS = []
            TARGS,DFARGS = {},{}
            for key, val in valb.items():
                s1 = '-'+key
                NARGS.append(s1)
                TARGS[s1]= type(val)
                DFARGS[s1] = val
                if key in IInternal:
                    s1 +=' '
                    #
                    val0 = val
                    if type(val)==list:
                        for v1 in val0:
                            s1 += "\"" + v1 + "\"" + "  "
                    else:
                        if type(val)==str :
                            s1 +=  "\"" + str(val0)+"\""
                        else:
                            s1 += str(val0)
                    s1 += "\n"
                    s1 = s1.replace('\\','/')
                    self.text2.insert(tk.END,s1)
            self.text1.configure(state='disabled')
        except:
            self.text1.insert(tk.END, se)

        self.text1.configure(state='disabled')
        self.text1.bind("<Key>", lambda event: "break")

    #
    def newPyIDE(self):#OK
        sres = 'editIDE\n-fi ' +'Untitled.py'+'\n'
        if self.outFile:
            saveString2File(self.outFile, sres)
        self.exit1()

    #
    def editPyIDE(self):
        self.currentPy = self.getCurrentPy()
        if self.currentPy=='':
            return
        sres = 'editIDE\n-fi '+self.currentPy+'\n'
        if self.outFile:
            saveString2File(self.outFile, sres)
        self.exit1()
    #
    def errorInputArgument(self,v1,se):
        sMain= se.ljust(30) +'\n'+ v1
        gui_info(sTitle="ERROR Input Argument",sMain=sMain)
        logging.info(sMain)
    #
    def getArgs(self):# check +correct args
        res = ''
        for a1 in (self.text2.get(1.0, tk.END)).split("\n"):
            v1 = a1.strip()
            if v1:
                k1 = v1.find(" ")
                if k1>0:
                    u1 = v1[:k1]
                    u2 = v1[k1+1:]
                else:
                    u1 = v1
                    try:
                        u2 = DFARGS[u1]
                    except:
                        u2 = ''
                if u1 not in NARGS:
                    self.errorInputArgument(v1,'Input not found:')
                    return None
                try:
                    if checkSpecialCharacter(u2):
                        self.errorInputArgument(v1,'Special Character found: ')
                        return None
                except:
                    pass

                t1 = TARGS[u1]
                #
                try:
                    u2n = t1(u2)
                    if t1==str:
                        u2n = u2n.strip('"')
                        u2n = u2n.strip("'")
                        u2n = '"'+u2n+'"'
                except:
                    self.errorInputArgument(v1, 'Type (' +str(t1.__name__) + ') not found:')
                    return None
                res += u1+' '+str(u2n) +'\n'
        #
        logging.info(res)
        return res
    #
    def launch(self):
        global ARGSR,IREQUIRED,VERBOSE
        #
        self.currentPy = self.getCurrentPy()
        if self.currentPy=='':
            return
        logging.info('run : '+self.currentPy)
        sres = 'run\n'+self.currentPy+'\n'
        # args = self.getArgs()
        # if args==None:
        #     return
        # sres += args
        if self.outFile:
            saveString2File(self.outFile, sres)
        #
        self.updateRecents()
        self.exit1()
    #
    def launchcmd(self):
        global ARGSR,IREQUIRED,VERBOSE
        #
        logging.info('run : '+self.currentPy)
        sres = 'run\n'+self.currentPy+'\n'
        # Toplevel object which will
        # be treated as a new window
        args = self.getArgs()
        if args==None:
            return
        sres += args
        if self.outFile:
            saveString2File(self.outFile, sres)
        #
        self.updateRecents()
        self.exit1()

    def getCurrentPy(self):
        f1 = self.currPy1.get()
        if not isPyFile(f1):
            gui_error('Error',f1+'\nPython file not found.\nCheck the file name and try again.')
            return ''
        return f1

    def showCmdLineWindow(self):
        #
        self.currentPy = self.getCurrentPy()
        if self.currentPy=='':
            return
        self.cmdparwin.deiconify()

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
        if a01.startswith('importargparse') or a01.startswith('PARSER_INPUTS.add_argument('):
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
    if ARGVS.python=='' and isEmbedded():
        raise( Exception('Python path is missing' ) )

    #-----------------------------------
    root = tk.Tk()
    app = MainGUI(root)
    fix_scaling(root)
    fixwordbreaks(root)
    app.outFile =  ARGVS.opath
    root.mainloop()

# "C:\\Program Files (x86)\\ASPEN\\Python38-32\\python.exe" "pyManager.pyw"
# "C:\\Program Files (x86)\\ASPEN\\Python38-32\\python.exe" pyManager.py -h
if __name__ == '__main__':
    logger2File(PY_FILE)
    main()
    logging.shutdown()









