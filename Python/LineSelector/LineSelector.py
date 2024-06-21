"""
SubSystemSelector: select subsystem by area/zone/kv range or all network

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2020, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__version__   = "0.1.1"
__email__     = "support@aspeninc.com"
__status__    = "In development"

# IMPORT -----------------------------------------------------------------------
import sys,time,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPI
import OlxAPILib
from OlxAPIConst import *

import tkinter as tk
import tkinter.filedialog as tkf
import tkinter.messagebox as tkm
from tkinter import ttk
import xml.etree.ElementTree as ET
OlxAPILib.chekPythonVersion(PY_FILE)
#
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= '')
PARSER_INPUTS.usage = '\nSubsystem selector by Area/Zone/kV range or All network'
PARSER_INPUTS.add_argument('-olxpath' , metavar='', help = ' (str) ASPEN folder (where olxapi.dll is located)', default = '')
PARSER_INPUTS.add_argument('-opath'   , metavar='', help = ' (str) Ouput folder ', default = '')
PARSER_INPUTS.add_argument('-fi' , metavar='', help = '*(str) OLR input file', default = '')
PARSER_INPUTS.add_argument('-fo' , metavar='', help = ' (str) XML ouput file', default = '')
ARGVS,_ = PARSER_INPUTS.parse_known_args()

#
class GuiSelectListBox(tk.Frame):
    def __init__(self,master,title,data,var):
        tk.Frame.__init__(self, master=master)
        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        self.var= var
        self.data = data
        w,h = 220,220
        ow,oh = 100,100
        self.master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2)+ow,int(sh/2-h/2)+oh))
        self.master.resizable(0,0)# fixed size
        self.master.wm_title(title)
        OlxAPILib.setIco(self.master,"","1LINER.ico")
        OlxAPILib.remove_button(self.master)
        #
        button_OK = tk.Button(self.master,text = 'OK',width = 10,command=self.master.destroy)
        button_OK.place(x=80, y=180)
        #
        self.lb = tk.Listbox(self.master,selectmode='extended')
        self.lb.pack(expand=1, fill=tk.BOTH, side=tk.TOP)
        for d1 in data:
            self.lb.insert(tk.END, d1[0])

        self.lb.place(x=50, y=10)
        self.lb.bind('<<ListboxSelect>>', self.cb)
    #
    def cb(self,event):
        va = list(self.lb.curselection())
        sr = ""
        for v1 in va:
            sr += str(self.data[v1][1]) + ","
        sr = sr[:len(sr)-1]
        self.var.set(sr)

#
class MainGUI(tk.Frame):
    def __init__(self,master):
        tk.Frame.__init__(self, master=master)
        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        w,h = 400,200
        self.master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2),int(sh/2-h/2)))
        self.master.resizable(0,0)# fixed size
        self.master.wm_title("Line Selector")
        OlxAPILib.setIco(self.master,"","1LINER.ico")
        OlxAPILib.remove_button(self.master)
        #
        self.var1 = tk.IntVar()
        self.r1=tk.Radiobutton(self.master, text="All lines in network",variable = self.var1,value=1,command=self.radioAll)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.master, text="Lines in area(s)"    ,variable = self.var1,value=2,command=self.radioArea)
        self.r2.pack(anchor=tk.W)
        self.r3=tk.Radiobutton(self.master, text="Lines in zone(s)"    ,variable = self.var1,value=3,command=self.radioZone)
        self.r3.pack(anchor=tk.W)

        self.r1.place(x=20,y=10)
        self.r2.place(x=20,y=40)
        self.r3.place(x=20,y=70)
        lv = tk.Label(self.master, text="kV range = ")
        lv.place(x=40,y=100)
        #
        self.ar = tk.StringVar()
        self.ta = tk.Entry(self.master,width= 20,textvariable=self.ar)
        self.ta.place(x=160,y=40)
        #
        self.zo = tk.StringVar()
        self.tz = tk.Entry(self.master,width= 20,textvariable=self.zo)
        self.tz.place(x=160,y=70)
        #
        self.kv = tk.StringVar()
        self.kv.set("0-9999")
        self.tkv = tk.Entry(self.master,width= 20,textvariable=self.kv)
        self.tkv.place(x=160,y=100)
        #
        self.button_ar = tk.Button(self.master,text = '...',width = 5,height =1,command=self.selectArea)
        self.button_ar.place(x=320, y=35)
        #
        self.button_zo = tk.Button(self.master,text = '...',width = 5,height =1,command=self.selectZone)
        self.button_zo.place(x=320, y=65)
        #
        button_OK = tk.Button(self.master,text = 'OK',width = 10,command=self.run_OK)
        button_OK.place(x=100, y=150)
        #
        button_exit = tk.Button(self.master,text = 'Cancel',width = 10,command=self.exit)
        button_exit.place(x=220, y=150)
        #
        self.getData()
        self.radioAll()
    #
    def exit(self):
        try:
            self.root.destroy()
        except:
            pass
        self.master.destroy()
    #
    def radioAll(self):
        self.var1.set(1)
        self.ar.set('')
        self.zo.set('')
        self.ta.config(state='disabled')
        self.tz.config(state='disabled')
        self.button_ar['state']='disabled'
        self.button_zo['state']='disabled'
    #
    def radioArea(self):
        self.var1.set(2)
        self.zo.set('')
        self.ta.config(state='normal')
        self.tz.config(state='disabled')
        self.button_ar['state']='active'
        self.button_zo['state']='disabled'
    #
    def radioZone(self):
        self.var1.set(3)
        self.ar.set('')
        self.ta.config(state='disabled')
        self.tz.config(state='normal')
        self.button_ar['state']='disabled'
        self.button_zo['state']='active'
    #
    def getData(self):
        self.busHnd = OlxAPILib.getEquipmentHandle(TC_BUS)
        ar = OlxAPILib.getEquipmentData(self.busHnd,BUS_nArea)
        zo = OlxAPILib.getEquipmentData(self.busHnd,BUS_nZone)
        self.arA = list(set(ar))
        self.zoA = list(set(zo))
    #
    def selectArea(self):
        data = []
        for a1 in self.arA:
            data.append(['AREA # ' +str(a1),a1])
        self.root = tk.Tk()
        GuiSelectListBox(self.root,'Select Area',data,self.ar)
    #
    def selectZone(self):
        data = []
        for z1 in self.zoA:
            data.append(['ZONE # ' +str(z1),z1])
        self.root = tk.Tk()
        GuiSelectListBox(self.root,'Select Zone',data,self.zo)
    #
    def run_OK(self):
        ##<LINESELECTOR AREAS="1,2,3" ZONES="4,5,6" KVRANGE="100-300"/>
        r = self.var1.get()
        if r==1:# all network
            sa = ""
            for v1 in self.arA:
                sa += str(v1) + ","
            sa = sa[:len(sa)-1]
            #
            sz = ""
            for v1 in self.zoA:
                sz += str(v1) + ","
            sz = sz[:len(sz)-1]
            #
            kv = str(self.kv.get()).strip()
        else: # subsystem
            sa = str(self.ar.get()).strip()
            sz = str(self.zo.get()).strip()
            kv = str(self.kv.get()).strip()
        # add
        data = ET.Element('LINESELECTOR')
        data.set('AREAS', sa)
        data.set('ZONES', sz)
        data.set('KVRANGE', kv)
        xml_string = ET.tostring(data).decode('UTF-8')
        ARGVS.fo = OlxAPILib.get_fo(ARGVS.opath,ARGVS.fo,PY_FILE,'_out','.XML')
        OlxAPILib.saveString2File(ARGVS.fo,xml_string)
        print("LineSelector_out saved as: "+ os.path.abspath(ARGVS.fo))
        self.exit()
#
def run1(olxpath,opath,fi,fo):
    ARGVS.olxpath = olxpath
    ARGVS.opath = opath
    ARGVS.fi = fi
    ARGVS.fo = fo
    run()
#
def run():
    OlxAPILib.open_olrFile_1(ARGVS.olxpath,ARGVS.fi,readonly=1)
    #
    root = tk.Tk()
    feedback = MainGUI(root)
    root.mainloop()
#
if __name__ == '__main__':
    if ARGVS.fi=='':
        ARGVS.fi = 'SAMPLE30.OLR'
    #
    run()


