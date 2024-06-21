"""
Purpose: Change all bus data: Area/Zone in OLR file
                          from one value to another

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Released"
__version__   = "2.1.2"


# IMPORT -----------------------------------------------------------------------
from OlxObj import *
import AppUtils
import os
import tkinter as tk
from tkinter import ttk
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))

# INPUT Command Line Arguments
import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage ="Change all bus data: \
        \nArea/Zone in OLR file from one value to another\
        \nselected by GUI tkinter"
ARGVS = PARSER_INPUTS.parse_known_args()[0]

txtfont = ("MS Sans Serif", 8)
# CORE--------------------------------------------------------------------------
class GUI(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        #
        self.area = OLCase.AREANO
        self.zone = OLCase.ZONENO
        #
        self.initGUI()
    #
    def initGUI(self):
        offX,offY,sw,sh = AppUtils.calOffsetMonitors(self.master)
        w,h = 600,250
        self.master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2)+offX,int(sh/2-h/2)+offY))
        self.master.wm_title("Change AREA/ZONE")
        AppUtils.setIco_1Liner(self.master)
        #
        la1 = tk.Label(self.master, text="Change All",font = txtfont)
        la1.pack()
        la1.place(x=20, y=20)
        #
        self.var1 = tk.IntVar()
        #
        self.r1=tk.Radiobutton(self.master, text="Area", variable = self.var1, value=1, command=self.getValArea, font = txtfont)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.master, text="Zone", variable = self.var1, value=2, command=self.getValZone, font = txtfont)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=200,y=20)
        self.r2.place(x=400,y=20)
        #
        self.az = "Area"
        self.r1.select()
        #
        la2 = tk.Label(self.master, text="From",font = txtfont)
        la2.pack()
        la2.place(x=200, y=70)
        la3 = tk.Label(self.master, text="To",font = txtfont)
        la3.pack()
        la3.place(x=400, y=70)
        #
        self.updateCombo()
        #
        self.sto = tk.Text(self.master, height=1, width=6,font = txtfont)
        self.sto.pack()
        self.sto.place(x=400, y=120)
        #
        button_OK = tk.Button(self.master,text = 'OK',width=12, command=self.run_OK,font = txtfont)
        button_OK.place(x=260, y=200)
    #
    def quit(self): # quit
        self.master.destroy()
        self.master.quit()

    #
    def updateCombo(self):
        if self.az=="Area":
            self.cb = ttk.Combobox(self.master, values=self.area,width = 6,font = txtfont)
        elif self.az=="Zone":
            self.cb = ttk.Combobox(self.master, values=self.zone,width = 6,font = txtfont)
        #
        self.cb.place(x=200, y=120)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "Area"
        print("selected: "+self.az)
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "Zone"
        print("selected: "+self.az)
        self.updateCombo()
    #
    def run_OK(self):
        valFrom = int(self.cb.get())
        try:
            valTo = int(self.sto.get("1.0",tk.END))
        except:
            AppUtils.gui_error(sTitle='ERROR',sMain="To value: "+self.sto.get("1.0",tk.END)+"\nValue To must have INTEGER type")
            return
        #
        if self.az=="Area":
            for b1 in OLCase.BUS:
                if b1.AREANO==valFrom:
                    b1.AREANO = valTo
                    b1.postData()
                    print('\t', b1.toString())
        elif self.az=="Zone":
            for b1 in OLCase.BUS:
                if b1.ZONENO==valFrom:
                    b1.ZONENO = valTo
                    b1.postData()
                    print('\t', b1.toString())
        AppUtils.gui_info(PY_FILE,'Change %s %i=>%i'%(self.az,valFrom, valTo))
        self.quit()


def main():
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened
    #
    master = tk.Tk()
    GUI(master)
    master.mainloop()


if __name__ == '__main__':
    main()





