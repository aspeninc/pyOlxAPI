"""
Purpose: Change all bus data: Area/Zone in OLR file
                          from one value to another

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

#
import sys,os
#
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPILib
from OlxAPIConst import *
import AppUtils

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage="\n\tTest GUI and data OLR access\
                      \n\tChange all bus data: Area/Zone in OLR file\
                      \n\t\t               from one value to another")#
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS)

import tkinter as tk
import tkinter.filedialog as tkf
import tkinter.messagebox as tkm
from tkinter import ttk

# GUI
class APP(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        #
        busHnd = OlxAPILib.getEquipmentHandle(TC_BUS)
        ar = OlxAPILib.getEquipmentData(busHnd,BUS_nArea)
        self.area = list(set(ar))
        zo = OlxAPILib.getEquipmentData(busHnd,BUS_nZone)
        self.zone = list(set(zo))
        #
        self.initGUI()
    #
    def initGUI(self):
        olrFile = os.path.basename(ARGVS.fi)
        self.sw = self.parent.winfo_screenwidth()
        self.sh = self.parent.winfo_screenheight()
        w,h = 320,200 #
        self.parent.geometry("{0}x{1}+{2}+{3}".format(w,h,int(self.sw/2-w/2),int(self.sh/2-h/2)))
        self.parent.resizable(0,0)# fixed size

        self.parent.wm_title("Change bus: "+olrFile)
        AppUtils.setIco_1Liner(self.parent)
        AppUtils.remove_button(self.parent)
        #
        la1 = tk.Label(self.parent, text="Change All")
        la1.pack()
        la1.place(x=20, y=20)
        #
        self.var1 = tk.IntVar()
        #
        self.r1=tk.Radiobutton(self.parent, text="Area",variable = self.var1,value=1, command=self.getValArea)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.parent, text="Zone",variable = self.var1,value=2, command=self.getValZone)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=100,y=20)
        self.r2.place(x=200,y=20)
        #
        self.az = "Area"
        self.r1.select()
        #
        la2 = tk.Label(self.parent, text="From                         To")
        la2.pack()
        la2.place(x=100, y=70)
        #
        self.updateCombo()
        #
        self.sto = tk.Text(self.parent, height=1, width=6)
        self.sto.pack()
        self.sto.place(x=200, y=90)
        #
        button_OK = tk.Button(self.parent,text =   '     OK     ',command=self.run_OK)
        button_OK.place(x=100, y=150)
        #
        button_exit = tk.Button(self.parent,text = '   Exit   ',command=self.cancel)
        button_exit.place(x=200, y=150)
    #
    def cancel(self):
        # quit
        self.parent.destroy()
    #
    def updateCombo(self):
        if self.az=="Area":
            self.cb=ttk.Combobox(self.parent, values=self.area,width = 6)
        elif self.az=="Zone":
            self.cb=ttk.Combobox(self.parent, values=self.zone,width = 6)
        #
        self.cb.place(x=100, y=90)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "Area"
        print("\tselected: "+self.az)
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "Zone"
        print("\tselected: "+self.az)
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
            nParamCode = BUS_nArea
            if valTo not in self.area:
                AppUtils.gui_error(sTitle='ERROR',sMain="To value: "+self.sto.get("1.0",tk.END)+"\nBus Area number does not exist")
                return
        elif self.az=="Zone":
            nParamCode = BUS_nZone
            if valTo not in self.zone:
                AppUtils.gui_error(sTitle='ERROR',sMain="To value: "+self.sto.get("1.0",tk.END)+"\nBus Zone number does not exist")
                return
        #
        bres = changebus(nParamCode,valFrom,valTo)
        if len(bres) ==0:
            AppUtils.gui_info(sTitle='INFO',sMain="No bus changed")
        else:
            AppUtils.gui_info(sTitle='INFO',sMain =str(len(bres))+" bus changed" )
            sres = changebus_str(nParamCode,valFrom,valTo,bres)
            print(sres)
            #
            olrNew = tkf.asksaveasfilename(defaultextension="OLR",filetypes=[("Python Files", "*.OLR"), ("All Files", "*.*")])
            if olrNew:
                OlxAPILib.saveAsOlr(olrNew)
                print("File had been saved in: ",olrNew)
                AppUtils.launch_OneLiner(olrNew)
            else:
                OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
# CORE
def changebus(nParamCode,valFrom,valTo):
    """
    Change all bus data
        Args :
            nParamCode : BUS_nArea/BUS_nZone
            valFrom,valTo : valFrom => valTo
        Returns:
            NONE
        Raises:
            OlxAPI.Exception
    """
    bres = []
    if valFrom==valTo:
        return bres
    #
    ba = OlxAPILib.getEquipmentHandle(TC_BUS) # get all bus handle
    az = OlxAPILib.getEquipmentData(ba,nParamCode) # get area/zone of bus
    #
    for i in range(len(ba)):
        b1 = ba[i]
        a1 = az[i]
        if a1 == valFrom:
            OlxAPILib.setData(b1,nParamCode,valTo)
            # if validation OK
            if OlxAPILib.postData(b1):
                bres.append(b1)
    return bres

def changebus_str(nParamCode,valFrom,valTo,bres):
    sres = str(len(bres))
    if nParamCode == BUS_nArea:
        sres+=" bus have changed Area ("
    elif nParamCode == BUS_nZone:
        sres+=" bus have changed Zone ("
    sres+= str(valFrom) + "=>"+ str(valTo) + ")"
    for i in range( len(bres)):
        sres+="\n\t" +str(i+1) + " " + OlxAPILib.fullBusName(bres[i])
    return sres

def unit_test():
    ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    sres = "OLR file: " + os.path.basename(ARGVS.fi) + "\n"
    #
    valFrom = 1
    valTo   = 10
    bres = changebus(BUS_nArea,valFrom,valTo)
    sres += changebus_str(BUS_nArea,valFrom,valTo,bres)
    #
    bres = changebus(BUS_nZone,valFrom,valTo)
    sres += '\n'+changebus_str(BUS_nZone,valFrom,valTo,bres)
    #
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)
#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    root = tk.Tk()
    root.geometry("350x200+300+200")
    d = APP(root)
    root.mainloop()
#
def main():
    if ARGVS.ut >0:
        return unit_test()
    run()
#
if __name__ == '__main__':
    # ARGVS.ut = 1
    try:
        main()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)




