"""
Purpose:
    Report tielines/tiebranches by Area/Zone

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2021, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Release"
__version__   = "1.2.1"

# IMPORT -----------------------------------------------------------------------
import logging
import sys,os
PATH_FILE,PY_FILE = os.path.split(os.path.abspath(__file__))
PATH_LIB = os.path.split(PATH_FILE)[0]
if PATH_LIB not in sys.path:
    os.environ['PATH'] = PATH_LIB + ";" + os.environ['PATH']
    sys.path.insert(0, PATH_LIB)
#
import OlxAPILib
from OlxAPIConst import *
import AppUtils
AppUtils.chekPythonVersion(PY_FILE)
#
import tkinter as tk
from tkinter import ttk

# INPUTS cmdline ---------------------------------------------------------------
PARSER_INPUTS = AppUtils.iniInput(usage=
         "\n\tTest GUI and data OLR access\
         \n\tReport tielines/tiebranches by Area/Zone")
PARSER_INPUTS.add_argument('-fi' , help = '*(str) OLR file path', default = '',type=str,metavar='')
ARGVS = AppUtils.parseInput(PARSER_INPUTS)

# GUI
class APP(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        # get all area/zone
        self.area = OlxAPILib.getAll_Area()
        self.area.append(0)
        self.area.sort()
        #
        self.zone = OlxAPILib.getAll_Zone()
        self.zone.append(0)
        self.zone.sort()
        #
        self.initGUI()
    #
    def initGUI(self):
        self.parent.wm_title("Tieline Report: "+os.path.basename(ARGVS.fi))
        AppUtils.setIco_1Liner(self.parent)
        #
        la1 = tk.Label(self.parent, text="Type")
        la1.pack()
        la1.place(x=30, y=20)
        #
        self.var1 = tk.IntVar()
        self.r1=tk.Radiobutton(self.parent, text="TieLines"   ,variable = self.var1,value=1,command=self.getValLn)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.parent, text="TieBranches",variable = self.var1,value=2,command=self.getValBr)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=100,y=20)
        self.r2.place(x=200,y=20)
        self.r1.select()
        self.az="Area"
        #
        la2 = tk.Label(self.parent, text="Tie by ")
        la2.pack()
        la2.place(x=30, y=60)
        #
        self.var2 = tk.IntVar()
        self.r3=tk.Radiobutton(self.parent, text="Area",variable = self.var2, value=1,command=self.getValArea)
        self.r3.pack(anchor=tk.W)
        self.r4=tk.Radiobutton(self.parent, text="Zone",variable = self.var2, value=2,command=self.getValZone)
        self.r4.pack(anchor=tk.W)
        self.r3.place(x=100,y=60)
        self.r4.place(x=200,y=60)
        self.r3.select()
        self.LnBr="Line"
        #
        la2 = tk.Label(self.parent, text="From (0=all)")
        la2.pack()
        la2.place(x=30, y=100)
        #
        self.updateCombo()
        #
        button_OK = tk.Button(self.parent,text =   '     OK     ',command=self.run_OK)
        button_OK.place(x=100, y=150)
        #
        button_exit = tk.Button(self.parent,text = '   Exit   ',command=self.cancel)
        button_exit.place(x=200, y=150)

    #
    def cancel(self):
        # close OLR file @ TODO
        # quit
        self.parent.destroy()
    #
    def updateCombo(self):
        if self.az=="Area":
            self.cb=ttk.Combobox(self.parent, values=self.area,width = 6)
        elif self.az=="Zone":
            self.cb=ttk.Combobox(self.parent, values=self.zone,width = 6)
        #
        self.cb.place(x=100, y=100)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "Area"
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "Zone"
        self.updateCombo()
    #
    def getValLn(self):
        self.LnBr = "Line"
    #
    def getValBr(self):
        self.LnBr = "Branch"
       #
    def run_OK(self):
        azNum= int(self.cb.get())
        if self.az=="Area":
            nParamCode = BUS_nArea
        elif self.az=="Zone":
            nParamCode = BUS_nZone
        #
        if self.LnBr=="Line":
            equiCode = [TC_LINE]
        elif self.LnBr=="Branch":
            equiCode = []
        #
        tielines(nParamCode,azNum,equiCode)
# CORE
def tielines(nParamCode,azNum,equiCode):
    """
     Report tie branches by area or zone
        Args :
            nParamCode = BUS_nArea tie lines by AREA
                       = BUS_nZone tie lines by ZONE
            azNum : area/zone number (if <=0 all area/zone)
            equiCode: code of equipement
               example:
                    = [TC_LINE] => report tie lines
                    = [NULL]    => report tie branches
                    = [TC_LINE, TC_SCAP] => report tie lines and Serie capacitor/reactor

        Returns:
            sa: [(branche, (area1, area2)),... ]

        Raises:
            OlxAPIException
    """
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    #
    equiCode1 = equiCode
    if len(equiCode1)==0:
        equiCode1 = [TC_LINE, TC_SCAP, TC_PS, TC_SWITCH, TC_XFMR, TC_XFMR3]
    #
    abr = {}
    for ec1 in equiCode1:
        eqa = OlxAPILib.getEquipmentHandle(ec1)
        for ehnd1 in eqa:
            bus = OlxAPILib.getBusByEquipment(ehnd1,ec1)
            az = OlxAPILib.getEquipmentData(bus,nParamCode)
            #
            for i in range(len(bus)):
                for j in range(i+1, len(bus)):
                    if  (az[i]!= az[j]) and (azNum<=0 or az[i]==azNum or az[j]==azNum):
                        abr[ehnd1] = (min(az[i],az[j]),max(az[i],az[j]))
    # sort result
    sa = sorted(abr.items(), key=lambda kv: kv[1])

    #
    if len(equiCode)==1 and equiCode[0]==TC_LINE:
        if nParamCode==BUS_nArea:
            sres = "Tie lines report:\nNo   ,Area1   ,Area2   ,Lines (" + str(len(sa))+ ")\n"
        else:
            sres = "Tie lines report:\nNo   ,Zone1   ,Zone2   ,Lines (" + str(len(sa))+ ")\n"
        #
        s1 = "Tie lines"
    else:
        if nParamCode==BUS_nArea:
            sres = "Tie branches report:\nNo   ,Area1   ,Area2   ,Branches (" + str(len(sa))+ ")\n"
        else:
            sres = "Tie branches report:\nNo   ,Zone1   ,Zone2   ,Branches (" + str(len(sa))+ ")\n"
        #
        s1 = "Tie branches"
    #
    for i in range(len(sa)):
        a1 = sa[i][1][0]
        a2 = sa[i][1][1]
        ec1= sa[i][0]
        sres += str(i+1).ljust(5)+","+str(a1).ljust(8)+","+str(a2).ljust(8)+"," + OlxAPILib.fullBranchName(ec1)+"\n"
    #
    flog = AppUtils.logger2File(PY_FILE)
    logging.info("Found: "+str(len(sa))+" " +s1)
    logging.info(sres)
    logging.shutdown()
    #
    if ARGVS.ut==0:
        AppUtils.launch_notepad(flog)
        print("Report file had been saved in: " +flog)
    return flog
#
def unit_test():
    ARGVS.fi = AppUtils.getASPENFile(PATH_FILE,'SAMPLE30.OLR')
    #
    sres = "OLR file: " + os.path.basename(ARGVS.fi) + "\n"
    #
    nParamCode,azNum,equiCode = BUS_nArea,1,[TC_LINE]
    flog = tielines(nParamCode,azNum,equiCode)
    sres += AppUtils.read_File_text_2(flog,5)
    #
    nParamCode,azNum,equiCode = BUS_nArea,-1,[]
    flog = tielines(nParamCode,azNum,equiCode)
    sres += AppUtils.read_File_text_2(flog,5)
    #
    return AppUtils.unit_test_compare(PATH_FILE,PY_FILE,sres)
#
def run():
    if not AppUtils.checkInputOLR(ARGVS.fi,PY_FILE,PARSER_INPUTS):
        return
    #
    OlxAPILib.open_olrFile(ARGVS.fi,ARGVS.olxpath)
    root = tk.Tk()
    root.geometry("350x200+300+200")
    d = APP(root)
    root.mainloop()
#
def main():
    if ARGVS.ut>0:
        return unit_test()
    run()
#
if __name__ == '__main__':
    # ARGVS.ut = 1
    try:
        main()
    except Exception as err:
        AppUtils.FinishException(PY_FILE,err)
