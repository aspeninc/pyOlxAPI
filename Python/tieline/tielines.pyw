"""
Purpose:
    Report tielines/tiebranches by Area/Zone

"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced System for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "Common"
__pyManager__ = "Yes"
__email__     = "support@aspeninc.com"
__status__    = "Released"
__version__   = "2.1.1"


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
PARSER_INPUTS.usage ="Test GUI and data OLR access \
        \n\tReport tielines/tiebranches by Area/Zone\
        \n\tselected by GUI tkinter"
PARSER_INPUTS.add_argument('-fr' , help = ' (str) Path name report file .CSV',default = "", type=str, metavar='')
ARGVS = PARSER_INPUTS.parse_known_args()[0]


txtfont = ("MS Sans Serif", 8)
# CORE--------------------------------------------------------------------------
class GUI(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        # get all area/zone
        self.area = OLCase.AREANO
        self.area.append(0)
        self.area.sort()
        #
        self.zone = OLCase.ZONENO
        self.zone.append(0)
        self.zone.sort()
        #
        self.initGUI()
    #
    def initGUI(self):
        offX,offY,sw,sh = AppUtils.calOffsetMonitors(self.master)
        w,h = 500,260
        self.master.geometry("{0}x{1}+{2}+{3}".format(w,h,int(sw/2-w/2)+offX,int(sh/2-h/2)+offY))
        self.master.wm_title(PY_FILE)
        AppUtils.setIco_1Liner(self.master)
        #
        la1 = tk.Label(self.master, text="Type",font = txtfont)
        la1.pack()
        la1.place(x=30, y=20)
        #
        self.var1 = tk.IntVar()
        self.r1=tk.Radiobutton(self.master, text="TieLines"   ,variable = self.var1,value=1,font = txtfont,command=self.getValLn)
        self.r1.pack(anchor=tk.W)
        self.r2=tk.Radiobutton(self.master, text="TieBranches",variable = self.var1,value=2,font = txtfont,command=self.getValBr)
        self.r2.pack(anchor=tk.W)
        self.r1.place(x=150,y=20)
        self.r2.place(x=300,y=20)
        self.r1.select()
        self.az = "AREANO"
        #
        la2 = tk.Label(self.master, text="Tie by ",font = txtfont)
        la2.pack()
        la2.place(x=30, y=80)
        #
        self.var2 = tk.IntVar()
        self.r3=tk.Radiobutton(self.master, text="Area",variable = self.var2, value=1,font = txtfont,command=self.getValArea)
        self.r3.pack(anchor=tk.W)
        self.r4=tk.Radiobutton(self.master, text="Zone",variable = self.var2, value=2,font = txtfont,command=self.getValZone)
        self.r4.pack(anchor=tk.W)
        self.r3.place(x=150,y=80)
        self.r4.place(x=300,y=80)
        self.r3.select()
        self.LnBr = "LINE"
        #
        la2 = tk.Label(self.master, text="From (0=all)",font = txtfont)
        la2.pack()
        la2.place(x=30, y=130)
        #
        self.updateCombo()
        #
        button_OK = tk.Button(self.master,text = 'OK', width = 10, command=self.run_OK,font = txtfont)
        button_OK.place(x=150, y=200)
        #
        button_exit = tk.Button(self.master,text = 'Exit',width = 10, command=self.cancel,font = txtfont)
        button_exit.place(x=300, y=200)

    #
    def cancel(self):
        self.master.destroy()
        self.master.quit()
    #
    def updateCombo(self):
        if self.az=="AREANO":
            self.cb=ttk.Combobox(self.master, values=self.area,width = 6,font = txtfont)
        elif self.az=="ZONENO":
            self.cb=ttk.Combobox(self.master, values=self.zone,width = 6,font = txtfont)
        #
        self.cb.place(x=150, y=130)
        self.cb.bind("<<ComboboxSelected>>")
        self.cb.current(0)
    #
    def getValArea(self):
        self.az = "AREANO"
        self.updateCombo()
    #
    def getValZone(self):
        self.az = "ZONENO"
        self.updateCombo()
    #
    def getValLn(self):
        self.LnBr = "LINE"
    #
    def getValBr(self):
        self.LnBr = "BRANCH"
       #
    def run_OK(self):
        azNum= int(self.cb.get())
        sres = "Tie %s report:\nNo   ,%s1 ,%s2 ,"%(self.LnBr,self.az,self.az)

        bra = OLCase.LINE
        if self.LnBr=='BRANCH':
            bra.extend(OLCase.SWITCH)
            bra.extend(OLCase.SERIESRC)
            bra.extend(OLCase.SHIFTER)
            bra.extend(OLCase.XFMR)
            bra.extend(OLCase.XFMR3)
        #
        k = 0
        for br1 in bra:
            ba = [b1.getData(self.az) for b1 in br1.BUS]
            ba.sort()
            if ba[0]!=ba[-1]:
                if (azNum==0) or azNum in ba:
                    k+=1
                    sres += '\n'+str(k).ljust(5)
                    for b1 in ba:
                        sres+=','+str(b1).ljust(8)
                    sres+=','+br1.toString()

        ARGVS.fr = AppUtils.get_file_out(fo=ARGVS.fr, fi=OLCase.olrFile , subf='' , ad='_'+PY_FILE[:-4] , ext='.csv')
        AppUtils.saveString2File(ARGVS.fr,sres)
        s1 = '\nReport file had been saved in:\n%s'%ARGVS.fr
        AppUtils.explorerDir(ARGVS.fr, s1, PY_FILE) #open dir of fo
        self.cancel()

#
def main():
    OLCase.checkInit(PY_FILE) # check if ASPEN OLR file is opened
    #
    master = tk.Tk()
    GUI(master)
    master.mainloop()

#
if __name__ == '__main__':
    main()

