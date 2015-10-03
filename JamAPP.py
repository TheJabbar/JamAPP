import tkinter as tk
from tkinter import *
import tkinter.font as tkFont
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror


import xlrd
import os
import math
import random
import itertools
import copy
import sys

Large_Font = ("CaviaDream", 16)
Small_Font = ("Consolas", 10)
Smaller_Font = ("Consolas", 8)
XY = "800x600"

class JAMapp(tk.Tk):

    def __init__(self, *args, **kwargs):

        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)

        container.pack(side="top", fill="both", expand = True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        
        self.frames = {}

        for F in (StartPage, InputPage, InputTestPage):
            
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    def show_frame(self, cont):

        frame = self.frames[cont]
        frame.tkraise()

    
class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        label = ttk.Label(self, text="JPA Ant-Miner Menu", font=Large_Font)
        label.pack(pady=10,padx=10)

        button1 = ttk.Button(self, text="Input Students Data and Train",
                            command=lambda: controller.show_frame(InputPage))
        button1.pack()

        button2 = ttk.Button(self, text="Input Training Rule and Test",
                            command=lambda: controller.show_frame(InputTestPage))
        button2.pack()
        label1 = ttk.Label(self, text="Muhammad Abdul Jabbar           1107110039", font=Small_Font)
        label1.pack(pady=10,padx=10)

class InputTestPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        label = ttk.Label(self, text="Testing Page", font=Large_Font)
        label.grid(row=0, column=3, columnspan=4)

        button1 = ttk.Button(self, text="Back To Menu",
                            command=lambda: controller.show_frame(StartPage))
        button1.grid(row=1, column=3, columnspan=4)

        button2 = ttk.Button(self, text="Input Students Data and Train",
                            command=lambda: controller.show_frame(InputPage))
        button2.grid(row=2,column=3, columnspan=4)
        
        EntryMaxMin = ttk.Label(self, text="Insert Minimum Value of Atributes", font=Small_Font)
        EntryMaxMin.grid(row=3,column=3, columnspan=4)

        SemLabel =[0]*17
        for i in range(1,17):
            Sems = 'Atr',i
            SemLabel[i] = ttk.Label(self, text=Sems, font=Smaller_Font)
            if i > 8:
                SemLabel[i].grid(row=6,column=i-8)
            else:
                SemLabel[i].grid(row=4,column=i)
        
        self.Min1 = StringVar()
        Min1Ent = ttk.Entry(self, width=3, textvariable=self.Min1)
        Min1Ent.grid(row=5,column=1)
        Min1Ent.delete(0, END)
        Min1Ent.insert(0, 60)
        self.Min1.set(60)

        self.Min2 = StringVar()
        Min2Ent = ttk.Entry(self, width=3, textvariable=self.Min2)
        Min2Ent.grid(row=5,column=2)
        Min2Ent.delete(0, END)
        Min2Ent.insert(0, 55)
        self.Min2.set(55)

        self.Min3 = StringVar()
        Min3Ent = ttk.Entry(self, width=3, textvariable=self.Min3)
        Min3Ent.grid(row=5,column=3)
        Min3Ent.delete(0, END)
        Min3Ent.insert(0, 60)
        self.Min3.set(60)

        self.Min4 = StringVar()
        Min4Ent = ttk.Entry(self, width=3, textvariable=self.Min4)
        Min4Ent.grid(row=5,column=4)
        Min4Ent.delete(0, END)
        Min4Ent.insert(0, 0)
        self.Min4.set(0)

        self.Min5 = StringVar()
        Min5Ent = ttk.Entry(self, width=3, textvariable=self.Min5)
        Min5Ent.grid(row=5,column=5)
        Min5Ent.delete(0, END)
        Min5Ent.insert(0, 60)
        self.Min5.set(60)

        self.Min6 = StringVar()
        Min6Ent = ttk.Entry(self, width=3, textvariable=self.Min6)
        Min6Ent.grid(row=5,column=6)
        Min6Ent.delete(0, END)
        Min6Ent.insert(0, 64)
        self.Min6.set(64)

        self.Min7 = StringVar()
        Min7Ent = ttk.Entry(self, width=3, textvariable=self.Min7)
        Min7Ent.grid(row=5,column=7)
        Min7Ent.delete(0, END)
        Min7Ent.insert(0, 50)
        self.Min7.set(50)

        self.Min8 = StringVar()
        Min8Ent = ttk.Entry(self, width=3, textvariable=self.Min8)
        Min8Ent.grid(row=5,column=8)
        Min8Ent.delete(0, END)
        Min8Ent.insert(0, 0)
        self.Min8.set(0)

        self.Min9 = StringVar()
        Min9Ent = ttk.Entry(self, width=3, textvariable=self.Min9)
        Min9Ent.grid(row=7,column=1)
        Min9Ent.delete(0, END)
        Min9Ent.insert(0, 70)
        self.Min9.set(70)

        self.Min10 = StringVar()
        Min10Ent = ttk.Entry(self, width=3, textvariable=self.Min10)
        Min10Ent.grid(row=7,column=2)
        Min10Ent.delete(0, END)
        Min10Ent.insert(0, 60)
        self.Min10.set(60)

        self.Min11 = StringVar()
        Min11Ent = ttk.Entry(self, width=3, textvariable=self.Min11)
        Min11Ent.grid(row=7,column=3)
        Min11Ent.delete(0, END)
        Min11Ent.insert(0, 60)
        self.Min11.set(60)

        self.Min12 = StringVar()
        Min12Ent = ttk.Entry(self,  width=3, textvariable=self.Min12)
        Min12Ent.grid(row=7,column=4)
        Min12Ent.delete(0, END)
        Min12Ent.insert(0, 0)
        self.Min12.set(0)
        
        self.Min13 = StringVar()
        Min13Ent = ttk.Entry(self, width=3, textvariable=self.Min13)
        Min13Ent.grid(row=7,column=5)
        Min13Ent.delete(0, END)
        Min13Ent.insert(0, 70)
        self.Min13.set(70)

        self.Min14 = StringVar()
        Min14Ent = ttk.Entry(self, width=3, textvariable=self.Min14)
        Min14Ent.grid(row=7,column=6)
        Min14Ent.delete(0, END)
        Min14Ent.insert(0, 65)
        self.Min14.set(65)

        self.Min15 = StringVar()
        Min15Ent = ttk.Entry(self, width=3, textvariable=self.Min15)
        Min15Ent.grid(row=7,column=7)
        Min15Ent.delete(0, END)
        Min15Ent.insert(0, 61)
        self.Min15.set(61)

        self.Min16 = StringVar()
        Min16Ent = ttk.Entry(self, width=3, textvariable=self.Min16)
        Min16Ent.grid(row=7,column=8)
        Min16Ent.delete(0, END)
        Min16Ent.insert(0, 0)
        self.Min16.set(0)

        EntryJurusan = ttk.Label(self, text="Insert Major", font=Small_Font)
        EntryJurusan.grid(row=13,column=3, columnspan=4)

        self.Jurusan = StringVar()
        JurusanEnt = ttk.Combobox(self, width=28, justify=CENTER, textvariable=self.Jurusan)
        JurusanEnt.grid(row=14,column=3, columnspan=4)
        JurusanEnt['values'] = ("S1 Teknik Telekomunikasi",
                 "S1 Teknik Informatika",
                 "S1 Teknik Industri",
                 "S1 International ICT Business",
                 "S1 MBTI","S1 Teknik Elektro",
                 "S1 Sistem Komputer",
                 "S1 Teknik Fisika",
                 "S1 Ilmu Komputasi",
                 "S1 Ilmu Komunikasi",
                 "S1 Sistem Informasi",
                 "S1 Akuntansi",
                 "S1 Administrasi Bisnis",
                 "S1 Desain Komunikasi Visual",
                 "S1 Desain Interior",
                 "D3 Teknik Informatika",
                 "D3 Teknik Telekomunikasi",
                 "S1 Kriya Tekstil dan Mode",
                 "S1 Desain Produk",
                 "S1 Seni Rupa Murni",
                 "D3 Teknik Komputer",
                 "D3 Manajemen Informatika",
                 "D3 Komputerisasi Akuntansi",
                 "D3 Manajemen Pemasaran",
                 "D3 Perhotelan")
        JurusanEnt.delete(0, END)
        JurusanEnt.insert(0, 'S1 Teknik Telekomunikasi')
        self.Jurusan.set('S1 Teknik Telekomunikasi')
        
        EntryBestRule = ttk.Label(self, text="Insert Best Rule", font=Small_Font)
        EntryBestRule.grid(row=8,column=3, columnspan=4)

        BestRuleLabel =[0]*19
        for i in range(1,19):
            BLabel = 'Atr',i
            BestRuleLabel[i] = ttk.Label(self, text=BLabel, font=Smaller_Font)
            if i==17:
                BestRuleLabel[i] = ttk.Label(self, text='Tier', font=Smaller_Font)
            if i==18:
                BestRuleLabel[i] = ttk.Label(self, text='Class', font=Smaller_Font)
            if i > 8:
                BestRuleLabel[i].grid(row=11,column=i-8)
            else:
                BestRuleLabel[i].grid(row=9,column=i)
        
        self.BestRule = [0]*19
        BestRuleEnt = [0]*19
        for i in range(1,19):
            self.BestRule[i] = StringVar()
            BestRuleEnt[i] = ttk.Combobox(self, width=8, textvariable=self.BestRule[i])
            if i > 8:
                BestRuleEnt[i].grid(row=12,column=i-8)
            else:
                BestRuleEnt[i].grid(row=10,column=i)    
            BestRuleEnt[i].delete(0, END)
            if i==17:
                BestRuleEnt[i]['values'] = ('Tier1', 'Tier2', 'Tier3')
                BestRuleEnt[i].insert(0, 'Tier1')
                self.BestRule[i].set('Tier1')
            elif i==18:
                BestRuleEnt[i]['values'] = ('Lulus', 'Tidak Lulus')
                BestRuleEnt[i].insert(0, 'Lulus')
                self.BestRule[i].set('Lulus')
            else:
                BestRuleEnt[i]['values'] = ('-','Tinggi', 'Menengah', 'Rendah')
                BestRuleEnt[i].insert(0, '-')
                self.BestRule[i].set('-')
        
        self.Nil = [0]*18
        NilEnt = [0]*18
        EntryNilai = ttk.Label(self, text="Insert Scores", font=Small_Font)
        EntryNilai.grid(row=15,column=3, columnspan=4)

        ScoreLabel =[0]*17
        for i in range(1,17):
            Scores = 'Atr',i
            ScoreLabel[i] = ttk.Label(self, text=Scores, font=Smaller_Font)
            ScoreLabel[i].grid(row=4,column=19)
            if i > 8:
                ScoreLabel[i].grid(row=18,column=i-8)
            else:
                ScoreLabel[i].grid(row=16,column=i)
        
        for i in range(1,17):
            self.Nil[i] = StringVar()
            NilEnt[i] = ttk.Entry(self, width=3, textvariable=self.Nil[i])
            if i > 8:
                NilEnt[i].grid(row=19,column=i-8)
            else:
                NilEnt[i].grid(row=17,column=i)    
            NilEnt[i].delete(0, END)
            NilEnt[i].insert(0, 0)
            self.Nil[i].set(0)

        Lulusbutton = ttk.Button(self, text="Find Out!", command= self.RunLulus)
        Lulusbutton.grid(row=20,column=3, columnspan=4)

        ResultLabel = ttk.Label(self, text="Result", font=Small_Font)
        ResultLabel.grid(row=21,column=3, columnspan=4)
        self.ket=StringVar()
        tex = tk.Entry(self, textvariable=self.ket)
        tex.grid(row=22, column=3, columnspan=4)       
        
       
    def RunLulus(self):
        min1 = int(self.Min1.get())
        min2 = int(self.Min2.get())
        min3 = int(self.Min3.get())
        min4 = int(self.Min4.get())
        min5 = int(self.Min5.get())
        min6 = int(self.Min6.get())
        min7 = int(self.Min7.get())
        min8 = int(self.Min8.get())
        min9 = int(self.Min9.get())
        min10 = int(self.Min10.get())
        min11 = int(self.Min11.get())
        min12 = int(self.Min12.get())
        min13 = int(self.Min13.get())
        min14 = int(self.Min14.get())
        min15 = int(self.Min15.get())
        min16 = int(self.Min16.get())

        Tier1 = ["S1 Teknik Telekomunikasi",
                 "S1 Teknik Informatika",
                 "S1 Teknik Industri",
                 "S1 International ICT Business",
                 "S1 MBTI"]
        Tier2 = ["S1 Teknik Elektro",
                 "S1 Sistem Komputer",
                 "S1 Teknik Fisika",
                 "S1 Ilmu Komputasi",
                 "S1 Ilmu Komunikasi",
                 "S1 Sistem Informasi",
                 "S1 Akuntansi",
                 "S1 Administrasi Bisnis",
                 "S1 Desain Komunikasi Visual",
                 "S1 Desain Interior",
                 "D3 Teknik Informatika",
                 "D3 Teknik Telekomunikasi"]
        Tier3 = ["S1 Kriya Tekstil dan Mode",
                 "S1 Desain Produk",
                 "S1 Seni Rupa Murni",
                 "D3 Teknik Komputer",
                 "D3 Manajemen Informatika",
                 "D3 Komputerisasi Akuntansi",
                 "D3 Manajemen Pemasaran",
                 "D3 Perhotelan"]
        
        LimLow1 = min1+0.333*(100-min1)
        LimMed1 = min1+0.666*(100-min1)
        LimLow2 = min2+0.333*(100-min2)
        LimMed2 = min2+0.666*(100-min2)
        LimLow3 = min3+0.333*(100-min3)
        LimMed3 = min3+0.666*(100-min3)
        LimLow4 = min4+0.333*(100-min4)
        LimMed4 = min4+0.666*(100-min4)
        LimLow5 = min5+0.333*(100-min5)
        LimMed5 = min5+0.666*(100-min5)
        LimLow6 = min6+0.333*(100-min6)
        LimMed6 = min6+0.666*(100-min6)
        LimLow7 = min7+0.333*(100-min7)
        LimMed7 = min7+0.666*(100-min7)
        LimLow8 = min8+0.333*(100-min8)
        LimMed8 = min8+0.666*(100-min8)
        LimLow9 = min9+0.333*(100-min9)
        LimMed9 = min9+0.666*(100-min9)
        LimLow10 = min10+0.333*(100-min10)
        LimMed10 = min10+0.666*(100-min10)
        LimLow11 = min11+0.333*(100-min11)
        LimMed11 = min11+0.666*(100-min11)
        LimLow12 = min12+0.333*(100-min12)
        LimMed12 = min12+0.666*(100-min12)
        LimLow13 = min13+0.333*(100-min13)
        LimMed13 = min13+0.666*(100-min13)
        LimLow14 = min14+0.333*(100-min14)
        LimMed14 = min14+0.666*(100-min14)
        LimLow15 = min15+0.333*(100-min15)
        LimMed15 = min15+0.666*(100-min15)
        LimLow16 = min16+0.333*(100-min16)
        LimMed16 = min16+0.666*(100-min16)

        nil = [0]*17
        for i in range(1,17):
            nil[i] = int(self.Nil[i].get())
            
        BestRule = [0]*19
        for i in range(1,19):
            BestRule[i] = self.BestRule[i].get()
        
        if nil[1] < LimLow1:
            nil[1] = 'Rendah'
        elif (nil[1] > LimLow1) and (nil[1] < LimMed1):
            nil[1] = 'Menengah'
        elif nil[1] > LimMed1:
            nil[1] = 'Tinggi'
            
        if nil[2] < LimLow2:
            nil[2] = 'Rendah'
        elif (nil[2] > LimLow2) and (nil[2] < LimMed2):
            nil[2] = 'Menengah'
        elif nil[2] > LimMed2:
            nil[2] = 'Tinggi'

        if nil[3] < LimLow3:
            nil[3] = 'Rendah'
        elif (nil[3] > LimLow3) and (nil[3] < LimMed3):
            nil[3] = 'Menengah'
        elif nil[3] > LimMed3:
            nil[3] = 'Tinggi'

        if nil[4] < LimLow4:
            nil[4] = 'Rendah'
        elif (nil[4] > LimLow4) and (nil[4] < LimMed4):
            nil[4] = 'Menengah'
        elif nil[4] > LimMed4:
            nil[4] = 'Tinggi'

        if nil[5] < LimLow5:
            nil[5] = 'Rendah'
        elif (nil[5] > LimLow5) and (nil[5] < LimMed5):
            nil[5] = 'Menengah'
        elif nil[5] > LimMed5:
            nil[5] = 'Tinggi'

        if nil[6] < LimLow6:
            nil[6] = 'Rendah'
        elif (nil[6] > LimLow6) and (nil[6] < LimMed6):
            nil[6] = 'Menengah'
        elif nil[6] > LimMed6:
            nil[6] = 'Tinggi'

        if nil[7] < LimLow7:
            nil[7] = 'Rendah'
        elif (nil[7] > LimLow7) and (nil[7] < LimMed7):
            nil[7] = 'Menengah'
        elif nil[7] > LimMed7:
            nil[7] = 'Tinggi'

        if nil[8] < LimLow8:
            nil[8] = 'Rendah'
        elif (nil[8] > LimLow8) and (nil[8] < LimMed8):
            nil[8] = 'Menengah'
        elif nil[8] > LimMed8:
            nil[8] = 'Tinggi'

        if nil[9] < LimLow9:
            nil[9] = 'Rendah'
        elif (nil[9] > LimLow9) and (nil[9] < LimMed9):
            nil[9] = 'Menengah'
        elif nil[9] > LimMed9:
            nil[9] = 'Tinggi'

        if nil[10] < LimLow10:
            nil[10] = 'Rendah'
        elif (nil[10] > LimLow10) and (nil[10] < LimMed10):
            nil[10] = 'Menengah'
        elif nil[10] > LimMed10:
            nil[10] = 'Tinggi'

        if nil[11] < LimLow11:
            nil[11] = 'Rendah'
        elif (nil[11] > LimLow11) and (nil[11] < LimMed11):
            nil[11] = 'Menengah'
        elif nil[11] > LimMed11:
            nil[11] = 'Tinggi'

        if nil[12] < LimLow12:
            nil[12] = 'Rendah'
        elif (nil[12] > LimLow12) and (nil[12] < LimMed12):
            nil[12] = 'Menengah'
        elif nil[12] > LimMed12:
            nil[12] = 'Tinggi'

        if nil[13] < LimLow13:
            nil[13] = 'Rendah'
        elif (nil[13] > LimLow13) and (nil[13] < LimMed13):
            nil[13] = 'Menengah'
        elif nil[13] > LimMed13:
            nil[13] = 'Tinggi'

        if nil[14] < LimLow14:
            nil[14] = 'Rendah'
        elif (nil[14] > LimLow14) and (nil[14] < LimMed14):
            nil[14] = 'Menengah'
        elif nil[14] > LimMed14:
            nil[14] = 'Tinggi'

        if nil[15] < LimLow15:
            nil[15] = 'Rendah'
        elif (nil[15] > LimLow15) and (nil[15] < LimMed15):
            nil[15] = 'Menengah'
        elif nil[15] > LimMed15:
            nil[15] = 'Tinggi'

        if nil[16] < LimLow16:
            nil[16] = 'Rendah'
        elif (nil[16] > LimLow16) and (nil[16] < LimMed16):
            nil[16] = 'Menengah'
        elif nil[16] > LimMed16:
            nil[16] = 'Tinggi'

        Jurusan = self.Jurusan.get()

        if (Jurusan in Tier1):
            Jurusan = "Tier1"
        elif (Jurusan in Tier2):
            Jurusan = "Tier2"
        elif (Jurusan in Tier3):
            Jurusan = "Tier3"
        Pass = [0]*17
        
        for i in range(1,17):
            if (nil[i] == 'Tinggi' and (BestRule[i]=='Tinggi' or BestRule[i]=='Menengah' or BestRule[i]=='Rendah' or BestRule[i]=='-')):
                Pass[i] = 1
            elif (nil[i] == 'Menengah' and (BestRule[i]=='Menengah' or BestRule[i]=='Rendah' or BestRule[i]=='-')):
                Pass[i] = 1
            elif (nil[i] == 'Rendah' and (BestRule[i]=='Rendah' or BestRule[i]=='-')):
                Pass[i] = 1
            elif (nil[i] == '-' and (BestRule[i]=='-')):
                Pass[i] = 1
            else:
                Pass[i] = 0
        del Pass[0]
        if (Pass == [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1]):
            YesPass = True
        else:
            YesPass = False
        if (YesPass and (BestRule[17] == 'Tier1') and (Jurusan=='Tier1' or Jurusan=='Tier2' or Jurusan=='Tier3')):
            self.ket.set(BestRule[18])
        elif (YesPass and (BestRule[17] == 'Tier2') and (Jurusan=='Tier2' or Jurusan=='Tier3')):
            self.ket.set(BestRule[18])
        elif (YesPass and (BestRule[17] == 'Tier3') and (Jurusan=='Tier3')):
            self.ket.set(BestRule[18])
        elif not YesPass:
            self.ket.set('Tidak Lulus')
        else:
            self.ket.set('?')
        

        
        

        

        
class InputPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        label = ttk.Label(self, text="Input Menu", font=Large_Font)
        label.pack(pady=10,padx=10)

        button1 = ttk.Button(self, text="Back to Menu",
                            command=lambda: controller.show_frame(StartPage))
        button1.pack()

        button2 = ttk.Button(self, text="Input Training Rule and Test",
                            command=lambda: controller.show_frame(InputTestPage))
        button2.pack()
        
        Entrybutton = ttk.Button(self, text="Browse and Run!", command= self.browseandrun)
        Entrybutton.pack()

        LabTotalAnt = ttk.Label(self, text="Number of Ants", font=Small_Font)
        LabTotalAnt.pack()
        self.TotalAntNum = StringVar()
        TotalAnt = ttk.Entry(self, textvariable=self.TotalAntNum)
        TotalAnt.pack()
        TotalAnt.delete(0, END)
        TotalAnt.insert(0, 40)
        self.TotalAntNum.set(40)
        
        

        LabMaxRules = ttk.Label(self, text="Maximum Number of Rules", font=Small_Font)
        LabMaxRules.pack()
        self.MaxRulesNum = StringVar()
        MaxRules = ttk.Entry(self, textvariable=self.MaxRulesNum)
        MaxRules.pack()
        MaxRules.delete(0, END)
        MaxRules.insert(0, 15)
        self.MaxRulesNum.set(15)

        label1 = ttk.Label(self, text="Run Result", font=Small_Font)
        label1.pack(pady=10,padx=5)
        
        self.tex = Text(self)
        self.tex.pack()
        sys.stdout = TextRedirector(self.tex, "stdout")
        sys.stderr = TextRedirector(self.tex, "stderr")

       
        
        
    def browseandrun(self):
        fname = askopenfilename(filetypes=(("Excel Files", "*.xlsx;*.xls;*.csv"),
                                           ("All files", "*.*") ))
        if fname:
            try:
                self.tex.delete("0.0",END)
                print("""Running!""")
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        self.update()
        book = xlrd.open_workbook(fname)
        sheet = book.sheet_by_index(0)
        datatrain =  [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]    
        datatestpercent = 0.1
        datatest = []
        
        alpha = sheet.ncols-2
        Kelas = 2
        datacol = [0]*sheet.ncols
        LimLow = [1]*sheet.ncols
        LimMed = [1]*sheet.ncols
        
        Tier1 = ["S1 Teknik Telekomunikasi ",
                 "S1 Teknik Informatika ",
                 "S1 Teknik Industri ",
                 "S1 International ICT Business ",
                 "S1 MBTI "]
        Tier2 = ["S1 Teknik Elektro ",
                 "S1 Sistem Komputer ",
                 "S1 Teknik Fisika ",
                 "S1 Ilmu Komputasi ",
                 "S1 Ilmu Komunikasi ",
                 "S1 Sistem Informasi ",
                 "S1 Akuntansi ",
                 "S1 Administrasi Bisnis ",
                 "S1 Desain Komunikasi Visual ",
                 "S1 Desain Interior ",
                 "D3 Teknik Informatika ",
                 "D3 Teknik Telekomunikasi "]
        Tier3 = ["S1 Kriya Tekstil dan Mode ",
                 "S1 Desain Produk ",
                 "S1 Seni Rupa Murni ",
                 "D3 Teknik Komputer ",
                 "D3 Manajemen Informatika ",
                 "D3 Komputerisasi Akuntansi ",
                 "D3 Manajemen Pemasaran ",
                 "D3 Perhotelan "]
        
        for c in range(1,sheet.ncols):
            datacol[c] = [sheet.cell_value(r,c) for r in range(1,len(datatrain))]
            if type(datacol[c][1]) == float:
                LimLow[c] = min(datacol[c])+0.333*(max(datacol[c])-min(datacol[c]))
                LimMed[c] = min(datacol[c])+0.666*(max(datacol[c])-min(datacol[c]))
            
        #print(LimLow)
        #print(LimMed)
        #print(datatrain[1][17])
        
        #if datatrain[1][17] not in Tier2:
        #    print("tidak ada")
        #else:
        #    print("ada")
            
        for r in range(1,len(datatrain)):
            for c in range(1,sheet.ncols):
                if (type(datatrain[r][c]) == float): #Diskretisasi, jika float, ubah ke string A,B,C,D,E, jika string dibiarkan
                    if (datatrain[r][c] < (LimLow[c])): 
                        datatrain[r][c] = "Rendah"
                if (type(datatrain[r][c]) == float): 
                    if ((datatrain[r][c] >= LimLow[c]) and (datatrain[r][c] < LimMed[c])): 
                        datatrain[r][c] = "Menengah"
                if (type(datatrain[r][c]) == float): 
                    if (datatrain[r][c] >= (LimMed[c])): 
                        datatrain[r][c] = "Tinggi" 
                if (datatrain[r][c] in Tier1):
                    datatrain[r][c] = "Tier1"
                if (datatrain[r][c] in Tier2):
                    datatrain[r][c] = "Tier2"
                if (datatrain[r][c] in Tier3):
                    datatrain[r][c] = "Tier3"      
        random.seed(50) #Pemberian seed untuk membuat pembangkitan angka random memiliki barisan yang tetap pada kombinasi nomor 50.
        randlist = random.sample(range(1,len(datatrain)), math.floor(datatestpercent*len(datatrain)))
        random.seed()
        #print(randlist)
        for r,index in enumerate(randlist):
            index -= r
            datatest.append(datatrain.pop(index))
        print('Jumlah datatrain :',len(datatrain)-1)
        print('Jumlah datatest :',len(datatest)) 
        self.update()
        RuleBest = []
        RuleBestQ = []
        Min_case = 10
        Antnum = int(self.TotalAntNum.get())-1
        Rulesnum = 2
        Max_rules = int(self.MaxRulesNum.get())
        #Max_iter = 10
        Countcase = 0
        Countbenar = 0
        Countsalah = 0
        iterasi = 0
        print('Variabel yang digunakan :\nJumlah Semut : ',Antnum+1,'\nBatas Konvergensi Rules : ',Rulesnum+1,'\n''Jumlah Rules yang diinginkan : ',Max_rules,'\n')
        print('Persentase Data Training dan Testing : ',(1-datatestpercent)*100,':',datatestpercent*100)
        self.update()
        while len(datatrain) > Min_case and (iterasi < Max_rules):
            #print('iterasi ke : ',iterasi)
            t=0 #index semut
            j=0 #index konvergenitas rule
            Rulelist = []
            TempRule = []
            RuleQ = []
            RulelistQ = []
            ProbEmpir = 2*(1/(3*alpha))
            H = -(ProbEmpir*math.log(ProbEmpir,2))
            eta = (math.log(Kelas,2)- H)/3*(math.log(Kelas,2)-H)
            ferrule =  [[0.3 for x in range(1,sheet.ncols+1)] for x in range(1,len(datatrain))]
            sumferomon = sum(sum(i) for i in zip(*ferrule))
            randlist = list(range(1,len(datatrain)-2))
            random.shuffle(randlist)
            pheromonearrayR = []
            pheromonearrayC = []
            
            while ((t<=Antnum) and (j<=Rulesnum)):
                TempRule =[]
                for c in range(1,sheet.ncols):
                    for r in randlist: #Baris yang dipilih random
                        try:
                            probrule = (eta*ferrule[r][c])/(3*(eta*ferrule[r][c]))
                        except ZeroDivisionError:
                            probrule = 0
                        except IndexError:
                            break
                        self.update()    
                        randnum = random.uniform(0,1)
                        if (randnum <= probrule):
                            #print(r,',',c)
                            TempRule.append(datatrain[r][c])
                            pheromonearrayR.append(r) 
                            pheromonearrayC.append(c)
                            break
                        #elif (randnum > probrule):
                            #print(r,',',c)
                            #unpheromonearrayR.append(r) 
                            #unpheromonearrayC.append(c)

                #print(t)                
                Rulelist.append(TempRule)
                #print(Rulelist[t])
                RuleQ = []
                      
                TP = 0
                FP = 0
                FN = 0
                TN = 0
                Rulelistcopy = copy.deepcopy(Rulelist)
                datatraincopy = copy.deepcopy(datatrain)
                datatemp = copy.deepcopy(datatrain)
                prunerule = True  
                for r in range(1,len(datatrain)):
                    if (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                        TP = TP + 1
                    elif (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                        FP = FP + 1
                    elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                        FN = FN+ 1
                    elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                        TN = TN + 1
                                    
                                        
                try:
                   initquality = (TP/(TP+FN))*(TN/(FP+TN))
                except ZeroDivisionError:
                   initquality = 0
                        
                #print(initquality)

                pruneiter = 0
                while prunerule:
                    RuleQ = [] #kosongkan list kualitas rule
                    for c in range(sheet.ncols-2):
                        TP = 0
                        FP = 0
                        FN = 0
                        TN = 0
                        for r in datatraincopy:
                            r[c+1] = '-'
                        Rulelistcopy[t][c] = '-'
                        self.update() 
                        for r in range(1,len(datatrain)-2):
                            if (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                                TP = TP + 1
                            elif (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                                FP = FP + 1
                            elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                                FN = FN + 1
                            elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                                TN = TN + 1
                        
                        try:
                            RuleQ.append((TP/(TP+FN))*(TN/(FP+TN)))
                        except ZeroDivisionError:
                            RuleQ.append(0)
                        #print(max(RuleQ),' ',initquality)
                        Rulelistcopy = copy.deepcopy(Rulelist)
                        if pruneiter == 0:
                            datatraincopy = copy.deepcopy(datatrain)
                        else:
                            datatraincopy = copy.deepcopy(datatemp)
                    if (max(RuleQ) > initquality):
                        #print('Prune index ke: ',RuleQ.index(max(RuleQ)))
                        Rulelist[t][RuleQ.index(max(RuleQ))] = '-'
                        Rulelistcopy[t][RuleQ.index(max(RuleQ))] = '-'
                        for r in datatraincopy:
                            r[RuleQ.index(max(RuleQ))+1] = '-'
                        datatemp = copy.deepcopy(datatraincopy)
                        initquality = max(RuleQ)
                        #print(initquality)
                        prunerule = True
                        pruneiter = pruneiter + 1
                    else:
                        prunerule = False
                        
                RuleQ = [] #kosongkan list kualitas rule
                Rulelistcopy[t] = copy.deepcopy(Rulelist[t])
                if Rulelistcopy[t][17] == "LULUS":
                    Rulelistcopy[t][17] = "TIDAK LULUS"
                    #print('Ganti Konsekuen Rule dari Lulus jadi Tidak?')
                elif Rulelist[t][17] == "TIDAK LULUS":
                    Rulelistcopy[t][17] = "LULUS"
                    #print('Ganti Konsekuen Rule dari Tidak Lulus jadi Lulus?')
                TP = 0
                FP = 0
                FN = 0
                TN = 0
                for r in range(1,len(datatrain)-2):
                    if (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                        TP = TP + 1
                    elif (Rulelistcopy[t][:16] == datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                        FP = FP + 1
                    elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] == datatraincopy[r][18]):
                        FN = FN + 1
                    elif (Rulelistcopy[t][:16] != datatraincopy[r][1:17]) and (Rulelistcopy[t][17] != datatraincopy[r][18]):
                        TN = TN + 1
                try:
                    RuleQ.append((TP/(TP+FN))*(TN/(FP+TN)))
                except ZeroDivisionError:
                    RuleQ.append(0)
                if (max(RuleQ) > initquality):
                    Rulelist[t][17] = Rulelistcopy[t][17]
                    initquality = max(RuleQ)
                    #print('Ya Ganti Konsekuen Rule')
                #else:
                    #print('Tidak Ganti Konsekuen Rule')
                    
                RulelistQ.append(initquality) 
                for r,c in zip(pheromonearrayR,pheromonearrayC):
                    try:
                        ferrule[r][c] = ferrule[r][c] + (ferrule[r][c] * initquality)
                    except IndexError:
                        break
                for r in range(1,len(datatrain)-1):
                    for c in range(1,sheet.ncols):
                        if (r not in pheromonearrayR) and (c not in pheromonearrayC):
                            sumferomon = sum(sum(i) for i in zip(*ferrule))
                            ferrule[r][c] = (ferrule[r][c])/sumferomon
                
                #print(Rulelist[t])        
                                            
                if t>0 and Rulelist[t] == Rulelist[t-1]:
                    j=j+1
                else:
                    j=0
                t+=1
                
            if Rulelist[RulelistQ.index(max(RulelistQ))] not in RuleBest:     
                RuleBest.append(Rulelist[RulelistQ.index(max(RulelistQ))])
                RuleBestQ.append(max(RulelistQ))
                datatraincopy = copy.deepcopy(datatrain)
                for c in range(sheet.ncols-2):
                    if RuleBest[iterasi][c] == '-':
                        for r in range (1,len(datatraincopy)):
                            datatraincopy[r][c+1] = '-'
                Coveredruleindex = []    
                for r in range(1,len(datatrain)):
                    if (RuleBest[iterasi][:16] == datatraincopy[r][1:17]) and (RuleBest[iterasi][17] == datatraincopy[r][18]):
                        Countbenar = Countbenar + 1
                        Coveredruleindex.append(r)
                    elif (RuleBest[iterasi][:16] == datatraincopy[r][1:17]) and (RuleBest[iterasi][17] != datatraincopy[r][18]):
                        Countsalah = Countsalah + 1

                #print(Coveredruleindex)
                for r,index in enumerate(Coveredruleindex):
                    index -= r
                    del datatrain[index]
                    
            print('Sisa Data Training Iterasi ke ',iterasi,' : ',len(datatrain))
            iterasi = iterasi + 1
            self.update()
        Countcase = (Countbenar+Countsalah)
        if Countcase > 0:
            akurasitrain = (Countbenar/Countcase)*100
        else:
            akurasitrain = 0
        print('Jumlah klasifikasi benar : ', Countbenar)
        print('Jumlah klasifikasi salah : ', Countsalah)
        print('Akurasi training : ', akurasitrain,'%')

        
        Counttestbenar = 0
        Counttestsalah = 0
        akurasirule = []
        recallrule = []
        precisionrule = []
        avermeasure = []
        for i in range(len(RuleBest)):
            a = 0
            b = 0
            c = 0
            d = 0
            print('Rule ke ',i+1,RuleBest[i],'\n')
            print('Kualitas rule ke ',i+1,':',RuleBestQ[i],'\n')
            datatestcopy = copy.deepcopy(datatest)
            for c in range(sheet.ncols-2):
                if RuleBest[i][c] == '-':
                    for r in range (1,len(datatestcopy)):
                        datatestcopy[r][c+1] = '-'
            Coveredruleindex = []
            for r in range(1,len(datatest)):
                if (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] == datatestcopy[r][18]):
                    Counttestbenar = Counttestbenar + 1
                    Coveredruleindex.append(r)
                elif (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] != datatestcopy[r][18]):
                    Counttestsalah = Counttestsalah + 1
                        
                if (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] == 'TIDAK LULUS') and (datatestcopy[r][18] == 'TIDAK LULUS'):
                    a = a+1
                if (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] == 'LULUS') and (datatestcopy[r][18] == 'TIDAK LULUS'):
                    b = b+1
                if (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] == 'TIDAK LULUS') and (datatestcopy[r][18] == 'LULUS'):
                    c = c+1
                if (RuleBest[i][:16] == datatestcopy[r][1:17]) and (RuleBest[i][17] == 'LULUS') and (datatestcopy[r][18] == 'LULUS'):
                    d = d+1
            try:        
                akurasirule.append((a+d)/(a+b+c+d))
            except ZeroDivisionError:
                akurasirule.append(0)
            try:
                recallrule.append(d/(c+d))
            except ZeroDivisionError:
                recallrule.append(0)
            try:
                precisionrule.append(d/(b+d))
            except ZeroDivisionError:
                precisionrule.append(0)
                
            for r,index in enumerate(Coveredruleindex):
                    index -= r
                    del datatest[index]
            
            print('Akurasi rule ke ',i+1,' : ',akurasirule[i],'%')           
            print('Recall rule ke ',i+1,' : ',recallrule[i],'%')
            print('Presisi rule ke ',i+1,' : ',precisionrule[i],'%')
            avermeasure.append((akurasirule[i]+recallrule[i]+precisionrule[i])/3)
        Counttestcase = Counttestbenar+Counttestsalah
        if Counttestcase > 0:
            akurasitest = (Counttestbenar/Counttestcase)*100
        else:
            akurasitest = 0
        print('Jumlah testing benar : ', Counttestbenar)
        print('Jumlah testing salah : ', Counttestsalah)
        print('Akurasi testing : ', akurasitest,'%')

        print('Akurasi rule Maksimum  : ',max(akurasirule)*100,'%')           
        print('Recall rule Maksimum : ',max(recallrule)*100,'%')
        print('Presisi rule Maksimum : ',max(precisionrule)*100,'%')
        print('Ukuran-rata Maksimum : ',max(avermeasure)*100,'%')
        print('Rule dengan Ukuran-rata-rata terbaik : ',RuleBest[avermeasure.index(max(avermeasure))])
        #return(self)        
            
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")
        


app = JAMapp()
app.geometry(XY)
app.mainloop()
