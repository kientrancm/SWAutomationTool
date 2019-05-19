'''
#Automation tool for software test team
#HCCV - Hella Vietnam
#Author: Kien Tran
#email: kien.tran@hella.com
#---------------------------------------
#Author         Date            Version     Description
#Kien Tran      May 17, 2019    1.0         Init tool
#---------------------------------------

'''

import os

import tkinter
from tkinter import *
from tkinter import filedialog

from HTMLtoCSV.HTMLtoTestSpec import ParseTable
from SABHandler.SABHandler import cvt_xls_to_xlsx



'''
#GUI for SWAutomationTool
'''
class GUI(tkinter.Frame):
    def __init__(self, root):
        tkinter.Frame.__init__(self, root)
        self.funtion = 0
        self.root = root
        self.initMenu()
        self.initGUI()

    def initMenu(self):
        self.root.title("Software Test Automation Tool - version 1.0")
        self.pack(fill=BOTH, expand=1)

        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        #File menu
        self.generate_excel_script = BooleanVar()
        self.generate_html_spec = BooleanVar()
        self.generate_sab_spec = BooleanVar()
        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_checkbutton(label="Generate TestScript form Excel",
                                 onvalue = True, offvalue=False,
                                 variable = self.generate_excel_script,
                                 command = self._excel)
        fileMenu.add_checkbutton(label="Generate HTML to TestSpec",
                                 onvalue = True, offvalue=False,
                                 variable = self.generate_html_spec,
                                 command = self._html)
        fileMenu.add_checkbutton(label="Generate SAB to TestSpec",
                                 onvalue = True, offvalue=False,
                                 variable = self.generate_sab_spec,
                                 command = self._sab)
        fileMenu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=fileMenu)

        #Project Menu
        self.GCAPE = BooleanVar()
        self.BCM = BooleanVar()

        PrjMenu = Menu(menubar, tearoff=0)
        PrjMenu.add_checkbutton(label="GCAPE", onvalue=True, offvalue=False, variable = self.GCAPE, command=self._gcape)
        PrjMenu.add_checkbutton(label="BCM", onvalue=True, offvalue=False, variable = self.BCM, command=self._bcm)
        menubar.add_cascade(label="Project", menu=PrjMenu)

        #Help menu
        helpMenu = Menu(menubar, tearoff=0)
        helpMenu.add_command(label="Help")
        helpMenu.add_command(label="About")
        menubar.add_cascade(label="Help", menu=helpMenu)
    '''
    #Define function for select project and select a function will be run in this GUI
    #we have 3 function for run
    #and we have 2 project
    '''
    def _excel(self):
        self.generate_html_spec.set(False)
        self.generate_sab_spec.set(False)

    def _html(self):
        self.generate_excel_script.set(False)
        self.generate_sab_spec.set(False)

    def _sab(self):
        self.generate_excel_script.set(False)
        self.generate_html_spec.set(False)

    def _bcm(self):
        self.GCAPE.set(False)

    def _gcape(self):
        self.BCM.set(False)


    def initGUI(self):

        self.file_paths = []

        #Giao dien
        self.inputFrame = LabelFrame(self)
        self.optionFrame = LabelFrame(self)
        self.LeftBotLabelFrame = LabelFrame(self)
        self.RightBotLabelFrame = LabelFrame(self)

        self.inputFrame.grid(row=1, column=1, sticky = W)
        self.optionFrame.grid(row=1, column=2, sticky = W)
        self.LeftBotLabelFrame.grid(row=2, column=1, sticky = W)
        self.RightBotLabelFrame.grid(row=2, column=2, sticky = W)

        self.PathValue = tkinter.StringVar()
        self.statusValue = StringVar()
        self.statusValue.set('Please select your option')
        self.NameVar = tkinter.StringVar()

        self.HTMLtoCSV = tkinter.IntVar()
        self.SABHandler = tkinter.IntVar()

        #input file
        self.InputGroupLabel = Label(self.inputFrame, text = "-----> Input files <-----", width = 60).grid(row = 1, column = 1, columnspan = 2)

        self.ChooserButton = Button(self.inputFrame, text='Inputs', command=self.OpenFile).grid(row=2, column=1)
        self.PathEntry = Entry(self.inputFrame, width=60, bd=2, textvariable=self.PathValue).grid(row=2, column=2)
        self.NameLabel = Label(self.inputFrame, text = 'Name').grid(row = 3, column = 1)
        self.NameEntry = Entry(self.inputFrame, width = 60, bd = 2, textvariable = self.NameVar).grid(row = 3, column = 2)

        #run
        self.RunGroupLabel = Label(self.LeftBotLabelFrame, text='-----> Run Automation Scripts <-----', width = 60).grid(row=1, column=1, columnspan = 2)
        self.generateButton = Button(self.LeftBotLabelFrame, text='Generate', command=self.MyGUI).grid(row=2, column=1, columnspan=2)
        self.statusLabel = Label(self.LeftBotLabelFrame, textvariable=self.statusValue).grid(row=3, column=1, columnspan=2)


    def OpenFile(self):
        file_paths = tkinter.filedialog.askopenfilenames(filetype = (("HTMLtoCSV", "*.html;*.htm"),("SABHandler", "*.xls;*.xlsx"), ("All files", "*.*")), parent=self,)
        self.file_paths = file_paths
        self.PathValue.set(self.file_paths)

    def setStatus(self, status):
        self.statusValue.set(status)

    def MyGUI(self):
        oktorun = False
        HTMLlabel = 0
        SABlabel = 0
        self.setStatus('Running...')
        if not str(self.PathValue.get()):
            self.setStatus('Missing inputs')
        elif not str(self.NameVar.get()):
            self.setStatus('Missing name')
        elif (self.HTMLtoCSV.get() == 0 and (self.SABHandler.get() == 0)):
            self.setStatus('Missing function')
        elif (self.HTMLtoCSV.get() == 1 and (self.SABHandler.get() == 1)):
            self.setStatus('Only one function, plz select again!')
        elif (self.HTMLtoCSV.get() == 1):
            if (self.GCAPE.get() == 0 and self.BCM.get() == 0):
                self.setStatus('Missing project')
                oktorun = False
            elif (self.GCAPE.get() == 1 and self.BCM.get() == 1):
                self.setStatus('Only one project, plz select again!')
                oktorun = False
            else:
                oktorun = True
                HTMLlabel = 1
        elif (self.SABHandler.get() == 1):
            oktorun = True
            SABlabel = 1
        else:
            oktorun = False

        if oktorun == True:
            self.setStatus("Running")
            files = self.file_paths
            if HTMLlabel == 1:
                if self.GCAPE.get() == 1:
                    project = "Auto"
                else:
                    project = "Fully"

                for file in range(0, len(files)):
                    html = files[file]
                    name = os.path.splitext(html)[0]
                    testverdict = self.NameVar.get()
                    ParseTable(html, name + "_TestSpec", testverdict, project)
            elif SABlabel == 1:
                print("Run SAB handler")

            self.setStatus('Finished')




def runGUI():
    root = tkinter.Tk()
    rGUI = GUI(root)
    rGUI.pack()
    root.mainloop()

if __name__ == "__main__":
    runGUI()