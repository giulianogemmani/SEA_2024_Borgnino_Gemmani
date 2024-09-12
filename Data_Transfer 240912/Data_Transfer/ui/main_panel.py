# coding: utf-8
'''
Created on 1 June 2024

@author: Michele Borgnino
'''

import os
os.environ["PATH"] = r"C:\Program Files\Digsilent\PowerFactory"+ os.environ["PATH"]

import wx
from wx.lib import masked

import excel_transfer.ExWoTransfer as excel_transfer
from enum import IntEnum

class OutputDetail(IntEnum):
    DISABLED = 0
    NORMAL = 1
    DEBUG = 2
    VERBOSEDEBUG = 3

class MainPanel(wx.Panel): # I create a panel where I can put the outher GUI elements
    def __init__(self, parent, interface, logic):
        
        self.output_detail = OutputDetail.NORMAL
        
        self.interface = interface
        self.logic = logic
   
        
        wx.Panel.__init__(self, parent) # panel initialization

        headingStyle = wx.Font(pointSize=8, family=wx.DEFAULT, style=wx.NORMAL, weight=wx.BOLD)

        self.mainSizer = wx.BoxSizer(wx.VERTICAL) #A vertical sizer that will manage the overall layout of the panel by stacking child elements vertically.
                
        # Script specific fields
        
        
        # Source File
        self.sizerSourceFile = wx.BoxSizer(wx.HORIZONTAL) #A horizontal box sizer to hold components related to file selection.
        self.PFDFileLabel = wx.StaticText(self, -1, "PFD/DZ, XLSX, SLX file: ")
        self.PFDFileLabel.SetFont(headingStyle)
        self.sizerSourceFile.Add(self.PFDFileLabel, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.source_file_name = wx.TextCtrl(self, -1, name="PFD/DZ, XLSX,SLX FileName", size=(500,23))
        self.sizerSourceFile.Add(self.source_file_name, 1, wx.EXPAND | wx.ALL, border=4) # when I use wx.expand the selected widget strechtes for all the available space
        

        self.browsePFDButton = wx.Button(self, -1, "Browse for PFD/DZ, XLSX, SLX file")
        self.sizerSourceFile.Add(self.browsePFDButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browsePFDButton.Bind(wx.EVT_BUTTON, self.OnBrowsePFDFile, self.browsePFDButton)# I bind the button to an event
        
        # Results File
        self.sizerResultsFile = wx.BoxSizer(wx.HORIZONTAL) #A horizontal box sizer to hold components related to file selection.
        self.ResultsFileLabel = wx.StaticText(self, -1, "Results file: ")
        self.ResultsFileLabel.SetFont(headingStyle)
        self.sizerResultsFile.Add(self.ResultsFileLabel, 0,  wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.results_file_name = wx.TextCtrl(self, -1, name="ResultsFileName", size=(400,23))
        self.sizerResultsFile.AddSpacer(55)
        self.sizerResultsFile.Add(self.results_file_name, 1, wx.EXPAND | wx.ALL, border=4)

        self.browseButton = wx.Button(self, -1, "Browse for DOCX file")
        self.sizerResultsFile.Add(self.browseButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browseButton.Bind(wx.EVT_BUTTON, self.OnBrowseResultsFile, self.browseButton)

        self.mainSizer.Add(self.sizerSourceFile, 0, wx.EXPAND)# add the horizontal sizer to the main vertical one
        self.mainSizer.Add(self.sizerResultsFile, 0, wx.EXPAND)# add the horizontal sizer to the main vertical one   
        
        # Commands to create Boxsizers that are objects to handle the layout of GUI elements
        # create an horizontal box sizer that is the main container for the layout
        self.outteroptionsGridSizer = wx.BoxSizer(wx.HORIZONTAL)
        # create three  child vertical box sizers that are contained in the horizontal one
        self.leftoptionsGridSizer = wx.BoxSizer(wx.VERTICAL) # vertical box sizers will organize their child components in a vertical manner
        self.middleoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
        self.rightoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
            
        
        # Middle Column: add a button to transfer values for the excel files
        self.btntransfervalues = wx.Button(self, -1, "Transfer values", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btntransfervalues, 0, wx.ALIGN_LEFT| wx.ALL, border=4)
        self.btntransfervalues.Bind(wx.EVT_BUTTON, \
                    self.transfer_excel_values, self.btntransfervalues)
          
        # Right Column: add the "run multiple study cases" button       
        self.btnRunMultipleStudyCases = wx.Button(self, -1, "Run All study cases", size=wx.Size(200,25))  
        self.middleoptionsGridSizer.Add(self.btnRunMultipleStudyCases, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.TOP , border=4)
        self.btnRunMultipleStudyCases.Bind(wx.EVT_BUTTON, self.OnRunMultipleStudyCases, self.btnRunMultipleStudyCases)
        
        # Right Column: add the "move all picture to word report" button       
        self.btnMoveallpicturetoWord = wx.Button(self, -1, "Move all picture to word report", size=wx.Size(200,25))  
        self.rightoptionsGridSizer.Add(self.btnMoveallpicturetoWord, 0, wx.ALIGN_RIGHT | wx.TOP , border=4)
        self.btnMoveallpicturetoWord.Bind(wx.EVT_BUTTON, self.OnMoveallpicturetoWord, self.btnMoveallpicturetoWord)
        
        # put everything together..... 
        self.outteroptionsGridSizer.Add(self.leftoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)
        self.outteroptionsGridSizer.Add(self.middleoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)            
        self.outteroptionsGridSizer.Add(self.rightoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        self.mainSizer.Add(self.outteroptionsGridSizer,0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_HORIZONTAL, border=10)      
                  

        self.SetSizer(self.mainSizer)        
        self.SetAutoLayout(1)
        self.mainSizer.Fit(self)
        #self.mainSizer.Layout()
        wx.Window.InitDialog(self)
        self.Show()
        
    def is_dialog_setting_ok(self, window, settings):
        ''' 
        function checking that the dialog settings are correct
        '''
        # Check if results file has been specified.      
        if window.results_file_name.GetValue() == "":
            dlg = wx.MessageDialog(window, "Please specify a valid results filename.",
                                           "No result filename", \
                                           wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return False
        
        # Check if source file has been specified.      
        if window.source_file_name.GetValue() == "":
            dlg = wx.MessageDialog(window, "Please specify a valid source filename.",
                                           "No source filename", \
                                           wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return False
       
        # Check if a PF project is active
        if self.interface.is_project_active() == False:
            dlg = wx.MessageDialog(
                window, "Please activate a project.", "No project is active", \
                wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return False        
    
        
        return True  
        
    def GetSettings(self):
        settings = {}        
        sizers = [self.sizerResultsFile, self.regionSizer, \
                  self.FaultTypesSizer,self.FaultResistanceValueSizer,
                  self.FaultResistanceSizer, self.StudySelectedBusMiddleSizer,\
                  self.mainSizer, self.ppValueSizer]
        
        for sizer in sizers:
            #print sizer
            children = sizer.GetChildren()

            for child in children:
                widget = child.GetWindow()
                if isinstance(widget, wx.CheckBox) or isinstance(widget, wx.TextCtrl) or isinstance(widget, wx.SpinCtrl):
                    settings[widget.GetName()] = widget.GetValue()
                    # print widget.GetName() + " ==> ",
                    # print widget.GetValue()
                elif isinstance(widget, wx.ListBox):
                    settings[widget.GetName()] = widget.GetSelections()
#                     print(settings[widget.GetName()])
                    


    def OnCancel(self, event):
        '''
        tracer_logic closing event
        '''
        #MainWindow.OnExit();
#         dlg.Destroy()


    
    def OnRunMultipleStudyCases(self, event):
        ''' 
        run the procedures to run 
            * a simulation for any available study case in PowerFActory
            * run the unique available simulation in Symulink
        '''
        # Get the selected file from the file input field
        selected_file = self.source_file_name.GetValue()
    # Ensure a file is selected
        if not selected_file:
            wx.MessageBox("Please select an input file before running study cases.", "No File Selected", wx.OK | wx.ICON_WARNING)
            return
    # Determine the file extension to decide which study cases to run
        file_extension = os.path.splitext(selected_file)[-1].lower()
        self.Parent.parent.Hide()
        if file_extension == ".pfd" or file_extension == ".dz":
            error = self.logic.run_PF_simulation_for_multiple_study_cases()  
        elif file_extension == ".slx":
            error = self.logic.run_Simulink_simulation(selected_file)  
        else:
            wx.MessageBox("Please select a PowerFactory or a Simulink input file", "No File Selected", wx.OK | wx.ICON_WARNING)  
        self.Parent.parent.Restore()  
        
    
    def OnMoveallpicturetoWord(self, event):
        ''' 
        run the procedures to move to the word report listed as output file all
        the pictures available in all study cases 
        '''
        self.Parent.parent.Hide()
        error = self.logic.move_all_pictures_to_word(self)   
        self.Parent.parent.Restore()   
         
    def transfer_excel_values(self, event):
        '''
        transfer the values from excel to word
        '''
        self.Parent.parent.Hide()
        error = excel_transfer.transfer_excel_values(self)   
        self.Parent.parent.Restore()        
        
    
    
    def OnBrowsePFDFile(self, event): # this method allows to browse and select a file of one of the types enlisted above
        dirname = os.getcwd()
        filename = ""
    
        # Default to current location of DYR file if present. Otherwise, use
        # location of case file if that value has been chosen.
        if self.source_file_name.GetValue() != "":
            (dirname, __) = os.path.split(self.source_file_name.GetValue())
    
        dlg = wx.FileDialog(self, "Choose a project file", dirname, "", "PFD/XLSX/SLX file (*.pfd, *.dz, *.xlsx, *.sls)|*.pfd;*.dz;*.xlsx;*.slx", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            filename = dlg.GetFilename()
            dirname = dlg.GetDirectory()
            if len(dirname)==2 and dirname[0]>= 'A' and dirname[0]<= 'Z':  # in case of root directory i.e. C: the \\ is missing and I have to add it
                dirname += "\\" 
        else:
            dirname = ""
        dlg.Destroy()
        full_path = os.path.join(dirname, filename)
        self.source_file_name.SetValue(full_path)
        if dirname != "":
            self.interface.import_project(full_path)
        

    def OnBrowseResultsFile(self, event):
        dirname = os.getcwd()
       
        # Default to current location of DYR file if present. Otherwise, use
        # location of case file if that value has been chosen.
        if self.results_file_name.GetValue() != "":
            (dirname, __) = os.path.split(self.results_file_name.GetValue())
    
        dlg = wx.FileDialog(self, "Choose a file", dirname, "", "XLSX file (*.xlsx)|*.xlsx | DOCX file (*.docx)|*.docx", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.results_file_name.SetValue(path)
        dlg.Destroy()
        
    
    def OnOpenOptionsFile(self, optionsFile):
        # Open the results file. Need to check for file extension to determine which call to use.
        (__, ext) = os.path.splitext(optionsFile)
        ext = ext.upper()
            
        
        if self.dyrFileName.GetValue() != '':
            self.btnWriteDYRFile.Enable()
                
 
    
