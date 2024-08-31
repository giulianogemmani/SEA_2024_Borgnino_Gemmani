# coding: utf-8
'''
Created on 2 Oct 2018

@author: SMG & AMB
'''

import os
os.environ["PATH"] = r"C:\Program Files\Digsilent\PowerFactory"+ os.environ["PATH"]
#from typing_extensions import Literal
import wx
from wx.lib import masked

import ui.busbar_selector_dlg as busbar_selector

#from multiprocessing import Process

class MainPanel(wx.Panel): # I create a panel where I can put the outher GUI elements
    def __init__(self, parent, interface, logic):
        
        self.interface = interface
        self.logic = logic
   
        
        wx.Panel.__init__(self, parent) # panel initialization

        headingStyle = wx.Font(pointSize=8, family=wx.DEFAULT, style=wx.NORMAL, weight=wx.BOLD)

        self.mainSizer = wx.BoxSizer(wx.VERTICAL) #A vertical sizer that will manage the overall layout of the panel by stacking child elements vertically.
                
        # Script specific fields
        # self.ppSizer = wx.BoxSizer(wx.VERTICAL) #Another vertical sizer that might be used to organize specific elements vertically.
        
        
        # Source File
        self.sizerResultsFile = wx.BoxSizer(wx.HORIZONTAL) #A horizontal box sizer to hold components related to file selection.
        self.PFDFileLabel = wx.StaticText(self, -1, "PFD/DZ, XLS, SLX file: ")
        self.PFDFileLabel.SetFont(headingStyle)
        self.sizerResultsFile.Add(self.PFDFileLabel, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.PFDFileName = wx.TextCtrl(self, -1, name="PFD/DZ, XLS,SLX FileName")
        self.sizerResultsFile.Add(self.PFDFileName, 1, wx.EXPAND | wx.ALL, border=4) # when I use wx.expand the selected widget strechtes for all the available space
        

        self.browsePFDButton = wx.Button(self, -1, "Browse for PFD/DZ, XLS, SLX file")
        self.sizerResultsFile.Add(self.browsePFDButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browsePFDButton.Bind(wx.EVT_BUTTON, self.OnBrowsePFDFile, self.browsePFDButton)# I bind the button to an event
        
        self.ResultsFileLabel = wx.StaticText(self, -1, "Results file: ")
        self.ResultsFileLabel.SetFont(headingStyle)
        self.sizerResultsFile.Add(self.ResultsFileLabel, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.results_file_name = wx.TextCtrl(self, -1, name="ResultsFileName")
        self.sizerResultsFile.Add(self.results_file_name, 1, wx.EXPAND | wx.ALL, border=4)

        self.browseButton = wx.Button(self, -1, "Browse")
        self.sizerResultsFile.Add(self.browseButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browseButton.Bind(wx.EVT_BUTTON, self.OnBrowseResultsFile, self.browseButton)


        self.mainSizer.Add(self.sizerResultsFile, 0, wx.EXPAND)# I add the horizontal sizer to the main vertical one
        
        
        
        # Commands to create Boxsizers that are objects to handle the layout of GUI elements
        # create an horizontal box sizer that is the main container for the layout
        self.outteroptionsGridSizer = wx.BoxSizer(wx.HORIZONTAL)
        # create three  child vertical box sizers that are contained in the horizontal one
        self.leftoptionsGridSizer = wx.BoxSizer(wx.VERTICAL) # vertical box sizers will organize their child components in a vertical manner
        self.middleoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
        self.rightoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
            

                
        # self.FaultTypesSizer = wx.BoxSizer(wx.VERTICAL)
        #
        #
        #
        # self.leftoptionsGridSizer.Add(self.FaultTypesSizer, 0, wx.ALIGN_LEFT | wx.LEFT, border=10)           
        #
        # self.FaultResistanceSizer = wx.BoxSizer(wx.HORIZONTAL)
        # self.FaultResistanceDescSizer = wx.BoxSizer(wx.VERTICAL)
        
        
        # # add an edit box which allows to enter the name of a device
        # self.LayoutSizer = wx.BoxSizer(wx.HORIZONTAL)
        # # ...here add the edit box
        # self.ppValueSizer = wx.BoxSizer(wx.VERTICAL)
        # self.ppTargetDevice = masked.TextCtrl(parent=self, value="",  size=(200, 20), name="TargetDevice")
        # self.ppTargetDevice.SetMaxLength(0)
        # self.ppValueSizer.Add(self.ppTargetDevice, 0, wx.ALIGN_CENTER | wx.ALL, border=2)   
        # #   add everything in the upper level
        # self.LayoutSizer.Add(self.ppValueSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=0)
        # self.ppSizer.Add(self.LayoutSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=0)
        # self.rightoptionsGridSizer.Add(self.ppSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=2)
        #
        # self.middleoptionsGridSizer.Add(self.FaultResistanceSizer, 0, wx.ALIGN_LEFT | wx.RIGHT, border=10) 
        
        # Middle Column: add a button to transfer values for the excel files
        self.btntransfervalues = wx.Button(self, -1, "Transfer values", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btntransfervalues, 0, wx.ALIGN_LEFT| wx.ALL, border=4)
        self.btntransfervalues.Bind(wx.EVT_BUTTON, \
                    self.transfer_set_values, self.btntransfervalues)
          
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
                  

        #self.mainSizer.SetSizeHints(self)
        self.SetSizer(self.mainSizer)        
        self.SetAutoLayout(1)
        self.mainSizer.Fit(self)
        #self.mainSizer.Layout()
        wx.Window.InitDialog(self)
        self.Show()
        
#     def GetRegions(self):
#         self.areaList.Clear()
#         areas = zip(iarray[0], carray[0])
#         for number, name in areas:
#             self.areaList.Append("{} - {}".format(number, name))
#         self.areaList.Enable()        
        
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
                    
        #top = wx.GetTopLevelParent(self)
        #settings.update(top.tab2.GetSettings())
        #settings.update(top.tab3.panel.GetSettings())
        #settings.update(top.tab4.panel.GetSettings())        
#         
        #settings.update({'VoltageRange': [settings['MinVoltageLimit'], settings['MaxVoltageLimit']]})
        #print settings['VoltageRange']
        # print settings['AreaNumbers']
        # print settings['ZoneNumbers']

        #return settings

    def OnCancel(self, event):
        '''
        tracer_logic closing event
        '''
        #MainWindow.OnExit();
#         dlg.Destroy()

        
    def OnCreateTDD(self, event):
        '''
        Run the Time Distance Diagram creation 
        '''
        # just hide the window
        self.Parent.parent.Hide()
        error = self.logic.create_TDD(self)   
        self.Parent.parent.Restore()
    
    
    def OnCreateTCD(self, event):
        '''
        Run the Time Current Diagram creation 
        '''
        # just hide the window
        self.Parent.parent.Hide()
        error = self.logic.create_TCD(self)   
        self.Parent.parent.Restore()
     
     
    def OnSetRelay(self, event):
        ''' 
        run the procedures to automatically set relays
        '''
        self.Parent.parent.Hide()
        error = self.logic.set_relays(self)   
        self.Parent.parent.Restore()
        
        
    def OnPFDB(self, event):
        ''' 
        run the procedures to save the relays settings in to the DB
        '''
        self.Parent.parent.Hide()
        error = self.logic.save_relay_settings_to_DB()   
        self.Parent.parent.Restore()  
        

    def OnDiffsheetsDB(self, event):
        ''' 
        run the procedures to save the differential relays settings from
        a sheet in to the DB
        '''
        self.Parent.parent.Hide()
        error = self.logic.save_differential_relay_settings_from_sheet_to_DB()   
        self.Parent.parent.Restore()  
        
    
    def OnRunMultipleStudyCases(self, event):
        ''' 
        run the procedures to save the relays settings in to the DB
        '''
        # Get the selected file from the file input field
        selected_file = self.PFDFileName.GetValue()
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
            error = self.logic.run_Simulink_simulation()  
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
        
        
    def calculate_line_SHCs(self, event):
        '''
        run the procedure to calculate the line SHC values
        '''
        self.Parent.parent.Hide()
        error = self.logic.calculate_line_SHCs(self)   
        self.Parent.parent.Restore()
        
    
    def calculate_shunt_SHCs(self, event):
        '''
        run the procedure to calculate the shunt SHC values
        '''
        self.Parent.parent.Hide()
        error = self.logic.calculate_shunt_SHCs(self)   
        self.Parent.parent.Restore()
        
        
    def calculate_generator_SHCs(self, event):
        '''
        run the procedure to calculate the generator SHC values
        '''
        self.Parent.parent.Hide()
        error = self.logic.calculate_generator_SHCs(self)   
        self.Parent.parent.Restore()
      
    def find_inverter_P_checking_V(self, event):
        '''
        run the procedure to find the inverter operating P using the inverter
        bus voltage
        '''
        self.Parent.parent.Hide()
        error = self.logic.find_inverter_P_checking_V(self)   
        self.Parent.parent.Restore()  
        
    def find_inverter_P_checking_tanphi(self, event):
        '''
        run the procedure to find the inverter operating P checking that 
        that the tan phi at the PoI is the one set by the station controller
        '''
        self.Parent.parent.Hide()
        error = self.logic.find_inverter_P_checking_tanphi(self)   
        self.Parent.parent.Restore()    
        
    def OnMoveTCDstoWord(self, event):
        '''
        run the procedures to move to the word report listed as output file the
        TCD diagrams available in the active study case 
        '''
        self.Parent.parent.Hide()
        error = self.logic.MoveTCDstoWord(self)   
        self.Parent.parent.Restore()  
            
        
    def run_replacer(self, event):
        '''
        run the procedure to replace values somewhere
        '''
        self.Parent.parent.Hide()
        error = self.logic.run_replacer(self)   
        self.Parent.parent.Restore()    
        
     
    def refresh_tdds(self, event):
        '''
        run the procedure to refresh all TDDs
        '''
        self.Parent.parent.Hide()
        error = self.logic.refresh_tdds()   
        self.Parent.parent.Restore() 
    
    
    def calculate_critical_times(self, event):
        '''
        calculate the generator critical time for all study cases
        '''
        self.Parent.parent.Hide()
        error = self.logic.calculate_all_critical_times(self)   
        self.Parent.parent.Restore()
     
    def calculate_optimal_lines(self, event):
        '''
        calculate the line optimal types
        '''
        self.Parent.parent.Hide()
        error = self.logic.calculate_lines_optimal_types()   
        self.Parent.parent.Restore()  
        
    def check_prc024_3_voltage(self, event):
        '''
        calculate the line optimal types
        '''
        self.Parent.parent.Hide()
        error = self.logic.check_prc024_3_voltage_settings()   
        self.Parent.parent.Restore()  
        
    def transfer_set_values(self, event):
        '''
        calculate the line optimal types
        '''
        self.Parent.parent.Hide()
        error = self.logic.transfer_set_values()   
        self.Parent.parent.Restore()         
            
    
    
    def OnBrowsePFDFile(self, event): # this method allows to browse and select a file of one of the types enlisted above
        dirname = os.getcwd()
        filename = ""
    
        # Default to current location of DYR file if present. Otherwise, use
        # location of case file if that value has been chosen.
        if self.PFDFileName.GetValue() != "":
            (dirname, __) = os.path.split(self.PFDFileName.GetValue())
    
        dlg = wx.FileDialog(self, "Choose a project file", dirname, "", "PFD/XLS/SLX file (*.pfd, *.dz, *.xls, *.slx)|*.pfd;*.dz;*.xls;*.slx", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            filename = dlg.GetFilename()
            dirname = dlg.GetDirectory()
            if len(dirname)==2 and dirname[0]>= 'A' and dirname[0]<= 'Z':  # in case of root directory i.e. C: the \\ is missing and I have to add it
                dirname += "\\" 
        else:
            dirname = ""
        dlg.Destroy()
        full_path = os.path.join(dirname, filename)
        self.PFDFileName.SetValue(full_path)
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
            
        self.fill_lists()
        
        if self.dyrFileName.GetValue() != '':
            self.btnWriteDYRFile.Enable()
                
    def OnSelectBusToStudy(self,event):
        frame = busbar_selector.BusbarSelector(self, self.interface)
        frame.Show()
        
    
