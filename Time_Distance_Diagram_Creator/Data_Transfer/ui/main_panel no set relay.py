# coding: utf-8
'''
Created on 2 Oct 2018

@author: SMG & AMB
'''

import os
import wx
from wx.lib import masked

import ui.busbar_selector_dlg as busbar_selector

#from multiprocessing import Process

class MainPanel(wx.Panel):
    def __init__(self, parent, interface, logic):
        
        self.interface = interface
        self.logic = logic
        
        self.area_list = []
        self.zone_list = []
        self.voltage_list = []
        self.path_list = []
        self.grid_list = []
        self.busbar_list = []
        
        wx.Panel.__init__(self, parent)


        headingStyle = wx.Font(pointSize=8, family=wx.DEFAULT, style=wx.NORMAL, weight=wx.BOLD)

        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        
        # Source File
        self.sizerResultsFile = wx.BoxSizer(wx.HORIZONTAL)
        self.PFDFileLabel = wx.StaticText(self, -1, "PFD/DZ file: ")
        self.PFDFileLabel.SetFont(headingStyle)
        self.sizerResultsFile.Add(self.PFDFileLabel, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.PFDFileName = wx.TextCtrl(self, -1, name="PfdFileName")
        self.sizerResultsFile.Add(self.PFDFileName, 1, wx.ALIGN_CENTER | wx.EXPAND | wx.ALL, border=4)
        

        self.browsePFDButton = wx.Button(self, -1, "Browse for PFD/DZ")
        self.sizerResultsFile.Add(self.browsePFDButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browsePFDButton.Bind(wx.EVT_BUTTON, self.OnBrowsePFDFile, self.browsePFDButton)
        
        self.ResultsFileLabel = wx.StaticText(self, -1, "Results file: ")
        self.ResultsFileLabel.SetFont(headingStyle)
        self.sizerResultsFile.Add(self.ResultsFileLabel, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.results_file_name = wx.TextCtrl(self, -1, name="ResultsFileName")
        self.sizerResultsFile.Add(self.results_file_name, 1, wx.ALIGN_CENTER | wx.EXPAND | wx.ALL, border=4)

        self.browseButton = wx.Button(self, -1, "Browse")
        self.sizerResultsFile.Add(self.browseButton, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.browseButton.Bind(wx.EVT_BUTTON, self.OnBrowseResultsFile, self.browseButton)


        self.mainSizer.Add(self.sizerResultsFile, 0, wx.EXPAND)
        
        self.selectRegionsLabel = wx.StaticText(self, -1, "Select region to study:")
        self.selectRegionsLabel.SetFont(headingStyle)
        self.mainSizer.Add(self.selectRegionsLabel, 1, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.regionSizer = wx.BoxSizer(wx.HORIZONTAL)
       
        self.regionSizer.Add(wx.StaticText(self, -1, "Voltage Levels: "), 0, wx.ALIGN_TOP | wx.ALL, border=2)
        self.voltageList = wx.ListBox(self, size=(150, 75), choices=[], style=wx.LB_EXTENDED, name="VoltageList")
        self.voltageList.Disable()
        self.regionSizer.Add(self.voltageList, 0, wx.ALIGN_CENTER | wx.LEFT, border=8)

        self.regionSizer.Add(wx.StaticText(self, -1, "Areas: "), 0, wx.ALIGN_TOP | wx.ALL, border=2)        
        self.areaList = wx.ListBox(self, size=(150, 75), choices=[], style=wx.LB_EXTENDED, name="AreaList")
        self.areaList.Disable()     
        self.regionSizer.Add(self.areaList, 0, wx.ALIGN_CENTER | wx.LEFT, border=8)
        
        self.regionSizer.Add(wx.StaticText(self, -1, "Zones: "), 0, wx.ALIGN_TOP | wx.ALL, border=4)    
        self.zoneList = wx.ListBox(self, size=(150, 75), style=wx.LB_EXTENDED, name="ZoneList")
        self.zoneList.Disable()
        self.regionSizer.Add(self.zoneList, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        
        self.regionSizer.Add(wx.StaticText(self, -1, "Grids: "), 0, wx.ALIGN_TOP | wx.ALL, border=4)    
        self.gridList = wx.ListBox(self, size=(150, 75), style=wx.LB_EXTENDED, name="GridList")
        self.gridList.Disable()
        self.regionSizer.Add(self.gridList, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
        self.regionSizer.Add(wx.StaticText(self, -1, "Paths: "), 0, wx.ALIGN_TOP | wx.ALL, border=4)    
        self.pathList = wx.ListBox(self, size=(150, 75), style=wx.LB_EXTENDED, name="PathList")
        self.pathList.Disable()
        self.regionSizer.Add(self.pathList, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        
                
        self.mainSizer.Add(self.regionSizer, 0, wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        self.StudySelectedBusSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.StudySelectedBusLeftSizer = wx.BoxSizer(wx.VERTICAL)
        self.StudySelectedBusMiddleSizer = wx.BoxSizer(wx.VERTICAL)
        self.StudySelectedBusRightSizer = wx.BoxSizer(wx.VERTICAL)        
        self.StudySelectedBusLeftSizer.Add(wx.StaticText(self, -1, "Study Around This Busbar :"), 0, wx.ALIGN_LEFT | wx.ALL, border=6)
        self.StudySelectedBusLeftSizer.Add(wx.StaticText(self, -1, "Number of lines away from selected busbar to study:"), 0, wx.ALIGN_LEFT | wx.ALL, border=6)        
        self.StudySelectedBus = wx.TextCtrl(parent=self, value="",  size=(150, 20),  name="StudySelectedBus")
        self.StudySelectedBusMiddleSizer.Add(self.StudySelectedBus, 0, wx.ALIGN_CENTER | wx.ALL, border=4)        
        self.StudySelectedBusExtent = masked.NumCtrl(self, integerWidth=1,value="0", size=(50, 20), autoSize=False, fractionWidth=0, min=0, max=5, name="StudySelectedBusExtent")
        self.StudySelectedBusMiddleSizer.Add(self.StudySelectedBusExtent, 0, wx.ALIGN_CENTER | wx.ALL, border=4)        

        self.btnSelectBusToStudy = wx.Button(self, -1, "Select Busbar To Study") #, size=wx.Size(200,25))        
        self.StudySelectedBusRightSizer.Add(self.btnSelectBusToStudy, 0, wx.ALIGN_CENTER | wx.ALL, border=4)
        self.btnSelectBusToStudy.Bind(wx.EVT_BUTTON, self.OnSelectBusToStudy, self.btnSelectBusToStudy)        

        self.StudySelectedBusSizer.Add(self.StudySelectedBusLeftSizer, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.StudySelectedBusSizer.Add(self.StudySelectedBusMiddleSizer, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.StudySelectedBusSizer.Add(self.StudySelectedBusRightSizer, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.mainSizer.Add(self.StudySelectedBusSizer, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        # Main Options Grid
        self.outteroptionsGridSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.leftoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
        self.middleoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
        self.rightoptionsGridSizer = wx.BoxSizer(wx.VERTICAL)
            

        self.optionsLabel = wx.StaticText(self, -1, "Faults to Study")
        self.optionsLabel.SetFont(headingStyle)
        self.leftoptionsGridSizer.Add(self.optionsLabel, 1, wx.ALIGN_LEFT | wx.LEFT, border=4)
        
        self.FaultTypesSizer = wx.BoxSizer(wx.VERTICAL)
         
        self.Fslgr = wx.CheckBox(self, label="Single Line to Ground with fault resistance", name="Fslgr")
        self.Fslgr.SetValue(False)
        self.FaultTypesSizer.Add(self.Fslgr, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.Fltlr = wx.CheckBox(self, label="Line to Line with fault resistance", name="Fltlr")
        self.Fltlr.SetValue(False)
        self.FaultTypesSizer.Add(self.Fltlr, 0, wx.ALIGN_LEFT | wx.ALL, border=4)

        self.Fdlgr = wx.CheckBox(self, label="Double Line to Ground with fault resistance", name="Fdlgr")
        self.Fdlgr.SetValue(False)
        self.FaultTypesSizer.Add(self.Fdlgr, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.Fslg = wx.CheckBox(self, label="Single Line to Ground", name="Fslg")
        self.Fslg.SetValue(True)
        self.FaultTypesSizer.Add(self.Fslg, 0, wx.ALIGN_LEFT | wx.ALL, border=4)

        self.Fltl = wx.CheckBox(self, label="Line to Line", name="Fltl")
        self.Fltl.SetValue(False)
        self.FaultTypesSizer.Add(self.Fltl, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.Fdlg = wx.CheckBox(self, label="Double Line to Ground", name="Fdlg")
        self.Fdlg.SetValue(False)
        self.FaultTypesSizer.Add(self.Fdlg, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        
        self.Ftph = wx.CheckBox(self, label="Three phase", name="Ftph")
        self.Ftph.SetValue(False)
        self.FaultTypesSizer.Add(self.Ftph, 0, wx.ALIGN_LEFT | wx.ALL, border=4)

        self.leftoptionsGridSizer.Add(self.FaultTypesSizer, 0, wx.ALIGN_LEFT | wx.LEFT, border=10)           

        self.FaultResistanceSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.FaultResistanceDescSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.optionsLabel = wx.StaticText(self, -1, "")
        self.optionsLabel.SetFont(headingStyle)
        self.FaultResistanceDescSizer.Add(self.optionsLabel, 1, wx.ALIGN_LEFT | wx.LEFT, border=4)
           
        self.FaultResistanceDescSizer.Add(wx.StaticText(self, -1, "Single Line to Ground Fault Resistance:"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceDescSizer.Add(wx.StaticText(self, -1, "Line to Line Fault Resistance:"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceDescSizer.Add(wx.StaticText(self, -1, "Double Line to Ground Fault Resistance:"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceSizer.Add(self.FaultResistanceDescSizer)
        
        self.FaultResistanceValueSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.optionsLabel = wx.StaticText(self, -1, "")
        self.optionsLabel.SetFont(headingStyle)
        self.FaultResistanceValueSizer.Add(self.optionsLabel, 1, wx.ALIGN_LEFT | wx.LEFT, border=4)
        
        self.FslgrValue = masked.NumCtrl(self, value="0.0", size=(50, 20), autoSize=False, fractionWidth=1, min=0, max=5000, name="FslgrValue")
        self.FaultResistanceValueSizer.Add(self.FslgrValue, 0, wx.ALIGN_CENTER | wx.ALL, border=2)
        self.FltlrValue = masked.NumCtrl(self, value="0.0", size=(50, 20), autoSize=False, fractionWidth=1, min=0, max=5000,  name="FltlrValue")
        self.FaultResistanceValueSizer.Add(self.FltlrValue, 0, wx.ALIGN_CENTER | wx.ALL, border=2)
        self.FdlgrValue = masked.NumCtrl(self, value="0.0", size=(50, 20), autoSize=False, fractionWidth=1, min=0, max=5000, name="FdlgrValue")
        self.FaultResistanceValueSizer.Add(self.FdlgrValue, 0, wx.ALIGN_CENTER | wx.ALL, border=2)
        self.FaultResistanceSizer.Add(self.FaultResistanceValueSizer)
        
        self.FaultResistanceUnitsSizer = wx.BoxSizer(wx.VERTICAL)
        self.optionsLabel = wx.StaticText(self, -1, "")
        self.optionsLabel.SetFont(headingStyle)
        self.FaultResistanceUnitsSizer.Add(self.optionsLabel, 1, wx.ALIGN_LEFT | wx.LEFT, border=4)
        self.FaultResistanceUnitsSizer.Add(wx.StaticText(self, -1, "ohm"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceUnitsSizer.Add(wx.StaticText(self, -1, "ohm"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceUnitsSizer.Add(wx.StaticText(self, -1, "ohm"), 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.FaultResistanceSizer.Add(self.FaultResistanceUnitsSizer)
        
        
        # Run TDD creation  Button
        self.btnRunStudy = wx.Button(self, -1, "Create TDD", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btnRunStudy, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btnRunStudy.Bind(wx.EVT_BUTTON, self.OnCreateTDD, self.btnRunStudy)
        
        self.middleoptionsGridSizer.Add(self.FaultResistanceSizer, 0, wx.ALIGN_LEFT | wx.RIGHT, border=10)
        
        # Run TCD creation  Button
        self.optionsLabel2 = wx.StaticText(self, -1, "")
        self.middleoptionsGridSizer.Add(self.optionsLabel2, 0, wx.ALIGN_CENTER | wx.ALL, border=36)  
        
        self.btnCreateTOD = wx.Button(self, -1, "Create TCD", size=wx.Size(200,25))        
        self.middleoptionsGridSizer.Add(self.btnCreateTOD, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btnCreateTOD.Bind(wx.EVT_BUTTON, self.OnCreateTCD, self.btnCreateTOD)

        # Right Column: add the "PF => DB" button       
        self.btnPFDB = wx.Button(self, -1, "PF => DB", size=wx.Size(200,25))  
        self.rightoptionsGridSizer.Add(self.btnPFDB, 0,  wx.ALIGN_RIGHT | wx.TOP , border=180)
        self.btnPFDB.Bind(wx.EVT_BUTTON, self.OnPFDB, self.btnPFDB)
        
        # Right Column: add the "differential sheets => DB" button       
        self.btnDiffsheetsDB = wx.Button(self, -1, "Differential sheets => DB", size=wx.Size(200,25))  
        self.rightoptionsGridSizer.Add(self.btnDiffsheetsDB, 0,  wx.ALIGN_RIGHT | wx.TOP , border=10)
        self.btnDiffsheetsDB.Bind(wx.EVT_BUTTON, self.OnDiffsheetsDB, self.btnDiffsheetsDB)
        
        # Right Column: add the "Set Relay" button       
        self.btnCreateSetRelay = wx.Button(self, -1, "Set Relays", size=wx.Size(200,25))  
        #self.rightoptionsGridSizer.Add(self.btnCreateSetRelay, 0, wx.ALIGN_RIGHT | wx.TOP , border=90)
        #self.btnCreateSetRelay.Bind(wx.EVT_BUTTON, self.OnSetRelay, self.btnCreateSetRelay)
        
        # add a button to calculate the line SHC values 
        self.btnlineshc = wx.Button(self, -1, "Calculate line shc", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btnlineshc, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btnlineshc.Bind(wx.EVT_BUTTON, self.calculate_line_SHCs, self.btnlineshc)
        
        # Middle Column: add a button to refresh the TDD diagrams 
        self.btntddrefresh = wx.Button(self, -1, "Refresh TDD", size=wx.Size(200,25))        
        self.middleoptionsGridSizer.Add(self.btntddrefresh, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btntddrefresh.Bind(wx.EVT_BUTTON, self.refresh_tdds, self.btntddrefresh)       
        
        # add a button to calculate the shunt SHC values 
        self.btnshuntshc = wx.Button(self, -1, "Calculate shunt shc", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btnshuntshc, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btnshuntshc.Bind(wx.EVT_BUTTON, self.calculate_shunt_SHCs, self.btnshuntshc)
        
        # add a button to calculate the generator SHC values 
        self.btngeneratorshc = wx.Button(self, -1, "Calculate generator shc", size=wx.Size(200,25))        
        self.leftoptionsGridSizer.Add(self.btngeneratorshc, 0, wx.ALIGN_LEFT | wx.ALL, border=4)
        self.btngeneratorshc.Bind(wx.EVT_BUTTON, self.calculate_generator_SHCs, self.btngeneratorshc)       
        
        # put everything together..... 
        self.outteroptionsGridSizer.Add(self.leftoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)
        self.outteroptionsGridSizer.Add(self.middleoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)            
        self.outteroptionsGridSizer.Add(self.rightoptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)
        
        self.mainSizer.Add(self.outteroptionsGridSizer,0, wx.ALIGN_LEFT | wx.LEFT, border=10)      
                  

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
        sizers = [self.sizerResultsFile, self.regionSizer, self.FaultTypesSizer,self.FaultResistanceValueSizer,
            self.FaultResistanceSizer, self.StudySelectedBusMiddleSizer,self.mainSizer]
        
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

        return settings

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
        
     
    def refresh_tdds(self, event):
        '''
        run the procedure to refresh all TDDs
        '''
        self.Parent.parent.Hide()
        error = self.logic.refresh_tdds()   
        self.Parent.parent.Restore() 
           
            
    def refresh_listboxes(self):
        self.fill_lists()
    
    
    def OnBrowsePFDFile(self, event):
        dirname = os.getcwd()
        filename = ""
    
        # Default to current location of DYR file if present. Otherwise, use
        # location of case file if that value has been chosen.
        if self.PFDFileName.GetValue() != "":
            (dirname, __) = os.path.split(self.PFDFileName.GetValue())
    
        dlg = wx.FileDialog(self, "Choose a file", dirname, "", "PFD file (*.pfd, *.dz)|*.pfd;*.dz", style=wx.DD_DEFAULT_STYLE)
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
        self.refresh_listboxes()

    def OnBrowseResultsFile(self, event):
        dirname = os.getcwd()
       
        # Default to current location of DYR file if present. Otherwise, use
        # location of case file if that value has been chosen.
        if self.results_file_name.GetValue() != "":
            (dirname, __) = os.path.split(self.results_file_name.GetValue())
    
        dlg = wx.FileDialog(self, "Choose a file", dirname, "", "XLSX file (*.xlsx)|*.xlsx", style=wx.DD_DEFAULT_STYLE)
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
        
    def fill_lists(self):
        '''
        Function filling the area and the zone lists
        '''
        self.fill_area_list()
        self.fill_zone_list()
        self.fill_voltage_list()
        self.fill_grid_list()
        self.fill_path_list()
        
        
    def fill_area_list(self):
        '''
        Function filling the content of the "area" list
        '''
        self.area_list = self.interface.get_available_areas()
        
        self.areaList.Clear()
        if self.area_list is not None:
            for area in self.area_list:
                self.areaList.Append(self.interface.get_name_of(area))   
        self.areaList.Enable()
    
    def fill_zone_list(self):
        '''
        Function filling the content of the "zone" list
        '''
        self.zone_list = self.interface.get_available_zones()
        self.zoneList.Clear()
        if self.zone_list is not None:
            for zone in self.zone_list:
                self.zoneList.Append(self.interface.get_name_of(zone))
        self.zoneList.Enable()
    
    def fill_voltage_list(self):
        '''
        Function filling the content of the "voltage" list
        '''
        self.voltage_list = self.interface.get_available_voltages()
        self.voltageList.Clear()
        if self.voltage_list is not None:
            for voltage in self.voltage_list:
                self.voltageList.Append(str("{0:.2f}".format(voltage)))
        self.voltageList.Enable()
        
    def fill_grid_list(self):
        '''
        Function filling the content of the "grid" list
        '''
        self.grid_list = self.interface.get_available_grids()
        self.gridList.Clear()
        if self.grid_list is not None:
            for grid in self.grid_list:
                self.gridList.Append(self.interface.get_name_of(grid))
        self.gridList.Enable()
        
    def fill_path_list(self):
        '''
        Function filling the content of the "path" list
        '''
        self.path_list = self.interface.get_available_paths()
        self.pathList.Clear()
        if self.path_list is not None:
            for path in self.path_list:
                self.pathList.Append(self.interface.get_name_of(path))
        self.pathList.Enable()
    
    
