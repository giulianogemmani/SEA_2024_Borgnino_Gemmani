'''
Created on 1 June 2024

@author: MB
'''


import sys
import os

import winreg
import traceback


from enum import Enum
from enum import IntEnum
from collections import namedtuple
from datetime import datetime

from itertools import repeat
from math import sqrt
 

class PowerFactoryInterface():
    '''
    Interface between The python app and PowerFactory
    '''   
#=========================================================================
# Initialization method
#=========================================================================

    def __init__(self):
        '''
        Constructor 
        '''
        #self.default_shc_slot_name = ['m:Ikss:bus1','m:Ikss:bus2']
        self.default_pf_attribute_name = ['bus1','bus2']
        
        #self.shc_trace = None
        
        self.last_active_study_case = None
        
        self.output_file = None # the file where the print is redirected for testing

    def create(self, username, powerfactory_path):
        '''
        Method which binds PowerFactory (and run it if it isn't running)
        '''
        #powerfactorypath = os.path.dirname(r"D:\\Materiale Lavoro DIgSILENT\\PF 2018\\build\\Win32\\pf\\python\\3.6")
        #powerfactorypath = os.path.dirname(r"C:\\Program Files\\DIgSILENT\\PowerFactory 2017 SP2\\")
        powerfactorypath = os.path.join(powerfactory_path if len(
         powerfactory_path) > 0 else self.get_pf_installation_dir(), 'python\\3.6')
        sys.path.append(powerfactorypath)
        os.environ['PATH'] += powerfactorypath

        import powerfactory as pf
        self.pf = pf
        self.app = self.pf.GetApplication()
        self.app.Show()
        
        
    def refresh_pf(self):
        '''
        function closing and reopening PF and triggering a full rebuild of all
        graphical objects in pf
        '''
        self.app.Hide()
        self.app.Show()
        self.app.Rebuild(2)
        
        
    def rebuild_pf(self):
        '''
        function triggering a full rebuild of all graphical objects in pf
        '''
        self.app.Rebuild(2)


    def get_pf_installation_dir(self):
        '''
        function getting the PowerFactory installation directory
        '''
        pf_name = 'SOFTWARE\\WOW6432Node\\DIgSILENT GmbH'
        pf_path = ''

        try:
            h_key = winreg.CreateKey(
                winreg.HKEY_LOCAL_MACHINE, pf_name)  # pf_name
            pf_version = []
            i = 0
            while True:                              # get the latest PF version
                try:
                    pf_version.append(winreg.EnumKey(h_key, i))
                    h_subkey = winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE, pf_name + '\\' + pf_version[-1])
                    pf_path = (winreg.EnumValue(h_subkey, 0))[1]
                    i += 1
                    #print("PowerFactory %s is installed in: %s" % (pf_version, pf_path))
                except OSError:
                    if len(pf_version) == 0:                    # no entry found
                        print("PowerFactory insn't correctly installed!! ")
                        self.print(traceback.format_exc())
                    break
        except (PermissionError, WindowsError):
            print("PowerFactory not found!! ")
            self.print(traceback.format_exc())
        return pf_path

#=========================================================================
# Get Methods
#=========================================================================

    def get_name_of(self, element):
        '''
        function returning the element loc_name
        '''
        try:
            return str(element.loc_name)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return ""


    def get_full_name_of(self, element):
        '''
        function returning the whole element path in the database and name 
        including the class type
        '''
        try:
            return element.GetFullName()
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return ""




    def get_attribute(self, element, attribute_name):
        '''
        Function returning the value of the attribute "attribute_name" beloging
         to the element passed as first parameter
        '''
        try:
            return element.GetAttribute(attribute_name)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Problem in the " + \
                       self.get_full_name_of(element) + " element")
            self.print(traceback.format_exc())
            return 0



    def get_study_cases(self):
        '''
        function returning a list of all available study cases objects
        '''
        try:
            active_project = self.app.GetActiveProject()
            return active_project.GetContents('Study Cases.*')[0].GetContents\
                ('*.IntCase') if active_project else []             
        except Exception as e:
            self.print("Get study cases failed!: " + str(e))
            self.print(traceback.format_exc())  
     
    def get_active_study_case(self):
        '''
        function returning the active study case objects
        '''
        try:
            active_study_case = self.app.GetActiveStudyCase()
            return active_study_case             
        except Exception as e:
            self.print("Get study cases failed!: " + str(e))
            self.print(traceback.format_exc())   
            
            
    def get_simulation_inits_of(self, study_case):
        '''
        function getting the simulation initialization objects of the given 
        study case. It returns a list
        '''   
        try:
            return study_case.GetContents('*.ComInc') if study_case else []             
        except Exception as e:
            self.print("Get simulation init failed!: " + str(e))
            self.print(traceback.format_exc()) 
    
            
    def get_simulation_objects_of(self, study_case):
        '''
        function getting the simulation  objects of the given 
        study case. It returns a list
        '''   
        try:
            return study_case.GetContents('*.ComSim') if study_case else []             
        except Exception as e:
            self.print("Get simulation init failed!: " + str(e))
            self.print(traceback.format_exc()) 
            
    
    def get_study_case_events(self, study_case, event_name = ''):
        '''
        function getting the events of the given study case
        an event_name can be provided to narrow the returned list otherwise
        all shc events are returned
        '''
        try:
            event_name = event_name + '.*'  if event_name else '*.EvtShc'
            return study_case.GetContents(event_name, 1)\
                 if study_case else []             
        except Exception as e:
            self.print("Get simulation events failed!: " + str(e))
            self.print(traceback.format_exc())
            
    
    def get_study_case_results(self, study_case, result_name = ''):
        '''
        function getting the results of the given study case
        a result_name can be provided to narrow the returned list
        '''
        try:
            result_name = result_name + '.ElmRes' if result_name else '*.ElmRes'
            return study_case.GetContents(result_name, 1)\
                 if study_case else []             
        except Exception as e:
            self.print("Get simulation results failed!: " + str(e))
            self.print(traceback.format_exc())
            
            
    def get_element_variable_results(self, results, variable_name):    
        '''
        function getting in a list all results coming from a simulation of the 
        given variable_name for the given "results" result  object
        '''
        try:
            #Load the result file
            results.Load()
            # get the number of rows of the result file
            number_of_data_rows = results.GetNumberOfRows()
            # the column number of the given variable_name
            column_number = results.FindColumn(variable_name)
            return_list = [results.GetValue(i, column_number)[1]\
                     for i in range(number_of_data_rows)]
            # release the result file
            results.Release()
            return return_list
        except Exception as e:
            self.print("Get element variable results failed!: " + str(e))
            self.print(traceback.format_exc())
            
     
    def get_max_of_element_variable_results(self, results, variable_name):    
        '''
        function getting the max value of the  results coming from a simulation of the 
        given variable_name belonging to the given element for the given
         "results" result  object
        '''
        try:
            #Load the result file
            results.Load()
            # the column number of the given variable_name
            column_number = results.FindColumn(variable_name)
            max_value = results.FindMaxInColumn(column_number)[1]
            # release the result file
            results.Release()
            return max_value
        except Exception as e:
            self.print("Get max element variable results failed!: " + str(e))
            self.print(traceback.format_exc()) 
     
        
#=========================================================================
# Is Methods
#=========================================================================

    def is_project_active(self):
        '''
        function checking if there a project active in PF
        '''
        return self.app.GetActiveProject() != None


   
            
    
#=========================================================================
# Set Methods
#=========================================================================


    def set_echo_on(self):
        '''
        Enable the PowerFactory output messaqes
        '''
        self.app.EchoOn()


    def set_echo_off(self):
        '''
        Disable the PowerFactory output messaqes
        '''
        self.app.EchoOff() 


#=========================================================================
# Commands
#=========================================================================

    def import_project(self, full_project_name):
        '''
        function loading a project from a .dz or a .pfd file
        '''
        # get just the project name + .dz or .pfd
        project_name = full_project_name.split('\\')[-1]
        project_name = project_name.replace('.pfd', '')
        project_name = project_name.replace('.dz', '')
        # try to find if a project with the same name is available
        project_list = self.app.GetCurrentUser().GetContents(project_name)
        if len(project_list) > 0:
            other_project = project_list[0]
        else:
            other_project = None
        # if it exists with the same name I rename it
        if other_project != None:
            other_project.loc_name = other_project.loc_name + '_old'
        
        active_project = self.app.GetActiveProject()
        if active_project != None:
            active_project.Deactivate()
        if ".dz" in full_project_name:
            self.app.ExecuteCmd('Rd iopt_def=1 iopt_rd=dz f=' + full_project_name)
        else:
            self.app.ExecuteCmd('Rd iopt_def=1 iopt_rd=pdf f=' + full_project_name)
        
        self.app.ActivateProject(project_name)


    def export_project(self, full_project_name):
        '''
        function exporting the actual active project in the given pfd file
        '''
        active_project = self.app.GetActiveProject()
        if active_project != None:
            active_project.Deactivate()
            try:
                script = self.app.GetCurrentScript()
                exportObj = script.CreateObject('CompfdExport','Export')
                exportObj.g_objects = active_project
                exportObj.SetAttribute("e:g_file", full_project_name)
                
                exportObj.Execute()
                active_project.Activate()
            except Exception as e:
                self.print("Project export failed: " + str(e))
                self.print(traceback.format_exc())
                return 1
    
    def activate_project(self, project):
        '''
        activate the given project
        '''
        try:
            if (type(project)) is str:
                self.app.ActivateProject(project)
            else:
                project.Activate()
        except Exception as e:
            self.print("Project open failed: " + str(e))
            self.print(traceback.format_exc())
            return 1   
     
    
    def deactivate_project(self):
        '''
        deactivate the active project and return it
        '''
        try:
            active_project = self.app.GetActiveProject()
            if active_project != None:
                active_project.Deactivate();
                return active_project
        except Exception as e:
            self.print("Project open failed: " + str(e))
            self.print(traceback.format_exc())
            return None;    
    
    def get_variation(self, variation_name):
        '''
        function getting the "variation" object (IntScheme) and the relevant
        Expansion Stage 
        '''
        try:
            active_project = self.app.GetActiveProject()
            variation_folder = active_project.GetContents('Variations.IntPrjfolder', 1)[0]
            variation_object = variation_folder.GetContents(variation_name, 1)[0]
            return variation_object
        except Exception as e:
            self.print("Variation not found!: " + str(e))
            self.print(traceback.format_exc())
            return 1     
    
    
        
        
    def activate_variation(self, variation_object):
        '''
        function activating the given "variation" object (IntScheme)
        '''
        try:
            return variation_object.Activate()
        except Exception as e:
            self.print("Variation activation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
      
        
    def deactivate_variation(self, variation_object):
        '''
        function deactivating the given "variation" object (IntScheme)
        '''
        try:
            return variation_object.Deactivate()
        except Exception as e:
            self.print("Variation deactivation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
        
        
    def activate_study_case(self, study_case_object):
        '''
        function activating the given study case object (IntCase)
        '''
        try:
            return study_case_object.Activate()
        except Exception as e:
            self.print("Study case activation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
      
        
    def deactivate_study_case(self, study_case_object):
        '''
        function deactivating the given study case object (IntScheme)
        '''
        try:
            return study_case_object.Deactivate()
        except Exception as e:
            self.print("Study case deactivation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
        

    def enable_pf_gui_update(self):
        '''
        enable the update of the pf graphical interface
        '''
        self.app.SetGuiUpdateEnabled(1)


    def disable_pf_gui_update(self):
        '''
        disable any update of the pf graphical interface
        '''
        self.app.SetGuiUpdateEnabled(0) 

            
    def clear_output_window(self):
        '''
        function clearing the PowerFactory output window
        '''
        try:
            self.app.ClearOutputWindow() 
        except Exception as e:
            self.print("Clear output window error: " + str(e))
            self.print(traceback.format_exc())
            
    
    def show_page(self, page): 
        '''
        refresh/show the given page (typically a diagram page)
        '''
        try:
            graphic_board = self.app.GetGraphicsBoard()
            graphic_board.Show(page)
        except Exception as e:
            self.print("Show page error: " + str(e))
            self.print(traceback.format_exc())   
            
            
    def save_page_in_wmf(self, page, file_name): 
        '''
        save the active page as wmf file in the given path
        '''
        try:
            self.show_page(page)
            #self.app.Rebuild(2)
            #self.refresh_pf()
            graphic_board = self.app.GetGraphicsBoard()
            #self.print("  Writing " + file_name)
            graphic_board.WriteWMF(file_name)
        except Exception as e:
            self.print("Save page in wmf error: " + str(e))
            self.print(traceback.format_exc())   
    
    
    def print(self, string_to_print):
        '''
        function sending a string to the PowerFactory outputwindow and to a file
        if the self.output_file variable has been set
        '''
        self.app.PrintPlain(string_to_print)
        if self.output_file != None:
            self.output_file.write(string_to_print + '\n')
            
    
     
            
    
    
    

    
    

        
