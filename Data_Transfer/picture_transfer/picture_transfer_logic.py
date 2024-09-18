# Encoding: utf-8
'''
Created on 1 June 2024

@author: Michele Borgnino
'''

import os
import sys

import wx

from enum import IntEnum

from docx import Document
from win32com import client

from xlsxwriter.workbook import Workbook
from openpyxl import load_workbook


from pydoc import doc
import win32com.client as win32

import shutil

class OutputDetail(IntEnum):
    DISABLED = 0
    NORMAL = 1
    DEBUG = 2
    VERBOSEDEBUG = 3



class TransferLogic():
    '''
    Main logic class
    '''

    def __init__(self, interface):
        '''
        Constructor
        '''
        self.interface = interface
        self.output_detail = OutputDetail.VERBOSEDEBUG
        self.new_file_name = ""
        
        



################################################################################
################################################################################
################################################################################
    def run_PF_simulation_for_multiple_study_cases(self):
        '''
        run a PowerFactory simulation along all available study cases
        '''
        self.run_PF_multiple_study_cases(\
                    self.run_rms_simulation)

################################################################################
################################################################################
################################################################################

    def run_Simulink_simulation(self, source_file_full_name):
        '''
        run a Simulink simulation
        '''
        # get the dir path of the source file
        source_file_dir = os.path.dirname(source_file_full_name)
        # remove the .xls extension from the source file full name
        source_file_full_name = os.path.basename(source_file_full_name).\
                                                             replace('.slx','')
        
        import subprocess
        #get the template file
        template_file = open("matlab_runner_template.py","r+")
        template_content = template_file.read()
        # replace the <TARGET> string with the file full path
        target_ID = "<TARGET>"
        out_content = template_content.replace(target_ID, source_file_full_name)
        # replace the <TARGETDIR> string with the file dircetory full path
        targetdir_ID = "<TARGETDIR>"
        out_content = out_content.replace(targetdir_ID, source_file_dir)
        out_content = out_content.replace("\\", "\\\\")
        # create the python script file to run
        with open("matlab_runner.py", "w") as out_text_file:
            out_text_file.write(out_content)
    
        # Run the command whcih should operate in the 3.11 virtual environment
        script_path = os.path.dirname(os.path.realpath(__file__)).\
                                                replace("\\picture_transfer",'')
        script_path = script_path.replace("\\", "\\\\")
        subprocess.Popen("activate.bat", cwd = script_path, stdin = subprocess.PIPE, shell = True)
        
################################################################################
################################################################################
################################################################################

    def run_PF_multiple_study_cases(self, payload_function, window=None):
        '''
        run a process along all available study cases running for each of them
        the given payload function which does something
        '''
        # Print some info in the PF output window
        try:
            if self.output_detail.value >= OutputDetail.NORMAL.value:
                self.interface.print("*********************************************************")
                self.interface.print("            Run multiple study cases           ")
                self.interface.print("*********************************************************\n")
            # self.interface.set_echo_off()
            study_cases = self.interface.get_study_cases()
        except Exception as e:
            dlg = wx.MessageDialog(window, "DIgSILENT PowerFactory not found!",
                                               "Warning", \
                                               wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return
        for study_case in study_cases:
            self.interface.activate_study_case(study_case)
            if self.output_detail >= OutputDetail.NORMAL:
                self.interface.set_echo_on()
                self.interface.print(self.interface.get_name_of(study_case) + \
                                  " study case has been activated")
                self.interface.set_echo_off()
            # defintion of the function which performs the required operations
            # like running a simulation, a LDF etc
            payload_function(study_case, window)
        self.interface.set_echo_on()
        if self.output_detail.value >= OutputDetail.NORMAL.value:
            self.interface.print("***     Run multiple study cases  completed     ***")

################################################################################
    def run_rms_simulation(self, study_case, window):
        '''
        run the single study case passed as parameter
        '''    
        self.interface.set_echo_off() 
        # here init and run the simulation
        #default name for the event folder in the PF data manager
        event_name = 'Simulation Events/Fault'
        events = self.interface.get_study_case_events(study_case, event_name)
        study_case_name = self.interface.get_name_of(study_case)
        simulation_init_objects = self.interface.get_simulation_inits_of(study_case)
        # check that the simulation init object and the events are availablle
        if simulation_init_objects and events:
            simulation_init_objects[0].Execute()
            simulation_objects = self.interface.\
                                        get_simulation_objects_of(study_case)
            if simulation_objects:
                # run the simulation
                simulation_objects[0].Execute()
            else:
                self.interface.set_echo_on()
                self.interface.print("ERROR: no sim object available for " + \
                self.interface.get_name_of(study_case) + " study case")
        else:
            self.interface.set_echo_on()
            if simulation_init_objects:
                self.interface.print("ERROR: no events available for " + \
                                            study_case_name + " study case")
            else:
                self.interface.print("ERROR: no simulation init available for " + \
                                            study_case_name + " study case")
        self.interface.set_echo_on()

 
################################################################################
################################################################################
################################################################################
   
    def move_all_pictures_to_word(self, window, input_settings=None):
        '''
        function getting all pictures available in all study cases, saving them
        in the Pictures
        '''
        # just for test
        # self.interface.rebuild_pf()
        # return
        # just check that the settings are ok
        word_file_name = window.results_file_name.GetValue()  
        excel_file_name = window.source_file_name.GetValue() 
        if word_file_name == "":
            dlg = wx.MessageDialog(window, "Please specify a valid results filename.",
                                           "No result filename", \
                                           wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return
        if excel_file_name == "":
            dlg = wx.MessageDialog(window, "Please specify a valid xls filename.",
                                           "No xls filename", \
                                           wx.OK | wx.ICON_WARNING)
            dlg.ShowModal()
            dlg.Destroy()
            return
        
        if self.output_detail >= OutputDetail.NORMAL:
                print("Saving all pictures ...")

        # create the Pictures directory if not present in the path where the
        # result file is
        picture_dir_path = os.path.split(window.results_file_name.GetValue())[0] + \
                                                 '\\Pictures'
        if not os.path.exists(picture_dir_path):
            os.mkdir(picture_dir_path)

        #just second step
#         self.add_wmfs_in_word(word_file_name = window.results_file_name.GetValue(),\
#                              wmf_path = picture_dir_path)
#         self.add_symulink_pictures_in_word(word_file_name=window.results_file_name.GetValue(), \
#                              emf_path=picture_dir_path)
#         sys.exit()
        # go throw all study cases
        #=======================================================================
        try:
            study_cases = self.interface.get_study_cases()
            for study_case in study_cases:
                # activate the study case
                self.interface.set_echo_off()
                self.interface.activate_study_case(study_case)
                if self.output_detail >= OutputDetail.NORMAL:
                    self.interface.set_echo_on()
                    self.interface.print(self.interface.get_name_of(study_case) + \
                                      " study case has been activated")
                    self.interface.set_echo_off()
                diagram_path = picture_dir_path + '\\' + \
                                            self.interface.get_name_of(study_case)
                # save the diagrams             
                self.save_diagrams_as_wmf(window, path=diagram_path, create_copy=True)
                self.interface.set_echo_on()       
            if self.output_detail >= OutputDetail.NORMAL:
                    self.interface.print("All pictures have been saved!")
                    self.interface.print("Trasfering pictures to word....")
        except Exception:
            pass
        # here process the word file to replace the pictures
        # here the pictures coming from PF
        self.add_wmfs_in_word(word_file_name=window.results_file_name.GetValue(), \
                             wmf_path=picture_dir_path)
        # here the pictures coming from Matlab
        self.add_symulink_pictures_in_word(word_file_name=window.results_file_name.\
                            replace(".docx", " with pictures.docx").GetValue(), \
                             picture_path=picture_dir_path)
        #=======================================================================
        try:
            if self.output_detail >= OutputDetail.NORMAL:
                self.interface.print("All pictures have been Transfered!")
            self.interface.print(" ***   Task completed    ***")
        except Exception:
            pass
    
            
###############################################################################
# auxiliary function
###############################################################################
   
    def create_copy_of(self, file_name, extension='.xls'):
        '''
        create a copy of the given faile name (if existing) adding the date/time in
        the file name
        '''
        import os.path
        from shutil import copyfile
        from datetime import datetime
        import time
        if os.path.isfile(file_name):
            file_date_str = time.ctime(os.path.getctime(file_name))
            file_date_str = file_date_str.replace(':', '-')
            new_file_name = file_name.replace(extension, file_date_str + extension)
            copyfile(file_name, new_file_name)
            return new_file_name
        return ""
    
    def save_diagrams_as_wmf(self, window, diagram_name='', path=None, \
                                    create_copy=True):
        '''
        save as a wmf file in the given path directory all graphical diagram
         with name containing or equal to diagram_name
        if no path is provided the path of the script output file is used
        if the create_copy parameter is True a copy of the file with the date/hour
        is created 
        '''
        from shutil import copyfile
        from datetime import datetime
        import time
        dir_path = path if path != None else\
                os.path.split(window.results_file_name.GetValue())[0] + '\\Pictures'
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
        
        diagram_pages = self.interface.get_diagram_pages(diagram_name)
        for diagram_page in diagram_pages:
            self.new_file_name = ""
            file_name = (dir_path + '\\' + self.interface.\
                         get_name_of(diagram_page))
            if create_copy == True:
                self.new_file_name =self.create_copy_of(file_name + '.wmf', extension='.wmf')
            try:
                os.remove(file_name + '.wmf')
            except:
                pass
            if self.output_detail >= OutputDetail.NORMAL:
                            self.interface.print("    Saving " + file_name)
            self.interface.save_page_in_wmf(diagram_page, file_name)
            # if the file has not created restored the previous one
            extension='.wmf'
            file_name = file_name + extension
            if os.path.isfile(file_name) == False:
                self.interface.print("ERROR: " + file_name + " not saved!") 
                if len(self.new_file_name) > 0:
                    copyfile(self.new_file_name, file_name)
                

    def create_word_with_wmf(self, window, diagram_name, file_name, file_path=None,
                             wmf_path=None):
        '''
        create the the given file_name word file containing the wmf with name  
        containing the or equal to diagram_name. the wmf file are searched in the 
        wmf_path is provided otherwise in the "wmf" directory in the result file
        path
        '''
        file_path = file_path if file_path != None else\
                os.path.split(window.results_file_name.GetValue())[0]
        wmf_path = wmf_path if wmf_path != None else\
                os.path.split(window.results_file_name.GetValue())[0] + '\\wmf'
        # open word
        try:
            # word = client.Dispatch("Word.Application")
            word = client.Dispatch("kwps.Application")
        except Exception as e:
            try:
                word = client.Dispatch("Word.Application")
            except Exception as e:
                self.interface.print("ERROR:    Exception Running Word!")
                return
        Doc = word.Documents.Open(file_path + '\\' + file_name)
        word.Visible = True

        diagram_pages = self.interface.get_diagram_pages(diagram_name)
        for index, diagram_page in enumerate(diagram_pages):
            pic = Doc.Paragraphs(index + 1).Range.Words(2).InlineShapes.AddPicture\
            (wmf_path + '\\' + self.interface.get_name_of(diagram_page) + '.wmf')

    def add_wmfs_in_word(self, word_file_name, wmf_path):
        '''
        replace inside the given word file all the tags with the relevant wmf 
        pictures found in the given wmf_path. The wmf pictures are from PF
        '''
        # open word
        try:
        #first of all really try to open wps...in case of no POLIMI Word license ;-)
            word = client.Dispatch("kwps.Application")
        except Exception as e:
            #otherwise try to use Word
            try:
                word = client.Dispatch("Word.Application")
            except Exception as e:
                self.interface.print("ERROR:    Exception Running Word!")
                return

        picture_tag = "<PIC>"
        picture_end_tag = "<\PIC>"

        doc = word.Documents.Open(word_file_name)
        word.Visible = True
        wmf_added = False
        for i in range(doc.Paragraphs.Count - 1):
            try:
                paragraph = doc.Paragraphs(i + 1).Range.Text
                if picture_tag in paragraph:
                    start_index = paragraph.find(picture_tag) + len(picture_tag)
                    end_index = paragraph.find(picture_end_tag)
                    wmf_info = paragraph[start_index:end_index].split(':')
                    #if we have only three set of info it isn't a PF picture
                    if len(wmf_info) < 4:
                        continue
                    # try to fet the data of a second picture on the same line
                    start_index2 = paragraph.find(picture_tag, end_index) + \
                                                                len(picture_tag)
                    end_index2 = paragraph.find(picture_end_tag, end_index + \
                                                        len(picture_end_tag))
                    wmf_info2 = paragraph[start_index2:end_index2].split(':')
                    two_pictures = False
                    if end_index2 > 0:
                        two_pictures = True
                    try:
                        if self.output_detail >= OutputDetail.NORMAL:
                            self.interface.print("Transfering " + wmf_info[0] + \
                                                 " -" + wmf_info[1])
                        doc.Paragraphs(i + 1).Range.Text = ""
                        inlineshapes = doc.Paragraphs(i + 1).Range.Words(1).InlineShapes
                        path = wmf_path + '\\' + wmf_info[0] + '\\' + wmf_info[1] + '.wmf'
                        new_picture = inlineshapes.AddPicture(path)
                        shape = inlineshapes.Item(1).ConvertToShape()
                        if two_pictures == True:
                            self.new_picture2 = inlineshapes.AddPicture\
                            ('"'+wmf_path + '\\' + wmf_info2[0] + '\\' + wmf_info2[1])         
                            self.shape2 = inlineshapes.Item(1).ConvertToShape()      
                        # magic numbers to crop the picture from the original PF wmf
                        # at home big screen 210/290
                        new_picture.PictureFormat.CropBottom = 120 #2 #145 #120 #175
                        new_picture.PictureFormat.CropRight = 170 # 5 #210 #170 #250
                        if two_pictures == True:
                            self.new_picture2.PictureFormat.CropBottom = 145 #2 #145 #120 #175
                            self.new_picture2.PictureFormat.CropRight = 210 # 5 #210 #170 #250
                        # new_picture.ScaleWidth = 87
                        # new_picture.ScaleHeight = 68
                        new_picture.ScaleWidth = wmf_info[2] if len(wmf_info) > 2 else 51.5
                        new_picture.ScaleHeight = wmf_info[3] if len(wmf_info) > 3 else 40 
                        if two_pictures == True:
                            self.new_picture2.ScaleWidth = wmf_info2[2] if len(wmf_info2) > 2 else 51.5
                            self.new_picture2.ScaleHeight = wmf_info2[3] if len(wmf_info2) > 3 else 40                      
                        if two_pictures == True:
                            shape.WrapFormat.Type = 3 #6
                            self.shape2.WrapFormat.Type = 3 
                        else:
                            shape.WrapFormat.Type = 4   # 4 wdWrapFront
                        shape.WrapFormat.AllowOverlap = True  # False
                        shape.Left = word.CentimetersToPoints(wmf_info[4] \
                                                    if len(wmf_info) > 4 else 0.001)
                        shape.Top = word.CentimetersToPoints(wmf_info[5]\
                                                    if len(wmf_info) > 5 else 0.01)
                        if two_pictures == True:
                            self.shape2.Left = word.CentimetersToPoints(wmf_info2[4] \
                                                    if len(wmf_info2) > 4 else 7.5)
                            self.shape2.Top = word.CentimetersToPoints(wmf_info2[5]\
                                                    if len(wmf_info2) > 5 else 0.01)
                        wmf_added = True
                    except Exception as e:
                        self.interface.print("ERROR: " + wmf_path + '\\' + \
                        wmf_info[0] + '\\' + wmf_info[1] + '.wmf' + " not found!")
            except Exception as e:
                self.interface.print("ERROR: at doc line " + str(i) + "(of" + \
                                     str(doc.Paragraphs.Count) + ")")
        if wmf_added:
            doc.SaveAs(word_file_name.replace(".docx", " with pictures.docx"))

    
    def add_symulink_pictures_in_word(self, word_file_name, picture_path):
        '''
        replace inside the given word file all the tags with the relevant symulink 
        pictures found in the given picture_path
        this is the code to import the matlab pictures, the approch is not the 
        same used for PF: there are no study cases and all pictures are together 
        '''
        # open word
        try:
        #first of all really try to open wps...no money for a Word license ;-)
            word = client.Dispatch("kwps.Application")
        except Exception as e:
            #otherwise try to use Word
            try:
                word = client.Dispatch("Word.Application")
            except Exception as e:
                self.interface.print("ERROR:    Exception Running Word!")
                return

        picture_tag = "<PIC>"
        picture_end_tag = "<\PIC>"

        doc = word.Documents.Open(word_file_name)
        word.Visible = True
        emf_added = False
        for i in range(doc.Paragraphs.Count - 1):
            try:
                paragraph = doc.Paragraphs(i + 1).Range.Text
                if picture_tag in paragraph:
                    start_index = paragraph.find(picture_tag) + len(picture_tag)
                    end_index = paragraph.find(picture_end_tag)
                    wmf_info = paragraph[start_index:end_index].split(':')
                    # try to get the data of a second picture on the same line
                    start_index2 = paragraph.find(picture_tag, end_index) + \
                                                                len(picture_tag)
                    end_index2 = paragraph.find(picture_end_tag, end_index + \
                                                        len(picture_end_tag))
                    wmf_info2 = paragraph[start_index2:end_index2].split(':')
                    two_pictures = False
                    ext = '.emf'
                    if end_index2 > 0:
                        two_pictures = True
                    try:
                        if self.output_detail >= OutputDetail.NORMAL:
                            self.interface.print("Transfering " + wmf_info[0])
                        doc.Paragraphs(i + 1).Range.Text = ""
                        inlineshapes = doc.Paragraphs(i + 1).Range.Words(1).InlineShapes
                        new_file = picture_path + '\\' + wmf_info[0] + ext
                        new_picture = inlineshapes.AddPicture\
                        ( picture_path + '\\' + wmf_info[0] + ext)
                        shape = inlineshapes.Item(1).ConvertToShape()
                        if two_pictures == True:
                            self.new_picture2 = inlineshapes.AddPicture\
                            (picture_path + '\\' + wmf_info2[0] + ext)
                            self.shape2 = inlineshapes.Item(1).ConvertToShape()      
                        # magic numbers to crop the picture from the original PF wmf
                        # at home big screen 210/290
                        new_picture.PictureFormat.CropBottom = 1 #2 #145 #120 #175
                        new_picture.PictureFormat.CropRight = 1 # 5 #210 #170 #250
                        if two_pictures == True:
                            self.new_picture2.PictureFormat.CropBottom = 145 #2 #145 #120 #175
                            self.new_picture2.PictureFormat.CropRight = 210 # 5 #210 #170 #250
                        # new_picture.ScaleWidth = 87
                        # new_picture.ScaleHeight = 68
                        new_picture.ScaleWidth = wmf_info[1] if len(wmf_info) > 1 else 51.5
                        new_picture.ScaleHeight = wmf_info[2] if len(wmf_info) > 2 else 40 
                        if two_pictures == True:
                            self.new_picture2.ScaleWidth = wmf_info2[1] if len(wmf_info2) > 1 else 51.5
                            self.new_picture2.ScaleHeight = wmf_info2[2] if len(wmf_info2) > 2 else 40                      
                        if two_pictures == True:
                            shape.WrapFormat.Type = 3 #6
                            self.shape2.WrapFormat.Type = 3 
                        else:
                            shape.WrapFormat.Type = 4   # 4 wdWrapFront
                        shape.WrapFormat.AllowOverlap = False
                        shape.Left = word.CentimetersToPoints(wmf_info[3] \
                                                    if len(wmf_info) > 4 else 0.001)
                        shape.Top = word.CentimetersToPoints(wmf_info[4]\
                                                    if len(wmf_info) > 5 else 0.01)
                        if two_pictures == True:
                            self.shape2.Left = word.CentimetersToPoints(wmf_info2[3] \
                                                    if len(wmf_info2) > 4 else 7.5)
                            self.shape2.Top = word.CentimetersToPoints(wmf_info2[4]\
                                                    if len(wmf_info2) > 5 else 0.01)
                        emf_added = True
                    except Exception as e:
                        self.interface.print("ERROR: " + picture_path + '\\' + \
                        wmf_info[0] + ext + " not found!")
            except Exception as e:
                self.interface.print("ERROR: at doc line " + str(i) + "(of" + \
                                     str(doc.Paragraphs.Count) + ")")
        if emf_added:
            doc.SaveAs(word_file_name.replace(".docx", " with pictures.docx"))
   
    
    def move_all_pictures_together(self, source_file_full_path, output_file_full_path):
        '''
        move all pictures save by simulink (if present) in the "Pictures" directoryloacted
        at the same level where the output file template is
        '''
        source_file_path = os.path.dirname(source_file_full_path) + "\\Pictures"
        file_names = os.listdir(source_file_path)
        output_file_dir = os.path.dirname(output_file_full_path) + "\\Pictures"
        if source_file_path != output_file_dir:
            for file_name in file_names:
                try:
                    shutil.move(os.path.join(source_file_path, file_name),\
                                os.path.join(output_file_dir, file_name))
                except:
                    pass
