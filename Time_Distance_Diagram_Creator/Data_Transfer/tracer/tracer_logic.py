# Encoding: utf-8
'''
Created on 2 Oct 2018

@author: AMB 
'''

import wx

from enum import Enum
from enum import IntEnum
from collections import namedtuple
import os

from pathlib import Path

from shutil import copy2

from docx import Document
from docx.shared import Inches
from PIL import Image, ImageFile
from PIL import WmfImagePlugin
from PIL.WmfImagePlugin import WmfHandler

from win32com import client

from tracer.grid import *

from itertools import repeat

from xlsxwriter.workbook import Workbook
from openpyxl import load_workbook
#from openpyxl import Workbook

from  calc_sw_interface.powerfactory_relay_interface import *

import webbrowser
import itertools
import tracer.branch

from calc_sw_interface.powerfactory_relay_interface import RelaySetting

import copy
from pydoc import doc

import win32com.client as win32

import datetime
import sys


class OutputDetail(IntEnum):
    DISABLED = 0
    NORMAL = 1
    DEBUG = 2
    VERBOSEDEBUG = 3



class PSET():
    '''
    Protective device setting calculation logic class
    '''

    def __init__(self, interface):
        '''
        Constructor
        '''
        self.interface = interface
        self.output_detail = OutputDetail.VERBOSEDEBUG
        self.new_file_name = ""
        
        

    def initialize(self, window, input_settings=None):
        '''
        function initializing all dictionaries to configure the network, 
        the fault type, the fault location and the other selection criteria
        ''' 

        # ldf status dictionary
        self.Ldfstatus = namedtuple("Ldfstatus", "ldf_already_calculated ldf_failed")

        self.ldf_status_list = { "Intact network" : self.Ldfstatus(ldf_already_calculated=False, ldf_failed=False)}
        # the actual ldf status
        self.ldf_status = self.ldf_status_list["Intact network"]
        return settings


################################################################################
################################################################################
################################################################################
    def run_PF_simulation_for_multiple_study_cases(self):
        '''
        run a PowerFCatory simulation along all available study cases
        '''
        self.run_PF_multiple_study_cases(\
                    self.run_rms_simulation)

################################################################################
################################################################################
################################################################################

    def run_Simulink_simulation(self):
        '''
        run a Simulink simulation along all available study cases
        '''
        a = 0
################################################################################
################################################################################
################################################################################

    def run_PF_multiple_study_cases(self, payload_function, window=None):
        '''
        run a process along all available study cases running for each of them
        the given payload function which does something
        '''
        # Print some info in the PF output window
        if self.output_detail.value >= OutputDetail.NORMAL.value:
            self.interface.print("*********************************************************")
            self.interface.print("            Run multiple study cases           ")
            self.interface.print("*********************************************************\n")
        # self.interface.set_echo_off()
        study_cases = self.interface.get_study_cases()
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
        settings = self.initialize(window, input_settings)
        if self.is_dialog_setting_ok(window, settings) == False:
            return 1
        if self.output_detail >= OutputDetail.NORMAL:
                self.interface.print("Saving all pictures ...")

        # create the Pictures directory if not present in the path where the
        # result file is
        picture_dir_path = os.path.split(window.results_file_name.GetValue())[0] + \
                                                 '\\Pictures'
        if not os.path.exists(picture_dir_path):
            os.mkdir(picture_dir_path)

        #just second step
        #self.add_wmfs_in_word(word_file_name = window.results_file_name.GetValue(),\
        #                     wmf_path = picture_dir_path)
        #sys.exit()
        # go throw all study cases
        #=======================================================================
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
        # here process the word file to replace the pictures
        self.add_wmfs_in_word(word_file_name=window.results_file_name.GetValue(), \
                             wmf_path=picture_dir_path)
        #=======================================================================
        if self.output_detail >= OutputDetail.NORMAL:
            self.interface.print("All pictures have been Transfered!")
        self.interface.print(" ***   Task completed    ***")
    
    
            
###############################################################################
# auxiliary function
###############################################################################
   
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
        pictures found in the given wmf_path
        '''
        # open word
        try:
        #first of all really try to open wps...no money for a Word license ;-)
            word = client.Dispatch("kwps.Application")
        except Exception as e:
            try:
                word = client.Dispatch("Word.Application")
            except Exception as e:
                self.interface.print("ERROR:    Exception Running Word!")
                return

        picture_tag = "<PIC>"
        picture_end_tag = "<\PIC>"

        doc = word.Documents.Open(word_file_name)
        word.Visible = True
        for i in range(doc.Paragraphs.Count - 1):
            try:
                paragraph = doc.Paragraphs(i + 1).Range.Text
                if picture_tag in paragraph:
                    start_index = paragraph.find(picture_tag) + len(picture_tag)
                    end_index = paragraph.find(picture_end_tag)
                    wmf_info = paragraph[start_index:end_index].split(':')
                    # try to fet the data of a second picture on the same line
                    start_index2 = paragraph.find(picture_tag, end_index) + \
                                                                len(picture_tag)
                    end_index2 = paragraph.find(picture_end_tag, end_index + \
                                                        len(picture_end_tag))
                    wmf_info2 = paragraph[start_index2:end_index2].split(':')
                    two_pictures = False
                    #if start_index2 > 0:
                    #    two_pictures = True
                    try:
                        if self.output_detail >= OutputDetail.NORMAL:
                            self.interface.print("Transfering " + wmf_info[0] + \
                                                 " -" + wmf_info[1])
                        doc.Paragraphs(i + 1).Range.Text = ""
                        inlineshapes = doc.Paragraphs(i + 1).Range.Words(1).InlineShapes
                        new_picture = inlineshapes.AddPicture\
                        (wmf_path + '\\' + wmf_info[0] + '\\' + wmf_info[1] + '.wmf', \
                        doc.Paragraphs(i + 2).Range)
                        shape = inlineshapes.Item(1).ConvertToShape()
                        if two_pictures == True:
                            self.new_picture2 = inlineshapes.AddPicture\
                            (wmf_path + '\\' + wmf_info2[0] + '\\' + wmf_info2[1] \
                             + '.wmf', doc.Paragraphs(i + 2).Range)         
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
                    except Exception as e:
                        self.interface.print("ERROR: " + wmf_path + '\\' + \
                        wmf_info[0] + '\\' + wmf_info[1] + '.wmf' + " not found!")
            except Exception as e:
                self.interface.print("ERROR: at doc line " + str(i) + "(of" + \
                                     str(doc.Paragraphs.Count) + ")")

        doc.SaveAs(word_file_name.replace(".docx", " with pictures.docx"))

    
    
    #===========================================================================
    # Object serialization
    #===========================================================================

    def serialize(self, object_to_serialize, file_name):
        '''
#         generic function to serialize the object passed as parameter
#         the object can be up a 2 dimension matrix
#         '''
#         file_name = file_name.replace('.xml', '.json')
#         file_name = file_name.replace('.XML', '.json')
#         file_name = file_name.replace('.', type(object_to_serialize).__name__+'.') # the file where I save the obejct contains the object class name
#
#
#         with open(file_name, mode='w') as output_file:
#             json.dump(object_to_serialize, output_file)
#             if type(object_to_serialize) is list:
#                 for item in object_to_serialize:
#                     if type(item) is list:
#                         for subitem in item:
#                             json.dump(len(item), output_file)
#                             json.dump(subitem, output_file)
#                     else:
#                         json.dump(len(object_to_serialize), output_file)
#                         json.dump(item, output_file)
#             else:
#                 json.dump(object_to_serialize, output_file)
#             output_file.close()

    #=========================================================================
    #   Input/ouput XML functions
    #=========================================================================

    def writeXSL(self, XMLfilename):
        XSLfilename = XMLfilename.replace('.xml', '.xsl')
        XSLfilename = XSLfilename.replace('.XML', '.XSL')

        XSDfilename = XMLfilename.rsplit("\\", 1)[-1]
        XSDfilename = XSDfilename.replace('.xml', '.xsd')
        XSDfilename = XSDfilename.replace('.XML', '.XSD')
        with open(XSLfilename, 'w') as XSLOutputFile:
            XSLOutputFile.write("<?xml version=\"1.0\" ?>\n")
            XSLOutputFile.write(
                "<xsl:stylesheet xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\" version=\"1.0\" xmlns:schemaLocation=\"" + XSDfilename + "\">\n")
            XSLOutputFile.write(
                "<xsl:output method=\"html\" version=\"1.0\" encoding=\"UTF-8\" indent=\"yes\" />\n")
            XSLOutputFile.write(
                "<!-- File automatically created by PSET for Digsilent PowerFactory-->\n")
            XSLOutputFile.write("<xsl:template match=\"/CESIPSETRESULTS\">\n")
            XSLOutputFile.write(
                "<html><head><title>CESI Time Distance Diagram Creator 2019 (BETA) Results v0.01</title>\n")
            XSLOutputFile.write("<style media=\"screen\" type=\"text/css\">\n")
            XSLOutputFile.write(
                "table{border-collapse: collapse; border-spacing: 0;}\n")
            XSLOutputFile.write(
                ".CESItableformat {margin:0px;padding:0px;width:100%;border:1px solid #000000;}\n")
            XSLOutputFile.write(".CESItableformat table{\n")
            XSLOutputFile.write(
                "    border-collapse: collapse; border-spacing: 0; width:100%; height:100%; margin:0px;padding:0px;}\n")
            XSLOutputFile.write(".CESItableformat tr:hover {\n")
            XSLOutputFile.write("background-color:#ffffff;}\n")
            XSLOutputFile.write(".CESItableformat td{\n")
            XSLOutputFile.write(
                "vertical-align:middle;background-color:#6d7175;border:1px solid #000000;border-width:0px 1px 1px 0px;text-align:left;padding:5px;\n")
            XSLOutputFile.write(
                "font-size:12px;font-family:verdana;font-weight:normal;color:#ffffff;}\n")
            XSLOutputFile.write(
                ".CESItableformat tr:hover td{background-color:#edeeef;font-size:12px;font-family:verdana;font-weight:bold;color:#6d7175;repeat-x 0 0;}\n")
            XSLOutputFile.write(".CESItableformat td:hover span{\n")
            XSLOutputFile.write(
                "display:inline; position:absolute;     border:2px solid #FFF;  \n")
            XSLOutputFile.write(
                "font-size:12px;font-family:verdana;font-weight:bold;color:#000000;    background:#edeeef repeat-x 0 0;}\n")
            XSLOutputFile.write(".CESItableformat th:hover span{\n")
            XSLOutputFile.write(
                "display:inline; position:absolute; border:2px solid #FFF;  \n")
            XSLOutputFile.write(
                "font-size:12px;font-family:verdana;font-weight:bold;color:#000000;    background:#edeeef repeat-x 0 0;}\n")
            XSLOutputFile.write(".CESItableformat th{\n")
            XSLOutputFile.write(
                "background-color:#003f7f;border:0px solid #000000;text-align:left;border-width:0px 0px 1px 1px;\n")
            XSLOutputFile.write(
                "font-size:12px;font-family:verdana;font-weight:bold;color:#ffffff; padding:7px;}\n")
            XSLOutputFile.write(".CESItableformat span {\n")
            XSLOutputFile.write("z-index:10;display:none; padding:3px 3px;\n")
            XSLOutputFile.write(
                "    margin-top:40px; margin-left:20px; width:1500px; line-height:16px;}\n")
            XSLOutputFile.write("div {padding: 300px 0px 0px 0px;}\n")
            XSLOutputFile.write("</style>\n")
            XSLOutputFile.write("</head>\n")
            XSLOutputFile.write(
                "<table border=\"1\" class=\"CESItableformat\">\n")
            XSLOutputFile.write("<thead>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"11\">CESI Time Distance Diagram Creator Tool 2019 (BETA) V0.01 Results</th></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Results Created:</th><td colspan=\"9\"><xsl:value-of select=\"SimulationStartTime\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write("<tr><th colspan=\"2\"><span><table>\n")
            XSLOutputFile.write("<tr><th>Parameter</th><th>Value</th></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Minimum CTI</th><td><xsl:value-of select=\"MinCTI\"/><xsl:text> </xsl:text> <xsl:value-of select=\"..//TimeUnit\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Max Permissible Clearance Time for close-in Faults</th><td><xsl:value-of select=\"MaxClearanceTimeNearEnd\"/><xsl:text> </xsl:text> <xsl:value-of select=\"..//TimeUnit\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Max Permissible Clearance Time for remote-end Faults</th><td><xsl:value-of select=\"MaxClearanceTimeFarEnd\"/><xsl:text> </xsl:text> <xsl:value-of select=\"..//TimeUnit\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Definition of Fast Trip Time Faults</th><td><xsl:value-of select=\"MinClearanceTimeFarEnd\"/><xsl:text> </xsl:text> <xsl:value-of select=\"..//TimeUnit\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Maximum Reach for Fast Trippping</th><td><xsl:value-of select=\"MinClearanceDistFarEnd\"/><xsl:text> </xsl:text> <xsl:text> % </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Max Overall Permissible Fault Clearance Time</th><td><xsl:value-of select=\"MaxCT\"/><xsl:text> </xsl:text> <xsl:value-of select=\"..//TimeUnit\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Total Short Circuits Applied</th><td><xsl:value-of select=\"NumberShortCircuits\"/><xsl:text> </xsl:text> </td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Total Violations Found</th><td><xsl:value-of select=\"NumberViolations\"/><xsl:text> </xsl:text> </td></tr>\n")
            XSLOutputFile.write(
                "</table></span>Database:</th><td colspan=\"9\"><xsl:value-of select=\"CAPEDatabase\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Network Study Date:</th><td colspan=\"9\"><xsl:value-of select=\"StudyDate\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Voltage:</th><td colspan=\"9\"><xsl:value-of select=\"StudyVoltage\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Area:</th><td colspan=\"9\"><xsl:value-of select=\"StudyArea\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Zone:</th><td colspan=\"9\"><xsl:value-of select=\"StudyZone\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Grid:</th><td colspan=\"9\"><xsl:value-of select=\"StudyGrid\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Path:</th><td colspan=\"9\"><xsl:value-of select=\"StudyPath\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th colspan=\"2\">Simulation Bus:</th><td colspan=\"9\"><xsl:value-of select=\"StudyBus\"/><xsl:text> </xsl:text></td></tr>\n")
            XSLOutputFile.write(
                "<tr><th>Fault Number</th><th>From Station</th><th>To Station</th><th>Voltage (kV)</th><th>Circuit ID</th>\n")
            XSLOutputFile.write(
                "<th>Distance To Fault(%)</th><th>Fault Type</th><th>Contingency</th><th>Outage(s)</th><th>Fault Clearance Time (s)</th><th>Result</th></tr>\n")
            XSLOutputFile.write("</thead>\n")
            XSLOutputFile.write("<tbody>\n")
            XSLOutputFile.write("<xsl:apply-templates select=\".//Fault\"/>\n")
            XSLOutputFile.write("</tbody>\n")
            XSLOutputFile.write("</table>\n")
            XSLOutputFile.write(
                "<div>Copyright 2024 Michele Borgnino Software. All rights reserved.</div>\n")
            XSLOutputFile.write("</html>\n")
            XSLOutputFile.write("</xsl:template> \n")
            XSLOutputFile.write("<xsl:template match=\"Fault\">\n")
            XSLOutputFile.write("<tr>\n")
            XSLOutputFile.write("<td>\n")
            XSLOutputFile.write(
                "<span><table><tr><th>Station</th><th>Circuit Breaker</th><th>Voltage (kV)</th><th>Ckt ID</th><th>Tripping Relay</th><th>Tripping Element</th><th>Trip time (SECONDS)</th><th>psetdigsilent_test Result</th>\n")
            XSLOutputFile.write(
                "<th>IA (pu)</th><th>IB (pu)</th><th>IC (pu)</th><th>IN (pu)</th><th>Relay Setting Phase(Ground)(pu)</th></tr>\n")
            XSLOutputFile.write("<xsl:apply-templates select=\".//Relay\"/>\n")
            XSLOutputFile.write("</table></span>\n")
            XSLOutputFile.write("<xsl:value-of select=\"FaultNumber\"/>\n")
            XSLOutputFile.write("</td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"FromStation\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"RemoteStation\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"Voltage\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"CircuitID\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"DistanceToFault\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"FaultType\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"Contingency\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"OutagedElement\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"FaultClearanceTime\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"ProtectionPerformanceAssessment\"/></td>\n")
            XSLOutputFile.write("</tr>\n")
            XSLOutputFile.write("</xsl:template>\n")
            XSLOutputFile.write("<xsl:template match=\"Relay\">\n")
            XSLOutputFile.write("<tr>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"FromStation\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"ToStation\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"Voltage\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"CircuitID\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"LZOPTag\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"TrippingElement\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"TripTime\"/></td>\n")
            XSLOutputFile.write(
                "<td><xsl:value-of select=\"RelayPerformanceAssessment\"/></td>\n")
            XSLOutputFile.write("<td><xsl:value-of select=\"IFA\"/></td>\n")
            XSLOutputFile.write("<td><xsl:value-of select=\"IFB\"/></td>\n")
            XSLOutputFile.write("<td><xsl:value-of select=\"IFC\"/></td>\n")
            XSLOutputFile.write("<td><xsl:value-of select=\"IFN\"/></td>\n")
            XSLOutputFile.write("<td><xsl:value-of select=\"Irelay\"/></td>\n")
            XSLOutputFile.write("</tr>\n")
            XSLOutputFile.write("</xsl:template>\n")
            XSLOutputFile.write("</xsl:stylesheet>\n")
            XSLOutputFile.close()

    def writeXSD(self, XSDfilename):
        # print ("Writing XSD file")
        # XMLfilename = self.results_file_name.GetValue()
        XSDfilename = XSDfilename.replace('.xml', '.xsd')
        XSDfilename = XSDfilename.replace('.XML', '.XSD')
        with open(XSDfilename, 'w') as XSDOutputFile:
            XSDOutputFile.write(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n")
            XSDOutputFile.write(
                "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\n")
            XSDOutputFile.write(
                "<!--File automatically created by PSET for Digsilent PowerFactory-->\n")
            XSDOutputFile.write("<xs:element name=\"CESIPSETRESULTS\">\n")
            XSDOutputFile.write("  <xs:complexType>\n")
            XSDOutputFile.write("    <xs:sequence>   \n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyDate\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"DatabaseFile\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationStartTime\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationID\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyVoltage\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyArea\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyZone\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyGrid\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyPath\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyBus\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"TimeUnit\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MinCTI\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MaxCT\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MaxClearanceTimeNearEnd\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MinClearanceDistFarEnd\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MinClearanceTimeFarEnd\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MaxClearanceTimeFarEnd\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"OvercurrentMargin\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"ImpedanceMargin\" type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationDepth\" type=\"xs:positiveInteger\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"MutualDepth\" type=\"xs:positiveInteger\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"StudyDate\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationDate\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Fault\"  maxOccurs=\"unbounded\">\n")
            XSDOutputFile.write("<xs:complexType>\n")
            XSDOutputFile.write("<xs:sequence>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationID\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NetworkCaseID\" type=\"xs:string\"  minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NetworkStateID\" type=\"xs:string\" minOccurs=\"0\"/>  \n")
            XSDOutputFile.write(
                "<xs:element name=\"FaultNumber\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FromStation\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FromStationID\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"ToStation\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"ToStationID\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RemoteStation\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RemoteStationID\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Voltage\"  type=\"xs:decimal\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"CircuitID\"  type=\"xs:positiveInteger\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"OutagedElement\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Contingency\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"DistanceToFault\"  type=\"xs:decimal\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FaultType\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FaultClearanceTime\"  type=\"xs:decimal\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"ProtectionPerformanceAssessment\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Relay\"  maxOccurs=\"unbounded\" minOccurs=\"0\">\n")
            XSDOutputFile.write("  <xs:complexType>\n")
            XSDOutputFile.write("<xs:sequence>  \n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationID\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NetworkCaseID\" type=\"xs:string\"  minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NetworkStateID\" type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FaultNumber\" type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"FromStation\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"ToStation\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RemoteStation\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"CircuitID\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Voltage\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"LZOPTag\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RelayTag\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RelayName\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RelayModel\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"TrippingElement\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"TrippingCharacteristicValue\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"SimulationMeasuredValue\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"TripTime\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"CBOpenTime\"  type=\"xs:string\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"RelayPerformanceAssessment\"  type=\"xs:string\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"IFA\"  type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"IFB\"  type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"IFC\"  type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"IFN\"  type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"Irelay\"  type=\"xs:decimal\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write("</xs:sequence>\n")
            XSDOutputFile.write("  </xs:complexType>\n")
            XSDOutputFile.write("</xs:element> \n")
            XSDOutputFile.write("</xs:sequence>\n")
            XSDOutputFile.write("  </xs:complexType>\n")
            XSDOutputFile.write("</xs:element> \n")
            XSDOutputFile.write(
                "<xs:element name=\"BusesStudied\"  type=\"xs:integer\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"LinesStudied\"  type=\"xs:integer\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"TransformersStudied\"  type=\"xs:integer\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NumberShortCircuits\"  type=\"xs:integer\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write(
                "<xs:element name=\"NumberViolations\"  type=\"xs:integer\" minOccurs=\"0\"/>\n")
            XSDOutputFile.write("    </xs:sequence>\n")
            XSDOutputFile.write("</xs:complexType>\n")
            XSDOutputFile.write("</xs:element>\n")
            XSDOutputFile.write("</xs:schema> \n")
            XSDOutputFile.close()

