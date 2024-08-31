'''
Created on 23 Oct 2018

@author: AMB
'''


from datetime import date  # @UnusedImport
from datetime import datetime  # @UnusedImport




class Report_maker(object):
    '''
    class generating the xml tracer result report
    Responsibility: generate the ouput without any assessment logic
    '''


    def __init__(self, xml_file_name, relays, grid,\
                         settings, number_of_applied_shc, assessment_result,
                         no_trip_time, fault_no_cleared_id, list_object):
        '''
        Constructor
        '''
        self.xml_file_name = xml_file_name
        self.relays = relays
        self.grid = grid
        self.settings = settings
        self.number_of_applied_shc = number_of_applied_shc
        
        self.now = datetime.now().strftime("%c")
        self.todaysdate = date.today().strftime("%x")
        self.simulationID = datetime.now().strftime("%Y%m%d%H%M%S%f")
        
        self.event = None
        self.is_ldf = False # flag  on when the last event is a ldf
         
        self.event_index = 0
        self.number_of_violations = 0
        
        self.list_object = list_object
        
        self.assessment_result = assessment_result
        
        self.output_data = []
        
        self.no_trip_time = no_trip_time
        self.fault_no_cleared_id = fault_no_cleared_id
        
        self.xsd_filename = xml_file_name.rsplit("\\", 1)[-1]
        self.xsd_filename = self.xsd_filename.replace('.xml', '.xsd')
        self.xsd_filename = self.xsd_filename.replace('.XML', '.xsd')
        self.xsl_filename = self.xsd_filename.replace('.xsd', '.xsl')
        
        self.xmlOutputFile = open(self.xml_file_name, 'w')
        self._write_string('?xml version="1.0" encoding="UTF-8"?')
        self._write_string('?xml-stylesheet type="text/xsl" href="'+ self.xsl_filename + '"?')
        
        
    def write_study_info(self):
        '''
        function saving in the xml file all the study initial info
        '''    
        study_parameters = {'StudyDate'             : self.todaysdate,
                            'DatabaseFile'          : str(self.settings['PfdFileName']),
                            'SimulationStartTime'   : self.now,
                            'SimulationID'          : self.simulationID,
                            'StudyVoltage'          : str(self.list_object.voltage_list\
        [self.settings['VoltageList'][0]]) if len(self.settings['VoltageList']) > 0 else '',
                            'StudyArea'             : str(self.list_object.area_list\
        [self.settings['AreaList'][0]]) if len(self.settings['AreaList']) > 0 else '',
                            'StudyZone'             : str(self.list_object.zone_list\
        [self.settings['ZoneList'][0]]) if len(self.settings['ZoneList']) > 0 else '',
                            'StudyGrid'             : str(self.list_object.grid_list\
        [self.settings['GridList'][0]]) if len(self.settings['GridList']) > 0 else '',
                            'StudyPath'             : str(self.list_object.path_list\
        [self.settings['PathList'][0]]) if len(self.settings['PathList']) > 0 else '',
                            'StudyBus'              : str(self.settings['StudySelectedBus']),
                            'StudyBusTiers'         : str(self.settings['StudySelectedBusExtent']),
                            'Nintact'               : str(int(self.settings['Nintact'])),
                            'Nmin'                  : str(int(self.settings['Nmin'] == True)),
                            'Nmin2nd'               : str(int(self.settings['Nmin2nd'] == True)),
                            'NminGlobal'            : str(int(self.settings['NminGlobal'])),
                            'NminAll'               : str(int(self.settings['NminAll'])),
                            'Nmin2'                 : str(int(self.settings['Nmin2'])),
                            'Ncbf'                  : str(int(self.settings['Ncbf'])),
                            'SLGFAULT'              : str(int(self.settings['Fslg'] == True)),
                            'LTLFAULT'              : str(int(self.settings['Fltl'] == True)),
                            'DLGFAULT'              : str(int(self.settings['Fdlg'] == True)),
                            'TPHFAULT'              : str(int(self.settings['Ftph'] == True)),
                            'SLGOHMFAULT'           : str(self.settings['FslgrValue']),
                            'LTLOHMFAULT'           : str(self.settings['FltlrValue']),
                            'DLGOHMFAULT'           : str(self.settings['FdlgrValue']) ,
                            'CLOSEINFAULT'          : str(int(self.settings['flCloseIn'])),
                            'REMOTEENDFAULT'        : str(int(self.settings['flRemoteEnd'])),
                            'REMOTEENDOPENFAULT'    : str(int(self.settings['flRemoteEndOpen'] == True)),
                            'MIDLINE1'              : str(self.settings['flMidLine1Value']) ,
                            'MIDLINE2'              : str(self.settings['flMidLine2Value']),
                            'MIDLINE3'              : str(self.settings['flMidLine3Value']),
                            'TimeUnit'              : "Seconds",
                            'MinCTI'                : str(self.settings['ppMinCoordinationTimeMargin']),
                            'MaxCT'                 : str(self.settings['ppMaxClearanceTime']),
                            'MaxClearanceTimeNearEnd': str(self.settings['ppMaxNearEndTime']) ,
                            'MaxClearanceTimeFarEnd': str(self.settings['ppMaxFarEndTime']),
                            'MinClearanceDistFarEnd': str(self.settings['ppMaxInstReach']),
                            'MinClearanceTimeFarEnd': str(self.settings['ppMaxFastTrippingTime']),
                            'OvercurrentMargin'     : str(self.settings['ppNearMissOCValue']),
                            'ImpedanceMargin'       : str(0),
                            'SimulationDepth'       : str(self.settings['ppProtSimDepthValue']),
                            'NumberShortCircuits'   : str(self.number_of_applied_shc),
                            'NumberViolations'      : str(self.number_of_violations)
                            }
        
        self._write_string('CESITDDCRESULTS xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" \
                    xsi:noNamespaceSchemaLocation="' + self.xsd_filename +'"')
        self._write_string('!--File automatically created by PSET for DIgSILENT PowerFactory--')
      
        self._write_parameters(study_parameters)
        
        
    def write_fault(self, fault_index, fault, error_message_id):
        '''
        function saving in the xml file all the fault (pset_logic) data
        it calls in turn the write_outage function to save the outage data
        '''
        #import pydevd
        #pydevd.settrace()
        error_message_id = fault.error if fault.error > 0 else error_message_id
           
        fault_parameters = { 'SimulationID'      :  self.simulationID,
                             'NetworkCaseID'     :  '1',
                             'NetworkStateID'    :  fault.network_status,
                             'FaultNumber'       :  str(fault_index),
                             'FromStation'       :  fault.from_station,
                             'ToStation'         :  fault.to_station,
                             'RemoteStation'     :  fault.to_station,
                             'Voltage'           :  str(fault.voltage),
                             'CircuitID'         :  '1',
                             'FaultArea'         :  fault.area,
                             'FaultZone'         :  fault.zone,
                             'Contingency'       :  fault.network_status,
                             'OutagedElement'    :  fault.disconnected_elements_names,
                             'DistanceToFault'   :  str(round(fault.fault_position, 0)),
                             'FaultType'         :  fault.type + ", R = {} (ohm)".format(fault.fault_resistance),
                             'ProtectionPerformanceAssessment' : str(int(error_message_id)),
                             'FaultClearanceTime':  str(fault.fault_clearance_time)
                            }
        
        if self.event != None:
            self._write_end('Outage')
            self._write_end('Fault')
        self.event = fault
        self.is_ldf = False
        self._write_start('Fault')
        
        self._write_parameters(fault_parameters)
        self.write_outage(fault_index, fault)
        if int(error_message_id) != 0:
            self.number_of_violations +=1
    
    
    def write_ldf(self, ldf_index, ldf, error_message_id):
        '''
        function saving in the xml file all the ldf (pset_logic) data
        it calls in turn the write_outage function to save the outage data
        '''
        error_message_id = ldf.error if ldf.error > 0 else error_message_id
           
        fault_parameters = { 'SimulationID'      :  self.simulationID,
                             'NetworkCaseID'     :  '1',
                             'NetworkStateID'    :  ldf.network_status,
                             'FaultNumber'       :  str(ldf_index),
                             'FromStation'       :  "",
                             'ToStation'         :  "",
                             'RemoteStation'     :  "",
                             'Voltage'           :  "",
                             'CircuitID'         :  '1',
                             'FaultArea'         :  "",
                             'FaultZone'         :  "",
                             'Contingency'       :  ldf.network_status,
                             'OutagedElement'    :  ldf.disconnected_elements_names,
                             'DistanceToFault'   :  "",
                             'FaultType'         :  "LDF",
                             'ProtectionPerformanceAssessment' : str(int(error_message_id)),
                             'FaultClearanceTime':  str(ldf.ldf_trip_time)
                            }
        
        if self.event != None and ldf != self.event:
            self._write_end('Outage')
            self._write_end('Fault')
        self.event = ldf
        self.is_ldf = True
        self._write_start('Fault')
        
        self._write_parameters(fault_parameters)
        self.write_outage(ldf_index, ldf)
        if int(error_message_id) != 0:
            self.number_of_violations +=1
    
    
    def write_outage(self, fault_index, fault):
        '''
        function saving in the xml file all the outage (pset_logic) data
        '''      
        outage_parameters = {  'SimulationID'      :  'self.simulationID',
                               'NetworkCaseID'     :  '1',
                               'NetworkStateID'    :  fault.network_status,
                               'FaultNumber'       :  str(fault_index),
                               'OutageID'          :  '1',
                               'OutageDescription' :  'Base Case with no user-defined outaged applied'  
                            }
        if self.event != None:
            self._write_start('Outage')
            self._write_parameters(outage_parameters)
      
       
    def write_relay(self, relay, fault, tripping_data, error_message_id):
        '''
        function saving in the xml file all the relay (pset_logic) data
        '''
        relay_parameters = {'SimulationID'              : self.simulationID,
                            'NetworkCaseID'             : '1',
                            'NetworkStateID'            : fault.network_status,
                            'FaultNumber'               : str(self.event_index),
                            'OutageID'                  : '1',
                            'FromStation'               : relay.from_station,
                            'ToStation'                 : relay.to_station,
                            'RemoteStation'             : relay.to_station,
                            'Voltage'                   : str(relay.voltage),
                            'CircuitID'                 : '1',
                            'LZOPTag'                   : relay.name + '(' + \
                            self._filterstr(relay.model) + ')',
                            'RelayTag'                  : relay.name + " " + \
                            self._filterstr(relay.manufacturer) + " " + self._filterstr(relay.model),
                            'RelayName'                 : self._filterstr(relay.name),
                            'RelayModel'                : self._filterstr(relay.model),
                            'TrippingElement'           : tripping_data.tripping_element_string,
                            'TripTime'                  : str(tripping_data.trip_time),
                            'CBOpenTime'                : str(relay.cbr_optime),
                            'IFA'                       : str(tripping_data.currents[0]),
                            'IFB'                       : str(tripping_data.currents[1]),
                            'IFC'                       : str(tripping_data.currents[2]),
                            'IFN'                       : str(tripping_data.currents[3]),
                            'Irelay'                    : str(relay.phase_minimum_threshold) +\
                                                          ' (' + str(relay.ground_minimum_threshold) +\
                                                          ')',
                            'RelayPerformanceAssessment': str(int(error_message_id)) 
                        }
        if self.event != None:
            relay_already_inserted , relay_index, fault_index = self._has_already_been_inserted(relay)
            if relay_already_inserted == True:
                insertion_line = self._find_line_of('RelayPerformanceAssessment',
                                                        relay_index)
                if insertion_line >= 0 and \
                str(error_message_id) not in self.output_data[insertion_line] :
                    self._append_string_at(insertion_line,  str(int(error_message_id)))
            else:
                self._write_start('Relay')
                self._write_parameters(relay_parameters)   
                self._write_end('Relay')
                
            # add  the error id also in the fault record
            insertion_line = self._find_line_of('ProtectionPerformanceAssessment',
                                                    fault_index) 
            if insertion_line >= 0 and \
            str(error_message_id) not in self.output_data[insertion_line] and \
            (self.is_ldf == True or error_message_id != self.fault_no_cleared_id or \
            self.event.fault_clearance_time == self.no_trip_time):
                self._append_string_at(insertion_line, str(int(error_message_id)))
    
        
    def write(self):
        '''
        function which writes the whole data array in the file on the disk 
        closing the file object
        '''   
        if self.event != None:     # at least one fault dat set has been written 
            self._write_end('Outage')
            self._write_end('Fault')
        self._write_end('CESITDDCTRESULTS')
        
        for output_data_line in self.output_data:
            # replace the number of violations with the right number
            if "<NumberViolations>" in output_data_line:                
                output_data_line = "<NumberViolations>" + \
                    str(self.number_of_violations) + "</NumberViolations>\n"
            # replace the greater error message id with the relevant error string         
            if 'RelayPerformanceAssessment' in output_data_line:    
                output_data_line = "<RelayPerformanceAssessment>" + self.assessment_result[self.\
                _get_max_error_message_id(output_data_line)] + "</RelayPerformanceAssessment>" 
            if 'ProtectionPerformanceAssessment' in output_data_line: 
                output_data_line = "<ProtectionPerformanceAssessment>" + self.assessment_result[self.\
                _get_max_error_message_id(output_data_line)] + "</ProtectionPerformanceAssessment>"     
            
            self.xmlOutputFile.write(output_data_line)
        
        self.xmlOutputFile.close()
        
        
#===============================================================
#  Auxiliary functions
#===============================================================
    def _has_already_been_inserted(self, relay):
        '''
        function going throw the stored data in the data array checking if the 
        given relay has already been inserted
        '''
        # come back up to the fault start line       
        fault_start_index = self._find_line_of('<Fault>', len(self.output_data), 
                                                        forwad_search = False)
        
        if fault_start_index >= 0:  # if a fault has been found
            # try to find a relay  with the same name and same to/from station
            relay_start_index = self._find_line_of('<Relay>', fault_start_index)
            while relay_start_index >= 0:            
                if relay.name in self.output_data[self._find_line_of('RelayName',\
                                            relay_start_index)]  and\
                relay.from_station in self.output_data[self._find_line_of('FromStation',\
                                            relay_start_index)] \
                and                                                               \
                relay.to_station in self.output_data[self._find_line_of('ToStation',\
                                            relay_start_index)]:
                    return True, relay_start_index, fault_start_index
                relay_start_index = self._find_line_of('<Relay>', relay_start_index+1)
                
        return False, relay_start_index, fault_start_index
        
        
    def  _find_line_of(self, searched_string,  starting_line = 0, 
                                                        forwad_search = True):
        '''
        function finding in the list of strings the first occurance of of 
            the given searched string. 
        
        Args:
        
        searched_string: the string we are looking for in the data array
        starting_line: the starting point of the research in the list of strings
        forwad_search: the direction of the research (from the starting line to
         the end of the string list is the forward direction)
            
        return: the index of the line containing the searched string.
                     If no match is found -1 is returned
        '''
        search_range =  range(starting_line, len(self.output_data)) \
                        if  forwad_search else range(starting_line-1, -1, -1)
        for index in search_range:
            if searched_string in self.output_data[index]:
                return index
        return -1
        
                
    def _write_parameters(self, parameters):
        '''
        write in the data array a dictionary of parameters
        
        Args:
            parameters: the dictionary of values containing setting names and 
            values
        '''           
        for parameter_name, parameter_value in parameters.items():
            self._write_parameter(parameter_name, parameter_value)  
            
            
    def _write_parameter(self, parameter_name, parameter_value): 
        '''
        function writing in the data array the <name>value</name> xml tag
        '''    
        self.output_data.append("<" + parameter_name + ">" + parameter_value + \
                                "</" + parameter_name + ">\n")    
            
            
    def _write_start(self, name):
        '''
        function writing in the data array the <name> xml tag
        '''
        self.output_data.append("<" + name + ">\n")
        
        
    def _write_end(self, name):
        '''
        function writing in the data array the </name> xml tag
        '''
        self.output_data.append("</" + name + ">\n")
        
        
    def _write_string(self, output_string): 
        '''
        function writing in the data array the given string between delimiters 
            + \n
        '''    
        self.output_data.append("<" + output_string + ">\n")
        
        
    def _append_string_at(self, line_number, additional_string): 
        '''
        function writing to the file the given string between delimiters + \n
        '''    
        insertion_index = self.output_data[line_number].rfind('<')
        if insertion_index >= 0:
            self.output_data[line_number] = \
            self.output_data[line_number].replace(
                                self.output_data[line_number][insertion_index:],
                                ", " + additional_string +
                                self.output_data[line_number][insertion_index:])
            
                  
    def _get_max_error_message_id(self, line_string): 
        '''
        function parsing the given string to extract the  greatest
        error message index      
        '''
        start_index = line_string.find('>') + 1
        end_index = line_string.rfind('<')
        return max(int(s) for s in line_string[start_index:end_index].split(', '))  
    
    def _filterstr(self, input):
        '''
        function transforming the input in a string without non ascii characters
        ''' 
        return "".join(filter(lambda x: ord(x)<128, str(input)))
    
    
        
        
        
            
            
            