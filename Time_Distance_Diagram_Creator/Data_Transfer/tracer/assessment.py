'''
Created on 22 Oct 2018

@author: AMB
'''

from collections import namedtuple
from enum import IntEnum
import inspect

from tracer.report_maker import * 



class BreakerID(IntEnum):
    BREAKER1    = 0
    BREAKER2    = 1
    

class Assessment_result_id(IntEnum):   
    '''
    the returned values are the first 6 (OK....LDF_TRIP)
    these values matches the position of the messages in the  
    ASSESSMENT_RESULT dictionary
    '''  
    OK                      = 0
    OVERCURRENT_NEAR_MISS   = 1
    MISSED_COORDINATION     = 2
    INSTANTANEOUS_OVERREACH = 3 
    INSTANTANEOUS_UNDERREACH= 4 
    MISOPERATION            = 5
    FAULT_NOT_CLEARED       = 6 
    LDF_TRIP                = 7
    LDF_ERROR               = 8
    SHCTRACE_ERROR          = 9 
    WARNING                 = 10
    ERROR                   = 20



class EventType(IntEnum):
    FAULT           = 0
    LDF             = 1
        
        
class Assessment(object):
    '''
    class containing the logic which analyzes the short circuit times calculated
     by PSET applying the "Protection Performances Requirements" 
    '''
     
    class OutputDetail(IntEnum):
        DISABLED        = 0
        NORMAL          = 1
        DEBUG           = 2
        VERBOSEDEBUG    = 3
        

    def __init__(self, events, relays, relay_matrix, trip_data_matrix,
                                     interface, settings , output_file , 
                                     grid, list_object):
        '''
        Constructor
        
        Args:
            events: a list of the faults/LDFs which have been applied
            relays: the list of available protective devices as listed in tracer 
                    logic
            relay_matrix: the logic interconnection between the relays as found 
                            by the grid object
            trip_data_matrix: the relay trip data (time and currents) as 
                            calculated applying all faults/LDFs
            interface: reference with the calculation software interface
            settings: a reference to the setting values available in the 
                    graphical interface  
            output_file: full path of the xml file which will contain the report
            grid: reference to the grid object (info regarding the power system
                    layout)
            list_object: reference to an object storing the content of the lists
            containing the available voltage levels, grids, zones, paths etc
        '''
        
        self.events = events
        self.relays = relays
        self.relay_matrix = relay_matrix
        self.interface = interface
        self.trip_data_matrix = trip_data_matrix
        self.settings = settings
        self.output_file = output_file
        self.grid = grid
        self.list_object = list_object
        
        self.output_detail = self.OutputDetail.VERBOSEDEBUG
                          
        self.assessment_result = {Assessment_result_id.OK        : 'Correct Operation',
                     Assessment_result_id.OVERCURRENT_NEAR_MISS  : 'Overcurrent Near Miss',
                     Assessment_result_id.MISSED_COORDINATION    : 'Missed Coordination',
                     Assessment_result_id.INSTANTANEOUS_OVERREACH: 'Instantaneous Over-reach',
                     Assessment_result_id.INSTANTANEOUS_UNDERREACH: 'Instantaneous Under-reach',
                     Assessment_result_id.MISOPERATION           : 'Misoperation',
                     Assessment_result_id.FAULT_NOT_CLEARED      : 'Fault Not Cleared',
                     Assessment_result_id.LDF_TRIP               : 'Trip during LDF',
                     Assessment_result_id.LDF_ERROR              : 'LDF failed',
                     Assessment_result_id.SHCTRACE_ERROR         : 'SHC Trace error'
                    }                    
                          

        Assessment_rule = namedtuple("Assessment_rule", 
                                     "result assessment_function output_function\
                                     event_type error_message_id") 
        
        # create the object sending results to the xml file
        self.report_maker = Report_maker(self.output_file, self.relays, # @UndefinedVariable\
                            self.grid, self.settings, len(trip_data_matrix),\
                            self.assessment_result, self.interface.\
                            get_relay_NO_TRIP_constant(), \
                            Assessment_result_id.FAULT_NOT_CLEARED, self.list_object)  # @UndefinedVariable

        self.assessment_rules = {
            "LDF error":                            Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_LDF,\
                                                        output_function     = self.report_maker.write_ldf,\
                                                        event_type = EventType.LDF, error_message_id = Assessment_result_id.LDF_ERROR),
            "SHC trace error":                      Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_SHCTrace,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.SHCTRACE_ERROR),
            "LDF no trip":                          Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_LDF_no_trip,\
                                                        output_function     = self.report_maker.write_ldf,\
                                                        event_type = EventType.LDF, error_message_id = Assessment_result_id.LDF_TRIP),
            "Max Fault Clearence Time":              Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_max_fault_clearance_time,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.MISOPERATION),
            "Max Trip Time for Near-End faults":     Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_max_trip_time_for_near_end_faults,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.MISOPERATION),
            "Max Trip Time for Far-End faults":      Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_max_trip_time_for_far_end_faults,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.MISOPERATION),
            "Max Reach for Fast Trip Time":          Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_max_reach__for_fast_trip_time,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.INSTANTANEOUS_OVERREACH),
            "Min Reach for Fast Trip Time":          Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_min_reach__for_fast_trip_time,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.INSTANTANEOUS_UNDERREACH),
            "Coordination Margin between Primary/Secondary Protection": Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_coordination_margin_between_primary_secondary_protection,\
                                                        output_function = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.MISSED_COORDINATION),\
            "Security":                              Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_security,\
                                                        output_function = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.MISSED_COORDINATION),\
            "Overcurrent_Near_Miss":                 Assessment_rule(result = Assessment_result_id.OK,\
                                                        assessment_function = self.verify_near_miss_condition,\
                                                        output_function     = self.report_maker.write_fault,\
                                                        event_type = EventType.FAULT, error_message_id = Assessment_result_id.OVERCURRENT_NEAR_MISS)
           # "Coordination Margin between Primary/Backup Protection": Assessment_rule(result = Assessment_result_id.OK,\
           # assessment_function = self.verify_coordination_margin_between_primary_backup_protection, \
           # output_function = self.report_maker.write_fault, event_type = EventType.FAULT, error_message_id = "error")
        }         
      
        
    def run(self):
        '''
        main entry point to run the anlysis
        '''
        #import pydevd
        #pydevd.settrace()
        if self.output_detail >= self.OutputDetail.DEBUG:
            self.interface.print("\n Performing Protection Performances \
                                                    Requirement verification:" )  
        
        self.report_maker.write_study_info()     # initial report info in the xml file
        
        
        for event_index, event in enumerate(self.events):
            if self.is_ldf(event) == False:
                if self.output_detail >= self.OutputDetail.DEBUG:
                    self.interface.print("\n Event #{} - Fault at ".format(event_index) + 
                                self.interface.get_name_of(event.faulted_line) + 
                                " ( at {}% from {}, {} ohm)".
                                format(event.fault_position, \
                                       event.from_station, event. fault_resistance)) 
                    self.interface.print(" Fault type: " + event.type + 
                                    "\t Network status: " + event.network_status)
            else:
                if self.output_detail >= self.OutputDetail.DEBUG:
                    self.interface.print("\n Event #{} - LDF ".format(event_index))
                    
            is_first_event_report = True  
            # iterate throw all all assessment rules                           
            for nassessment_rule_key, nassessment_rule in self.assessment_rules.items():  
                # if it's fault the "LDF no trip" or the "LDF error"  rule desn't apply 
                if (nassessment_rule_key == "LDF no trip" or \
                   nassessment_rule_key == "LDF error")  and \
                   self.is_ldf(event) == False:
                    continue;        # so I skip it
                if self.output_detail >= self.OutputDetail.DEBUG:
                    self.interface.print("\t\tPerforming " + 
                                         nassessment_rule_key + 
                                " verification" )     
                result, affected_relay_index_list = nassessment_rule.\
                                        assessment_function(event, event_index)
                if result != Assessment_result_id.OK:
                    # for any fault we could have multiple relay error reports write the data only once
                    if is_first_event_report == True:      
                        # we use the error message associated to the assessment rule
                        # only if the error message is a generic ERROR
                        if result == Assessment_result_id.ERROR:            
                            result = nassessment_rule.error_message_id     
                        if self.is_ldf(event) == False:
                            if event.fault_clearance_time > \
                                self.interface.get_relay_NO_TRIP_constant()-0.1:
                                result = Assessment_result_id.FAULT_NOT_CLEARED
                        # check if at least one relay has a reportable error
                        # relays with at_load_bus or breaker_failure = true are 
                        # not reportable
                        if self.is_ldf(event) or \
                        any([self.trip_data_matrix[event_index] \
                        [affected_relay_index].at_load_bus == False and \
                        self.trip_data_matrix[event_index] \
                        [affected_relay_index].breaker_failure == False for \
                        affected_relay_index in affected_relay_index_list]):
                            # fault info in the xml file      
                            nassessment_rule.output_function(event_index, event, result)  
                            is_first_event_report = False
                    for affected_relay_index in affected_relay_index_list:
                        error_message = nassessment_rule.error_message_id 
                        if self.trip_data_matrix[event_index][affected_relay_index].trip_time > \
                             self.interface.get_relay_NO_TRIP_constant()-0.1:
                            error_message = Assessment_result_id.FAULT_NOT_CLEARED 
                        # if the relay is at a bus with only loads the errors are
                        # not considered. Also if the relay is set to simulate
                        # a breaker failure
                        if self.trip_data_matrix[event_index]\
                        [affected_relay_index].at_load_bus == False and \
                        self.trip_data_matrix[event_index]\
                        [affected_relay_index].breaker_failure == False:
                            # relay info in the xml file
                            self.report_maker.write_relay(self.relays[affected_relay_index], 
                                                          event, 
                                                          self.trip_data_matrix[event_index][affected_relay_index],
                                                          error_message)  
                            if self.output_detail >= self.OutputDetail.DEBUG:
                                self.interface.print("\t\t\t " + 
                                          self.interface.get_name_of(
                                    self.relays[affected_relay_index].pf_relay) + 
                                          ": " + self.assessment_result.get(error_message))
                        else:
                            if self.output_detail >= self.OutputDetail.DEBUG:
                                if self.trip_data_matrix[event_index]\
                                    [affected_relay_index].at_load_bus == True:
                                    self.interface.print("\t\t\t " + 
                                          self.interface.get_name_of(
                                    self.relays[affected_relay_index].pf_relay) +\
                                    " ignored due to the connection to a load bus.")
                                if self.trip_data_matrix[event_index]\
                                [affected_relay_index].breaker_failure == True:
                                    self.interface.print("\t\t\t " + 
                                          self.interface.get_name_of(
                                    self.relays[affected_relay_index].pf_relay) +\
                                    " ignored due to the breaker failure simulation")                             
                else: 
                    if self.output_detail >= self.OutputDetail.DEBUG:
                        self.interface.print("\t\t\t Result: OK")
                if self.is_ldf(event) == True:  # the ldf has an unique assessment
                    break                       # so I leave the inner loop
            # no event reported so everythink is OK and send  info in the xml file 
            if is_first_event_report == True:     
                nassessment_rule.output_function(event_index, event, result)                            
        self.report_maker.write()            # write the report on the disk
                            
#==========================================================================
#  Verification functions
#==========================================================================
    def verify_LDF(self, ldf, ldf_index):
        '''
        function checking that the LDF has been calculated correctly
        '''
        return (Assessment_result_id.OK, []) if ldf.error == 0 else \
        (Assessment_result_id.ERROR, list())
    
    
    def verify_SHCTrace(self, shctrace, fault_index):
        '''
        function checking that the SHC trace has been calculated correctly
        '''
        return (Assessment_result_id.OK, []) if shctrace.error == 0 else \
        (Assessment_result_id.ERROR, list())
    
    
    def verify_LDF_no_trip(self, ldf, ldf_index):
        '''
        function checking that no relay tripped during the ldf
        '''
        # get all relays indexes belonging to that event (== all relays indexes!) 
        affected_relay_index_list = [index for index in range(0, len(self.relays))]
        
        return_affected_relay_index_list = [relay_index for relay_index \
                                            in affected_relay_index_list \
                if self.trip_data_matrix[ldf_index][relay_index].trip_time <\
                            self.interface.get_relay_NO_TRIP_constant()]
        
        return (Assessment_result_id.OK, []) if len(return_affected_relay_index_list) == 0\
                     else (Assessment_result_id.ERROR, return_affected_relay_index_list) 
     
     
    def verify_max_fault_clearance_time(self, fault, fault_index):
        '''
        function applying the "Max Fault Clearence Time" performance requirement
        '''
        return self.check_max_time(fault, fault_index, \
                                   self.settings['ppMaxClearanceTime'], \
                            (lambda time, allowed_time: time < allowed_time),\
                            reverse_condition_for_remote = False)
                
    
    def verify_max_trip_time_for_near_end_faults(self, fault, fault_index):
        '''
        function applying the "Max Trip Time for Near-End faults" performance 
        requirement
        '''
        return self.check_max_time(fault, fault_index, \
                                   self.settings['ppMaxNearEndTime'], \
                            (lambda time, allowed_time: time < allowed_time), \
                            consider_remote_relay = True \
                            if fault.fault_position > 50 else False,
                            consider_local_relay = True \
                            if fault.fault_position <= 50 else False,\
                            reverse_condition_for_remote = False)
    
    
    def verify_max_trip_time_for_far_end_faults(self, fault, fault_index):
        '''
        function applying the "Max Trip Time for Far-End faults" performance 
        requirement
        '''
        return self.check_max_time(fault, fault_index, \
                                   self.settings['ppMaxFarEndTime'], \
                            (lambda time, allowed_time: time < allowed_time),
                            consider_remote_relay = True \
                            if fault.fault_position <= 50 else False,
                            consider_local_relay = True \
                            if fault.fault_position > 50 else False,\
                            reverse_condition_for_remote = False)
    
    
    def verify_max_reach__for_fast_trip_time(self, fault, fault_index):
        '''
        function applying the Max Reach for Fast Trip Time"  performance 
        requirement
        '''
        if self.settings['ppMaxFastTrippingTime'] <= 0:   # function disabled
            return (Assessment_result_id.OK, [])
        if fault.fault_position > self.settings['ppMaxInstReach'] or \
        fault.fault_position < 100 - self.settings['ppMaxInstReach']:        
            return self.check_max_time(fault, fault_index,
                            self.settings['ppMaxFastTrippingTime'], \
                            (lambda time, fast_trip_time: time > fast_trip_time), \
                            reverse_condition_for_remote = False,\
                            consider_remote_relay = True if \
                            fault.fault_position < 100 - self.settings['ppMaxInstReach']\
                            else False, \
                            consider_local_relay = True if \
                            fault.fault_position > self.settings['ppMaxInstReach'] \
                            else False)
        else:
            if self.output_detail >= self.OutputDetail.DEBUG:
                self.interface.print("\t\t\t Fault position under the max inst reach. Rule not applicable" )
            return Assessment_result_id.OK, [] # we leave, this rule is not valid there
    
    
    def verify_min_reach__for_fast_trip_time(self, fault, fault_index):
        '''
        function applying the Min Reach for Fast Trip Time"  performance 
        requirement. This is an additional/auxiliary requirement 
        '''
        if self.settings['ppMaxFastTrippingTime'] <= 0:   # function disabled
            return (Assessment_result_id.OK, []) 
        if fault.fault_position <= self.settings['ppMaxInstReach'] and \
        fault.fault_position > 100 - self.settings['ppMaxInstReach']:
            
            return self.check_max_time(fault, fault_index, 
                            self.settings['ppMaxFastTrippingTime'], \
                            (lambda time, fast_trip_time: time < fast_trip_time), \
                            reverse_condition_for_remote = False)
        else:
            if self.output_detail >= self.OutputDetail.DEBUG:
                self.interface.print("\t\t\t Fault position over the max inst reach. Rule not applicable" )
            return Assessment_result_id.OK, [] # we leave, this rule is not valid there
    
    
    def verify_coordination_margin_between_primary_backup_protection(self, \
                                                            fault, fault_index):
        '''
        function verifying that the minimum margin is present between the 
        primary and the backup relays present in the same cubicle 
        NOTE: at the moment not used
        '''
        if self.settings['ppMinCoordinationTimeMargin'] <= 0:   # function disabled
            return (Assessment_result_id.OK, [])
        # get the line zone 1 relay
        affected_relay_list = [relay for relay in self.relay_matrix \
                               if relay.protected_line == fault.faulted_line and\
                                relay.step == 1]  
        # get their indexes in the self.relays array (from pset_logic)
        affected_relay_index_list = self.get_relay_index_list(affected_relay_list)   
        backup_relay_trip_time_list = [self.trip_data_matrix[fault_index][relay_index].trip_time \
                                       for relay_index in affected_relay_index_list \
                                       if self.relays[relay_index].is_backup_relay == True]   
        main_relay_trip_time_list = [self.trip_data_matrix[fault_index][relay_index].trip_time \
                                     for relay_index in affected_relay_index_list \
                                     if self.relays[relay_index].is_backup_relay == False]
        if len(backup_relay_trip_time_list) > 0 and len(main_relay_trip_time_list) > 0 :  # only if both list are not void
            min_time_backup = min(backup_relay_trip_time_list) 
            max_time_primary = max(main_relay_trip_time_list) 
            delta_time = min_time_backup - max_time_primary
            return (Assessment_result_id.OK, []) if delta_time > \
                self.settings['ppMinCoordinationTimeMargin']  \
                else (Assessment_result_id.MISOPERATION, affected_relay_index_list)
        else:
            return (Assessment_result_id.OK, []) 
            
            
    def verify_coordination_margin_between_primary_secondary_protection(self, fault, fault_index):  
        '''
        function verifying the Coordination Minimum Time Interval between the 
        main and the secondary protections(s)
        '''
        # the list of relays with problems which will be returned
        return_affected_relay_index_list = []  
        # the minimum delta time between relays to have selectivity                                 
        minimum_delta_time = self.settings['ppMinCoordinationTimeMargin'] 
        if minimum_delta_time <= 0:   # function disabled
            return (Assessment_result_id.OK, []) 
        # iterate only for the number specified as simulation depth
        for step in range (self.settings['ppProtSimDepthValue']):       
            # get the line zone "step" relay 
            affected_relay_list = [relay for relay in self.relay_matrix \
                                if relay.protected_line == fault.faulted_line \
                                and self.interface.get_name_of(self.interface.\
                                get_relay_protected_item_of(relay.pf_relay)) not in \
                                fault.disconnected_elements_names\
                                and relay.step == step]  
            # get their indexes in the self.relays array (from pset_logic)
            affected_relay_index_list = self.get_relay_index_list(affected_relay_list) 
            error = Assessment_result_id.ERROR
            primary_min_trip_time = min([self.trip_data_matrix[fault_index]\
            [affected_relay_index].trip_time for affected_relay_index \
            in affected_relay_index_list]) if len(affected_relay_index_list) > 0 \
            else self.interface.get_relay_NO_TRIP_constant() 
            # iterate throw all relay_matrix relays
            for index, affected_relay_index in enumerate(affected_relay_index_list):  
                # if the relay is at load bus or has been disabled don't consider it
                if self.trip_data_matrix[fault_index][affected_relay_index].\
                at_load_bus == True or \
                self.trip_data_matrix[fault_index][affected_relay_index].\
                breaker_failure == True:
                    continue
                # and look in the secondary relay list 
                for secondary_relay in affected_relay_list[index].selective_relay_links:
                    if self.interface.get_name_of(self.interface.\
                    get_relay_protected_item_of(secondary_relay)) in \
                    fault.disconnected_elements_names:
                        continue                   
                    primary_trip_time = self.trip_data_matrix[fault_index]\
                                                [affected_relay_index].trip_time
                    secondary_trip_time = self.trip_data_matrix[fault_index]\
                            [self.get_relay_index(secondary_relay)].trip_time
                    # verify the secondary relay trip
                    if secondary_trip_time > \
                    self.interface.get_relay_NO_TRIP_constant()-0.1:
                        error = Assessment_result_id.MISOPERATION
                    if primary_min_trip_time + minimum_delta_time > \
                    secondary_trip_time:
                        # add the relay in the list of relays with problems
                        return_affected_relay_index_list.append(affected_relay_index)
                        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                            self.interface.print("\t\t\t Primary relay: " +  \
                            self.interface.get_name_of(self.relays[\
                            affected_relay_index].pf_relay) + "(" + str(primary_trip_time) + " s)")
                            self.interface.print("\t\t\t Secondary relay: " +  \
                            self.interface.get_name_of(self.relays[\
                            self.get_relay_index(secondary_relay)].pf_relay) + \
                            "(" + str(secondary_trip_time) + " s)")
                                                 
                            
                            
                            
        return (Assessment_result_id.OK, []) if len(return_affected_relay_index_list) == 0\
                     else (error, return_affected_relay_index_list) 
    
    
    def verify_security(self, fault, fault_index):
        '''
        function verifying that only the relay supposed to trip are tripping
        for the given fault
        '''
        # the minimum delta time between relays to have selectivity                                 
        minimum_delta_time = self.settings['ppMinCoordinationTimeMargin'] 
        if minimum_delta_time <= 0:   # function disabled
            return (Assessment_result_id.OK, [])
        # collect the indexes all the relays which are tripped  
        tripped_relay_index_list = [relay_index for relay_index in range(len(self.relays))\
                                    if self.trip_data_matrix[fault_index][relay_index].trip_time <\
                                     self.interface.get_relay_NO_TRIP_constant()-0.01]
        # get the relays supposed to trip in any step 
        all_supposed_to_trip_relay_list = [relay for relay in self.relay_matrix \
                                    if relay.protected_line == fault.faulted_line]
        all_supposed_to_trip_relay_index_list = self.get_relay_index_list(\
                                                all_supposed_to_trip_relay_list)
        # add the secondary relays (i.e. selective relays belonging to 
        # branches without further branches)
        all_supposed_to_trip_relay_index_list = self.add_selective_relays_to(\
                all_supposed_to_trip_relay_list, all_supposed_to_trip_relay_index_list)  
        #remove relays disabled by the breaker failure
        all_supposed_to_trip_relay_index_list = [index for index in \
                all_supposed_to_trip_relay_index_list if \
                self.trip_data_matrix[fault_index][index].breaker_failure == False]
        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
            self.interface.print("\t\t\t All relay(s) supposed to trip:")
            for supposed_to_trip_relay_index in all_supposed_to_trip_relay_index_list:
                self.interface.print("\t\t\t\t" + self.interface.get_name_of(\
                         self.relays[supposed_to_trip_relay_index].pf_relay))
            
        return_affected_relay_index_list = []
        for step in range(1, self.settings['ppProtSimDepthValue']+1):
            # collect all the relays which protect the given line for the given step
            supposed_to_trip_relay_list = [relay for relay in self.relay_matrix \
                                    if relay.protected_line == fault.faulted_line \
                                    and relay.step == step]   
            # void list...skip
            if len(supposed_to_trip_relay_list) == 0:
                continue
                 
            #get their indexes in the rip data matrix
            supposed_to_trip_relay_index_list = self.get_relay_index_list(\
                                                    supposed_to_trip_relay_list)         
            #make unique and remove relays disabled by the breaker failure
            supposed_to_trip_relay_index_list = list(set([ index for index in \
                supposed_to_trip_relay_index_list if \
                self.trip_data_matrix[fault_index][index].breaker_failure == False]))
            
            # add the secondary relays (i.e. selective relays belonging to 
            # branches without further branches)
            supposed_to_trip_relay_index_list = self.add_selective_relays_to(\
                supposed_to_trip_relay_list, supposed_to_trip_relay_index_list)  
            
            supposed_to_trip_max_time = max([self.trip_data_matrix[fault_index]\
                                  [relay_index].trip_time for relay_index \
                                  in supposed_to_trip_relay_index_list]) \
                                if len(supposed_to_trip_relay_index_list) > 0 else\
                                self.interface.get_relay_NO_TRIP_constant()
            
            # compare the two lists and the relays tripped but not supposed to trip
            # will be returned
            return_affected_relay_index_list += [tripped_relay_index for tripped_relay_index\
                                                in tripped_relay_index_list\
                                                if tripped_relay_index not in \
                                                supposed_to_trip_relay_index_list and\
                                                self.trip_data_matrix[fault_index]\
                                                [tripped_relay_index].trip_time <\
                                                supposed_to_trip_max_time + minimum_delta_time]
            #make unique 
            return_affected_relay_index_list = list(set(return_affected_relay_index_list))
            
            if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                self.interface.print("\t\t\tStep {} Min allowed time: ".format(step) + \
                    str(round(supposed_to_trip_max_time + minimum_delta_time,3)))
                self.interface.print("\t\t\t Relay(s) supposed to trip:")
                for supposed_to_trip_relay_index in supposed_to_trip_relay_index_list:
                    trip_time = self.trip_data_matrix[fault_index]\
                                    [supposed_to_trip_relay_index].trip_time
                    self.interface.print("\t\t\t\t" + self.interface.get_name_of(\
                        self.relays[supposed_to_trip_relay_index].pf_relay) + "(" + \
                                                    str(trip_time) + " s)")
                self.interface.print("\t\t\t Other relay(s) tripping too fast:")
                for affected_relay_index in return_affected_relay_index_list:
                    trip_time = self.trip_data_matrix[fault_index]\
                                                [affected_relay_index].trip_time
                    self.interface.print("\t\t\t\t" + self.interface.get_name_of(\
                        self.relays[affected_relay_index].pf_relay) + "(" + \
                                                    str(trip_time) + " s)")
                    
                    
        return (Assessment_result_id.OK, []) if len(return_affected_relay_index_list) == 0\
                     else (Assessment_result_id.ERROR, return_affected_relay_index_list) 
           
        
    def verify_near_miss_condition(self, fault, fault_index):    
        '''
        function verifying the near miss condition for the line main 
        protection(s) both the max phase current and the zero sequnece current
        are compared with the phase and the ground thresholds
        '''
        # get the line zone 1 relay
        affected_relay_list = [relay for relay in self.relay_matrix \
                               if relay.protected_line == fault.faulted_line and\
                               relay.step == 1]  
        # get their indexes in the self.relays array (from pset_logic)
        affected_relay_index_list = self.get_relay_index_list(affected_relay_list)  
        
        return_affected_relay_index_list = []  # the list of relays with problems which will be returned
        for affected_relay_index in affected_relay_index_list:
            # calculate the phase and the ground difference
            phase_current_difference = max(self.trip_data_matrix[fault_index]\
                [affected_relay_index].currents[0:2])-self.relays[affected_relay_index].phase_minimum_threshold 
            ground_current_difference = self.trip_data_matrix[fault_index]\
                [affected_relay_index].currents[3]-self.relays[affected_relay_index].phase_minimum_threshold  
            # if there is not enough difference between the measured I and the threshold
            if  phase_current_difference > 0 and phase_current_difference < \
                    self.relays[affected_relay_index].phase_minimum_threshold * \
                    self.settings['ppNearMissOCValue']/100.0 or \
                ground_current_difference > 0 and ground_current_difference < \
                self.relays[affected_relay_index].ground_minimum_threshold * \
                    self.settings['ppNearMissOCValue']/100.0 : 
                # add the relay in the list of relays with problems 
                return_affected_relay_index_list.append(affected_relay_index) 
            if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                self.interface.print("\t\t\t" + self.interface.get_name_of(\
                self.relays[affected_relay_index].pf_relay))
                self.interface.print("\t\t\t\t" + "  Phase Threshold= "  + \
                str(self.relays[affected_relay_index].phase_minimum_threshold) + \
                " A      Ground Threshold= " + 
                str(self.relays[affected_relay_index].ground_minimum_threshold) + " A")
                self.interface.print("\t\t\t\t" + "  Max Phase I= "  + \
                str(max(self.trip_data_matrix[fault_index][affected_relay_index].currents[0:2]))
                + " A     Ground I= "  + \
                str(self.trip_data_matrix[fault_index][affected_relay_index].currents[3])
                + " A")                                                       
                                                                              
        return (Assessment_result_id.OK, []) if len(return_affected_relay_index_list) == 0\
                     else (Assessment_result_id.ERROR, return_affected_relay_index_list) 
                
        
        
            
#===============================================================
#    Auxiliary functions
#===============================================================
    
    def add_selective_relays_to(self, grid_relay_list, grid_relay_index_list):
        '''
        function adding to the given grid relay list (link relays) the
        selective relays in case they are not already part of list
        (it happens when the branch is generator)
        it returns the index list of the merged relays
        '''
        return_index_list = grid_relay_index_list
        for grid_relay in grid_relay_list:
            selective_relay_index_list = [self.get_relay_index(pf_relay) \
                        for pf_relay in grid_relay.selective_relay_links]    
            return_index_list = list(set(return_index_list + \
                                          selective_relay_index_list))
        return return_index_list   
            
            
    def get_relay_index_list(self, grid_relay_list):
        '''
        function returning a list of relay indexes (indexes in the list of the 
        available tracer-logic relays) of the grid relays passed as parameter 
        '''
        return  [logic_relay_index for grid_relay in grid_relay_list \
                  for logic_relay_index,logic_relay in enumerate(self.relays)\
                   if grid_relay.pf_relay == logic_relay.pf_relay]
    
    
    def get_relay_index(self, pf_relay):
        '''
        function returning the relay index (index in the list of the available 
        tracer-logic relays) of the PowerFactory relay passed as parameter 
        '''
        # iterate throw all tracer logic relays
        for relay_index, relay in enumerate(self.relays):
            # if the tracer logic relay refers to the pf_relay passed as a param  
            if relay.pf_relay == pf_relay:                  
                return relay_index                     # return relevant index
            
    
    def check_max_time(self, fault, fault_index, allowed_time, comparison,
                       reverse_condition_for_remote = True,
                       consider_remote_relay = True,
                       consider_local_relay = True):
        '''
        function checking that for the given fault the trip time is not 
        greater/not smaller than the given allowed_time
        the not greater/not smaller behavior is defined by the comparison 
        function passed as last parameter 
        '''
        if allowed_time <= 0:   # function disabled
            return (Assessment_result_id.OK, [])  
        affected_local_relay_index_list = []
        affected_remote_relay_index_list = []
        # get the line zone 1 remote relays 
        affected_remote_relay_list = [relay for relay in self.relay_matrix \
                    if relay.protected_line == fault.faulted_line and\
                    relay.step == 1 and self.is_relay_local(relay, fault) == False and\
                    self.interface.is_out_of_service(relay.pf_relay) == False and \
                        consider_remote_relay == True]
        if len(affected_remote_relay_list) == 0: 
            if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                self.interface.print("\t\t\t Use 2nd step remote relays.") 
            affected_remote_relay_index_list = [self.get_relay_index(selective_relay) \
                                          for relay in self.relay_matrix \
                        if relay.protected_line == fault.faulted_line and\
                        relay.step == 1 and self.is_relay_local(relay, fault) == False and\
                        consider_remote_relay == True for selective_relay\
                                         in relay.selective_relay_links]
            # make the relay index list unique
            affected_remote_relay_index_list = list(set(affected_remote_relay_index_list)) 
        else:   
            # get their indexes in the self.relays array (from pset_logic)
            affected_remote_relay_index_list = \
                            self.get_relay_index_list(affected_remote_relay_list)
                            
        # get the line zone 1 local relays 
        affected_local_relay_list = [relay for relay in self.relay_matrix \
                    if relay.protected_line == fault.faulted_line and \
                    relay.step == 1 and self.is_relay_local(relay, fault) == True and \
                    self.interface.is_out_of_service(relay.pf_relay) == False and\
                        consider_local_relay == True]
        #if the breaker failure option is set get the 2nd step relays 
        if len(affected_local_relay_list) == 0: 
            if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                self.interface.print("\t\t\t Use 2nd step local relays.") 
            affected_local_relay_index_list = [self.get_relay_index(selective_relay) \
                                          for relay in self.relay_matrix \
                        if relay.protected_line == fault.faulted_line and\
                        relay.step == 1 and self.is_relay_local(relay, fault) == True and\
                        consider_local_relay == True for selective_relay\
                                         in relay.selective_relay_links]
            # make the relay index list unique
            affected_local_relay_index_list = list(set(affected_local_relay_index_list))
            # get all relays indexes in the self.relays array (from pset_logic)                     
        else:
            # get the local relay indexes in the self.relays array (from pset_logic)
            affected_local_relay_index_list = \
                            self.get_relay_index_list(affected_local_relay_list)   
        # get their indexes in the self.relays array (from pset_logic)     
        affected_relay_index_list = affected_local_relay_index_list +\
                                                affected_remote_relay_index_list
                                         
        #debug output
        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
            self.interface.print("\t\t\t Affected local relay:")                    
            for affected_local_relay_index in affected_local_relay_index_list:
                self.interface.print("\t\t\t\t" + self.interface.get_name_of(\
                         self.relays[affected_local_relay_index].pf_relay) + "(" + \
                         str(self.trip_data_matrix[fault_index]\
                        [affected_local_relay_index].trip_time) + " s)")
            self.interface.print("\t\t\t Affected remote relay:")                    
            for affected_remote_relay_index in affected_remote_relay_index_list:
                self.interface.print("\t\t\t\t" + self.interface.get_name_of(\
                         self.relays[affected_remote_relay_index].pf_relay) + "(" + \
                         str(self.trip_data_matrix[fault_index]\
                        [affected_remote_relay_index].trip_time) + " s)")
                     
        # if at least one relay is listed
        if affected_relay_index_list != None and len(affected_relay_index_list) > 0:                                                                                      
            # get the min trip time
            min_local_time = min(self.trip_data_matrix[fault_index][relay_index].trip_time \
                       for relay_index in affected_local_relay_index_list) \
                       if len(affected_local_relay_index_list) > 0 else 0
            min_remote_time = min(self.trip_data_matrix[fault_index][relay_index].trip_time \
                       for relay_index in affected_remote_relay_index_list)  \
                       if len(affected_remote_relay_index_list) > 0 else 0   
            if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
                self.interface.print("\t\t\t Trip Time: local " + str(min_local_time) +
                "(s)   remote "  + str(min_remote_time) + "(s)    Allowed Time (s): " + str(allowed_time))
                self.interface.print("\t\t\t Rule :" + inspect.getsource(comparison))                    
            # the list of relays with problems which will be returned
            condition_for_remote = False if reverse_condition_for_remote else True
            return_affected_relay_index_list = [affected_relay_index
                        for affected_relay_index in affected_relay_index_list  
                if (comparison(min_local_time, allowed_time) != True and 
                self.get_relay_index(self.relays[affected_relay_index].pf_relay) in \
                affected_local_relay_index_list)
                or (comparison(min_remote_time, allowed_time) != condition_for_remote and 
                self.get_relay_index(self.relays[affected_relay_index].pf_relay) in \
                affected_remote_relay_index_list)]
            return (Assessment_result_id.OK, []) if len(return_affected_relay_index_list) == 0\
                 else (Assessment_result_id.ERROR, return_affected_relay_index_list)
        else:    
            return (Assessment_result_id.OK, []) 
     
    
    def is_relay_local(self, relay, fault): 
        '''
        function returning true if the given relay is located in the "from" cubicle
        of the line where it's located 
        '''
        logic_relay = self.relays[self.get_relay_index(relay.pf_relay)]
        bus_name = self.interface.get_name_of(\
                            self.interface.get_branch_bus1_of(fault.faulted_line)) \
                    if fault.reference_breaker == BreakerID.BREAKER1 else \
                    self.interface.get_name_of(\
                            self.interface.get_branch_bus2_of(fault.faulted_line))
        
        return True if bus_name in logic_relay.from_station else False
    
            
    def is_ldf(self, fault):
        ''' 
        function checking the kind of givne "fault" (which can be also a LDF...))
        '''
        return False if 'Fault' in  type(fault).__name__ else True