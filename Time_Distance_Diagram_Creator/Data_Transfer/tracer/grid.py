'''
Created on 16 Oct 2018

@author: AMB
'''

from collections import namedtuple
from enum import IntEnum
import copy

from tracer.branch import *
from tracer.measurement_path import MeasurementPath  # @UnresolvedImport


class Grid(object):
    '''
    class managing all grid data, collecting the branches and calculating the 
    relay interconnection matrix
    '''

    Relay_link = namedtuple("Relay_link",
                            "pf_relay selective_relay_links monitored_relay_links \
                             protected_line step branch_list")
    
    class OutputDetail(IntEnum):
        DISABLED        = 0
        NORMAL          = 1
        DEBUG           = 2
        VERBOSEDEBUG    = 3

    def __init__(self, interface):
        '''
        Constructor
        '''
        self.output_detail = self.OutputDetail.VERBOSEDEBUG
        self.interface = interface
        # dictionary containing for each busbar the list of the branches
        self.busbar_list = {}
        # list of all branches available
        self.branch_list = []
        # matrix of the relays roles (primary, secondary, tertiary protection)
        self.relay_matrix = []
        # dictionary which states for each busbar for a given network configuration
        # if the busbar is a load busbar or not 
        self.load_busbars = {}
        
        # flag storing the result of the _follow_branches_bus_of function
        self.relay_link_added = False
        
        self.create_lists()
        
        
#===========================================================================
#   Branch list creation functions
#===========================================================================
    
    def create_lists(self):
        '''
        function getting throw all lines available in the active project and 
        creating a set of branch objects which are stored inside the branch_list 
        class object
        in parallel also the busbar_list list is filled with the node objects 
        which connects the branches
        it calls the <_follow_lines_bus_of> function
        '''        
        #initialize class variables
        self.busbar_list.clear()
        self.branch_list.clear()
        
        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
            self.interface.print("\tBranches\n")
        all_busbars = self.interface.get_busbars()
        for busbar in all_busbars:  # go throw all busbars
            # this busbar has already been inserted as part of a branch or as
            # terminal bus
            busbar_connections = self.interface.get_bus_connections_of(busbar) 
            for busbar_connection in busbar_connections:   
                # this connection has already been inserted as part of a branch 
                # or is disabled
                if self._has_already_been_collected(busbar_connection) == True \
                or self.interface.is_out_of_service(busbar_connection) == True:   
                    continue                                    # so I ignore it                                    
                branch_name = self.interface.get_name_of(busbar_connection)                                  
                for secondary_side in [self.interface.Side.Side_2,\
                                       self.interface.Side.Side_3]:
                    if secondary_side == self.interface.Side.Side_3 and\
                    self.interface.get_connection_number(busbar_connection) < 3:
                        continue
                    # return variables
                    element_list = []
                    relay_list = [[],[],[]] 
                    terminal_bus_list =[None, None, None, None]
                    
                    element_list.append(busbar_connection)
                    # get the element relays
                    self._get_element_relays_of(busbar_connection, relay_list) 
                    # collect the breakers associated to the relays
                    element_list.extend(self.interface.get_branch_breakers_of(\
                                                            busbar_connection))                   
                    for side in  [self.interface.Side.Side_1, 
                                  secondary_side]: 
                        # recursive function     
                        branch_name = self._follow_lines_bus_of(
                                            line = busbar_connection,
                                            busbar = None, 
                                            side = side,
                                            branch_name = branch_name,
                                            element_list = element_list,
                                            terminal_bus_list = terminal_bus_list,
                                            relay_list = relay_list) 
                        # check that terminal bus list first position is filled
                        # it could happen that for a branch with only one terminal bus
                        # that terminal has been set in the 2 position
                        if terminal_bus_list[0] is None and \
                           terminal_bus_list[1] is not None:
                            terminal_bus_list[0] = terminal_bus_list[1]
                            terminal_bus_list[1] = None  
                        if side == self.interface.Side.Side_3:
                                terminal_bus_list[1], terminal_bus_list[2] = \
                                terminal_bus_list[2], terminal_bus_list[1] 
                    #create the new branch object   
                    new_branch = Branch(name = branch_name,  # @UndefinedVariable
                                        pf_element_list = element_list,
                                        relay_list1 = relay_list[0],
                                        relay_list2 = relay_list[side],
                                        terminal_bus_list = terminal_bus_list,
                                        interface = self.interface)       
                    self.branch_list.append(new_branch)       # and add it the list               
                    # fill the terminal bus object with the new branch info and add 
                    # it in the list
                    for terminal_bus in terminal_bus_list:    
                        if terminal_bus != None:
                            if terminal_bus in self.busbar_list :
                                self.busbar_list[terminal_bus].append(new_branch)
                            else:
                                self.busbar_list.update({terminal_bus:[new_branch]})
                         
                    # output code for debugging
#                     if self.output_detail >= self.OutputDetail.VERBOSEDEBUG: 
#                         self.interface.print("\n\t\t     Branch:" + new_branch.name)                      
#                         for collected_element in new_branch.pf_element_list:
#                             self.interface.print("\t\t%r " % \
#                             (self.interface.get_full_name_of(collected_element),  ))
#                         for i in range(0,2):
#                             self.interface.print("\t\tRelay list {}:".format(i))
#                             for relay in new_branch.get_relay_list(i):
#                                 self.interface.print("\t\t%r " % \
#                                     (self.interface.get_full_name_of(relay),  )) 
#                         self.interface.print("\t\tTerminal busbars:")
#                         for j in range(0,2): 
#                             if len(terminal_bus_list) > j and \
#                                             terminal_bus_list[j] != None:
#                                 self.interface.print("\t\t%r " % \
#                             (self.interface.get_full_name_of(terminal_bus_list[j]), )) 
        # output code for debugging
#         if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
#             self.interface.print("\nBus bars: \n")
#             for busbar_key, busbarvalue in self.busbar_list.items():
#                 self.interface.print("\t" +
#                                  self.interface.get_full_name_of(busbar_key), )
#                 self.interface.print("\t\t Connessions:")
#                 for connection in busbarvalue:
#                     self.interface.print("\t\t" + connection.name, )
                
#------------------------------------------------------------------------------ 
# Branch list creation Recursive function
#------------------------------------------------------------------------------      
     
    def _follow_lines_bus_of(self, line, busbar, side, branch_name, element_list, 
                                                terminal_bus_list, relay_list):
        '''
        it collects the lines/ bus of the following couples of busbar /line
        Recursive function called by <create_branch_list>
        
        Args:
            line: the line from which we are looking for the next busbar
            busbar: the last busbar
            side : the line side at which we get the relays
            branch_name: return string containg the name of the new branch 
                        (= collection of all connected elements names) 
            element_list: return list containing all the elements contained by
                         the branch
            terminal_bus_list: return list containing the busses which delimit 
                                the branch
            relay_list: list of all relays
        Returns:
            the name of the branch     
        '''
        # avoid to try to get Side_2 if not available
        first_side = side if (side == self.interface.Side.Side_1 or \
                          (self.interface.get_connection_number(line) > 1 and
                           side == self.interface.Side.Side_2) or \
                          (self.interface.get_connection_number(line) > 2 and
                           side == self.interface.Side.Side_3)) \
                          else self.interface.Side.Side_1 
        # try to get the busbar connected at the given "side"                     
        connected_busbar = self.interface.get_branch_busses_of(line, first_side)   
        # the found busbar is the same from which we come from 
        if connected_busbar == busbar:   # we get the other one if available
            second_side = self.interface.Side.Side_2\
            if first_side == self.interface.Side.Side_1 and \
            self.interface.get_connection_number(line) > 1 else \
            self.interface.Side.Side_1
            connected_busbar = self.interface.get_branch_busses_of(line, second_side) 
        # the found busbar has already been collected                    
        if connected_busbar in element_list:                       
            connected_busbar = None                       #  force to leave
        # just in case that the terminal is not a "final terminal"
        if  connected_busbar != None and \
                    self._is_terminal_bus(connected_busbar) ==  False:  
            # add the bus bar in the list of already collected elements    
            element_list.append(connected_busbar)                                               
            connected_element_list = self.interface.get_bus_connections_of(connected_busbar) 
            # iterate throw all connections of the busbar
            for connected_element in connected_element_list:                                    
                if connected_element == line or \
                self.interface.is_out_of_service(connected_element) == True:
                    continue      
                # get the element relays               
                self._get_element_relays_of(connected_element,  relay_list)  
                #if not already inserted                                                        
                if connected_element not in element_list:    
                    # add the connection element in the list of already 
                    #collected elements                   
                    element_list.append(connected_element)    
                    # collect the breakers associated to the relays
                    element_list.extend(self.interface.get_branch_breakers_of(\
                                                        connected_element))
                    # fill the name string of the branch element                  
                    branch_name = branch_name + ' - ' +   \
                            self.interface.get_name_of(connected_element)  
                # next iteration                 
                branch_name = self._follow_lines_bus_of( line = connected_element,
                                                        busbar = connected_busbar,
                                                        side = side, 
                                                        branch_name = branch_name, 
                                                        element_list = element_list, 
                                                        terminal_bus_list = terminal_bus_list, 
                                                        relay_list = relay_list)  
        elif connected_busbar != None:
            # if not already stored.....
            if connected_busbar not in terminal_bus_list:  
                # the busbar is  a "final terminal" so it's a true terminal and
                # I add it in the list                    
                terminal_bus_list[side] = connected_busbar                      
        return  branch_name                 
    
    
        
#------------------------------------------------------------------------------ 
#   Branches list creation auxilary functions
#------------------------------------------------------------------------------    
   
    def _get_element_relays_of(self, element, relay_list):
        '''
        function returning the relays present in the cubibles of the given 
        element at the given side
        Args:
            element: the network element from which we are trying to retrieve 
                    the relays
            relay_list: the list where the found relays are returned 
                    (2 places, one for each side)
        '''
        for side in [self.interface.Side.Side_1, 
                     self.interface.Side.Side_2,
                     self.interface.Side.Side_3]:
            #get relays at both side only if 2 sides are available
            if side == self.interface.Side.Side_1 or \
                          (self.interface.get_connection_number(element) > 1 and
                           side == self.interface.Side.Side_2) or \
                          (self.interface.get_connection_number(element) > 2 and
                           side == self.interface.Side.Side_3):
                rel_list = self.interface.get_branch_relays_of(element, side)
                if len(rel_list) > 0:            # if I got a relay list
                    relay_list[side] = rel_list  # replace the relays list field
   
   
    def _get_relay_breakers_of(self, relay_list):
        '''
        functions returning all breakers present in the same cubicle of the given
        relays. It's used to collect the breakers not graphically present in the
        one line diagram.
        '''
        return_breaker_list = []
        for side in [self.interface.Side.Side_1,\
                     self.interface.Side.Side_2,\
                     self.interface.Side.Side_3]:
            for relay in relay_list[side]:
                return_breaker_list.extend(
                    self.interface.get_relay_breaker_of(relay))
#                 cubicle = self.interface.get_relay_cubicle_of(relays[side])
#                 if cubicle != None: 
#                     return_breaker_list.append(self.interface.\
#                                              get_content(cubicle, '*ElmCoup'))
                    
        return_breaker_list = list(set(return_breaker_list)) 
        return return_breaker_list
    
        
    def _is_terminal_bus(self, busbar):
        '''
        function checkingf if the given busbar is completing a branch.
        It's completing a branch when it has more than 2 connections, or one 
        single connection
        or when one the two connections host at least one relay
        '''
        connected_element_list = self.interface.get_bus_connections_of(busbar) 
        number_of_connected_elements = len(connected_element_list)
        # terminal bus is the number of connection is 1 or greater than 2
        if  number_of_connected_elements> 2 or  number_of_connected_elements == 1: 
            return True
        else:
            # just 2 connections! check if there are some relays
            for nconnected_element in connected_element_list:        
                side = self.interface.Side.Side_1  \
                    if self.interface.get_branch_busses_of(nconnected_element,\
                                 self.interface.Side.Side_1) == busbar else \
                            (self.interface.Side.Side_2 \
                            if self.interface.get_connection_number(nconnected_element) > 1 \
                            else self.interface.Side.Side_1)            
                relay_list = self.interface.get_branch_relays_of(nconnected_element, side)
                # if relays are available and at least one them is active
                if len(relay_list) > 0 and\
                    any([self.interface.is_out_of_service(nrelay)==False \
                         for nrelay in relay_list]):
                    return True                     # it's a "final terminal"
            return False
        
        
    def _has_already_been_collected(self, element):
        '''
        is the given element already present in the collected elements in the 
        existing branches?
        '''
        # check if it's between the terminal busbars
        if element in self.busbar_list:     
            return True
        for branch in self.branch_list:
            # check if it's between one of the branch elements
            if element in branch.pf_element_list:  
                return True
        return False
                
        
    def get_branch_of(self, element):
        '''
        function returning the branch at which the given element belongs
        If the given element is a busbar all branches connected to that busbar 
        are returned in a list  
        ''' 
        if element != None:
            class_name = self.interface.get_class_name_of(element)  
            if class_name == 'ElmTerm':
                return [branch for branch in self.branch_list \
                                if element in branch.terminal_bus_list]
            elif class_name == 'ElmRelay':
                for branch in self.branch_list:
                    # check if it's between one of the relays elements
                    if element in branch.relay_list1 or \
                                element in branch.relay_list2:  
                        return branch 
            else:
                for branch in self.branch_list:
                    # check if it's between one of the branch elements
                    if element in branch.pf_element_list: 
                        return branch
           
                    
    def get_branches_of(self, element):
        '''
        function returning in a list the branch(es) at which the given
         element belongs 
        ''' 
        if element != None:
            class_name = self.interface.get_class_name_of(element)  
            if class_name == 'ElmTerm':
                return [branch for branch in self.branch_list \
                                if element in branch.terminal_bus_list]
            elif class_name == 'ElmRelay':               
                # check if it's between one of the relays elements
                return [branch for branch in self.branch_list \
                        if element in branch.relay_list1 or \
                            element in branch.relay_list2]           
            else:
                # check if it's between one of the branch elements
                return [branch for branch in self.branch_list \
                        if element in branch.pf_element_list]
                    
                    
    def find_relay_of_pf_relay(self, pf_relay, relay_list):
        '''
        function finding the Relay object of the given pf_relay
        '''
        found_relays =  [relay for relay in relay_list \
                if relay.pf_relay == pf_relay]
        return found_relays[0] if found_relays else None

#===============================================================================
# relay interconnection matrix functions
#===============================================================================
        
    def create_relay_matrix_for(self, lines, number_of_levels):
        '''
        Function returning the matrix which links logically the relay to express
        the multizone protection concept
        the matrix is generated for all relays protecting the line objects 
        passed as parameter  
        Args:
            lines: the lines used to get the relays, from each line we get the 
                    relevant branch 
            number_of_levels: how many zones/branches must be "investigated" 
        returns:
            the relay matrix
        '''
        #initialize class variables
        self.relay_matrix.clear()
        #import pydevd
        #pydevd.settrace(stdoutToServer=True, stderrToServer=True)
        number_of_lines = len(lines)
        # iterate throw all lines
        for line_index, line in enumerate(lines):
            if self.output_detail >= self.OutputDetail.NORMAL: 
                self.interface.print("Processing line " + self.interface.get_name_of(line))  
            if line_index % 20 == 0:
                if self.output_detail >= self.OutputDetail.NORMAL:
                    self.interface.print("Processing line#" + str(line_index) +\
                                         " of " + str(number_of_lines))
            # in both directions Side_1 and Side_2                                
            for side in [self.interface.Side.Side_1, self.interface.Side.Side_2]:  
                # get the branch of the given line            
                line_branch = self.get_branch_of(line)        
                if line_branch == None:              # a not connected line 
                    continue                         # is ignored
                # get the busbar at which the branch of the given line is connected 
                busbar = line_branch.get_busbar_at_side(side)   
                
                actual_level = 0
                self.relay_link_added = False
                # get a list of all branches connected to the bus bar
                branches = self.get_branch_of(busbar)                 
                if branches != None:          
                    for branch in branches:        # iterate throw all branches
                        #skip the branch we are protecting, ignore terminal branches
                        # out service branches and open branches
                        if branch != line_branch and \
                        branch.is_terminal_branch() == False and \
                        branch.is_out_of_service() == False and \
                        branch.is_open() == False:
                            # list of the busbars where we had already passed throw, 
                            # used to avoid loops     
                            already_collected_busbar_list = []   
                            already_collected_branch_list = []
                            already_collected_branch_list.append(line_branch)
                            already_collected_branch_list.append(branch)
                            self._follow_branches_bus_of(first_line = line, 
                                relay_branch = line_branch,
                                initial_branch = line_branch,
                                following_branch = branch, 
                                initial_busbar = busbar, 
                                down_relay_busbar =  busbar, 
                                number_of_levels = number_of_levels,
                                actual_level = actual_level, 
                                already_collected_busbar_list = already_collected_busbar_list,
                                relay_list = self.relay_matrix,
                                already_collected_branch_list = already_collected_branch_list)  # recursive call
                else:
                    if self.output_detail >= self.OutputDetail.DEBUG:
                        if busbar != None:
                            self.interface.print("\nWarning: " + 
                                self.interface.get_full_name_of(busbar) + 
                                "has no branch connected. \n" )
                        else:
                            self.interface.print("\nWarning: " + 
                                line_branch.name + " has no busbar at side {}".\
                                        format(side) + "\n" )
#                             self.interface.print("Line branch = " + line_branch.name)
#                             self.interface.print("Terminals: ")
#                             for terminal in line_branch.terminal_bus_list:
#                                 if terminal != None:
#                                     self.interface.print(self.interface.\
#                                                          get_name_of(terminal))        
                            
        # output code for debugging
#         if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
#             self.interface.print("\nRelay matrix: \n")
#             for relay in self.relay_matrix:
#                 self.interface.print("\n\t" + 
#                                 self.interface.get_full_name_of(relay.pf_relay), )
#                 self.interface.print("\t Protected branch: " + 
#                         self.interface.get_full_name_of(relay.protected_line), )
#                 self.interface.print("\t Step: {}".format(relay.step), )
#                 self.interface.print("\t\t Selective relay links:")
#                 for link in relay.selective_relay_links:
#                     self.interface.print("\t\t\t" + 
#                                         self.interface.get_full_name_of(link), )        
#                 self.interface.print("\t\t Monitored  relay links:")
#                 for link in relay.monitored_relay_links:
#                     self.interface.print("\t\t\t" + 
#                                         self.interface.get_full_name_of(link), ) 
        self.interface.print("\nRelay matrix completed. \n")
        return self.relay_matrix
        
#------------------------------------------------------------------------------ 
# Recursive function
#------------------------------------------------------------------------------ 
        
    def _follow_branches_bus_of(self, first_line, relay_branch, initial_branch,
                             following_branch, initial_busbar, 
                             down_relay_busbar, number_of_levels, 
                             actual_level, already_collected_busbar_list,
                             relay_list, already_collected_branch_list):
        '''
        it goes throw the branches / bus of the following couples of busbar / branches
        Recursive function called by <create_relay_matrix_for>
        
        Args:
            first_line: the first line  at which the selectivity steps refer
            realy_branch = the branch where the first set of relays is
            initial_branch: the branch going to the initial_busbar
            following_branch: the branch from which we are looking for the 
                              next busbar
            initial_busbar: the busbar from which we start
            number_of_levels: the max number of selectivity steps we are 
                              investigating
            actual_level: the actual selectivity step
            already_collected_busbar_list : the list of all busbars which have 
                                          already been parsed
            relay_list: list of all relay links  (it's what we are filling!)  
            already_collected_branch_list: the list of all branches which have 
                                          already been parsed
        '''  
        actual_level += 1                                  # increase the level
        # check if we had already reached the max number of levels 
        if actual_level > number_of_levels:                    
            return  
        # store the new intial bus bar in the list of already investigated busbars 
        already_collected_busbar_list.append(initial_busbar)                                                                                 
                                                
        # we get the list of the relays which protect the branch we are investigating
        down_relay_list = relay_branch.get_relay_at_busbar(down_relay_busbar)        
        # get new branch relays        
        opposite_relay_list = following_branch.get_relays_opposite_to(initial_busbar)       
        close_relay_list = following_branch.get_relay_at_busbar(initial_busbar) 
        # we prepare the data for the next couple line/busbar
        next_bus = following_branch.get_other_busbar_of(initial_busbar)  
        
        # check if the next busbar has already been processed
        if next_bus in already_collected_busbar_list:
            return
        # get a list of all branches connected to the bus bar
        branches = self.get_branch_of(next_bus)                 
        if len(down_relay_list) > 0:        # the branch is directly protected by some relays            
            if len(opposite_relay_list) > 0: # also the opposite relays are available
                for down_relay in down_relay_list: # so we create the relay matrix record(s)
                    relay_list.append(self.Relay_link(pf_relay = down_relay,  
                              selective_relay_links = opposite_relay_list,
                              monitored_relay_links = close_relay_list, 
                              protected_line = first_line,
                              step = actual_level,
                              branch_list = copy.copy(already_collected_branch_list)))
                new_relay_branch =  following_branch  # we move one step further
                self.relay_link_added = True
            else:                        # no opposite relay
                # the down_relay_list will be retrieved again from this initial_branch
                new_relay_branch = relay_branch  
            if branches != None:            
                for branch in branches:                # iterate throw all branches
                    # skip the branch from which we are coming from, out of service
                    # and open branches
                    if branch == following_branch or\
                     branch.is_out_of_service() == True or\
                     branch.is_open() == True:     
                        continue    
                    already_collected_branch_list.append(branch)                        
                    self._follow_branches_bus_of(first_line = first_line, 
                                relay_branch = new_relay_branch,
                                initial_branch = following_branch,
                                following_branch = branch, 
                                initial_busbar = next_bus,
                                down_relay_busbar = down_relay_busbar, 
                                number_of_levels = number_of_levels,
                                actual_level = actual_level, 
                                already_collected_busbar_list = already_collected_busbar_list,
                                relay_list = relay_list,
                                already_collected_branch_list = copy.copy(already_collected_branch_list))  # recursive call
        else:   # we look for the first branch where we have protection relays        
            if branches != None: 
                for branch in branches:                 # iterate throw all branches
                    # skip the branch from which we are coming from
                    if branch == following_branch:         
                        continue    
                    already_collected_branch_list.append(branch)                        
                    self._follow_branches_bus_of(first_line = first_line, 
                                relay_branch = following_branch,
                                initial_branch = following_branch,
                                following_branch = branch, 
                                initial_busbar = next_bus,
                                down_relay_busbar = down_relay_busbar, 
                                number_of_levels = number_of_levels,
                                actual_level = actual_level-1, \
                                already_collected_busbar_list = already_collected_busbar_list, 
                                relay_list = relay_list,
                                already_collected_branch_list = copy.copy(already_collected_branch_list))  # recursive call
        # add the relay link if no selective relay has been found
        if self.relay_link_added == False and len(down_relay_list) > 0 and \
        len(opposite_relay_list) == 0 and actual_level == 1:
            for down_relay in down_relay_list: # so we create the relay matrix record(s)
                relay_list.append(self.Relay_link(pf_relay = down_relay,  
                                  selective_relay_links = opposite_relay_list,
                                  monitored_relay_links = close_relay_list, 
                                  protected_line = first_line,
                                  step = actual_level,
                                  branch_list = copy.copy(already_collected_branch_list)))
            self.relay_link_added = True
            
            
#===============================================================================
# Transformer collection function
#===============================================================================
    def get_bus_trnsformers_from(self, busbar_name, number_of_steps):
        '''      
        function returning a collection of all trafo connected to the given 
        busbar for the given number of number_of_steps
        '''
        lines = self.get_bus_lines_from(busbar_name, number_of_steps)
        return [line for line in lines if 'ElmTr' in self.get_class_name_of(line)]
    
    
#===============================================================================
# Line collection functions
#===============================================================================
    def get_bus_lines_from(self, busbar_name, number_of_steps):
        '''      
        function returning a collection of all lines connected to the given 
        busbar for the given number of number_of_steps
        '''
        # list of the busbars where we had already passed throw
        collected_busbar_list = []  
        # the list of the lines we are looking for, it's the returned list
        collected_lines_list = []    
        # get the busbar object from the given name (a list can be returned also)
        initial_busbar_list = self.interface.get_element_by_name(busbar_name) 
        actual_level = 0   
        if initial_busbar_list != None:    
            # we iterate throw all initial bus bars
            for initial_busbar in initial_busbar_list:  
                # we get all branches of the first busbar
                connected_branches_list = self.get_branch_of(initial_busbar) 
                # get all lines of all branches connected to the first bus bar
                collected_lines_list += [element for connected_branch \
                                         in connected_branches_list \
                                         for element in connected_branch.pf_element_list \
                    if self.interface.is_line(element) ]
                # we store as investigated the first busbar
                collected_busbar_list.append(initial_busbar)   
                # iterate throw all branches
                for connected_branch in connected_branches_list: 
                    # get the other busbar
                    new_busbar = connected_branch.get_other_busbar_of(initial_busbar)  
                    # collect all the lines for that branch (recursive)
                    if new_busbar != None:
                        self._collect_lines_bus_from(initial_busbar = new_busbar, 
                                                number_of_levels = number_of_steps-1, 
                                                actual_level = actual_level, 
                                                collected_busbar_list = collected_busbar_list, 
                                                line_list = collected_lines_list) 
            
        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
            self.interface.print("\nCollected lines: \n")
            for collected_line in collected_lines_list:
                self.interface.print("\t\t\t" + 
                            self.interface.get_full_name_of(collected_line), )
                
        return collected_lines_list
#------------------------------------------------------------------------------ 
# Recursive function
#------------------------------------------------------------------------------         
    def _collect_lines_bus_from(self, initial_busbar, number_of_levels, 
                               actual_level, collected_busbar_list, line_list):
        '''
        it goes throw the lines / bus of the following couples of busbar/lines
        Recursive function called by <get_bus_lines_from>
        
        Args:
        initial_busbar: the busbar from which we start
        number_of_levels: the max number of selectivity steps we are 
                            investigating
        actual_level: the actual selectivity step
        already_collected_busbar_list : the list of all busbars which have
                                        already been parsed
        line_list: list of all lines  (it's what we are filling!)  
        '''
        
        actual_level += 1                       # increase the level
        # check if we had already reached the max number of levels
        # or the busbar has already been processed 
        if actual_level > number_of_levels or \
            initial_busbar in collected_busbar_list:             
            return          
        # store the intial bus bar in the list of already investigated busbars
        collected_busbar_list.append(initial_busbar)                                       
        # we get all branches of the first busbar
        connected_branches_list = self.get_branch_of(initial_busbar) 
        # get all lines of all branches connected to the first bus bar
        line_list += [element for connected_branch in connected_branches_list \
                      if connected_branch.get_other_busbar_of(initial_busbar) not in \
                                                        collected_busbar_list
                      for element in connected_branch.pf_element_list \
                      if self.interface.is_line(element) ]
        # we store as investigated the first busbar
        for connected_branch in connected_branches_list:
            new_busbar = connected_branch.get_other_busbar_of(initial_busbar)
            if new_busbar != None and new_busbar not in collected_busbar_list:
                self._collect_lines_bus_from(initial_busbar = new_busbar, 
                                            number_of_levels = number_of_levels,
                                            actual_level =  actual_level,
                                            collected_busbar_list = collected_busbar_list,
                                            line_list = line_list) 
  
  
#===============================================================================
# Function to identify if a bus bar is a load busbar
#===============================================================================              

    def is_load_bus(self, busbar, feeder, network_configuration):
        '''
        check if the given busbar is fed by any branch which is not the given
        feeder
        '''
        #check if the given busbar has already been detected for the given network
        #configuration as a load busbar
        busbar_network_string = network_configuration + \
                                     self.interface.get_full_name_of(busbar) + \
                                     self.interface.get_full_name_of(feeder)
        load_busbar_value = self.load_busbars.get(busbar_network_string, "Not found")
        if load_busbar_value != "Not found":
            return load_busbar_value       
        # list of the busbars where we had already passed throw
        collected_busbar_list = []  
        #get the branch at which the given feeder belongs
        given_feeder_branch = self.get_branch_of(feeder)
        # if that branch is terminal then it's a load bus
        # probably wrong: this is the given branch and should not be considered
#         if given_feeder_branch != None and \
#         given_feeder_branch.is_terminal_branch():
#             self.load_busbars.update({busbar_network_string:True})
#             return True
        collected_busbar_list.append(busbar)
        is_load_bus =  self._is_follow_branch_load_bus_of(busbar, given_feeder_branch,\
                                              collected_busbar_list)
        self.load_busbars.update({busbar_network_string:is_load_bus})
        return is_load_bus
    
    
#------------------------------------------------------------------------------ 
# Recursive function
#------------------------------------------------------------------------------

    def _is_follow_branch_load_bus_of(self, busbar, feeder_branch, collected_busbar_list):
        '''
        function going throw all busbar/line to find if the given bus bar is a
        a load busbar (no feeding source except the given feeder)
        '''
        # we get all branches of the given busbar
        connected_branches_list = self.get_branch_of(busbar)
        if connected_branches_list != None:
            for connected_branch in connected_branches_list:
                # not consider the branch of the given feeder and the branches
                # out of service or open
                if connected_branch == feeder_branch or \
                connected_branch.is_open() or \
                connected_branch.is_out_of_service():
                    continue
                # in case of a connected generator it isn not a load bus...
                if connected_branch.is_generator_branch():
                    return False
                else: # recursive research
                    other_busbar = connected_branch.get_other_busbar_of(busbar)
                    # if the following bus bar has already been collected skip it!
                    if other_busbar in collected_busbar_list:
                        continue
                    collected_busbar_list.append(other_busbar)
                    # recursive call
                    is_load = self._is_follow_branch_load_bus_of(other_busbar,\
                                            connected_branch, collected_busbar_list)
                    if is_load == False:
                        return False
            
        return True
        


#===============================================================================
# Function indentifying the paths between loads (= measurement units)
#===============================================================================              

    def get_load_paths(self):
        '''
        function identifying for each load all the possible paths to another load
        and storing the identified path in the "paths" class variable 
        '''
        # create the path data
        for load in self.interface.get_loads():
            self.paths[self.interface.get_name_of(load)] = self.get_paths_of(load)
        # output code for debugging
        if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
            self.interface.print("\nPath collection: \n")
            for load_key, path_list in self.paths.items():
                self.interface.print("\t Load: " + load_key), 
                self.interface.print("\t\t Paths:")
                for path in path_list:
                    self.interface.print("\t\t" + self.interface.
                                    get_name_of(path.Measurement1) + " - " + 
                                     self.interface.get_name_of(path.Measurement2))
                    self.interface.print("\t\t\tBranches:")
                    for branch in path.branch_list:
                        self.interface.print("\t\t\t" + branch.name)                                          
                    self.interface.print("\t\t\t Path Z: {:f} + {:f}J".\
                                         format(path.zp.real , path.zp.imag)) 
                    self.interface.print("\t\t\t Path Zn: {:f} + {:f}J".\
                                         format(path.zn.real , path.zn.imag))   
                    self.interface.print("\t\t\tBusbars:")
                    for busbar in path.busbar_list:                       
                        top_bus_string = "(top)" if busbar.top_busbar == True \
                                            else ""
                        self.interface.print("\t\t\t" + self.interface.
                                                get_name_of(busbar.pf_busbar) +\
                                                top_bus_string )
            self.interface.print("\nEnd Path\n")       
            
            
    
    def get_paths_of(self, load, network_configuration = ''):
        '''
        get all paths connecting the given load with another load
        it returns a list of path objects
        '''      
        # list of the loads which have already been connected
        collected_loads_list = []  
        # list of the path branchess
        path_branches_list = []
        # list of all found path
        returned_path_list = []
        # list with the busbar which have been already connected
        already_collected_busbar_list = []
        # get the branch at which the given load belongs
        given_load_branch = self.get_branch_of(load)
        # create a new path directly in the path list
        returned_path_list.append(MeasurementPath())
        returned_path_list[len(returned_path_list)-1].Measurement1 = load
        # remove from the research the starting load
        collected_loads_list.append(load)
        #recursive call
        returned_path_list =  self._follow_path_of(load,
                                                   given_load_branch,
                                                   None,
                                                   collected_loads_list,
                                                   path_branches_list,
                                                   returned_path_list,
                                                   already_collected_busbar_list)  
        # remove the last path if not complete 
        if len(returned_path_list) > 0 and \
        returned_path_list[len(returned_path_list)-1].Measurement2 == None:
            returned_path_list.pop()
        # set the "Top path" flag, remove duplicated bubar and calculate the 
        # Z for all found paths
        for path in returned_path_list:
            path.busbar_list = list(set(path.busbar_list))
            path.set_top_busbars(self)
            path.calculate_impedances()
        return returned_path_list
    
    
#------------------------------------------------------------------------------ 
# Recursive function
#------------------------------------------------------------------------------

    def _follow_path_of(self, given_load, branch, busbar, collected_loads_list,
                        path_branches_list, returned_path_list, already_collected_busbar_list):
        '''
        function going throw all busbar/line to find a not yet found another
        load
        args:
            given_load: the load from which we are trying to collect all paths
            branch: the branch where we are in the path detection algorithm
            busbar: the bus bar at which the given branch goes
            collected_loads_list: the list of all loads which have already been 
                    detected and which are not used anymore as valide path ends
            path_branches_list: the list containing all the branches which are 
            part of the possible path we are trying to define   
            returned_path_list: the list of the paths already found    
        returns:
            the list containing all found paths
        '''
        # add to the new path the branch we are evaluating
        new_path_branches_list = copy.copy(path_branches_list)
        new_path_branches_list.append(branch) 
        # check if the branch includes a load not yet evaluated
        branch_load_list = branch.get_load()
        branch_load = branch_load_list[0] if len(branch_load_list) > 0 else None
        if branch_load is not None and branch_load not in collected_loads_list:
            # add the 2nd load in the path object and store it
            returned_path_list[len(returned_path_list)-1].branch_list = \
                                                            new_path_branches_list
            returned_path_list[len(returned_path_list)-1].Measurement2 = branch_load
            # store the load in the list already evaluated load
            collected_loads_list.append(branch_load)
            # create a new path
            returned_path_list.append(MeasurementPath())
            # add in the bus list the busses already collected
            returned_path_list[len(returned_path_list)-1].busbar_list.extend(\
            copy.copy(returned_path_list[len(returned_path_list)-2].busbar_list))
            # add the busbar in the path busbar list
            returned_path_list[len(returned_path_list)-1].busbar_list.\
                    append(MeasurementPath.Busbar(pf_busbar = busbar,
                                                  top_busbar = False))
            returned_path_list[len(returned_path_list)-1].Measurement1 = given_load
            # reset the list of already collected bus bar = reinit search algorithm
            already_collected_busbar_list.clear()
            return returned_path_list
        else:    
            # get the busbar at the other side of the branch
            opposite_busbar = branch.get_other_busbar_of(busbar) if busbar != None \
                        else branch.get_busbar_at_side(self.interface.Side.Side_1)
            # avoid the iteration throw the same busbars
            if opposite_busbar in already_collected_busbar_list:
                return returned_path_list
            # add the busbar in the path busbar list
            returned_path_list[len(returned_path_list)-1].busbar_list.\
                    append(MeasurementPath.Busbar(pf_busbar = opposite_busbar,
                                                  top_busbar = False))
            # get all branches of the new busbar
            connected_branches_list = self.get_branch_of(opposite_busbar)
            if connected_branches_list != None:
                for connected_branch in connected_branches_list:
                    # not consider the branch of the given feeder and the branches
                    # out of service or open
                    if connected_branch == branch or \
                    connected_branch.is_open() or \
                    connected_branch.is_out_of_service():
                        continue
                    # in case of a connected generator it is not a load bus...
                    if connected_branch.is_generator_branch():
                        continue
                    else: # recursive research
                        # against any infinite loop...
                        already_collected_busbar_list.append(opposite_busbar)
                        # recursive call
                        returned_path_list = self._follow_path_of(given_load,
                                                                  connected_branch,
                                                                  opposite_busbar,\
                                                                  collected_loads_list,
                                                                  new_path_branches_list,
                                                                   returned_path_list,
                                                                   already_collected_busbar_list) 
                
        return returned_path_list   
   
   
        
#===============================================================================
# Function creating all 2 lines paths for the given lines
#===============================================================================   
     
     
    def create_paths(self, lines, relay_list):
        '''
        collect all paths of the actual project for the given lines
        '''
        path_list = []
        for line in lines:
            line_branch = self.get_branch_of(line)
            if  line_branch != None:    # collect the fault info
                bus_1 = line_branch.get_busbar_at_side(\
                                            self.interface.Side.Side_1)
                bus_2 = line_branch.get_busbar_at_side(\
                                            self.interface.Side.Side_2)
                
                for side in [self.interface.Side.Side_1, self.interface.Side.Side_2]:
                    branch_cub_relays = line_branch.get_relay_at_side(side)
                    if branch_cub_relays:
                        for relay in branch_cub_relays:
                            for relay_link in self.relay_matrix:
                                if relay_link.pf_relay == relay and relay_link.step == 1:                                                                                              
                        
                                    path_elements = []                                                                              
                                    selective_relay = None
                                   
                                    path_name = ''
                                    bus_just_added = False
                                    number_of_branches = len(relay_link.branch_list)
                                    for index, branch in enumerate(reversed(relay_link.branch_list)):
                                        bus_1 = branch.get_busbar_at_side(self.interface.Side.Side_1)
                                        bus_2 = branch.get_busbar_at_side(self.interface.Side.Side_2)
                                        next_bus_1 = None
                                        next_bus_2 = None
                                        if index < number_of_branches - 1:
                                            next_index = number_of_branches - index - 2
                                            next_bus_1 = relay_link.branch_list[next_index].get_busbar_at_side(self.interface.Side.Side_1)
                                            next_bus_2 = relay_link.branch_list[next_index].get_busbar_at_side(self.interface.Side.Side_2)
#                                             if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
#                                                 if next_bus_1:
#                                                     self.interface.print("Next Bus 1: " +\
#                                                      self.interface.get_name_of(next_bus_1))
#                                                 if next_bus_2:
#                                                     self.interface.print("Next Bus 2: " +\
#                                                      self.interface.get_name_of(next_bus_2))
#                                         if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
#                                             if bus_1:
#                                                 self.interface.print("Bus 1: " + \
#                                                 self.interface.get_name_of(bus_1))
#                                             if bus_2:
#                                                 self.interface.print("Bus 2: " + \
#                                                 self.interface.get_name_of(bus_2))                                                                           
    
                                        if bus_just_added == False:
                                            if bus_1 not in [next_bus_1, next_bus_2] and\
                                            bus_1 not in path_elements and bus_1:
                                                path_elements.append(bus_1)
                                                bus_just_added = True
                                            elif bus_2 and bus_2 not in path_elements:
                                                path_elements.append(bus_2)
                                                bus_just_added = True
                                        # remove switches and add elements... 
                                        new_elements = self.get_branch_filtered_elements_of(\
                                                        branch, path_elements)   
                                        path_elements += new_elements
                                        if len(new_elements) > 0:
                                            bus_just_added = False
#                                         else:
#                                             if self.output_detail >= self.OutputDetail.VERBOSEDEBUG:
#                                                 self.interface.print("No element has been added")
                                        if bus_just_added == False:
                                            if branch.get_number_of_terminal_busses() == 1 and\
                                            bus_1 not in path_elements and bus_1:
                                                path_elements.append(bus_1)
                                                bus_just_added = True
                                            if branch.get_number_of_terminal_busses() > 1 and\
                                            bus_2 not in path_elements and bus_2:
                                                path_elements.append(bus_2)
                                                bus_just_added = True
                                      
                                        if relay_link.selective_relay_links:
                                            pf_selective_relay = relay_link.\
                                                    selective_relay_links[0]
                                            selective_relay = self.find_relay_of_pf_relay\
                                                (pf_selective_relay, relay_list)
                                        
                                    # generic code to set the name
                                    for element in path_elements:
                                        if self.interface.\
                                            get_class_name_of(element) == 'ElmTerm':
                                            if len(path_name) > 0:
                                                path_name += '-'
                                            path_name += self.get_bus_code(\
                                                self.interface.get_name_of(element)) 
                                        
                                    if path_name not in [self.interface.get_name_of(path)\
                                        for path in path_list]:  
                                        if self.output_detail >= self.OutputDetail.NORMAL and\
                                        selective_relay != None:
                                            self.interface.print("  Creating path: "\
                                                                + path_name)  
                                            self.interface.print("     Relay 1: "\
                                            + self.interface.get_name_of(relay))  
                                            self.interface.print("     Relay 2: "\
                                            + self.interface.get_name_of(pf_selective_relay))                                     
                                        path_list.append(self.interface.\
                                                    create_path(path_name, \
                                                            path_elements))
        return path_list
    
    
#------------------------------------------------------------------------------ 
# Recursive function
#------------------------------------------------------------------------------
    
    def follow_tap_elements(self, initial_bus, initial_branch,  path_elements):
        '''
        recursive function getting all tap elements
        '''
        # get a list of all branches connected to the bus bar
        branches = self.get_branch_of(initial_bus)                 
        if branches != None:          
            for branch in branches:        # iterate throw all branches
                #skip the branch we are protecting, ignore terminal branches
                # out service branches and open branches
                if branch != initial_branch and \
                branch.is_terminal_branch() == False and \
                branch.is_out_of_service() == False and \
                branch.is_open() == False:
                    pass
                
#------------------------------------------------------------------------------ 
# Auxiliary functiona
#------------------------------------------------------------------------------

    def get_branch_filtered_elements_of(self, branch, path_elements):
        '''
        get the given branch elements filtering them to remove
        shunt, breaker, loop elements etc
        '''
        return [element for element in \
                branch.get_elements() if self.interface.\
                get_class_name_of(element) != 'StaSwitch' and \
                self.interface.\
                get_class_name_of(element) != 'ElmLod' and \
                self.interface.\
                get_class_name_of(element) != 'ElmShnt' and \
                element not in path_elements]
        
    
    def get_bus_code(self, bus_name):
        '''
        ancillary function getting the bus code (first part of the name before ' ')
        '''
        string_parts = bus_name.split()
        return string_parts[0] if len(string_parts) > 0 else bus_name    
        
    
