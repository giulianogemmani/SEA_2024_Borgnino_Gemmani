'''
Created on 11 Feb 2019

@author: AMB
'''

from collections import namedtuple
from tracer import branch
from builtins import complex



class MeasurementPath(object):
    '''
    class reprensenting the path between two measurement elements
    '''
    Busbar = namedtuple("Busbar", "pf_busbar top_busbar")

    def __init__(self):
        '''
        Constructor
        '''
        self.branch_list = []
        self.busbar_list = []
        self.Measurement1 = None
        self.Measurement2 = None
        self.zp = complex(0, 0)  # path phase impedance
        self.zn = complex(0, 0)  # path neutral impedance
    
    
    def set_top_busbars(self, grid):
        '''
        function finding which bus bar is the path "top bas bar" = the busbar
        fed by a branch which doesn't belong to the path
        the found busbar "top_busbar" variable is set equal to True
        '''
        for index, busbar in enumerate(self.busbar_list):
            bus_connections_list = grid.get_branch_of(busbar.pf_busbar)
            for  connection in bus_connections_list:
                # if a connection is not part of the path and is feeding the bus
                # then we have found or "top bus"
                if connection not in self.branch_list and \
                            grid.is_load_bus(busbar.pf_busbar, connection, "") == True:
                    self.busbar_list[index] = self.Busbar(busbar.pf_busbar,\
                                                           top_busbar = True)
                    break
            if busbar.top_busbar == True:
                break
       
            
    def calculate_impedances(self):
        '''
        function calculating the phase and the neutral branch impedances
        as sum of all the impedances of the lines belonging to the branch
        '''
        self.zp = sum([z.Z for branch in self.branch_list \
                       for z in branch.Z.values()]) 
        self.zn = sum([z.Zn for branch in self.branch_list \
                       for z in branch.Z.values()])   
          
    def get_delta_neutral_v_at(self, time_index): 
        '''
        function returning the delta V alog the branch neutral at the given 
        time index
        '''
        return_delta_v = 0
        # go throw the busbar list and get the top busbar
        top_busbar = None
        for busbar in self.busbar_list:
            if busbar.top_busbar == True:
                top_busbar = busbar.pf_busbar
                break
        # flag to add/subtract the v along the path
        reverse_sign = False
        for branch in self.branch_list:
            if reverse_sign == False and \
            (branch.terminal_bus_list[0] == top_busbar or \
                (len(branch.terminal_bus_list) > 1 and \
                    branch.terminal_bus_list[1] == top_busbar)):
                reverse_sign = True
            sign = -1 if reverse_sign == True else 1
            return_delta_v += branch.get_delta_neutral_v_at(time_index) * sign
                 
        return return_delta_v
    
    def get_delta_v_at(self, time_index, phase_index = None): 
        '''
        function returning the delta V as complex number along the branch 
        neutral at the given time index for the given phase index
        '''
        return_delta_v = 0
        # go throw the busbar list and get the top busbar
        top_busbar = None
        for busbar in self.busbar_list:
            if busbar.top_busbar == True:
                top_busbar = busbar.pf_busbar
                break
        # flag to add/subtract the v along the path
        reverse_sign = False
        for branch in self.branch_list:
            sign = -1 if reverse_sign == True else 1
            return_delta_v += branch.get_delta_v_at(time_index, phase_index) * sign
            if reverse_sign == False and \
            (branch.terminal_bus_list[0] == top_busbar or \
                (len(branch.terminal_bus_list) > 1 and \
                    branch.terminal_bus_list[1] == top_busbar)):
                reverse_sign = True
            
            
                 
        return return_delta_v 
        
    