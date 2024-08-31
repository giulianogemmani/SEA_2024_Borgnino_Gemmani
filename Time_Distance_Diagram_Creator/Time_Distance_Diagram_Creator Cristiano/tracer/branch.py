'''
Created on 18 Oct 2018

@author: AMB
'''


class Branch(object):
    '''
    Class representing an aggregator of network elements
    It's delimited by any bus having more than 2 or one single connection or 
    by any cubicle containing a relay
    It's also delimited by any open switch.
    '''


    def __init__(self, name, pf_element_list, relay_list1, relay_list2,
                                                 terminal_bus_list, interface):
        '''
        Constructor
        '''
        self.transformer = None
        self.generator = None
        self.lines = []
        
        self.name = name
        self.pf_element_list = pf_element_list
        self.relay_list1  = relay_list1
        self.relay_list2  = relay_list2
        self.terminal_bus_list = terminal_bus_list
        self.interface = interface
        self.generator_available = False
        self.transformer_available = False
        self.line_available = False
        self.check_generator_availability()
        self.check_transformer_availability()
        self.check_line_availability()
        
        # case of a generator branch with relays looking only inside the generator
        if self.generator_available == True and \
        len(self.relay_list1) == 0:
            self.relay_list1 = self.relay_list2
   
    def check_generator_availability(self):
        '''
        function checking if any syncrhonous or asynchronous generator, or 
        network is part of the branch
        '''
        for element in self.pf_element_list:
            if self.interface.is_generator(element):
                self.generator_available = True
                self.generator = element
                break  
       
            
    def check_transformer_availability(self):
        '''
        function checking if any kind of transformer is part of the branch
        '''
        for element in self.pf_element_list:
            if self.interface.is_transformer(element):
                self.transformer_available = True
                self.transformer = element
                break  
         
            
    def check_line_availability(self):
        '''
        function checking if any kind of line is part of the branch
        '''
        for element in self.pf_element_list:
            if self.interface.is_line(element):
                self.line_available = True
                self.lines.append(element)
            
        
    def is_terminal_branch(self):
        '''
        function returning true if only one "terminal busbar" is present and 
        no generator is present in the pf_element_list
        '''
        return self.get_number_of_terminal_busses() == 1 and\
                         self.generator_available == False
       
            
    def is_generator_branch(self):
        '''
        function returning true if only one "terminal busbar" is present and 
        a generator is present in the pf_element_list
        '''
        return self.get_number_of_terminal_busses() == 1 and\
                        self.generator_available == True
                        
                        
    def is_transformer_branch(self):
        '''
        function returning true if at least one transformer is present 
        in the pf_element_list
        '''
        return True if self.transformer_available == True else False
    
    
    def is_line_branch(self):
        '''
        function returning true if at least one line is part of the branch
        '''
        return True if self.lines else False
    
    
    def is_out_of_service(self):
        '''
        function checking if the branch is in service
        it returns true if no internal element is out of service
        '''
        return any(self.interface.is_out_of_service(element) == True \
                                            for element in self.pf_element_list)
        
        
    def is_open(self):
        '''
        function checking if the branch in open
        it returns true if no internal element is open
        '''
        return any(self.interface.is_open(element) == True \
                                            for element in self.pf_element_list)    
        
        
    def get_busbar_at_side(self, side):
        '''
        function returning the busbar at which the branch is connected at 
        the given side   
        '''
        return self.terminal_bus_list[0] if side == self.interface.Side.Side_1 \
            else self.terminal_bus_list[1] if len(self.terminal_bus_list) > 1 \
            else None
    
    
    def get_other_busbar_of(self, busbar):
        '''
        function returning the branch terminal busbar which is not the busbar passed 
        as parameter
        '''
        other_bus = self.get_busbar_at_side(self.interface.Side.Side_1)       
        return other_bus if other_bus != busbar else \
                        self.get_busbar_at_side(self.interface.Side.Side_2)
                        
                        
    def get_busbars(self):
        '''
        function returning all basbar which are part of the branch
        '''
        return [element for element in self.pf_element_list \
                                    if self.interface.is_busbar(element)]
    
    
    def get_last_busbar_from(self, initial_busbar):
        '''
        function returning the from the given "initial bus bar"
        '''
        busbars = self.get_busbars()
        # if there is no internal busbar add at least the terminal bus bars
        if len(busbars) == 0:
            busbars += self.terminal_bus_list
            # remove the 'none' elements
            busbars = [busbar for busbar in busbars if busbar != None]
        if busbars:
            return busbars[0] if busbars[0] != initial_busbar else busbars[-1]
        else:
            return None
    
        
    def get_relay_at_side(self, side):
        '''
        function returning the relays located at the given side
        '''
        return self.relay_list1 if side == self.interface.Side.Side_1 else \
                                                             self.relay_list2
        
     
    def get_relay_at_busbar(self, busbar):
        '''
        function returning the relays located in a cubicle of the given busbar
        A special procedure is implemented for generator branches with  only 
        one relay slot is populated
        In that case the available relay slot is always returned 
        '''
        if self.is_generator_branch() and (len(self.relay_list1) == 0 or \
                                           len(self.relay_list2) == 0):
            return self.relay_list1 if len(self.relay_list1) > 0  else self.relay_list2
        if len(self.relay_list1) > 0:
            return self.relay_list1 \
                if self.interface.get_relay_busbar_of(self.relay_list1[0]) == busbar \
                else self.relay_list2
        elif len(self.relay_list2) > 0:
            return self.relay_list2 \
                if self.interface.get_relay_busbar_of(self.relay_list2[0]) == busbar \
                else self.relay_list1
        return self.relay_list1   # both lists are void return the first void string   
        
        
    def get_relays_opposite_to(self, busbar):
        '''
        function returning the relays located at the cubicle belonging to the 
        other busbar (other = not the given busbar)
        A special procedure is implemented for generator branches with only one
        relay slot is populated
        In that case the available relay slot is always returned 
        '''
        if self.is_generator_branch() and (len(self.relay_list1) == 0 or \
                                           len(self.relay_list2) == 0):
            return self.relay_list1 if len(self.relay_list1) > 0  else self.relay_list2
        if len(self.relay_list1) > 0:
            return self.relay_list2 \
                if self.interface.get_relay_busbar_of(self.relay_list1[0]) == busbar \
                else self.relay_list1
        elif len(self.relay_list2) > 0:
            return self.relay_list1 \
                if self.interface.get_relay_busbar_of(self.relay_list2[0]) == busbar \
                else self.relay_list2
        return self.relay_list1   # both lists are void return the first void string
    
        
    def get_relay_list(self, index):
        '''
        function returning relay_list1 or relay_list2 depending up on the 
        given index
        '''
        return self.relay_list1 if index == 0 else  self.relay_list2
    
    
    def get_lines(self):
        '''
        functions returning all the lines belonging to this branch
        '''
        return [element for element in self.pf_element_list \
                                    if self.interface.is_line(element)]
        
        
    def get_transformers(self):
        '''
        functions returning all the transformers belonging to this branch
        '''
        return [element for element in self.pf_element_list \
                                    if self.interface.is_transformer(element)]
    
    def get_elements(self):
        '''
        function returning a list with all elements belonging to the branch
        '''
        return self.pf_element_list
    
     
        
    def get_length(self):
        '''
        function returning the sum of the lenths of all branch lines     
        '''
        total_length = 0
        for line in self.get_lines():
            total_length += self.interface.get_line_length_of_(line)        
        return total_length
    
    def get_branch_z(self):
        '''
        function returning the total branch of the branch considering the lines
        and the transformers impedance
        '''
        from math import sin, cos, pi
        return_z = complex(0, 0)
        lines = self.get_lines()
        for line in lines:
            line_angle = self.interface.get_line_angle(line) * pi/180.
            line_z = self.interface.get_line_z(line)
            return_z += complex(line_z * cos(line_angle), line_z * sin(line_angle))
        transformers = self.get_transformers()
        for transformer in transformers:
            return_z += self.interface.get_transformer_z(transformer)
        return return_z
    
        
    def get_number_of_terminal_busses(self):
        '''
        return the number of connected terminals counting the number of objects 
        present inside self.terminal_bus_list
        '''
        return sum(bus is not None for bus in self.terminal_bus_list)