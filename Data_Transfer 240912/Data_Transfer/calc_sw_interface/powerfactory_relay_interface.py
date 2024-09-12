'''
Created on 14 Jan 2019

@author: AMB
'''

from collections import namedtuple


# relay setting data structure
# all info to access the setting are stored here together with the setting value
# and the connection with the value in the  interface (the dictionary key)
RelaySetting = namedtuple("RelaySetting", "subrelay_name\
                                               element_name \
                                               setting_name \
                                               value")

class PowerFactoryRelayInterface():
    '''
    class collecting all functions which allow to read and write the relay 
    settings
    '''

    
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        self.interface = interface
        self.pf_relay = pf_relay
     
     
    def get_curve_object(self, part_of_curve_name):
        '''
        get from the TypChatoc objects contained inside the relay the first curve 
        containing the given name or part of the given name
        '''
        relay_type = self.interface.get_relay_type_of(self.pf_relay)
        if relay_type:
            curves = relay_type.GetContents("*.TypChatoc", 1)
            if curves:
                for curve in curves:
                    if part_of_curve_name in self.interface.get_name_of(curve):
                        return curve
        return None    
     
     
    def read_settings(self):
        '''
        function reading all relay settings
        ''' 
        return_values = {}
        for setting_key in self.setting_list.keys():
            if len(self.setting_list[setting_key].subrelay_name) > 0 :
                parent = self.interface.get_element_by_name(\
                                    self.setting_list[setting_key].subrelay_name)
            else:
                parent = self.pf_relay
            element_list = self.interface.get_element_by_name_and_parent(\
                                    self.setting_list[setting_key].element_name,\
                                            parent)
            if element_list:
                value = self.interface.get_attribute(element = element_list[0],\
                                attribute_name = self.setting_list[setting_key].\
                                                                setting_name)
            else:
                value = None
            if value != None:
                return_values.update({setting_key : value})
        return return_values
     
     
    def write_settings(self, value_list):  
        '''
        function writing all relay settings
        Inputs:
        value_list: dictionary consisting of the setting name as key and setting
        value
        ''' 
    
        for value_key in value_list.keys():
            if len(self.setting_list[value_key].subrelay_name) > 0 :
                parent = self.interface.get_element_by_name(\
                                    self.setting_list[value_key].subrelay_name)
            else:
                parent = self.pf_relay           
            element_list = self.interface.get_element_by_name_and_parent(\
                                self.setting_list[value_key].element_name,\
                                        parent) if parent else None     
            if element_list:
                self.interface.set_attribute(element = element_list[0],\
                                    attribute_name = self.setting_list[value_key].\
                                                                       setting_name,
                                    attribute_value = value_list[value_key])
            else:
                pass
 
#===============================================================================
# Overcurrent relay interface
#===============================================================================
  
class PowerFactoryOvercurrentRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF overcurrent relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Enable #1': RelaySetting(subrelay_name = '', element_name = 'I>', setting_name = 'outserv', value = 0),    
        'Characteristic #1': RelaySetting(subrelay_name = '', element_name = 'I>', setting_name = 'pcharac', value = 0),
        'Current #1': RelaySetting(subrelay_name = '', element_name = 'I>', setting_name = 'Ipset', value = 0),
        'Time #1': RelaySetting(subrelay_name = '', element_name = 'I>', setting_name = 'Tpset', value = 0),
        'Enable #2': RelaySetting(subrelay_name = '', element_name = 'I>>', setting_name = 'outserv', value = 0), 
        'Characteristic #2': RelaySetting(subrelay_name = '', element_name = 'I>>', setting_name = 'pcharac', value = 0),
        'Current #2': RelaySetting(subrelay_name = '', element_name = 'I>>', setting_name = 'Ipset', value = 0),
        'Time #2': RelaySetting(subrelay_name = '', element_name = 'I>>', setting_name = 'Tpset', value = 0),
        'Enable #3': RelaySetting(subrelay_name = '', element_name = 'I>>>', setting_name = 'outserv', value = 1), 
        'Current #3': RelaySetting(subrelay_name = '', element_name = 'I>>>', setting_name = 'Ipset', value = 0),
        'Time #3': RelaySetting(subrelay_name = '', element_name = 'I>>>', setting_name = 'Tset', value = 0),
        'Enable #4': RelaySetting(subrelay_name = '', element_name = 'I>>>>', setting_name = 'outserv', value = 1), 
        'Current #4': RelaySetting(subrelay_name = '', element_name = 'I>>>>', setting_name = 'Ipset', value = 0),
        'Time #4': RelaySetting(subrelay_name = '', element_name = 'I>>>>', setting_name = 'Tset', value = 0)
        }   
 
#===============================================================================
# Neutral Overcurrent relay interface
#===============================================================================
  
class PowerFactoryNeutralOvercurrentRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF overcurrent relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Enable #1': RelaySetting(subrelay_name = '', element_name = 'Ig>', setting_name = 'outserv', value = 0),    
        'Characteristic #1': RelaySetting(subrelay_name = '', element_name = 'Ig>', setting_name = 'pcharac', value = 0),
        'Current #1': RelaySetting(subrelay_name = '', element_name = 'Ig>', setting_name = 'Ipset', value = 0),
        'Time #1': RelaySetting(subrelay_name = '', element_name = 'Ig>', setting_name = 'Tpset', value = 0),
        'Enable #2': RelaySetting(subrelay_name = '', element_name = 'Ig>>', setting_name = 'outserv', value = 0), 
        'Current #2': RelaySetting(subrelay_name = '', element_name = 'Ig>>', setting_name = 'Ipset', value = 0),
        'Time #2': RelaySetting(subrelay_name = '', element_name = 'Ig>>', setting_name = 'Tset', value = 0),
        'Enable #3': RelaySetting(subrelay_name = '', element_name = 'Ig>>>', setting_name = 'outserv', value = 1), 
        'Current #3': RelaySetting(subrelay_name = '', element_name = 'Ig>>>', setting_name = 'Ipset', value = 0),
        'Time #3': RelaySetting(subrelay_name = '', element_name = 'Ig>>>', setting_name = 'Tset', value = 0)
        }  
 
  
  
#===============================================================================
# Mho Distance relay interface
#===============================================================================
  
class PowerFactoryMhoDistanceRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF Mho distance relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Phase Phase Mho 1 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 1', setting_name = 'outserv', value = 0),
        'Phase Phase Mho 1 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 1', setting_name = 'idir', value = 0),
        'Phase Phase Mho 1 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 1', setting_name = 'Zm', value = 0),
        'Phase Phase Mho 1 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 1', setting_name = 'phi', value = 0),  
        'Phase Phase Mho 1 Delay': RelaySetting(subrelay_name = '', element_name = 'Mho 1 Delay', setting_name = 'Tdelay', value = 0),     
        'Phase Phase Mho 2 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 2', setting_name = 'outserv', value = 0),
        'Phase Phase Mho 2 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 2', setting_name = 'idir', value = 0),
        'Phase Phase Mho 2 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 2', setting_name = 'Zm', value = 0),
        'Phase Phase Mho 2 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 2', setting_name = 'phi', value = 0),
        'Phase Phase Mho 2 Delay': RelaySetting(subrelay_name = '', element_name = 'Mho 2 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Phase Mho 3 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 3', setting_name = 'outserv', value = 0),
        'Phase Phase Mho 3 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 3', setting_name = 'idir', value = 0),
        'Phase Phase Mho 3 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 3', setting_name = 'Zm', value = 0),
        'Phase Phase Mho 3 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 3', setting_name = 'phi', value = 0),
        'Phase Phase Mho 3 Delay': RelaySetting(subrelay_name = '', element_name = 'Mho 3 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Phase Mho 4 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 4', setting_name = 'outserv', value = 0),
        'Phase Phase Mho 4 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 4', setting_name = 'idir', value = 0),
        'Phase Phase Mho 4 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 4', setting_name = 'Zm', value = 0),
        'Phase Phase Mho 4 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Mho 4', setting_name = 'phi', value = 0),
        'Phase Phase Mho 4 Delay': RelaySetting(subrelay_name = '', element_name = 'Mho 4 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Earth Mho 1 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 1', setting_name = 'outserv', value = 0),
        'Phase Earth Mho 1 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 1', setting_name = 'idir', value = 0),
        'Phase Earth Mho 1 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 1', setting_name = 'Zm', value = 0),
        'Phase Earth Mho 1 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 1', setting_name = 'phi', value = 0),
        'Phase Earth Mho 2 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 2', setting_name = 'outserv', value = 0),
        'Phase Earth Mho 2 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 2', setting_name = 'idir', value = 0),
        'Phase Earth Mho 2 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 2', setting_name = 'Zm', value = 0),
        'Phase Earth Mho 2 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 2', setting_name = 'phi', value = 0),
        'Phase Earth Mho 3 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 3', setting_name = 'outserv', value = 0),
        'Phase Earth Mho 3 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 3', setting_name = 'idir', value = 0),
        'Phase Earth Mho 3 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 3', setting_name = 'Zm', value = 0),
        'Phase Earth Mho 3 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 3', setting_name = 'phi', value = 0),
        'Phase Earth Mho 4 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 4', setting_name = 'outserv', value = 0),
        'Phase Earth Mho 4 Tripping Direction': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 4', setting_name = 'idir', value = 0),
        'Phase Earth Mho 4 Replica Impedance': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 4', setting_name = 'Zm', value = 0),
        'Phase Earth Mho 4 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Mho 4', setting_name = 'phi', value = 0),
        'Phase Directional Angle': RelaySetting(subrelay_name = '', element_name = 'Phase Directional', setting_name = 'phi', value = 0),
        'Ground Directional Angle': RelaySetting(subrelay_name = '', element_name = 'Ground Directional', setting_name = 'phi', value = 0),
        'Starting Phase Current #1': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ip1', value = 0),
        'Starting Phase Voltage Current #1': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'u', value = 0),
        'Starting Phase Current #2': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ip2', value = 0),
        'Starting Earth Current': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ie', value = 0),
        'k0': RelaySetting(subrelay_name = '', element_name = 'Polarizing', setting_name = 'k0', value = 0),
        'k0 Angle': RelaySetting(subrelay_name = '', element_name = 'Polarizing', setting_name = 'phik0', value = 0)
        }  

#===============================================================================
# Polygonal Distance relay interface
#===============================================================================
  
class PowerFactoryPolygonalDistanceRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF Polygonal distance relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Phase Phase Polygonal 1 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 1', setting_name = 'outserv', value = 0),
        'Phase Phase Polygonal 1 X': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 1', setting_name = 'cpXmax', value = 0),
        'Phase Phase Polygonal 1 R': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 1', setting_name = 'cpRmax', value = 0),
        'Phase Phase Polygonal 1 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 1', setting_name = 'phi', value = 0),
        'Phase Phase Polygonal 1 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 1', setting_name = 'beta', value = 0),
        'Phase Phase Polygonal 1 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal  1 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Phase Polygonal 2 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 2', setting_name = 'outserv', value = 0),
        'Phase Phase Polygonal 2 X': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 2', setting_name = 'cpXmax', value = 0),
        'Phase Phase Polygonal 2 R': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 2', setting_name = 'cpRmax', value = 0),
        'Phase Phase Polygonal 2 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 2', setting_name = 'phi', value = 0),
        'Phase Phase Polygonal 2 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 2', setting_name = 'beta', value = 0),
        'Phase Phase Polygonal 2 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal  2 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Phase Polygonal 3 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 3', setting_name = 'outserv', value = 0),
        'Phase Phase Polygonal 3 X': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 3', setting_name = 'cpXmax', value = 0),
        'Phase Phase Polygonal 3 R': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 3', setting_name = 'cpRmax', value = 0),
        'Phase Phase Polygonal 3 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 3', setting_name = 'phi', value = 0),
        'Phase Phase Polygonal 3 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 3', setting_name = 'beta', value = 0),
        'Phase Phase Polygonal 3 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal  3 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Phase Polygonal 4 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 4', setting_name = 'outserv', value = 0),
        'Phase Phase Polygonal 4 X': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 4', setting_name = 'cpXmax', value = 0),
        'Phase Phase Polygonal 4 R': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 4', setting_name = 'cpRmax', value = 0),
        'Phase Phase Polygonal 4 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 4', setting_name = 'phi', value = 0),
        'Phase Phase Polygonal 4 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal 4', setting_name = 'beta', value = 0),
        'Phase Phase Polygonal 4 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-Ph Polygonal  4 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Earth Polygonal 1 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1', setting_name = 'outserv', value = 0),
        'Phase Earth Polygonal 1 X': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1', setting_name = 'cpXmax', value = 0),
        'Phase Earth Polygonal 1 R': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1', setting_name = 'cpRmax', value = 0),
        'Phase Earth Polygonal 1 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1', setting_name = 'phi', value = 0),
        'Phase Earth Polygonal 1 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1', setting_name = 'beta', value = 0),
        'Phase Earth Polygonal 1 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 1 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Earth Polygonal 2 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2', setting_name = 'outserv', value = 0),
        'Phase Earth Polygonal 2 X': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2', setting_name = 'cpXmax', value = 0),
        'Phase Earth Polygonal 2 R': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2', setting_name = 'cpRmax', value = 0),
        'Phase Earth Polygonal 2 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2', setting_name = 'phi', value = 0),
        'Phase Earth Polygonal 2 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2', setting_name = 'beta', value = 0),
        'Phase Earth Polygonal 2 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 2 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Earth Polygonal 3 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3', setting_name = 'outserv', value = 0),
        'Phase Earth Polygonal 3 X': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3', setting_name = 'cpXmax', value = 0),
        'Phase Earth Polygonal 3 R': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3', setting_name = 'cpRmax', value = 0),
        'Phase Earth Polygonal 3 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3', setting_name = 'phi', value = 0),
        'Phase Earth Polygonal 3 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3', setting_name = 'beta', value = 0),
        'Phase Earth Polygonal 3 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 3 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Earth Polygonal 4 Out service': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4', setting_name = 'outserv', value = 0),
        'Phase Earth Polygonal 4 X': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4', setting_name = 'cpXmax', value = 0),
        'Phase Earth Polygonal 4 R': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4', setting_name = 'cpRmax', value = 0),
        'Phase Earth Polygonal 4 Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4', setting_name = 'phi', value = 0),
        'Phase Earth Polygonal 4 X Angle': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4', setting_name = 'beta', value = 0),
        'Phase Earth Polygonal 4 delay': RelaySetting(subrelay_name = '', element_name = 'Ph-E Polygonal 4 Delay', setting_name = 'Tdelay', value = 0),
        'Phase Directional Angle': RelaySetting(subrelay_name = '', element_name = 'Phase Directional', setting_name = 'phi', value = 0),
        'Ground Directional Angle': RelaySetting(subrelay_name = '', element_name = 'Ground Directional', setting_name = 'phi', value = 0),
        'Starting Phase Current #1': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ip1', value = 0),
        'Starting Phase Voltage Current #1': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'u', value = 0),
        'Starting Phase Current #2': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ip2', value = 0),
        'Starting Earth Current': RelaySetting(subrelay_name = '', element_name = 'Starting', setting_name = 'ie', value = 0),
        'k0': RelaySetting(subrelay_name = '', element_name = 'Polarizing', setting_name = 'k0', value = 0),
        'k0 Angle': RelaySetting(subrelay_name = '', element_name = 'Polarizing', setting_name = 'phik0', value = 0)
        }  


#===============================================================================
# Polygonal OOS relay interface
#===============================================================================
  
class PowerFactoryPolygonalOOSRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF polygonal OOS relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Outer Polygonal X': RelaySetting(subrelay_name = '', element_name = 'Outer Poly', setting_name = 'Xmax', value = 0),
        'Outer Polygonal R': RelaySetting(subrelay_name = '', element_name = 'Outer Poly', setting_name = 'Rmax', value = 0),
        'Outer Polygonal -R': RelaySetting(subrelay_name = '', element_name = 'Outer Poly', setting_name = 'Rmin', value = 0),
        'Outer Polygonal Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Outer Poly', setting_name = 'phi', value = 0),
        'Outer Polygonal X Angle': RelaySetting(subrelay_name = '', element_name = 'Outer Poly', setting_name = 'beta', value = 0),
        'Inner Polygonal X': RelaySetting(subrelay_name = '', element_name = 'Inner Poly', setting_name = 'Xmax', value = 0),
        'Inner Polygonal R': RelaySetting(subrelay_name = '', element_name = 'Inner Poly', setting_name = 'Rmax', value = 0),
        'Inner Polygonal -R': RelaySetting(subrelay_name = '', element_name = 'Inner Poly', setting_name = 'Rmin', value = 0),
        'Inner Polygonal Relay Angle': RelaySetting(subrelay_name = '', element_name = 'Inner Poly', setting_name = 'phi', value = 0),
        'Inner Polygonal X Angle': RelaySetting(subrelay_name = '', element_name = 'Inner Poly', setting_name = 'beta', value = 0),
        }   


  
#===============================================================================
# Over Frequency relay interface
#=============================================================================== 
       
class PowerFactoryOverFrequencyRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF over frequency relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Pickup #1': RelaySetting(subrelay_name = '', element_name = 'F>1', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #1': RelaySetting(subrelay_name = '', element_name = 'F>1', setting_name = 'Tpset', value = 0),
        'Pickup #2': RelaySetting(subrelay_name = '', element_name = 'F>2', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #2': RelaySetting(subrelay_name = '', element_name = 'F>2', setting_name = 'Tpset', value = 0),
        'Pickup #3': RelaySetting(subrelay_name = '', element_name = 'F>3', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #3': RelaySetting(subrelay_name = '', element_name = 'F>3', setting_name = 'Tpset', value = 0),
        'Pickup #4': RelaySetting(subrelay_name = '', element_name = 'F>4', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #4': RelaySetting(subrelay_name = '', element_name = 'F>4', setting_name = 'Tpset', value = 0),
        }
        
        
#===============================================================================
# Under Frequency relay interface
#=============================================================================== 
       
class PowerFactoryUnderFrequencyRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF under frequency relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Pickup #1': RelaySetting(subrelay_name = '', element_name = 'F<1', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #1': RelaySetting(subrelay_name = '', element_name = 'F<1', setting_name = 'Tpset', value = 0),
        'Pickup #2': RelaySetting(subrelay_name = '', element_name = 'F<2', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #2': RelaySetting(subrelay_name = '', element_name = 'F<2', setting_name = 'Tpset', value = 0),
        'Pickup #3': RelaySetting(subrelay_name = '', element_name = 'F<3', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #3': RelaySetting(subrelay_name = '', element_name = 'F<3', setting_name = 'Tpset', value = 0),
        'Pickup #4': RelaySetting(subrelay_name = '', element_name = 'F<4', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #4': RelaySetting(subrelay_name = '', element_name = 'F<4', setting_name = 'Tpset', value = 0),
        }
        
        
#===============================================================================
# Over Voltage relay interface
#=============================================================================== 
       
class PowerFactoryOverVoltageRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF over voltage relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Pickup #1': RelaySetting(subrelay_name = '', element_name = 'Upp>1', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #1': RelaySetting(subrelay_name = '', element_name = 'Upp>1', setting_name = 'Tpset', value = 0),
        'Pickup #2': RelaySetting(subrelay_name = '', element_name = 'Upp>2', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #2': RelaySetting(subrelay_name = '', element_name = 'Upp>2', setting_name = 'Tpset', value = 0),
        'Pickup #3': RelaySetting(subrelay_name = '', element_name = 'Upp>3', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #3': RelaySetting(subrelay_name = '', element_name = 'Upp>3', setting_name = 'Tpset', value = 0),
        'Pickup #4': RelaySetting(subrelay_name = '', element_name = 'Upp>4', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #4': RelaySetting(subrelay_name = '', element_name = 'Upp>4', setting_name = 'Tpset', value = 0)
        }
        
        
#===============================================================================
# Under Voltage relay interface
#=============================================================================== 
       
class PowerFactoryUnderVoltageRelayInterface(PowerFactoryRelayInterface):
    '''
    class interfacing the PF under voltage relay
    '''
    def __init__(self, interface, pf_relay):
        '''
        Constructor
        '''
        PowerFactoryRelayInterface.__init__(self, interface, pf_relay)
        self.setting_list = {
        'Pickup #1': RelaySetting(subrelay_name = '', element_name = 'Upp<1', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #1': RelaySetting(subrelay_name = '', element_name = 'Upp<1', setting_name = 'Tpset', value = 0),
        'Pickup #2': RelaySetting(subrelay_name = '', element_name = 'Upp<2', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #2': RelaySetting(subrelay_name = '', element_name = 'Upp<2', setting_name = 'Tpset', value = 0),
        'Pickup #3': RelaySetting(subrelay_name = '', element_name = 'Upp<3', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #3': RelaySetting(subrelay_name = '', element_name = 'Upp<3', setting_name = 'Tpset', value = 0),
        'Pickup #4': RelaySetting(subrelay_name = '', element_name = 'Upp<4', setting_name = 'Ipsetr', value = 0),    
        'Time Delay #4': RelaySetting(subrelay_name = '', element_name = 'Upp<4', setting_name = 'Tpset', value = 0)
        }