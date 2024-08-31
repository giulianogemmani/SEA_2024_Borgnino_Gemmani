'''
Created on 1 October 2018

@author: AMB
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
    Interface between PSET and PowerFactory
    '''
    # relay types as listed in the PF relay typ dialog
    class RelayType(Enum):
        OVERCURRENT     = 'Overcurrent'
        DIRECTIONAL     = 'Directional'
        DISTANCE        = 'Distance'
        FREQUENCY       = 'Frequency'
        VOLTAGE         = 'Voltage'
        DIFFERENTIAL    = 'Differential'
        SUB_FUNCTION    = 'Subfunction'
        RECLOSER        = 'Recloser'
        SECTIONALIZER   = 'Sectionalizer'
        ANY = 'Any'

    # shc_trace calculation types as available inside PF
    class SHC_Mode(IntEnum):
        VDE0101         = 0
        IEC60909        = 1
        ANSI            = 2
        COMPLETE        = 3
        IEC61363        = 4
        IEC61660DC      = 5
        ANSIIEEE946DC   = 6
        DINEN61660DC    = 7

    # fault types for the short circuit trace 
    FautType_Shctrace = {
        '3rst'          : 0,
        '2psc'          : 1,
        'spgf'          : 2,
        '2pgf'          : 3
    }  
    
    # TD diagram methods 
    Time_Distance_Diagram_Method = {
        'short_circuit_sweep'    : 'iec',
        'kilometrical'          : 'len'
    } 
    
    Time_Current_Diagram_Line_Type = {
        "Undefined"                     :  0,
        "Phase Current"                 :  1,
        "Earth Current(3I0)"            :  2,
        "Zero Sequence Current(I0)"     :  3,
        "Negative Sequence Current(I2)" :  4
    }
    
    Time_Current_Diagram_Displayed_Relays = {
        "All"                      :  0,
        "Phase Relays"             :  1,
        "Phase & Earth Relays"     :  2,
        "Earth Relays"             :  3,
        "Negative Sequence Relays" :  4
    }
    
    # element id used to retrieve the relay phase, ground etc element types 
    ElementType = {
        'phase'         : ['Phase Current (3ph)',
                           'Phase Current (1ph)', 
                           'Phase A Current',
                           'Phase B Current',
                           'Phase C Current'],
        'ground'        : ['Earth Current (3*I0)',
                           'Sensitive Earth Current (3*I0)',
                           'Zero Sequence Current (I0)'],
        'any'           : ['any id']
    }


    # id to identify which relays must be returned in a connection element (side connection 1 or 2 or etc)
    class Side(IntEnum):
        Side_1  = 0
        Side_2  = 1    
        Side_3  = 2 
        Side_4  = 3 
        AllSide = 99

    Slot = namedtuple("Slot", "slot_1 slot_2 slot_3 slot_4")
    # relationship betweden the PF element types and the variables where to get the short circuit current
    i_slot_name_list = {
        'ElmLne' : Slot(slot_1 = 'm:Ikss:bus1'  , slot_2 = 'm:Ikss:bus2'  , slot_3 = '', slot_4 = ''),
        'ElmTr2' : Slot(slot_1 = 'm:Ikss:bushv' , slot_2 = 'm:Ikss:buslv' , slot_3 = '', slot_4 = ''),
        'ElmTr3' : Slot(slot_1 = 'm:Ikss:bushv' , slot_2 = 'm:Ikss:busmv' , slot_3 = 'm:Ikss:buslv', slot_4 = ''),
        'ElmTr4' : Slot(slot_1 = 'm:Ikss:bushv0', slot_2 = 'm:Ikss:buslv1', slot_3 = 'm:Ikss:buslv2', slot_4 = 'm:Ikss:buslv3')
    }
    
    PF_Attribute = namedtuple("PF_Attribute", "attr_1 attr_2 attr_3 attr_4")
    # relationship betweden the PF element types and the variables where to get the connected objects
    pf_attribute_name_list = {
        'ElmLne'   : PF_Attribute(attr_1 = 'bus1'  , attr_2 = 'bus2'  , attr_3 = '', attr_4 = ''),
        'ElmTr2'   : PF_Attribute(attr_1 = 'bushv' , attr_2 = 'buslv' , attr_3 = '', attr_4 = ''),
        'ElmTr3'   : PF_Attribute(attr_1 = 'bushv' , attr_2 = 'busmv' , attr_3 = 'buslv', attr_4 = ''),
        'ElmTr4'   : PF_Attribute(attr_1 = 'bushv0', attr_2 = 'buslv1', attr_3 = 'buslv2', attr_4 = 'buslv3'),
        'ElmCoup'    : PF_Attribute(attr_1 = 'bus1' , attr_2 = 'bus2' , attr_3 = '', attr_4 = ''),
        'ElmLod'     : PF_Attribute(attr_1 = 'bus1' , attr_2 = ''  , attr_3 = '', attr_4 = ''),
        'ElmShnt'    : PF_Attribute(attr_1 = 'bus1' , attr_2 = ''  , attr_3 = '', attr_4 = ''),
        'ElmSym'     : PF_Attribute(attr_1 = 'bus1' , attr_2 = ''  , attr_3 = '', attr_4 = ''),
        'ElmGenstat' : PF_Attribute(attr_1 = 'bus1' , attr_2 = ''  , attr_3 = '', attr_4 = '')
    }
    
    # create the selection criteria dictionary
    Criteria = namedtuple("Criteria", "itemlist function")

#=========================================================================
# Initialization method
#=========================================================================

    def __init__(self):
        '''
        Constructor 
        '''
        self.default_shc_slot_name = ['m:Ikss:bus1','m:Ikss:bus2']
        self.default_pf_attribute_name = ['bus1','bus2']
        
        self.shc_trace = None
        
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
        info retrived from the windows registry at  HKEY_LOCAL_MACHINE\\
        SOFTWARE\\WOW6432Node\\DIgSILENT GmbH
        if more installations i.e 2018 SP2, SP3 etc are present the latest one 
        is retrieved
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
        it returns a void string if the object type is not DataObject
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


    def get_class_name_of(self, element):
        '''
        function returning the class name (i.e. ElmTerm) of the given PF object
        '''
        try:
            return element.GetClassName()
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


    def get_element_by_name(self, element_name):
        '''
        function getting the PF element having the given name or part of the 
        given name
        '''
        try:
            return self.app.GetCalcRelevantObjects(element_name + "*.*")
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        
    
    def get_element_by_name_and_parent(self, element_name, element_parent):
        '''
        function getting the PF element having the given name and parent
        '''
        try:
            return element_parent.GetContents(element_name +".*")
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        
    
    def get_element_by_foreign_key(self, element_foreign_key):
        '''
        function getting the PF element having the given foreign key
        '''
        try:
            return self.app.SearchObjectByForeignKey(element_foreign_key)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        

    def get_relays(self, relay_type=[RelayType.ANY]):
        '''
        Function returning the relays of the given type (distance, overcurrent, 
        differential etc) list in the active project
        '''
        relay_type_values = list(map(lambda typval: typval.value, relay_type))             
        return [relay for relay in self.app.GetCalcRelevantObjects("*.ElmRelay") \
                if ( relay.typ_id and relay.c_category in relay_type_values) or\
                 relay_type[0].name == self.RelayType.ANY]
        # the relay category is defined only if the relay has a type
#         return [line for line in pf_lines
#                 if all([len(criteria.itemlist) == 0 or criteria.function(line, key) 
#                     in criteria.itemlist for key, criteria \
#                                                 in research_criteria.items()])]                

    def get_relay_cubicle_of(self, relay):
        ''' 
        Function returning the cubicle where the relay is located
        '''
        try:
            return relay.GetAttribute('fold_id')
            #return relay.GetParent()   #not to use, too slow
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        


    def get_relay_breaker_of(self, relay):
        '''
        function getting the breaker located in the cubicle where the given 
        relay is
        '''
        return [breaker for breaker in self.get_relay_cubicle_of(relay).\
                                    GetContents("*.ElmCoup, *.StaSwitch")]
    
    
    def get_cubicle_relay_of(self, cubicle):
        '''
        function getting the relay(s) contained in the given cubicle
        '''
        return cubicle.GetContents("*.ElmRelay")
    
    
    def get_cubicle_CT_of(self, cubicle):
        '''
        function getting the ct(s) contained in the given cubicle
        '''
        return cubicle.GetContents("*.StaCt")
    
    
    def get_cubicle_VT_of(self, cubicle):
        '''
        function getting the vt(s) contained in the given cubicle
        '''
        return cubicle.GetContents("*.StaVt")
    
    
    def get_ct_ratio(self, ct):
        '''
        return the CT ratio of the given CT
        '''
        return ct.ptapset / ct.stapset
    
    
    def get_vt_ratio(self, vt):
        '''
        return the VT ratio of the given VT
        '''
        return vt.ptapset / vt.stapset
    
    
    def get_breaker_operating_time(self, breaker):
        '''
        function retrieving the operating time of the given breaker.
        The breaker can be a StaSwitch and in this case the'Tprot' variable is 
        returned or a ElmCoup and then the 'Tb' variable is returmed
        '''
        try:
            return round(breaker.GetAttribute('Tb'), 4) \
            if self.get_class_name_of(breaker) == 'ElmCoup' else \
            round(breaker.GetAttribute('Tprot'), 4)
        except Exception as e:
            if breaker != None: 
                self.print("Script data type error (brk oper time): " + str(e))
                self.print(traceback.format_exc())
            return 0
     
    
    def get_breaker_bus(self, breaker):
        '''
        function retrieving the bus of the given breaker. 
        '''
        try:
            return self.get_parent(breaker.GetAttribute('fold_id'))
        except Exception as e:
            if breaker != None: 
                self.print("Script data type error (breaker bus): " + str(e))
                self.print(traceback.format_exc())
            return 0
        
        
    def get_relay_busbar_of(self, relay):
        ''' 
        Function returning the bus bar of the cubicle where the relay is located
        '''
        try:
            return relay.GetAttribute('cn_bus')
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_relay_protected_item_of(self, relay):
        ''' 
        Function returning the object (line, transformer ) protected by the 
        relay passed as parameter
        '''
        try:
            return relay.GetAttribute('cbranch')
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
    
        
    def get_relay_manufacturer_of(self, relay):
        '''
        function returning the manufacturer of the relay passed as parameter
        Note: no check that the passed parameter is a relay
        '''
        type_relay = self.get_attribute(relay, 'typ_id')
        return self.get_attribute(type_relay, 'manuf') if type_relay is not None else ""
    
    
    def get_relay_model_name_of(self, relay):
        '''
        function returning the model of the relay passed as parameter
        Note: no check that the passed parameter is a relay
        '''
        relay_type = self.get_attribute(relay, 'typ_id')
        return self.get_attribute(relay_type, 'loc_name') if relay_type is not None else ""
    
    
    def get_relay_type_of(self, relay):
        '''
        function returning the model of the relay passed as parameter
        Note: no check that the passed parameter is a relay
        '''
        try:
            return self.get_attribute(relay, 'typ_id')
        except Exception as e:
            self.print("Script data type error (get_relay_type_of): " + str(e))
            self.print(traceback.format_exc())
            return None
                
     
    def get_relay_category_of(self, relay):
        '''
        function returning the category (distance, overcurrent ect) of the relay 
        passed as parameter
        Note: no check that the passed parameter is a relay
        '''
        try:
            return self.get_attribute(relay, 'c_category')
        except Exception as e:
            self.print("Script data type error (get_relay_category_of): " + str(e))
            self.print(traceback.format_exc())
            return None 
        
    
    def get_relay_time_defined_overcurrent_elements_thresholds_of(self, relay,
                                            element_type = ElementType['any']):
        '''
        function returning a list containing the threshold values in primary 
        amps of all ioc elements of the given type (phase or ground or any)
        also the element in thr subrerlays are considered 
        if no element has been found a void list is returned
        '''
        return [round(self.get_attribute(ioc_element, 'Ipsetr'), 3) \
                for ioc_element in relay.GetContents('*.RelIoc', 1) \
                if self.is_out_of_service(ioc_element) == False and \
                (element_type == self.ElementType['any'] or \
                self.get_attribute(ioc_element, 'c_type') in element_type)]
        
        
    def get_relay_tripping_element_of(self, relay):
        '''
        function returning the tripped elements list of the given relay
        '''    
        element_list =  relay.GetContents('*.Relioc',    1) + \
                        relay.GetContents('*.RelToc',    1) + \
                        relay.GetContents('*.RelDismho', 1) + \
                        relay.GetContents('*.RelDispoly',1)
        return [element for element in element_list \
                if self.is_out_of_service(element) == False and \
                self.get_attribute(element, 's:yout') < \
                                            self.get_relay_NO_TRIP_constant()]
      
      
    def get_relay_current_measures_of(self, measurement):
        '''
        function returning the phase and the zero sequence currents of the given
         measurement element
        ''' 
        if self.is_out_of_service(measurement) == False: 
            return (self.get_attribute(measurement, 's:I_A'),
                    self.get_attribute(measurement, 's:I_B'),
                    self.get_attribute(measurement, 's:I_C'),
                    self.get_attribute(measurement, 's:I0x3')) 
        else:
            return (0, 0, 0, 0)
        
    
    def get_relay_complex_current_measures_of(self, measurement):
        '''
        function returning the phase and the zero sequence currents of the given
         measurement element as complex values
        ''' 
        if self.is_out_of_service(measurement) == False: 
            return (complex(self.get_attribute(measurement, 's:Ir_A'),
                            self.get_attribute(measurement, 's:Ii_A')),
                    complex(self.get_attribute(measurement, 's:Ir_B'),
                            self.get_attribute(measurement, 's:Ii_B')),
                    complex(self.get_attribute(measurement, 's:Ir_C'),
                            self.get_attribute(measurement, 's:Ii_C')),
                    complex(self.get_attribute(measurement, 's:I0x3r'),
                            self.get_attribute(measurement, 's:I0x3i'))) 
        else:
            return (complex(0,0), complex(0,0), complex(0,0), complex(0,0))


    def get_relay_measurement_element(self, relay):
        '''
        get the measurement element of the given relay
        '''
        # only these types provide the 3 phase currents
        try:
            measurement_allowed_types = ['3rms', '3pui', '3dui', 'abbd']  
            measurements_list =  relay.GetContents('*.RelMeasure', 1)
            for measurement in measurements_list:
                measurement_type = self.get_attribute(measurement, 'typ_id')
                if measurement_type != None and \
                self.get_attribute(measurement_type, 'atype') \
                in measurement_allowed_types:
                    return measurement
            raise Exception('Measurement search failure')  
        except Exception as e:      
            self.print("No measurement element found in : " +\
                        self.get_full_name_of(relay) + " "  + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_relay_polarizing_element(self, relay):
        '''
        get the polarizing element of the given relay
        '''
        try:      
            polarizings_list =  relay.GetContents('*.RelZpol', 1)
            if polarizings_list:
                return polarizings_list[0]
            else:
                raise Exception('Polarizing  search failure')  
        except Exception as e:      
            self.print("No polarizing element found in : " +\
                        self.get_full_name_of(relay) + " "  + str(e))
            self.print(traceback.format_exc())
            return None
        
    
    def get_relay_secondary_z_measures_of(self, polarizing):
        '''
        function returning the phase Z magnitude of the given polarizing element
        ''' 
        try: 
            if self.is_out_of_service(polarizing) == False: 
                R_A = self.get_attribute(polarizing, 's:R_A')
                R_B = self.get_attribute(polarizing, 's:R_B')
                R_C = self.get_attribute(polarizing, 's:R_C')
                X_A = self.get_attribute(polarizing, 's:X_A')
                X_B = self.get_attribute(polarizing, 's:X_B')
                X_C = self.get_attribute(polarizing, 's:X_C')
                return (round(sqrt(R_A*R_A + X_A*X_A),2),
                        round(sqrt(R_B*R_B + X_B*X_B),2),
                        round(sqrt(R_C*R_C + X_C*X_C),2))
    #             return (round(self.get_attribute(polarizing, 'c:Zlp_A'),2),
    #                     round(self.get_attribute(polarizing, 'c:Zlp_B'),2),
    #                     round(self.get_attribute(polarizing, 'c:Zlp_C'),2)) 
            else:
                return (0, 0, 0)
        except Exception as e:      
            self.print("Problem getting the polarizing Z in : " +\
                        self.get_full_name_of(polarizing) + " "  + str(e))
            self.print(traceback.format_exc())
            return (0, 0, 0)
        
        
    def get_relay_secondary_complex_z_measures_of(self, polarizing):
        '''
        function returning the phase Z as complex of the given polarizing element
        '''
        try:  
            if self.is_out_of_service(polarizing) == False: 
                R_A = self.get_attribute(polarizing, 's:R_A')
                R_B = self.get_attribute(polarizing, 's:R_B')
                R_C = self.get_attribute(polarizing, 's:R_C')
                X_A = self.get_attribute(polarizing, 's:X_A')
                X_B = self.get_attribute(polarizing, 's:X_B')
                X_C = self.get_attribute(polarizing, 's:X_C')
                return (complex(round(R_A, 4), round(X_A, 4)),
                        complex(round(R_B, 4), round(X_B, 4)),
                        complex(round(R_C, 4), round(X_C, 4)))
            else:
                return (0, 0, 0)
        except Exception as e:      
            self.print("Problem getting the polarizing Z in : " +\
                        self.get_full_name_of(polarizing) + " "  + str(e))
            self.print(traceback.format_exc())
            return (0, 0, 0)
        

    def get_busbars(self, research_criteria = {"": Criteria([], get_attribute)}):
        '''
        Function returning all busbars in the active project which respect the
         research criteria
        '''
        pf_busbars = self.app.GetCalcRelevantObjects("*.ElmTerm")
        return [busbar for busbar in pf_busbars
                if all([len(criteria.itemlist) == 0 or criteria.function(busbar,
             key) in criteria.itemlist for key, criteria \
                                                in research_criteria.items()])]


    def get_lines(self, research_criteria = {"": Criteria([], get_attribute)}):
        '''
        Function returning all lines in the active project which respect the 
        research criteria
        '''
        pf_lines = self.app.GetCalcRelevantObjects("*.ElmLne")
        return [line for line in pf_lines
                if all([len(criteria.itemlist) == 0 or criteria.function(line, key) 
                    in criteria.itemlist for key, criteria \
                                                in research_criteria.items()])]


    def get_generators(self, research_criteria = {"": Criteria([], get_attribute)}):
        '''
        Function returning all generators in the active project which respect the
         research criteria
        '''
        pf_generators = [generator for generator in\
                         self.app.GetCalcRelevantObjects("*.ElmSym, *.ElmAsm")\
                         if self.get_attribute(generator, "i_mot") == 0] 
        return [generator for generator in pf_generators
                if all([len(criteria.itemlist) == 0 or criteria.function(generator,
             key) in criteria.itemlist for key, criteria \
                                                in research_criteria.items()])]
        
    
    def get_transformers(self, research_criteria = {"": Criteria([], get_attribute)}):
        '''
        Function returning all transformers in the active project which respect the
         research criteria
        '''
        pf_transformers = [transformer for transformer in\
                self.app.GetCalcRelevantObjects("*.ElmTr*")]     
        return [transformer for transformer in pf_transformers
                if all([len(criteria.itemlist) == 0 or criteria.function(transformer,
             key) in criteria.itemlist for key, criteria \
                                                in research_criteria.items()])]    
        
    
    def get_transformer_cubicle_of(self, transformer):
        '''
        function returning in a list all transformer cubicles
        '''
        class_name = self.get_class_name_of(transformer)
        if class_name == "ElmTr2":
            return [transformer.GetAttribute('bushv'),\
                    transformer.GetAttribute('buslv')]
        elif class_name == "ElmTr3":
            return [transformer.GetAttribute('bushv'),\
                    transformer.GetAttribute('busmv'),\
                    transformer.GetAttribute('buslv')]
        elif class_name == "ElmTr4":
            return [transformer.GetAttribute('bush0'),\
                    transformer.GetAttribute('busl1'),\
                    transformer.GetAttribute('busl2'),\
                    transformer.GetAttribute('busl3')]
        else:
            return []
        
    
    def get_transformer_voltages_of(self, transformer):
        '''
        function returning in a list all transformer voltages
        '''
        class_name = self.get_class_name_of(transformer)
        if class_name == "ElmTr2":
            return [self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('bushv'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('buslv')))]
        elif class_name == "ElmTr3":
            return [self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('bushv'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('busmv'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('buslv')))]
        elif class_name == "ElmTr4":
            return [self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('bush0'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('busl1'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('busl2'))),\
                    self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('busl3')))]
        else:
            return []

    def get_transformer_voltage_of(self, transformer):
        '''
        function returning the transformer highest voltage
        '''
        class_name = self.get_class_name_of(transformer)
        if class_name == "ElmTr2" or class_name == "ElmTr3":
            return self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('bushv')))
        elif class_name == "ElmTr4":
            return self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute('bush0')))
        else:
            return 0
        
    
    def get_transformer_winding_voltage_of(self, transformer, winding):
        '''
        function returning the transformer voltage of the given winding
        '''
        class_name = self.get_class_name_of(transformer)
        if class_name == "ElmTr2" or class_name == "ElmTr3" or class_name == "ElmTr4":
            return self.get_busbar_rated_voltage_in(self.get_cubicle_busbar_of\
                                            (transformer.GetAttribute(winding)))
        else:
            return 0


    def get_line_length_of_(self, line):
        '''
        function returning the lenght of the given line
        '''
        try:
            return round(line.GetAttribute('dline'), 4)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_line_fault_removal_time_of(self, line):
        '''
        function returning the time of the fault removal for the given line
        '''
        try:
            return round(line.GetAttribute('m:Tfct:bus1'), 3)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_available_zones(self):
        '''
        Function returning a list of DataObject which are the available zones
        '''
        zone_folder = self.app.GetDataFolder("ElmZone", 0)
        if zone_folder != None:
            return zone_folder.GetContents("*.ElmZone")


    def get_available_areas(self):
        '''
        Function returning a list of DataObject which are the available areas
        '''
        area_folder = self.app.GetDataFolder("ElmArea", 0)
        if area_folder != None:
            return area_folder.GetContents("*.ElmArea")


    def get_available_voltages(self):
        '''
        Function returning a list of all available voltages
        '''
        return list(set([x.uknom for x \
                         in self.app.GetCalcRelevantObjects("*.ElmTerm")]))


    def get_relay_tripping_time_of(self, relay, include_breaker_time = True):
        '''
        Function returning the tripping time of the relay passed as parameter
        '''
        try:
            breaker_list = self.get_relay_breaker_of(relay)
            breaker = breaker_list[0] if len(breaker_list) >0 else None
            return round(relay.GetAttribute('s:yout') + \
                        (self.get_breaker_operating_time(breaker) \
                        if breaker != None and include_breaker_time else 0),3) \
                if self.is_out_of_service(relay) == False else \
                                            self.get_relay_NO_TRIP_constant()
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            self.print(relay.GetFullName())
            return 0
    
    
    def get_relay_NO_TRIP_constant(self):
        '''
        function returning the time constant used inside PF to declare the no 
        trip condition
        '''
        return 9999.999
    
    
    def get_available_grids(self):
        '''
        Function returning a list of DataObject which are the available grids
        '''
        area_folder = self.app.GetDataFolder("ElmArea", 0)
        if area_folder != None:
            network_data = area_folder.GetParent()
            if network_data != None:
                return network_data.GetContents("*.ElmNet")

    
    def get_parent(self, element):
        '''
        Function returning the parent of the given element
        '''
        try:
            return element.GetParent()
        except Exception as e:
            self.print("Script data type error (get parent): " + str(e))
            self.print(traceback.format_exc())
            return 0


    def get_available_paths(self):
        '''
        Function returning a list of DataObject which are the available paths
        '''
        area_folder = self.app.GetDataFolder("ElmArea", 0)
        if area_folder != None:
            network_data = area_folder.GetParent()
            if network_data != None:
                path_folder = network_data.GetContents("*.IntPath")
                if path_folder[0] != None:
                    return path_folder[0].GetContents("*.SetPath")


    def get_available_busbars(self):
        '''
        Function returning a list of DataObject which are the available busbars
        '''
        return self.app.GetCalcRelevantObjects("*.ElmTerm")
    
    
    def get_available_shunts(self):
        '''
        Function returning a list of DataObject which are the available shunts
        '''
        return self.app.GetCalcRelevantObjects("*.ElmShnt")
    
    
    def get_available_generators(self):
        '''
        Function returning a list of DataObject which are the available generators
        '''
        return self.app.GetCalcRelevantObjects("*.ElmSym")

    
    def get_generator_voltage(self, generator, key = ''):
        '''
        get the voltage of the given generator
        '''
        try:
            return self.get_attribute(generator.typ_id, 'ugn')
        except Exception as e:
            self.print("Script data type error (get generator voltage): " + str(e))
            self.print(traceback.format_exc())
            return 0


    def get_area_bus_voltage(self, line, key = ''):
        '''
        auxiliary fuction which returns the rated voltage of the bus found by 
        the get_area_bus_of function
        '''
        try:
            return self.get_area_bus_of(line).GetAttribute('uknom')
        except Exception as e:
            self.print("Script data type error (get_area_bus_voltage): " + str(e))
            self.print(traceback.format_exc())
            return 0


    def get_content(self, pf_object, search_string = '*.*', recursive = 0):
        '''
        return a list with all elements contained inside the pf object passed 
        as parameter. an additional search string can be used to filter the 
        returned objects
        '''
        try:
            return pf_object.GetContents(search_string, recursive)
        except Exception as e:
            self.print("Script data type error (get content): " + str(e))
            self.print(traceback.format_exc())
            return None
        
    
    def get_referenced_object_of(self, reference):
        '''
        function returning the object referenced by the given reference
        '''
        try:
            return reference.GetAttribute('obj_id')
        except Exception as e:
            self.print("No reference available " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_connection_number(self, pf_object):
        '''
        function returning the number of connection of the given Pf element
        '''
        #import pydevd
        #pydevd.settrace(stdoutToServer=True, stderrToServer=True)
        # try to get the connection attribute name list using the given elemnt class name
        attribute_name_list = self.pf_attribute_name_list.\
                                        get(self.get_class_name_of(pf_object))
        if  attribute_name_list == None:   # if no list has been found
            attribute_name_list = self.default_pf_attribute_name # use the default
        else:           # return the number of items of its list
            return sum(len(name) > 0 for name in attribute_name_list)    
        index = 1
        try:
            for index in range(1,len(attribute_name_list)+1):
                if len(attribute_name_list[index]) == 0:  # if the name is void 
                    break                                 # found the last index
                else:                       # try to get the value
                    pf_object.get_attribute(attribute_name_list[index])
        except:          
            pass
        return index
    

    def get_branch_relays_of(self, branch, branch_side = Side.AllSide):
        '''
        function returning all active relays protecting a line, a transformer or
         a generator at the given side (2nd parameter)
        '''
        try:            
            attribute_name_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            start_side = branch_side if branch_side != self.Side.AllSide else self.Side.Side_1
            end_side  = branch_side+1 if branch_side != self.Side.AllSide else self.Side.Side_3
            if self.get_connection_number(branch) == 1:
                start_side = self.Side.Side_1
                end_side = self.Side.Side_2
            relay_list = []
            relay_type_values = [self.RelayType.OVERCURRENT.value, 
                                 self.RelayType.DISTANCE.value,
                                 self.RelayType.DIRECTIONAL.value]
            for side in range(start_side,end_side):
                pf_attribute_name = attribute_name_list[side] \
                                    if  attribute_name_list != None    \
                                    else self.default_pf_attribute_name[side]
                cubicle = self.get_attribute(branch, pf_attribute_name)
                relay_list = relay_list + cubicle.GetContents('*.ElmRelay')
            return [relay for relay in relay_list if \
                    self.is_out_of_service(relay) == False and \
                    relay.typ_id != None and\
                    relay.c_category in relay_type_values]         
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Attribute name: " + pf_attribute_name)
            self.print("Branch name: " + self.get_name_of(branch))
            self.print(traceback.format_exc())
            return None
        
        
    def get_branch_cubicles_of(self, branch, branch_side = Side.AllSide):
        '''
        function returning the connection cubicles of a line, a transformer or
         a generator at the given side (2nd parameter)
        '''
        try:            
            attribute_name_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            start_side = branch_side if branch_side != self.Side.AllSide else self.Side.Side_1
            end_side  = branch_side+1 if branch_side != self.Side.AllSide else self.Side.Side_3
            if self.get_connection_number(branch) == 1:
                start_side = self.Side.Side_1
                end_side = self.Side.Side_2
            cubicle_list = []
            for side in range(start_side,end_side):
                pf_attribute_name = attribute_name_list[side] \
                                    if  attribute_name_list != None    \
                                    else self.default_pf_attribute_name[side]
                cubicle = self.get_attribute(branch, pf_attribute_name)
                cubicle_list.append(cubicle)
            return cubicle_list         
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        
    
    def get_branch_breakers_of(self, branch, branch_side = Side.AllSide):
        '''
        function returning all breakers protecting a line, a transformer or
         a generator at the given side (2nd parameter)
        '''
        try:            
            attribute_name_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            start_side = branch_side if branch_side != self.Side.AllSide else self.Side.Side_1
            end_side  = branch_side+1 if branch_side != self.Side.AllSide else self.Side.Side_3
            if self.get_connection_number(branch) == 1:
                start_side = self.Side.Side_1
                end_side = self.Side.Side_2
            breaker_list = []
            for side in range(start_side,end_side):
                pf_attribute_name = attribute_name_list[side] \
                                    if  attribute_name_list != None    \
                                    else self.default_pf_attribute_name[side]
                cubicle = self.get_attribute(branch, pf_attribute_name)
                breaker_list = breaker_list + cubicle.GetContents('*.ElmCoup, *.StaSwitch')
            return breaker_list         
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_branch_busses_of(self, branch, branch_side):
        '''
        function returning all busbars connected to a line, a transformer or a 
        generator
        '''
        try:            
            attribute_name_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            pf_attribute_name = attribute_name_list[branch_side] \
                                if  attribute_name_list != None else \
                                    self.default_pf_attribute_name[branch_side]
            if pf_attribute_name:
                cubicle = self.get_attribute(branch, pf_attribute_name)
                return cubicle.GetAttribute('fold_id')
            else:
                return None
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_cubicle_busbar_of(self, cubicle):
        '''
        function returning the busbar at which the givne cubicle belongs
        '''
        try:
            if cubicle:
                return cubicle.__getattr__("cterm")
            else:
                return None        
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None    
            


    def get_busbar_substation_of(self, busbar):
        '''
        function returning the pf substation object at which the given busbar 
        belong
        '''
        try:
            return busbar.__getattr__("cpSubstat")
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return 0

    def get_busbar_shc_I_in(self, busbar):
        '''
        function returning the value of the short circuit current at the given 
        busbar
        '''
        try:
            return busbar.__getattr__("m:Ikss")
        except Exception as e:
            self.print("Script data type error: " + str(e) +
                        " (" + self.get_name_of(busbar) +  ")")
            self.print(traceback.format_exc())
            return 0
        
    
    def get_busbar_rated_voltage_in(self, busbar):
        '''
        function returning the value of the rated voltage  at the given 
        busbar
        '''
        try:
            return round(busbar.__getattr__("uknom"), 2)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return 0    
    
    

    def get_shc_I_bus1_in(self, branch):
        '''
        function returning the value of the short circuit current at the given
         connection (side bus1)
        '''
        try:       
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.slot_1 if slot_list != None else \
                            self.default_shc_slot_name[self.Side.Side_1]
            return branch.__getattr__(slot_name)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Item: " + self.get_full_name_of(branch))
            self.print(traceback.format_exc())
            return 0


    def get_shc_I_list_bus1_in(self, branch):
        '''
        function returning the value of the short circuit currents at the given
         connection (side bus1)
        '''
        try:       
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            i_list = []
            for phase in [":A", ":B", ":C"]:
                slot_name = slot_list.slot_1 if slot_list != None else \
                                self.default_shc_slot_name[self.Side.Side_1]
                i_list.append(branch.__getattr__(slot_name + phase))
            return i_list
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Item: " + self.get_full_name_of(branch))
            self.print(traceback.format_exc())
            return [0, 0, 0]


    def get_shc_I_bus2_in(self, branch):
        '''
        function returning the value of the short circuit current at the given 
        connection (side bus2)
        '''
        try:
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.slot_2 if slot_list != None else \
                            self.default_shc_slot_name[self.Side.Side_2]
            return branch.__getattr__(slot_name)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return 0
     
    
    def get_shc_I_list_bus2_in(self, branch):
        '''
        function returning the value of the short circuit currents at the given
         connection (side bus2)
        '''
        try:       
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            i_list = []
            for phase in [":A", ":B", ":C"]:
                slot_name = slot_list.slot_2 if slot_list != None else \
                                self.default_shc_slot_name[self.Side.Side_2]
                i_list.append(branch.__getattr__(slot_name + phase))
            return i_list
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Item: " + self.get_full_name_of(branch))
            self.print(traceback.format_exc())
            return 0 
        
        
    def get_shc_I_bus3_in(self, branch):
        '''
        function returning the value of the short circuit current at the given
         connection (side bus3)
        '''
        try:
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.slot_3 if slot_list != None else \
                            self.default_shc_slot_name[self.Side.Side_3]
            return branch.__getattr__(slot_name)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return 0
    
    
    def get_shc_I_list_bus3_in(self, branch):
        '''
        function returning the value of the short circuit currents at the given
         connection (side bus3)
        '''
        try:       
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            i_list = []
            for phase in [":A", ":B", ":C"]:
                slot_name = slot_list.slot_3 if slot_list != None else \
                                self.default_shc_slot_name[self.Side.Side_3]
                i_list.append(branch.__getattr__(slot_name + phase))
            return i_list
        except Exception as e:
#             self.print("Script data type error: " + str(e))
#             self.print("Item: " + self.get_full_name_of(branch))
#             self.print(traceback.format_exc())
            i_list = [0, 0, 0]
            return i_list
    
    
    def get_shc_I_bus4_in(self, branch):
        '''
        function returning the value of the short circuit current at the given
         connection (side bus4)
        '''
        try:
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.slot_4 if slot_list != None else \
                            self.default_shc_slot_name[self.Side.Side_4]
            return branch.__getattr__(slot_name)
        except Exception as e:
#             self.print("Script data type error: " + str(e))
#             self.print(traceback.format_exc())
            i_list = [0, 0, 0]
            return i_list    
    
    
    def get_shc_I_list_bus4_in(self, branch):
        '''
        function returning the value of the short circuit currents at the given
         connection (side bus4)
        '''
        try:       
            slot_list = self.i_slot_name_list.get(self.get_class_name_of(branch))
            i_list = []
            for phase in [":A", ":B", ":C"]:
                slot_name = slot_list.slot_4 if slot_list != None else \
                                self.default_shc_slot_name[self.Side.Side_4]
                i_list.append(branch.__getattr__(slot_name + phase))
            return i_list
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print("Item: " + self.get_full_name_of(branch))
            self.print(traceback.format_exc())
            return 0
        

    def get_connections_of(self, element):
        '''
        function returning all elements connected to the given element
        NOTE: THE PF FUNCTION CONTAINS A BUG.  USE "get_bus_connections_of" FOR BUSSES
        '''
        try:
            return element.GetConnectedElements(0, 0, 0)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_bus_connections_of(self, bus):
        '''
        function returning all elements connected to the given bus
        NOTE: TO USE INSTEAD OF "get_connections_of "FOR BUSSES
        '''
        try:
            return [connection.obj_id for connection \
                    in bus.GetContents("*.StaCubic") if connection.obj_id != None and\
                    all([switch.on_off == 1 for switch in \
                         connection.GetContents("*.StaSwitch")])]
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None

    def get_bus_cubicles(self, bus):
        '''
        function returning a list containing all cubicles of the given busbar
        '''
        try:
            return bus.GetContents("*.StaCubic") 
        except Exception as e:
            self.print("Script data type error (get_bus_cubicle): " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_line_bus_of(self, line):
        '''
        function returning the i index bus connected to the given line
        '''
        try:
            return line.bus1.GetAttribute('fold_id')
        except Exception as e:
            self.print("Script data type error (fold_id): " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_line_cubicle_i_of(self, line):
        '''
        function returning the i index cubicle connected to the given line
        '''
        try:
            return line.bus1
        except Exception as e:
            self.print("Script data type error (line cubicle i): " + str(e))
            self.print(traceback.format_exc())
            return None
    
    
    def get_line_cubicle_j_of(self, line):
        '''
        function returning the j index cubicle connected to the given line
        '''
        try:
            return line.bus2
        except Exception as e:
            self.print("Script data type error (-line cubicle j-): " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_line_voltage_of(self, line):
        '''
        function returning the rated voltage of the line passed as parameter
        Note: no check that the passed parameter is a line
        '''
        try:
            line_type = self.get_attribute(line, 'typ_id')
            if line_type is not None:
                if (self.get_class_name_of(line_type) == "TypLne"):
                    return round(self.get_attribute(line_type, 'uline'), 2) 
                else:
                    return self.get_busbar_rated_voltage_in(\
                                                    self.get_line_bus_of(line))
            else:
                return 0
        except Exception as e:
            self.print("Script data type error (line voltage): " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_line_z(self, line):
        '''
        function the returning the positive sequence impedance of the given line
        '''
        try:
            return line.GetAttribute('Z1')
        except Exception as e:
            self.print("Script data type error (Z1): " + str(e))
            self.print(traceback.format_exc())
            return 0
        
    
    def get_line_angle(self, line):
        '''
        function the returning the line angle in radiant of the given line
        '''
        try:
            return line.GetAttribute('phiz1')
        except Exception as e:
            self.print("Script data type error (phiz1): " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_transformer_z(self, transformer):
        '''
        function the returning the positive sequence impedance of the given
        transformer
        '''
        try:
            s_nom = transformer.GetAttribute('Snom_a')
            v_nom = 1
            vcc = 1
            trafo_type = self.get_attribute(transformer, 'typ_id')
            if trafo_type is not None:
                if self.get_class_name_of(trafo_type) == "TypTr2":
                    v_nom = trafo_type.GetAttribute('utrn_h')
                elif self.get_class_name_of(trafo_type) == "TypTr3":
                    v_nom = trafo_type.GetAttribute('uktr3_h') 
                vcc = trafo_type.GetAttribute('uktr') 
                # vcc real part
                vccr = trafo_type.GetAttribute('uktrr') 
                # calculate vcc immaginary part
                vccx = sqrt(vcc*vcc-vccr*vccr)       
            return complex(vccr, vccx) * v_nom * v_nom/s_nom
        except Exception as e:
            self.print("Script data type error (Z1 trafo ): " + str(e))
            self.print(traceback.format_exc())
            return 0 
        
    def get_transformer_rated_i(self, transformer):
        '''
        function the returning the rated current of the given transformer
        '''
        from math import sqrt
        try:
            v_nom = 1
            trafo_type = self.get_attribute(transformer, 'typ_id')
            if trafo_type is not None:
                if self.get_class_name_of(trafo_type) == "TypTr2":
                    v_nom = trafo_type.GetAttribute('utrn_h')
                    s_nom = transformer.GetAttribute('Snom_a')
                elif self.get_class_name_of(trafo_type) == "TypTr3":
                    v_nom = trafo_type.GetAttribute('uktr3_h')  
                    s_nom = trafo_type.GetAttribute('strn3_h')         
            return  s_nom / (sqrt(3) * v_nom)
        except Exception as e:
            self.print("Script data type error (Z1 trafo ): " + str(e))
            self.print("Trafo: " + self.get_name_of(transformer))
            self.print(traceback.format_exc())
            return None  
               
    
    def get_shunt_rated_i(self, shunt):
        '''
        function the returning the rated current of the given shunt
        '''
        try:
            total_i =  self.get_attribute(shunt, 'cutot')
            return self.get_attribute(shunt, 'cucap') if total_i == 0 else \
                   total_i          
        except Exception as e:
            self.print("Script data type error (In shunt ): " + str(e))
            self.print("Shunt: " + self.get_name_of(shunt))
            self.print(traceback.format_exc())
            return None


    def get_branch_bus1_of(self, branch):
        '''
        function returning the i index bus connected to the given branch
        '''
        try:
            slot_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.attr_1 if slot_list != None else \
            self.default_pf_attribute_name[self.Side.Side_1]
            return branch.__getattr__(slot_name).GetAttribute('fold_id')       
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_branch_bus2_of(self, branch):
        '''
        function returning the j index bus connected to the given branch
        '''
        try:
            slot_list = self.pf_attribute_name_list.get(self.get_class_name_of(branch))
            slot_name = slot_list.attr_2 if slot_list != None else \
            self.default_pf_attribute_name[self.Side.Side_2]
            return branch.__getattr__(slot_name).GetAttribute('fold_id')
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_line_grid_of(self, line):
        '''
        function returnin the grid object at which the given line belongs
        '''
        try:
            return line.GetAttribute('fold_id')
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None


    def get_grid_lines_of(self, grid):
        '''
        function returning all the lines beloging to the given grid 
        '''
        try:
            return grid.GetContents("*.ElmLne")
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return None
        
        
    def get_shc_trace_tripped_devices(self):
        '''
        function getting from the SHC trace calculation currently set in the 
        active study case the list of the tripped devices
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.GetTrippedDevices()
        except Exception as e:
            self.print("SHC Trace Tripped Devices failed!: " + str(e))
            self.print(traceback.format_exc())
            
            
    def get_shc_trace_started_devices(self):
        '''
        function getting from the SHC trace calculation currently set in the 
        active study case the list of the started devices
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.GetStartedDevices()
        except Exception as e:
            self.print("SHC Trace Tripped Devices failed!: " + str(e))
            self.print(traceback.format_exc())
            
            
    def get_shc_trace_device_trip_time_of(self, pf_element):
        '''
        function getting from the SHC trace calculation currently set in the 
        active study case the trip time of the given device
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return round(shctrace_object.GetDeviceTime(pf_element), 3)
        except Exception as e:
            self.print("SHC Trace Device Trip Time failed!: " + str(e))
            self.print(traceback.format_exc())
            
        
    def get_shc_trace_actual_time_step(self):
        '''
        function returnig the current time of the shc_trace trace static simulation
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return round(shctrace_object.GetCurrentTimeStep(), 3)
        except Exception as e:
            self.print("SHC Trace Current Time Step failed!: " + str(e))
            self.print(traceback.format_exc())
     
     
    def get_shc_sweep_object(self):
        '''
        function returning the shc object used for the time distance diagram short
        circuit sweep
        '''
        # get directly from the test case the shc object
        try:
            SC_sweep = self.app.GetFromStudyCase("ComShcsweep")
            SC_sweep_objects = SC_sweep.GetContents()
            # if no shc object is present create it!
            if len(SC_sweep_objects) == 0:
                SC_sweep.CreateObject("ComShc")
                SC_sweep_objects = SC_sweep.GetContents()      
            SC_object = SC_sweep_objects[0]
            return SC_object
        except Exception as e:
            self.print("Get SHC sweep shc object failed!: " + str(e))
            self.print(traceback.format_exc())
      
      
    def get_diagram_pages(self, diagram_name = ''):
        '''
        function creating a time distance diagram of the given name 
        '''
        try:           
            graphic_board = self.app.GetGraphicsBoard()
            # get VI pages
            pages_list = graphic_board.GetContents(diagram_name + '*.SetVipage')
            if len(pages_list) == 0:
                pages_list = graphic_board.GetContents(diagram_name + '*.GrpPage')
            return pages_list
        except Exception as e:
            self.print("Get diagram pages failed!: " + str(e))
            self.print(traceback.format_exc())   
            
    def get_TCDdiagram_pages(self, diagram_name = ''):
        '''
        function getting time distance diagram(s) of the given name 
        '''
        try:           
            graphic_board = self.app.GetGraphicsBoard()
            # get VI pages
            pages_list = graphic_board.GetContents(diagram_name + '*.SetVipage')
            if len(pages_list) == 0:
                pages_list = graphic_board.GetContents(diagram_name + '*.GrpPage')   
            pages_list = [page for page in pages_list \
                          if len(page.GetContents('*.VisOcplot'))!=0]                           
            return pages_list
        except Exception as e:
            self.print("Get diagram pages failed!: " + str(e))
            self.print(traceback.format_exc())    
                
    
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


    def is_out_of_service(self, pf_element):
        '''
        function checking if the PowerFactory element passed as parameter is 
        active
        '''
        try:
            return (True if pf_element.IsOutOfService() == 1 else False)
        except Exception as e:          
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return True
    
    
    def is_open(self, pf_element):   
        '''
        function checking if the PowerFactory element passed as parameter is 
        open. 
        It returns always False for any element except ElmCoup and StaSwitch 
        elements for which we get the 'on_off' parameter value 
        '''
        try:
            element_class_name = self.get_class_name_of(pf_element)
            if element_class_name == "ElmCoup" or element_class_name == "StaSwitch":
                return True if pf_element.GetAttribute('on_off') == 0 else False
            else:
                return False
        except Exception as e:          
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return True
        
        
    def is_energized(self, pf_element):
        '''
        function checking if the PowerFactory element passed as parameter is 
        energized
        '''
        try:
            return (True if pf_element.IsEnergized() == 1 else False)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return True
       
        
    def is_generator(self, pf_element):
        '''
        function checking if the given element is a synchronous or asynchronous
         generator, or a network 
        '''
        class_name = self.get_class_name_of(pf_element)
        if ((class_name == "ElmAsm" or class_name == "ElmSym") 
            and self.get_attribute(pf_element, "i_mot") == 0): # i_mot=0 means acting as generator
            return True
        elif class_name == "ElmXnet":
            return True
        elif class_name == "ElmGenstat":
            return True
        return False
    
    
    def is_transformer(self, pf_element):
        '''
        function checking if the given element is a transformer 
        (2,3,4 windings, ElmTrb booster transformer)
        '''
        class_name = self.get_class_name_of(pf_element)
        if "ElmTr" in class_name: 
            return True
        return False
    
    
    def is_line(self, pf_element):
        '''
        function checking if the given element is a line
        '''
        return True if self.get_class_name_of(pf_element) == "ElmLne" else False
    
    
    def is_busbar(self, pf_element):
        '''
        function checking if the given element is a bus bar
        '''
        return True if self.get_class_name_of(pf_element) == "ElmTerm" else False
    
    
    def is_main_protection(self, pf_relay):
        '''
        function returning true if the application variable inside ElmRealy is 0 
        '''
        return True if self.get_attribute(pf_relay, "application") == 0 else False
    
    
    def is_subrelay(self, relay):
        '''
        function checking if the given relay is a subrelay
        '''
        return True if self.get_class_name_of(relay.GetAttribute('fold_id'))\
                                             == "ElmRelay" else False
                                             

    def is_shc_trace_next_step_available(self):
        '''
        function checking if another step of the SHC trace calculation currently  
        set in the active study  case is still availble
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.NextStepAvailable()
        except Exception as e:
            self.print("SHC Trace is Next Step Available  calculation failed!: "\
                                                             + str(e))
            self.print(traceback.format_exc())
            
    def is_shc_valid(self):
        '''
        function returning true if a shc_trace has been run and valid results have
        been obtained
        '''
        return False if self.app.IsShcValid() == 0 else True
    
    def is_ldf_valid(self):
        '''
        function returning true if  a ldf has been run and valid results have
        been obtained
        '''
        return False if self.app.IsLdfValid() == 0 else True
    
    def is_overcurrent_relay(self, relay):
        '''
        function returning true if the given relay is an overcurrent relay
        '''
        try:          
            return True if (relay.typ_id and relay.c_category in \
                            [self.RelayType.OVERCURRENT.value,
                             self.RelayType.DIRECTIONAL.value]) else False
        except Exception as e:
            self.print("Is overcurrent relay failed!: " + str(e))
            self.print(traceback.format_exc())
            
    def is_distance_relay(self, relay):
        '''
        function returning true if the given relay is an distance relay
        '''
        try:          
            return True if (relay.typ_id and relay.c_category in \
                            [self.RelayType.DISTANCE.value]) else False
        except Exception as e:
            self.print("Is distance relay failed!: " + str(e))
            self.print(traceback.format_exc())

#=========================================================================
# Set Methods
#=========================================================================

    def set_output_file(self, output_file):
        '''
        Function setting the output file to redirect the print instruction
        '''
        self.output_file = output_file
        
    
    def set_attribute(self, element, attribute_name, attribute_value):
        '''
        Function setting equal to the given attribute_value the value of the 
        attribute "attribute_name" beloging to the element passed as first
         parameter
        '''
        try:
            element.SetAttribute(attribute_name, attribute_value)
        except Exception as e:
            self.print("Script data type error: " + str(e))
            self.print(traceback.format_exc())
            return ""
    
    
    def set_ldf_basic_configuration(self):
        '''
        Function setting the LDF object, the calculation method is set as 
        asymmetric to avoid problems with asymmetric faults 
        '''
        ldf_object = self.app.GetFromStudyCase('ComLdf')
        try:
            ldf_object.opt_net = 1  # asymmetric
        except Exception as e:
            self.print("Setting LDF basic values failed!: " + str(e))
            self.print(traceback.format_exc())


    def set_fault(self, faultype, resistance = 0, single_shc = False, shc_object = None):
        '''
        Function setting the actual SHC object fault type and resistance
        '''
        if shc_object == None:
            shc_object = self.app.GetFromStudyCase('ComShc')\
                        if self.shc_trace == None or single_shc == True else self.shc_trace
        try:
            if self.shc_trace == None or single_shc == True:
                shc_object.iopt_shc = faultype.value
                shc_object.Rf = resistance
            else: # shc_trace trace settings
                shc_object.i_shc = self.FautType_Shctrace[faultype.value]
                shc_object.R_f = resistance
        except Exception as e:
            self.print("Setting fault type/resistance failed!: " + str(e))
            self.print(traceback.format_exc())


    def set_fault_position(self, element, position=0, \
                           single_shc = False, reference_busbar =  None):
        '''
        Function setting the actual SHC object fault position along the 
        "element" (a line or a busbar) given as parameter. 
        If position is smaller than 0 no position is set.
        If the reference_busbar is given the position value is adapted to use 
        such bus bar as reference
        '''
        shc_object = self.app.GetFromStudyCase('ComShc')\
                        if self.shc_trace == None or single_shc == True else self.shc_trace
        try:
            if self.shc_trace == None or single_shc == True:
                shc_object.shcobj = element
            else:
                shc_object.p_target = element
            if position >= 0 and position <= 100:
                if self.shc_trace == None or single_shc == True:
                    # only for the lines adapt the position if it's required
                    if self.get_class_name_of(element) == 'ElmLne':
                        # get from the line the bus at i
                        bus1 = self.get_line_bus_of(element)
                        if reference_busbar and reference_busbar != bus1:
                            position = 100 - position
                    shc_object.ppro = position    
                    # set the reference for the fault (distance from i)
                    shc_object.iopt_dfr = 0
                else:   # shc_trace trace settings
                    shc_object.p_target.fshcloc = position 
            else:
                try:
                    raise ValueError('Fault position wrong value! ')
                except Exception as e:
                    self.print(
                        'Fault position wrong value! Value = {} '.format(position))
                    self.print(traceback.format_exc())
        except Exception as e:
            self.print("Setting fault position failed!: " + str(e))
            self.print(traceback.format_exc())
        return position


    def set_shc_basic_configuration(self, calc_method):
        '''
        Function setting the SHC object calculation method, "Calculate" = max 
        SHC current and the "fault distance from..." parameter
        '''
        shc_object = self.app.GetFromStudyCase('ComShc')\
                                                if self.shc_trace == None else self.shc_trace
        try:
            shc_object.iopt_mde = int(calc_method)
            shc_object.iopt_cur = 0   # calculate max SHC I
            shc_object.iopt_dfr = 0   # fault distance from terminal i
            shc_object.iopt_allbus = 0# set fault location at "user selection"
        except Exception as e:
            self.print("Setting SHC basic parameters failed!: " + str(e))
            self.print(traceback.format_exc())
            

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
                return 1;
    
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
            return 1;   
     
    
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
            return 1;     
    
    def create_variation(self, variation_name):
        '''
        function creating the "variation" object (IntScheme) and the relevant
        Expansion Stage 
        '''
        try:
            active_project = self.app.GetActiveProject()
            variation_folder = active_project.GetContents('Variations.IntPrjfolder', 1)[0]
            variation_object = variation_folder.CreateObject('IntScheme', variation_name)
            date_time = datetime.utcnow()
            date_time_integer = date_time.timestamp() 
            variation_object.NewStage("Expansion Stage", date_time_integer, 1)
            return variation_object
        except Exception as e:
            self.print("Variation creation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
        
        
    def activate_variation(self, variation_object):
        '''
        function activating the given "variation" object (IntScheme)
        '''
        try:
            return variation_object.Activate()
        except Exception as e:
            self.print("Variation activation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
      
        
    def deactivate_variation(self, variation_object):
        '''
        function deactivating the given "variation" object (IntScheme)
        '''
        try:
            return variation_object.Deactivate()
        except Exception as e:
            self.print("Variation deactivation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
        
        
    def activate_study_case(self, study_case_object):
        '''
        function activating the given study case object (IntCase)
        '''
        try:
            return study_case_object.Activate()
        except Exception as e:
            self.print("Study case activation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
      
        
    def deactivate_study_case(self, study_case_object):
        '''
        function deactivating the given study case object (IntScheme)
        '''
        try:
            return study_case_object.Deactivate()
        except Exception as e:
            self.print("Study case deactivation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
        

    def run_ldf(self):
        '''
        function running the LDF calculation currently set in the active study 
        case
        '''
        try:
            ldf_object = self.app.GetFromStudyCase('ComLdf')
            return ldf_object.Execute()
        except Exception as e:
            self.print("LDF calculation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1;
        

    def run_shc(self, single_shc = False):
        '''
        function running the SHC calculation currently set in the active study 
        case
        '''
        try:
            shc_object = self.app.GetFromStudyCase('ComShc')\
                        if self.shc_trace == None or single_shc == True else self.shc_trace
            return shc_object.Execute()
        except Exception as e:
            self.print("SHC calculation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1; 
            
    def use_shc_trace_shc(self):
        '''
        function forcing the interface to use the shc_trace trace short circuit object
        as shc_trace object 
        '''
        try:
            shc_container = self.app.GetFromStudyCase('IntEvtshc')
            shc_object_list = shc_container.GetContents("*.EvtShc")
            # use 
            if len(shc_object_list) > 0:
                self.shc_trace = shc_object_list[0]
            else:       # otherwise create a short circuit
                self.shc_trace = shc_container.CreateObject('EvtShc', 'Generic shc_trace')
        except Exception as e:
            self.print("SHC Trace set SHC failed!: " + str(e))
            self.print(traceback.format_exc())
            
    
    def run_initialize_shc_trace(self):
        '''
        function running the initial step of the SHC trace calculation currently 
        set in the active study  case
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.ExecuteInitialStep()
        except Exception as e:
            self.print("SHC Trace calculation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
             
            
    def run_whole_shc_trace(self):
        '''
        function running the whole SHC trace calculation currently 
        set in the active study  case
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            result = shctrace_object.ExecuteInitialStep()
            if result > 0:
                raise Exception('Short circuit trace initialization failed!')
            return shctrace_object.ExecuteAllSteps()
        except Exception as e:
            self.print("SHC Trace calculation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1


    def run_shc_trace_next_step(self):
        '''
        function running the following step of the SHC trace calculation 
        currently set in the active study  case
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.ExecuteNextStep()
        except Exception as e:
            self.print("SHC Trace Next Step  calculation failed!: " + str(e))
            self.print(traceback.format_exc())
            return 1
    
    
    def interface_get_shc_trace_trip_time_of(self, relay):
        '''
        function returning the tripping time of the given relay after that the
        the short circuit trace has been performed
        '''
        try:
            shctrace_object = self.app.GetFromStudyCase('ComShctrace')
            return shctrace_object.GetDeviceTime(relay)
        except Exception as e:
            self.print("SHC Trace GetDeviceTime  calculation failed!: " + str(e))
            self.print(traceback.format_exc())    


    def print(self, string_to_print):
        '''
        function sending a string to the PowerFactory outputwindow and to a file
        if the self.output_file variable has been set
        '''
        self.app.PrintPlain(string_to_print)
        if self.output_file != None:
            self.output_file.write(string_to_print + '\n')

    
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


    def enable(self, element):
        '''
        function enabling the given element
        '''
        try:
            self.set_attribute(element, "outserv", 0)
        except Exception as e:
            self.print("Enable command error: " + str(e))
            self.print(traceback.format_exc())


    def disable(self, element):
        '''
        function disabling the given element
        '''
        try:
            self.set_attribute(element, "outserv", 1)
        except Exception as e:
            self.print("Disable command error: " + str(e))
            self.print(traceback.format_exc())


    def close_(self, breaker):
        '''
        functionc closing the given breaker
        '''
        try:
            self.set_attribute(breaker, "on_off", 1)
        except Exception as e:
            self.print("Switch close error: " + str(e))
            self.print(traceback.format_exc()) 
            
            
    def open_(self, breaker):
        '''
        functionc opening the given breaker
        '''
        try:
            self.set_attribute(breaker, "on_off", 0)
        except Exception as e:
            self.print("Switch open error: " + str(e))
            self.print(traceback.format_exc())


    def switch_on(self, element):
        '''
        function switching on (energizing) the given element
        '''
        try:
            return element.SwitchOn(1)
        except Exception as e:
            self.print("Switch on error: " + str(e))
            self.print(traceback.format_exc())


    def switch_off(self, element):
        '''
        function switching off (de energizing) the given element
        '''
        try:
            return element.SwitchOff(1)
        except Exception as e:
            self.print("Switch off error: " + str(e))
            self.print(traceback.format_exc())
            
            
    def clear_output_window(self):
        '''
        function clearing the PowerFactory output window
        '''
        try:
            self.app.ClearOutputWindow() 
        except Exception as e:
            self.print("Clear output window error: " + str(e))
            self.print(traceback.format_exc())
            

    def run_verification(self):
        '''
        Function run the generic PF verification over all relays
        NOTE: at the moment not used, custom verification implemented in 
        tracer_logic
        '''
        try:
            raise Exception('Verification failed!')
        except Exception as e:
            self.print("Run verification failed!: " + str(e))
            self.print(traceback.format_exc())

            
    def create_relay(self, relay_type, cubicle, relay_name):
        '''
        function creating the relay of the given type and name in the given cubicle
        '''
        try:
            new_relay = cubicle.CreateObject("ElmRelay", relay_name) 
            if new_relay != None:
                self.set_attribute(new_relay, "typ_id", relay_type)
                new_relay.SlotUpdate()
            return new_relay
        except Exception as e:
            self.print("Create relay error: " + str(e))
            self.print(traceback.format_exc())
            
            
    def delete_relay(self, pf_relay):
        '''
        function deleting from the PowerFactory database the given relay
        '''
        try:
            pf_relay.Delete()
        except Exception as e:
            self.print("Delete relay error: " + str(e))
            self.print(traceback.format_exc())
            
        
    def create_time_distance_diagram(self, diagram_name, path = None, \
                                     relays_in_path = None, \
                                     relay_type_line_format = None):
        '''
        function creating a time distance diagram of the given name with the given
        path (if provided)
        '''
        try:           
            graphic_board = self.app.GetGraphicsBoard()
            # create VI page
            tcd_page = graphic_board.GetPage(diagram_name, 1)
            # create TD plot
            new_time_distance_diagram =  tcd_page.GetVI("TD plot", "VisPlottz", 1)
            if path != None:
                new_time_distance_diagram.pPath = path
                # to improve the graphical layout set only 3 columns for the names
                new_time_distance_diagram.leg_numcol = 3
                # to reduce the text length don't show the terminal name
                new_time_distance_diagram.leg_bus = 0
            if relays_in_path != None:
                # the relays in path are a list of tuples, unpack the tuples and create
                # a "flat" list
                relays_in_path_list = []
                for relay_tuple in relays_in_path:
                    relays_in_path_list += list(relay_tuple)
                self.set_attribute(new_time_distance_diagram, \
                                         "fObjs", relays_in_path_list)  
                # set the colors using the same color for the reays in a tuple
                # all colors from red except green (==3)
                color_set = [color for color in range(2,len(relays_in_path) + 3) \
                             if color != 3]
                fcolors = []
                for color_index, color in enumerate(color_set):
                    for relay_index in range(0, len(relays_in_path[color_index])):
                        fcolors.append(color) 
                
                self.set_attribute(new_time_distance_diagram, \
                                         "fColor", fcolors) 
                # set the line style by default equal to a continous line
                fstyles = list(repeat(1.0, len(relays_in_path_list)))
                # if a relay is an overcurrent relay set the style as not continous line
                for index, relay in enumerate(relays_in_path_list):
                    if self.is_overcurrent_relay(relay) == True:
                        fstyles[index] = 9.0                    
                    # get the line style from the dictionary which links relay type
                    # and line style
                    try:
                        fstyles[index] = relay_type_line_format[self.\
                                        get_relay_model_name_of(relay)].style
                    except Exception as e:
                        pass     
                self.set_attribute(new_time_distance_diagram, \
                                         "fStyle", fstyles) 
                # set the line width
                fwidths = list(repeat(50.0, len(relays_in_path_list)))
                self.set_attribute(new_time_distance_diagram, \
                                         "fWidth", fwidths) 
            return new_time_distance_diagram
            
        except Exception as e:
            self.print("Create Time distance Diagram error: " + str(e))
            self.print(traceback.format_exc())
    
    def create_voltage_time_diagram(self, diagram_name):
        '''
        function creating a voltage time diagram of the given name
        '''
         #Get current graphic board
        graphic_board = self.app.GetGraphicsBoard()
        #Create VI Page
        plot_page = graphic_board.GetPage(diagram_name,1)
        #Create a new subplot
        plot = plot_page.GetVI('VT plot','VisXyplot',1)    
        return plot
          
            
    def create_time_current_diagram(self, diagram_name, items):
        '''
        function creating a time current diagram of the given name with the given
        items displayed. items is a list of tuples
        '''
        try:           
            graphic_board = self.app.GetGraphicsBoard()
            # create VI page
            tcd_page = graphic_board.GetPage(diagram_name, 1)
            # create TD plot
            toc_diagram =  tcd_page.GetVI("TC plot", "VisOcplot", 1)
            items_list = []
            for items_tuple in items:
                    items_list += list(items_tuple)
        
            # set the trafo and the relays to display
            rsl = self.set_attribute(toc_diagram, "gObjs", items_list)             
            # set the color
            # all colors from red except green (==3)
            color_set = [color for color in range(2,len(items) + 3) if color != 3]
            fcolors = []
            for color_index, color in enumerate(color_set):
                for relay_index in range(0, len(items[color_index])):
                    fcolors.append(color)                 
            self.set_attribute(toc_diagram, "gColor", fcolors)             
            # set items style
            gStyles = list(repeat(1.0, len(items_list)))
            rsl = self.set_attribute(toc_diagram, "gStyle", gStyles)
            gWidths = list(repeat(50.0, len(items_list)))
            # set items width
            rsl = self.set_attribute(toc_diagram, "gWidth", gWidths)
            
            self.set_attribute(toc_diagram, "x_min", 100) 
            self.set_attribute(toc_diagram, "x_max", 30000) 
            self.set_attribute(toc_diagram, "y_min", 0.01) 
            self.set_attribute(toc_diagram, "y_max", 1000) 
            
            return toc_diagram
            
        except Exception as e:
            self.print("Create Time current Diagram error: " + str(e))
            self.print(traceback.format_exc()) 
       
            
    def set_time_current_diagram_type(self, toc_diagram, relay_type_string):
        '''
        set the given time current diagram as "phase", Ground", "Negative sequence" ...
        '''
        try:
            option_object = toc_diagram.GetContents("*.SetOcplt")
            self.set_attribute(option_object[0], "ishow", \
                           self.Time_Current_Diagram_Displayed_Relays[relay_type_string])
        except Exception as e:
            self.print("Set TOC diagram type error: " + str(e))
            self.print(traceback.format_exc())  
            
            
    def create_TOC_diagram_verticalline(self, line_name, TOC_diagram, value = 100,\
                        line_type = Time_Current_Diagram_Line_Type["Undefined"]): 
        '''
        function creating a vertical line in the given TOC_diagram at the given
        value. Before running it the active study case must be deactivated
        '''
        try:
            new_line = TOC_diagram.CreateObject("VisXvalue", line_name) 
            if new_line != None:
                self.set_attribute(new_line, "value", value)
                self.set_attribute(new_line, "xis", line_type)
                self.set_attribute(new_line, "style", 2)
                # set the label as user defined
                self.set_attribute(new_line, "label", 1)
                self.set_attribute(new_line, "iopt_lab", 1)
                line_name_list = []
                line_name_list.append(line_name)
                self.set_attribute(new_line, "lab_text", line_name_list)               
            return new_line
        except Exception as e:
            self.print("Create TOC diagram vertical line error: " + str(e))
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
    
    
    def create_study_case_from_last_enabled_study_case(self):
        '''
        create a new test case 
        '''    
        study_case_folder = self.app.GetProjectFolder("study")
        initial_study_cases = study_case_folder.GetContents()
        if self.last_active_study_case != None:
            last_active_study_case_name = self.get_name_of(self.last_active_study_case)
            #new_study_case_name = last_active_study_case_name
            
            study_case_folder.PasteCopy(self.last_active_study_case)
            final_study_cases = study_case_folder.GetContents()
            new_study_case = [testcase for testcase in final_study_cases if \
                            testcase not in initial_study_cases][0]
            self.last_active_study_case = new_study_case
            self.last_active_study_case.Activate()
            return self.last_active_study_case
            
            
    def disable_current_study_case(self):
        '''
        deactivate the current study case and put it in the self.last_active_study_case
        variable
        '''
        try:
            #deactivate the study case
            actual_study_case = self.app.GetActiveStudyCase()           
            if actual_study_case != None:
                self.last_active_study_case = actual_study_case
                self.last_active_study_case.Deactivate()
        except Exception as e:
            self.print("Disable current study case error: " + str(e))
            self.print(traceback.format_exc())  
     
            
    def enable_last_enabled_study_case(self):
        '''
        enable the last study case which has been enabled and which has been stored in
         the
        '''
        try:         
            if self.last_active_study_case != None:
                self.last_active_study_case.Activate()
        except Exception as e:
            self.print("Enable latest active study case error: " + str(e))
            self.print(traceback.format_exc())  
    
    
    def delete_TCC_pages(self, study_case):
        '''
        delete all graphical objects starting with 'TCC-'
        '''
        graphic_board = self.app.GetGraphicsBoard()
        tcc_diagrams = graphic_board.GetContents("TCC-*")
        for tcc_diagram in tcc_diagrams:
            tcc_diagram.Delete()
        self.last_active_study_case.Deactivate()
        self.last_active_study_case.Activate()
                
    
    def create_path(self, path_name, path_elements):
        '''
        function creating a path with the given name and conatining the given elements
        '''
        try:     
            path_folder = self.app.GetDataFolder("IntPath") 
            path_object = path_folder.CreateObject("SetPath", path_name)
            # add the references 
            #path_object.AddRef(path_elements)
            for element in path_elements:
                #path_object.AddRef(element)
                self.add_reference_to(element, containing_path = path_object)               
            return path_object
        except Exception as e:
            self.print("Create Path error: " + str(e))
            self.print(traceback.format_exc())
            
    
    def add_reference_to(self, ref_object, containing_path):
        '''
        function creating a reference to the given object inside the given
        conatining_path
        '''
        try:     
            reference_object = containing_path.CreateObject("IntRef")
            self.set_attribute(reference_object, "obj_id", ref_object)     
            return reference_object
        except Exception as e:
            self.print("Create Reference error: " + str(e))
            self.print(traceback.format_exc())        
    
            
    def set_path_method(self, path, method):
        '''
        function stting the given path as kilometric or short circuit sweep
        '''
        try:     
            path.iopt_mod = self.Time_Distance_Diagram_Method[method]
        except Exception as e:
            self.print("Set Path method error: " + str(e))
            self.print(traceback.format_exc())
            
     
    def refresh_TD_diagram(self, input_diagram):   
        '''
        refresh the givne time distance diagram running a short circuit sweep
        at the moment implemented only  calling the Rebuild function which works
        only in PF V17
        '''
        import re
        # create TD plot
        try:
            diagram_list= []
            if input_diagram == None:
                # get the graphic board
                graphic_board = self.app.GetGraphicsBoard()
                # get the list of pages
                diagram_list = graphic_board.GetContents()
            else:
                diagram_list.append(input_diagram)
            
            for diagram in diagram_list:
                # skip other diagrams 
                if 'KM' not in diagram.loc_name and 'SW-' not in diagram.loc_name:
                    continue
                # skip the kilometric diagrams
                if 'SW-' in diagram.loc_name:
                    SC_sweep = self.app.GetFromStudyCase("ComShcsweep")
                    # set the shc sweep step size equal to constant
                    self.set_attribute(SC_sweep, 'iopt_itr', 'con')
                    # get the shc object
                    SC_sweep_objects = SC_sweep.GetContents('*.ComShc')
                    SC_object = SC_sweep_objects[0]
                    # set the calculation method as "Complete"
                    SC_object.iopt_mde = 3  
                     
                    # set the fault type and the resistance
                    faulttype_ids = {'spgf':['1PHR' , '1PH'],
                                    '2psc':['2PHR', '2PH'],
                                    '3psc':['3PH'],
                                    '2pgf':['2PHGR', '2PHG']}   
                    params_set = False      
                    for faulttype, ids in faulttype_ids.items():
                        self.set_attribute(SC_object, "iopt_shc", faulttype)
                        self.set_attribute(SC_object, "Rf", 0)
                        for id_val in ids:
                            if id_val in diagram.loc_name:                     
                                if 'R' in id_val:
                                    r_index = diagram.loc_name.index(id_val) + len(id_val)
                                    R_list = re.findall(r"[-+]?\d*\.\d+|\d+", \
                                                        diagram.loc_name[r_index:])
                                    if R_list:
                                        self.set_attribute(SC_object, "Rf",\
                                                           float(R_list[0]))
                                        self.print("R: " + str(R_list[0]))
                                    params_set = True 
                                    break
                                else:
                                    params_set = True 
                                    break  
                        if params_set:
                            break    
#                     self.print("Fault type: " + SC_object.iopt_shc + "---" +  diagram.loc_name)                    
                self.show_page(diagram)
                self.app.Rebuild(2)
                diagram.DoAutoScaleX()
        except Exception as e:
            self.print("Refresh TD diagram error: " + str(e))
            self.print(traceback.format_exc())
            
            
    def SHC_sweep_update(self, fault_type_str, refresh_all = False):
        '''
        function get the SHW sweep pages updater object and setting it
        to refresh the shc sweep and to refresh just the user defined pages,
        settting the given list of shc sweep pages and setting the given shc type
        '''
        graphic_updater = self.app.GetFromStudyCase("ComProtgraphic") 
        if graphic_updater != None:
#             self.print("DEBUG 2")
            # set refresh shc swep
            graphic_updater.iopt_action = 2 
            # set user defined pages
            graphic_updater.iopt_pages = 1           
            SC_sweep = self.app.GetFromStudyCase("ComShcsweep")
            
            if refresh_all == True:
                # set all pages refresh
                iopt_pages = 0
                # run the TD diagrams refresh
                graphic_updater.Execute()
            # set only custom  pages refresh
            iopt_pages = 1
            
            # set the shc sweep step size equal to constant
            self.set_attribute(SC_sweep, 'iopt_itr', 'con')
            # get the shc object
            SC_sweep_objects = SC_sweep.GetContents()
            SC_object = SC_sweep_objects[0]
            # set the calculation method as "Complete"
            SC_object.iopt_mde = 3       
#             self.print("DEBUG 3")    
            # get the graphic board
            graphic_board = self.app.GetGraphicsBoard()
            # get the list of pages
            page_list = graphic_board.GetContents()
            # collect the TD diagrams for each topology
            page_list1PH0 = [page for page in page_list if "1PH" in page.loc_name]
            page_list1PH25 = [page for page in page_list if "1PH(25" in page.loc_name]
            page_list1PH50 = [page for page in page_list if "1PH(50" in page.loc_name]
            page_list3PH = [page for page in page_list if "3PH" in page.loc_name]
            
#             self.print("DEBUG 4")
            if fault_type_str == 'spgf' and page_list1PH0:
#                 self.print("DEBUG 4.01")
                self.print(page_list1PH0)
                try:
                    graphic_updater.cUpdatePages = page_list1PH0
                except Exception as e:
                    self.print("set list: " + str(e))
                    self.print(traceback.format_exc())
            
                # set SHC
#                 self.print("DEBUG 4.1")
                SC_object.iopt_shc = fault_type_str
#                 self.print("DEBUG 4.2")
                SC_object.Rf = 0
#                 self.print("DEBUG 5")
                # run the TD diagrams refresh
                graphic_updater.Execute()
#                 self.print("DEBUG 6")
            elif fault_type_str == 'spgf25' and page_list1PH25:
                graphic_updater.cUpdatePages = page_list1PH25
                # set SHC
                SC_object.iopt_shc = 'spgf'
                SC_object.Rf = 25
                # run the TD diagrams refresh
                graphic_updater.Execute()
            elif fault_type_str == 'spgf50' and page_list1PH50:
                graphic_updater.cUpdatePages = page_list1PH50
                # set SHC
                SC_object.iopt_shc = 'spgf'
                SC_object.Rf = 50
                # run the TD diagrams refresh
                graphic_updater.Execute()
            elif fault_type_str == '3psc' and page_list3PH:
                graphic_updater.cUpdatePages = page_list3PH
                # set SHC
                SC_object.iopt_shc = fault_type_str
                SC_object.Rf = 0
                # run the TD diagrams refresh
#                 self.print("DEBUG 8") 
                graphic_updater.Execute()
#         self.print("DEBUG 9")    
    

#=========================================================================
# Auxiliary functions
#=========================================================================

    def get_area_bus_of(self, line):
        '''
        function returning for the line passed as parameter the bus which is 
        used as reference for the area  
        '''
        try:
            cubicle = (line.bus1 if line.iZoneBus == 0 else line.bus2)
            bus = cubicle.GetAttribute('fold_id')
            return bus
        except Exception as e:
            self.print("Get area bus error: " + str(e))
            self.print(traceback.format_exc())
            return None
    
    

        
