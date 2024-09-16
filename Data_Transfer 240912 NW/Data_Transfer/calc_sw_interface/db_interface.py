'''
Created on 1 April 2020

@author: AMB
'''

from collections import namedtuple

import os , sys

import calc_sw_interface.powerfactory_interface
import calc_sw_interface.powerfactory_relay_interface

from sqlalchemy import create_engine, MetaData
from sqlalchemy.sql import func
from sqlalchemy.orm import Session
from sqlalchemy.ext.automap import automap_base

from configparser import ConfigParser


class DBInterface():
    '''
    parent class collecting all functions which allow to access a databse
    '''
    def __init__(self, interface):
        '''
        Constructor
        '''
        self.interface = interface
        
     
    def load_config_data(self):
        '''
        function loading the database connection data from the config.ini file 
        '''
        self_path = os.path.dirname(os.path.abspath(__file__))
        parent_dir_path = os.path.realpath(os.path.join(self_path, os.pardir))
        config_full_path = os.path.join(parent_dir_path , "config.ini")
        if not os.path.isfile(config_full_path):
            self.interface.print('Config file doesn''t exist at %s' % config_full_path)
            sys.exit(1)
        config = ConfigParser()
        config.read(config_full_path)
        return config
        
        
      
class MYSQLInterface(DBInterface):
    '''
    interface with the MySQL database
    '''
    def __init__(self, interface):
        '''
        Constructor
        '''
        super().__init__(interface)
        self.Base = automap_base()
        # load from the config.ini file the configuration data to access the DB
        config = self.load_config_data()
        self.engine = create_engine(
            "{0:s}://{1:s}:{2:s}@{3:s}:{4:s}/{5:s}".format(
                config.get('database_connection', 'database_type'),
                config.get('database_connection', 'database_username'),
                config.get('database_connection', 'database_password'),
                config.get('database_connection', 'database_host'),
                config.get('database_connection', 'database_port'),
                config.get('database_connection', 'database_schema_name'))
        , echo=False)
        
        self.meta_data = MetaData(bind = self.engine, reflect=True)
#         self.connection = self.engine.connect()
        self.Base.prepare(self.engine, reflect=True)   
        self.schema_dictionary = self.Base.classes._data          
        self.variation = config.get('network', 'variation')
        
        
    def get_table(self, table_name):
        '''
        return the db table of the given name
        '''
        try:
            return self.meta_data.tables[table_name]
        except Exception as e:
            self.interface.print('Not able to get %s DB table' % table_name + str(e))
            return None
        
    
    def execute(self, query):
        '''
        execute the given query
        '''
        try:
            self.engine.execute(query)
        except Exception as e:
            self.interface.print('Failed executing query: ')
            
            
        