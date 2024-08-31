# Encoding: utf-8

"""
 Title: Time Distance Diagram Creator
 Description: 
 This application automatize the creation of the Time Distance diagram finding
 automatically the paths between the protection devices in a given area
             
 Author: Alberto Borgnino
 Company: CESI
 E-mail: alberto.borgnino@esterni.cesi.it
 Date: 10-October-2019
 Version: 2019v1
 Please contact me with any bugs that you find or any improvements that you'd like to suggest. I mean that!

 Changelog:
 10-October, 2019 :1) Initial File Created
               2) Added basic graphical user interface
 17-October, 2019 :1) 

"""
__author__ = "CESI"
__copyright__ = "Copyright 2019, CESI"
__license__ = "All rights reserved"
__version__ = "0.1"
__email__ = "alberto.borgnino@esterni.cesi.it"
__status__ = "In development"


import os


# module_path = os.path.abspath(os.getcwd())
# module_path = os.path.join(module_path, "tracer")
# print("Actual module path  " + module_path)
# 
# print ("Initial")
# print(sys.path)
# 
# try:
#     user_paths = os.environ['PYTHONPATH'].split(os.pathsep)
# except KeyError:
#     user_paths = []
# 
# if module_path not in sys.path:
#     sys.path.append(user_paths)
# 
# print("Final  ")
# print(sys.path)

import wx


from ui.main_window import MainWindow
from ui.main_window import PROGRAMFOLDER
from ui.splash_screen import SplashScreen


from calc_sw_interface.powerfactory_interface import *
from tracer.tracer_logic import *

PRE_PROD_RELEASE = False

REQUIRED_FILES = ['ui\\about_dlg.py', 'ui\\splash_screen.py', 'ui\\main_tab.py', 'ui\\main_window.py', 'ui\\main_panel.py', 'tracer\\tracer_logic.py']

for reqfile in REQUIRED_FILES: 
    path = os.path.dirname(os.path.realpath(__file__)).replace("tracer","")  # this file is in the "tracer" directory
    if os.path.isfile(os.path.join(path, reqfile)) == False:    
        errorString = "Required file '{}' could not be found.".format(reqfile)
        raise IOError(errorString)




def run(pf_path = "D:\\Materiale Lavoro DIgSILENT\\PF 2019 (April 18)\\build\\Win32\\pf"):
    
    # Create the program folder if it doesn't already exist.
    if not os.path.exists(PROGRAMFOLDER):
        os.makedirs(PROGRAMFOLDER)
        
    ## add the PYTHONPATH for the remote debugging    
    import sys
    try:
        sys.path.index("C:\\Users\\Alberto's laptop\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pydevd.py") # Or os.getcwd() for this directory
    except ValueError:
        sys.path.append("C:\\Users\\Alberto's laptop\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pydevd.py") # Or os.getcwd() for this directory    
    ##
    

    app = wx.App(False)
       
    interface = PowerFactoryInterface()   # interface with the calculation software @UndefinedVariable
    interface.create(username = "", powerfactory_path = pf_path)
    
    logic = PSET(interface)   # protective devices validation logic class @UndefinedVariable
  
    frame = MainWindow(None, \
        "CESI Time Distance Diagram Creator Tool (tracer_logic) 2019 Beta Release v0.01", interface, logic)
   
    
    if PRE_PROD_RELEASE:
        splash_screen = SplashScreen(frame)
        splash_screen.Show()
    else:    
        frame.Show()
    
    app.MainLoop()
    del app
    
    
    
    
    
    
