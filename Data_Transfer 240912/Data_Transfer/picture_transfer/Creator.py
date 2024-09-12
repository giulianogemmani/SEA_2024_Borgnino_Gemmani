# Encoding: utf-8

"""
 Title: Data Transfer
 Description: 
 This application automatize the diagrams picture transfer from PF/Matlab to 
 word. The Excel data transfer is supported as well
             
 Author: Michele Borgnino
 Company: 
 E-mail: borgnino.michele@gmail.com
 Date: Summer 2024
 Version: 0.01 Beta
 Please contact me with any bugs that you find or any improvements that you'd like to suggest. I mean that!

 Changelog:
  

"""
__author__ = "MB&GG"
__copyright__ = "Copyright 2024"
__license__ = "All rights reserved"
__version__ = "0.01"
__email__ = "borgnino.michele@gmail.com"
__status__ = "In development"


import os


import wx


from ui.main_window import MainWindow
from ui.main_window import PROGRAMFOLDER
from ui.splash_screen import SplashScreen


from calc_sw_interface.powerfactory_interface import *
from picture_transfer.picture_transfer_logic import *

PRE_PROD_RELEASE = False

REQUIRED_FILES = ['ui\\about_dlg.py', 'ui\\splash_screen.py', 'ui\\main_tab.py', 'ui\\main_window.py', 'ui\\main_panel.py', 'picture_transfer\\picture_transfer_logic.py']

for reqfile in REQUIRED_FILES: 
    path = os.path.dirname(os.path.realpath(__file__)).replace("picture_transfer","")  # this file is in the "tracer" directory
    if os.path.isfile(os.path.join(path, reqfile)) == False:    
        errorString = "Required file '{}' could not be found.".format(reqfile)
        raise IOError(errorString)




def run(pf_path = "C:\Program Files\Digsilent\PowerFactory"):
    
    # Create the program folder if it doesn't already exist.
    if not os.path.exists(PROGRAMFOLDER):
        os.makedirs(PROGRAMFOLDER)
        
    os.environ["PATH"] = r"C:\Program Files\Digsilent\PowerFactory;"+ os.environ["PATH"]    
        
    ## add the PYTHONPATH for the remote debugging    
    import sys
    try:
        sys.path.index("C:\\Users\\borgn\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pydevd.py") # Or os.getcwd() for this directory
    except ValueError:
        sys.path.append("C:\\Users\\borgn\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pydevd.py") # Or os.getcwd() for this directory    
    ##
    

    app = wx.App(False)
       
    interface = PowerFactoryInterface()   # interface with the calculation software @UndefinedVariable
    interface.create(username = "", powerfactory_path = pf_path)
    
    logic = TransferLogic(interface)   # protective devices validation logic class @UndefinedVariable
  
    frame = MainWindow(None, \
        "Data Transfer 2024 Beta Release v0.01", interface, logic)
    
    if PRE_PROD_RELEASE:
        splash_screen = SplashScreen(frame)
        splash_screen.Show()
    else:    
        frame.Show()
    
    app.MainLoop()
    del app
    
    
    
    
    
