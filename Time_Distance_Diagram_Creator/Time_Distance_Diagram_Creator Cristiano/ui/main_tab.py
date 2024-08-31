'''
Created on 2 Oct 2018

@author: SMG & AMB
'''

import wx
from ui.main_panel import MainPanel

class MainTab(wx.Panel):
    def __init__(self, notebook, interface, logic, parent):
        wx.Panel.__init__(self, notebook)
        self.SetBackgroundColour( wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))
        self.panel = MainPanel(self, interface, logic)
        self.parent = parent
        
    def refresh_listboxes(self):
        self.panel.refresh_listboxes()