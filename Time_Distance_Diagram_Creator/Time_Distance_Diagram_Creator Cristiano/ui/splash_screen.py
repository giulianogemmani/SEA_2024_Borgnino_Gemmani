# This Python file uses the following encoding: utf-8
# Copyright © 2019 CESI spa. 				
# 
# CESI reserves all rights in the Program as delivered.  The Program or any portion thereof 
# may not be reproduced in any form whatsoever except as provided by license without the     
# written consent of CESI.  A license under CESI's rights in the Program may be available  
# directly from CESI.

"""
Pre-production splash screen.

Author: Penn Markham
Created: 6/30/2017

"""

import os
import sys
import wx.html
import webbrowser

 
class SplashScreen(wx.Frame):
 
    def __init__(self, parent):
 
        wx.Frame.__init__(self, parent, wx.ID_ANY, title="CESI Pre-Production Software Disclaimer", size=(600,525), style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)
#         panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        
        html = wxHTML(self)
        file_path = os.path.dirname(os.path.abspath(__file__))
        sys.path.append(file_path)
 
        with open(file_path + '\\' +'disclaimer.html', 'rb') as htmlFile:
            html.SetPage(htmlFile.read())
        
        sizer.Add(html, 1, wx.EXPAND, wx.ALL, 10)
        
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        acceptBtn = wx.Button(self, wx.ID_ANY, "Accept")
        self.Bind(wx.EVT_BUTTON, self.OnAccept, acceptBtn)
        
        declineBtn = wx.Button(self, wx.ID_ANY, "Decline")
        self.Bind(wx.EVT_BUTTON, self.OnDecline, declineBtn)
        
        buttonSizer.Add(acceptBtn, -1, wx.ALIGN_CENTER)
        
        buttonSizer.Add(declineBtn, -1, wx.ALIGN_CENTER)
        
        sizer.Add(buttonSizer, 0, wx.EXPAND, wx.ALIGN_CENTER | wx.ALL, 10)
        
        self.SetSizer(sizer)
        
        self.frame = parent
        
        
    def OnAccept(self, event):
        self.frame.Show()
        self.Close()
        self.Destroy()
    
    
    def OnDecline(self, event):
        self.frame.Close()
        self.frame.Destroy()
        self.Close()
        self.Destroy()
        
        
class wxHTML(wx.html.HtmlWindow):
    def OnLinkClicked(self, link):
        webbrowser.open(link.GetHref())
 
# Run the program
if __name__ == '__main__':
    app = wx.App(False)
    frame = SplashScreen(None)
    frame.Show()
    app.MainLoop()
