'''
Created on 20 July 24

@author: MB
'''

import os

import wx.html
import webbrowser

class AboutDlg(wx.Frame):
    '''
    Application about dialog
    '''
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, wx.ID_ANY, title="About Data Transfer", size=(600,700),\
                           style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)
        #panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        html = wxHTML(self)
        pathFile = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'about_text.html')
        with open(pathFile, 'rb') as html_file:
            html.SetPage(html_file.read())

        sizer.Add(html, 1, wx.EXPAND, wx.ALL, 10)
        ok_btn = wx.Button(self, wx.ID_ANY, "OK", size=(100, 25))
        self.Bind(wx.EVT_BUTTON, self.on_ok, ok_btn)
        sizer.Add(ok_btn, flag=wx.ALIGN_CENTER | wx.ALL, border=10)
        self.SetSizer(sizer)

    def on_ok(self, event):
        '''
        ok button pushed event
        '''
        self.Destroy()

        
class wxHTML(wx.html.HtmlWindow):
    def OnLinkClicked(self, link):
        webbrowser.open(link.GetHref()) 