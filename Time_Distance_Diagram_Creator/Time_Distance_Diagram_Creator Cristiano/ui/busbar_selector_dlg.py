'''
Created on 11 Oct 2018

@author: AMB
'''

import wx

class BusbarSelector ( wx.Frame ):
    
    def __init__( self, parent, interface):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 400,600 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL | wx.STAY_ON_TOP)
        
        self.SetSizeHints( wx.Size( 330,600 ), wx.Size( 330,600 ) )
        
        bSizer1 = wx.BoxSizer( wx.VERTICAL )
        
        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, u"Select one of the available busbars", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )
        
        bSizer1.Add( self.m_staticText1, 0, wx.ALL, 5 )
        
        m_BusbarslistBoxChoices = []
        self.m_BusbarslistBox = wx.ListBox( self, wx.ID_ANY, wx.Point( 0,0 ), wx.Size( 500,450 ), m_BusbarslistBoxChoices, 0 )
        bSizer1.Add( self.m_BusbarslistBox, 0, wx.ALL, 5 )
        
        bSizer2 = wx.BoxSizer( wx.VERTICAL )
        
        gSizer1 = wx.GridSizer( 0, 2, 0, 0 )
        
        self.m_OKbutton = wx.Button( self, wx.ID_ANY, u"Ok", wx.Point( 100,300 ), wx.DefaultSize, 0 )
        
        self.m_OKbutton.SetBitmapPosition( wx.BOTTOM )
        gSizer1.Add( self.m_OKbutton, 0, wx.ALL, 5 )
        
        self.m_Cancelbutton = wx.Button( self, wx.ID_ANY, u"Cancel", wx.Point( 10,300 ), wx.DefaultSize, 0 )
        
        self.m_Cancelbutton.SetBitmapPosition( wx.BOTTOM )
        gSizer1.Add( self.m_Cancelbutton, 0, wx.ALL, 5 )
        
        
        bSizer2.Add( gSizer1, 1, wx.EXPAND, 5 )
        
        
        bSizer1.Add( bSizer2, 1, wx.EXPAND, 5 )
        
        
        self.SetSizer( bSizer1 )
        self.Layout()
        
        self.Centre( wx.BOTH )
        
        self.interface = interface
        self.busbar_list = self.interface.get_available_busbars()  
        self.window = parent
        
        # Connect Events
        self.Bind( wx.EVT_ACTIVATE, self.BusbarSelectorOnActivate )
        self.m_OKbutton.Bind( wx.EVT_BUTTON, self.m_OKbuttonOnButtonClick )
        self.m_Cancelbutton.Bind( wx.EVT_BUTTON, self.m_CancelbuttonOnButtonClick )
    
    def __del__( self ):
        pass
    
    
    # Virtual event handlers, overide them in your derived class
    def BusbarSelectorOnActivate( self, event ):
        event.Skip()
        self.m_BusbarslistBox.Clear()
        if self.busbar_list is not None:
            for busbar in self.busbar_list:
                self.m_BusbarslistBox.Append(self.interface.get_name_of(busbar))
        self.m_BusbarslistBox.Enable()
        
    
    def m_OKbuttonOnButtonClick( self, event ):
        selected_busbar_index = self.m_BusbarslistBox.GetSelection()
        if selected_busbar_index != wx.NOT_FOUND:
            self.window.StudySelectedBus.AppendText(self.interface.get_name_of(self.busbar_list[selected_busbar_index]))
        event.Skip()
        self.Destroy()
    
    def m_CancelbuttonOnButtonClick( self, event ):
        event.Skip()
        self.Destroy()