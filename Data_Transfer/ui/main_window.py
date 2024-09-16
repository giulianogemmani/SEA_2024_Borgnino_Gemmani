'''
Created on July 20, 24

@author: MB & GG
'''

import os
import wx

import wx.lib.agw.persist as PM

from ui.main_tab import MainTab
import ui.about_dlg as about



PROGRAMFOLDER = os.path.join(os.path.expanduser(
    '~'), 'Documents', 'MB & GG', 'Data Transfer')


class MainWindow(wx.Frame):
    def __init__(self, parent, title, interface, logic):
        wx.Frame.__init__(self, parent, title=title, size=(
            1100, 400), style=wx.STAY_ON_TOP | wx.DEFAULT_FRAME_STYLE)
        
        self.interface = interface
        self.logic = logic
        
        self.CreateStatusBar()  # A Statusbar in the bottom of the window

        if self.GetName() == 'frame':
            self.SetName('MainWindow')

        # Setting up the menu.
        filemenu = wx.Menu()
        helpmenu = wx.Menu()

        menuOpen = filemenu.Append(
            wx.ID_OPEN, "&Open", " Open project options file")
        menuSave = filemenu.Append(
            wx.ID_SAVE, "Save", " Save project options file")
        menuSaveAs = filemenu.Append(
            wx.ID_SAVEAS, "Save As...", "Save project options file with new name")
        menuExit = filemenu.Append(wx.ID_EXIT, "E&xit", " Exit the program")
        menuAbout = helpmenu.Append(
            wx.ID_ABOUT, "&About", " Information about this program")

        # Creating the menubar.
        menubar = wx.MenuBar()
        # Adding the "filemenu" to the MenuBar
        menubar.Append(filemenu, "&File")
        menubar.Append(helpmenu, "&Help")
        self.SetMenuBar(menubar)  # Adding the MenuBar to the Frame content.

        # Bind menu items to event handlers.
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnSave, menuSave)
        self.Bind(wx.EVT_MENU, self.OnSaveAs, menuSaveAs)

        # Other event handler
      
        self.Bind(wx.EVT_CHAR_HOOK, self.OnKeyDown)

        # Create panel to hold widgets
        panel = wx.Panel(self)

        # Construct notebook and tabs.
        self.nb = wx.Notebook(panel, name="MainNotebook")
        self.tab1 = MainTab(self.nb, interface, logic, self)


        self.nb.AddPage(self.tab1, "Main Options")

  
        sizer = wx.BoxSizer()
        sizer.Add(self.nb, 1, wx.EXPAND)

        self._persistMgr = PM.PersistenceManager.Get()
        self._persistMgr.SetPersistenceFile(
            os.path.join(PROGRAMFOLDER, 'DefaultOptions.opt'))

        self.RegisterandRestore(self)

        self.Bind(wx.EVT_CLOSE, self.OnExit)
        self.ProjectFile = ''

        self.nb.SetSelection(0)  # Always start on main tab.

        panel.SetSizerAndFit(sizer)
   

    def RegisterandRestore(self, win):
        """Register widgets so that they can be persisted and load them with the file content"""
        if win and win.Name not in PM.BAD_DEFAULT_NAMES:
            self._persistMgr.RegisterAndRestore(win)
        for child in win.Children:
            self.RegisterandRestore(child)

    def Register(self, win):
        """Register widgets so that they can be persisted."""
        if win and win.Name not in PM.BAD_DEFAULT_NAMES:
            self._persistMgr.Register(win)
        for child in win.Children:
            self.Register(child)


    def OnKeyDown(self, Event):
        char = Event.GetKeyCode()
        if char == wx.WXK_F12:
            try:
                self.logic.workbook.close()
                self.logic.workbook = None
            except:
                pass
        Event.Skip()

    def OnAbout(self, e):
        """Display About dialog."""
        frame = about.AboutDlg(None)
        frame.Show()

    def OnExit(self, e):
        """Cleanup actions to take upon exit."""
        
        self._persistMgr.SetPersistenceFile(
            os.path.join(PROGRAMFOLDER, 'DefaultOptions.opt'))
        self.Register(self)
        self._persistMgr.SaveAndUnregister()
        e.Skip()
        self.Destroy()

    def closeWindow(self, e):
        """Cleanup actions to take upon exit."""
        #self._persistMgr = PM.PersistenceManager.Get()
        # self._persistMgr.SaveAndUnregister()
        e.Skip()
        self.Destroy()

    def OnOpen(self, e):
        """ Open project settings file."""
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(
        ), "", "Options file (.opt)|*.opt", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            self.ProjectFile = os.path.join(
                dlg.GetDirectory(), dlg.GetFilename())
            self._persistMgr.SetPersistenceFile(self.ProjectFile)
            self.RegisterandRestore(self)
        dlg.Destroy()

    def OnSave(self, e):
        """Save the configuration settings"""
        if self.ProjectFile == '':
            dlg = wx.FileDialog(self, "Save project options file", PROGRAMFOLDER, "",
                                "Project options file (.opt)|*.opt", wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            else:
                self.ProjectFile = os.path.join(
                    dlg.GetDirectory(), dlg.GetFilename())

        if self.ProjectFile != '':
            self._persistMgr.SetPersistenceFile(self.ProjectFile)
            self.Register(self)
            self._persistMgr.SaveAndUnregister()
            e.Skip()

    def OnSaveAs(self, e):
        """Save the configuration settings in a different file"""
        dlg = wx.FileDialog(self, "Save project options file", os.getcwd(
        ), "", "Project options file (.opt)|*.opt", wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return
        else:
            self.ProjectFile = dlg.GetPath()

        if self.ProjectFile != '':
            #mgr = PM.PersistenceManager.Get()
            self._persistMgr.SetPersistenceFile(self.ProjectFile)
            self.Register(self)
            self._persistMgr.SaveAndUnregister()
            e.Skip()
