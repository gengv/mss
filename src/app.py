# coding: utf-8

from ui.MyPanel import MyPanel
import wx

__version__ = '1.0'
        
if __name__ == '__main__':
    app = wx.App(False)
    
    frame = wx.Frame(None, 
                     title='MSS Splitter v%s' % __version__ ,
                     size=(600, 200))
        
    _panel = MyPanel(frame, style=wx.BORDER_DEFAULT)
    
    frame.Show()
        
    app.MainLoop()