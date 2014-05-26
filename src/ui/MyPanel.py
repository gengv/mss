# coding: utf-8

from core import Splitter
from wx import Panel, Button, TextCtrl, BoxSizer, StaticText, FileDialog
from wx.lib.dialogs import ScrolledMessageDialog
from xlsxwriter import Workbook
import datetime
import os
import wx


class MyPanel(Panel):
    def __init__(self, parent, style=wx.TAB_TRAVERSAL|wx.NO_BORDER):
        Panel.__init__(self, parent, style=style)

        self.textfield_filepath = TextCtrl(self, wx.ID_ANY, size=(200, 25))
        
        self.panel_control = Panel(self, size=(200, 150), )
        self.button_choose_file = Button(self.panel_control, wx.ID_ANY, 'Browse...')
        self.button_process = Button(self.panel_control, wx.ID_ANY, 'Process')
        
        
        self.Bind(wx.EVT_BUTTON, self.OnChooseFile, self.button_choose_file)
        self.Bind(wx.EVT_BUTTON, self.OnProcess, self.button_process)
        
        self.__init_layout()
        
        self.splitter = Splitter()
        
    
    def __init_layout(self):
        self.h_sizer_11 = BoxSizer(wx.HORIZONTAL)
        self.h_sizer_11.Add(self.button_choose_file, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        self.h_sizer_11.Add(self.button_process, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        self.panel_control.SetSizerAndFit(self.h_sizer_11)
        
        
        self.v_sizer_01 = BoxSizer(wx.VERTICAL)
        self.v_sizer_01.Add(StaticText(self, label="File Path:"), 0, wx.ALIGN_LEFT | wx.ALL, 10)
        self.v_sizer_01.Add(self.textfield_filepath, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        self.v_sizer_01.Add(self.panel_control, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        
        self.SetSizerAndFit(self.v_sizer_01)
        
        
    def _config(self, local_db_file, bom_querier):
        self.local_db_file = local_db_file
        self.bom_querier = bom_querier
        
        
    def OnChooseFile(self, evt):
        _file_obj = None
        _file_dlg = FileDialog(self, message=u'选择源文件...', defaultDir=os.getcwd(),
                               defaultFile='', wildcard='text file (*.txt)|*.txt', 
                               style=wx.OPEN)
        
        _file_path = None
        if _file_dlg.ShowModal() == wx.ID_OK:
            _file_path = _file_dlg.GetPath()
            
        _file_dlg.Destroy()
        
        if _file_path: 
            self.textfield_filepath.SetValue(_file_path)
            
        
    def OnProcess(self, evt):
        _source_filepath = self.textfield_filepath.GetValue().strip()
        
        if not _source_filepath:
            msg_dlg = ScrolledMessageDialog(self, u'请先选择正确的文件！', 'Error')
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
        
        else:
            try:
                if not os.path.exists(_source_filepath):
                    raise Exception(u'所选文件并不存在，请重新检查！')
                
                self.splitter.source_file = _source_filepath
                self.splitter.output_file = os.path.join(
                                                         os.path.dirname(_source_filepath), 
                                                         'MSS_Split_%s.xlsx' % datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d_%H%M%S'))
                
                self.splitter.process()
                
                _title = 'Complete!'
                _message = u'输出的文件保存于 [%s].' % self.splitter.output_file
                
                self.textfield_filepath.Clear()
                        
            except:
                _title = 'Error!'
                
                import sys
                _message = str(sys.exc_info()[1])
                
            finally:
                msg_dlg = ScrolledMessageDialog(self, _message, _title)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
            
