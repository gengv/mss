# coding: utf-8
from contextlib import contextmanager
import os
import pyodbc
import xlsxwriter

class Splitter(object):
    def __init__(self):
        self.mss_db_file = r'mss.mdb'
        
        if not os.path.exists(self.mss_db_file):
            raise Exception('Fatal: Failed to find db file!')
        
        self.mss_dict_by_name = {}
        self.mss_dict_by_group = {}
        self.source_file = r'wo_bom_inv.txt'
        self.output_file = r'test.xlsx'
        
        
    def process(self):
        self.get_mss_dicts()
        
        with open_workbook(self.output_file) as _wb:
            self.output_workbook = _wb
            
            _splitting_results, _origninal_lines = self.parse_source_file()
            self.export_splitting_result(_splitting_results)
            self.export_original(_origninal_lines)
        
        
        
    def get_mss_dicts(self):
        _conn = pyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};Dbq=%s;Uid=;Pwd=flex1234;' % self.mss_db_file)
        
        _cursor = _conn.cursor()
        
        _sql = '''
            SELECT GROUP_ID, ITEM_NO, AWARD from mss
        '''
        
        _cursor.execute(_sql)
        _rs = _cursor.fetchall()
        
        for _r in _rs:
            _group_id, _item_no, _award = list(_r)
            
            self.mss_dict_by_name[_item_no] = _group_id
            
            if not self.mss_dict_by_group.has_key(_group_id):
                self.mss_dict_by_group[_group_id] = []
                
            self.mss_dict_by_group[_group_id].append((_item_no, _award))
            
            
    def parse_source_file(self):
        start_flag = False
        
        _results = []
        _origninal_lines = []
        
        with open(self.source_file, 'r') as _f:
            for _i, _ln in enumerate(_f.readlines()):
                _origninal_lines.append(_ln)
                
                if _ln.startswith('ITEM_ID'):
                    start_flag = True
                    continue
                
                if start_flag:
                    if _ln.startswith('--'): 
                        continue                    
                    elif _ln.strip()=='':
                        start_flag = False
                    else:
                        _results.extend(self.parse_line(_ln))
                
        return _results, _origninal_lines
                
            
    def parse_line(self, _ln):
        _strs = [_s for _s in _ln.split(' ') if _s]        
        
        _item_no, _quantity = _strs[0], float(_strs[1])
        
        _group_id = self.mss_dict_by_name.get(_item_no, None)
            
        if _group_id is None:
            return [(_item_no, _quantity)]
        
        else:
            _results = []
            _items = self.mss_dict_by_group.get(_group_id)
            
            _total_share = 0
            for _itm in _items:
                _total_share += _itm[1]
                
            for _itm in _items:
                _results.append((_itm[0], _quantity*_itm[1]/_total_share, '%s(%s)' % (_item_no, _quantity)))
            
            return _results
        
    
        
    def export_original(self, _lines):
        _sht = self.output_workbook.add_worksheet('original')
        
        for _row, _ln in enumerate(_lines):
            for _col, _v in enumerate([_s for _s in _ln.split(' ') if _s]):
                if _col == 0:
                    _sht.write_string(_row, _col, _v)
                else:
                    try:
                        _sht.write_number(_row, _col, int(_v))
                    except:
                        _sht.write(_row, _col, _v)
                
        
        
    def export_splitting_result(self, _results):
        _sht = self.output_workbook.add_worksheet('split')
        
        _sht.write(0, 0, 'Item_Number')
        _sht.write(0, 1, 'Qauntity')
        _sht.write(0, 2, 'Original')
        
        for _row, _item in enumerate(_results):
            _sht.write(_row+1, 0, _item[0])
            _sht.write(_row+1, 1, _item[1])
            
            if len(_item)>2:
                _sht.write(_row+1, 2, _item[2])
                
    
@contextmanager
def open_workbook(filepath):
    workbook = xlsxwriter.Workbook(filepath)
    yield workbook
    workbook.close()
    
        
if __name__ == '__main__':
    spl = Splitter()
    
#     spl.process()
