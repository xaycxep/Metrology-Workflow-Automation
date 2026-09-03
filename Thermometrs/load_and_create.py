import win32com.client as win32
import os
import datetime
import random
import re
from pathlib import Path

class ConnExcel:
    def __init__(self):
        self.filename = 'Предварительно.xlsx'
        self.input_file = os.path.join(os.getcwd(), self.filename)
        self.excel = None
        self.wb = None
        self.ws = None
        self.prot_path = os.path.join(os.getcwd(), 'Протокол.xlsx')
        self.prot_wb = None
        self.prot_8_279_78 = None
        self.for_db_data = []
        

    def __enter__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False

        if os.path.exists(self.prot_path):
            self.wb = self.excel.Workbooks.Open(self.input_file)
            self.ws = self.wb.Worksheets('Реестр')
            self.prot_wb = self.excel.Workbooks.Open(self.prot_path)
            self.prot_8_279_78 = self.prot_wb.Worksheets('ГОСТ_8_279_78')


        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        try:
            if self.wb:
                self.wb.Close()

            if self.prot_wb:
                self.prot_wb.Saved = True
                self.prot_wb.Close()

            if self.excel:
                self.excel.Visible = True
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()

        except Exception as e:
            print(f'Err clos Excel: {e}')

        finally:
            self.wb = None
            self.prot_wb = None
            self.excel = None
    
    def gen_err(self, division_price, start_point, end_point):
        POINTS = {'0.01': 1,
                  '0.02': 2,
                  '0.05': 5,
                  '0.1': 10,
                  '0.2': 10,
                  '0.5': 50,
                  '1.0': 50,
                  '2.0': 50,
                  '5.0': 50,
                  '10.0': 50}
        multiplicity = POINTS[str(division_price)]
        distance = abs(start_point) + abs(end_point)
        rng = range(start_point, end_point + 1, multiplicity)
        if len(rng) < 3:
            rng = [start_point, 0, end_point]

        data = dict()
        for s in rng:
            data[s] = ['', '', '', '', '']

        return data
    
    def format_cell(self, ws, r, start_c, end_c, val):
        rng = ws.Range(ws.Cells(r, start_c), ws.Cells(r, end_c))
        rng.Merge()
        rng.Borders.LineStyle = 1
        rng.NumberFormat = "@"
        rng.Value = str(val)
        rng.HorizontalAlignment = -4108

    def ensure_file_in_directory(self, directory: str, filename: str) -> str:

        dir_path = Path(directory)
        dir_path.mkdir(parents=True, exist_ok=True)

        file_path = dir_path / filename

        return str(file_path)

    def save_ws(self, ws, num, miOwner):
        chars = r'[<>:"/\\|?*]'
        num = re.sub(chars, '_', num)
        miOwner = re.sub(chars, '_', miOwner)
        directory = f'{os.getcwd()}/Протоколы/{miOwner}'
        xlsx_path = self.ensure_file_in_directory(directory, f'Protocol_{num}.xlsx')
        ws.Copy()
        new_wb = self.excel.ActiveWorkbook
        new_wb.SaveAs(xlsx_path, FileFormat=51)
        new_wb.Close(SaveChanges=False)
        ws.Activate()
        return xlsx_path
        
    def get_mits_data(self):
        reg = self.ws
        lr = reg.Cells(reg.Rows.Count, 1).End(-4162).Row
        data = []
        for r in range(2, lr + 1):
            type_mit = reg.Cells(r, 1).Value
            name_mit = reg.Cells(r, 2).Value
            start_point = int(reg.Cells(r, 3).Value)
            end_point = int(reg.Cells(r, 4).Value)
            multiplicity = reg.Cells(r, 5).Value
            number_reg = reg.Cells(r, 6).Value
            interval = reg.Cells(r, 7).Value
            standart = reg.Cells(r, 8).Value
            instrument_name = reg.Cells(r, 9).Value
            manufacturer_num = reg.Cells(r, 10).Value
            if not isinstance(manufacturer_num, str):
                manufacturer_num = str(int(manufacturer_num))
            manufacturer_year = reg.Cells(r, 11).Value
            if manufacturer_year:
                manufacturer_year = str(int(manufacturer_year))
            reasons = reg.Cells(r, 12).Value
            vrf_date = datetime.datetime.fromtimestamp(reg.Cells(r, 13).Value.timestamp())
            vr_date = vrf_date.strftime("%d.%m.%Y")
            valid_date = ((vrf_date - datetime.timedelta(days=1)).replace(year=vrf_date.year + int(interval))).strftime("%Y-%m-%d")
            val_date = datetime.datetime.strptime(valid_date, "%Y-%m-%d").strftime("%d.%m.%Y")
            if reasons:
                val_date = None
            mi_owner = reg.Cells(r, 14).Value
            act = reg.Cells(r, 15).Value
            options = {
                'type_mit': type_mit,
                'name_mit': name_mit,
                'start_point': start_point,
                'end_point': end_point,
                'multiplicity': multiplicity,
                'number_reg': number_reg,
                'mpi': interval,
                'standart': standart,
                'manufacturer_num': manufacturer_num,
                'manufacturer_year': manufacturer_year,
                'reasons': reasons,
                'vrf_date': vrf_date.strftime("%Y-%m-%d"),
                'vr_date': vr_date,
                'valid_date': valid_date,
                'val_date': valid_date,
                'mi_owner': mi_owner,
                'instrument_name': instrument_name,
                'act': act
                       }
            data.append(options)

        return data
            
    def paste_data(self, data):
        prot = self.prot_8_279_78
        lr = prot.Cells(prot.Rows.Count, 1).End(-4162).Row
        prot.Range(prot.Cells(35, 1), prot.Cells(lr, 41)).Clear()
        points = self.gen_err(data['multiplicity'], data['start_point'], data['end_point'])
        prot.Range("AA5").Value = data['vr_date']
        prot.Range("K7").Value = f"{data['name_mit']} {data['type_mit']}"
        manufacturer_num = data['manufacturer_num']
        prot.Range("K9").Value = data['number_reg']
        prot.Range("K11").Value = manufacturer_num
        prot.Range("K13").Value = data['mi_owner']
        prot.Range("AG11").Value = data.get('manufacturer_year', '')
        prot.Range("K16").Value = f"- {data['standart']}"
        prot.Range("K31").Value = data['type_mit']
        prot.Range("U31").Value = f"от {data['start_point']} до {data['end_point']}"
        prot.Range("AF31").Value = data['multiplicity']
        prot.Range("A28").Value = f"4.1. Внешний осмотр: соответствует требованиям п. 5.1 {data['standart']}"
        for r, key in enumerate(points.items()):
            el = key[0]
            str_etalon, avr_etalon, str_mit, avr_mit, correction = key[1]
            row = r + 35
            self.format_cell(prot, row, 1, 5, el)
            self.format_cell(prot, row, 6, 11, str_etalon)
            self.format_cell(prot, row, 12, 20, avr_etalon)
            self.format_cell(prot, row, 21, 26, str_mit)
            self.format_cell(prot, row, 27, 35, avr_mit)
            self.format_cell(prot, row, 36, 41, correction)

        prot.Cells(35 + len(points) + 1, 1).Value = 'Заключение: Признан пригодным к применению'
        prot.Cells(35 + len(points) + 3, 1).Value = 'Поверитель:'
        prot.Cells(35 + len(points) + 3, 33).Value = 'Фамилия И. О.'

        prot_link = self.save_ws(prot, str(manufacturer_num), data['mi_owner'])

        self.for_db_data.append({
                'act': data.get('act'),
                'instrument_name': data.get('instrument_name'),
                'serial_number': data.get('manufacturer_num'),
                'verification_date': data.get('vrf_date'),
                'next_verification_date':  data.get('val_date'),
                'note': data.get('reasons'),
                'mit_link': prot_link               
            })
        return

    def add_data_to_bd(self):
        import sys
        import os

        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        sys.path.append(parent_dir)
        from all_base_mits.conn_db import add_mits

        add_mits(self.for_db_data)
    

