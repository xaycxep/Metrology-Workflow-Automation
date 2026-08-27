import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os
import datetime
import requests
import xml.etree.ElementTree as ET
from xml.dom import minidom
import time

SNILS_DATA = {
        'ФАМИЛИЯ': {
            'Last': 'Фамилич',
            'First': 'Имя',
            'SNILS': 'Снилс'
            }
    }


class AppExcel:
    def __init__(self, filepath):
        self.filepath = filepath
        self.excel = None
        self.wb = None
        self.ws = None
        self.lr = None


    def __enter__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False

        if os.path.exists(self.filepath):
            self.wb = self.excel.Workbooks.Open(self.filepath)
            self.ws = self.wb.Worksheets('Sheet1')
            self.lr = self.ws.Cells(self.ws.Rows.Count, 1).End(-4162).Row

        return self


    def __exit__(self, exc_type, exc_value, exc_tb):
        try:
            if self.wb:
                #self.wb.Save()
                self.wb.Close()

            if self.excel:
                self.excel.Quit()

        except Exception as e:
            print(f'Err closing: {e}')

        finally:
            self.wb = None
            self.excel.DisplayAlerts = True
            self.excel.ScreenUpdating = True
            self.excel = None
            

    def get_sert_nums(self):
        res = dict()
        for r in range(4, self.lr + 1):
            cell = self.ws.Cells
            res[cell(r, 11).Value] = cell(r, 12).Value

        return res


root = tk.Tk()
root.withdraw()
path = filedialog.askopenfilename(
        title='Select Excel',
        filetypes=[('Excel file', '*.xls*'), ('All files', '*.*')]
    )

def format_date(date):
    if isinstance(date, str):
        res = datetime.datetime.strptime(date, '%d.%m.%Y')
        return res.strftime('%Y-%m-%d')
    return False

DATA_SET = dict()

if path:
    with AppExcel(path) as app:
        sert_nums = app.get_sert_nums()

        base_url = 'https://fgis.gost.ru/fundmetrology/eapi/vri'

        for num, val in sert_nums.items():
            print(num, val)
            year_ = num.split('/')[1].split('-')[-1]
            params = {'year': year_, 'result_docnum': num}

            req = requests.get(base_url, params)

            if req.status_code == 200:
                print(req.json())
                item = req.json()['result']['items'][0]
                
                mit_title = item['mit_title']
                ver_date = format_date(item['verification_date'])
                valid_date = format_date(item.get('valid_date', None))
                
                name = val.upper()
                last = SNILS_DATA[name]['Last']
                first = SNILS_DATA[name]['First']
                snils = SNILS_DATA[name]['SNILS']
                
                DATA_SET[num] = {
                    'ver_date': ver_date,
                    'valid_date': valid_date,
                    'mit_title': mit_title,
                    'last': last,
                    'first': first,
                    'snils': snils
                    }
                time.sleep(0.5)

    def create_xml(data, output_file='example.xml'):
        app = ET.Element('Message', {'xsi:noNamespaceSchemaLocation': "schema.xsd",
                                 'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"})
    
        ver_mes_inst_data = ET.SubElement(app, 'VerificationMeasuringInstrumentData')
        for item, val in data.items():
            ver_mes_inst = ET.SubElement(ver_mes_inst_data, 'VerificationMeasuringInstrument')
            ET.SubElement(ver_mes_inst, 'NumberVerification').text = item
            ET.SubElement(ver_mes_inst, 'DateVerification').text = val['ver_date']
            valid_date = val['valid_date']
            res_ver = 2
            if valid_date:
                ET.SubElement(ver_mes_inst, 'DateEndVerification').text = valid_date
                res_ver = 1

            ET.SubElement(ver_mes_inst, 'TypeMeasuringInstrument').text = val['mit_title']
            app_emp = ET.SubElement(ver_mes_inst, 'ApprovedEmployees')
            ver_name = ET.SubElement(app_emp, 'Name')
            ET.SubElement(ver_name, 'Last').text = val['last']
            ET.SubElement(ver_name, 'First').text = val['first']
            ET.SubElement(app_emp, 'SNILS').text = val['snils']
            ET.SubElement(ver_mes_inst, 'ResultVerification').text = str(res_ver)
        ET.SubElement(app, 'SaveMethod').text = '2'

        rough_str = ET.tostring(app, encoding='unicode')
        reparsed = minidom.parseString(rough_str)
        pretty_xml = reparsed.toprettyxml(indent="  ", encoding=None)

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
            
if __name__ == '__main__':
    create_xml(DATA_SET)
