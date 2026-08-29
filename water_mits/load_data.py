#load_data.py
import win32com.client as win32
import os
import datetime
import re
import random
import tkinter as tk
from tkinter import filedialog

class LoadData:
    def __init__(self, filename='Предварительно.xlsx'):
        self.filename = filename
        self.path = os.path.join(os.getcwd(), self.filename)
        self.excel = None
        self.wb = None
        self.ws = None
        self.path_prot = os.path.join(os.getcwd(), 'Протокол.xlsm')
        self.wb_prot = None
        self.ws_prot = None
        self.ws_const = None
        self._last_row = 2
        self.error = 0
        self.error_num = 0

    def __enter__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        print(os.path.join(os.getcwd(), self.filename))

        if os.path.exists(self.path):
            self.wb = self.excel.Workbooks.Open(self.path)
            self.ws = self.wb.Worksheets('Реестр')
            self.wb_prot = self.excel.Workbooks.Open(self.path_prot)
            self.ws_prot = self.wb_prot.Worksheets('Протокол')
            self.ws_const = self.wb_prot.Worksheets('Проливы')
                
            self._last_row = self.ws.Cells(self.ws.Rows.Count, 1).End(-4162).Row + 1

        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        try:
            if self.wb:
                self.wb.Close()
                self.wb_prot.Save()
                self.wb_prot.Close()
            if self.excel:
                self.excel.Quit()
        except Exception as e:
            print(f"Error closing Excel: {e}")
        finally:
            self.wb = None
            self.wb_prot = None
            self.excel = None


    def get_mits_data(self):
        data = []
        for row in range(2, self._last_row):
            instrument_name = self.ws.Cells(row, 1).Value
            diameter = self.get_diameter(instrument_name)
            number = self.ws.Cells(row, 2).Value
            if 'ДНР' in number:
                number = 'Счетчик воды'
            mpi = int(self.ws.Cells(row, 3).Value)
            type_mit = self.ws.Cells(row, 4).Value
            if not isinstance(type_mit, str):
                type_mit = str(int(type_mit))
            num_manufacturer = str(self.ws.Cells(row, 5).Value)
            manufactureYear = self.ws.Cells(row, 6).Value
            if manufactureYear:
                manufactureYear = int(manufactureYear)
            docTitle = self.ws.Cells(row, 7).Value
            miOwner = self.ws.Cells(row, 9).Value
            reasons = self.ws.Cells(row, 8).Value
            date_now = datetime.datetime.now()
            act = str(self.ws.Cells(row, 10).Value)
            
            #date_now = datetime.date(2026,4,28)
            vrDate = date_now.strftime("%Y-%m-%d")
            validDate = ((date_now - datetime.timedelta(days=1)).replace(year=date_now.year + mpi)).strftime("%Y-%m-%d")
            if reasons:
                validDate = None
            self.add_data_in_protocol(type_mit, number, num_manufacturer, manufactureYear,
                                      miOwner, vrDate, validDate, reasons, diameter)
            res = {
                'mitypeNumber': number,
                'mpi': mpi,
                'modification': type_mit,
                'manufactureNum': num_manufacturer,
                'manufactureYear': manufactureYear,
                'docTitle': docTitle,
                'miOwner': miOwner,
                'reasons': reasons,
                'vrfDate': vrDate,
                'validDate': validDate,
                'act': act,
                'instrument_name': instrument_name,
                "mis": [
                    {"typeNum": "Рег. номер", "manufactureNum": "Заводской"},
                    {"typeNum": "Рег. номер", "manufactureNum": "Заводской"}
                ], # СИ применяемые при поверке
                "temperature": "Температура ˚C",
                "pressure": "Давление кПа",
                "hymidity": "Влажность %"
            }
            data.append(res)
        return data
    
    def get_diameter(self, full_name):
        ws_reg_mits = self.wb.Worksheets('Перечень')
        print(full_name)
        print(ws_reg_mits.Columns(9).Cells.Find(What=full_name, LookIn=-4163))
        finder = ws_reg_mits.Columns(9).Cells.Find(What=full_name, LookIn=-4163).Row
        diameter = ws_reg_mits.Cells(finder, 3).Value
        return int(diameter)

    def format_cell(self, obj, row, start_column, end_column, val):
        rng = obj.Range(obj.Cells(row, start_column), obj.Cells(row, end_column))
        rng.Merge()
        rng.Borders.LineStyle = 1
        if start_column == 12:
            rng.NumberFormat = "@"
        rng.Value = val

        return val

    def format_number(self, number):
        res = str(f'{float(number):.1f}').split(".")
        res = res[0][-3:] + "." + res[1][:1]
        return float(res)

    def add_data_in_protocol(self, modification, mitypeNumber, manufactureNum, manufactureYear,
                             miOwner, vrfDate, validDate, reasons, diameter):
        self.error = 0
        self.error_num = 0
        file_path = os.path.join(os.getcwd(), self.filename)
        lr = 23
        if os.path.exists(file_path):
            lr = self.ws_prot.Cells(self.ws_prot.Rows.Count, 3).End(-4162).Row + 1
            self.ws_prot.Cells(lr, 82).Value = str(vrfDate)
            self.ws_prot.Cells(lr, 83).Value = str(validDate)
            self.format_cell(self.ws_prot, lr, 72, 81, '')
            self.format_cell(self.ws_prot, lr, 1, 2, '')
            self.format_cell(self.ws_prot, lr, 3, 7, modification)
            self.format_cell(self.ws_prot, lr, 8, 11, mitypeNumber)
            self.format_cell(self.ws_prot, lr, 12, 16, str(manufactureNum))
            self.format_cell(self.ws_prot, lr, 17, 19, manufactureYear)
            self.format_cell(self.ws_prot, lr, 20, 23, miOwner)
            self.format_cell(self.ws_prot, lr, 24, 25, diameter)
            if not reasons:
                self.format_cell(self.ws_prot, lr, 26, 28, 'Соотв.')
                self.format_cell(self.ws_prot, lr, 29, 31, 'Соотв.')
                self.format_cell(self.ws_prot, lr, 68, 71, 'Годен')

            if reasons:
                if 'погрешност' in reasons.lower():
                    self.format_cell(self.ws_prot, lr, 26, 28, 'Соотв.')
                    self.format_cell(self.ws_prot, lr, 29, 31, 'Соотв.')
                    self.format_cell(self.ws_prot, lr, 68, 71, 'Не годен')

                else:
                    if 'герм' in reasons.lower():
                        self.format_cell(self.ws_prot, lr, 26, 28, 'Соотв.')
                        self.format_cell(self.ws_prot, lr, 29, 31, 'Не соотв.')

                    if 'внешний' in reasons.lower():
                        self.format_cell(self.ws_prot, lr, 26, 28, 'Не соотв.')
                        self.format_cell(self.ws_prot, lr, 29, 31, '')


                    self.format_cell(self.ws_prot, lr, 32, 34, '')
                    self.format_cell(self.ws_prot, lr, 35, 36, '')
                    self.format_cell(self.ws_prot, lr, 37, 39, '')
                    self.format_cell(self.ws_prot, lr, 40, 42, '')
                    self.format_cell(self.ws_prot, lr, 43, 45, '')
                    self.format_cell(self.ws_prot, lr, 46, 47, '')
                    self.format_cell(self.ws_prot, lr, 48, 50, '')
                    self.format_cell(self.ws_prot, lr, 51, 53, '')
                    self.format_cell(self.ws_prot, lr, 54, 56, '')
                    self.format_cell(self.ws_prot, lr, 57, 58, '')
                    self.format_cell(self.ws_prot, lr, 59, 61, '')
                    self.format_cell(self.ws_prot, lr, 62, 64, '')
                    self.format_cell(self.ws_prot, lr, 65, 67, '')
                    self.format_cell(self.ws_prot, lr, 68, 71, 'Не годен')
                    return


            #Первый пролив
            start_l = ''
            self.format_cell(self.ws_prot, lr, 32, 34, f"{start_l}")
            self.format_cell(self.ws_prot, lr, 35, 36, 120)
            etalon = self.ws_const.Columns(1).Cells.Find(diameter, LookAt=1).Row
            etalon_1 = self.ws_const.Cells(etalon, 2).Value * 120 / 3.6
            self.format_cell(self.ws_prot, lr, 40, 42, etalon_1)
            first_del = self.format_cell(self.ws_prot, lr, 43, 45, '')
            first_res = ''
            self.format_cell(self.ws_prot, lr, 37, 39, f'{first_res}')

            #Второй пролив
            self.format_cell(self.ws_prot, lr, 46, 47, 360)
            etalon_2 = self.ws_const.Cells(etalon, 4).Value * 360 / 3.6
            self.format_cell(self.ws_prot, lr, 51, 53, etalon_2)
            second_del = self.format_cell(self.ws_prot, lr, 54, 56, '')
            second_res = ''
            self.format_cell(self.ws_prot, lr, 48, 50, f'{second_res}')

            #Третий пролив
            self.format_cell(self.ws_prot, lr, 57, 58, 720)
            etalon_3 = self.ws_const.Cells(etalon, 6).Value * 720 / 3.6
            self.format_cell(self.ws_prot, lr, 62, 64, etalon_3)
            third_del = self.format_cell(self.ws_prot, lr, 65, 67, '')
            third_res = ''
            self.format_cell(self.ws_prot, lr, 59, 61, f'{third_res}')

