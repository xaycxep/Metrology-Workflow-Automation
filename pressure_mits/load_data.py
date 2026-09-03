#load_data.py
import win32com.client as win32
import os
import datetime
import re
from utils import reformat_data
from pathlib import Path

class LoadData:
    def __init__(self, filename='Предварительно.xlsx'):
        self.filename = filename
        self.path = os.path.join(os.getcwd(), self.filename)
        self.excel = None
        self.wb = None
        self.ws = None
        self.ws_mits = None
        self._last_row = 2
        self.mits_dict = dict()
        self.mits_cache = dict()
        self.protocol = 'Протокол.xlsm'
        self.prot_path = os.path.join(os.getcwd(), self.protocol)
        self.prot_wb = None
        self.prot_ws = None
        self.prot_2124_90 = None
        self.prot_003_04_2003 = None
        self.prot_925_85 = None

    def _load_mits_dict(self):
        ws = self.ws_mits
        lr = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        mits_dict = dict()
        for row in range(3, lr + 1):
            full_name = ws.Cells(row, 11).Value
            mits_dict[full_name] = {
                'mitype': ws.Cells(row, 1).Value,
                'down': float(ws.Cells(row, 2).Value.split('; ')[0].replace(',', '.')),
                'up': float(ws.Cells(row, 2).Value.split('; ')[1].replace(',', '.')),
                'measurement': ws.Cells(row, 3).Value,
                'accuracy': float(ws.Cells(row, 4).Value),
                'mi_name': ws.Cells(row, 5).Value,
                'minumber': ws.Cells(row, 6).Value,
                'gost_title': ws.Cells(row, 8).Value
            }
        return mits_dict
    
    def __enter__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False

        if os.path.exists(self.path):
            self.wb = self.excel.Workbooks.Open(self.path)
            self.ws = self.wb.Worksheets('Реестр')
            self.ws_mits = self.wb.Worksheets('Перечень')
            self.mits_dict = self._load_mits_dict()
            self._last_row = self.ws.Cells(self.ws.Rows.Count, 1).End(-4162).Row + 1

            self.prot_wb = self.excel.Workbooks.Open(self.prot_path)
            self.prot_2124_90 = self.prot_wb.Worksheets('МИ_2124_90')
            self.prot_003_04_2003 = self.prot_wb.Worksheets('МПУ_003_04_2003')
            self.prot_925_85 = self.prot_wb.Worksheets('МИ_925_85')
            

        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        try:
            if self.wb:
                self.wb.Close()
            if self.prot_wb:
                #self.prot_wb.Save()
                self.prot_wb.Saved = True
                self.prot_wb.Close()
            if self.excel:
                self.excel.Visible = True
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
                self.excel.Quit()            
        except Exception as e:
            print(f"Error closing Excel: {e}")
        finally:
            self.wb = None
            self.prot_wb = None
            self.excel = None


    def get_mits_data(self):
        data = []
        for row in range(2, self._last_row):
            instrument_name = self.ws.Cells(row, 1)
            number = self.ws.Cells(row, 8).Value
            mpi = int(self.ws.Cells(row, 9).Value)
            type_mit = self.ws.Cells(row, 7).Value
            if not isinstance(type_mit, str):
                type_mit = str(int(type_mit))
            num_manufacturer = self.ws.Cells(row, 2).Value
            if not isinstance(num_manufacturer, str):
                num_manufacturer = str(int(num_manufacturer))
            manufactureYear = self.ws.Cells(row, 3).Value
            if manufactureYear:
                manufactureYear = int(manufactureYear)
            docTitle = self.ws.Cells(row, 10).Value
            miOwner = self.ws.Cells(row, 5).Value
            reasons = self.ws.Cells(row, 6).Value
            date_now = datetime.datetime.fromtimestamp(self.ws.Cells(row, 4).Value.timestamp())
            vrDate = date_now.strftime("%Y-%m-%d")
            validDate = ((date_now - datetime.timedelta(days=1)).replace(year=date_now.year + mpi)).strftime("%Y-%m-%d")
            if reasons:
                validDate = None
            mieta = self.ws.Cells(row, 11).Value.split(', ')
            
            mis = self.ws.Cells(row, 12).Value.split('; ')
            temperature = self.ws.Cells(row, 13).Value
            hymidity = self.ws.Cells(row, 15).Value
            pressure = self.ws.Cells(row, 14).Value
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
                'mieta': mieta,
                "mis": mis,
                "temperature": f"{temperature} ˚C",
                "pressure": f"{pressure} кПа",
                "hymidity": f"{hymidity} %"
            }
            data.append(res)
        return data   


    def ensure_file_in_directory(self, directory: str, filename: str) -> str:

        dir_path = Path(directory)
        dir_path.mkdir(parents=True, exist_ok=True)

        file_path = dir_path / filename

        return str(file_path)


    def _protocol_to_pdf(self, num, type_mit, miOwner):
        chars = r'[<>:"/\\|?*]'
        type_mit = re.sub(chars, '_', type_mit)
        num = re.sub(chars, '_', num)
        miOwner = re.sub(chars, '_', miOwner)
        directory = f'{os.getcwd()}/Протоколы/{miOwner}'
        filename = f'Protocol_{type_mit}_{num}.xlsx'
        xlsx_path = self.ensure_file_in_directory(directory, filename)
        '''
        self.prot_ws.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
        '''
        self.prot_ws.Copy()  # создаётся новая книга с одним листом
        new_wb = self.excel.ActiveWorkbook

        # Сохраняем новую книгу
        new_wb.SaveAs(xlsx_path, FileFormat=51)

        # Закрываем новую книгу, не сохраняя (уже сохранена)
        new_wb.Close(SaveChanges=False)

        # Возвращаем фокус на исходную книгу (на случай, если потребуется дальше работать с ней)
        self.prot_wb.Activate()

        return xlsx_path

    
    def _format_cell(self, r, start, end, val):
        cell = self.prot_ws.Range(self.prot_ws.Cells(r, start), self.prot_ws.Cells(r, end))
        cell.Merge()
        cell.Borders.LineStyle = 1
        cell.Value = val
        
    def save_protocol(self, data):
        data_prots = {
            'МПУ 003/04-2003': self.prot_003_04_2003,
            'МИ 2124-90': self.prot_2124_90,
            'МИ 925-85': self.prot_925_85,
            'МП 4212-114-64115539-2022': self.prot_003_04_2003,
            'МП 406123-2018': self.prot_2124_90
            }
        mitype = data['mitype']
        num = data['num']
        down = data['down']
        up = data['up']
        measurement = data['measurement']
        accuracy = data['accuracy']
        mi_name = data.get('mi_name', ' ')
        minumber = data['minumber']
        gost_name = data['gost_name']
        gost = data['gost']
        P = data['P']
        N1 = data['N1']
        N2 = data['N2']
        abs_error_direct = data['abs_error_direct']
        abs_error_reverse = data['abs_error_reverse']
        reduced_error_direct = data['reduced_error_direct']
        reduced_error_reverse = data['reduced_error_reverse']
        variation_value = data['variation_value']
        get_name_gost = data_prots.get(gost_name, self.prot_003_04_2003)
        self.prot_ws = get_name_gost
        self.prot_ws.Range(self.prot_ws.Cells(41, 1), self.prot_ws.Cells(55, 38)).Clear()
        cell = self.prot_ws.Range
        cell("K7").Value = f'{mi_name}\n{mitype}, пределы измерения: {down}...{up} {measurement}, кл. т. {accuracy}'
        cell("K9").Value = minumber
        cell("K11").Value = str(num)
        miOwner = data['miOwner']
        cell("K13").Value = miOwner
        cell("K16").Value = gost
        cell("AA5").Value = data['vrDate']
        
        for i, el in enumerate(P):
            r = 41
            self._format_cell(r + i, 1, 4, el)

        for i, el in enumerate(N1):
            r = 41
            self._format_cell(r + i, 5, 9, el)

        for i, el in enumerate(N2):
            r = 41
            self._format_cell(r + i, 10, 14, el)

        for i, el in enumerate(abs_error_direct):
            r = 41
            self._format_cell(r + i, 15, 19, el)

        for i, el in enumerate(abs_error_reverse):
            r = 41
            self._format_cell(r + i, 20, 24, el)

        for i, el in enumerate(reduced_error_direct):
            r = 41
            self._format_cell(r + i, 25, 29, el)

        for i, el in enumerate(reduced_error_reverse):
            r = 41
            self._format_cell(r + i, 30, 34, el)

        for i, el in enumerate(variation_value):
            r = 41
            self._format_cell(r + i, 35, 38, el)

        cell("A53").Value = 'Заключение: признан пригодным к применению'
        cell("A55").Value = 'Поверитель:'
        cell("AD55").Value = 'Фамилия И. О.'
            
        return self._protocol_to_pdf(num, mitype, miOwner)
        
    def get_mit_info(self, full_name: str, num) -> dict:
        gosts = {
                'МПУ 003/04-2003': 'МАНОМЕТРЫ, ВАКУУММЕТРЫ, МАНОВАКУУММЕТРЫ, НАПОРОМЕРЫ, ТЯГОМЕРЫ И ТЯГОНАПОРОМЕРЫ ПОКАЗЫВАЮЩИЕ И САМОПИШУЩИЕ',
                'МИ 2124-90': 'Манометры, вакуумметры, мановакуумметры, напоромеры, тягомеры и тягонапоромеры показывающие и самопишущие',
                'МИ 925-85': 'МАНОМЕТРЫ, ВАКУУММЕТРЫ И МАНОВАКУУММЕТРЫ ПОКАЗЫВАЮЩИЕ И САМОПИШУЩИЕ',
                'МП 4212-114-64115539-2022': 'ГСИ. Манометры, вакуумметры, мановакуумметры, напоромеры, тягомеры и тягонапоромеры ФТ',
                'РТ-МП-50-443-2023': 'Манометры, вакуумметры и мановакуумметры показывающие деформационные',
                'МЦКЛ.0315.МП': 'Манометры, вакуумметры, мановакуумметры, напоромеры, тягомеры и тягонапоромеры',
                'ГОСТ 15614-70': 'Манометры избыточного давления показывающие',
                'Инструкция Госкомитета 4-53': 'Манометры показывающие обыкновенные',
                'МП 406121-2018 С изменением №1': 'Манометры показывающие',
                'МП 406123-2018': 'Манометры показывающие',
                'ГОСТ 8.053-73': ' Тягомеры, напоромеры, тягонапоромеры мембранные показывающие',
                'Раздел 3 2В0.283.979 РЭ': 'Тягомеры, напоромеры, тягонапоромеры, дифманометры-тягомеры, дифманометры-напоромеры, дифманометры-тягонапоромеры'
            }
        
        if full_name in self.mits_cache:
            return self.mits_cache[full_name]
        
        result = self.mits_dict[full_name]
        result['gost'] = f'- {result["gost_title"]} "{gosts[result["gost_title"]]}"'
        result['num'] = num
        
        
        self.mits_cache[full_name] = result

        return result
    
    def create_protocol(self):
        
        POINTS = {
            0.4: 10,
            0.6: 10,
            1.0: 8,
            1.5: 8,
            1.6: 8,
            2.5: 8,
            4: 5
        }
        
        # Функция для форматирования чисел
        def fmt_num(value, decimals=3, width=8, with_sign=False):
            """
            Форматирует число, убирая минус у значений, близких к нулю
            
            Args:
                value: число для форматирования
                decimals: количество знаков после запятой
                width: ширина поля
                with_sign: True - со знаком, False - без знака
            """
            # Порог для сравнения (половина последнего разряда)
            threshold = 0.5 * 10**(-decimals)
            
            # Если значение близко к нулю, считаем его нулем
            if abs(value) < threshold:
                value = 0.0
            
            # Форматирование
            if with_sign:
                if value == 0.0:
                    # Для нуля без знака
                    return f"{0.0:>{width}.{decimals}f}"
                else:
                    # Для ненулевых со знаком
                    return f"{value:>{width}.{decimals}f}"  # или f"{value:+>{width}.{decimals}f}"
            else:
                # Без знака
                return f"{value:>{width}.{decimals}f}"
        
        # Определяем последнюю строку
        lr = self.ws.Cells(self.ws.Rows.Count, 1).End(-4162).Row
        
        for row_idx in range(2, lr + 1):
            data_mit = dict()
            full_name = self.ws.Cells(row_idx, 1).Value
            num = self.ws.Cells(row_idx, 2).Value
            vrDate = self.ws.Cells(row_idx, 4).Value
            miOwner = self.ws.Cells(row_idx, 5).Value
            if not isinstance(num, str):
                num = str(int(num))
            act = self.ws.Cells(row_idx, 16).Value
            mi_info = self.get_mit_info(full_name, num)
            
            # Извлекаем данные
            mitype = mi_info['mitype']
            down = mi_info['down']
            up = mi_info['up']
            measurement = mi_info['measurement']
            accuracy = mi_info['accuracy']
            
            # Количество точек и диапазон
            error_rate = POINTS[accuracy]
            ambit = up - down
            step = ambit / error_rate
            
            # Точки измерения
            mi_points = [round(down + j * step, 2) for j in range(error_rate + 1)]
            delta_max = abs(round(accuracy * up / 100, 3))
            
            direct_readings = []
            reverse_readings = []
            
            # Прямой ход
            for P in mi_points:
                direct_readings.append('')
            
            # Обратный ход
            for P in reversed(mi_points):
                reverse_readings.append('')
            
            reverse_readings.reverse()
            
            # === Вывод результатов ===
            print(f"\n{'='*80}")
            print(f"Тип СИ: {mitype}")
            print(f"Диапазон: {down}...{up} {measurement}")
            print(f"Класс точности: {accuracy}%")
            print(f"{'='*80}")
            data_mit['mitype'] = mitype
            data_mit['num'] = num
            data_mit['down'] = down
            data_mit['up'] = up
            data_mit['measurement'] = measurement
            data_mit['accuracy'] = accuracy
            data_mit['mi_name'] = mi_info['mi_name']
            data_mit['minumber'] = mi_info['minumber']
            data_mit['gost'] = mi_info['gost']
            data_mit['gost_name'] = mi_info['gost_title']
            data_mit['miOwner'] = miOwner
            data_mit['vrDate'] = vrDate
            data_mit['P'] = []
            data_mit['N1'] = []
            data_mit['N2'] = []
            data_mit['abs_error_direct'] = []
            data_mit['abs_error_reverse'] = []
            data_mit['reduced_error_direct'] = []
            data_mit['reduced_error_reverse'] = []
            data_mit['variation_value'] = []
            
            for j in range(len(mi_points)):
                data_mit['P'].append('')
                data_mit['N1'].append('')
                data_mit['N2'].append('')
                data_mit['abs_error_direct'].append('')
                data_mit['abs_error_reverse'].append('')
                data_mit['reduced_error_direct'].append('')
                data_mit['reduced_error_reverse'].append('')
                data_mit['variation_value'].append('')

            
            link_prot = self.save_protocol(data_mit)

            date_now = datetime.datetime.fromtimestamp(vrDate.timestamp())
            vrDate = date_now.strftime("%Y-%m-%d")
            mpi = int(self.ws.Cells(row_idx, 9).Value)
            validDate = ((date_now - datetime.timedelta(days=1)).replace(year=date_now.year + mpi)).strftime("%Y-%m-%d")
            reform = reformat_data([{
                    'act': act,
                    'instrument_name': full_name,
                    'serial_number': num,
                    'verification_date': str(vrDate),
                    'next_verification_date':  str(validDate),
                    'note': self.ws.Cells(row_idx, 6).Value,
                    'mit_link': link_prot
                }])

            import sys
            import os

            parent_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

            sys.path.append(parent_dir)
            from all_base_mits.conn_db import add_mits

            add_mits(reform)
            

with LoadData() as ld:
    ld.create_protocol()

