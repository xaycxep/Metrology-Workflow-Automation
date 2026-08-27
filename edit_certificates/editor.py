import os
import zipfile
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tkinter as tk
from tkinter import filedialog
import fitz
import tempfile
import time
import subprocess

class EditorPdf:

    def __init__(self):
        self.SUMATRA_PATH = r"вставляете свой путь\SumatraPDF\SumatraPDF.exe"
        # === НАСТРОЙКИ ===
        self.OUTPUT_PDF = 'output_files.pdf'     
        self.FONT_PATH = 'DejaVuSans.ttf'        # Шрифт (должен лежать рядом со скриптом)
        self.FONT = 'CyrillicFont'
        
        # Текст и координаты (единые для всех файлов)
        self.UNIQ_NUM = "НОМЕР" # здесь вставте уникальный номер записи об аккредитации
        self.AUTH_PERSON = 'Начальник отдела' # должность руководителя или уполномоченного лица
        self.ИМЯ_ПЕРЕМЕННОЙ = 'ФАМИЛИЯ И. О.'
        self.FONT_SIZE = 9
        self.COORD_X = 445
        self.COORD_Y = 695

        self.root = tk.Tk()
        self.root.withdraw()
        self.INPUT_ZIP = filedialog.askopenfilename(
                title='Select ZIP',
                filetypes=[('ZIP files', '*.zip'),('All files', '*.*')]
            )

        self.phrase = 'должность руководителя или'
        self.width = 0
        self.height = 0

        self.writer = PdfWriter()

    def find_cords(self, pdf_file):
        doc = fitz.open(stream=pdf_file, filetype='pdf')
        res = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            height = page.rect.height
            text_instances = page.search_for(self.phrase)
            for inst in text_instances:
                res.append(
                        {
                        'page': page_num,
                        'rect': self.height - inst.y0
                        }
                    )
        doc.close()

        return res

    def create_overlay(self, text):

        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=(self.width, self.height))

        pdfmetrics.registerFont(TTFont(self.FONT, self.FONT_PATH))
        c.setFont(self.FONT, self.FONT_SIZE)

        if isinstance(text, list):
            for item in text:
                c.drawString(item['x'], item['y'], item['text'])
        else:
            c.drawString(self.COORD_X, self.COORD_Y, text)

        c.showPage()
        c.save()

        packet.seek(0)

        return packet.read()

    def print_pdf(self):
        out_buffer = BytesIO()
        self.writer.write(out_buffer)
        out_buffer.seek(0)

        pdf_bytes = out_buffer.getvalue()

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        try:
            if os.path.exists(self.SUMATRA_PATH):
                # Тихая печать через SumatraPDF
                subprocess.run([self.SUMATRA_PATH, "-print-to-default", tmp_path], check=True)
                print(f"   🖨️ Файл отправлен на печать через SumatraPDF")
            else:
                # Запасной вариант через ShellExecute
                printer = win32print.GetDefaultPrinter()
                win32api.ShellExecute(0, "print", tmp_path, f'"{printer}"', ".", 0)
                print(f"   🖨️ Файл отправлен на печать (ShellExecute)")
            time.sleep(1)
        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass
    
    def process_zip(self):
        try:
            with zipfile.ZipFile(self.INPUT_ZIP, 'r') as zin:

                print(f"\nНайдено файлов в архиве: {len(zin.namelist())}")
                print("Начинаю обработку...\n")

                for item in zin.infolist():
                    if item.filename.lower().endswith('.pdf'):
                        pdf_data = zin.read(item.filename)

                        
                        reader = PdfReader(BytesIO(pdf_data))
                        
                        page_ = reader.pages[0]
                        self.width = float(page_.mediabox.width)
                        self.height = float(page_.mediabox.height)

                        cords = self.find_cords(pdf_data)
                        y = cords[0]['rect']

                        TEXT = [
                                {'text': self.AUTH_PERSON, 'x': 50, 'y': y + 5},
                                {'text': self.ИМЯ_ПЕРЕМЕННОЙ, 'x': 435, 'y': y + 5}
                            ]
                          
                        overlay_first_bytes = self.create_overlay(self.UNIQ_NUM)
                        overlay_first = PdfReader(BytesIO(overlay_first_bytes)).pages[0]
                              
                        overlay_last_bytes = self.create_overlay(TEXT)
                        overlay_last = PdfReader(BytesIO(overlay_last_bytes)).pages[0]

                        print(cords[0]['page'], item.filename)

                        for i, page in enumerate(reader.pages):
                            
                            if i == 0:
                                page.merge_page(overlay_first)
                                    
                            if i == cords[0]['page']:
                                page.merge_page(overlay_last)

                            self.writer.add_page(page)


            with open(self.OUTPUT_PDF, 'wb') as pdf:
                self.writer.write(pdf)

            self.print_pdf()
            
        except Exception as e:
            print(f"Err: {e}")

            

if __name__ == '__main__':
    e = EditorPdf()
    e.process_zip()
        

        
