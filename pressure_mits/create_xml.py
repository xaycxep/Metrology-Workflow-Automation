#create_xml.py
import xml.etree.ElementTree as ET
from xml.dom import minidom

from load_data import LoadData

import datetime

def create_verification_xml(results: list, output_file: str):
    """
    Создаёт XML-файл для импорта в АРШИН с несколькими результатами.
    :param results: список словарей с данными для каждого СИ
    :param output_file: путь для сохранения XML
    """
    # Корневой элемент
    app = ET.Element("gost:application", {
        "xmlns:gost": "urn://fgis-arshin.gost.ru/module-verifications/import/2020-06-19",
        "xmlns:xs": "http://www.w3.org/2001/XMLSchema",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
    })

    # Добавляем несколько <result>
    for data in results:
        result = ET.SubElement(app, "gost:result")

        # gost:miInfo
        mi_info = ET.SubElement(result, "gost:miInfo")
        single_mi = ET.SubElement(mi_info, "gost:singleMI")

        if '-' not in data.get('mitypeNumber'):
            ET.SubElement(single_mi, "gost:crtmitypeTitle").text = data.get("mitypeNumber", "")
        else:
            ET.SubElement(single_mi, "gost:mitypeNumber").text = data.get("mitypeNumber", "")

        ET.SubElement(single_mi, "gost:manufactureNum").text = str(data.get("manufactureNum", ""))

        if data.get('manufactureYear'):
            ET.SubElement(single_mi, "gost:manufactureYear").text = str(data.get("manufactureYear", "-"))
        ET.SubElement(single_mi, "gost:modification").text = data.get("modification", "")

        # Организация и владелец
        ET.SubElement(result, "gost:signCipher").text = 'Шифр организации'
        ET.SubElement(result, "gost:miOwner").text = data.get("miOwner", "")

        # Даты
        ET.SubElement(result, "gost:vrfDate").text = data.get("vrfDate", "")
        if data.get("validDate"):
            ET.SubElement(result, "gost:validDate").text = data.get("validDate", "")

        # Тип (1 — первичная, 2 — периодическая)
        ET.SubElement(result, "gost:type").text = str(2)

        # Поверка / калибровка
        ET.SubElement(result, "gost:calibration").text = "false"

        # Пригодность
        if data.get('reasons'):
            inapplicable = ET.SubElement(result, "gost:inapplicable")
            ET.SubElement(inapplicable, "gost:reasons").text = data.get('reasons')
        else:
            applicable = ET.SubElement(result, "gost:applicable")
            ET.SubElement(applicable, "gost:signPass").text = "false"
            ET.SubElement(applicable, "gost:signMi").text = "true"

        # Документ и поверитель
        ET.SubElement(result, "gost:docTitle").text = data.get("docTitle", "")
        ET.SubElement(result, "gost:metrologist").text = "Фамилия И. О."

        # Средства поверки
        means = ET.SubElement(result, "gost:means")

        #uve = ET.SubElement(means, "gost:uve")
        #ET.SubElement(uve, "gost:number").text = "Рег номер ЭТАЛОНА"
        mieta = ET.SubElement(means, "gost:mieta")
        for el in data['mieta']:
            ET.SubElement(mieta, "gost:number").text = str(el)

        if "mis" in data:
            mis = ET.SubElement(means, "gost:mis")
            for mi in data["mis"]:
                mi_elem = ET.SubElement(mis, "gost:mi")
                ET.SubElement(mi_elem, "gost:typeNum").text = str(mi.split(', ')[0])
                ET.SubElement(mi_elem, "gost:manufactureNum").text = str(mi.split(', ')[1])

        # Условия
        conditions = ET.SubElement(result, "gost:conditions")
        ET.SubElement(conditions, "gost:temperature").text = data.get("temperature", "")
        ET.SubElement(conditions, "gost:pressure").text = data.get("pressure", "")
        ET.SubElement(conditions, "gost:hymidity").text = data.get("hymidity", "")

        if '-' not in data.get('mitypeNumber'):
            ET.SubElement(result, 'gost:additional_info').text = 'Примечание для СИ'


    # Форматирование XML
    rough_string = ET.tostring(app, encoding='unicode')
    reparsed = minidom.parseString(rough_string)
    pretty_xml = reparsed.toprettyxml(indent="  ", encoding=None)

    # Сохранение
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(pretty_xml)


# Пример: несколько результатов
if __name__ == "__main__":
    with LoadData() as ld:
        results_list = ld.get_mits_data()
        docName = f'pressure_mits_{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")}'
        create_verification_xml(results_list, f"{docName}.xml")
        print("XML с несколькими результатами создан: f'{docName}.xml'")

