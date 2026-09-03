import xml.etree.ElementTree as ET
from xml.dom import minidom
import datetime

from load_and_create import ConnExcel

def create_ver_xml(results, filename):
    app = ET.Element("gost:application", {
        "xmlns:gost": "urn://fgis-arshin.gost.ru/module-verifications/import/2020-06-19",
        "xmlns:xs": "http://www.w3.org/2001/XMLSchema",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
    })

    for data in results:
        res = ET.SubElement(app, 'gost:result')

        mi_info = ET.SubElement(res, 'gost:miInfo')
        single_mi = ET.SubElement(mi_info, 'gost:singleMI')

        if '-' not in data['number_reg']:
            ET.SubElement(single_mi, 'gost:crtmitypeTitle').text = data['name_mit']
        else:
            ET.SubElement(single_mi, 'gost:mitypeNumber').text = data['number_reg']

        ET.SubElement(single_mi, 'gost:manufactureNum').text = data['manufacturer_num']
        if data.get('manufacturer_year'):
            ET.SubElement(single_mi, 'gost:manufactureYear').text = data['manufacturer_year']
        ET.SubElement(single_mi, 'gost:modification').text = data['type_mit']

        ET.SubElement(res, "gost:signCipher").text = 'Шифр организации'
        ET.SubElement(res, "gost:miOwner").text = data['mi_owner']

        ET.SubElement(res, "gost:vrfDate").text = data['vrf_date']
        if not data.get('reasons'):
            ET.SubElement(res, "gost:validDate").text = data['valid_date']

        ET.SubElement(res, "gost:type").text = str(2)

        ET.SubElement(res, "gost:calibration").text = "false"

        if data.get('reasons'):
            inapplicable = ET.SubElement(res, "gost:inapplicable")
            ET.SubElement(inapplicable, "gost:reasons").text = data.get('reasons')
        else:
            applicable = ET.SubElement(res, "gost:applicable")
            ET.SubElement(applicable, "gost:signPass").text = "false"
            ET.SubElement(applicable, "gost:signMi").text = "true"

        ET.SubElement(res, "gost:docTitle").text = data['standart']
        ET.SubElement(res, "gost:metrologist").text = "Фамилия И. О."

        means = ET.SubElement(res, "gost:means")

        mieta = ET.SubElement(means, "gost:mieta")
        ET.SubElement(mieta, "gost:number").text = 'Рег номер Эталона'

      # СИ примениемы при поверке
        mis = ET.SubElement(means, "gost:mis")
        mi_elem = ET.SubElement(mis, "gost:mi")
        ET.SubElement(mi_elem, "gost:typeNum").text = 'Рег номер СИ'
        ET.SubElement(mi_elem, "gost:manufactureNum").text = 'Заводской номер СИ'
        mi_elem = ET.SubElement(mis, "gost:mi")
        ET.SubElement(mi_elem, "gost:typeNum").text = 'Рег номер СИ'
        ET.SubElement(mi_elem, "gost:manufactureNum").text = 'Заводской номер СИ'

        conditions = ET.SubElement(res, "gost:conditions")
        ET.SubElement(conditions, "gost:temperature").text = '... ˚C'
        ET.SubElement(conditions, "gost:pressure").text = '... кПа'
        ET.SubElement(conditions, "gost:hymidity").text = '... %'

        if '-' not in data.get('number_reg'):
            ET.SubElement(res, 'gost:additional_info').text = 'Примечание для СИ'

    rough_string = ET.tostring(app, encoding='unicode')
    reparsed = minidom.parseString(rough_string)
    pretty_xml = reparsed.toprettyxml(indent="  ", encoding=None)


    with open(filename, "w", encoding="utf-8") as f:
        f.write(pretty_xml)

if __name__ == '__main__':
    with ConnExcel() as ce:
        res = ce.get_mits_data()
        for el in res:
            ce.paste_data(el)
        ce.add_data_to_bd()
        doc_name = f'temperature_mits_{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")}'
        create_ver_xml(res, f'{doc_name}.xml')
        print('Complited')
