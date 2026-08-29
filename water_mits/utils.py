#utils.py
def reformat_data(data: list) -> list:
    res = []
    for el in data:
        res.append({
                'act': el.get('act'),
                'instrument_name': el.get('instrument_name'),
                'serial_number': el.get('manufactureNum'),
                'verification_date': el.get('vrfDate'),
                'next_verification_date':  el.get('validDate'),
                'note': el.get('reasons'),
                'mit_link': 'Путь к Excel-файлу\Протокол.xlsm'
            })
    return res
