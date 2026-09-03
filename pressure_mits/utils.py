#utils.py
def reformat_data(data: list) -> list:
    res = []
    for el in data:
        res.append({
                'act': el.get('act'),
                'instrument_name': el.get('instrument_name'),
                'serial_number': el.get('serial_number'),
                'verification_date': el.get('verification_date'),
                'next_verification_date':  el.get('next_verification_date'),
                'note': el.get('note'),
                'mit_link': el.get('mit_link')
            })
    return res
