import json
import os
import re

import gspread

gc = gspread.service_account(filename=r'C:\Avrora\googleAvh.json')
sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1_yRgXKLGO0y4f_-CxaiBeeHkajkR1Jqx3TSJYGt4e7o/')
code_worksheet = sh.get_worksheet_by_id(854780455)


def replace_some_address(address):
    replace_words = {
        "ВКО": "Восточно-Казахстанская область",
        "СКО": "Северо-Казахстанская область",
        "ЗКО": "Западно-Казахстанская область",
        "ЮКО": "Южно-Казахстанская область"
    }

    for word in replace_words:
        address = address.replace(word, replace_words[word])

    return address


def remove_suffix(word):
    if not word:
        return ""
    # Список регулярных выражений для удаления окончаний
    suffixes = [
        r'ый$', r'ий$',  # -ый, -ий
        r'ого$', r'его$',  # -ого, -его
        r'ский$', r'ской$',  # -ский, -ской
        r'ой$',  # -ой
        r'ому$', r'ему$',  # -ому, -ему
        r'ым$', r'им$',  # -ым, -им
        r'ом$',  # -ом
        r'ая$',  # -ая
        r'ой$',  # -ой
        r'ую$',  # -ую
        r'е$',  # -е
        r'ой$',  # -ой
        r'ы$',  # -ы
        r'и$',  # -и
        r'у$',  # -у
        r'а$',  # -а
        r'я$',  # -я
        r'ь$',  # -ь
        r'и$',  # -и
        r'е$',  # -е
        r'ы$',  # -ы
        r'у$',  # -у
        r'о$',  # -о
    ]

    # Применим каждое регулярное выражение к слову
    for suffix in suffixes:
        word = re.sub(suffix, '', word)

    return word


def get_choice_licensor(address):
    adrs = replace_some_address(address)

    match_obl = re.search(r'([а-яА-ЯёЁ-]+)\s*область|область\s*([а-яА-ЯёЁ-]+)|([а-яА-ЯёЁ-]+)\s*обл|обл\s*([а-яА-ЯёЁ-]+)', adrs, re.IGNORECASE)
    match_sel = re.search(r'([а-яА-я]+)\s*сельский округ|сельский округ\s*([а-яА-я]+)|([а-яА-я]+)\s*с\Wо|с\Wо\s*([а-яА-я]+)', adrs, re.IGNORECASE)
    match_region = re.search(r'([а-яА-я]+)\s*район|район\s*([а-яА-я]+)|([а-яА-я]+)\s*р\Wн|р\Wн\s*([а-яА-я]+)', adrs, re.IGNORECASE)
    match_city = re.search(r'([а-яА-Я]+)(?:\s*\([^)]+\))?\s*(г\.|город)\s*|(г\.|город)\s*(?:\s*\([^)]+\))?\s*([а-яА-Я]+)', adrs, re.IGNORECASE)

    def first_non_none(groups):
        for group in groups:
            if group is not None and group not in ["г.", "город", "область", "обл", "сельский округ", "с.о", "район", "р-н"]:
                return group
        return ""

    data = {
        "oblast": remove_suffix(first_non_none(match_obl.groups() if match_obl else [])),
        "sel_okrug": remove_suffix(first_non_none(match_sel.groups() if match_sel else [])),
        "region": remove_suffix(first_non_none(match_region.groups() if match_region else [])),
        "city": remove_suffix(first_non_none(match_city.groups() if match_city else []))
    }

    return data


def get_data_of_pm(region):
    json_file = 'contractor_data.json'
    with open(json_file, 'r') as file:
        data = json.load(file)
        for item in data:
            if region in item:
                return data[item]


def get_full_name_representative(region):
    data = get_data_of_pm(region)
    if data:
        return data[0]

    return ""


def get_phone_number(region):
    data = get_data_of_pm(region)
    if data:
        return data[1]

    return ""


def get_purpose_use_land_plot(is_lep):
    if "лэп" in is_lep.casefold():
        return "размещения и эксплуатации опор линии электропередач"

    return "размещения и эксплуатации антенна-мачтового сооружения"


def get_order_template(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'схема' == file.casefold().split('_')[-1].split('.pdf')[0]:
            return os.path.join(directory, file)


def get_area(is_lep, area):
    if "лэп" in is_lep.casefold():
        return area.split(" ")[-1]

    return "0,0225"


def get_all_values(path):
    directory = fr'{path}'
    object_name = directory.split('\\')[-2]
    ex_dict = {}
    lep_match = {}
    non_lep_match = {}

    for row in code_worksheet.get_all_values()[2:]:
        name = row[3]
        region = row[2]
        address = row[4]
        status = row[10]
        is_lep = row[8]
        area = row[40]

        if name in object_name and region != "" and address != "":
            if "лэп" in directory.split('\\')[-3].casefold():
                if "лэп" in is_lep.casefold():
                    lep_match = {
                        "choice_licensor": get_choice_licensor(address),
                        "full_name_representative": get_full_name_representative(region),
                        "phone_number": get_phone_number(region),
                        "purpose_use_land_plot": get_purpose_use_land_plot(is_lep),
                        "order_template": get_order_template(directory),
                        "area": get_area(is_lep, area),
                        "region": region,
                        "name": name,
                        "address": address,
                        "status": status,
                        "is_lep": is_lep,
                    }
            else:
                if "лэп" not in is_lep.casefold():
                    non_lep_match = {
                        "choice_licensor": get_choice_licensor(address),
                        "full_name_representative": get_full_name_representative(region),
                        "phone_number": get_phone_number(region),
                        "purpose_use_land_plot": get_purpose_use_land_plot(is_lep),
                        "order_template": get_order_template(directory),
                        "area": get_area(is_lep, area),
                        "region": region,
                        "name": name,
                        "address": address,
                        "status": status,
                        "is_lep": is_lep,
                    }

    if lep_match:
        return lep_match
    elif non_lep_match:
        return non_lep_match

    return ex_dict


def get_choice_licensor_by_object_name(object_name):
    for row in code_worksheet.get_all_values()[2:]:
        name = row[3]
        address = row[4]

        if address and name in object_name:
            return get_choice_licensor(address)

    return ""


def get_full_name_representative_by_object_name(object_name):
    for row in code_worksheet.get_all_values()[2:]:
        name = row[3]
        region = row[2]

        if region and name in object_name:
            return get_full_name_representative(region)

    return ""


def get_phone_number_by_object_name(object_name):
    for row in code_worksheet.get_all_values()[2:]:
        name = row[3]
        region = row[2]

        if region and name in object_name:
            return get_phone_number(region)

    return ""


def get_address_by_object_name(object_name):
    for row in code_worksheet.get_all_values()[2:]:
        name = row[3]
        address = row[4]

        if address and object_name in name:
            return address

    return ""


app = get_all_values(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\Восточно-Казахстанская область\OSK_Pervomay\1 этап')
print(app)
