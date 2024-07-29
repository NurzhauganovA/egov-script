import os
import re
import fitz


context = {}


def set_order_main_info(apz_file: str):
    """
        This function reads the APZ.pdf file and returns the information in a string
    """

    number_of_order = ''
    date_of_order = ''

    with fitz.open(apz_file) as pdf:
        text_lines = pdf[0].get_text().split('\n')[0]
        pattern = r'Исх:\s(\d+/\d+)\sот\s(\d{2}\.\d{2}\.\d{4})'
        match = re.search(pattern, text_lines)

        if match:
            number_of_order = match.group(1)
            date_of_order = match.group(2)

    context['number_of_order'] = number_of_order
    context['date_of_order'] = date_of_order


def get_region_address(expertise_file):
    with fitz.open(expertise_file) as pdf:
        for page in pdf:
            text = page.get_text()
            text = text.replace('\n', ' ')

            # Поиск адреса объекта
            match = re.search(r'Адрес проектируемого объекта:\s*(.+?)(?:Прошу|$)', text)
            if match:
                address = match.group(1).strip()
                return address

    return ""


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
        r'ская',  # -ская
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
        r'я$',  # -я
        r'ь$',  # -ь
        r'и$',  # -и
        r'о$',  # -о
    ]

    # Применим каждое регулярное выражение к слову
    for suffix in suffixes:
        word = re.sub(suffix, '', word)

    return word


def set_correct_address_from_portal(address):
    correct_version = {
        "Улытау": "Ұлытау",
        "Жетысу": "Жетісу",
        "Жетису": "Жетісу"
    }

    for word in correct_version:
        address = address.replace(word, correct_version[word])

    return address


def first_non_none(groups):
    suffixes = ["область", "обл", "сельский округ", "с/о", "c\о", "район", "р-н", "город", "г."]
    for group in groups:
        if group is not None and group not in suffixes:
            return group
    return ""


def set_region_name(expertise_file):
    """
    This function reads the expertise.pdf file and returns the region name and subregion name as a dictionary
    """

    address = get_region_address(expertise_file)
    adrs = replace_some_address(address)

    match_obl = re.search(
        r'([а-яА-ЯёЁ-]+)\s*область|область\s*([а-яА-ЯёЁ-]+)|([а-яА-ЯёЁ-]+)\s*обл.|обл.\s*([а-яА-ЯёЁ-]+)', adrs,
        re.IGNORECASE)
    match_sel = re.search(
        r'([а-яА-я]+)\s*сельский округ|сельский округ\s*([а-яА-я]+)|([а-яА-я]+)\s*с\Wо|с\Wо\s*([а-яА-я]+)', adrs,
        re.IGNORECASE)
    match_region = re.search(r'([а-яА-я]+)\s*район|район\s*([а-яА-я]+)|([а-яА-я]+)\s*р\Wн|р\Wн\s*([а-яА-я]+)', adrs,
                             re.IGNORECASE)
    match_city = re.search(
        r'([а-яА-Я]+)(?:\s*\([^)]+\))?\s*(г\.|город)\s*|(г\.|город)\s*(?:\s*\([^)]+\))?\s*([а-яА-Я]+)', adrs,
        re.IGNORECASE)

    context["oblast"] = set_correct_address_from_portal(remove_suffix(first_non_none(match_obl.groups() if match_obl else [])))
    context["sel_okrug"] = set_correct_address_from_portal(remove_suffix(first_non_none(match_sel.groups() if match_sel else [])))
    context["region"] = set_correct_address_from_portal(remove_suffix(first_non_none(match_region.groups() if match_region else [])))
    context["city"] = set_correct_address_from_portal(remove_suffix(first_non_none(match_city.groups() if match_city else [])))


def set_questionnaire_examination(apz_file):
    """
        This function reads the APZ.pdf file and returns the information in a string
    """

    customer = ''
    object_name_rus = ''
    object_name_kaz = ''

    with fitz.open(apz_file) as pdf:
        for page in pdf:
            text = page.get_text()
            text = text.replace('\n', ' ')
            match_customer = re.search(r'Заказчик \(застройщик, инвестор\):\s*(.+?)\s*_', text)
            match_kaz = re.search(r'Объектің атауы:\s*(.+?)\s*;', text)
            match_rus = re.search(r'Наименование объекта:\s*(.+?)\s*;', text)

            if match_customer:
                customer = match_customer.group(1).strip()
            if match_kaz:
                object_name_kaz = match_kaz.group(1).strip()
            if match_rus:
                object_name_rus = match_rus.group(1).strip()
            break

    context['customer'] = customer
    context['object_name_rus'] = object_name_rus
    context['object_name_kaz'] = object_name_kaz


def extract_last_line_with_keyword(file_names, path):
    keyword_lines = []
    keyword_string = fr'{path}'.split('\\')[-2]
    for file_name in file_names:
        archive_match = re.search(r'_([^_]+)\.rar$', file_name)
        if archive_match:
            keyword_lines.append(file_name)
        pdf_math = re.search(r'_([^_]+)\.pdf$', file_name)
        if pdf_math:
            keyword_lines.append(file_name)
        docx_match = file_name.endswith(".docx")
        if docx_match:
            keyword_lines.append(file_name)

    return keyword_lines, keyword_string


def set_file_names(files, main_directory, object_name):
    """
        This function reads the file names and returns the information in a string
    """

    context_data = []

    for file in files:
        last_part_line = file.split(object_name)[-1]
        if 'письмо' in last_part_line.casefold():
            context_data.append({'письмо': f'{main_directory}/{file}'})
        elif "задание на проектирование" in last_part_line.casefold():
            context_data.append({'задание на проектирование': f'{main_directory}/{file}'})
        elif "договор аренды" in last_part_line.casefold():
            context_data.append({'договор аренды': f'{main_directory}/{file}'})
        elif "прав док" in last_part_line.casefold():
            context_data.append({'прав-док': f'{main_directory}/{file}'})
        elif "реквизит" in last_part_line.casefold():
            if "аврора" in last_part_line.casefold():
                context_data.append({'рек-проект': f'{main_directory}/{file}'})
            if "кар-тел" in last_part_line.casefold():
                context_data.append({'рек-генпроект': f'{main_directory}/{file}'})
            if "теле" in last_part_line.casefold():
                context_data.append({'рек-генпроект': f'{main_directory}/{file}'})
        elif "апз" in last_part_line.casefold():
            context_data.append({'апз': f'{main_directory}/{file}'})
        elif "топо" in last_part_line.casefold():
            context_data.append({'топо': f'{main_directory}/{file}'})
        elif "иги" in last_part_line.casefold():
            context_data.append({'иги': f'{main_directory}/{file}'})
        elif "пир" in last_part_line.casefold():
            context_data.append({'пир': f'{main_directory}/{file}'})
        elif "ТУ" in last_part_line:
            context_data.append({'тех-условие': f'{main_directory}/{file}'})
        elif "эп" in last_part_line.casefold():
            context_data.append({'эп': f'{main_directory}/{file}'})

    return context_data


def set_project_files(files, additional_directory, object_name):
    context_data = []

    for file in files:
        last_part_line = file.split(object_name)[-1]
        if 'гп' in last_part_line.casefold():
            context_data.append({'гп': f'{additional_directory}/{file}'})
        elif 'опз' in last_part_line.casefold():
            context_data.append({'опз': f'{additional_directory}/{file}'})
        elif 'пос' in last_part_line.casefold():
            context_data.append({'пос': f'{additional_directory}/{file}'})
        elif 'пп' in last_part_line.casefold():
            context_data.append({'пп': f'{additional_directory}/{file}'})
        elif 'эг' in last_part_line.casefold():
            context_data.append({'эг': f'{additional_directory}/{file}'})
        elif 'расчет' in last_part_line.casefold():
            context_data.append({'расчет': f'{additional_directory}/{file}'})
        elif 'овос' in last_part_line.casefold():
            context_data.append({'овос': f'{additional_directory}/{file}'})
        elif 'км оиф' in last_part_line.casefold():
            context_data.append({'км-оиф': f'{additional_directory}/{file}'})
        elif 'км' in last_part_line.casefold():
            context_data.append({'км': f'{additional_directory}/{file}'})
        elif 'оиф' in last_part_line.casefold():
            context_data.append({'оиф': f'{additional_directory}/{file}'})
        elif 'рп' in last_part_line.casefold():
            context_data.append({'рп': f'{additional_directory}/{file}'})

    return context_data


def get_directories(path):
    subdirectories = ['1', '2']
    found_subdirectories = []

    for subdirectory in subdirectories:
        subdirectory_path = os.path.join(fr'{path}', subdirectory)
        if os.path.isdir(subdirectory_path):
            found_subdirectories.append(subdirectory_path)

    if len(found_subdirectories) == 2:
        return found_subdirectories[0], found_subdirectories[1]


def config_main(path):
    main_directory, additional_directory = get_directories(fr'{path}')
    object_name = os.path.basename(os.path.dirname(os.path.dirname(main_directory)))
    print("object_name:", object_name)

    files = [file for file in os.listdir(main_directory) if file.endswith(".pdf") or file.endswith(".rar") or file.endswith(".zip") or file.endswith(".docx")]
    additional_files = [file for file in os.listdir(additional_directory) if file.endswith(".pdf") or file.endswith(".rar") or file.endswith(".zip")]

    keyword_lines_list, keyword_string = extract_last_line_with_keyword(files, fr'{path}')

    for line in keyword_lines_list:
        last_part_line = line.split(object_name)[-1]
        directory = f'{main_directory}/{line}'
        if 'письмо' in last_part_line.casefold():
            set_order_main_info(directory)
        if "эп" in last_part_line.casefold():
            set_region_name(directory)
        if "апз" in last_part_line.casefold():
            set_questionnaire_examination(directory)

    context_data = set_file_names(keyword_lines_list, main_directory, object_name)
    context['files'] = context_data
    context['project_files'] = set_project_files(additional_files, additional_directory, object_name)

    for file in files:
        if 'лицензия_пд' in file.casefold():
            context['files'].append({'лицензия_пд': f'{main_directory}/{file}'})

    print(context)
    return context


if __name__ == '__main__':
    config_main(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\1. СМР ВВОД\Восточно-Казахстанская область\OSK_Pervomay\экспертиза картел')
