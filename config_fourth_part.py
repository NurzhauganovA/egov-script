import os
import fitz
import re

from parse_google_docs import get_choice_licensor_by_object_name, get_full_name_representative_by_object_name, \
    get_phone_number_by_object_name, get_address_by_object_name

context = {}


def append_to_context(key, value):
    context[key] = value


def get_file_from_directory(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.docx') and 'опросник' in file.casefold():
            return os.path.join(directory, file)


def get_text_of_file(file):
    text = ''
    with fitz.open(file) as pdf:
        for page in pdf:
            text += page.get_text()

    return text


def choice_licensor(object_name):
    append_to_context('choice_licensor', get_choice_licensor_by_object_name(object_name))


def personal_data(object_name):
    full_name = get_full_name_representative_by_object_name(object_name)
    phone_number = get_phone_number_by_object_name(object_name)
    append_to_context('full_name_representative', full_name)
    append_to_context('phone_number', phone_number)


def set_iin_bin():
    append_to_context('iin_bin_applicant', '980540000397')


def set_file_data(file, directory):
    append_to_context('doc_land_plot', file)
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'прав док' in file.casefold():
            append_to_context('request_list_tech_doc', f'{directory}/{file}')


def set_name_object(text_lines):
    object_name = ''
    start_idx = None
    end_idx = None

    for i, line in enumerate(text_lines):
        if 'Наименование объекта' in line:
            start_idx = i + 1
        elif 'Срок строительства по нормам' in line:
            end_idx = i
            break

    if start_idx is not None and end_idx is not None:
        object_name = ' '.join(line.strip() for line in text_lines[start_idx:end_idx]).replace('\xa0', ' ')

    return object_name


def set_name_object_in_russian(text_lines):
    object_name = set_name_object(text_lines)
    object_name_rus = object_name.split('/')
    if len(object_name_rus) > 1:
        object_name_rus = object_name_rus[0].strip()
    else:
        object_name_rus = object_name_rus[0].strip()
    append_to_context('full_name_object_rus', object_name_rus)


def set_name_object_in_kazakh(text_lines):
    object_name = set_name_object(text_lines)
    object_name_kaz = object_name.split('/')
    if len(object_name_kaz) > 1:
        object_name_kaz = object_name_kaz[1].strip()
    else:
        object_name_kaz = object_name_kaz[0].strip()
    append_to_context('full_name_object_kaz', object_name_kaz)


def set_region_name(object_name):
    address = get_address_by_object_name(object_name)
    append_to_context('region', address)


def set_customer_name(text_lines):
    customer_name = ''
    start_idx = None
    end_idx = None

    for i, line in enumerate(text_lines):
        if 'Заказчик' in line:
            start_idx = i + 1
        elif 'Наименование объекта' in line:
            end_idx = i
            break

    if start_idx is not None and end_idx is not None:
        customer_name = ' '.join(line.strip() for line in text_lines[start_idx:end_idx]).replace('\xa0', ' ')

    return customer_name


def set_customer_in_russian(text_lines):
    customer_name = set_customer_name(text_lines)
    append_to_context('customer_name_rus', customer_name)


def set_customer_in_kazakh(text_lines):
    customer_name = set_customer_name(text_lines)
    if 'ТОО' in customer_name:
        customer_name = customer_name.split('ТОО')[1].strip()

    append_to_context('customer_name_kaz', f'{customer_name} ЖШС')


def get_data_values(path):
    directory = fr'{path}'
    file = get_file_from_directory(directory)
    text = get_text_of_file(file)
    text_lines = text.split('\n')

    choice_licensor(directory.split('\\')[-2])
    personal_data(directory.split('\\')[-2])
    set_name_object_in_russian(text_lines)
    set_name_object_in_kazakh(text_lines)
    set_iin_bin()
    set_region_name(directory.split('\\')[-2])
    set_customer_in_russian(text_lines)
    set_customer_in_kazakh(text_lines)
    set_file_data(file, directory)

    print("context: ", context)
    return context


if __name__ == '__main__':
    get_data_values(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\Костанайская область\KOS_Vladimirovka2\АПЗ')
