import os
import fitz

from parse_google_docs import get_full_name_representative_by_object_name, get_phone_number_by_object_name, \
    get_address_by_object_name, get_choice_licensor_by_object_name

context = {}


def append_to_context(key, value):
    context[key] = value


def get_file_from_directory(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'эп' == file.casefold().split('_')[-1].split('.pdf')[0]:
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


def set_file_data(file):
    append_to_context('doc_file', file)


def set_customer_name():
    append_to_context('customer', 'ТОО "КаР-Тел"')


def set_projector_name():
    append_to_context('projector', 'лицензия 20006068 от 13.04.2020 г., II категория')


def set_full_object_name(text_lines):
    name = ''
    start_idx = None
    end_idx = None

    for i, line in enumerate(text_lines):
        if 'ЭСКИЗНЫЙ ПРОЕКТ' in line:
            start_idx = i + 1
            end_idx = i + 3
            break

    if start_idx is not None and end_idx is not None:
        name = ' '.join(line.strip() for line in text_lines[start_idx:end_idx]).replace('\xa0', ' ')

    append_to_context('full_object_name', name)


def set_address_name(object_name):
    address = get_address_by_object_name(object_name)
    append_to_context('address', address)


def get_data_values(path):
    directory = fr'{path}'
    file = get_file_from_directory(directory)
    text = get_text_of_file(file)
    text_lines = text.split('\n')

    choice_licensor(directory.split('\\')[-2])
    personal_data(directory.split('\\')[-2])
    set_customer_name()
    set_projector_name()
    set_full_object_name(text_lines)
    set_address_name(directory.split('\\')[-2])
    set_file_data(file)

    print("context: ", context)
    return context


if __name__ == '__main__':
    get_data_values(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\Восточно-Казахстанская область\OSK_Montana\эп')