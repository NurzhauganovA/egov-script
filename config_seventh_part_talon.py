import os
import fitz
import re

from parse_google_docs import get_full_name_representative_by_object_name, get_phone_number_by_object_name, \
    get_choice_licensor_by_object_name, get_all_addresses_by_object_name, \
    get_full_name_construction_inspector_by_object_name

context = {}


def append_to_context(key, value):
    context[key] = value


def get_file_from_directory(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'заключение экспертизы' in file.casefold():
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


def set_chief_construction_data(object_name):
    full_name = get_full_name_construction_inspector_by_object_name(object_name)
    append_to_context('full_name_construction_inspector', full_name)


def set_customer_data():
    customer_data = {
        "company_name": 'ТОО "КаР-Тел"',
        "postcode": "10000",
        "city": "г. Астана",
        "locality": 'район "Алматы"',
        "street": "Қадырғали Жалайыри",
        "number_house": "2",
        "phone_number": "87784611005",
        "bin": "980540000397"
    }

    context['customer_data'] = customer_data


def set_location_of_the_object(object_name):
    address = get_all_addresses_by_object_name(object_name)
    context['location_of_the_object'] = address


def get_date_number_expertise_conclusion(text_lines):
    print(text_lines)
    conclusion = "ҚОРЫТЫНДЫ"

    date_number_expertise_conclusion = None

    for i, line in enumerate(text_lines):
        if conclusion in line:
            date_number_expertise_conclusion = text_lines[i - 1]
            break

    match_expertise_conclusion = re.search(r'(\d{2}\.\d{2}\.\d{4}) ж. № (.*)', date_number_expertise_conclusion)
    date_expertise_conclusion = match_expertise_conclusion.group(1)
    number_expertise_conclusion = match_expertise_conclusion.group(2)

    append_to_context('date_expertise_conclusion', date_expertise_conclusion)
    append_to_context('number_expertise_conclusion', number_expertise_conclusion)


def get_data_values(directory):
    file = get_file_from_directory(directory)
    text = get_text_of_file(file)
    text_lines = text.split('\n')
    object_name = directory.split('\\')[-1]

    choice_licensor(object_name)
    personal_data(object_name)
    set_chief_construction_data(object_name)
    set_customer_data()
    set_location_of_the_object(object_name)
    get_date_number_expertise_conclusion(text_lines)

    print("context: ", context)
    return context


if __name__ == '__main__':
    get_data_values(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\Восточно-Казахстанская область\OSK_Pervomay')
