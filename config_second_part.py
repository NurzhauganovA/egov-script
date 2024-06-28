import os
import fitz
import re

from parse_google_docs import get_full_name_representative_by_object_name, get_phone_number_by_object_name, \
    get_choice_licensor_by_object_name

context = {}


def append_to_context(key, value):
    context[key] = value


def get_file_from_directory(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'комиссия' in file.casefold():
            return os.path.join(directory, file)


def get_order_file_from_directory(directory):
    files = os.listdir(directory)
    for file in files:
        if file.endswith('.pdf') and 'зем.проект' in file.casefold():
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


def set_file_data(commission_file, order_file):
    append_to_context('conclusion_land_commission', commission_file)
    append_to_context('order_approval_land_plot', order_file)


def extract_data(text_lines):
    print(text_lines)
    before_right_type_marker = 'предоставляется право на земельный участок)'
    right_type_marker = '(вид права на земельный участок)'
    usage_purpose_marker = '(целевое назначение земельного участка)'
    area_marker = '(площадь земельного участка)'
    location_marker = '(местоположение испрашиваемого земельного участка)'

    right_type = ""
    usage_purpose = None
    area = None
    location = None

    for i, line in enumerate(text_lines):
        if before_right_type_marker in line:
            while text_lines[i + 1] != right_type_marker:
                right_type += text_lines[i + 1].strip()
                i += 1
        elif usage_purpose_marker in line:
            if "для" in text_lines[i - 1] or "-" in text_lines[i - 1]:
                usage_purpose = text_lines[i - 1].strip()
            else:
                for j in range(i + 1, len(text_lines)):
                    if "для" in text_lines[j]:
                        usage_purpose = text_lines[j].strip()
                        break
        elif area_marker in line:
            area = text_lines[i - 1].strip().split(' ')[0]
        elif location_marker in line:
            location = text_lines[i - 1].strip()

    if "(площадь земельного участка)" in right_type:
        right_type = right_type.split("(площадь земельного участка)")[-1]

    print(f"Requested right use: {right_type}")
    print(f"Purpose use land plot: {usage_purpose}")
    print(f"Estimated dimensions land plot: {area}")
    print(f"Location land plot: {location}")

    append_to_context('requested_right_use', right_type)
    append_to_context('purpose_use_land_plot', usage_purpose)
    append_to_context('estimated_deminsions_land_plot', area)
    append_to_context('location_land_plot', location)


def get_data_values(directory):
    file = get_file_from_directory(directory)
    order_file = get_order_file_from_directory(directory)
    text = get_text_of_file(file)
    text_lines = text.split('\n')

    choice_licensor(directory.split('\\')[-1])
    personal_data(directory.split('\\')[-1])
    set_iin_bin()
    extract_data(text_lines)
    set_file_data(file, order_file)

    print("context: ", context)
    return context


if __name__ == '__main__':
    get_data_values(r'\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\СМР ВВОД\Восточно-Казахстанская область\OSK_Pervomay')
