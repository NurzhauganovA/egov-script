import ctypes
import os

import fitz
import pygetwindow as gw
import openpyxl
import pyperclip
import pyautogui
import time
import sys

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QLineEdit

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from parse_google_docs import get_choice_licensor_by_object_name, get_full_name_representative_by_object_name, \
    get_phone_number_by_object_name, get_address_by_object_name
from logging_modules import get_file_logger


context = {}
file_name = "fourth-part.log"


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


created_orders = []
absolute_path = os.path.dirname(os.path.abspath(__file__))


def moveAllWindows():
    time.sleep(2)
    window = None

    all_windows = gw.getAllWindows()
    for window in all_windows:
        if window.title != "":
            window = window
            break

    window.minimize()
    window.restore()
    window.activate()
    window.maximize()
    window.moveTo(0, 0)
    window.resizeTo(1920, 1080)
    time.sleep(2)


def get_current_keyboard_layout():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
    layout_id = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    return layout_id


def switch_to_english_keyboard():
    english_layout = 67699721
    current_layout = get_current_keyboard_layout()

    while current_layout != english_layout:
        pyautogui.hotkey('alt', 'shift')
        current_layout = get_current_keyboard_layout()
        time.sleep(0.1)
        print(current_layout)


def sing_in_ncalayer(search):
    search_certificate_points = (1632, 80)  # (1574, 105)
    pyautogui.click(search_certificate_points)
    pyperclip.copy(search)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(1)
    AUTH_file_points = (444, 214)  # (711, 276)
    pyautogui.click(AUTH_file_points)
    pyautogui.press('enter')


def automate_ncalayer(search):
    # time.sleep(10)
    try:
        moveAllWindows()
        eds_directory_path = r"\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\ЭЦП КаР-Тел"
        password = "May2021"
        choose_eds = (538, 606)  # (710, 622)
        pyautogui.click(choose_eds)
        time.sleep(1)
        moveAllWindows()

        switch_to_english_keyboard()

        type_eds_path = (1380, 80)  # (1220, 100)
        pyautogui.click(type_eds_path)

        pyperclip.copy(eds_directory_path)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1)
        sing_in_ncalayer(search)

        write_password_points = (356, 654)  # (498, 697)
        pyautogui.click(write_password_points)
        pyperclip.copy(password)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1)
        moveAllWindows()
        pyautogui.press('enter')
        time.sleep(1)
        sign_certificate_points = (870, 911)  # (843, 1001)
        pyautogui.click(sign_certificate_points)
    except Exception as e:
        raise ValueError(f"Ошибка automate_ncalayer(): {e}")


def main(path):
    context = get_data_values(path)
    choice_licensor_arg = context['choice_licensor']
    full_name_representative_arg = context['full_name_representative']
    phone_number_arg = context['phone_number']
    full_name_object_rus_arg = context['full_name_object_rus']
    full_name_object_kaz_arg = context['full_name_object_kaz']
    region_arg = context['region']
    customer_rus_arg = context['customer_name_rus']
    customer_kaz_arg = context['customer_name_kaz']
    bin_arg = context['iin_bin_applicant']
    doc_land_plot_arg = context['doc_land_plot']
    request_list_tech_doc = context['request_list_tech_doc']

    empty_fields = [key for key, value in context.items() if value == ""]
    if empty_fields:
        get_file_logger(file_name).error(f"Не заполнены поля: {', '.join(empty_fields)}")

    try:
        options = Options()
        os.makedirs(os.path.join(absolute_path, "files"), exist_ok=True)
        download_directory = os.path.join(absolute_path, "files")  # enter the path of the download file
        options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        driver = webdriver.Chrome(options=options)
        driver.get("https://www.egov.kz")
        driver.maximize_window()

        driver.implicitly_wait(5)

        usermenu_block = driver.find_elements(By.CLASS_NAME, "usermenu-block")
        login_button = usermenu_block[1].find_elements(By.TAG_NAME, "a")[0]
        login_button.click()
        driver.implicitly_wait(5)

        login_menu_navbar_blocks = driver.find_elements(By.ID, "myTab2")[0]
        login_menu_navbar_blocks.find_elements(By.TAG_NAME, "li")[1].click()
        button_select_certificate = driver.find_element(By.ID, "buttonSelectCert")
        button_select_certificate.click()
        driver.implicitly_wait(5)

        automate_ncalayer("AUTH")

        auth = authorization(driver)
        if auth:
            change_egov_language_to_russian(driver)
            search_by_query_in_portal(driver, "эскиз", choice_licensor_arg, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                                      full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc)

    except Exception as e:
        get_file_logger(file_name).error(f"Ошибка при обработке заявки: {e}")
        raise ValueError(f"Ошибка при обработке заявки: {e}")


def authorization(driver):
    try:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.CLASS_NAME, "full-name"))
        )
        return element
    except Exception as e:
        print(e)


def change_egov_language_to_russian(driver):
    languages = driver.find_element(By.CLASS_NAME, "languages")
    languages.find_elements(By.TAG_NAME, "a")[1].click()


def change_elicense_language_to_russian(driver):
    languages = driver.find_element(By.CLASS_NAME, "new-languages")
    languages.find_elements(By.TAG_NAME, "a")[1].click()


def download_order_file(driver, order_number):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "mainDropbtnIndLK"))
    ).click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "myDropdownMainLK"))
    ).find_elements(By.TAG_NAME, 'a')[2].click()

    time.sleep(2)
    close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

    if close_table:
        close_table[0].click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'GlobalNumberStr'))
    ).send_keys(order_number)
    time.sleep(2)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'submit'))
    ).click()

    time.sleep(2)
    close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

    if close_table:
        close_table[0].click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[5]/div[2]/div[3]/div/div/table/tbody/tr[2]/td[4]/div/a'))
    ).click()

    time.sleep(2)
    close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

    if close_table:
        close_table[0].click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'btnDownloadRequest'))
    ).click()

    time.sleep(2)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'menuitem-1010-itemEl'))
    ).click()

    time.sleep(5)

    file_path = os.path.join(absolute_path, 'files', order_number + '_ru.pdf')

    if os.path.exists(file_path):
        get_file_logger(file_name).info(f"Файл заказа {order_number} скачен.")
    else:
        get_file_logger(file_name).error(f"Файл заказа {order_number} не скачен.")

    return file_path


def search_by_query_in_portal(driver, query, choice_licensor_arg, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                              full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc):

    try:
        search = driver.find_element(By.ID, "edit-query")
        search.send_keys(query, Keys.ENTER)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "search-results"))
        )
        search_results = driver.find_element(By.CLASS_NAME, "search-results")
        search_results.find_elements(By.TAG_NAME, "div")[0].find_element(By.TAG_NAME, "p").find_element(By.TAG_NAME,
                                                                                                        "a").click()
        driver.find_element(By.ID, "sticky-wrapper").find_element(By.TAG_NAME, "div").find_elements(By.TAG_NAME, "a")[
            0].click()
        driver.implicitly_wait(5)
        driver.switch_to.window(driver.window_handles[1])
        elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                         full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc)
    except Exception as e:
        raise ValueError(f"Ошибка search_by_query_in_portal(): {e}")


def elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                     full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc):
    try:
        driver.get("https://elicense.kz")
        change_elicense_language_to_russian(driver)
        # TODO: Осы жерден бастап `Строительство` деген бағананы таңдап, әрі қарай жалғастырып код жазу керек. ФИО, және т.б. данные толтыратын жерге дейін жазу керек. Дальше `create_order()` метод жалғастырады.
        driver.execute_script("window.scrollBy(0, 500);")
        driver.find_element(By.ID, "new-service-button-af-14").click()  # click 'Строительство'
        driver.implicitly_wait(30)
        driver.find_elements(By.CLASS_NAME, "new-detail-single")[0].find_elements(By.TAG_NAME, "a")[1].click()  # click 'Предоставление исходных материалов при разработке'
        driver.implicitly_wait(100)
        numeric = driver.find_elements(By.CLASS_NAME, "numeric")[0]
        first_li = numeric.find_elements(By.TAG_NAME, "li")[0]
        links = first_li.find_elements(By.CLASS_NAME, "new-status-a")
        links[0].click()
        driver.implicitly_wait(100)
        driver.find_element(By.CLASS_NAME, "new-order-online").click()

        search_data = choice_licensor_arg

        oblast = search_data.get("oblast", "")
        sel_okrug = search_data.get("sel_okrug", "")
        region = search_data.get("region", "")
        city = search_data.get("city", "")

        regions_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "treeSection"))
        ).find_element(By.CLASS_NAME, "tree").find_element(By.TAG_NAME, "ul")

        # Находим все элементы списка областей
        oblast_list = regions_box.find_elements(By.TAG_NAME, "li")

        try:
            if city:
                for oblast_item in oblast_list:
                    oblast_name = oblast_item.find_element(By.TAG_NAME, "span").text
                    if city in oblast_name:
                        oblast_item.find_element(By.TAG_NAME, "ins").click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.TAG_NAME, "ul"))
                        )
                        oblast_item.find_element(By.TAG_NAME, "li").find_element(By.TAG_NAME, "span").click()
                        break
            found_list = []
            for oblast_item in oblast_list:
                oblast_name = oblast_item.find_element(By.TAG_NAME, "span").text
                if oblast in oblast_name:
                    oblast_item.find_element(By.TAG_NAME, "ins").click()
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "ul"))
                    )
                    sel_okrug_list = oblast_item.find_elements(By.TAG_NAME, "li")
                    city_list = oblast_item.find_elements(By.TAG_NAME, "li")

                    if sel_okrug:
                        for sel_okrug_item in sel_okrug_list:
                            sel_okrug_name = sel_okrug_item.find_element(By.TAG_NAME, "span").text
                            if sel_okrug in sel_okrug_name:
                                clickable_obj = sel_okrug_item.find_element(By.TAG_NAME, "span")
                                found_list.append(clickable_obj)
                                break
                    if region:
                        for region_item in sel_okrug_list:
                            region_name = region_item.find_element(By.TAG_NAME, "span").text
                            if region in region_name:
                                clickable_obj = region_item.find_element(By.TAG_NAME, "span")
                                if len(found_list) > 1 and clickable_obj.text in [item.text for item in found_list]:
                                    found_list = [clickable_obj]
                                else:
                                    found_list.append(clickable_obj)
                                break
                    if city:
                        for city_item in city_list:
                            city_name = city_item.find_element(By.TAG_NAME, "span").text
                            if city in city_name:
                                clickable_obj = city_item.find_element(By.TAG_NAME, "span")
                                if len(found_list) > 1 and clickable_obj.text in [item.text for item in found_list]:
                                    found_list = [clickable_obj]
                                else:
                                    found_list.append(clickable_obj)
                                break

                    if found_list:
                        found_list[0].click()
                        break
                    break
        except Exception as e:
            get_file_logger(file_name).error(f"Ошибка при выборе адреса лицензий: {e}")

        time.sleep(10)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "newRequest"))
        ).click()

        time.sleep(5)

        create_order(driver, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                     full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc)
    except Exception as e:
        get_file_logger(file_name).error(f"Ошибка при создании заявки: {e}")
        raise ValueError(f"Ошибка elicense_new_tab(): {e}")


def create_order(driver, full_name_representative_arg, phone_number_arg, full_name_object_rus_arg,
                 full_name_object_kaz_arg, region_arg, customer_rus_arg, customer_kaz_arg, bin_arg, doc_land_plot_arg, request_list_tech_doc):
    try:
        time.sleep(3)
        number_orders = driver.find_element(By.ID, "panel-1010-formTable").find_elements(
            By.TAG_NAME, "tbody"
        )[0].find_element(By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td"
        )[0].find_element(By.TAG_NAME, "input").get_attribute('value')
        print(number_orders)

        time.sleep(5)
        driver.implicitly_wait(100)

        fio = \
            driver.find_element(By.ID, "panel-1011-formTable").find_elements(By.TAG_NAME, "tbody")[1].find_element(
                By.TAG_NAME,
                "tr").find_elements(
                By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input")
        phone_number = \
            driver.find_element(By.ID, "panel-1014-formTable").find_elements(By.TAG_NAME, "tbody")[-2].find_element(
                By.TAG_NAME,
                "tr").find_elements(
                By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input")
        phone_number.clear()

        driver.implicitly_wait(100)
        fio.send_keys(full_name_representative_arg)
        phone_number.send_keys(phone_number_arg)

        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"

        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.implicitly_wait(100)

        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[1].click()  # button "Next"

        time.sleep(15)
        driver.implicitly_wait(100)

        # TODO: Менің комментировать етіп қойған кодыма қарап, түсініп, жазып көрсең болады. Важный момент, `+` басып, файл таңдайтын кезде `switch to iframe` метод қолдану керек.
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[0].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(full_name_object_rus_arg)  # Полное наименование объекта russian
        print(driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[0].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").text)
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[1].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(full_name_object_kaz_arg)  # Полное наименование объекта kazakh
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[2].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(region_arg)  # region
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[3].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(customer_rus_arg)  # customer_rus
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[4].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(customer_kaz_arg)  # customer_kaz
        driver.find_element(By.ID, "panel-1012-formTable").find_elements(By.TAG_NAME, "tbody")[7].find_element(
            By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "input").send_keys(bin_arg)  # bin

        # driver.find_element(By.ID, "button-1023-btnWrap").click()
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1023-btnIconEl"))
        ).click()

        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )
        driver.switch_to.frame(iframe)

        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "new-lk-wrapper"))
            ).find_element(By.TAG_NAME, "div").find_elements(By.TAG_NAME, "button")[1].click()
        except Exception as e:
            print(e)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "file"))
        ).send_keys(request_list_tech_doc)

        time.sleep(5)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-btns"))
        ).find_elements(By.TAG_NAME, "div")[0].find_element(By.TAG_NAME, "input").click()

        time.sleep(15)
        driver.implicitly_wait(100)

        try:
            documents = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DocumentTable"))
            ).find_element(By.TAG_NAME, "tbody")

            documents.find_elements(By.TAG_NAME, "tr")[1].find_elements(By.TAG_NAME, "td")[0].find_element(
                By.TAG_NAME, "a").click()
        except Exception as e:
            get_file_logger(file_name).error(f"Ошибка при загрузке документа: {e}")
            print(e)

        time.sleep(5)
        driver.switch_to.default_content()

        time.sleep(2)

        # Create and upload the document request_list_tech_doc
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1028-btnWrap"))
        ).click()

        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )
        driver.switch_to.frame(iframe)

        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "new-lk-wrapper"))
            ).find_element(By.TAG_NAME, "div").find_elements(By.TAG_NAME, "button")[1].click()
        except Exception as e:
            print(e)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "file"))
        ).send_keys(doc_land_plot_arg)

        time.sleep(5)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-btns"))
        ).find_elements(By.TAG_NAME, "div")[0].find_element(By.TAG_NAME, "input").click()

        time.sleep(20)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div/div/div[3]/div/form/div/table/tbody/tr[2]/td[1]/a"))
        ).click()

        driver.switch_to.default_content()

        time.sleep(5)
        driver.implicitly_wait(20)

        driver.find_element(By.ID, "toolbar-1017-targetEl").find_elements(By.TAG_NAME, "div")[
            0].click()  # button "Save"
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1017-targetEl").find_elements(By.TAG_NAME, "div")[
            2].click()  # button "Next"
        driver.implicitly_wait(100)

        #  Confirm the order with EDS
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-login-tab"))
        ).find_elements(By.TAG_NAME, "a")[1].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "nextButton"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "certTypeList"))
        ).find_element(By.TAG_NAME, "input").click()

        try:
            while check_last_window() != "NCALayer":
                time.sleep(2)
        except Exception as e:
            print(e)
            print("NCALayer window not found")
            time.sleep(15)

        automate_ncalayer("GOST")
        time.sleep(20)

        close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

        if close_table:
            close_table[0].click()

        # after confirmation use EDS
        file_path = download_order_file(driver, number_orders)

        save_file_in_server(number_orders, file_path, request_list_tech_doc)
        time.sleep(5)

        driver.quit()
    except Exception as e:
        raise ValueError(f"Ошибка create_order(): {e}")


def check_last_window():
    time.sleep(2)
    window = ""

    all_windows = gw.getAllWindows()
    for window in all_windows:
        print(window.title)
        if "NCALayer" in window.title:
            window = "NCALayer"
            break

    return window


def save_file_in_server(number_order, file_path, template):
    directory = os.path.dirname(template)
    new_file_name = f'{number_order}_заявление.pdf'
    new_file_path = os.path.join(directory, new_file_name)

    os.system(f'copy "{file_path}" "{new_file_path}"')
    get_file_logger(file_name).info(f"Файл {new_file_name} сохранен в директории {directory}")
    os.remove(file_path)


class EgovForm(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Выбор рабочей директории")
        self.setGeometry(100, 100, 400, 200)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.path_input = QLineEdit()
        self.layout.addWidget(self.path_input)

        self.select_button = QPushButton("Выбрать директорию")
        self.select_button.clicked.connect(self.select_directory)
        self.layout.addWidget(self.select_button)

        self.start_button = QPushButton("Запустить программу")
        self.start_button.clicked.connect(self.start_program)
        self.layout.addWidget(self.start_button)

        self.central_widget.setLayout(self.layout)

    def select_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Выбрать директорию")
        if directory:
            self.path_input.setText(directory)

    def start_program(self):
        path = self.path_input.text().replace("/", "\\")
        if path:
            print("Запуск программы с директорией:", fr'{path}')
            while True:
                try:
                    main(fr'{path}')
                    break
                except Exception as e:
                    print("Произошла ошибка:", e)
                    time.sleep(2)
        else:
            print("Не выбрана директория")


app = QApplication(sys.argv)
window = EgovForm()
window.show()
sys.exit(app.exec_())
