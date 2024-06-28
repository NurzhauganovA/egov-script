import json
import os
from datetime import timezone, datetime, timedelta

import pygetwindow as gw
import openpyxl
import pyperclip
import pyautogui
import time
import sys

import pytz
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QLineEdit

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config_seventh_part_talon import get_data_values

created_orders = []
absolute_path = os.path.dirname(os.path.abspath(__file__))


def moveAllWindows():
    time.sleep(2)
    window = gw.getAllWindows()[0]
    window.minimize()
    window.restore()
    window.activate()
    window.maximize()
    window.moveTo(0, 0)
    window.resizeTo(1920, 1080)
    time.sleep(2)


def sing_in_ncalayer(search):
    search_certificate_points = (1681, 70)
    pyautogui.click(search_certificate_points)
    time.sleep(2)
    pyperclip.copy(search)
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    AUTH_file_points = (390, 190)
    pyautogui.click(AUTH_file_points)
    pyautogui.press('enter')


def automate_ncalayer(search):
    time.sleep(10)
    # try:
    #     moveAllWindows()
    #
    #     # time.sleep(20)
    #     # start_x, start_y = pyautogui.position()
    #     # print(start_x, start_y)
    #
    #     eds_directory_path = r"\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\ЭЦП КаР-Тел"
    #     password = "May2021"
    #     pyperclip.copy(eds_directory_path)
    #     time.sleep(2)
    #     choose_eds = (538, 606)
    #     pyautogui.click(choose_eds)
    #     time.sleep(2)
    #     moveAllWindows()
    #
    #     type_eds_path = (1458, 71)
    #     pyautogui.click(type_eds_path)
    #
    #     pyautogui.hotkey('ctrl', 'v')
    #     pyautogui.press('enter')
    #     time.sleep(2)
    #
    #     sing_in_ncalayer(search)
    #
    #     write_password_points = (356, 654)
    #     pyautogui.click(write_password_points)
    #     time.sleep(2)
    #     pyperclip.copy(password)
    #     time.sleep(2)
    #     pyautogui.hotkey('ctrl', 'v')
    #     time.sleep(2)
    #     pyautogui.press('enter')
    #     time.sleep(2)
    #
    #     moveAllWindows()
    #
    #     pyautogui.press('enter')
    #     time.sleep(2)
    #
    #     sign_certificate_points = (870, 911)
    #     pyautogui.click(sign_certificate_points)
    # except Exception as e:
    #     raise ValueError(f"Ошибка automate_ncalayer(): {e}")


def main(path):
    context = get_data_values(path)
    choice_licensor_arg = context['choice_licensor']
    full_name_representative_arg = context['full_name_representative']
    phone_number_arg = context['phone_number']
    customer_arg = context['customer']
    projector_arg = context['projector']
    full_object_name_arg = context['full_object_name']
    address_arg = context['address']
    doc_file_arg = context['doc_file']

    try:
        options = Options()
        os.makedirs(os.path.join(absolute_path, "files"), exist_ok=True)
        download_directory = os.path.join(absolute_path, "files")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        driver = webdriver.Chrome(options=options)
        driver.get("https://www.egov.kz")
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
            search_by_query_in_portal(driver, "эскиз", choice_licensor_arg, full_name_representative_arg,
                                      phone_number_arg,
                                      customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg)

    except Exception as e:
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


def search_by_query_in_portal(driver, query, choice_licensor_arg, full_name_representative_arg, phone_number_arg,
                              customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg):
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
        elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg,
                         customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg)
    except Exception as e:
        raise ValueError(f"Ошибка search_by_query_in_portal(): {e}")


def elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg,
                     customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg):
    try:
        driver.get("https://elicense.kz")
        change_elicense_language_to_russian(driver)
        driver.execute_script("window.scrollBy(0, 500);")
        driver.find_element(By.ID, "new-service-button-af-14").click()  # click 'Строительство'
        driver.implicitly_wait(30)
        driver.find_elements(By.CLASS_NAME, "new-detail-single")[2].find_element(By.TAG_NAME,
                                                                                 "a").click()  # click 'Согласование эскиза (эскизного проекта)'
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div/div/ol/li[1]/a"))
        ).click()
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

        driver.implicitly_wait(100)
        # Находим все элементы списка областей
        oblast_list = regions_box.find_elements(By.TAG_NAME, "li")
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

        time.sleep(10)
        driver.implicitly_wait(100)

        driver.implicitly_wait(100)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "newRequest"))
        ).click()
        driver.implicitly_wait(100)

        create_order(driver, full_name_representative_arg, phone_number_arg,
                     customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg)
    except Exception as e:
        raise ValueError(f"Ошибка elicense_new_tab(): {e}")


def navigate_and_click(driver, addresses, tbody_path):
    idx = 1
    for address in addresses:
        print("address:", address)
        tbody_regions = driver.find_element(By.XPATH, tbody_path)
        regions = tbody_regions.find_elements(By.TAG_NAME, "tr")[2:]
        print("regions:", regions)
        found = False

        for region in regions:
            region_object = region.find_element(By.TAG_NAME, "td").find_element(By.TAG_NAME, "div")
            print("region name:", region_object.text)
            if address in region_object.text:
                if address == addresses[-1]:
                    driver.execute_script(
                        "var event = new MouseEvent('dblclick', { bubbles: true, cancelable: true, view: window }); arguments[0].dispatchEvent(event);",
                        region_object.find_elements(By.TAG_NAME, "img")[idx]
                    )
                    time.sleep(5)
                else:
                    region_object.find_elements(By.TAG_NAME, "img")[idx].click()
                    time.sleep(5)
                found = True
                break
        idx += 1

        if not found:
            print(f"Address '{address}' not found.")
            return False
    return True


def element_displayed(driver, locator):
    element = driver.find_element(*locator)
    return element.is_displayed() and element.value_of_css_property('display') != 'none'


def create_order(driver, full_name_representative_arg, phone_number_arg,
                 customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg):
    try:
        time.sleep(3)
        number_orders = driver.find_element(By.ID, "panel-1010-formTable").find_elements(
            By.TAG_NAME, "tbody"
        )[0].find_element(By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td"
        )[0].find_element(By.TAG_NAME, "input").get_attribute('value')
        print(number_orders)
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
        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[
            0].click()  # button "Save"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[
            1].click()  # button "Next"

        time.sleep(10)

        driver.find_element(By.ID, "panel-1041-formTable").find_elements(By.TAG_NAME, "tbody")[0].find_element(
            By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "textarea").send_keys("full_object_name_arg")
        time.sleep(2)

        driver.find_element(By.ID, "panel-1041-formTable").find_elements(By.TAG_NAME, "tbody")[1].find_element(
            By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "table").click()
        time.sleep(1)
        driver.find_element(By.ID, "boundlist-1048-listEl").find_elements(By.TAG_NAME, "li")[11].click()
        time.sleep(2)

        # Нажимаем на кнопку "Добавить"
        driver.find_element(By.ID, "toolbar-1015-targetEl").find_elements(By.TAG_NAME, "div")[0].click()
        time.sleep(2)

        list_addresses = ["Актюбинская область", "Шалкар", "Бозой", "Коянкулак"]

        # Нажимаем на кнопку "..." (для ввода адреса)
        detail = driver.find_element(By.XPATH,
                                     "/html/body/div[13]/div[2]/div/div/table/tbody[2]/tr/td/div/div/div/div/div/em/button/span[1]")
        if not detail:
            detail = driver.find_element(By.XPATH,
                                         "/html/body/div[12]/div[2]/div/div/table/tbody[2]/tr/td/div/div/div/div/div/em/button/span[1]")

        detail.click()
        time.sleep(2)

        # Запуск функции с вашим списком адресов
        navigate_and_click(driver, list_addresses, "/html/body/div[15]/div[2]/div[1]/div[2]/div/table/tbody")
        time.sleep(2)

        # Ожидаем появления следующего поля для ввода и вводим значение
        street_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[13]/div[2]/div/div/table/tbody[3]/tr/td[1]/input"))
        )
        if not street_input:
            street_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "/html/body/div[12]/div[2]/div/div/table/tbody[3]/tr/td[1]/input"))
            )

        driver.execute_script("arguments[0].value = arguments[1];", street_input, "object_address_street_arg")
        print("value attribute:", street_input.get_attribute("value"))
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", street_input)

        # Нажимаем на кнопку "Сохранить"
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1054-btnInnerEl"))
        ).click()
        time.sleep(5)

        driver.find_element(By.ID, "panel-1042-formTable").find_elements(By.TAG_NAME, "tbody")[0].find_element(
            By.TAG_NAME, "tr"
        ).find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input").send_keys(
            "full_name_main_building_inspector_arg")

        end_date = (datetime.now(tz=pytz.timezone('Asia/Almaty')) + timedelta(days=15)).strftime("%d.%m.%Y")
        print("end_date:", end_date)

        input_field = driver.find_element(By.XPATH,
                                          "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[4]/tr/td/div/div[2]/table/tbody[3]/tr/td[1]/table/tbody/tr/td[1]/input")

        driver.execute_script("arguments[0].value = arguments[1];", input_field, end_date)
        print("value attribute:", input_field.get_attribute("value"))
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", input_field)

        time.sleep(10)
        driver.implicitly_wait(100)
        driver.find_element(By.ID, "panel-1011-formTable").find_elements(By.TAG_NAME, "tbody")[0].find_element(
            By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "table").click()
        time.sleep(2)
        driver.find_element(By.ID, "boundlist-1062-listEl").find_elements(By.TAG_NAME, "li")[1].click()

        tbody_objects = driver.find_element(By.ID, "panel-1011-formTable").find_elements(By.TAG_NAME, "tbody")
        result = []

        for obj in tbody_objects:
            style = obj.get_attribute("style")
            if "table-layout: fixed;" in style:
                result.append(obj)

        if len(result) > 5:
            inputs = [
                ("customer_organization_name_arg", 1),
                ("030004", 2),
                ("customer_city_arg", 3),
                ("customer_locality_arg", 4),
                ("customer_street_arg", 5),
                (int("2"), 6),
                ("customer_phone_number_arg", 7),
                ("980540000397", 9)
            ]

            for value, index in inputs:
                input_element = result[index].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                    0].find_element(By.TAG_NAME, "input")

                # Очистка поля ввода перед отправкой текста
                input_element.clear()

                input_element.send_keys(value)

                # Добавление небольшой задержки
                time.sleep(2)

        else:
            print("Ошибка в заполнении полей")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1044-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[0].click()  # button "Сохранить"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1044-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[2].click()  # button "Дальше"

        time.sleep(10)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1013-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[0].click()  # Добавить
        time.sleep(5)

        add_new_order = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "form-1075-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")

        add_new_order[0].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(
            By.TAG_NAME, "input").send_keys("name_local_executive_body_arg")
        add_new_order[1].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(
            By.TAG_NAME, "input").send_keys("number_decision_arg")
        add_new_order[2].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(
            By.TAG_NAME, "input").send_keys("29.05.2024")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1078-btnInnerEl"))
        ).click()  # button "Сохранить"
        time.sleep(5)

        info_project_doc = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1036-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")

        inputs = (
            ("name_project_org_developed_doc_arg", 0),
            ("number_license_arg", 1),
            ("25.05.2024", 2),
            ("", 3),
        )

        for value, index in inputs:
            try:
                if index == 3:
                    clickable_element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, f"tbody:nth-child({index + 1}) tr td:nth-child(1) table"))
                    )
                    clickable_element.click()

                    dropdown_list = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "boundlist-1080-listEl"))
                    )
                    dropdown_items = dropdown_list.find_elements(By.TAG_NAME, "li")

                    if dropdown_items:
                        dropdown_items[1].click()
                    else:
                        print(f"No items found in dropdown for index {index}")
                else:
                    input_element = info_project_doc[index].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                        0].find_element(By.TAG_NAME, "input")

                    input_element.send_keys(value)

                time.sleep(2)
            except Exception as e:
                print(f"Exception at index {index}: {e}")
                break

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[2]/tr/td/div/div[2]/table/tbody[5]/tr/td[1]/input"))
        ).send_keys("name_project_org_approved_doc_arg")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[2]/tr/td/div/div[2]/table/tbody[6]/tr/td[1]/input"))
        ).send_keys("number_order_approval_arg")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[2]/tr/td/div/div[2]/table/tbody[7]/tr/td[1]/table/tbody/tr/td[1]/input"))
        ).send_keys("29.05.2024")

        positive_expert_opinion = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1067-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")

        # Найти поле ввода даты
        date_input = positive_expert_opinion[1].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(
            By.TAG_NAME, "table").find_element(By.TAG_NAME, "tbody").find_element(By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input")

        # Установить значение атрибута value с помощью JavaScript
        date_value = "16.04.2024"
        driver.execute_script("arguments[0].setAttribute('value', arguments[1])", date_input, date_value)
        time.sleep(2)
        date_input.send_keys(date_value)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[3]/tr/td/div/div[2]/table/tbody[3]/tr/td/div/div/div/div/table/tbody/tr/td[1]/input"))
        ).send_keys("Aқ-0140/24")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1037-btnInnerEl"))
        ).click()  # button "Отправить"

        time.sleep(10)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1041-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[0].click()  # button "Добавить"
        time.sleep(5)

        add_new_order = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "form-1084-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")

        inputs = (
            ("name_expert_performed_expertise_arg", 0),
            ("phone_number_expert_arg", 1),
            ("number_certificate_expert_arg", 2),
            ("name_org_have_certification_expert_arg", 4)
        )

        for value, index in inputs:
            try:
                input_element = add_new_order[index].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                    0].find_element(By.TAG_NAME, "input")

                input_element.send_keys(value)

                time.sleep(2)
            except Exception as e:
                print(f"Exception at index {index}: {e}")
                break

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1083-btnInnerEl"))
        ).click()  # button "..." choose city imprisonment

        if driver.find_elements(By.XPATH, "/html/body/div[18]/div[2]/div[1]/div[2]/div/table/tbody"):
            tbody_object = "/html/body/div[18]/div[2]/div[1]/div[2]/div/table/tbody"
        else:
            tbody_object = "/html/body/div[17]/div[2]/div[1]/div[2]/div/table/tbody"

        navigate_and_click(driver, list_addresses, tbody_object)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"/html/body/div[15]/div[2]/div/div/table/tbody[7]/tr/td[1]/input"))
        ).send_keys("name_street_imprisonment_arg")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"/html/body/div[15]/div[2]/div/div/table/tbody[8]/tr/td[1]/input"))
        ).send_keys("number_house_imprisonment_arg")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1087-btnInnerEl"))
        ).click()  # button "Сохранить"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1068-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[0].find_element(By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "table").click()
        time.sleep(1)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "boundlist-1095-listEl"))
        ).find_elements(By.TAG_NAME, "li")[0].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[4]/tr/td/div/div[2]/table/tbody[2]/tr/td[1]/input"))
        ).send_keys(12)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1070-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[0].click()  # button "Сохранить"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))
        ).click()  # button "OK"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "toolbar-1070-targetEl"))
        ).find_elements(By.TAG_NAME, "div")[2].click()  # button "Далее"

        time.sleep(2000)

        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "RequesterSatisfactionWindow"))
            ).find_elements(By.TAG_NAME, "input")[0].click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "RequesterSatisfactionWindow"))
            ).find_elements(By.TAG_NAME, "input")[4].click()
            automate_ncalayer("GOST")
            time.sleep(10)
            driver.quit()
        except Exception as e:
            print(e)
    except Exception as e:
        raise ValueError(f"Ошибка create_order(): {e}")


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
