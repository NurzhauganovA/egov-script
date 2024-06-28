import json
import os
from datetime import timezone, datetime

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

from config_sixth_part import get_data_values

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


def sing_in_ncalayer(search):
    search_certificate_points = (1574, 105)
    pyautogui.click(search_certificate_points)
    time.sleep(1)
    pyperclip.copy(search)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    AUTH_file_points = (711, 276)
    pyautogui.click(AUTH_file_points)
    pyautogui.press('enter')


def automate_ncalayer(search):
    # time.sleep(10)
    try:
        moveAllWindows()
        eds_directory_path = r"\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\ЭЦП КаР-Тел"
        password = "May2021"
        pyperclip.copy(eds_directory_path)
        time.sleep(1)
        choose_eds = (710, 622)
        pyautogui.click(choose_eds)
        time.sleep(1)
        moveAllWindows()
        type_eds_path = (1220, 100)
        pyautogui.click(type_eds_path)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(1)
        sing_in_ncalayer(search)
        write_password_points = (498, 697)
        pyautogui.click(write_password_points)
        time.sleep(1)
        pyperclip.copy(password)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        moveAllWindows()
        pyautogui.press('enter')
        time.sleep(1)
        sign_certificate_points = (843, 1001)
        pyautogui.click(sign_certificate_points)
    except Exception as e:
        raise ValueError(f"Ошибка automate_ncalayer(): {e}")


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
            search_by_query_in_portal(driver, "эскиз", choice_licensor_arg, full_name_representative_arg, phone_number_arg,
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
        # TODO: Осы жерден бастап `Строительство` деген бағананы таңдап, әрі қарай жалғастырып код жазу керек. ФИО, және т.б. данные толтыратын жерге дейін жазу керек. Дальше `create_order()` метод жалғастырады.
        driver.execute_script("window.scrollBy(0, 500);")
        driver.find_element(By.ID, "new-service-button-af-14").click()  # click 'Строительство'
        driver.implicitly_wait(100)
        driver.find_elements(By.CLASS_NAME, "new-detail-single")[0].find_elements(By.TAG_NAME, "a")[8].click()  # click 'Согласование эскиза (эскизного проекта)'
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

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "newRequest"))
        ).click()
        driver.implicitly_wait(10)

        create_order(driver, full_name_representative_arg, phone_number_arg,
                     customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg)
    except Exception as e:
        raise ValueError(f"Ошибка elicense_new_tab(): {e}")


def create_order(driver, full_name_representative_arg, phone_number_arg,
                 customer_arg, projector_arg, full_object_name_arg, address_arg, doc_file_arg):
    try:
        driver.implicitly_wait(100)

        number_orders = driver.find_element(By.ID, "panel-1010-formTable").find_elements(
            By.TAG_NAME, "tbody"
        )[0].find_element(By.TAG_NAME, "tr").find_elements(
            By.TAG_NAME, "td"
        )[0].find_element(By.TAG_NAME, "input").get_attribute('value')
        print(number_orders)

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

        time.sleep(30)
        driver.implicitly_wait(100)

        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        )
        if table:
            table.find_elements(By.TAG_NAME, "tbody")[0].find_elements(
                By.TAG_NAME, "tr"
            )[0].find_elements(
                By.TAG_NAME, "td"
            )[0].find_element(By.TAG_NAME, "input").send_keys(customer_arg)  # Заказчик

            table.find_elements(By.TAG_NAME, "tbody")[1].find_elements(
                By.TAG_NAME, "tr"
            )[0].find_elements(
                By.TAG_NAME, "td"
            )[0].find_element(By.TAG_NAME, "input").send_keys(projector_arg)  # Проектировщик № ГСЛ, категория

            table.find_elements(By.TAG_NAME, "tbody")[2].find_elements(
                By.TAG_NAME, "tr"
            )[0].find_elements(
                By.TAG_NAME, "td"
            )[0].find_element(By.TAG_NAME, "input").send_keys(full_object_name_arg)  # Наименование проектируемого объекта

            table.find_elements(By.TAG_NAME, "tbody")[3].find_elements(
                By.TAG_NAME, "tr"
            )[0].find_elements(
                By.TAG_NAME, "td"
            )[0].find_element(By.TAG_NAME, "input").send_keys(address_arg)  # Адрес проектируемого объекта

        time.sleep(10)
        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"

        driver.implicitly_wait(100)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[2].click()  # button "Next"

        time.sleep(10)
        driver.implicitly_wait(100)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1020-btnIconEl"))
        ).click()

        time.sleep(3)
        driver.implicitly_wait(100)

        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "iframe"))
        )
        driver.switch_to.frame(iframe)

        time.sleep(5)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-lk-wrapper"))
        ).find_element(By.TAG_NAME, "div").find_elements(By.TAG_NAME, "button")[1].click()

        time.sleep(5)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "file"))
        ).send_keys(doc_file_arg)

        time.sleep(10)
        driver.implicitly_wait(100)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-btns"))
        ).find_elements(By.TAG_NAME, "div")[0].click()

        time.sleep(10)
        driver.implicitly_wait(100)

        try:
            documents = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DocumentTable"))
            ).find_element(By.TAG_NAME, "tbody")

            documents.find_elements(By.TAG_NAME, "tr")[1].find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME,
                                                                                                           "a").click()
        except Exception as e:
            print(e)

        time.sleep(5)
        driver.switch_to.default_content()

        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1014-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"

        driver.implicitly_wait(100)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.implicitly_wait(100)
        driver.find_element(By.ID, "toolbar-1014-targetEl").find_elements(By.TAG_NAME, "div")[2].click()  # button "Next"
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

        time.sleep(15)

        while check_last_window() != "NCALayer":
            time.sleep(2)

        automate_ncalayer("GOST")
        time.sleep(20)

        close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

        if close_table:
            close_table[0].click()

        # after confirmation use EDS
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "LkBox"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "myDropdownMainLK"))
        ).find_elements(By.TAG_NAME, 'a')[2].click()

        driver.implicitly_wait(100)
        time.sleep(2)

        close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

        if close_table:
            close_table[0].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'GlobalNumberStr'))
        ).send_keys(number_orders)
        time.sleep(2)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'submit'))
        ).click()
        driver.implicitly_wait(100)

        close_table = driver.find_elements(By.CLASS_NAME, "ui-dialog-titlebar-close")

        if close_table:
            close_table[0].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[5]/div[2]/div[3]/div/div/table/tbody/tr[2]/td[4]/div/a/img'))
        ).click()
        driver.implicitly_wait(100)
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

        time.sleep(15)
        file_path = os.path.join(absolute_path, 'files', number_orders + '_ru.pdf')

        save_file_in_server(number_orders, file_path, doc_file_arg)
        time.sleep(15)

        driver.quit()
    except Exception as e:
        raise ValueError(f"Ошибка create_order(): {e}")


def check_last_window():
    time.sleep(2)
    window = ""

    all_windows = gw.getAllWindows()
    for window in all_windows:
        if "NCALayer" in window.title:
            window = "NCALayer"
            break

    return window


def save_file_in_server(number_order, file_path, template):
    directory = os.path.dirname(template)
    new_file_name = f'{number_order}_заявление.pdf'
    new_file_path = os.path.join(directory, new_file_name)

    os.system(f'copy "{file_path}" "{new_file_path}"')


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
