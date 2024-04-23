import json
from datetime import timezone, datetime

import pygetwindow as gw
import openpyxl
import pyperclip
import pyautogui
import time
import xml.etree.ElementTree as ET
import sys

import pytz
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from send_request_egov import send_auth_request


created_orders = []


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
    try:
        moveAllWindows()

        # time.sleep(20)
        # start_x, start_y = pyautogui.position()
        # print(start_x, start_y)

        eds_directory_path = r"\\10.10.10.144\Serv-55\Отдел аренды\1.КаР-Тел\ЭЦП КаР-Тел"
        password = "May2021"
        pyperclip.copy(eds_directory_path)
        time.sleep(2)
        choose_eds = (538, 606)
        pyautogui.click(choose_eds)
        time.sleep(2)
        moveAllWindows()

        type_eds_path = (1458, 71)
        pyautogui.click(type_eds_path)

        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(2)

        sing_in_ncalayer(search)

        write_password_points = (356, 654)
        pyautogui.click(write_password_points)
        time.sleep(2)
        pyperclip.copy(password)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)

        moveAllWindows()

        pyautogui.press('enter')
        time.sleep(2)

        sign_certificate_points = (870, 911)
        pyautogui.click(sign_certificate_points)
    except Exception as e:
        raise ValueError(f"Ошибка automate_ncalayer(): {e}")


def main(count_loop, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
         estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
         iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg):

    try:
        driver = webdriver.Chrome()
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

        # await auth_with_eds(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, location_land_plot_arg,
        #                     requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg)
        # driver.quit()

        auth = authorization(driver)
        if auth:
            change_egov_language_to_russian(driver)
            search_by_query_in_portal(driver, "эскиз", choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                                      estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                                      iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg)

    except Exception as e:
        raise ValueError(f"Ошибка при обработке заявки {count_loop}: {e}")


def auth_with_eds(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                  estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                  iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg):
    xml_value = driver.find_element(By.ID, "xmlToSign")
    root = ET.fromstring(xml_value.get_attribute("value"))

    timeTicket = root.find('timeTicket').text
    sessionID = root.find('sessionid').text

    try:
        send_auth_request(sessionID, timeTicket, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                          estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                          iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg)
        return True
    except Exception as e:
        print(e)
        return False


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


def search_by_query_in_portal(driver, query, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                              estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                              iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg):

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
        elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                         estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                         iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg)
    except Exception as e:
        raise ValueError(f"Ошибка search_by_query_in_portal(): {e}")


def elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                     estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                     iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg):
    try:
        driver.get("https://elicense.kz")
        change_elicense_language_to_russian(driver)
        driver.find_element(By.ID, "new-service-button-af-4").click()
        driver.implicitly_wait(100)
        driver.find_elements(By.CLASS_NAME, "new-detail-single")[0].find_elements(By.TAG_NAME, "a")[4].click()
        driver.implicitly_wait(100)
        driver.find_elements(By.TAG_NAME, "li")[1].find_element(By.TAG_NAME, "a").click()
        driver.implicitly_wait(100)
        driver.find_element(By.CLASS_NAME, "new-order-online").click()

        search_text = choice_licensor_arg
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "licensiarSearch"))
        )
        for char in search_text:
            search_box.send_keys(char)
            time.sleep(0.5)

        licensor_section = driver.find_element(By.ID, "treeSection").find_element(By.CLASS_NAME, "tree").find_element(
            By.TAG_NAME, "ul")
        licensors_list = licensor_section.find_elements(By.TAG_NAME, "li")
        licensors = []

        for licensor in licensors_list:
            licensor_style = licensor.get_attribute("style")
            if "display: none" not in licensor_style:
                licensors.append(licensor)

        licensors[0].click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "newRequest"))
        ).click()

        create_order(driver, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                     estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                     iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg)
    except Exception as e:
        raise ValueError(f"Ошибка elicense_new_tab(): {e}")


def create_order(driver, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                 estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                 iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg):
    try:
        time.sleep(3)
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

        driver.implicitly_wait(10)
        fio.send_keys(full_name_representative_arg)
        phone_number.send_keys(phone_number_arg)

        driver.implicitly_wait(10)
        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()

        driver.find_element(By.ID, "toolbar-1016-targetEl").find_elements(By.TAG_NAME, "div")[1].click()  # button "Next"

        time.sleep(10)
        driver.implicitly_wait(100)
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        )
        if table:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
            ).find_elements(By.TAG_NAME, "tbody")[0].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                0].find_element(By.TAG_NAME, "input").send_keys(purpose_use_land_plot_arg)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
            ).find_elements(By.TAG_NAME, "tbody")[1].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                0].find_element(By.TAG_NAME, "input").send_keys(estimated_deminsions_land_plot_arg)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
            ).find_elements(By.TAG_NAME, "tbody")[2].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                0].find_element(By.TAG_NAME, "input").send_keys(location_land_plot_arg)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
            ).find_elements(By.TAG_NAME, "tbody")[3].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                0].find_element(By.TAG_NAME, "input").send_keys(requested_right_use_arg)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
            ).find_elements(By.TAG_NAME, "tbody")[4].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
                0].find_element(By.TAG_NAME, "table").click()
            driver.find_element(By.ID, "boundlist-1062-listEl").find_elements(By.TAG_NAME, "li")[2].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1046-formTable"))
        ).find_element(By.CSS_SELECTOR, "input[type='text']").send_keys(iin_bin_applicant_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1051-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[0].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[
            0].find_element(By.TAG_NAME, "table").click()
        driver.find_element(By.ID, "boundlist-1064-listEl").find_element(By.TAG_NAME, "li").click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[3]/tr/td/div/div[2]/table/tbody[2]/tr/td[1]/table"))
        ).click()
        driver.find_element(By.ID, "boundlist-1066-listEl").find_elements(By.TAG_NAME, "li")[2].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1049-btnIconEl"))
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
        ).send_keys(conclusion_land_commission_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-btns"))
        ).find_elements(By.TAG_NAME, "div")[0].click()

        time.sleep(15)
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

        time.sleep(2)

        # WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[4]/tr/td/div/div[2]/table/tbody[1]/tr/td[1]/table"))
        # ).click()
        # time.sleep(1)
        # driver.implicitly_wait(10)
        #
        # WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.ID, "boundlist-1064-listEl"))
        # ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[1]/div/div[2]/div/div[1]/table/tbody[4]/tr/td/div/div[2]/table/tbody[2]/tr/td[1]/table"))
        ).click()

        time.sleep(2)
        driver.implicitly_wait(10)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "boundlist-1066-listEl"))
        ).find_elements(By.TAG_NAME, "li")[2].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1054-btnIconEl"))
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
        ).send_keys(order_approval_land_plot_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "new-btns"))
        ).find_elements(By.TAG_NAME, "div")[0].click()

        time.sleep(10)
        driver.implicitly_wait(20)

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

        time.sleep(10)
        driver.implicitly_wait(20)

        driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
        ).click()
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[2].click()  # button "Next"

        time.sleep(10)
        driver.implicitly_wait(20)

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

        driver.implicitly_wait(100)

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
        self.setWindowTitle("Egov")
        self.setGeometry(100, 100, 600, 100)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        self.order_template = QLabel("Файл с шаблоном заявки")
        layout.addWidget(self.order_template)

        self.select_file_button = QPushButton("Выбрать файл")
        self.select_file_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_file_button)

        self.submit_button = QPushButton("Отправить")
        self.submit_button.clicked.connect(self.submit_form)
        layout.addWidget(self.submit_button)

        central_widget.setLayout(layout)

    def select_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", "", "All Files (*);;PNG Files (*.png);;JPEG Files (*.jpg);;PDF Files (*.pdf)", options=options)
        if filename:
            self.order_template.setText(filename)

    def get_data_from_selected_file(self, file_path):
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        data = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            choice_licensor, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg, estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg, iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg = row[:10]

            data.append(
                {
                    "choice_licensor": choice_licensor,
                    "full_name_representative": full_name_representative_arg,
                    "phone_number": phone_number_arg,
                    "location_land_plot": location_land_plot_arg,
                    "requested_right_use": requested_right_use_arg,
                    "purpose_use_land_plot": purpose_use_land_plot_arg,
                    "estimated_deminsions_land_plot": estimated_deminsions_land_plot_arg,
                    "iin_bin_applicant": iin_bin_applicant_arg,
                    "conclusion_land_commission": conclusion_land_commission_arg,
                    "order_approval_land_plot": order_approval_land_plot_arg,
                }
            )

        self.start_program(data)

    def submit_form(self):
        order_template: str = self.order_template.text()

        self.get_data_from_selected_file(order_template)

    @staticmethod
    def get_kz_time():
        # Устанавливаем временную зону для Казахстана, например, Астану
        kz_time_zone = pytz.timezone('Asia/Almaty')
        kz_time = datetime.now(kz_time_zone)
        formatted_time = kz_time.strftime('%d-%m-%Y %H:%M:%S')

        return formatted_time

    def start_program(self, data):
        file_path = "created_orders.json"
        try:
            with open(file_path, "r") as file:
                created_orders = json.load(file)
        except FileNotFoundError:
            created_orders = []

        count_loop = 0
        while count_loop < len(data):
            try:
                row = data[count_loop]
                choice_licensor = row["choice_licensor"]
                full_name_representative_arg = row["full_name_representative"]
                phone_number_arg = row["phone_number"]
                location_land_plot_arg = row["location_land_plot"]
                requested_right_use_arg = row["requested_right_use"]
                purpose_use_land_plot_arg = row["purpose_use_land_plot"]
                estimated_deminsions_land_plot_arg = row["estimated_deminsions_land_plot"]
                iin_bin_applicant_arg = row["iin_bin_applicant"]
                conclusion_land_commission_arg = row["conclusion_land_commission"]
                order_approval_land_plot_arg = row["order_approval_land_plot"]

                main(count_loop, choice_licensor, full_name_representative_arg, phone_number_arg, purpose_use_land_plot_arg,
                     estimated_deminsions_land_plot_arg, location_land_plot_arg, requested_right_use_arg,
                     iin_bin_applicant_arg, conclusion_land_commission_arg, order_approval_land_plot_arg)

                try:
                    row["created_time"] = self.get_kz_time()
                except Exception as e:
                    print(f"Ошибка при получении времени: {e}")
                    row["created_time"] = datetime.now(timezone.utc).isoformat()

                created_orders.append(row)
                count_loop += 1
            except Exception as e:
                print(f"Ошибка при обработке заявки {count_loop}: {e}")

        with open(file_path, "w") as file:
            json.dump(created_orders, file, indent=4, ensure_ascii=False)
            print(f"Файл {file_path} сохранен")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EgovForm()
    window.show()
    sys.exit(app.exec_())
