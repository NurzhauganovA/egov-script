import time

import pygetwindow as gw
import pyautogui
import pyperclip

from config_expertise import config_main
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLineEdit

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class ExpertiseDirectory(QMainWindow):
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
                    start(fr'{path}')
                    break
                except Exception as e:
                    print("Произошла ошибка:", e)
                    time.sleep(2)
        else:
            print("Не выбрана директория")


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


def sign_with_eds():
    time.sleep(5)
    password = 'May2021'
    moveAllWindows()
    pyautogui.click(1840, 1040)
    time.sleep(2)
    pyautogui.typewrite(password)
    time.sleep(2)
    pyautogui.hotkey('enter')
    time.sleep(2)
    pyautogui.hotkey('enter')


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


def start(path):
    driver = webdriver.Chrome()
    try:
        driver.get('https://epsd.kz/Home/Main')
        driver.maximize_window()
        driver.find_element(By.CLASS_NAME, 'main-menu').find_elements(By.TAG_NAME, 'div')[0].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'cl-eds-heading'))
        ).find_elements(By.TAG_NAME, 'li')[1].click()

        driver.find_elements(By.CLASS_NAME, 'btn-eds-certificate')[0].click()

        automate_ncalayer('AUTH')

        auth = auth_with_eds(driver)
        if auth:
            main(driver, fr'{path}')
        else:
            print('Ошибка авторизации')
            raise Exception('Ошибка авторизации')
    except Exception as e:
        print("Произошла ошибка:", e)
        raise
    finally:
        driver.quit()


def auth_with_eds(driver):
    try:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.CLASS_NAME, "top-bar-info"))
        )
        return element
    except Exception as e:
        print(e)
        raise


def successCreatedOrder(driver):
    try:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.ID, "multi-derevo"))
        )
        return element
    except Exception as e:
        print(e)
        raise


def main(driver, path):
    try:
        while True:
            context = config_main(fr'{path}')
            print(context)

            modal_window = driver.find_elements(By.ID, "modalConcNotice")

            if modal_window:
                try:
                    modal_window[0].find_element(By.CLASS_NAME, "modal-footer").find_element(By.TAG_NAME, "button").click()
                    time.sleep(2)
                except Exception as e:
                    print("Произошла ошибка:", e)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'btn-group'))
            ).click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-menu'))
            ).find_elements(By.TAG_NAME, 'li')[0].click()

            driver.find_elements(By.CLASS_NAME, 'frame')[0].find_elements(By.TAG_NAME, 'input')[0].click()
            time.sleep(2)

            #  first step - type of document
            type_doc = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'refDesignStages'))
            )
            type_doc.click()
            doc_options = type_doc.find_elements(By.TAG_NAME, 'option')
            for option in doc_options:
                if option.get_attribute('value') == '10196':
                    option.click()
                    break

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'CorrespondentOutcomeNumber'))
            ).send_keys(context['number_of_order'])

            date_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'CorrespondentOutcomeDate'))
            )
            date_input.clear()
            date_input.send_keys(context['date_of_order'])
            time.sleep(5)
            button_next = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'button_next'))
            )
            button_next.click()

            driver.implicitly_wait(100)

            #  second step - type of work
            work_type_list = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'FormStep2'))
            ).find_element(By.TAG_NAME, 'fieldset').find_element(By.TAG_NAME, 'ul').find_elements(By.TAG_NAME, 'li')

            for work_type in work_type_list:
                value = work_type.find_element(By.TAG_NAME, 'label').get_attribute('for')
                if value == '162':
                    work_type.find_element(By.TAG_NAME, 'input').click()
                    break

            button_next.click()

            driver.implicitly_wait(100)

            #  third step - choose type of object
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'o20796'))
            ).find_element(By.TAG_NAME, 'span').click()
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'o20829'))
            ).find_element(By.TAG_NAME, 'span').click()
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'o20829'))
            ).find_element(By.TAG_NAME, 'ul').find_elements(By.TAG_NAME, 'li')[1].find_element(By.TAG_NAME, 'label').click()

            button_next.click()

            driver.implicitly_wait(100)

            #  fourth step - choose of object location

            oblast = context["oblast"]
            region = context["region"]
            sel_okrug = context["sel_okrug"]
            city = context["city"]

            driver.implicitly_wait(100)
            time.sleep(10)

            region_object = driver.find_element(By.ID, "checkpanel").find_element(By.TAG_NAME, "li").find_element(By.TAG_NAME, "ul").find_elements(By.TAG_NAME, "li")
            for region_obj in region_object:
                region_name = region_obj.find_element(By.TAG_NAME, "span").find_element(By.TAG_NAME, "strong").text
                print("region_name:", region_name)
                if oblast and oblast in region_name:
                    region_obj.find_element(By.TAG_NAME, "div").click()
                    time.sleep(7)
                    subregions = region_obj.find_element(By.TAG_NAME, "ul").find_elements(By.TAG_NAME, "li")
                    for subregion in subregions:
                        subregion_name = subregion.find_element(By.TAG_NAME, "label").text
                        print("subregion_name:", subregion_name)
                        if (city and city in subregion_name) or (region and region in subregion_name) or (sel_okrug and sel_okrug in subregion_name):
                            subregion.find_elements(By.TAG_NAME, "input")[0].click()
                            break
                elif city and city in region_name:
                    region_obj.find_element(By.TAG_NAME, "div").click()
                    time.sleep(7)
                    subregions = region_obj.find_element(By.TAG_NAME, "ul").find_elements(By.TAG_NAME, "li")
                    for subregion in subregions:
                        subregion_name = subregion.find_element(By.TAG_NAME, "label").text
                        if (region and region in subregion_name) or (sel_okrug and sel_okrug in subregion_name):
                            subregion.find_elements(By.TAG_NAME, "input")[0].click()
                            break

            button_next.click()

            driver.implicitly_wait(100)

            #  fifth step - questionnaire of the object of examination
            time.sleep(2)
            technological_complexity = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'refTechnologicalComplexity'))
            ).find_elements(By.TAG_NAME, 'option')

            for complexity in technological_complexity:
                if complexity.get_attribute('value') == '2':
                    complexity.click()
                    break

            time.sleep(2)

            driver.find_elements(By.ID, 'levpanel')[0].find_elements(By.TAG_NAME, 'li')[4].find_element(By.TAG_NAME,
                                                                                                        'input').click()

            time.sleep(2)
            driver.implicitly_wait(100)
            level_of_responsibility = driver.find_element(By.ID, 'refLevelOfResponsibility').find_elements(By.TAG_NAME,
                                                                                                           'option')

            for responsibility in level_of_responsibility:
                if responsibility.get_attribute('value') == '7':
                    responsibility.click()
                    break

            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ProjectName_RU'))
            ).send_keys(context['object_name_rus'])

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ProjectName_KZ'))
            ).send_keys(context['object_name_kaz'])

            construction_industry = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'refConstructionIndustry'))
            ).find_elements(By.TAG_NAME, 'option')

            for industry in construction_industry:
                if industry.get_attribute('value') == '72':
                    industry.click()
                    break

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'customerTerm'))
            ).send_keys('980540000397')
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'controlgroup'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'div')[1].find_element(By.TAG_NAME, 'img').click()
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'designTerm'))
            ).send_keys('080840010555')
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="FormStep5"]/div[1]/div[7]/div/div[1]/div[2]/img'))
            ).click()
            time.sleep(2)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContractDate'))
            ).send_keys('05.01.2021')

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ContractNumber'))
            ).send_keys('09682-20')

            source_finance = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'refSourcesOfFinancing'))
            ).find_elements(By.TAG_NAME, 'option')

            for source in source_finance:
                if source.get_attribute('value') == '18':
                    source.click()
                    break

            time.sleep(5)

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            is_potentially_dangerous_object = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'bootstrap-switch-id-IsPotentialDangerous'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'span')[-1]
            is_potentially_dangerous_object.click()
            time.sleep(1)
            is_potentially_dangerous_object.click()

            time.sleep(2)
            is_expo2017_object = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'bootstrap-switch-id-IsExpoObject'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'span')[-1]
            is_expo2017_object.click()
            time.sleep(1)
            is_expo2017_object.click()

            time.sleep(2)
            is_using_tims_object = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'bootstrap-switch-id-IsUsingTimsoObject'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'span')[-1]
            is_using_tims_object.click()
            time.sleep(1)
            is_using_tims_object.click()

            time.sleep(2)
            with_expert_support_object = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'bootstrap-switch-id-WithExpertSupportObject'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'span')[-1]
            with_expert_support_object.click()
            time.sleep(1)
            with_expert_support_object.click()

            time.sleep(2)
            is_special_industry = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'bootstrap-switch-id-IsSpecialIndustry'))
            ).find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'span')[-1]
            is_special_industry.click()
            time.sleep(1)
            is_special_industry.click()

            time.sleep(2)
            driver.find_elements(By.ID, 'levpanel')[1].find_elements(By.TAG_NAME, 'li')[4].find_element(By.TAG_NAME, 'input').click()

            construction_territorial_subdivisions = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'refConstructionTerritorialSubdivisions'))
            ).find_elements(By.TAG_NAME, 'option')

            for subdivision in construction_territorial_subdivisions:
                if subdivision.get_attribute('value') == '565':
                    subdivision.click()
                    break

            time.sleep(2)
            button_next.click()
            time.sleep(2)
            button_next.click()
            time.sleep(2)

            button_ready = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'button_ready'))
            )

            button_ready.click()
            time.sleep(2)
            sign_with_eds()
            success_created = successCreatedOrder(driver)
            if success_created:
                try:
                    documentation_part(driver, fr'{path}')
                except Exception as e:
                    print("Произошла ошибка:", e)
                    raise
            else:
                print('Ошибка создания заявки')
                raise
    except Exception as e:
        print("Произошла ошибка:", e)
        raise


def set_materials(path):
    try:
        context = config_main(fr'{path}')

        files = context.get('files', [])
        project_files = context.get('project_files', [])
        apz_files, contract_files, task_files, igi_files, permission_docs, pir_files, letter_files, topo_files, ep_files, tech_cond, license_files, acc_project, acc_genproject = [], [], [], [], [], [], [], [], [], [], [], [], []
        rp_files, km_oif_files, ovos_files, calculation_files, gp_files, km_files, oif_files, opz_files, pos_files, pp_files, eg_files = [], [], [], [], [], [], [], [], [], [], []

        for file in files:
            for key, value in file.items():
                if 'км-оиф' in key:
                    km_oif_files.append(value)
                if 'апз' in key:
                    apz_files.append(value)
                elif 'договор аренды' in key:
                    contract_files.append(value)
                elif 'задание на проектирование' in key:
                    task_files.append(value)
                elif 'иги' in key:
                    igi_files.append(value)
                elif 'пир' in key:
                    pir_files.append(value)
                elif 'письмо' in key:
                    letter_files.append(value)
                elif 'топо' in key:
                    topo_files.append(value)
                elif 'эп' in key:
                    ep_files.append(value)
                elif 'лицензия_пд' in key:
                    license_files.append(value)
                elif 'прав-док' in key:
                    permission_docs.append(value)
                elif 'тех-условие' in key:
                    tech_cond.append(value)
                elif 'рек-проект' in key:
                    acc_project.append(value)
                elif 'рек-генпроект' in key:
                    acc_genproject.append(value)

        for file in project_files:
            for key, value in file.items():
                if 'овос' in key:
                    ovos_files.append(value)
                elif 'расчет' in key:
                    calculation_files.append(value)
                elif 'гп' in key:
                    gp_files.append(value)
                elif 'км' in key:
                    km_files.append(value)
                elif 'оиф' in key:
                    oif_files.append(value)
                elif 'опз' in key:
                    opz_files.append(value)
                elif 'пос' in key:
                    pos_files.append(value)
                elif 'пп' in key:
                    pp_files.append(value)
                elif 'эг' in key:
                    eg_files.append(value)
                elif 'рп' in key:
                    rp_files.append(value)

        source_materials = {
            '7': [  # done
                {
                    'file': task_files[0] if len(task_files) > 0 else '',
                    'sign': False
                }
            ],
            '157': [  # done
                {
                    'file': license_files[0] if len(license_files) > 0 else '',
                    'sign': False
                }
            ],
            '240': [  # done
                {
                    'file': topo_files[0] if len(topo_files) > 0 else '',
                    'sign': False
                }
            ],
            '241': [  # done
                {
                    'file': igi_files[0] if len(igi_files) > 0 else '',
                    'sign': False
                }
            ],
            '173': [  # done
                {
                    'file': letter_files[0] if len(letter_files) > 0 else '',
                    'sign': False
                }
            ],
            '180': [  # done
                {
                    'file': pir_files[0] if len(pir_files) > 0 else '',
                    'sign': False
                }
            ],
            '192': [  # done
                {
                    'file': apz_files[0] if len(apz_files) > 0 else '',
                    'sign': False
                }
            ],
            '194': [  # done
                {
                    'file': acc_project[0] if len(acc_project) > 0 else '',
                    'sign': False
                }
            ],
            '175': [  # done
                {
                    'file': ep_files[0] if len(ep_files) > 0 else '',
                    'sign': False
                }
            ],
            '243': [  # done
                {
                    'file': permission_docs[0] if len(permission_docs) > 0 else '',
                    'sign': False
                }
            ],
            '189': [  # done
                {
                    'file': acc_genproject[0] if len(acc_genproject) > 0 else '',
                    'sign': False
                }
            ]
        }

        composition_documentation = {
            '225': [  # done
                {
                    'file': km_oif_files[0] if len(km_oif_files) > 0 else km_files[0] if len(km_files) > 0 else oif_files[0] if len(oif_files) > 0 else '',
                    'sign': False
                }
            ],
            '236': [  # done
                {
                    'file': gp_files[0] if len(gp_files) > 0 else '',
                    'sign': False
                }
            ],
            '299': [  # done
                {
                    'file': pp_files[0] if len(pp_files) > 0 else '',
                    'sign': False
                },
            ],
            '300': [  # done
                {
                    'file': opz_files[0] if len(opz_files) > 0 else '',
                    'sign': False
                },
            ],
            '301': [  # done
                {
                    'file': ovos_files[0] if len(ovos_files) > 0 else '',
                    'sign': False
                }
            ],
            '304': [  # done
                {
                    'file': rp_files[0] if len(rp_files) > 0 else eg_files[0] if len(eg_files) > 0 else '',
                    'sign': False
                },
            ],
            '305': [  # done
                {
                    'file': pos_files[0] if len(pos_files) > 0 else '',
                    'sign': False
                },
            ],
            '330': [  # done
                {
                    'file': calculation_files[0] if len(calculation_files) > 0 else '',
                    'sign': False
                }
            ]
        }

        return [source_materials, composition_documentation]
    except Exception as e:
        print("Произошла ошибка:", e)
        raise


def documentation_part(driver, path):
    try:
        driver.implicitly_wait(100)
        list_materials = driver.find_elements(By.ID, 'root')[0].find_elements(By.TAG_NAME, 'li')

        materials = set_materials(fr'{path}')

        for material in list_materials:
            for key, value in materials[0].items():
                if int(key) == int(material.get_attribute('data-id')):
                    for item in value:
                        if len(item['file']) > 0:
                            if item['file'] == '':
                                print('No file')
                                continue
                            material.find_element(By.TAG_NAME, 'span').find_element(By.TAG_NAME, 'a').click()
                            time.sleep(2)
                            file = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, 'file'))
                            )
                            file.send_keys(item['file'])
                            time.sleep(2)
                            if item['sign']:
                                WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((By.ID, 'subscribe-files'))
                                ).click()
                                time.sleep(2)

                                sign_with_eds()
                                time.sleep(5)
                        else:
                            print('No file')

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btCheck'))
        ).click()

        time.sleep(8)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btCheckAll'))
        ).click()

        time.sleep(8)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btMultipleSign'))
        ).click()

        time.sleep(2)
        sign_with_eds()

        time.sleep(5)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'tablink2'))
        ).click()

        driver.implicitly_wait(100)
        list_materials = driver.find_elements(By.ID, 'root')[1].find_elements(By.TAG_NAME, 'li')

        idx = ['']
        for material in list_materials:
            for key, value in materials[1].items():
                if int(key) == int(material.get_attribute('data-id')):
                    for item in value:
                        if len(item['file']) > 0:
                            if item['file'] == '':
                                print('No file')
                                continue
                            material.find_element(By.TAG_NAME, 'span').find_element(By.TAG_NAME, 'a').click()
                            time.sleep(2)
                            tom = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, 'TomNumber'))
                            )
                            tom.send_keys(str(len(idx)))
                            time.sleep(2)

                            album = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, 'BookNumber'))
                            )
                            album.send_keys("1")
                            time.sleep(2)

                            file = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, 'file'))
                            )
                            file.send_keys(item['file'])
                            time.sleep(2)

                            if item['sign']:
                                WebDriverWait(driver, 10).until(
                                    EC.presence_of_element_located((By.ID, 'subscribe-files'))
                                ).click()
                                time.sleep(2)

                                sign_with_eds()

                                time.sleep(5)

                            idx.append('')
                        else:
                            print('No file')

        print(idx)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btCheck'))
        ).click()

        time.sleep(8)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btCheckAll'))
        ).click()

        time.sleep(8)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btMultipleSign'))
        ).click()

        time.sleep(2)

        sign_with_eds()

        time.sleep(3600)
    except Exception as e:
        print("Произошла ошибка:", e)
        raise


app = QApplication(sys.argv)
window = ExpertiseDirectory()
window.show()
sys.exit(app.exec_())
