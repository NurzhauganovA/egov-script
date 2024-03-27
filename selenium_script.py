import sys
import time
import xml.etree.ElementTree as ET

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from send_request_egov import send_auth_request


def main():
    choice_licensor_arg = sys.argv[1]
    full_name_representative_arg = sys.argv[2]
    phone_number_arg = sys.argv[3]
    location_land_plot_arg = sys.argv[4]
    requested_right_use_arg = sys.argv[5]
    area_arg = sys.argv[6]
    purpose_use_land_plot_arg = sys.argv[7]
    schema_plan_land_plot_arg = sys.argv[8]

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

    # auth_with_eds(driver)

    # Авторизация через ЭЦП вручную
    auth = authorization(driver)
    if auth:
        change_egov_language_to_russian(driver)
        search_by_query_in_portal(driver, "эскиз", choice_licensor_arg, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg)


def auth_with_eds(driver):
    xml_value = driver.find_element(By.ID, "xmlToSign")
    root = ET.fromstring(xml_value.get_attribute("value"))

    timeTicket = root.find('timeTicket').text
    sessionID = root.find('sessionid').text

    send_auth_request(sessionID, timeTicket)


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


def search_by_query_in_portal(driver, query, choice_licensor_arg, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg):
    search = driver.find_element(By.ID, "edit-query")
    search.send_keys(query, Keys.ENTER)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "search-results"))
    )
    search_results = driver.find_element(By.CLASS_NAME, "search-results")
    search_results.find_elements(By.TAG_NAME, "div")[0].find_element(By.TAG_NAME, "p").find_element(By.TAG_NAME, "a").click()
    driver.find_element(By.ID, "sticky-wrapper").find_element(By.TAG_NAME, "div").find_elements(By.TAG_NAME, "a")[0].click()
    driver.implicitly_wait(5)
    driver.switch_to.window(driver.window_handles[1])
    elicense_new_tab(driver, choice_licensor_arg, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg)


def elicense_new_tab(driver, choice_licensor, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg):
    driver.get("https://elicense.kz")
    change_elicense_language_to_russian(driver)
    driver.find_element(By.ID, "new-service-button-af-4").click()
    driver.implicitly_wait(300)
    driver.find_elements(By.CLASS_NAME, "new-detail-single")[0].find_elements(By.TAG_NAME, "a")[4].click()
    driver.implicitly_wait(300)
    driver.find_elements(By.TAG_NAME, "li")[0].find_element(By.TAG_NAME, "a").click()
    driver.implicitly_wait(300)
    driver.find_element(By.CLASS_NAME, "new-order-online").click()

    search_text = choice_licensor
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

    create_order(driver, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg)


def create_order(driver, full_name_representative_arg, phone_number_arg, location_land_plot_arg, requested_right_use_arg, area_arg, purpose_use_land_plot_arg, schema_plan_land_plot_arg):
    fio = driver.find_element(By.ID, "panel-1011-formTable").find_elements(By.TAG_NAME, "tbody")[1].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input")
    phone_number = driver.find_element(By.ID, "panel-1014-formTable").find_elements(By.TAG_NAME, "tbody")[-2].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input")
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
        ).find_elements(By.TAG_NAME, "tbody")[0].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input").send_keys(location_land_plot_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[1].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input").send_keys(requested_right_use_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[2].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input").send_keys(area_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[3].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "input").send_keys(purpose_use_land_plot_arg)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "panel-1010-formTable"))
        ).find_elements(By.TAG_NAME, "tbody")[4].find_element(By.TAG_NAME, "tr").find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "table").click()
        driver.find_element(By.ID, "boundlist-1015-listEl").find_elements(By.TAG_NAME, "li")[1].click()

    time.sleep(10)
    driver.implicitly_wait(20)
    driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
    ).click()
    driver.implicitly_wait(10)
    driver.find_element(By.ID, "toolbar-1012-targetEl").find_elements(By.TAG_NAME, "div")[2].click()  # button "Next"

    time.sleep(10)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "button-1020-btnIconEl"))
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

    time.sleep(10)
    driver.find_element(By.ID, "file").send_keys(schema_plan_land_plot_arg)
    driver.find_element(By.CLASS_NAME, "new-btns").find_elements(By.TAG_NAME, "div")[0].click()

    documents = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "DocumentTable"))
    ).find_element(By.TAG_NAME, "tbody")

    documents.find_elements(By.TAG_NAME, "tr")[1].find_elements(By.TAG_NAME, "td")[0].find_element(By.TAG_NAME, "a").click()
    driver.switch_to.default_content()

    driver.find_element(By.ID, "toolbar-1014-targetEl").find_elements(By.TAG_NAME, "div")[0].click()  # button "Save"
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "button-1005-btnInnerEl"))  # button "OK"
    ).click()
    driver.implicitly_wait(10)
    driver.find_element(By.ID, "toolbar-1014-targetEl").find_elements(By.TAG_NAME, "div")[2].click()  # button "Next"
    time.sleep(10)


main()
