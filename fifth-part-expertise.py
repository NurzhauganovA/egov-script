# import json
# from datetime import timezone, datetime
#
# import openpyxl
# import pyperclip
# import pyautogui
import time
# import sys
# import pytz
# from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog
from fourth_part_apz import automate_ncalayer

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def authorization():
    driver = webdriver.Chrome()
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
        main(driver)
    else:
        print('Ошибка авторизации')


def auth_with_eds(driver):
    try:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.CLASS_NAME, "top-bar-info"))
        )
        return element
    except Exception as e:
        print(e)


def main(driver):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'btn-group'))
    ).click()
    time.sleep(5)
