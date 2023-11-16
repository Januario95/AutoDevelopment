# -*- coding: utf-8 -*-
"""
Created on Tue Jul 16 11:00:42 2023

@author: C835910
"""

import pandas as pd
import win32com.client
import datetime
import os

const = win32com.client.constants
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#from webdriver_manager.chrome import ChromeDriverManager
from difflib import SequenceMatcher

import logger_config
from unidecode import unidecode
import requests

MAIN_PATH = "C://Users/a248433/Documents/SB Mozambique/EDO/Robotics/Bots/Automatização do Processo Abertura de contas Corporate"
T24_URL = 'http://10.245.207.50:8080/BrowserWeb/'
FILE_NAME = 'C://Users/C835910/Documents/Corporate_Account_Opening/temp/accounts.xlsx'

logger = logger_config.get_logger('Corporate Account Opening')
today = datetime.datetime.now()

def init_t24(is_inputter, is_menu):
    logger.info('Iniciando o T24...')
    driver = webdriver.Chrome(f'{MAIN_PATH}/chromedriver.exe')
    driver.implicitly_wait(45)
    driver.get(T24_URL)
    driver.maximize_window()
    if is_inputter:
        driver.find_element(By.NAME, 'signOnName').send_keys('A221320')
        driver.find_element(By.NAME, 'password').send_keys('AAbb123')
    else:
        driver.find_element(By.NAME, 'signOnName').send_keys('IVOFAQ')
        driver.find_element(By.NAME, 'password').send_keys('@AAbb123')
    driver.find_element(By.CLASS_NAME, 'sign_in').click()
    if is_menu:
        driver.switch_to.frame(driver.find_element(By.XPATH, "//*[contains(@name, 'menu')]"))
        logger.info('T24 iniciado com sucesso!')
    else:
        driver.switch_to.frame(driver.find_element(By.XPATH, "//*[contains(@name, 'banner')]"))
        driver.find_element(By.XPATH, "//a[contains(text(), 'Sign Off')]").click()
        logger.info('Logout completed.')
        driver.quit() # Close T24
    return driver

def authorize_customer(driver, customer_id):
    logger.info('Autorizando o cliente {}'.format(customer_id))
    sleep(2)
    driver.find_elements(By.XPATH, "//span[contains(text(), 'Menu Origination Maintenance Team Leader')]")[0].click()
    driver.find_elements(By.XPATH, "//span[contains(text(), 'Cliente')]")[0].click()
    driver.find_elements(By.XPATH, "//a[contains(text(), 'Abertura Cliente em Nome Individual')]")[0].click()

    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    driver.maximize_window()
    element_id = driver.find_element(By.CLASS_NAME, 'idbox_CUSTOMER_INDIVID')
    element_id.clear()
    element_id.send_keys(customer_id)
    driver.find_element(By.XPATH, "//img[contains(@title, 'Perform an action on the contract')]").click()
    driver.find_element(By.XPATH, "//img[contains(@title, 'Authorises a deal')]").click()

    try:
        message = driver.find_elements(By.XPATH, "//td[contains(@class, 'message')]")[0].text
        logger.info('Cliente autorizado.')
        logger.info(message)
    except: message = ''

    driver.close() # Close T24

def create_customer(driver, data):
    logger.info('Iniciando a criação do cliente: {}'.format(data.iloc[3][1]))
    sleep(2)
    driver.find_elements(By.XPATH, "//span[contains(text(), 'Menu Account Origination Officer')]")[0].click()
    driver.find_elements(By.XPATH, "//span[contains(text(), 'Cliente')]")[0].click()
    driver.find_elements(By.XPATH, "//a[contains(text(), 'Abertura Cliente em Nome Individual')]")[0].click()
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    driver.maximize_window()
    driver.find_element(By.XPATH, "//img[contains(@title, 'New Deal')]").click()

    # Informacao dos Documentos de Identificacao
    full_name = str(unidecode(data.iloc[2][1])).split(' ')
    last_name = full_name[len(full_name) - 1]
    if len(last_name) < 3 and len(full_name) > 3: last_name = full_name[len(full_name) - 2]
    #print(str(data.iloc[0][8]))
    type_in(driver, unidecode(data.iloc[2][1]), "fieldName:NAME.1:1")
    type_in(driver, full_name[0] + ' ' + last_name, "fieldName:SHORT.NAME:1")

    if len(str(data.iloc[6][1])) == 9 or 'AB' in str(data.iloc[6][1]): type_in(driver, "1", "fieldName:ID.TYPE")
    else: type_in(driver, "2", "fieldName:ID.TYPE")

    type_in(driver, str(data.iloc[7][1]), "fieldName:LEGAL.ID:1")
    type_in(driver, data.iloc[8][1], "fieldName:LEGAL.ISS.DATE:1")
    type_in(driver, data.iloc[9][1], "fieldName:LEGAL.EXP.DATE:1")
    type_in(driver, data.iloc[10][1], "fieldName:LEGAL.HOLDER.NAME:1")
    type_in(driver, data.iloc[12][1], "fieldName:PLACE.BIRTH")
    type_in(driver, data.iloc[13][1], "fieldName:BIRTH.INCORP.DATE")
    if 'MALE' in data.iloc[14][1]: driver.find_elements(By.ID, 'radio:tab1:GENDER')[1].click()
    else: driver.find_elements(By.ID, 'radio:tab1:GENDER')[0].click()
    type_in(driver, "3", "fieldName:LEGAL.STATUS")
    type_in(driver, unidecode(data.iloc[16][1]), "fieldName:FATHERS.NAME")
    type_in(driver, unidecode(data.iloc[17][1]), "fieldName:MOTHERS.NAME")

    marital = "Unmarried" #str(data[18][1]).strip()

    try: driver.find_element(By.XPATH, "//select[@name='fieldName:MARITAL.STATUS']/option[text()='"+ marital +"']").click()
    except: driver.find_element(By.XPATH, "//select[@name='fieldName:MARITAL.STATUS']/option[text()='de outros']").click()

    type_in(driver, "9", "fieldName:TYPE.MARRIAGE")
    if 'Casad' in str(data.iloc[18][1]): type_in(driver, "Desconhecido", "fieldName:SPOUSE.NAME")

    # Informacao Fiscal
    type_in(driver, data.iloc[20][1], "fieldName:TAX.ID:1")
    driver.find_element(By.NAME, 'radio:tab1:EMIGRANT').click()
    # Informacao Enderecos do Cliente
    type_in(driver, unidecode(short_address(data.iloc[24][1])), "fieldName:STREET:1")
    type_in(driver, "MZ", "fieldName:POST.ADD.CTY")
    type_in(driver, data.iloc[30][1], "fieldName:PROVINCIAS")
    type_in(driver, data.iloc[31][1], "fieldName:LOCALIDADE")
    type_in(driver, unidecode(data.iloc[28][1]), "fieldName:TOWN.COUNTRY:1")
    type_in(driver, unidecode(data.iloc[34][1]), "fieldName:DISTRICT.2")
    type_in(driver, "4", "fieldName:STAT.FREQ")
    type_in(driver, data.iloc[27][1], "fieldName:RESIDENCE.SINCE:1")
    try: type_in(driver, data.iloc[38][1], "fieldName:EMAIL.1:1")
    except: pass
    type_in(driver, data.iloc[37][1], "fieldName:SMS.1:1")
    # Informacao Professional
    type_in(driver, str(data.iloc[41][1]).strip(), "fieldName:JOB.TITLE:1")
    type_in(driver, unidecode(data.iloc[42][1]), "fieldName:EMPLOYERS.NAME:1")
    type_in(driver, "0000", "fieldName:EMPLOYERS.ADD:1:1")
    type_in(driver, "MZN", "fieldName:CUSTOMER.CURRENCY:1")

    salary = str(data.iloc[45][1]).strip()
    type_in(driver, salary, "fieldName:SALARY:1")
    type_in(driver, data.iloc[48][1], "fieldName:EMPLOYMENT.START:1")
    type_in(driver, unidecode(data.iloc[47][1]), "fieldName:OCCUPATION:1")

    #contract_type = 'Outro'
    """if 'PERMANENT' in str(data.iloc[2][7]).strip(): contract_type = 'Contracto por tempo Indeterminado'
    elif 'TEMPORARY' in str(data.iloc[2][7]).strip(): contract_type = 'Contracto por tempo determinado'
    elif 'INDETERMINATED' in str(data.iloc[2][7]).strip(): contract_type = 'Contracto por tempo Indeterminado'"""

    #try: driver.find_element(By.XPATH, "//select[@name='fieldName:EMP.CONT.TYPE']/option[text()='"+ str(data.iloc[49][1]) +"']").click()
    #except: pass

    # Informacao de Gestao
    type_in(driver, "116", "fieldName:ACCOUNT.OFFICER")
    type_in(driver, "2222", "fieldName:TARGET")
    # Final
    now = datetime.datetime.now()
    if len(str(now.month)) == 2: month = now.month
    else: month = f"0{now.month}"
    if len(str(now.day)) == 2: day = now.day
    else: day = f"0{now.day}"

    try:
        driver.find_elements(By.NAME, 'radio:tab1:CLIENTE.PEP')[0].click()
        if data.iloc[61][1] == 'C': driver.find_elements(By.NAME, 'radio:tab1:CUSSICCODE')[2].click()
        if data.iloc[61][1] == 'B': driver.find_elements(By.NAME, 'radio:tab1:CUSSICCODE')[1].click()
        if data.iloc[61][1] == 'A': driver.find_elements(By.NAME, 'radio:tab1:CUSSICCODE')[0].click()
    except Exception as err: print(str(err))

    if data.iloc[61][1] == 'C': driver.find_elements(By.NAME, 'radio:tab1:KYC.FLAG')[2].click()
    if data.iloc[61][1] == 'B': driver.find_elements(By.NAME, 'radio:tab1:KYC.FLAG')[1].click()
    if data.iloc[61][1] == 'A': driver.find_elements(By.NAME, 'radio:tab1:KYC.FLAG')[0].click()

    driver.find_elements(By.ID, 'radio:tab1:KYC')[0].click()
    driver.find_elements(By.ID, 'radio:tab1:INIB.CHQ')[0].click()
    type_in(driver, f"{now.year}{month}{day}", "fieldName:KYC.DATE")
    # Validate
    type_in(driver, "MZ", "fieldName:NATIONALITY")
    sleep(10)
    driver.find_element(By.XPATH, "//img[contains(@title, 'Validate a deal')]").click()
 
    """try:
        errors = driver.find_element(By.NAME, 'errors').text
        if len(errors) > 5:
            logger.info('Erro encontrado durante o processo...')
            #send_email(data.iloc[0][0], errors)
            logger.info(errors)
    except: errors = ''"""

    customer_id = ''
    if False: pass
        #len(errors) > 2: status = 'completed'
    else:
        customer_id = driver.find_element(By.CLASS_NAME, 'iddisplay_CUSTOMER').text
        driver.find_element(By.XPATH, "//img[contains(@title, 'Commit the deal')]").click()
        logger.info('Cliente criado com o CIF: {}'.format(customer_id))
        #status = 'completed'

    driver.close() # Close T24
    return customer_id

def short_address(address):
    return str(address).replace('Esquerdo', 'Esq').replace('Predio ', 'Prd').replace('BAIRRO ', 'B. ').replace('CASA ' , 'C.')

def type_in(driver, text, element):
    input_ = driver.find_element(By.NAME, element)
    input_.clear()
    input_.send_keys(text)