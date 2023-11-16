#!/usr/bin/env python3
# -*- encoding: utf-8 -*-
"""
Created on Mon May 08, 2023  10:23:41

@author: a248433
"""

import os
import sys
import glob
import json
import time
import string
import random
import shutil
import getpass
import logging
import keyring
import datetime 
import traceback
import pandas as pd
import datetime as dt
from faker import Faker
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

from _helpers import (
    login_user, logout_user, get_credentials, get_driver, 
    display_text_std, 
    find_element_by_xpath_and_click,
    common_handler, handle_user_auth, change_branch,
    change_to_last_window, find_input_field_and_fill_value,
    INDEXES, ClientesSociedadePrenchido,
    ContasSociedadeConstituidaPr, ClientCIFs,
)
from generate_letter_ import (
    LetterGeneration,
)


f = open("config.txt")
configInfo = json.loads(f.read())

input_username = configInfo["input_username"]
input_key_name = configInfo["input_key_name"]
input_password = get_credentials(input_key_name, input_username)

# auth_username = configInfo["auth_username"]
# auth_key_name = configInfo["auth_key_name"]
# auth_password = input_password

print({'input_username': input_username, 'input_password': input_password})
# print({'auth_username': auth_username, 'auth_password': auth_password})

f = Faker()
URL = configInfo['url']
digits = string.digits
preffix = ['86', '87', '84', '85', '82']


# def find_input_field_and_fill_value(driver, input_value, text_to_console, x_path):
# 	global INDEXES
# 	try:
# 		tag = driver.find_element(By.XPATH, x_path)
# 		tag.clear()
# 		for char in input_value:
# 			tag.send_keys(char)
# 			time.sleep(0.0000001)
# 		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
# 		# display_text_std(text)
# 	except Exception as e:
# 		text_to_console = e.args
# 		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
# 		# display_text_std(text)

def create_letter(driver, generated_CIF, filename):
# def create_letter(generated_CIF, filename):
    global INDEXES
    # client_nrs = random.choices(CLIENT_NRs, k=4)
    print('=' * 100)
    print(filename)
    print('=' * 100)
    path = os.getcwd() + f'\\{filename}\\data.xlsx'
    print('FILE EXISTS:', os.path.exists(filename))
    print('=' * 100)
    filename_ = f'C:/Users/a248433/Documents/SB Mozambique/EDO/Robotics/Bots/Automatização do Processo Abertura de contas Corporate/Clientes/approved/{filename}/data.xlsx'
    client_cif_df = ClientCIFs(filename=filename_)
    client_sociedade = ClientesSociedadePrenchido(filename=filename_)
    sociedade_constituida = ContasSociedadeConstituidaPr(filename=filename_)
    
    index = 0
    text = f'\n\t{chr(482)} = pass.\t{chr(530)} = Fail.'
    # display_text_std(text)

    driver = get_driver()
    driver.get(URL)
    driver.maximize_window()
    initial_window = driver.current_window_handle

    start = time.perf_counter()

    print('\n')
    login_user(driver, 'input_user')

    text = f'{INDEXES.pop(0)}. Bot is attempting to switch frame.   {chr(482)}'
    # display_text_std(text)
    frames = driver.find_elements(By.TAG_NAME, 'frame')
    frame2 = frames[1]
    driver.switch_to.frame(frame2)
    text = f'{INDEXES.pop(0)}. Bot switched to a new frame.   {chr(482)}'
    # display_text_std(text)

    text = 'Menu Account Origination Officer'
    find_element_by_xpath_and_click(
        driver,
        selector='/html/body/div[3]/ul/li/span',
        delay=0.001)
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)

    text = 'Contas'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/span')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Sociedades'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/ul/li[2]/span')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Constituidas'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/ul/li[2]/ul/li[1]/a')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    window_handles = driver.window_handles
    size = len(window_handles)
    another_window = window_handles[size-1]
    driver.switch_to.window(another_window)
    driver.maximize_window()
    text = f'{INDEXES.pop(0)}. Bot switched driver to newly opened window.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Current Account - Corporate'
    input_field = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/input[1]')
    input_field.clear()
    for char in '100':
        input_field.send_keys(char)
    input_field.submit()
    time.sleep(1)


    text = 'Curstomer Id'
    customer_id_tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[3]/input')
    # customer_id_tag.clear()
    for char in str(generated_CIF):
        customer_id_tag.send_keys(char)
    customer_id_tag.submit()

    Mnemonic = str(sociedade_constituida.get_value('Mnemonico'))
    find_input_field_and_fill_value(
        driver, 
        input_value=Mnemonic,
        text_to_console='Mnemonico',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[3]/input')

    GB_Nome_da_Conta = str(sociedade_constituida.get_value('GB # Nome da Conta'))
    find_input_field_and_fill_value(
        driver, 
        input_value=GB_Nome_da_Conta,
        text_to_console='GB # Nome da Conta',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[4]/td[3]/input')

    GB_Nome_da_conta_Cont = str(sociedade_constituida.get_value('GB # Nome da Conta (Cont)'))
    find_input_field_and_fill_value(
        driver, 
        input_value=GB_Nome_da_conta_Cont,
        text_to_console='GB # Nome da Conta (Cont)',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[3]/input')

    GB_Nome_Curto = str(sociedade_constituida.get_value('GB # Nome Curto'))
    find_input_field_and_fill_value(
        driver, 
        input_value=GB_Nome_Curto,
        text_to_console='GB # Nome Curto',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[6]/td[3]/input')

    Moeda = str(sociedade_constituida.get_value('Moeda'))
    find_input_field_and_fill_value(
        driver, 
        input_value=Moeda,
        text_to_console='Moeda',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/input')

    Nome_do_Ecosistema_1 = str(sociedade_constituida.get_value('Nome do Ecosistema.1'))
    find_input_field_and_fill_value(
        driver, 
        input_value=Nome_do_Ecosistema_1,
        text_to_console='Nome do Ecosistema.1',
        x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[8]/td[3]/input')

    text = 'Tipo de Ecosistema.1'
    select_tag = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[9]/td[3]/select'))
    # select by visible text
    select_tag.select_by_visible_text('1N/a')

    Inib_Chq = str(sociedade_constituida.get_value('Inib Chq'))
    if Inib_Chq == 'Y':
        path = '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[13]/td[3]/table/tbody/tr/td[3]/input'
    elif Inib_Chq == 'N':
        path = '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[13]/td[3]/table/tbody/tr/td[2]/input'
    else:
        path = '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[13]/td[3]/table/tbody/tr/td[1]/input'
    find_element_by_xpath_and_click(
        driver,
        selector=path,
        delay=0.001)
    

    counter = 0
    for row in client_cif_df.df.itertuples():
        if counter > 0:
            # In case there are more signers of the account
            text = 'Outro(s) Titular(es).1'
            add_more_signers = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[15]/td[2]/a[1]/img')
            add_more_signers.click()
            text = f'{INDEXES.pop(0)}. Bot selected "D" from selet tag under {text!r}.   {chr(482)}'
            # display_text_std(text)
            time.sleep(1)
        counter += 1

    index = 15
    counter = 0
    for row in client_cif_df.df.itertuples():
        NOME_da_PESSOA = row.NOME
        CARGO_NA_EMPRESA = str(row.CARGO)
        NUMERO_DO_CLIENTE = str(row.CIF)
        
        path_cif = f'/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[{index}]/td[3]/input'
        path_relationsh = f'/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[{index+1}]/td[3]/input'
        
        find_input_field_and_fill_value(
            driver, 
            input_value=NUMERO_DO_CLIENTE,
            text_to_console='Relacao 1',
            x_path=path_cif)

        Relacao_1 = 'None'
        if CARGO_NA_EMPRESA == 'A':
            Relacao_1 = '16'
        find_input_field_and_fill_value(
            driver, 
            input_value=Relacao_1,
            text_to_console='Relacao 1',
            x_path=path_relationsh)
        
        index += 2
        counter += 1
    
    Confirm_Doc_Suporte_ME = str(sociedade_constituida.get_value('Confirm Doc Suporte M.E'))
    Confirm_Doc_Suporte_ME_tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[14]/td[3]/input[1]')
    if Confirm_Doc_Suporte_ME == 'Y':
        Confirm_Doc_Suporte_ME_tag.click()

    
    created_account = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span')
    created_account = created_account.text
    text = f'{INDEXES.pop(0)}. Bot clicked on {created_account!r}.   {chr(482)}'
    display_text_std(text)
    sociedade_constituida.update_val(filename, 'B15', created_account)

    text = 'Validate Deal'
    validate_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[2]/a/img')
    validate_deal.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    # time.sleep(10)

    text = 'Commit the Deal'
    commit_the_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[1]/a/img')
    commit_the_deal.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'	
    
    driver.switch_to.window(initial_window)
    text = f'{INDEXES.pop(0)}. Bot switched driver to newly opened window.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)
    # driver.refresh()
    logout_user(driver)

    login_user(driver, 'auth_user')

    text = f'{INDEXES.pop(0)}. Bot is attempting to switch frame.   {chr(482)}'
    # display_text_std(text)
    frames = driver.find_elements(By.TAG_NAME, 'frame')
    frame2 = frames[1]
    driver.switch_to.frame(frame2)
    text = f'{INDEXES.pop(0)}. Bot switched to a new frame.   {chr(482)}'

    
    text = 'Menu Account Origination Officer'
    find_element_by_xpath_and_click(
        driver,
        selector='/html/body/div[3]/ul/li/span',
        delay=0.001)
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)

    text = 'Contas'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/span')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Sociedades'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/ul/li[2]/span')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Constituidas'
    client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/ul/li/ul/li[2]/ul/li[2]/ul/li[1]/a')))
    client_menu.click()
    text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    window_handles = driver.window_handles
    size = len(window_handles)
    another_window = window_handles[size-1]
    driver.switch_to.window(another_window)
    driver.maximize_window()
    text = f'{INDEXES.pop(0)}. Bot switched driver to newly opened window.   {chr(482)}'
    # display_text_std(text)
    time.sleep(0.001)

    text = 'Current Account - Corporate'
    input_field = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/input[1]')
    input_field.clear()
    created_account = str(created_account).replace('-', '')
    for char in created_account:
        input_field.send_keys(char)
    input_field.submit()
    time.sleep(1)

    submit_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[3]/a/img')
    submit_deal.click()
    time.sleep(2)

    try:
        authorize_deal = driver.find_element(By.XPATH, '//*[@id="goButton"]/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[5]/img')
        authorize_deal.click()
    except Exception as e:
        pass

    generator = LetterGeneration(
        sociedade_constituida.get_value('GB # Nome:'),
        str(client_sociedade.get_value('GB # Endereco Residencial')),
        str(client_sociedade.get_value('Telefone Celular.1')),
        client_sociedade.get_value('GB # Nome:'),
        str(created_account),
        '117', 'S17 AGENCIA DA MATOLA',
        '000301170903044100587', 'MZ59000301170903044100587',
        'SBICMZMX', 'MZN', 'Ivone Faquir', 'Nacima Khan'
    )
    generator.generate(filename=client_sociedade.get_value('GB # Nome:'))
    # time.sleep(5)




# create_letter(generated_CIF='756001')

