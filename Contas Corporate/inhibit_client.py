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
	display_text_std, find_element_by_xpath_and_click,
	common_handler, handle_user_auth, change_branch,
	change_to_last_window,
	INDEXES,
)


f = open("config.txt")
configInfo = json.loads(f.read())

df = pd.read_excel('data.xlsx')
CLIENT_NRs_AND_BRANCHES = df[['cus_n', 'BCH_Desc']].values.tolist()
# print(CLIENT_NRs_AND_BRANCHES)

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

def generate_date():
	date = dt.date(random.randint(2010, 2022), random.randint(1, 12), random.randint(1, 28))
	date = date.strftime('%Y%m%d')
	return date

def get_suffix():
	return ''.join([random.choice(digits) for _ in range(7)])

def generate_phone_number():
	return '258' + random.choice(preffix) + get_suffix()


def find_input_field_and_fill_value(driver, input_value, text_to_console, x_path):
	global INDEXES
	try:
		tag = driver.find_element(By.XPATH, x_path)
		tag.clear()
		for char in input_value:
			tag.send_keys(char)
			time.sleep(0.0001)
		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
		display_text_std(text)
	except Exception as e:
		text_to_console = e.args
		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
		display_text_std(text)


def process_action(client_name):
	global INDEXES
	# client_nrs = random.choices(CLIENT_NRs, k=4)
	index = 0
	text = f'\n\t{chr(482)} = pass.\t{chr(530)} = Fail.'
	display_text_std(text)

	driver = get_driver()
	driver.get(URL)
	driver.maximize_window()
	# time.sleep(1)
	initial_window = driver.current_window_handle

	# while len(client_nrs) > index:
	start = time.perf_counter()
	# client_nr = client_nrs[index]

	print('\n')
	display_text_std(text=f'Attempting to search client = {client_name!r}')
	login_user(driver, 'input_user')

	text = f'{INDEXES.pop(0)}. Bot is attempting to switch frame.   {chr(482)}'
	display_text_std(text)
	frames = driver.find_elements(By.TAG_NAME, 'frame')
	frame2 = frames[1]
	driver.switch_to.frame(frame2)
	text = f'{INDEXES.pop(0)}. Bot switched to a new frame.   {chr(482)}'
	display_text_std(text)

	# change_branch(driver, branch_name)
	# change_to_last_window(driver)

	'menu135144127802'

	# common_handler(driver, client_nr)

	text = 'Menu Account Origination Officer'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/ul/li/span',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	display_text_std(text)


	text = 'Consultas'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/ul/li/ul/li[15]/span',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	display_text_std(text)


	text = 'Consultas de Clientes'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/ul/li/ul/li[15]/ul/li[9]/span',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	display_text_std(text)


	text = 'Informacao Lista de Clientes'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/ul/li/ul/li[15]/ul/li[9]/ul/li[2]/a',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	display_text_std(text)

	window_handles = driver.window_handles
	size = len(window_handles)
	another_window = window_handles[size-1]
	driver.switch_to.window(another_window)
	driver.maximize_window()
	text = f'{INDEXES.pop(0)}. Bot switched driver to a new window.   {chr(482)}'
	display_text_std(text)
	time.sleep(0.02)


	# short_name_field = driver.find_element(By.XPATH, '/html/body/div[3]/form/table/tbody/tr/td[2]/table/tbody/tr[3]/td/div/table/tbody/tr[1]/td[3]/input[1]')
	# short_name_field.clear()
	# for char in client_name:
	# 	short_name_field.send_keys(char)
	# 	time.sleep(0.01)
	# text = f'{INDEXES.pop(0)}. Bot inserted {client_name!r} into input field.   {chr(482)}'
	# display_text_std(text)

	find_input_field_and_fill_value(
		driver, 
		input_value=client_name,
		text_to_console='Nome',
		x_path='/html/body/div[3]/form/table/tbody/tr/td[2]/table/tbody/tr[3]/td/div/table/tbody/tr[1]/td[3]/input[1]')

	select_tag = Select(driver.find_element(By.XPATH, '/html/body/div[3]/form/table/tbody/tr/td[2]/table/tbody/tr[3]/td/div/table/tbody/tr[1]/td[2]/select'))
	# select by visible text
	select_tag.select_by_visible_text('matches')
	text = f'{INDEXES.pop(0)}. Bot selected "matches" from selet tag.   {chr(482)}'
	display_text_std(text)
	
	find_btn = driver.find_element(By.XPATH, '/html/body/div[3]/form/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/table/tbody/tr/td/a')
	find_btn.click()
	text = f'{INDEXES.pop(0)}. Bot clicked on "Find" button.   {chr(482)}'
	display_text_std(text)
	
	text = f'{INDEXES.pop(0)}. Bot awaiting visibility of result table.   {chr(482)}'
	display_text_std(text)
	WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/form/div')))

	try:
		WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'datadisplay')))
		text = f'{INDEXES.pop(0)}. Bot found result table.   {chr(482)}'
		display_text_std(text)

		client_cif = None
		client_fullname = None
		for tr_id in range(1, 100):
			tr_id = f'r{tr_id}'
			try:
				tr_tag = driver.find_element(By.ID, tr_id)
				td_elements = tr_tag.find_elements(By.TAG_NAME, 'td')
				td_elements = [tag.text for tag in td_elements]
				client_number, client_full_name, *tags = td_elements

				first_name, last_name = client_name.split('...')
				client_first_name, *middle_name, client_last_name = client_full_name.split(' ')
				
				if first_name == client_first_name and last_name == client_last_name:
					client_cif = client_number
					client_fullname = client_full_name
			except Exception as e:
				# print(e.args)
				pass
		
		text = f'{INDEXES.pop(0)}. Bot found client_cif={client_cif}.   {chr(482)}'
		display_text_std(text)
		text = f'{INDEXES.pop(0)}. Bot found client_fullname={client_fullname}.   {chr(482)}'
		display_text_std(text)

		driver.switch_to.window(initial_window)
		text = f'{INDEXES.pop(0)}. Bot switched to first window.   {chr(482)}'
		display_text_std(text)
		driver.refresh()







		login_user(driver, 'input_user')

		text = f'{INDEXES.pop(0)}. Bot is attempting to switch frame.   {chr(482)}'
		display_text_std(text)
		frames = driver.find_elements(By.TAG_NAME, 'frame')
		frame2 = frames[1]
		driver.switch_to.frame(frame2)
		text = f'{INDEXES.pop(0)}. Bot switched to a new frame.   {chr(482)}'
		display_text_std(text)

		text = 'Menu Account Origination Officer'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/ul/li/span',
			delay=0.001)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Cliente'
		client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pane_"]/ul/li/ul/li[1]/span')))
		client_menu.click()
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)
		time.sleep(0.001)

		text = 'Abertura Clientes - Sociedades'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/ul/li/ul/li[1]/ul/li[3]/a',
			delay=0.001)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		window_handles = driver.window_handles
		size = len(window_handles)
		another_window = window_handles[size-1]
		driver.switch_to.window(another_window)
		driver.maximize_window()
		text = f'{INDEXES.pop(0)}. Bot switched driver to newly opened window.   {chr(482)}'
		display_text_std(text)
		time.sleep(0.001)

	
		text = 'New Deal'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[6]/a/img',
			delay=0.001)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Select Language Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Portuguese'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[5]/div[3]/table/tbody/tr[3]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		display_text_std(text)

		# DB_nome = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[6]/td[3]/input')
		# DB_nome.clear()
		# name = f.name()
		# for char in name:
		# 	DB_nome.send_keys(char)
		# 	time.sleep(0.001)
		name = f.name()
		find_input_field_and_fill_value(
			driver, 
			input_value=name,
			text_to_console='Nome',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[6]/td[3]/input')

		# DB_nome_cont = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/input')
		# DB_nome_cont.clear()
		# for char in name:
		# 	DB_nome_cont.send_keys(char)
		# 	time.sleep(0.02)
		find_input_field_and_fill_value(
			driver, 
			input_value=name,
			text_to_console='Nome (Cont)',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/input')

		# DB_nome_curto = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[8]/td[3]/input')
		# DB_nome_curto.clear()
		# for char in name:
		# 	DB_nome_curto.send_keys(char)
		# 	time.sleep(0.02)
		find_input_field_and_fill_value(
			driver, 
			input_value=name,
			text_to_console='Nome Curto',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[8]/td[3]/input')

		
		text = 'Select "Group Economico" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[10]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = '9999 - N/A'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow26"]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		display_text_std(text)

		# date_field = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[14]/td[3]/input')
		# date_field.clear()
		# date = dt.date(random.randint(2010, 2022), random.randint(1, 12), random.randint(1, 28))
		# text = f'{INDEXES.pop(0)}. Bot is adding "Data de Constituicao"={date.strftime("%d de %m de %Y")}.   {chr(482)}'
		# display_text_std(text)
		# date = generate_date()
		# for char in date:
		# 	date_field.send_keys(char)
		# 	time.sleep(0.01)
		date = generate_date()
		find_input_field_and_fill_value(
			driver, 
			input_value=date,
			text_to_console='Date Field',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[14]/td[3]/input')

		# text = 'Numero de Registo Comercial'
		# numero_de_registo_comercial_tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[15]/td[3]/input')
		# numero_de_registo_comercial_tag.clear()
		# numero_de_registo_comercial = 'NJDNJF'
		# text = f'{INDEXES.pop(0)}. Bot is inserting {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in numero_de_registo_comercial:
		# 	numero_de_registo_comercial_tag.send_keys(char)
		# 	time.sleep(0.01)
		find_input_field_and_fill_value(
			driver, 
			input_value='NJDNJF',
			text_to_console='Numero de Registo Comercial',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[15]/td[3]/input')

		text = 'Select "Legal Doc Name" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[17]/td[3]/select',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Others'
		select_tag = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[17]/td[3]/select'))
		# select by visible text
		select_tag.select_by_visible_text('Others')
		text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		display_text_std(text)

		# text = 'Conservatoria Registo Criminal'
		# conservatoria_reg_criminal_tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[18]/td[3]/input')
		# conservatoria_reg_criminal_tag.clear()
		# conservatoria_reg_criminal = 'Maputo'
		# text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in conservatoria_reg_criminal:
		# 	conservatoria_reg_criminal_tag.send_keys(char)
		# 	time.sleep(0.01)
		find_input_field_and_fill_value(
			driver, 
			input_value='Maputo',
			text_to_console='Conservatoria Registo Criminal',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[18]/td[3]/input')
		

		# text = 'Legal Issue Date'
		# legal_issue_date = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[20]/td[3]/input')
		# date = generate_date()
		# text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in date:
		# 	legal_issue_date.send_keys(char)
		# 	time.sleep(0.01)
		date = generate_date()
		find_input_field_and_fill_value(
			driver, 
			input_value=date,
			text_to_console='Legal Issue Date',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[20]/td[3]/input')

		
		# text = 'Legal Expire Date'
		# legal_expire_date = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[21]/td[3]/input')
		# # date = generate_date()
		# date = dt.date(random.randint(2023, 2025), random.randint(1, 12), random.randint(1, 28))
		# date = date.strftime('%Y%m%d')
		# text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in date:
		# 	legal_expire_date.send_keys(char)
		# 	time.sleep(0.01)
		legal_expire_date = generate_date()
		date = dt.date(random.randint(2023, 2025), random.randint(1, 12), random.randint(1, 28))
		legal_expire_date = date.strftime('%Y%m%d')
		find_input_field_and_fill_value(
			driver,
			input_value=legal_expire_date,
			text_to_console='Legal Expire Date',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[21]/td[3]/input')

		# text = 'Data de Escritura / Alrava'
		# data_escritura_alvara = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[22]/td[3]/input')
		# data_escritura_alvara.clear()
		# date = generate_date()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in date:
		# 	data_escritura_alvara.send_keys(char)
		# 	time.sleep(0.01)
		data_escritura_alvara = generate_date()
		find_input_field_and_fill_value(
			driver, 
			input_value=data_escritura_alvara,
			text_to_console='Data de Escritura / Alrava',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[22]/td[3]/input')

		# text = 'Numero Publicado BR'
		# numero_publicado_BR = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[23]/td[3]/input')
		# numero_publicado_BR.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# numero = "12345"
		# for char in numero:
		# 	numero_publicado_BR.send_keys(char)
		# 	time.sleep(0.01)
		numero = "12345"
		find_input_field_and_fill_value(
			driver, 
			input_value=numero,
			text_to_console='Numero Publicado BR',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[23]/td[3]/input')

		# text = 'Detalhes de Alteracao'
		# detalhes_de_alteracao = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[26]/td[3]/input')
		# detalhes_de_alteracao.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in 'N/A':
		# 	detalhes_de_alteracao.send_keys(char)
		# 	time.sleep(0.01)
		numero = "N/A"
		find_input_field_and_fill_value(
			driver, 
			input_value=numero,
			text_to_console='Detalhes de Alteracao',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[26]/td[3]/input')

	
		# text = 'Moeda de Salario'
		# salary_currency = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[27]/td[3]/input')
		# salary_currency.clear()
		# name = 'MZN'
		# text = f'{INDEXES.pop(0)}. Bot inserted salary currency: {name!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in name:
		# 	salary_currency.send_keys(char)
		# 	time.sleep(0.01)
		currency = "MZN"
		find_input_field_and_fill_value(
			driver, 
			input_value=currency,
			text_to_console='Moeda de Salario',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[27]/td[3]/input')

		# capital_social = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[28]/td[3]/input')
		# capital_social.clear()
		# valor = 'MZN 251'
		# for char in valor:
		# 	capital_social.send_keys(char)
		# 	time.sleep(0.01)
		valor = "MZN 251"
		find_input_field_and_fill_value(
			driver, 
			input_value=valor,
			text_to_console='Capital Social',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[28]/td[3]/input')

		text = 'Select "Sector" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[32]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)


		text = 'Company'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow72"]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r} under sector section.   {chr(482)}'
		display_text_std(text)

		
		text = 'Select "Tipo de Entidade" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[33]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Associacao - ONG'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow7"]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r} under sector section.   {chr(482)}'
		display_text_std(text)


		text = 'Select "Ramo de Actividade" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[34]/td[3]/a[1]/img',
			delay=0.2)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'Industria de Turismo'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow7"]/td[2]',
			delay=0.2)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r} under sector section.   {chr(482)}'
		display_text_std(text)


		
		text = 'Select "No. Pessoa Colectiva" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[35]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'EMPRESARIO EM NOME INDIVIDUAL'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow32"]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r} under sector section.   {chr(482)}'
		display_text_std(text)

		# text = 'NUIT'
		# NUIT_field = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[44]/td[3]/input')
		# NUIT_field.clear()
		# NUIT = '129342846'
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under NUIT section.   {chr(482)}'
		# display_text_std(text)
		# for char in NUIT:
		# 	NUIT_field.send_keys(char)
		# 	time.sleep(0.01)
		NUIT = '129342846'
		find_input_field_and_fill_value(
			driver, 
			input_value=NUIT,
			text_to_console='NUIT',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[44]/td[3]/input')

		# text = 'Maputo'
		# Bairro_Fiscal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[45]/td[3]/input')
		# Bairro_Fiscal.clear()
		# bairro = 'Maputo'
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under Bairro Fiscal section.   {chr(482)}'
		# display_text_std(text)
		# for char in bairro:
		# 	Bairro_Fiscal.send_keys(char)
		# 	time.sleep(0.01)
		bairro = 'Maputo'
		find_input_field_and_fill_value(
			driver, 
			input_value=bairro,
			text_to_console='Bairro Fiscal',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[45]/td[3]/input')


		# text = 'Nome da Pessoa'
		# nome_da_pessoa = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[53]/td[3]/input')
		# nome_da_pessoa.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under Numero de Cliente section.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Joe Doe'
		# for char in nome:
		# 	nome_da_pessoa.send_keys(char)
		# 	time.sleep(0.01)
		nome = 'Joe Doe'
		find_input_field_and_fill_value(
			driver, 
			input_value=nome,
			text_to_console='Nome da Pessoa',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[53]/td[3]/input')
			

		text = 'Cargo na Empresa'
		Cargo_na_Empresa = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[54]/td[3]/select'))
		# select by visible text
		Cargo_na_Empresa.select_by_visible_text('D')
		text = f'{INDEXES.pop(0)}. Bot selected "D" from selet tag under {text!r}.   {chr(482)}'
		display_text_std(text)

		# text = 'Client CIF'
		# numero_client = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[55]/td[3]/input')
		# numero_client.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under Numero de Cliente section.   {chr(482)}'
		# display_text_std(text)
		# for char in client_cif:
		# 	numero_client.send_keys(char)
		# 	time.sleep(0.01)
		find_input_field_and_fill_value(
			driver, 
			input_value=client_cif,
			text_to_console='Client CIF',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[55]/td[3]/input')

		
		# text = 'Endereco Residencial'
		# endereco_residencial = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[59]/td[3]/input')
		# endereco_residencial.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Downtown Manhattan'
		# for char in nome:
		# 	endereco_residencial.send_keys(char)
		# 	time.sleep(0.01)
		endereco_residencial = 'Downtown Manhattan'
		find_input_field_and_fill_value(
			driver, 
			input_value=endereco_residencial,
			text_to_console='Endereco Residencial',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[59]/td[3]/input')

		''
		# text = 'New York'
		# cidade = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[60]/td[3]/input')
		# cidade.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'New York'
		# for char in nome:
		# 	cidade.send_keys(char)
		# 	time.sleep(0.01)
		cidade = 'New York'
		find_input_field_and_fill_value(
			driver, 
			input_value=cidade,
			text_to_console='Cidade',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[60]/td[3]/input')

		
		# text = 'Pais'
		# pais = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[61]/td[3]/input')
		# pais.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'MZ'
		# for char in nome:
		# 	pais.send_keys(char)
		# 	time.sleep(0.01)
		pais = 'MZ'
		find_input_field_and_fill_value(
			driver, 
			input_value=pais,
			text_to_console='Pais',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[61]/td[3]/input')

		
		# text = 'Provincia'
		# provincia = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[62]/td[3]/input')
		# provincia.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = '100000'
		# for char in nome:
		# 	provincia.send_keys(char)
		# 	time.sleep(0.01)
		provincia = '100000'
		find_input_field_and_fill_value(
			driver, 
			input_value=provincia,
			text_to_console='Provincia',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[62]/td[3]/input')

		
		# text = 'Localidade-1'
		# localidade = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[63]/td[3]/input')
		# localidade.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = '100000'
		# for char in nome:
		# 	localidade.send_keys(char)
		# 	time.sleep(0.01)
		localidade = '100000'
		find_input_field_and_fill_value(
			driver, 
			input_value=localidade,
			text_to_console='Localidade-1',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[63]/td[3]/input')

		
		# text = 'Localidade-2'
		# localidade2 = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[67]/td[3]/input')
		# localidade2.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = '100000'
		# for char in nome:
		# 	localidade2.send_keys(char)
		# 	time.sleep(0.01)
		localidade2 = '100000'
		find_input_field_and_fill_value(
			driver,
			input_value=localidade2,
			text_to_console='Localidade-2',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[67]/td[3]/input')

		
		# text = 'Distrito'
		# distrito = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[68]/td[3]/input')
		# distrito.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Kampfumo'
		# for char in nome:
		# 	distrito.send_keys(char)
		# 	time.sleep(0.01)
		distrito = 'Kampfumo'
		find_input_field_and_fill_value(
			driver,
			input_value=distrito,
			text_to_console='Distrito',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[68]/td[3]/input')

		
		# text = 'Telefone Celular'
		# telefone_celular = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[71]/td[3]/input')
		# telefone_celular.clear()
		# phone_nr = generate_phone_number()
		# text = f'{INDEXES.pop(0)}. Bot inserted {phone_nr!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in phone_nr:
		# 	telefone_celular.send_keys(char)
		# 	time.sleep(0.01)
		telefone_celular = generate_phone_number()
		find_input_field_and_fill_value(
			driver,
			input_value=telefone_celular,
			text_to_console='Telefone Celular',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[71]/td[3]/input')


		# text = 'Email'
		# email = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[72]/td[3]/input')
		# email.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = f.email()
		# for char in nome:
		# 	email.send_keys(char)
		# 	time.sleep(0.01)
		email = f.email()
		find_input_field_and_fill_value(
			driver,
			input_value=email,
			text_to_console='Email',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[72]/td[3]/input')

		time.sleep(0.1)

		text = 'Select Country Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[50]/td[3]/a[1]/img',
			delay=0.2)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)


		text = 'Mozambique'
		find_element_by_xpath_and_click(
			driver,
			selector='//*[@id="dropDownRow161"]/td[2]',
			delay=0.2)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		display_text_std(text)

		
		# text = 'Endereco Postal'
		# endereco_postal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[65]/td[3]/input')
		# endereco_postal.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'XYZ Main Street'
		# for char in nome:
		# 	endereco_postal.send_keys(char)
		# 	time.sleep(0.01)
		endereco_postal = 'XYZ Main Street'
		find_input_field_and_fill_value(
			driver, 
			input_value=endereco_postal,
			text_to_console='Endereco Postal',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[65]/td[3]/input')

		
		# text = 'Select "Pais Postal" Dropdown'
		# pais_postal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[66]/td[3]/input')
		# pais_postal.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'MZ'
		# for char in nome:
		# 	pais_postal.send_keys(char)
		# 	time.sleep(0.01)
		pais_postal = 'MZ'
		find_input_field_and_fill_value(
			driver, 
			input_value=pais_postal,
			text_to_console='Select "Pais Postal" Dropdown',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[66]/td[3]/input')
		
		text = 'Periodicidade de Extrato'
		periodicidade_de_extrato = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[75]/td[3]/input')
		periodicidade_de_extrato.clear()
		periodicidade_de_extrato.send_keys(2)
		text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		display_text_std(text)

		
		# text = 'Codigo de Gestor'
		# codigo_de_gestor = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[79]/td[3]/input')
		# codigo_de_gestor.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# codigo = '865'
		# for char in codigo:
		# 	codigo_de_gestor.send_keys(char)
		# 	time.sleep(0.01)
		codigo_de_gestor = '865'
		find_input_field_and_fill_value(
			driver, 
			input_value=codigo_de_gestor,
			text_to_console='Codigo de Gestor',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[79]/td[3]/input')

		text = 'Select "Segmento" Dropdown'
		find_element_by_xpath_and_click(
			driver,
			selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[80]/td[3]/a[1]/img',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		text = 'COMMERCIAL-T2'
		find_element_by_xpath_and_click(
			driver,
			# '//*[@id="dropDownRow1"]/td[2]'
			# selector='/html/body/div[5]/div[3]/table/tbody/tr[23]/td[2]',
			selector='//*[@id="dropDownRow22"]/td[2]',
			delay=0.02)
		text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
		display_text_std(text)

		
		# text = 'Vendas Mensais'
		# vendas_mensais = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[85]/td[3]/input')
		# vendas_mensais.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# valor = 'MZN 2345'
		# for char in valor:
		# 	vendas_mensais.send_keys(char)
		# 	time.sleep(0.01)
		vendas_mensais = 'MZN 2345'
		find_input_field_and_fill_value(
			driver, 
			input_value=vendas_mensais,
			text_to_console='Vendas Mensais',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[85]/td[3]/input')
		

		
		# text = 'Principais Accionistas e Quota'
		# Pricipais_Accionistas_e_Quotas = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[87]/td[3]/input')
		# Pricipais_Accionistas_e_Quotas.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Principais Accionistas e Quotas'
		# for char in nome:
		# 	Pricipais_Accionistas_e_Quotas.send_keys(char)
		# 	time.sleep(0.01)
		Pricipais_Accionistas_e_Quotas = 'Principais Accionistas e Quotas'
		find_input_field_and_fill_value(
			driver, 
			input_value=Pricipais_Accionistas_e_Quotas,
			text_to_console='Principais Accionistas e Quotas',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[87]/td[3]/input')

		text = 'CUSMKTPARTID'
		CUSMKTPARTID = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[93]/td[3]/select'))
		# select by visible text
		CUSMKTPARTID.select_by_visible_text('50')
		text = f'{INDEXES.pop(0)}. Bot selected "50" from selet tag under {text!r}.   {chr(482)}'
		display_text_std(text)

		
		text = 'Res Nonres'
		# This selects yes checkbox
		Res_Nonres = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[94]/td[3]/table/tbody/tr/td[2]/input')
		Res_Nonres.click()
		text = f'{INDEXES.pop(0)}. Bot selected "Yes" {text!r}.   {chr(482)}'
		display_text_std(text)

		
		# text = 'VAT Number'
		# VAT_Number = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[95]/td[3]/input')
		# VAT_Number.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = '762364'
		# for char in nome:
		# 	VAT_Number.send_keys(char)
		# 	time.sleep(0.01)
		VAT_Number = '762364'
		find_input_field_and_fill_value(
			driver, 
			input_value=VAT_Number,
			text_to_console='VAT Number',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[95]/td[3]/input')

		
		# text = 'Cidade de Maputo'
		# Convervatoria = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[96]/td[3]/input')
		# Convervatoria.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Cidade de Maputo'
		# for char in nome:
		# 	Convervatoria.send_keys(char)
		# 	time.sleep(0.01)
		Convervatoria = 'Cidade de Maputo'
		find_input_field_and_fill_value(
			driver, 
			input_value=Convervatoria,
			text_to_console='Convervatoria',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[96]/td[3]/input')

		
		text = 'Categoria KYC'
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[97]/td[3]/table/tbody/tr/td[2]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked "No" under {text!r}.   {chr(482)}'
		display_text_std(text)


		# text = 'Issue Cheques'
		# Issue_Checkes = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[98]/td[3]/select'))
		# # select by visible text
		# Issue_Checkes.select_by_visible_text('No')
		# text = f'{INDEXES.pop(0)}. Bot selected "NO" from selet tag under {text!r}.   {chr(482)}'
		# display_text_std(text)


		text = 'Kyc'
		Kyc = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[99]/td[3]/table/tbody/tr/td[3]/input')
		Kyc.click()
		text = f'{INDEXES.pop(0)}. Bot clicked "Y" under {text!r}.   {chr(482)}'
		display_text_std(text)

		# text = 'Data KYC'
		# date_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[100]/td[3]/input')
		# date_KYC.clear()
		# date = dt.date(random.randint(2010, 2022), random.randint(1, 12), random.randint(1, 28))
		# text = f'{INDEXES.pop(0)}. Bot is adding "{text}"={date.strftime("%d de %m de %Y")}.   {chr(482)}'
		# display_text_std(text)
		# date = generate_date()
		# for char in date:
		# 	date_KYC.send_keys(char)
		# 	time.sleep(0.01)
		date_KYC = generate_date()
		find_input_field_and_fill_value(
			driver, 
			input_value=date_KYC,
			text_to_console='Data KYC',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[100]/td[3]/input')

		
		
		# text = 'CAE'
		# CAE = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[101]/td[3]/input')
		# CAE.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r}.   {chr(482)}'
		# display_text_std(text)
		# text = 'CAE'
		# for char in text:
		# 	CAE.send_keys(char)
		# 	time.sleep(0.01)
		CAE = 'CAE'
		find_input_field_and_fill_value(
			driver, 
			input_value=CAE,
			text_to_console='CAE',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[101]/td[3]/input')


		text = 'Market Participant Type ID'
		Market_Participant_Type_ID = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[3]/td/table/tbody/tr[5]/td[4]/select'))
		# select by visible text
		Market_Participant_Type_ID.select_by_visible_text('30')
		text = f'{INDEXES.pop(0)}. Bot selected "30" from selet tag under {text!r}.   {chr(482)}'
		display_text_std(text)


		# In case there are more signers of the account
		text = 'More Signers'
		add_more_signers = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[53]/td[2]/a/img')
		add_more_signers.click()
		text = f'{INDEXES.pop(0)}. Bot selected "D" from selet tag under {text!r}.   {chr(482)}'
		display_text_std(text)
		time.sleep(1)

		
		# text = 'Nome do Primeiro Assinante'
		# nome_da_pessoa2 = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[56]/td[3]/input')
		# nome_da_pessoa2.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under Numero de Cliente section.   {chr(482)}'
		# display_text_std(text)
		# nome = 'Jane Doe'
		# for char in nome:
		# 	nome_da_pessoa2.send_keys(char)
		# 	time.sleep(0.01)
		nome_da_pessoa2 = 'Jane Doe'
		find_input_field_and_fill_value(
			driver, 
			input_value=nome_da_pessoa2,
			text_to_console='Nome do Segundo Assinante',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[56]/td[3]/input')
		

		text = 'Cargo na Empresa'
		Cargo_na_Empresa = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[57]/td[3]/select'))
		# select by visible text
		Cargo_na_Empresa.select_by_visible_text('P')
		text = f'{INDEXES.pop(0)}. Bot selected "P" from selet tag under {text!r}.   {chr(482)}'
		display_text_std(text)

		# text = 'Client CIF'
		# numero_client = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[58]/td[3]/input')
		# numero_client.clear()
		# text = f'{INDEXES.pop(0)}. Bot inserted {text!r} under Numero de Cliente section.   {chr(482)}'
		# display_text_std(text)
		# for char in client_cif:
		# 	numero_client.send_keys(char)
		# 	time.sleep(0.01)
		numero_client_2 = client_cif
		find_input_field_and_fill_value(
			driver, 
			input_value=numero_client_2,
			text_to_console='Client CIF do Segundo Assinante',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[58]/td[3]/input')



		text = 'Validate Deal'
		validate_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[2]/a/img')
		validate_deal.click()
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)

		# time.sleep(3)

		# text = 'Telefone Celular'
		# telefone_celular = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[74]/td[3]/input')
		# telefone_celular.clear()
		# phone_nr = generate_phone_number()
		# text = f'{INDEXES.pop(0)}. Bot inserted {phone_nr!r}.   {chr(482)}'
		# display_text_std(text)
		# for char in phone_nr:
		# 	telefone_celular.send_keys(char)
		# 	time.sleep(0.01)
		telefone_celular_client_2 = generate_phone_number()
		find_input_field_and_fill_value(
			driver, 
			input_value=telefone_celular_client_2,
			text_to_console='Numero de telefone do Segundo Assinante',
			x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[74]/td[3]/input')

		# text = 'Validate Deal'
		# validate_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[2]/a/img')
		# validate_deal.click()
		# text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		# display_text_std(text)

		text = 'Commit the Deal'
		commit_the_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[1]/a/img')
		commit_the_deal.click()
		text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
		display_text_std(text)
	except Exception as e:
		error = e.args
		text = f'{INDEXES.pop(0)}. Bot did not find result for client {error!r}.   {chr(482)}'
		display_text_std(text)

	time.sleep(60)
	driver.close()
	# driver.switch_to.window(initial_window)
	# display_text_std(f'{INDEXES.pop(0)}. Switched driver to initial window')

	# logout_user(driver)

	index += 1
	INDEXES = list(range(1, 100))
	text = f'>>> Bot processed in {time.perf_counter() - start} secs <<<'
	display_text_std(text)

# client_nr = 643574
# branch_name = 'C05 AGENCIA CHIMOIO'
# client_nr, branch_name = random.choice(CLIENT_NRs_AND_BRANCHES)
process_action(client_name="CARMEN...ZIMBA")
# process_action(client_nr, branch_name)
# process_action()


