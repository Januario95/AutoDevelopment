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
	# display_text_std, 
	find_element_by_xpath_and_click,
	common_handler, handle_user_auth, change_branch,
	change_to_last_window,
	INDEXES, ClientesSociedadePrenchido,
	ContasSociedadeConstituidaPr, ClientCIFs,
)

client_cif_df = ClientCIFs()
client_sociedade = ClientesSociedadePrenchido()
sociedade_constituida = ContasSociedadeConstituidaPr()

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


def find_input_field_and_fill_value(driver, input_value, text_to_console, x_path):
	global INDEXES
	try:
		tag = driver.find_element(By.XPATH, x_path)
		tag.clear()
		for char in input_value:
			tag.send_keys(char)
			time.sleep(0.0000001)
		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
		# display_text_std(text)
	except Exception as e:
		text_to_console = e.args
		text = f'{INDEXES.pop(0)}. Bot inserted {text_to_console!r} into input field.   {chr(482)}'
		# display_text_std(text)

def process_action(client_name):
	global INDEXES
	# client_nrs = random.choices(CLIENT_NRs, k=4)
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

	text = 'Cliente'
	client_menu = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pane_"]/ul/li/ul/li[1]/span')))
	client_menu.click()
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)
	time.sleep(0.00001)

	text = 'Abertura Clientes - Sociedades'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/ul/li/ul/li[1]/ul/li[3]/a',
		delay=0.001)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)

	window_handles = driver.window_handles
	size = len(window_handles)
	another_window = window_handles[size-1]
	driver.switch_to.window(another_window)
	driver.maximize_window()
	text = f'{INDEXES.pop(0)}. Bot switched driver to newly opened window.   {chr(482)}'
	# display_text_std(text)
	time.sleep(0.001)


	text = 'New Deal'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[6]/a/img',
		delay=0.001)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)

	language = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[3]/input')
	language.clear()
	language.send_keys(client_sociedade.get_value('Lingua Mae:'))
	text = f"{INDEXES.pop(0)}. Bot inserted language code: {client_sociedade.get_value('Lingua Mae:')!r}.   {chr(482)}"
	# display_text_std(text)

	name = client_sociedade.get_value('GB # Nome:')
	find_input_field_and_fill_value(
		driver, 
		input_value=name,
		text_to_console='Nome',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[6]/td[3]/input')

	find_input_field_and_fill_value(
		driver, 
		input_value=name,
		text_to_console='Nome (Cont)',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/input')

	find_input_field_and_fill_value(
		driver, 
		input_value=name,
		text_to_console='Nome Curto',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[8]/td[3]/input')

	date = str(client_sociedade.get_value('Data da Constituicao')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=date,
		text_to_console='Data de Constituicao',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[14]/td[3]/input')
	text = f"{INDEXES.pop(0)}. Bot inserted 'Data de Constituicao': {client_sociedade.get_value('Data da Constituicao')!r}.   {chr(482)}"
	# display_text_std(text)

	numero_de_registo_comercial = str(client_sociedade.get_value('Numero Registo Comercial'))
	find_input_field_and_fill_value(
		driver, 
		input_value=numero_de_registo_comercial,
		text_to_console='Numero de Registo Comercial',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[15]/td[3]/input')

	legal_id = client_sociedade.get_value('Legal ID.1')
	find_input_field_and_fill_value(
		driver, 
		input_value=legal_id,
		text_to_console='Legal ID.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[16]/td[3]/input')

	text = 'Select "Legal Doc Name" Dropdown'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[17]/td[3]/select',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)

	text = 'Others'
	select_tag = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[17]/td[3]/select'))
	# select by visible text
	select_tag.select_by_visible_text('Others')
	text = f'{INDEXES.pop(0)}. Bot selected {text!r}.   {chr(482)}'
	# display_text_std(text)

	conservatoria_registo_criminal = client_sociedade.get_value('Conservat. Reg.Comercial.1')
	find_input_field_and_fill_value(
		driver, 
		input_value=conservatoria_registo_criminal,
		text_to_console='Conservatoria Registo Criminal',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[18]/td[3]/input')
	
	legal_iss_auth_1 = str(client_sociedade.get_value('Legal Iss Auth.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=legal_iss_auth_1,
		text_to_console='Legal Iss Auth.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[19]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted LEGAL ISSUE AUTH {legal_iss_auth_1!r}.   {chr(482)}'
	# display_text_std(text)

	legal_iss_date_1 = str(client_sociedade.get_value('Legal Iss Date.1')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=legal_iss_date_1,
		text_to_console='Legal Iss Date.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[20]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted LEGAL ISSUE DATE {legal_iss_date_1!r}.   {chr(482)}'
	# display_text_std(text)

	legal_exp_date_1 = str(client_sociedade.get_value('legal Exp Date.1')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=legal_exp_date_1,
		text_to_console='legal Exp Date.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[21]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted LEGAL EXP DATE {legal_exp_date_1!r}.   {chr(482)}'
	# display_text_std(text)

	data_escritura_alvara = str(client_sociedade.get_value('Data Escritura/Alvara')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=data_escritura_alvara,
		text_to_console='Data Escritura/Alvara',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[22]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted DATA ESCRITURA ALVARA {data_escritura_alvara!r}.   {chr(482)}'
	# display_text_std(text)


	numero_de_publicacao_BR = client_sociedade.get_value('Numero de Publicacao B.R')
	find_input_field_and_fill_value(
		driver, 
		input_value=numero_de_publicacao_BR,
		text_to_console='Numero de Publicacao B.R',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[23]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted NUMERO DE PUBLICACAO BR {numero_de_publicacao_BR!r}.   {chr(482)}'
	# display_text_std(text)

	data_de_publicacao_BR = str(client_sociedade.get_value('Data de Publicacao B.R.')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=data_de_publicacao_BR,
		text_to_console='Data de Publicacao B.R.',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[24]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted DATA DE PUBLICACAO BR {data_de_publicacao_BR!r}.   {chr(482)}'
	# display_text_std(text)

	moeda_de_salario_1 = str(client_sociedade.get_value('Moeda de Salario.1')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=moeda_de_salario_1,
		text_to_console='Moeda de Salario.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[27]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted MOEDA DE SALARIO {moeda_de_salario_1!r}.   {chr(482)}'
	# display_text_std(text)


	capital_social = client_sociedade.get_value('Capital Social')
	find_input_field_and_fill_value(
		driver, 
		input_value=capital_social,
		text_to_console='Capital Social',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[28]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted CAPICAL SOCIAL {capital_social!r}.   {chr(482)}'
	# display_text_std(text)

	capital_emitido = client_sociedade.get_value('Capital Emitido')
	find_input_field_and_fill_value(
		driver, 
		input_value=capital_emitido,
		text_to_console='Capital Emitido',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[29]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted CAPICAL EMITIDO {capital_emitido!r}.   {chr(482)}'
	# display_text_std(text)

	Sector = str(client_sociedade.get_value('Sector')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=Sector,
		text_to_console='Sector',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[32]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted SECTOR {Sector!r}.   {chr(482)}'
	# display_text_std(text)

	tipo_de_entidade = str(client_sociedade.get_value('Tipo de Entidade')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=tipo_de_entidade,
		text_to_console='Tipo de Entidade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[33]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted TIPO DE ENTIDADE {tipo_de_entidade!r}.   {chr(482)}'
	# display_text_std(text)

	ramo_de_actividade = str(client_sociedade.get_value('Ramo de Actividade')).replace('-', '')[:8]
	find_input_field_and_fill_value(
		driver, 
		input_value=ramo_de_actividade,
		text_to_console='Ramo de Actividade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[34]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted RAMO DE ACTIVIDADE {ramo_de_actividade!r}.   {chr(482)}'
	# display_text_std(text)

	text = 'Select "Isento de Taxa" Dropdown'
	find_element_by_xpath_and_click(
		driver,
		selector='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[36]/td[3]/select',
		delay=0.02)
	text = f'{INDEXES.pop(0)}. Bot selected {text!r} under "Isento de Taxa".   {chr(482)}'
	# display_text_std(text)

	N_contribuinte_NUIT_1 = str(client_sociedade.get_value('N.Contribuinte (NUIT).1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=N_contribuinte_NUIT_1,
		text_to_console='N.Contribuinte (NUIT).1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[44]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted NUMERO CONTRIBUINTE NUIT {N_contribuinte_NUIT_1!r}.   {chr(482)}'
	# display_text_std(text)

	GB_bairro_fiscal = client_sociedade.get_value('GB # Bairro Fiscal')
	find_input_field_and_fill_value(
		driver, 
		input_value=GB_bairro_fiscal,
		text_to_console='GB # Bairro Fiscal',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[45]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted DB BAIRRO FISCAL {GB_bairro_fiscal!r}.   {chr(482)}'
	# display_text_std(text)

	codigo_fiscal = str(client_sociedade.get_value('Codigo Fiscal'))
	find_input_field_and_fill_value(
		driver, 
		input_value=codigo_fiscal,
		text_to_console='Codigo Fiscal',
		x_path='//html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[46]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted CODIGO FISCAL {codigo_fiscal!r}.   {chr(482)}'
	# display_text_std(text)

	estado_cliente = str(client_sociedade.get_value('Estado Cliente'))
	find_input_field_and_fill_value(
		driver, 
		input_value=estado_cliente,
		text_to_console='Estado Cliente',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[49]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted ENTIDADE CLIENTE {estado_cliente!r}.   {chr(482)}'
	# display_text_std(text)

	Nacionalidade = str(client_sociedade.get_value('Nacionalidade'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Nacionalidade,
		text_to_console='Nacionalidade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[50]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted NACIONALIDADE {Nacionalidade!r}.   {chr(482)}'
	# display_text_std(text)


	GB_endereco_residencial = str(client_sociedade.get_value('GB # Endereco Residencial'))
	find_input_field_and_fill_value(
		driver, 
		input_value=GB_endereco_residencial,
		text_to_console='GB # Endereco Residencial',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[59]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted GB ENDERECO RESIDENCIAL {GB_endereco_residencial!r}.   {chr(482)}'
	# display_text_std(text)

	GB_Cidade = str(client_sociedade.get_value('GB # Cidade'))
	find_input_field_and_fill_value(
		driver, 
		input_value=GB_Cidade,
		text_to_console='GB # Cidade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[60]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted GB CIDADE {GB_Cidade!r}.   {chr(482)}'
	# display_text_std(text)

	Pais = str(client_sociedade.get_value('Pais'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Pais,
		text_to_console='Pais',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[61]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted PAIS {Pais!r}.   {chr(482)}'
	# display_text_std(text)

	Provincia = str(client_sociedade.get_value('Provincia'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Provincia,
		text_to_console='Provincia',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[62]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted PROVINCIA {Provincia!r}.   {chr(482)}'
	# display_text_std(text)

	Localidade = str(client_sociedade.get_value('Localidade'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Localidade,
		text_to_console='Localidade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[63]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted LOCALIDADE {Localidade!r}.   {chr(482)}'
	# display_text_std(text)

	Codigo_Postal = str(client_sociedade.get_value('Codigo Postal'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Codigo_Postal,
		text_to_console='Codigo Postal',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[64]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted CODIGO POSTAL {Codigo_Postal!r}.   {chr(482)}'
	# display_text_std(text)

	Endereco_Postal_1 = str(client_sociedade.get_value('Endereco Postal.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Endereco_Postal_1,
		text_to_console='Endereco Postal.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[65]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted ENDERECO POSTAL 1 {Endereco_Postal_1!r}.   {chr(482)}'
	# display_text_std(text)

	Pais_Postal = str(client_sociedade.get_value('Pais Postal'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Pais_Postal,
		text_to_console='Pais Postal',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[66]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted PAIS POSTAL {Pais_Postal!r}.   {chr(482)}'
	# display_text_std(text)

	Localidade = str(client_sociedade.get_value('Localidade'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Localidade,
		text_to_console='Localidade',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[67]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted LOCALIDADE {Localidade!r}.   {chr(482)}'
	# display_text_std(text)

	Distrito = str(client_sociedade.get_value('Distrito'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Distrito,
		text_to_console='Distrito',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[68]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted DISTRITO {Distrito!r}.   {chr(482)}'
	# display_text_std(text)

	Provincia = str(client_sociedade.get_value('Provincia'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Provincia,
		text_to_console='Provincia',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[69]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted PROVINCIA {Provincia!r}.   {chr(482)}'
	# display_text_std(text)

	Telefone_Fixo_1 = str(client_sociedade.get_value('Telefone Fixo.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Telefone_Fixo_1,
		text_to_console='Telefone Fixo.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[70]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted TELEFONE FIXO 1 {Telefone_Fixo_1!r}.   {chr(482)}'
	# display_text_std(text)

	Telefone_Celular_1 = str(client_sociedade.get_value('Telefone Celular.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Telefone_Celular_1,
		text_to_console='Telefone Celular.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[71]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted TELEFONE CELULAR 1 {Telefone_Celular_1!r}.   {chr(482)}'
	# display_text_std(text)

	Email_1 = str(client_sociedade.get_value('Email.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Email_1,
		text_to_console='Email.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[72]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted EMAIL 1 {Email_1!r}.   {chr(482)}'
	# display_text_std(text)

	Fax_Emprego_1 = str(client_sociedade.get_value('Fax (Emprego).1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Fax_Emprego_1,
		text_to_console='Fax (Emprego).1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[73]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted FAX EMPREGO 1 {Fax_Emprego_1!r}.   {chr(482)}'
	# display_text_std(text)

	E_Mail = str(client_sociedade.get_value('E-Mail'))
	find_input_field_and_fill_value(
		driver, 
		input_value=E_Mail,
		text_to_console='E-Mail',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[74]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted E_EMAIL {E_Mail!r}.   {chr(482)}'
	# display_text_std(text)

	Periocidade_do_Extracto = str(client_sociedade.get_value('Periocidade do Extracto'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Periocidade_do_Extracto,
		text_to_console='Periocidade do Extracto',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[75]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted PERIODICIDADE DE EXTRACAO {Periocidade_do_Extracto!r}.   {chr(482)}'
	# display_text_std(text)

	Codigo_de_Gestor = str(client_sociedade.get_value('Codigo de Gestor'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Codigo_de_Gestor,
		text_to_console='Codigo de Gestor',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[79]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted CODIGO DE GESTOR {Codigo_de_Gestor!r}.   {chr(482)}'
	# display_text_std(text)

	Segmento = str(client_sociedade.get_value('Segmento'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Segmento,
		text_to_console='Segmento',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[80]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted SEGMENTO {Segmento!r}.   {chr(482)}'
	# display_text_std(text)

	Vendas_Mensais = str(client_sociedade.get_value('Vendas Mensais'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Vendas_Mensais,
		text_to_console='Vendas Mensais',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[85]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted {Vendas_Mensais!r} under VENDAS MENSAIS.   {chr(482)}'
	# display_text_std(text)

	Princ_Accion_e_Quotas_1 = str(client_sociedade.get_value('Princ.Accion.e Quotas.1'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Princ_Accion_e_Quotas_1,
		text_to_console='Princ.Accion.e Quotas.1',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[87]/td[3]/input')
	text = f"{INDEXES.pop(0)}. Bot inserted {Princ_Accion_e_Quotas_1!r} under 'Princ.Accion.e Quotas.1'.   {chr(482)}"
	# display_text_std(text)

	text = 'CUSMKTPARTID'
	CUSMKTPARTID = str(client_sociedade.get_value('CUSMKTPARTID'))
	CUSMKTPARTID_TAG = Select(driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[93]/td[3]/select'))
	# select by visible text
	CUSMKTPARTID_TAG.select_by_visible_text(CUSMKTPARTID)
	text = f'{INDEXES.pop(0)}. Bot selected {CUSMKTPARTID} from selet tag under {text!r}.   {chr(482)}'
	# display_text_std(text)

	text = 'Res Nonres'
	# This selects yes checkbox
	Res_Nonres = str(client_sociedade.get_value('Res Nonres'))
	Res_Nonres_Tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[94]/td[3]/table/tbody/tr/td[2]/input')
	Res_Nonres_Tag.click()
	text = f'{INDEXES.pop(0)}. Bot selected "Yes" under {text}.   {chr(482)}'
	# display_text_std(text)	

	VAT_NUMBER = str(client_sociedade.get_value('VAT NUMBER'))
	find_input_field_and_fill_value(
		driver, 
		input_value=VAT_NUMBER,
		text_to_console='VAT NUMBER',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[95]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted VAT NUMBER {VAT_NUMBER!r}.   {chr(482)}'
	# display_text_std(text)

	Conservatoria = str(client_sociedade.get_value('Conservatoria'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Conservatoria,
		text_to_console='Conservatoria',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[96]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted Conservatoria {Conservatoria!r}.   {chr(482)}'
	# display_text_std(text)

	text = 'Categoria KYC'
	KYCFlag = str(client_sociedade.get_value('KYC Flag'))
	if KYCFlag == 'N':
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[97]/td[3]/table/tbody/tr/td[2]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked {KYCFlag!r} under KYC Flag.   {chr(482)}'
	elif KYCFlag == 'Y':
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[97]/td[3]/table/tbody/tr/td[3]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked {KYCFlag!r} under KYC Flag.   {chr(482)}'
	else:
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[97]/td[3]/table/tbody/tr/td[1]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked "[None]" under KYC Flag.   {chr(482)}'
	# display_text_std(text)


	text = 'Kyc'
	KYC = str(client_sociedade.get_value('KYC Flag'))
	if KYC == 'N':
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[99]/td[3]/table/tbody/tr/td[2]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked {KYC!r} under KYC.   {chr(482)}'
	elif KYC == 'Y':
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[99]/td[3]/table/tbody/tr/td[3]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked {KYC!r} under KYC.   {chr(482)}'
	else:
		Categoria_KYC = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[99]/td[3]/table/tbody/tr/td[1]/input')
		Categoria_KYC.click()
		text = f'{INDEXES.pop(0)}. Bot clicked "[None]" under KYC.   {chr(482)}'
	# display_text_std(text)


	Data_do_KYC = str(client_sociedade.get_value('Data do KYC')).replace('-', '')
	find_input_field_and_fill_value(
		driver, 
		input_value=Data_do_KYC,
		text_to_console='Data do KYC',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[100]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted {Data_do_KYC!r} under Data do KYC.   {chr(482)}'
	# display_text_std(text)

	CAE = str(client_sociedade.get_value('CAE'))
	find_input_field_and_fill_value(
		driver, 
		input_value=CAE,
		text_to_console='CAE',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[101]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted {CAE!r} under Data do KYC.   {chr(482)}'
	# display_text_std(text)

	Customer_Since = str(client_sociedade.get_value('Customer Since'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Customer_Since,
		text_to_console='Customer Since',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[102]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted {Customer_Since!r} under Customer Since.   {chr(482)}'
	# display_text_std(text)
	time.sleep(3)

	Share_Data_for_Direct_Mkt = str(client_sociedade.get_value('Share Data for Direct Mkt'))
	Share_Data_for_Direct_Mkt_tag = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[103]/td[3]/input[1]')
	Share_Data_for_Direct_Mkt_tag.click()
	text = f'{INDEXES.pop(0)}. Bot clicked on {Share_Data_for_Direct_Mkt!r}.   {chr(482)}'
	# display_text_std(text)

	index = 53
	counter = 0
	for row in client_cif_df.df.itertuples():
		NOME_da_PESSOA = row.NOME
		CARGO_NA_EMPRESA = row.CARGO
		NUMERO_DO_CLIENTE = str(row.CIF)

		if counter > 0:
			# In case there are more signers of the account
			text = 'More Signers'
			add_more_signers = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[53]/td[2]/a/img')
			add_more_signers.click()
			text = f'{INDEXES.pop(0)}. Bot selected "D" from selet tag under {text!r}.   {chr(482)}'
			# display_text_std(text)
			time.sleep(1)

		path_name = f'/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[{index}]/td[3]/input'
		find_input_field_and_fill_value(
			driver, 
			input_value=NOME_da_PESSOA,
			text_to_console='NOME_da_PESSOA',
			x_path=path_name)
		text = f'{INDEXES.pop(0)}. Bot inserted {NOME_da_PESSOA!r} under NOME_da_PESSOA.   {chr(482)}'
		# display_text_std(text)

		path_cargo = f'/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[{index+1}]/td[3]/select'
		Cargo_na_Empresa = Select(driver.find_element(By.XPATH, path_cargo))
		# select by visible text
		Cargo_na_Empresa.select_by_visible_text(CARGO_NA_EMPRESA)
		text = f'{INDEXES.pop(0)}. Bot selected {CARGO_NA_EMPRESA!r} under CARGO_NA_EMPRESA.   {chr(482)}'
		# display_text_std(text)
		
		path_CIF = f'/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[{index+2}]/td[3]/input'
		find_input_field_and_fill_value(
			driver, 
			input_value=NUMERO_DO_CLIENTE,
			text_to_console='NUMERO_DO_CLIENTE',
			x_path=path_CIF)
		text = f'{INDEXES.pop(0)}. Bot inserted {NUMERO_DO_CLIENTE!r} under NUMERO_DO_CLIENTE.   {chr(482)}'
		# display_text_std(text)

		index += 3
		counter += 1
	

	text = 'Validate Deal'
	validate_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[2]/a/img')
	validate_deal.click()
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)
	time.sleep(3)

	Customer_Since = str(client_sociedade.get_value('Customer Since'))
	find_input_field_and_fill_value(
		driver, 
		input_value=Customer_Since,
		text_to_console='Customer Since',
		x_path='/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[1]/td/table/tbody/tr[102]/td[3]/input')
	text = f'{INDEXES.pop(0)}. Bot inserted {Customer_Since!r} under Customer Since.   {chr(482)}'
	# display_text_std(text)
	time.sleep(3)

	text = 'Commit the Deal'
	commit_the_deal = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/form[1]/div[2]/table/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[1]/a/img')
	commit_the_deal.click()
	text = f'{INDEXES.pop(0)}. Bot clicked on {text!r}.   {chr(482)}'
	# display_text_std(text)

	time.sleep(120)
	driver.close()
	# driver.switch_to.window(initial_window)
	# display_text_std(f'{INDEXES.pop(0)}. Switched driver to initial window')

	# logout_user(driver)

	index += 1
	INDEXES = list(range(1, 100))
	text = f'>>> Bot processed in {time.perf_counter() - start} secs <<<'
	# display_text_std(text)

process_action(client_name="CARMEN...ZIMBA")



