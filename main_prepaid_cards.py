# -*- coding: utf-8 -*-
"""
Created on Tue Apr  6 08:34:49 2021

@author: C835910
"""

import pandas as pd
import datetime
import os
from os import path
import shutil
import uuid

#import splunk
import keyring
import getpass
import socket

from fnmatch import filter
from time import sleep
import win32com.client
from win32com.client import Dispatch, constants
import logger_config

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from configparser import SafeConfigParser

const = win32com.client.constants
parser = SafeConfigParser()
parser.read('config.ini')

MS_TEAM_CHANNEL = "https://standardbank.webhook.office.com/webhookb2/bba80477-b70b-4896-9436-ed20783007b7@7369e6ec-faa6-42fa-bc0e-4f332da5b1db/IncomingWebhook/6eee66f3505d4bf8b3f67e4318553e01/68d36e73-ca81-4c44-8fd2-f311b855ada1"
MAIN_PATH = os.path.expanduser("~\Desktop\Data")
T24_URL = parser.get('geral', 'T24_LINK')
logger = logger_config.get_logger('Prepaid Card Process')

def send_email(file, main_account, total, success, error, mail=None):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "PrePaid Report"
    newMail.To = parser.get('geral', 'MAIL_BOT_ADMIN')
    newMail.HTMLBody = str(parser.get('geral', 'MAIL')).format(success, error, main_account)
    newMail.BCC = parser.get('geral', 'MAIL_BOT_ADMIN')
    newMail.Attachments.Add(Source=file)
    newMail.Send()

def send_approval_email(subject, file, sender):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.To = 'anisio.marrima@standardbank.co.mz'
    newMail.HTMLBody = f'<b>Caro cliente</b><p>Recebemos o ficheiro em anexo De: {sender}.</p><p>Queira por favor confirmar respondento a este e-mail com a palavra `Aprovado` caso seja legítimo ou ignorar caso não seja.</p>'
    newMail.Attachments.Add(Source=file)
    newMail.Send()

def init_t24():
    password = keyring.get_password("T24_PREPAID", parser.get('geral', 'T24_USERNAME'))
    print('Iniciando o T24...')
    driver = webdriver.Chrome(parser.get('geral', 'DRIVER_PATH'))
    driver.implicitly_wait(45)
    driver.get(T24_URL)
    driver.maximize_window()
    driver.find_element_by_name('signOnName').send_keys(parser.get('geral', 'T24_USERNAME'))
    driver.find_element_by_name('password').send_keys("AAbb124")
    driver.find_element_by_class_name('sign_in').click()
    print('T24 Iniciado com sucesso.')
    return driver

def open_prepaid_card(driver):
    sleep(10)
    driver.switch_to.frame(driver.find_element(By.XPATH, "//frame[contains(@id, 'menu')]"))
    driver.find_element(By.XPATH, "//span[contains(text(), 'Menu Service & Tel Team Leader')]").click()
    driver.find_elements_by_xpath("//span[contains(text(), 'Cartao Pre-pago')]")[0].click()
    driver.find_elements_by_xpath("//a[contains(text(), 'Load Prepaid Card by Dr Customer Account')]")[0].click()
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    driver.maximize_window()

def fill_form(driver, data, debit_account, description):
    amount = str(data['valor_carregamento']).strip()
    card_formated = str(data['numero_cartao']).replace(' ', '').replace(str(data['numero_cartao'])[5:12], '**********')
    msg = 'Iniciando processo para o cartão: {} Montante: {}'.format(card_formated, data['valor_carregamento'])
    if not len(str(debit_account)) == 13:
        return [False, 'Invalid Account: {}'.format(debit_account)]

    print(msg)
    #print(int(float(dd)))
    # order inputs
    input_debit = driver.find_element_by_id("fieldName:DEBIT.ACCT.NO")
    input_debit.clear()
    input_debit.send_keys(int(debit_account)) # type debit account number
    sleep(1)
    input_amount = driver.find_element_by_id("fieldName:DEBIT.AMOUNT")
    input_amount.clear()
    input_amount.send_keys(amount) # type amount
    sleep(1)
    # customer inputs
    input_beneficiary = driver.find_element_by_id("fieldName:DEPOSITOR")
    input_beneficiary.clear()
    input_beneficiary.send_keys("{}".format(str(data['nome_produtor']).strip())) # type customer name
    sleep(1)
    input_card_number = driver.find_element_by_id("fieldName:CARD.NUM.FULL")
    input_card_number.clear()
    input_card_number.send_keys(int(str(data['numero_cartao']).replace(' ', '').strip())) # type customer card number
    sleep(1)
    input_desc = driver.find_element_by_id("fieldName:PAYMENT.DETAILS:1")
    input_desc.clear()
    input_desc.send_keys(description) # type description
    sleep(1)
    # driver.find_element_by_xpath("//img[@title='Validate a deal']").click()
    sleep(1)
    checkError = check_exists_by('errorText', 1, driver)
    if checkError[0]:
        return [False, checkError[1]]
    else:
        try:
            driver.implicitly_wait(5)
            driver.find_element_by_xpath("//img[@title='Commit the deal']").click()
            driver.implicitly_wait(45)
        except: pass
        try:
            driver.find_element(By.XPATH, "//a[contains(text(), 'Accept Overrides')]").click()
            print('Accept Overrides.')


        except: pass
        print('Carregamento submetido...')
        checkError = check_exists_by('OVE1', 0, driver)
        if checkError[0]:
            return [False, checkError[1]]
        else:
            try:
                message = driver.find_element_by_xpath("//td[@class='message']").text
                if 'Complete' in message:
                    print(message)
            except Exception as e:
                logger.error('Erro ao capturar mensagem: ' + str(e))
            print('O cartão {} foi carregado com sucesso.'.format(card_formated))
            driver.find_element_by_xpath("//img[@title='New Deal']").click()
    return [True, message] #Transaction commited

def check_exists_by(name, type, driver):
    driver.implicitly_wait(5)
    element = None
    try:
        element = driver.find_element_by_class_name(name) if type == 1 else driver.find_element_by_id(name)
        logger.error('Alguns erros foram encontrados durante o carregamento.')
        logger.error(element.text)
        driver.implicitly_wait(10)
    except NoSuchElementException:
        print('> Sem errors de validação.')
        driver.implicitly_wait(10)
        return [False, '']
    return [True, element.text]

def make_payment(data, file, mail):
    #splunkConfig()
    driver = init_t24()
    open_prepaid_card(driver)
    file_excel = path.join(MAIN_PATH, file).replace('csv', 'xlsx')
    data_copy = data.copy()
    data_copy = data_copy.loc[data_copy['estado'] == 'Pending']
    row_order = data.loc[(data['numero_cartao'] == 0) | (data['numero_cartao'] == '0')]
    debit_account = str(row_order['lfc']).strip()[2:].strip()[0:13]
    description = str(row_order['valor_carregamento']).strip()

    is_processed = False
    total = len(data_copy) - 1
    success = 0
    error = 0

    for index, row in data_copy.iterrows():
        card = str(row['numero_cartao']).strip().replace(' ', '').strip()
        if len(card) >= 3:
            res = fill_form(driver, row, debit_account, description)
            if res[0]:
                is_processed = True
                data.loc[index, 'estado'] = res[1]
                success = success + 1
            else:
                error = error + 1
                data.loc[index, 'estado'] = res[1]
    to_send = path.join('{}\Completed'.format(MAIN_PATH), file)
    shutil.move(file_excel, to_send)

    if is_processed:
        data.to_excel(to_send, index=False)
        send_email(to_send, debit_account, total, success, error, mail)
    driver.close() # close prepaid form
    driver.quit() # kill driver scope
    #try: splunkFinish()
    #except: pass

def get_dataframe(filename):
    file = path.join(MAIN_PATH, filename)
    file_excel = file.replace('csv', 'xlsx')
    if not path.exists(file_excel):
        try:
            d_frame = pd.read_csv(file)
            d_frame['estado'] = 'Pending'
            print('Ajustando as colunas no report...')
            d_frame.columns = ['lfc', 'nome_produtor', 'entidade', 'numero_cartao', 'valor_carregamento', 'descritivo', 'estado']
            d_frame = d_frame.drop(d_frame.index[[0]])
            d_frame['estado'] = d_frame['estado'].fillna("Pending")
            d_frame.to_excel(file_excel, index=False)
            os.remove(file)
            print('> Ajuste completado.')
        except Exception as e:
            shutil.move(file, path.join(MAIN_PATH+'\Error', filename))
            logger.exception(str(e))
            return pd.DataFrame(columns=['lfc'])
    else:
        d_frame = pd.read_excel(file_excel, engine="openpyxl", dtype=str)
        if not 'estado' in d_frame.columns:
            d_frame = d_frame.drop(d_frame.index[[0]])
            try:
                d_frame.columns = ['lfc', 'nome_produtor', 'entidade', 'numero_cartao', 'valor_carregamento', 'descritivo', 'estado']
            except:
                d_frame.columns = ['lfc', 'nome_produtor', 'entidade', 'numero_cartao', 'valor_carregamento', 'descritivo']
            d_frame = d_frame[d_frame.numero_cartao.notnull()]
            d_frame['estado'] = 'Pending'
            d_frame.to_excel(file_excel, index=False)
    return d_frame

def read_files():
    total = filter(os.listdir(MAIN_PATH), '*.xlsx')
    return total if len(total) > 0 else filter(os.listdir(MAIN_PATH), '*.csv')

def get_reports():
    dt_approvers = pd.read_excel('approvers.xlsx')
    dt_inputters = pd.read_excel('inputters.xlsx')
    subject = 'Carregamento'
    print('Searching file in Outlook by subject: {}'.format(subject))
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    mail = None
    for message in messages:
        if (message.Subject == subject) and message.Unread:
            if message.SenderEmailType == 'EX':
                email = str(message.Sender.GetExchangeUser().PrimarySmtpAddress)
            else:
                email = str(message.SenderEmailAddress)
            if 'standardbank.co.mz' in email:
                print('Prepaid Request found in *email, approval required...')
                dt_inputters = dt_inputters[dt_inputters['NAMES'] == str(message.Sender)]
                if not dt_inputters.empty:
                    for attachment in message.Attachments:
                        file_name = str('Carregamento EDM - ') + str(uuid.uuid4()).replace('-', '')
                        file = os.path.join(MAIN_PATH + '\Pending',  file_name + '.xlsx')
                        attachment.SaveAsFile(file)
                        send_approval_email(file_name, file, str(message.Sender))
                        message.Unread = False
                else:
                    print(str(message.Sender))
                    print('Inputter from EDM *{}* not recognized.'.format(str(message.Sender)))
                    message.Unread = False
            else:
                for attachment in message.Attachments:
                    extension = path.splitext(str(attachment))[1][1:].strip()
                    if message.Subject == subject and message.Unread and extension == 'xlsx' or extension == 'csv':
                        print('Foi encontrado ficheiro no *e-mail para o processamento: {}'.format(attachment))
                        attachment.SaveAsFile(os.path.join(MAIN_PATH, str(attachment)))
                        message.Unread = False
                        mail = message
                        break

        elif message.Unread and 'RE: Carregamento EDM' in message.Subject:
            dt_approvers = dt_inputters[dt_inputters['NAMES'] == str(message.Sender)]
            if not dt_approvers.empty:
                print(f'Checking... prepaid request approved by *{str(message.Sender)}')
                list_pending = os.listdir(MAIN_PATH + '\Pending')
                for file_pending in list_pending:
                    pending_f = file_pending.replace('.xlsx', '')
                    if pending_f in str(message.Subject):
                        print(f'The file {pending_f} has been approved by: {str(message.Sender)}.')
                        shutil.move(os.path.join(MAIN_PATH + '\Pending',  file_pending), os.path.join(MAIN_PATH, file_pending))
                        message.Unread = False
                        mail = message
                        break
            else:
                print(str(message.Sender))
                print('Approver from EDM *{}* not recognized.'.format(str(message.Sender)))
                message.Unread = False
    return [mail, read_files()]

files = read_files()
is_checkinemail = True

while True:
    now = datetime.datetime.now()
    if now.hour < 20 and now.hour > 6:
        if not is_checkinemail and len(files) > 0: # already exists files to be processed
            print('File found in *folder for processing: {}'.format(files[0]))
            d_frame = get_dataframe(files[0])
            if not d_frame.empty: make_payment(d_frame, files[0])
            else: is_checkinemail = False
        else:
            is_checkinemail = True
            email_report = get_reports()
            files = email_report[1]
            if len(files):
                d_frame = get_dataframe(files[0])
                if not d_frame.empty: make_payment(d_frame, files[0], email_report[0])
                else:
                    if email_report[0].Reply():
                        newMail = email_report[0].Reply()
                        newMail.HTMLBody = parser.get('geral', 'MAIL_NOT_PROCESSED')
                        print('ERROR FILE NOT PROCESSED')
                        newMail.Send()
            else: print('>>> No files were found for processing.')
        print('')
        print(f'{now} | Waiting for new processes...')
        sleep(10)
    else:
        print(f'{now} | Cut of time.')
        sleep(1800)