#!/usr/bin/env python3
# -*- encoding: utf-8 -*-
"""
Created on Fri Jun 30, 2023  08:52:15

@author: a248433
"""
import os
import shutil
import shutil
import win32com.client
from string import Template
from win32com.client import Dispatch, constants
import time

from _helpers import (
	login_user, logout_user, get_credentials, get_driver, 
	# display_text_std, 
	find_element_by_xpath_and_click,
	common_handler, handle_user_auth, change_branch,
	change_to_last_window, find_input_field_and_fill_value,
	INDEXES, ClientesSociedadePrenchido,
	ContasSociedadeConstituidaPr, ClientCIFs,
)
from inhibit_client import process_action
from create_letter import create_letter

INDEXES = list(range(1, 100))

template = Template(
    '<b>Caro cliente</b><p>Queira por favor confirmar respondendo a este e-mail com a palavra "Aprovado" caso seja legítimo ou ignorar caso não seja.</p>'
)

def send_request_approval_email(subject, sender):
    global INDEXES
    olMailItem = 0x0
    subject = subject.replace('_', ' ')
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.To = sender # 'anisio.marrima@standardbank.co.mz'
    newMail.HTMLBody = template.substitute(name=sender)
    # newMail.HTMLBody = f'<b>Caro cliente</b><p>Recebemos o ficheiro em anexo De: {sender}.</p><p>Queira por favor confirmar respondendo a este e-mail com a palavra "Aprovado" caso seja legítimo ou ignorar caso não seja.</p>'
    # newMail.Attachments.Add(Source=file)
    newMail.Send()
    text = f'{INDEXES.pop(0)}. Bot is sending email to {sender!r} requesting approval.   {chr(482)}'
    # display_text_std(text)

def format_filename(filename):
    for char in '!"#$%&\'()*+,-./:;<=>?@[\\]^`{|}~':
        filename = filename.replace(char, '')
    # filename = filename.replace(' ', '_', -1)
    return filename.strip()

def create_folder(filename):
    global INDEXES
    for char in '!"#$%&\'()*+,-./:;<=>?@[\\]^`{|}~':
        filename = filename.replace(char, '')
    # filename = filename.replace(' ', '_', -1).strip()
    if not os.path.exists(filename):
        try:
            os.system(f'mkdir "{filename}"')
        except Exception as e:
            print(e)

def create_folder_if_not_exist(foldername):
    if not os.path.exists(foldername):
        os.mkdir(foldername)
        text = f'{INDEXES.pop(0)}. Bot is creating folder: {foldername!r}.   {chr(482)}'
        # display_text_std(text)

def check_if_folders_exist():
    create_folder_if_not_exist('Clientes')
    os.chdir('./Clientes/')
    # create_folder_if_not_exist('pending')
    # create_folder_if_not_exist('done')
    # create_folder_if_not_exist('approved')
    for foldername in ['pending', 'approved', 'done']:
        create_folder_if_not_exist(foldername)
    os.chdir('./pending/')

# if os.path.exists('Clientes'):
#     shutil.rmtree('Clientes')


check_if_folders_exist()

def read_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    text = f'{INDEXES.pop(0)}. Bot is connecting to Outlook.   {chr(482)}'
    # display_text_std(text)
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    text = f'{INDEXES.pop(0)}. Bot is fetching emails.   {chr(482)}'
    # display_text_std(text)

    # if os.path.exists('Clientes'):
    #     shutil.rmtree('Clientes')
    #     text = f'{INDEXES.pop(0)}. Bot is removing "Clients" folder.   {chr(482)}'
        # display_text_std(text)

    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        # print(account.DisplayName)
        folders = inbox.Folders
        # print('NR OF FOLDERS:', len(folders))

        for folder in folders:
            if folder.Name == 'Inbox':
                messages = folder.Items
                for message in messages:
                    if message.Unread:
                        subject = message.Subject
                        # print(f'Subject: {message.Unread}')
                        if message.SenderName == 'Marrima, Anisio': # 'Cipriano, Januario J':  # Rodrigues, Joao':
                            
                            # print(template.substitute(name=message.SenderName))
                            print('READING MESSAGE')
                            subject = format_filename(subject)
                            folder_name = subject
                            if message.Subject.startswith('RE:'):
                                os.chdir('C:/Users/a248433/Documents/SB Mozambique/EDO/Robotics/Bots/Automatização do Processo Abertura de contas Corporate/Clientes/pending')
                                print(subject, ': ', message.SenderName)
                                # print(f'{folder_name!r}')
                                print(os.getcwd())
                                print(os.listdir())
                                message.Unread = False
                                # print(message.SenderName.upper(), ':', message.Subject.upper())
                                subject_name = message.Subject.split('RE: ')[-1].strip()
                                # subject_name = ' '.join(subject_name)
                                print(subject_name in os.listdir())
                                if os.path.exists(f'./{subject_name}'):
                                    print(subject_name, 'EXISTS')
                                    shutil.move(os.path.join(os.getcwd() + f'\\{subject_name}'),
                                                '../approved')
                                    text = f'{INDEXES.pop(0)}. Bot found {subject_name!r} folder inside "peding" folder, moving it to "approved" folder.   {chr(482)}'
                                    # display_text_std(text)
                                    os.chdir('../approved')
                                    if os.path.exists(f'./{subject_name}/data.xlsx'):
                                        process_action(subject_name)                                        
                                    else:
                                        text = f'{INDEXES.pop(0)}. Bot did not find folder {subject_name!r}.   {chr(482)}'
                                        # display_text_std(text)
                                else:
                                    print(f'{subject_name!r}', 'DOES NOT EXISTS')
                                    text = f'{INDEXES.pop(0)}. Bot did not find folder {subject_name!r}.   {chr(482)}'
                                    # display_text_std(text)
                                os.chdir('C:/Users/a248433/Documents/SB Mozambique/EDO/Robotics/Bots/Automatização do Processo Abertura de contas Corporate/Clientes/pending')
                            else:
                                try:
                                    create_folder(subject)
                                    text = f'{INDEXES.pop(0)}. Bot found {len(message.Attachments)} attachments inside the email.   {chr(482)}'
                                    # display_text_std(text)
                                    # send_request_approval_email(subject, 'januario.cipriano@standardbank.co.mz')
                                    # send_request_approval_email(subject, 'anisio.marrima@standardbank.co.mz')
                                    for attachment in message.Attachments:
                                        file = attachment.FileName.split('.') 
                                        # print(file)
                                        if len(file) == 1:
                                            file_ext = 'pdf'
                                            filename = file[0] + '.' + file_ext
                                        else:
                                            file_ext = file[-1]
                                            filename = ' '.join(file[:-1]) + '.' + file_ext
                                        
                                        text = f'{INDEXES.pop(0)}. Bot found {filename!r} attachments inside the email.   {chr(482)}'
                                        # display_text_std(text=f'CURRENT FOLDER: {os.getcwd()!r}')

                                        if os.path.exists(subject):
                                            try:
                                                # print(f'{subject} HAS FILES = {filename}')
                                                attachment.SaveAsFile(os.path.join(os.getcwd() + f'\\{subject}', filename))
                                                text = f'{INDEXES.pop(0)}. Bot is saving {filename!r} attachment into {subject!r} in "pending" folder.   {chr(482)}'
                                                # display_text_std(text)
                                            except Exception as e:
                                                print(e.args)
                                                # pass
                                        else:
                                            print(f'{subject} does not exist')
                                        # send_approval_email(file_name, file, str(message.Sender))
                                    # os.chdir('../')
                                    message.Unread = False
                                    # send_request_approval_email(subject, 'januario.cipriano@standardbank.co.mz')
                                    send_request_approval_email(subject, 'anisio.marrima@standardbank.co.mz')
                                except Exception as e:
                                    print(e.args)
                    folders = inbox.Folders   
while True:
    read_emails()
    time.sleep(17)