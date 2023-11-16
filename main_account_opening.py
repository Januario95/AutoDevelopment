# -*- coding: utf-8 -*-
"""
Created on Tue Jul 16 10:53:55 2023

@author: C835910
"""

import logger_config
import shutil
from time import sleep
import os
import pandas as pd
import t24_account_opening as t24

logger = logger_config.get_logger('Corporate Account Opening')
main_folder = "C://Users/C835910/Documents/Corporate_Account_Opening/temp"

while True:
    approved_folder = main_folder + '/approved'
    done_folder = main_folder + '/done'
    rows = os.listdir(approved_folder)
    main_excel = ''
    working_folder = ''
    data = list()

    for row in rows:
        working_folder = row
        main_excel = approved_folder + '/' + row + '/data.xlsx'
        sheets = pd.ExcelFile(main_excel).sheet_names
        sheets = list(filter(lambda x: ("CUS_" in x), sheets))
        for sheet in sheets:
            data.append(pd.read_excel(main_excel, sheet_name=sheet))
        if len(data) > 0: break

    if len(data) > 0:
        for cus in data:
            driver = t24.init_t24(True, True)
            customer = t24.create_customer(driver, cus)
            driver.quit()

            driver = t24.init_t24(False, True)
            logger.info('Logado como Authorizer...')
            t24.authorize_customer(driver, customer)
            driver.quit()

            cifs = pd.read_excel(main_excel, engine='openpyxl', sheet_name='CIF_CUS')
            cifs.append([cus.iloc[2][1], 'A', customer])
            writer = pd.ExcelWriter(main_excel, engine = 'xlsxwriter')
            cifs.to_excel(writer, sheet_name = 'CIF_CUS')
            writer.close()

        os.makedirs(done_folder+'/'+working_folder)
        files = os.listdir(approved_folder + '/' + working_folder)
        for file in files: shutil.move(approved_folder + '/' + working_folder + '/' + file, done_folder+'/'+working_folder)
        os.rmdir(approved_folder + '/' + working_folder)

    sleep(60)
