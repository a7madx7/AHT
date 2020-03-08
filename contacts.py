"""
    CopyRight Dr. Ahmad Hamdi Emara 2020
    Adam Medical Company
    This module is intended for Converting ABC System contacts file to Google Contacts File.
    The contacts file must be in the following format:

    CUST. NO	GENDER	FIRST	MIDDLE	FAMILY	MOBILE NO.	RECEIPTS #	AMOUNT	CITY	AREA	STREET	MAIN BRANCH	LANGUAGE	DELIVERY	ACTIVE	CREDIT

    With no more or no less columns.
"""
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# IMPORTS REGION.
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
import pandas as pd
import numpy as np

import datetime

from pathlib import Path
import os

def convertToCSV (df, file):    
    try:
        export_file_path = get_export_file_path(file) 

        # edit the dataframe here before saving it to csv
        # change column names to fit google contacts format
        df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location"]

        # convert any entry with empty phone number to np.nan for later dropping.
        df['Phone 1 - Value'].replace('', np.nan, inplace=True)

        # drop unnecessary columns from the data frame.
        df = df.drop(["Receipts", "Amount", "Delivery", "Active", "Id"], axis=1)

        # insert the "phone type" column and assign every value to "Mobile" before each phone number in the data frame.
        df.insert(4, 'Phone 1 - Type', 'Mobile', allow_duplicates = True)

        df.insert(1, 'Name', '', allow_duplicates = True)
        df.insert(1, 'Nickname', '', allow_duplicates = True)
        
        # drop completely empty entries.
        df.dropna(subset=['Phone 1 - Value'], inplace=True)

        # add the middle name to the first name.
        df['Name'] = df['Given Name'] + ' ' + df['Maiden Name'] + ' ' + df['Family Name']
        df['Nickname'] = df['Family Name']
        
        m = df['Maiden Name'].notna() & df['Maiden Name'].ne('')
        df.loc[m, 'Given Name'] += ' ' + df.loc[m, 'Maiden Name']

        # drop duplicates
        df.drop_duplicates()

        # drop last row
        df.drop(df.tail(1).index, inplace=True) 

        print(df.head())
        # save the ready google contacts file.
        df.to_csv(export_file_path, index = None, header=True, index_label = True, encoding='utf-8')
        return True
    except Exception as e:
        print(e)
        return False

def get_export_file_path(file):
    try:
        # renew the now variable in case of leaving the app open for days.
        now = datetime.datetime.now()

        file_name = now.strftime("%A") + ' ' + str(now.day) + '-' + str(now.month)
        _dir = Path(os.path.dirname(file)) ## directory of file
        export_file_path = _dir / str("Google Contacts " + file_name + '.csv')
        return export_file_path
    except Exception as e:
        print(e)
        return e
