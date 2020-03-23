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
from random import randrange

from pathlib import Path
import os

def convertToCSV (df, file, gender = 'Neutral'):    
    try:
        export_file_path = getExportFilePath(file) 

        # print(df.head())
        if not dropUnnecessaryColumns(df):
            return False

        dropUnnecessaryEntries(df)

        formatTheName(df)

        # print(df.head())
        
        exportCustomers(df, export_file_path, gender)
        return True
    except Exception as e:
        print(e)
        return False

def dropUnnecessaryColumns(df):
    # edit the dataframe here before saving it to csv.
    # change column names to fit google contacts format.
    # the try except part is to deal with different excel file formats.
    # drop unnecessary columns from the data frame.
    try:
        df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location", "CARDS", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "EXTRA"]
        df = df.drop(["Receipts", "Amount", "Delivery", "Active", "Id", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "CARDS", "EXTRA"], axis=1)
    except:
        try:
            df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location", "CARDS", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY"]
            df = df.drop(["Receipts", "Amount", "Delivery", "Active", "Id", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "CARDS"], axis=1)
        except:
            try:
                df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location"]
                df = df.drop(["Receipts", "Amount", "Delivery", "Active", "Id"], axis=1)
            except:
                try:
                    df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch"]
                    df = df.drop(["Receipts", "Amount", "Id"], axis=1)
                except:
                    try:
                        df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area"]
                        df = df.drop(["Receipts", "Amount", "Id"], axis=1)
                    except:
                        try:
                            df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value"]
                            df = df.drop(["Id"], axis=1)
                        except:
                            return False
    return True

def insertRequiredColumns(df):
    # insert the "phone type" column and assign every value to "Mobile" before each phone number in the data frame.
    df.insert(4, 'Phone 1 - Type', 'Mobile', allow_duplicates = True)
    df.insert(1, 'Name', '', allow_duplicates = True)
    df.insert(1, 'Nickname', '', allow_duplicates = True)

def dropUnnecessaryEntries(df):
    # convert any entry with empty phone number to np.nan for  dropping.
    df['Phone 1 - Value'].replace('', np.nan, inplace=True)
    # drop completely empty entries.
    df.dropna(subset=['Phone 1 - Value'], inplace=True)
    # drop duplicates
    df.drop_duplicates()
        # drop last row
    df.drop(df.tail(1).index, inplace=True) 

def formatTheName(df):
    # add the middle name to the first name.
    df['Name'] = df['Given Name'] + ' ' + df['Maiden Name'] + ' ' + df['Family Name']
    df['Nickname'] = df['Family Name']
        
    m = df['Maiden Name'].notna() & df['Maiden Name'].ne('')
    df.loc[m, 'Given Name'] += ' ' + df.loc[m, 'Maiden Name']

def exportCustomers(df, export_file_path, gender):
    # save the ready google contacts file.
    gender = str(gender).upper()
    if gender != 'NEUTRAL':
        df = filterCustomers(df, gender)
        
    df.to_csv(export_file_path, index = None, header=True, index_label = True, encoding='utf-8')

def filterCustomers(df, gender):
    return df[df['Gender'] == gender]

def getFileId():
    # renew the now variable in case of leaving the app open for days.
    now = datetime.datetime.now()
    return f'{now.strftime("%A")} {str(now.day)}-{str(now.month)}--{randrange(1000)}'
    

def getBranchName(original_filename):
    branch = ''
    try:
        branch = original_filename.split('loyality ph ')[1]
        print(branch)
        branch = branch.split('.xlsx')[0]
        print(branch)
        branch = f'Adam {branch}.csv'
    except:
       branch = 'default.csv'

    return branch

def getExportFilePath(file):
    try:
        file_name =  getFileId()
        file_name = f'{file_name}=={getBranchName(file)}'
        _dir = Path(os.path.dirname(file)) ## directory of file
        export_file_path = _dir / str("Google Contacts " + file_name)
        return export_file_path
    except Exception as e:
        print(e)
        return e
