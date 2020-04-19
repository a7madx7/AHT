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


def convertToCSV(df, file, gender = 'Neutral', mode = 'separate'):    
    try:
        export_file_path = getContactsExportFilePath(file) 
        
        # print(df.head())
        if not dropUnnecessaryColumns(df):
            return False

        dropUnnecessaryEntries(df)

        formatTheName(df)

        gender = str(gender).upper()
        if gender != 'NEUTRAL':
            df = filterCustomers(df, gender)
        
        # print(df.head())

        if mode == 'separate':
            # if the user wants every branch in a single file.
            exportCustomers(df, export_file_path, gender)
            return True
        else:
            return df
    except Exception as e:
        print(e)
        return False

def executeMergeToCSV(export_file_path, dfs, gender):
    finalDf = mergeBranchesAndDropDuplicateCustomers(dfs)
    exportCustomers(finalDf, export_file_path, gender)

def dropUnnecessaryColumns(df):
    # edit the dataframe here before saving it to csv.
    # change column names to fit google contacts format.
    # the try except part is to deal with different excel file formats.
    # drop unnecessary columns from the data frame.
    try:
        df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location", "CARDS", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "EXTRA"]
        df.drop(["Receipts", "Amount", "Delivery", "Active", "Id", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "CARDS", "EXTRA"], axis=1, inplace = True)
        # print(df.head())
        print('Done dropping columns on first try')
    except:
        try:
            df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location", "CARDS", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY"]
            df.drop(["Receipts", "Amount", "Delivery", "Active", "Id", "CREATED AT", "CREATED BY", "MODIFIED AT", "MODIFIED BY", "CARDS"], axis=1, inplace = True)
            print('Done dropping columns on second try')
        except:
            try:
                df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Location"]
                df.drop(["Receipts", "Amount", "Delivery", "Active", "Id"], axis=1, inplace = True)
                print('Done dropping columns on third try')
            except:
                try:
                    df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch"]
                    df.drop(["Receipts", "Amount", "Id"], axis=1, inplace = True)
                    print('Done dropping columns on fourth try')
                except:
                    try:
                        df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area"]
                        df.drop(["Receipts", "Amount", "Id"], axis=1, inplace = True)
                        print('Done dropping columns on fifth try')
                    except:
                        try:
                            df.columns = ["Id", "Gender", "Given Name", "Maiden Name", "Family Name", "Phone 1 - Value"]
                            df.drop(["Id"], axis=1, inplace = True)
                            print('Done dropping columns on sixth try')
                        except:
                            print('Failed to drop columns')
                            return False
    return True

def mergeBranchesAndDropDuplicateCustomers(dfs):
    return pd.concat(dfs).drop_duplicates(subset= 'Phone 1 - Value').reset_index(drop=True)

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
        branch = branch.split('.xlsx')[0]
        branch = f'Adam {branch}'
    except:
       branch = 'Default'

    return branch

def getContactsExportFilePath(file, merge = False):
    try:
        file_name =  getFileId()
        if merge:
            file_name = f'Merged Contacts-{file_name}.csv'
        else:
            file_name = f'{getBranchName(file)} Contacts-{file_name}.csv'
        _dir = Path(os.path.dirname(file)) ## directory of file
        export_file_path = _dir / str(file_name)
        return export_file_path
    except Exception as e:
        print(e)
        return e    