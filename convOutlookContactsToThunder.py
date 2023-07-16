# This program converts the .csv format of Microsoft Outlook into the same format of Thunderbird.
# Author: Sherman Chen
# Date of creation: 20230715
# Last update: 20230716

import pandas as pd
import numpy as np

INFILE = 'contactsOutlook.csv'
OUTFILE = 'outContactsThunderbird.csv'


def importContacts():
    """ Read csv into df. Remove columns not needed.

    Paramaters:
        None

    Returns:
        df2 (dataFrame): contacts of Outlook
    """
    df = pd.read_csv(INFILE, encoding='utf_8', low_memory=False) # If there is encoding problem, use detectEncoding.py to find the encoding, and substitute it in this line.
    # df = pd.read_csv(INFILE, encoding='Windows-1252', low_memory=False)
    df2 = df.loc[:, ['First Name','Last Name','Company','Department','Job Title','Business Street','Business City','Business State','Business Postal Code','Business Fax','Business Phone','Business Phone 2','Home Phone','Mobile Phone','E-mail Address','E-mail 2 Address','E-mail 3 Address','Notes']]
    return df2


def createEmptyThundDF():
    """ Create an empty dataframe which has all the columns of Thunderbird contact csv. Add one row with the value NaN for every column.

    Parameters:
        None

    Returns:
        dfo (dataFrame): a one-row dataframe which has all the columns of Thunderbird contact csv.
    """
    thundColNames = ['First Name', 'Last Name', 'Display Name', 'Nickname', 'Primary Email', 'Secondary Email', 'Screen Name', 'Work Phone', 'Home Phone', 'Fax Number', 'Pager Number', 'Mobile Number', 'Home Address', 'Home Address 2', 'Home City', 'Home State', 'Home Postal Code', 'Home Country', 'Work Address', 'Work Address 2', 'Work City', 'Work State', 'Work Postal Code', 'Work Country', 'Job Title', 'Department', 'Organization', 'Web Page 1', 'Web Page 2', 'Birth Year', 'Birth Month', 'Birth Day', 'Custom 1', 'Custom 2', 'Custom 3', 'Custom 4', 'Notes']
    dfo = pd.DataFrame(columns = thundColNames)
    numCol = len(dfo.columns)
    rowOfNaN = [np.nan] * numCol
    dfo.loc[0, :] = rowOfNaN
    return dfo


def convert(df2):
    """ Change Outlook contact column names into Thunderbird contact column names.

    Parameters:
        df2 (dataFrame): contacts of Outlook

    Returns:
        df2 (dataFrame): Outlook contacts in which column names have been changed into Thunderbird contact column names.
        """
    columnMapping = {
            'E-mail Address': 'Primary Email',
            'E-mail 2 Address': 'Secondary Email',
            'Business Phone': 'Work Phone',
            'Business Fax': 'Fax Number',
            'Mobile Phone': 'Mobile Number',
            'Business Street': 'Work Address',
            'Business City': 'Work City',
            'Business State': 'Work State',
            'Business Postal Code': 'Work Postal Code',
            'Business Country/Region': 'Work Country',
            'Company': 'Organization',
            }
    df2.rename(columns=columnMapping, inplace=True)
    return df2


def produceOutDF(df2, dfo):
    """ Concatenate Outlook df with Thunderbird one-row df. Change the order of column names as per the column name order of Thunderbird contacts.

    Parameters:
        df2 (dataFrame): Outlook contacts in which column names have been changed into Thunderbirdcontact column names.
    Returns:
        dfo3 (dataFrame): produced Thunderbird contact dataframe. The order of the column names has been changed as per Thunderbird contacts.
    """
    dfo2 = pd.concat([df2, dfo]) # After concat, index of last row is '0'
    new_order = ['First Name', 'Last Name', 'Display Name', 'Nickname', 'Primary Email', 'Secondary Email', 'Screen Name', 'Work Phone', 'Home Phone', 'Fax Number', 'Pager Number', 'Mobile Number', 'Home Address', 'Home Address 2', 'Home City', 'Home State', 'Home Postal Code', 'Home Country', 'Work Address', 'Work Address 2', 'Work City', 'Work State', 'Work Postal Code', 'Work Country', 'Job Title', 'Department', 'Organization', 'Web Page 1', 'Web Page 2', 'Birth Year', 'Birth Month', 'Birth Day', 'Custom 1', 'Custom 2', 'Custom 3', 'Custom 4', 'Notes']
    dfo3 = dfo2.loc[:, new_order]
    return dfo3


def exportContacts(dfo3):
    """ Export the produced contacts dataframe into a csv file.

    Parameters:
        dfo3 (dataFrame): produced Thunderbird contact dataframe. The order of the column names has been changed as per Thunderbird contacts.

    Returns:
        None
    """
    dfo3.to_csv(OUTFILE, index=False)


def main():
    df2 = importContacts()
    dfo = createEmptyThundDF()
    df2 = convert(df2)
    dfo3 = produceOutDF(df2, dfo)
    exportContacts(dfo3)


if __name__ == '__main__':
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    main()
