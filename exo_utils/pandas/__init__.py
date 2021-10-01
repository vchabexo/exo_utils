import re
import unidecode
import pandas as pd
import logging
import time
import datetime

from exo_utils.database import insertData, buildTable

def removeAccents(input_str):
    """Removes accents from a strin
        intput_str :(str) string you want to format
    Return the string without accent
    """
    output_str = unidecode.unidecode(input_str)
    return (output_str)

def removeSpaces(input_str):
    """Removes spaces from a strin
        intput_str :(str) string you want to format
    Return the string without spaces
    """
    wordsList = input_str.split(" ")
    if len(wordsList) > 1:
        output_str = wordsList[0]
        for i in range(1, len(wordsList)):
            output_str += wordsList[i].capitalize()
    else:
        output_str = input_str
    return (output_str)

def removeSpecialChar(input_str):
    """Replace anny non alpha numerical char with _
        intput_str :(str) string you want to format
    Return the string without special char
    """
    output_str = re.sub(r"\$", "_dollar_", input_str)
    output_str = re.sub(r"\W+", "_", output_str)
    return (output_str)


def columnNameFormating(name):
    """Normalizes a string. It will put lettre in lower, remove accent spaces and special char
        name :(str) name you want to format
    Return normalized name 
    """
    name = name.lower()
    name = removeAccents(name)
    name = removeSpaces(name)
    name = removeSpecialChar(name)
    if name in ["index"]:
        name = name + '_'
    return (name)

def normalizeDfColumns(df, inplace=False):
    """Normalizes df columns name
        df :(p.DataFrame) df to normalize
        inplace :(bool) If true modify df inplace else return a new df
    Return True if inplace else df with normalized columns names
    """
    newColumnNamesDict = {oldName: columnNameFormating(oldName) for oldName in df.columns}
    if inplace:
        df.rename(columns=newColumnNamesDict, inplace=True)
        return True
    else:
        newDf = df.rename(columns=newColumnNamesDict, inplace=False)
        return newDf


def importFile(pathToFile, fileType='excel',  filterColumns=None, notToKeepColumns=None, dateColumns=None, datetimeColumns=None, timeColumns=None, addImportDate=False, **kwds):
    """Import an excel or csv file to a pd dataframe
        pathToFile      : (str) path to excel sheet
        fileType        : (str) csv or excel - type of file to import (default excel)
        filterColumns   : ([str]) name of field to aply a filter on. All record with no value in this field won't be put in the db
        notToKeepColumns: ([str]) list a columsnot to transfert in the db
        dateColumns     : ([str]) list of columns that are date
        datetimeColumns : ([str]) list of columns that are datetime
        addImportDate   : (bool) Add a column with import date if true
        timeColumns     : ([str]) list of columns that are time only. Must be in format hh:mm or hh:mm:ss
        **kwds optional keyword argument can be pass to pd.read_execl
    Returns a pd DataFrame with data from excel file
    """
    t0 = time.time()
    #import data from excel or csv
    if fileType == 'excel':
        df = pd.read_excel(pathToFile, **kwds)
        
    elif fileType == 'csv':
        df = pd.read_csv(pathToFile, **kwds)

    #formating columns name
    normalizeDfColumns(df, inplace=True)

    #drop rows in filterColumn with emty data
    if not filterColumns is None:
        filterColumns = list(map(lambda name: columnNameFormating(name), filterColumns))
        for filterColumn in filterColumns:
            df = df[df[filterColumn].notnull()]
    
    #drop columns
    if not notToKeepColumns is None:
        notToKeepColumns = list(map(lambda name: columnNameFormating(name), notToKeepColumns))
        df.drop(notToKeepColumns, axis=1, inplace=True)
    
    #format datetime columns
    if not datetimeColumns is None:
        datetimeColumns = list(map(lambda name: columnNameFormating(name), datetimeColumns))
        for column in datetimeColumns:
            df[column] = df[column].apply(lambda date: pd.to_datetime(date))    

    #format date columns
    if not dateColumns is None:
        dateColumns = list(map(lambda name: columnNameFormating(name), dateColumns))
        def toDate(date):
            if pd.isna(date):
                return pd.NaT
            elif (isinstance(date, datetime.datetime)):
                return date.date()
            elif (isinstance(date, datetime.date)):
                return date
            else:
                return pd.to_datetime(date).date()
        for column in dateColumns:
            df[column] = df[column].apply(lambda date: toDate(date))    
            
    #format time columns
    if not timeColumns is None:
        timeColumns = list(map(lambda name: columnNameFormating(name), timeColumns))
        def toTime(date):
            if pd.isna(date):
                return pd.NaT
            elif (isinstance(date, datetime.datetime)) or (isinstance(date, datetime.date)):
                return date.time()
            elif (isinstance(date, datetime.time)):
                return date
            else:
                date = date.replace("h", ":")
                date = date.replace("H", ":")
                if len(date)<=5:
                    return pd.to_datetime(date, format = "%H:%M" ).time()
                else:
                    pd.to_datetime(date, format = "%H:%M:%S" ).time()
            
        for column in timeColumns:
            df[column] = df[column].apply(lambda date: toTime(date))

    if addImportDate:
        df['importDate'] = datetime.date.today()

    t1 = time.time()
    logging.info(f"Successfuly imported file {pathToFile} - nb of reocrds: {len(df)} - duration: {t1 - t0:.2f}s")
    return df


def exportDfToMssql(df, server, database, schema, table, dropFirst=False, customTypes=None, indexAsPrimaryKey=False, SQLLogin=False, username=None, password=None):
    """Export data from df to a mssql db
        df                  : (pd.DataFrame) data frame to export
        server              : (str) server adress
        database            : (str) database name
        schema              : (str) name of table schame
        table               : (str) name of table
        dropFirst           : (bool) If true will rebuild table and erease all old data first. Default is False
        customTypes         : (dict) custom types for tqble colums. Key of the dict is name of column and value is MSSQL type
        indexAsPrimaryKey   : (bool) if True if index column is table primary key (default is False).
        SQLLogin            : (bool) Set sql connection protocol. If True use user/pwd connection if False is AD for user info connection. Delfault is False
        username            : (str) username used for connection (Only needed if SQLLogin=True)
        password            : (str) passeword used for conneciton (Only needed if SQLLogin=True)
    
    Returns void
    """
    t0 = time.time()
    #To avoid type conflict we must replace all na values with None
    df.fillna("None", inplace=True)
    df.replace({"None": None, pd.NaT: None}, inplace=True)

    #if drop first we rebuild df table

    if dropFirst:
        buildTable(
            df, server, database, schema, table, 
            customTypes=customTypes, indexAsPrimaryKey=indexAsPrimaryKey, SQLLogin=SQLLogin, username=username, password=password
            )
    #insertion of data in db
    insertData (
        df, server, database, schema, table, 
        customTypes=customTypes, indexAsPrimaryKey=indexAsPrimaryKey, SQLLogin=SQLLogin, username=username, password=password
        )
    t1 = time.time()
    logging.info(f"Successfuly exported {len(df)} line to ({server}) {database}.{schema}.{table} with dropFirst: {dropFirst} - duration: {t1 - t0:.2f}s")
    return

