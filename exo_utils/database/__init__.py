import pyodbc
import pandas as pd
import numpy as np


# use to convert pd types to sql typs
pdTypeToSQLTypeDict = {
    "datetime64[ns]": "DATETIME",
    "object"        : "VARCHAR(MAX)",
    "float64"       : "FLOAT",
    "int64"         : "INT"
}



def odbcConnect(server, database, SQLLogin=False, username=None, password=None):
    """ Return a pyodbc connection to the db
        server      : (str) server adress
        database    : (str) database name
        SQLLogin    : (bool) Set sql connection protocol. If True use user/pwd connection if False is AD for user info connection. Delfault is False
        username    : (str) username used for connection (Only needed if SQLLogin=True)
        password    : (str) passeword used for conneciton (Only needed if SQLLogin=True)
    """
    if SQLLogin:
        id_string = "Server={0}; Database={1}; UID={2}; PWD={3}".format(server, database, username, password)
    else:
        id_string = "Server={0}; Database={1}; Trusted_Connection=yes;".format(server, database)
    #depanding on your computer, you may have ODBC 13 or ODBC 17 driver.
    #We test both to make sure you can connect
    try:
        driver_string = "Driver={ODBC Driver 17 for SQL Server};"
        connection_string = driver_string + id_string
        cnxn = pyodbc.connect(connection_string)
    except:
        try:
            driver_string = "Driver={ODBC Driver 13 for SQL Server};"
            connection_string = driver_string + id_string
            cnxn = pyodbc.connect(connection_string)
        except Exception as e:
            raise TypeError(e)
    
    return cnxn


def buildCreationQuery (df, schema, table, customTypes=None,indexAsPrimaryKey=False):
    """Build query for creation an sql table based on the df structure
    df                  :(pd.DataFrame) input dataframe you want to base your table on
    schema              : (str) name of table schame
    table               : (str) name of table
    customTypes         : (dict) custom types for tqble colums. Key of the dict is name of column and value is MSSQL type
    indexAsPrimaryKey   : (bool) if True add "NOT NULL PRIMARY KEY" constraint on index column (default is False). The idx column will be called "id" in the db
    Return the query as a string
    """
    tableCreatationQuery = "CREATE TABLE " +  schema + "." + table + " ("
    if indexAsPrimaryKey is True: 
        tableCreatationQuery += "id INT NOT NULL PRIMARY KEY,"
    for column in df.columns:
        if (customTypes) and (column in customTypes):
            tableCreatationQuery += column + " " + customTypes[column] + ","
        elif str(df[column].dtype) in pdTypeToSQLTypeDict.keys():
            if df[column].dtype == 'int64':
                #limit between big int and small int
                if df[column].max() > 2147483647:
                    tableCreatationQuery += column + " BIGINT ,"
                else: 
                    tableCreatationQuery += column + " " + pdTypeToSQLTypeDict[str(df[column].dtype)] + ","

            else:    
                tableCreatationQuery += column + " " + pdTypeToSQLTypeDict[str(df[column].dtype)] + ","
        else:
            tableCreatationQuery += column + " VARCHAR(MAX),"
    tableCreatationQuery = tableCreatationQuery[:-1] + ");"
    
    return tableCreatationQuery


def buildTable (df, server, database, schema, table, customTypes=None, indexAsPrimaryKey=False, SQLLogin=False, username=None, password=None):
    """Build query for creation an sql table based on the df structure
    df                  :(pd.DataFrame) input dataframe you want to base your table on
    server              : (str) server adress
    database            : (str) database name
    schema              : (str) name of table schame
    table               : (str) name of table
    customTypes         : (dict) custom types for tqble colums. Key of the dict is name of column and value is MSSQL type
    indexAsPrimaryKey   : (bool) if True add "NOT NULL PRIMARY KEY" constraint on index column (default is False). The idx column will be called "id" in the db
    SQLLogin            : (bool) Set sql connection protocol. If True use user/pwd connection if False is AD for user info connection. Delfault is False
    username            : (str) username used for connection (Only needed if SQLLogin=True)
    password            : (str) passeword used for conneciton (Only needed if SQLLogin=True)
    Return void
    """
    cnxn = odbcConnect(server, database, SQLLogin=SQLLogin, username=username, password=password)
    cursor = cnxn.cursor()
    #first we drop table if already exist
    cursor.execute("DROP TABLE IF EXISTS " + schema + "." + table + ";")
    cnxn.commit()
    #then we create the new table
    tableCreatationQuery = buildCreationQuery(df, schema, table, customTypes=customTypes, indexAsPrimaryKey=indexAsPrimaryKey)
    cursor.execute(tableCreatationQuery)
    cnxn.commit()

    cnxn.close()
    return


def buildInsertQuery (df, schema, table, indexAsPrimaryKey=False):
    """ Build an insertion query to export df data in to sql.
    df                  :(pd.DataFrame) input dataframe you want to base your table on
    schema              : (str) name of table schame
    table               : (str) name of table
    indexAsPrimaryKey   : (bool) if True if index column is table primary key (default is False).
    Return the query as a string. This query can then be used with pyodbc importmany function.
    """

    tableInsertQuery = "INSERT INTO " + schema + "." + table +" ("
    if indexAsPrimaryKey is True:
        tableInsertQuery += "id,"
    for column in df.columns:
        tableInsertQuery += column + ", "
    nbColumns = len(df.columns) + 1 if indexAsPrimaryKey else len(df.columns)
    tableInsertQuery = tableInsertQuery[:-2] + ") values (" + "?,"*(nbColumns - 1) + "?);"
    return tableInsertQuery


def insertData (df, server, database, schema, table, customTypes=None, indexAsPrimaryKey=False, SQLLogin=False, username=None, password=None):
    """
    df                  :(pd.DataFrame) input dataframe you want to base your table on
    server              : (str) server adress
    database            : (str) database name
    schema              : (str) name of table schame
    table               : (str) name of table
    customTypes         : (dict) custom types for tqble colums. Key of the dict is name of column and value is MSSQL type
    indexAsPrimaryKey   : (bool) if True if index column is table primary key (default is False).
    SQLLogin            : (bool) Set sql connection protocol. If True use user/pwd connection if False is AD for user info connection. Delfault is False
    username            : (str) username used for connection (Only needed if SQLLogin=True)
    password            : (str) passeword used for conneciton (Only needed if SQLLogin=True)
    Return void

    """
    tableInsertQuery = buildInsertQuery (df, schema, table, indexAsPrimaryKey)
    if indexAsPrimaryKey is True:
        outputList = np.concatenate((np.array([df.index]).T, df.values), axis=1).tolist()
    else:
        outputList = df.values.tolist()

    #filtering nan vallues
    outputList = list(map(lambda row: [x if not pd.isna(x) else None for x in row], outputList))
    
    cnxn = odbcConnect(server, database, SQLLogin=SQLLogin, username=username, password=password)
    cursor = cnxn.cursor()

    if not (cursor.tables(table=table, schema=schema, tableType="TABLE").fetchone()):
        buildTable(df, server, database, schema, table, customTypes=customTypes, indexAsPrimaryKey=indexAsPrimaryKey,
            SQLLogin=SQLLogin, username=username, password=password
            )


    cursor.executemany(tableInsertQuery, outputList)
    cnxn.commit()

    cnxn.close()
    return

