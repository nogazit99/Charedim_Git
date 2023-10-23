import pyodbc
import pandas as pd

# Connection details
server = 'machon.database.windows.net'
database = 'Machon_DB'
username = 'Noga_Gazit'
password = '1231!#ASDF!a'
driver = '{ODBC Driver 17 for SQL Server}'
# Construct the connection string
connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
# Establish a connection
connection = pyodbc.connect(connection_string)

#############################################

# Create a cursor
cursor = connection.cursor()

table_name = 'Inst_12102023'
schema_name = 'stage'

cursor.execute(f"TRUNCATE TABLE {schema_name}.{table_name}")
cursor.fast_executemany = True

#############################################

# # Write and execute the SQL CREATE TABLE query
# create_table_query = '''
# CREATE TABLE [stage].[Inst_12102023] (
#     סמל_בית_ספר INT,
#     שם_מוסד NVARCHAR(100),
#     מגדר_תלמידים_במוסד NVARCHAR(100),
#     סיווג_מוסד NVARCHAR(100),
#     סיווג_מוסד_מורחב NVARCHAR(100),
#     מבנה_מוסד NVARCHAR(100),
#     זרם NVARCHAR(100),
#     זרם_מורחב NVARCHAR(100),
#     קבוצת_השתייכות_דתית NVARCHAR(10)
# )
# '''
#
# cursor.execute(create_table_query)
# connection.commit()

######################################

# dictionaries for editing the dataframe
chosen_columns = ['school_symbol', 'שם מוסד', 'מגדר התלמידים במוסד', 'סיווג מוסד', 'סיווג מוסד מורחב', 'מבנה מוסד',
                  'זרם', 'זרם מורחב', 'קבוצת השתייכות דתית']
desired_types = [int, str, str, str, str, str, str, str, str]
# dictionary
dtype_dict = dict(zip(chosen_columns, desired_types))
fillna_dict = {'school_symbol': 0, 'זרם מורחב': 0}

#############################################

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'educatoion'
blob_name = 'Inst_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
print(url)
sheet_name = 'Sheet1'
df = pd.read_excel(url, sheet_name=sheet_name)
print('done read')

df.fillna(value=fillna_dict, inplace=True)
df = df.astype(dtype_dict)

##########################################

extracted_df = df.loc[:, chosen_columns]

###########################################

# query = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table_name}'"
query = f'''
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = '{schema_name}' AND TABLE_NAME = '{table_name}';
'''
cursor.execute(query)

# Fetch the column names
column_names_sql = [row.COLUMN_NAME for row in cursor.fetchall()]
column_names_sql2 = ', '.join(column_names_sql)
placeholders = ', '.join(['?'] * len(extracted_df.columns))

# Prepare the data for bulk insert
values = [tuple(row) for row in extracted_df.values]

# # Insert Data into SQL Server in Bulk
query = f"INSERT INTO {schema_name}.{table_name} ({column_names_sql2}) VALUES ({placeholders})"
cursor.executemany(query, values)
connection.commit()

cursor.close()
connection.close()
