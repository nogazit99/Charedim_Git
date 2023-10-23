import pyodbc
import pandas as pd

# Connection details
server = 'machon.database.windows.net'
database = 'Machon_DB'
# username = 'Data_Dev'
username = 'Noga_Gazit'
password = '1231!#ASDF!a'
# password = 'Hapoel2010987654%^'
driver = '{ODBC Driver 17 for SQL Server}'
# Construct the connection string
connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
# Establish a connection
connection = pyodbc.connect(connection_string)

#############################################

# Create a cursor
cursor = connection.cursor()

table_name = 'Demographic_Sett_12102023'
schema_name = 'stage'

#############################################

# # Write and execute the SQL CREATE TABLE query
# create_table_query = '''
# CREATE TABLE [stage].[Demographic_Sett_12102023] (
#     סמל_יישוב INT,
#     סהכ_תושבים INT,
#     חרדים INT,
#     יהודים_שאינם_חרדים INT,
#     ערבים INT,
#     חרדים_ללא_מספר NVARCHAR(10),
#     ספרדים INT,
#     ליטאים INT,
#     חסידים INT,
#     חבד INT
# )
# '''
#
# cursor.execute(create_table_query)
# connection.commit()

######################################

# dictionaries for editing the dataframe
chosen_columns = ['סמל יישוב', 'סה"כ אוכלוסייה ביישוב', 'חרדים', 'יהודים ואחרים שאינם חרדים', 'ערבים',
                  'יש חרדים ביישוב אך לא ניתן להציג את מספרם', 'ספרדים', 'ליטאים', 'חסידים', 'חב"ד']
desired_types = [int, int, int, int, int, str, int, int, int, int]
# dictionary
dtype_dict = dict(zip(chosen_columns, desired_types))
# dtype_dict = {'סמל יישוב': int, 'סה"כ אוכלוסייה ביישוב': int, 'חרדים': int,
#               'יהודים ואחרים שאינם חרדים': int, 'ערבים': int, 'ספרדים': int, 'ליטאים': int,
#               'חסידים': int, 'חב"ד': int}
fillna_dict = {'חרדים': 0, 'ערבים': 0, 'יהודים ואחרים שאינם חרדים': 0, 'יש חרדים ביישוב אך לא ניתן להציג את מספרם': 0}

#######################################

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'demographic'
blob_name = 'Demographic_Sett_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
sheet_name = 'יישובים'  # Replace with the actual sheet name or index
df = pd.read_excel(url, sheet_name=sheet_name)

df.fillna(value=fillna_dict, inplace=True)
df = df.astype(dtype_dict)

##########################################

extracted_df = df.loc[:, chosen_columns]

###########################################

# query = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{schema_name}.{table_name}'"
query = f'''
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = '{schema_name}' AND TABLE_NAME = '{table_name}';
'''
cursor.execute(query)

# Fetch the column names
column_names_sql = [row.COLUMN_NAME for row in cursor.fetchall()]
# Join the strings into one big string with a space separator
column_names_sql2 = ', '.join(column_names_sql)
# column_names_sql = 'סמל_יישוב, סהכ_תושבים, חרדים, יהודים_שאינם_חרדים, ערבים, חרדים_ללא_מספר, ספרדים, ליטאים, חסידים, חבד'

placeholders = ', '.join(['?'] * len(extracted_df.columns))

# Prepare the data for bulk insert
values = [tuple(row) for row in extracted_df.values]

# Insert Data into SQL Server in Bulk
query = f"INSERT INTO [{schema_name}].[{table_name}] ({column_names_sql2}) VALUES ({placeholders})"

cursor.executemany(query, values)
connection.commit()
cursor.close()
connection.close()
print('done')
