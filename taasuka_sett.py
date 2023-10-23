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

table_name = 'Taasuka_Sett_12102023'
schema_name = 'stage'

cursor.execute(f"TRUNCATE TABLE {schema_name}.{table_name}")
cursor.fast_executemany = True

#############################################

# # Write and execute the SQL CREATE TABLE query
# create_table_query = f'''
# CREATE TABLE [stage].[Taasuka_Sett_12102023] (
#     שנה INT,
#     מין NVARCHAR(50),
#     קבוצת_אוכלוסייה NVARCHAR(100),
#     זרם_חרדי NVARCHAR(100),
#     סמל_יישוב INT,
#     תושבים_25_64 INT,
#     מועסקים_25_64 INT,
#     הכנסה_שנתית_שכירים INT,
#     הכנסה_שנתית_עצמאים INT,
#     הכנסה_שנתית_עש INT,
#     הכנסה_חודשית_שכירים INT,
#     ממוצע_חודשי_עבודה FLOAT
# )
# '''
#
# cursor.execute(create_table_query)
# connection.commit()

######################################

# dictionaries for editing the dataframe
chosen_columns = ['שנה', 'מין', 'קבוצת אוכלוסייה', 'זרם חרדי', 'סמל יישוב', 'תושבים בגילאי 25-64', 'מועסקים בגילי 25-64',
                  'הכנסה שנתית ממוצעת (עצמאים)',
                  'הכנסה שנתית ממוצעת (שכירים)', 'הכנסה חודשית ממוצעת (שכירים)', 'הכנסה שנתית ממוצעת (עצמאים ושכירים)',
                  'ממוצע חודשי עבודה לעובד']
desired_types = [int, str, str, str, int, int, int, int, int, int, int, float]

# Create a dictionary by pairing columns with their desired types
dtype_dict = dict(zip(chosen_columns, desired_types))
fillna_dict = {'זרם חרדי': 0, 'הכנסה שנתית ממוצעת (עצמאים)': 0}

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'taasuka'
blob_name = 'Taasuka_Sett_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
print(url)
sheet_name = 'נתונים לדשבורד תעסוקה מינהלי'
df = pd.read_excel(url, sheet_name=sheet_name)
print('done read')

df.fillna(value=fillna_dict, inplace=True)
df = df.astype(dtype_dict)
df['ממוצע חודשי עבודה לעובד'] = df['ממוצע חודשי עבודה לעובד'].round(1)

##########################################

extracted_df = df.loc[:, chosen_columns]

###########################################

query = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table_name}'"
cursor.execute(query)

# Fetch the column names
column_names_sql = [row.COLUMN_NAME for row in cursor.fetchall()]
column_names_excel = ', '.join(extracted_df.columns)
column_names_sql2 = ', '.join(column_names_sql)
placeholders = ', '.join(['?'] * len(extracted_df.columns))

# Prepare the data for bulk insert
values = [tuple(row) for row in extracted_df.values]

# Insert Data into SQL Server in Bulk
query = f"INSERT INTO {schema_name}.{table_name} ({column_names_sql2}) VALUES ({placeholders})"

cursor.executemany(query, values)
connection.commit()
cursor.close()
connection.close()
