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

table_name = 'Taasuka_Lamas_12102023'
schema_name = 'stage'

#############################################
#
# # Write and execute the SQL CREATE TABLE query
# create_table_query = f'''
# CREATE TABLE [stage].[Taasuka_Lamas_12102023] (
#     קבוצת_אוכלוסייה NVARCHAR(100),
#     גיל NVARCHAR(50),
#     מין NVARCHAR(50),
#     שנה NVARCHAR(50),
#     סהכ_אוכלוסייה FLOAT,
#     רבעון INT,
#     מועסקים FLOAT,
#     אחוז_השתתפות_כוח_עבודה FLOAT
# )
# '''
#
# cursor.execute(create_table_query)
# connection.commit()

cursor.execute(f"TRUNCATE TABLE {schema_name}.{table_name}")
cursor.fast_executemany = True

# Sheet1
######################################

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'taasuka'
blob_name = 'Taasuka_Lamas_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
print(url)
sheet_name = 'סה"כ אוכלוסייה - אלפים'
df1 = pd.read_excel(url, sheet_name=sheet_name)

# Delete the first row and last
df1.drop(df1.index[-5:], inplace=True)
# delete the first column
df1 = df1.drop(df1.columns[0], axis=1)

df1.iloc[:, 0] = df1.iloc[:, 0].ffill()
df1.iloc[0, 0] = 'קבוצת אוכלוסייה'
df1.iloc[:, 1] = df1.iloc[:, 1].ffill()


# Get the range of columns from column 14 to the end
columns_to_process = df1.columns[13:]
# Iterate over the DataFrame and update the values in the row below
for col in columns_to_process:
    upper_row_value = str(df1.at[0, col])
    lower_row_value = str(df1.at[1, col])
    new_value = lower_row_value + ' ' + upper_row_value
    df1.at[0, col] = new_value

df1.drop(df1.index[1], inplace=True)
# Reset the index after deletion
df1.reset_index(drop=True, inplace=True)

df1.iloc[0, 1] = 'גיל'
df1.iloc[0, 2] = 'מין'

# Set the first row as column names and remove that row
df1.columns = df1.iloc[0]
df1 = df1[1:]

id_vars = ['קבוצת אוכלוסייה', 'גיל', 'מין']
value_vars = df1.columns[3:].tolist()
# Unpivot the DataFrame to long format
df_sheet1 = pd.melt(df1, value_vars=value_vars, id_vars=id_vars,
                    var_name='שנה', value_name='סה"כ אוכלוסייה')
df_sheet1['שנה'] = df_sheet1['שנה'].astype(str)

df_sheet1['רבעון'] = df_sheet1['שנה'].apply(lambda x: '1' if (x.count('I') == 1 and 'V' not in x) else
                                                  ('2' if x.count('I') == 2 else
                                                   ('3' if x.count('I') == 3 else
                                                    ('4' if 'IV' in x else
                                                     ('0' if '.0' in x else
                                                      '')))))

# Sheet 2
##################################################

import pandas as pd

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'taasuka'
blob_name = 'Taasuka_Lamas_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
print(url)
sheet_name = 'מועסקים - אלפים'
df2 = pd.read_excel(url, sheet_name=sheet_name)

# Delete the first row and last
# df = df.drop(df.index[0])
df2.drop(df2.index[-5:], inplace=True)
# delete the first column
df2 = df2.drop(df2.columns[0], axis=1)

df2.iloc[:, 0] = df2.iloc[:, 0].ffill()
df2.iloc[0, 0] = 'קבוצת אוכלוסייה'
df2.iloc[:, 1] = df2.iloc[:, 1].ffill()

# Get the range of columns from column 14 to the end
columns_to_process = df2.columns[13:]
# Iterate over the DataFrame and update the values in the row below
for col in columns_to_process:
    upper_row_value = str(df2.at[0, col])
    lower_row_value = str(df2.at[1, col])
    new_value = lower_row_value + ' ' + upper_row_value
    df2.at[0, col] = new_value

df2.drop(df2.index[1], inplace=True)
# Reset the index after deletion
df2.reset_index(drop=True, inplace=True)

df2.iloc[0, 1] = 'גיל'
df2.iloc[0, 2] = 'מין'

# Set the first row as column names and remove that row
df2.columns = df2.iloc[0]
df2 = df2[1:]

id_vars = ['קבוצת אוכלוסייה', 'גיל', 'מין']
value_vars = df2.columns[3:].tolist()
# Unpivot the DataFrame to long format
df_Sheet2 = pd.melt(df2, value_vars=value_vars, id_vars=id_vars,
                    var_name='שנה', value_name='מועסקים')
df_Sheet2['שנה'] = df_Sheet2['שנה'].astype(str)

df_Sheet2['רבעון'] = df_Sheet2['שנה'].apply(lambda x: '1' if (x.count('I') == 1 and 'V' not in x) else
                                                  ('2' if x.count('I') == 2 else
                                                   ('3' if x.count('I') == 3 else
                                                    ('4' if 'IV' in x else
                                                     ('0' if '.0' in x else
                                                      '')))))

# Sheet 3
################################################

username = 'Data_Dev@machon.org.il'
account_name = 'filedatacentermachon2'
account_key = 'YallaHapoel4Ever1923!'
container_name = 'taasuka'
blob_name = 'Taasuka_Lamas_02102023.xlsx'
# Read the Excel File
url = f"https://{account_name}.blob.core.windows.net/{container_name}/{blob_name}"
print(url)
sheet_name = 'אחוז השתתפות  בכוח העבודה'
df3 = pd.read_excel(url, sheet_name=sheet_name)

# Delete the first row and last
# df = df.drop(df.index[0])
df3.drop(df3.index[-5:], inplace=True)
# delete the first column
df3 = df3.drop(df3.columns[0], axis=1)

df3.iloc[:, 0] = df3.iloc[:, 0].ffill()
df3.iloc[:, 1] = df3.iloc[:, 1].ffill()

columns_to_add_0 = df3.columns[:13]
for col in columns_to_add_0:
    new_col_name = str(col) + '.0'
    df3.rename(columns={col: new_col_name}, inplace=True)

# Get the range of columns from column 14 to the end
columns_to_process = df3.columns[13:]
# Iterate over the DataFrame and update the values in the row below
for col in columns_to_process:
    upper_row_value = col[:-2]
    lower_row_value = str(df3.at[0, col])
    new_value = lower_row_value + ' ' + upper_row_value
    df3.rename(columns={col: new_value}, inplace=True)

# Set the first row as column names and remove that row
df3.drop(df3.index[0], inplace=True)
df3.rename(columns={df3.columns[0]: 'קבוצת אוכלוסייה'}, inplace=True)
df3.rename(columns={df3.columns[1]: 'גיל'}, inplace=True)
df3.rename(columns={df3.columns[2]: 'מין'}, inplace=True)

id_vars = ['קבוצת אוכלוסייה', 'גיל', 'מין']
value_vars = df3.columns[3:].tolist()
# Unpivot the DataFrame to long format
df_Sheet3 = pd.melt(df3, value_vars=value_vars, id_vars=id_vars,
                    var_name='שנה', value_name='אחוז השתתפות בכוח העבודה')
df_Sheet3['שנה'] = df_Sheet3['שנה'].astype(str)

df_Sheet3['רבעון'] = df_Sheet3['שנה'].apply(lambda x: '1' if (x.count('I') == 1 and 'V' not in x) else
                                                  ('2' if x.count('I') == 2 else
                                                   ('3' if x.count('I') == 3 else
                                                    ('4' if 'IV' in x else
                                                     ('0' if '.0' in x else
                                                      '')))))

# merge 3 df
################################################

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
merged = df_sheet1.merge(df_Sheet2, on=['קבוצת אוכלוסייה', 'גיל', 'מין', 'שנה', 'רבעון'], how='inner')
merged = merged.merge(df_Sheet3, on=['קבוצת אוכלוסייה', 'גיל', 'מין', 'שנה', 'רבעון'], how='inner')
print(merged.head())

chosen_columns = ['סה"כ אוכלוסייה', 'רבעון', 'מועסקים', 'אחוז השתתפות בכוח העבודה']
desired_types = [float, int, float, float]
# Create a dictionary by pairing columns with their desired types
dtype_dict = dict(zip(chosen_columns, desired_types))
merged = merged.astype(dtype_dict)

fillna_dict = {'סה"כ אוכלוסייה': -1, 'מועסקים': -1, 'אחוז השתתפות בכוח העבודה': -1}
merged.fillna(value=fillna_dict, inplace=True)
##############################################

query = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table_name}'"
cursor.execute(query)
# Fetch the column names
column_names_sql = [row.COLUMN_NAME for row in cursor.fetchall()]
column_names_sql2 = ', '.join(column_names_sql)

column_names_excel = ', '.join(merged.columns)
placeholders = ', '.join(['?'] * len(merged.columns))

# Prepare the data for bulk insert
values = [tuple(row) for row in merged.values]

# Insert Data into SQL Server in Bulk
query = f"INSERT INTO {schema_name}.{table_name} ({column_names_sql2}) VALUES ({placeholders})"

cursor.executemany(query, values)
connection.commit()
cursor.close()
connection.close()
print('doneso')
