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
df = pd.read_excel(url, sheet_name=sheet_name)

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Delete the first row and last
# df = df.drop(df.index[0])
df.drop(df.index[-5:], inplace=True)
# delete the first column
df = df.drop(df.columns[0], axis=1)

df.iloc[:, 0] = df.iloc[:, 0].ffill()
df.iloc[0, 0] = 'קבוצת אוכלוסייה'
df.iloc[:, 1] = df.iloc[:, 1].ffill()

# Get the range of columns from column 14 to the end
columns_to_process = df.columns[13:]
# Iterate over the DataFrame and update the values in the row below
for col in columns_to_process:
    upper_row_value = str(df.at[0, col])
    lower_row_value = str(df.at[1, col])
    new_value = lower_row_value + ' ' + upper_row_value
    df.at[0, col] = new_value

df.drop(df.index[1], inplace=True)
# Reset the index after deletion
df.reset_index(drop=True, inplace=True)

df.iloc[0, 1] = 'גיל'
df.iloc[0, 2] = 'מין'

# Set the first row as column names and remove that row
df.columns = df.iloc[0]
df = df[1:]

##################################################

id_vars = ['קבוצת אוכלוסייה', 'גיל', 'מין']
value_vars = df.columns[3:].tolist()
# Unpivot the DataFrame to long format
unpivoted_df = pd.melt(df, value_vars=value_vars, id_vars=id_vars,
                       var_name='שנה', value_name='מועסקים')
unpivoted_df['שנה'] = unpivoted_df['שנה'].astype(str)

unpivoted_df['רבעון'] = unpivoted_df['שנה'].apply(lambda x: '1' if (x.count('I') == 1 and 'V' not in x) else
                                                  ('2' if x.count('I') == 2 else
                                                   ('3' if x.count('I') == 3 else
                                                    ('4' if 'IV' in x else
                                                     ('0' if '.0' in x else
                                                      '')))))

print(unpivoted_df)

# # Define the Excel file name
# excel_file_name = 'output1.xlsx'
# # Write the DataFrame to an Excel file in the local directory
# new.to_excel(excel_file_name, index=False)