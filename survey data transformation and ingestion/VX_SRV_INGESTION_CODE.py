import pandas as pd
import json
from mssql_connector import MSSQLConnector as CON
import glob, os

# Define the Excel file path
excel_file = "E:\\Data_Files\\VX_New\\VX-DataMart\\"
path = "E:\\Data_Files\\VX_New\\VX-DataMart\\"
files = glob.glob(os.path.join(path , "*.xlsx"))
df_new = pd.DataFrame()

for f in files:
    print(files)
    frame = pd.read_excel(f, header=1)
    if ('D2' not in frame.columns) and ('D2_International' in frame.columns):
        frame.rename(columns={'D2_International':'D2'}, inplace=True)
    if ('TO3.2C8' not in frame.columns) and ('TO3.2C16' in frame.columns):
        frame.rename(columns={'TO3.2C16':'TO3.2C8'}, inplace=True)
    df_new = pd.concat([df_new,frame], ignore_index=True)

# Load database configuration from JSON file
with open("E:\\ETL\\VX_V2\\VX_SRV_INGESTION\\d_config.json", "r") as f:
    d_config = json.load(f)

# Create a database connection
conn = CON(d_config=d_config)

# Define the SQL statement to retrieve table names and columns from SRV.EXCEL_MAPPING
statement = "SELECT TBL_NAME, TBL_COLUMN FROM SRV.VX_EXCEL_MAPPING"
df_mapping = conn.query(statement)
#df_new = pd.read_excel(excel_file, header=1)
df_new['KEY'] = df_new['INTNR'].apply(str) + '_' + df_new['Date'].apply(str)
# Initialize an empty list to store DataFrames
dataframes_to_push = []

# Re-create the connection for pushing data
conn = CON(d_config=d_config)

# Iterate through the rows of df_mapping
for _, row in df_mapping.iterrows():
    table_name = row['TBL_NAME']
    columns_to_keep = row['TBL_COLUMN'].split(",")

    # Load the data from Excel
    df = pd.DataFrame()

    for col in columns_to_keep:
        if col in df_new.columns:
            df[col] = df_new[col]
        else:
            df[col] = None

    # Filter the DataFrame to keep only the desired columns
    #df = df[columns_to_keep]
    print(table_name)
    
    #print(df.columns)
    # Append the filtered DataFrame to the list
    #dataframes_to_push.append((table_name, df))

    conn.push_data(df, table_name)

# Now, concatenate and push all filtered DataFrames to the database in one batch
# for table_name, df in dataframes_to_push:
#     conn.push_data(df, table_name)