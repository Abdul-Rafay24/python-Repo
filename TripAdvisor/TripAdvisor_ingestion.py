import pandas as pd
import json
from mssql_connector import MSSQLConnector as CON

# Path to your Excel file
excel_file_path = r'D:\TRIPADVISOR_KSA\DETAILS\attraction_details.xlsx'

# Read specific columns from Excel into a DataFrame
columns_to_ingest = ['ATTRACTION_NAME', 'CONTINENT', 'COUNTRY', 'PROVINCE', 'RATING', 'TIMING']
df = pd.read_excel(excel_file_path, usecols=columns_to_ingest)

# Load database configuration from JSON file
with open("C:\\Users\\HP\\Desktop\\d_config.json", "r") as f:
    d_config = json.load(f)

# Create a database connection
conn = CON(d_config=d_config)

# Specify the target table name in your database
target_table_name = '[SIDR].[TRIP_ADVISOR].[ATTRACTION]'

# Ingest data into the SQL Server table
conn.push_data(df, target_table_name)
