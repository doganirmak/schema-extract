import os
import pyodbc
import pandas as pd

# Retrieve database connection details from environment variables
MSSQL_SERVER = os.getenv('MSSQL_SERVER')
MSSQL_DATABASE = os.getenv('MSSQL_DATABASE')
MSSQL_USER = os.getenv('MSSQL_USER')
MSSQL_PASSWORD = os.getenv('MSSQL_PASSWORD')
schema_name = 'dbo'  # Schema you want to export tables from
directory_to_save_excel = 'path/to/save/excel/files/'  # Directory to save Excel files

print("Connecting to the database...")
# Connection string
conn_str = f'DRIVER={{SQL Server}};SERVER={MSSQL_SERVER};DATABASE={MSSQL_DATABASE};UID={MSSQL_USER};PWD={MSSQL_PASSWORD}'

# Connect to the database
conn = pyodbc.connect(conn_str)
print("Successfully connected to the database.")

# Fetching all table names in the specified schema
cursor = conn.cursor()
print(f"Fetching table names from schema: {schema_name}...")
cursor.execute(f"SELECT table_name FROM information_schema.tables WHERE table_schema = '{schema_name}'")
tables = cursor.fetchall()
print(f"Found {len(tables)} tables. Starting to export to Excel files...")
