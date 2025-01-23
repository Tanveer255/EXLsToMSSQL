import pandas as pd
from sqlalchemy import create_engine

# MSSQL database credentials
hostname = 'your-servername'  # SQL Server hostname or IP
username = 'your-username'  # SQL Server username
password = 'your-password'  # SQL Server password
database = 'your-database-name'  # SQL Server database name

# Create a connection to the SQL Server database using SQLAlchemy and pyodbc
engine = create_engine(f'mssql+pyodbc://{username}:{password}@{hostname}/{database}?driver=ODBC+Driver+17+for+SQL+Server')

# Path to the Excel file (D:\\EdgeCTP.Net8.0\\edge_authServer_db.xlsx)
excel_file_path = "your-excel-file-path"  # Update this path to your Excel file location

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_file_path)

# Upload the DataFrame to SQL Server table
table_name = 'User'  # Add your SQL Server table name here
df.to_sql(table_name, con=engine, if_exists='replace', index=False)

print(f'Excel file has been uploaded to the SQL Server table "{table_name}".')
