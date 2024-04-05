import pandas as pd
import pypyodbc as odbc
import sqlalchemy
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

server_name = '{SELECT @@servername}'
database = 'testdb'
driver = '{SQL Server}'

connection_url = URL.create('mssql+pyodbc',query={'odbc_connect': f'DRIVER={driver};SERVER={server_name};DATABASE={database}'})
engine =create_engine(connection_url,module=odbc)

excel_file = 'app_usage_final_v2.xlsx'
df_excel = pd.read_excel(excel_file,sheet_name=None)
#print(df_excel.keys())
#df_excel.values()

for sheet_name, df in df_excel.items():
    if df.empty:
        continue
    df.to_sql(sheet_name,con=engine,if_exists='replace',index=False, dtype={'date':sqlalchemy.types.VARCHAR(length=10)})