import os
import pyodbc
import pandas as pd

# Veritabanı bağlantı detaylarını al
MSSQL_SERVER = os.getenv('MSSQL_SERVER')
MSSQL_DATABASE = os.getenv('MSSQL_DATABASE')
MSSQL_USER = os.getenv('MSSQL_USER')
MSSQL_PASSWORD = os.getenv('MSSQL_PASSWORD')
schema_name = ''  # İlgili şema adı
directory_to_save_excel = r'C:\Users\dogan\OneDrive\Masaüstü\Çalışmalar\Digirent\Faz-2\Faz-2 Çalışmalar\Faz-2 Konular\Veri kalite kontrol çalışmaları\Veriler\\'  # Directory to save Excel files

print("Veritabanına bağlanıyor...")
# Bağlantı dizesi
conn_str = f'DRIVER={{SQL Server}};SERVER={MSSQL_SERVER};DATABASE={MSSQL_DATABASE};UID={MSSQL_USER};PWD={MSSQL_PASSWORD}'

# Veritabanına bağlan
conn = pyodbc.connect(conn_str)
print("Veritabanına başarıyla bağlanıldı.")

# Tüm tablo isimlerini ve boyutlarını çek
cursor = conn.cursor()
table_size_query = f"""
SELECT 
    t.NAME AS TableName,
    s.Name AS SchemaName,
    p.rows AS RowCounts,
    SUM(a.total_pages) * 8 / 1024.0 AS TotalSpaceMB
FROM 
    sys.tables t
INNER JOIN      
    sys.indexes i ON t.OBJECT_ID = i.object_id
INNER JOIN 
    sys.partitions p ON i.object_id = p.OBJECT_ID AND i.index_id = p.index_id
INNER JOIN 
    sys.allocation_units a ON p.partition_id = a.container_id
INNER JOIN 
    sys.schemas s ON t.schema_id = s.schema_id
WHERE 
    s.Name = '' 
GROUP BY 
    t.Name, s.Name, p.Rows
HAVING 
    SUM(a.total_pages) * 8 / 1024.0 <= 11
ORDER BY 
    TotalSpaceMB DESC;
"""

# Sorguyu çalıştır ve sonuçları DataFrame'e kaydet
table_sizes_df = pd.read_sql(table_size_query, conn)
print(table_sizes_df)

# DataFrame üzerinde döngü yap ve tabloları indir
for index, row in table_sizes_df.iterrows():
    table_name = row['TableName']
    print(f"Processing table: {table_name}")
    user_input = input(f"Do you want to download {table_name}? (yes/no): ")
    if user_input.lower() == "y":
        sql_query = f'SELECT * FROM {schema_name}.{table_name}'
        df = pd.read_sql(sql_query, conn)
        excel_file_path = f'{directory_to_save_excel}{table_name}.xlsx'
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print(f'Successfully saved {table_name} records to {excel_file_path}')
    else:
        print(f"Skipping table: {table_name}")

# Veritabanı bağlantısını kapat
conn.close()
print("İşlem tamamlandı.")
