import pyodbc

try:
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=192.168.52.180,1433;"
        "DATABASE=factura_email;"
        "UID=sa;"
        "PWD=loucoste9850053;"
    )
    print("Conectado com sucesso!")
except Exception as e:
    print("Erro de conexão:", e)