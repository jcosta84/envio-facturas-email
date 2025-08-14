from sqlalchemy import create_engine
import pandas as pd


# Configurações
usuario = "paulo"
senha = "loucoste9309323"
servidor = "192.168.38.254"  # ou nome DNS
porta = "1433"
banco = "factura_email"

# String de conexão
conn_str = f"mssql+pyodbc://{usuario}:{senha}@{servidor}:{porta}/{banco}?driver=ODBC+Driver+17+for+SQL+Server"

# Criar engine
engine = create_engine(conn_str)

query = "SELECT * FROM clientes"
cliente = pd.read_sql(query, engine)

print(cliente)