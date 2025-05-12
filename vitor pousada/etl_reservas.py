import pandas as pd
import pyodbc

# === CONEXÃO COM O SQL SERVER ===
conn = pyodbc.connect(
    r"Driver={SQL Server};"
    r"Server=DESKTOP-ABC123\SQLEXPRESS;"
    r"Database=HotelDB;"                # Substitua pelo nome do seu banco
    r"Trusted_Connection=yes;"
)
cursor = conn.cursor()

# === LEITURA DO EXCEL ===
arquivo = "reservas_hotel.xlsx"

cliente = pd.read_excel(arquivo, sheet_name="cliente")
quarto = pd.read_excel(arquivo, sheet_name="quarto")
status = pd.read_excel(arquivo, sheet_name="status")
data = pd.read_excel(arquivo, sheet_name="data")
reserva = pd.read_excel(arquivo, sheet_name="reserva")

# === INSERÇÃO EM CADA TABELA ===

# Cliente
for _, row in cliente.iterrows():
    cursor.execute("""
        INSERT INTO Cliente (Id, Nome, Telefone, Cpf)
        VALUES (?, ?, ?, ?)
    """, int(row['Id']), row['Nome'], row['Telefone'], row['Cpf'])

# Quarto
for _, row in quarto.iterrows():
    cursor.execute("""
        INSERT INTO Quarto (Nro, Preco, Descricao)
        VALUES (?, ?, ?)
    """, int(row['Nro']), float(row['Preco']), row['Descricao'])

# Status
for _, row in status.iterrows():
    cursor.execute("""
        INSERT INTO Status (Nome)
        VALUES (?)
    """, row['Nome'])

# Data
for _, row in data.iterrows():
    cursor.execute("""
        INSERT INTO Data (Data, Nro_quarto, Status)
        VALUES (?, ?, ?)
    """, row['Data'], int(row['Nro_quarto']), row['Status'])

# Reserva
for _, row in reserva.iterrows():
    cursor.execute("""
        INSERT INTO Reserva (
            Nro, Nome_status, Data_inicio_data,
            Numero_quarto_data_inicio, Data_fim_data,
            Nro_quarto_data_fim, Data, Id_cliente
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, int(row['Nro']), row['Nome_status'], row['Data_inicio_data'],
         int(row['Numero_quarto_data_inicio']), row['Data_fim_data'],
         int(row['Nro_quarto_data_fim']), row['Data'], int(row['Id_cliente']))

conn.commit()
conn.close()

print("✅ Dados carregados com sucesso no SQL Server!")
