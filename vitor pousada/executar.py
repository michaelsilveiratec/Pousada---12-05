import pandas as pd
import pyodbc

# === CONEXÃO COM O SQL SERVER ===
try:
    conn = pyodbc.connect(
        r"Driver={SQL Server};"
        r"Server=SEU_SERVIDOR\SQLEXPRESS;"  # Substitua pelo nome correto do servidor
        r"Database=SEU_BANCO;"              # Substitua pelo nome correto do banco
        r"Trusted_Connection=yes;"
    )
    print("Conexão bem-sucedida!")
    conn.close()
except Exception as e:
    print(f"Erro ao conectar ao SQL Server: {e}")
cursor = conn.cursor()

# === LEITURA DAS PLANILHAS DO EXCEL ===
arquivo = "reservas_hotel.xlsx"

cliente = pd.read_excel(arquivo, sheet_name="cliente")
quarto = pd.read_excel(arquivo, sheet_name="quarto")
status = pd.read_excel(arquivo, sheet_name="status")
data = pd.read_excel(arquivo, sheet_name="data")
reserva = pd.read_excel(arquivo, sheet_name="reserva")

# === INSERÇÃO EM CADA TABELA ===
try:
    # Cliente
    for _, row in cliente.iterrows():
        if pd.isnull(row['Id']) or pd.isnull(row['Nome']):
            continue  # Ignorar linhas inválidas
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
        """, pd.to_datetime(row['Data']).strftime('%Y-%m-%d'),
             int(row['Nro_quarto']), row['Status'])

    # Reserva
    for _, row in reserva.iterrows():
        cursor.execute("""
            INSERT INTO Reserva (
                Nro, Nome_status, Data_inicio_data,
                Numero_quarto_data_inicio, Data_fim_data,
                Nro_quarto_data_fim, Data, Id_cliente
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, int(row['Nro']), row['Nome_status'],
             pd.to_datetime(row['Data_inicio_data']).strftime('%Y-%m-%d'),
             int(row['Numero_quarto_data_inicio']),
             pd.to_datetime(row['Data_fim_data']).strftime('%Y-%m-%d'),
             int(row['Nro_quarto_data_fim']), row['Data'], int(row['Id_cliente']))

    conn.commit()
    print("✅ Dados do Excel carregados com sucesso no SQL Server!")

except Exception as e:
    print(f"Erro durante o ETL: {e}")

finally:
    conn.close()