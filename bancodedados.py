import sqlite3
from tabulate import tabulate
import pandas as pd

# Nome do banco de dados
DB_NAME = "apitoners.db"

# Função para cadastrar uma impressora
def add_printer(unidade, marca, modelo, toner, setor, ip):
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO impressoras (unidade, marca, modelo, toner, setor, ip) VALUES (?, ?, ?, ?, ?, ?)
        ''', (unidade, marca, modelo, toner, setor, ip))
        conn.commit()
        print(f"Impressora adicionada com sucesso.")
    except sqlite3.IntegrityError:
        print(f"Erro: Já existe uma impressora com o IP '{ip}'.")
    except Exception as e:
        print(f"Erro ao adicionar impressora: {e}")
    finally:
        conn.close()

# Função para buscar todas as impressoras
def get_all_printers():
    """
    Retorna uma lista de todas as impressoras no banco de dados.
    :return: Lista de dicionários com as informações das impressoras
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('SELECT id, unidade, marca, modelo, toner, setor, ip FROM impressoras')
    printers = cursor.fetchall()
    conn.close()
    
    headers = ["ID", "Unidade", "Marca", "Modelo", "Toner", "Setor", "IP"]
    return tabulate(printers, headers=headers, tablefmt="fancy_grid")

def cadastrar_excel():
    try:
        df = pd.read_excel('impressoras.xlsx')
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return
    
    # Verifica se as colunas necessárias estão presentes
    colunas_necessarias = {'unidade', 'marca', 'modelo', 'toner', 'setor', 'ip'}
    if not colunas_necessarias.issubset(df.columns):
        print(f"A planilha deve conter as seguintes colunas: {colunas_necessarias}")
        return
    
    # Conecta ao banco de dados
    conn = sqlite3.connect('apitoners.db')
    cursor = conn.cursor()

    # Insere as impressoras na tabela
    for _, row in df.iterrows():
        try:
            cursor.execute('''
                INSERT INTO impressoras (unidade, marca, modelo, toner, setor, ip)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (row['unidade'], row['marca'], row['modelo'], row['toner'], row['setor'], row['ip']))
        except Exception as e:
            print(f"Erro ao inserir a impressora {row['modelo']} ({row['ip']}): {e}")
    
    conn.commit()
    conn.close()
    print("Impressoras cadastradas com sucesso!")

sair = False

while sair == False:
    resposta = int(input("\nO que deseja fazer?\n1 - Buscar todas as impressoras cadastradas\n2 - Cadastrar uma impressora\n3 - Cadastrar impressoras através de planilha Excel\n4 - Sair\n"))
    
    if resposta == 1:
        print(get_all_printers())
    elif resposta == 2:
        unidade = input("\nUnidade em que a impressora está: ")
        marca = input("Marca da impressora: ")
        modelo = input("Modelo da impressora: ")
        toner = input("Toner que a impressora utiliza: ")
        setor = input("Setor em que a impressora está: ")
        ip = input("IP da impressora: ")
        
        add_printer(unidade, marca, modelo, toner, setor, ip)
    elif resposta == 3:
        cadastrar_excel()
    elif resposta == 4:
        sair = True
    else:
        print("Resposta inválida")

