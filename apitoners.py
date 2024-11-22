import sqlite3
from flask import Flask, jsonify
import requests
from bs4 import BeautifulSoup

def initialize_db():
    conn = sqlite3.connect('apitoners.db')
    cursor = conn.cursor()
    
    # Criar tabela
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS impressoras (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unidade TEXT NOT NULL,
            marca TEXT NOT NULL,
            modelo TEXT NOT NULL,
            toner TEXT NOT NULL,
            setor TEXT NOT NULL,
            ip TEXT NOT NULL
        )
    ''')
    
    conn.commit()
    conn.close()

app = Flask(__name__)

# Função para acessar o banco
def get_impressoras():
    conn = sqlite3.connect('apitoners.db')
    cursor = conn.cursor()
    cursor.execute('SELECT unidade, setor, marca, ip FROM impressoras')
    impressoras = cursor.fetchall()
    conn.close()
    return impressoras

# Função para extrair nível de toner (Exemplo Simples)
def get_toner_level(ip):
    response = requests.get(f'http://{ip}', verify=False)
    response.raise_for_status()

    # Parse do HTML para extrair informações
    soup = BeautifulSoup(response.text, 'html.parser')
    image = soup.find('img', {'class': 'tonerremain'})

    if image:
        # Tenta obter a altura do atributo "height"
        height = image.get('height')

        # Se a altura não estiver presente como atributo, retorna mensagem
        if height:
            return f"{height}"
        else:
            return "A altura da imagem não está especificada no atributo HTML."
    else:
        soup = BeautifulSoup(response.text, 'html.parser')
        image = soup.find('img', {'class': 'tonerremain'})
    
        return "Imagem com a classe 'tonerremain' não encontrada."

# Endpoint da API
@app.route('/toner-levels', methods=['GET'])
def toner_levels():
    impressoras = get_impressoras()
    results = []

    for unidade, setor, marca, ip in impressoras:
        toner_level = get_toner_level(ip)
        results.append({'unidade': unidade, 'setor': setor, 'marca': marca, 'ip': ip, 'toner_level': toner_level})

    return jsonify(results)

if __name__ == '__main__':
    initialize_db()
    app.run(debug=True)
