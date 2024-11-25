import sqlite3
from flask import Flask, jsonify
import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

avisos = []

def enviar_email(assunto, corpo):
    email_remetente = "luca.silvestri@cortesiaconcreto.com.br"
    senha = "C0ncreto"
    email_destinatario = "lucas.dias@cortesiaconcreto.com.br"

    mensagem = MIMEMultipart()
    mensagem['From'] = email_remetente
    mensagem['To'] = email_destinatario
    mensagem['Subject'] = assunto

    mensagem.attach(MIMEText(corpo, 'plain'))
    
    try:
        servidor = smtplib.SMTP('smtp.office365.com', 587)
        servidor.starttls()
        servidor.login(email_remetente, senha)
        servidor.sendmail(email_remetente, email_destinatario, mensagem.as_string())
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
    finally:
        try:
            servidor.quit()
        except NameError:
            pass  # Variável não foi definida; não há conexão para fechar
        except Exception as e:
            print(f"Erro ao fechar o servidor SMTP: {e}")

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

def get_toner_level(ip):
    try:
        toner_level = 0
        response = requests.get(f'http://{ip}', verify=False, timeout=5)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        image = soup.find('img', {'class': 'tonerremain'})

        if image:
            height = int(image.get('height', 0))
            toner_level = (height * 100) / 56
            
            if toner_level <= 100:
                toners_baixos = toners_baixos + 1
                avisos.append(f"{toners_baixos}. Toner baixo da impressora: {ip}. Apenas {toner_level:.2f}%.")
        
            return f"{toner_level:.2f}%"
        else:
            return "Imagem com a classe 'tonerremain' não encontrada."
    
    except requests.exceptions.RequestException as e:
        return f"Erro ao conectar ao IP {ip}: {e}"
    
    except Exception as e:
        return f"Erro inesperado ao processar IP {ip}: {e}"

# Endpoint da API
@app.route('/toner-levels', methods=['GET'])
def toner_levels():
    impressoras = get_impressoras()
    results = []

    for unidade, setor, marca, ip in impressoras:
        toner_level = get_toner_level(ip)
        results.append({'unidade': unidade, 'setor': setor, 'marca': marca, 'ip': ip, 'toner_level': toner_level})
        
    avisos_formatados = "\n".join(avisos)
    enviar_email("Feedback de impressoras", f"{avisos_formatados}")
    return jsonify(results)


if __name__ == '__main__':
    initialize_db()
    app.run(debug=True)
