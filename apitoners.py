import sqlite3
from flask import Flask, jsonify
import requests
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import time
import requests
import threading
from openpyxl import Workbook
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import warnings

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
            pass
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

def get_impressoras():
    conn = sqlite3.connect('apitoners.db')
    cursor = conn.cursor()
    cursor.execute('SELECT unidade, setor, marca, ip FROM impressoras')
    impressoras = cursor.fetchall()
    conn.close()
    return impressoras

def get_toner_level(ip, marca):
    try:
        toner_level = 0
        if marca == "BROTHER":
            response = requests.get(f'http://{ip}', verify=False, timeout=5)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            image = soup.find('img', {'class': 'tonerremain'})

            
            height = int(image.get('height', 0))
            toner_level = (height * 100) / 56
                
            if toner_level <= 100:
                avisos.append(f"Toner baixo da impressora: {ip}. Apenas {toner_level:.2f}%.")
            
                return f"{toner_level:.2f}%"
        elif marca == "RICOH":
            options = Options()
            # options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--ignore-certificate-errors")
            
            service = Service('chromedriver.exe')  
            driver = webdriver.Chrome(service=service, options=options)
            
            toner_level = 0
            url = f"http://{ip}"
            driver.get(url)
                
            wait = WebDriverWait(driver, 10)
            elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'settingCategoryL11')))
        
            if len(elements) >= 5:
                print(elements)
                fifth_element = elements[4]
                classToner = fifth_element.text.strip()
            else:
                warnings.warn(f"Menos de 5 elementos encontrados com a classe 'settingCategoryL11' no IP {ip}")
                return "Erro: Elementos insuficientes", []
            
            if classToner == "Remaining Level 5":
                toner_level = 100
            elif classToner == "Remaining Level 4":
                toner_level = 75
            elif classToner == "Remaining Level 3":
                toner_level = 50
            elif classToner == "Remaining Level 2":
                toner_level = 25
            else:
                toner_level = 0
                
            if toner_level <= 100:
                avisos.append(f"Toner baixo da impressora: {ip}. Apenas {toner_level:.2f}%.")
                
            return f"{toner_level}%"
        elif marca == "SAMSUNG":
            options = Options()
            # options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--ignore-certificate-errors")
            
            service = Service('chromedriver.exe')  
            driver = webdriver.Chrome(service=service, options=options)
            
            toner_level = 0
            url = f"http://{ip}"
            driver.get(url)
                
            wait = WebDriverWait(driver, 10)
            elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'x-column')))
        
            elemento28 = elements[27]
            toner_level = int(elemento28.text.strip().replace('%', ''))
                
            if toner_level <= 100:
                avisos.append(f"Toner baixo da impressora: {ip}. Apenas {toner_level:.2f}%.")
                
            return f"{toner_level}%"
    
    except requests.exceptions.RequestException as e:
        return f"Erro ao conectar ao IP {ip}: {e}"
    
    except Exception as e:
        return f"Erro inesperado ao processar IP {ip}: {e}"

def gerar_planilha(impressoras, resultados):
    # Nome do arquivo da planilha
    arquivo_planilha = os.path.join(os.getcwd(), "status_impressoras.xlsx")
    
    # Criar workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Status Impressoras"
    
    # Cabeçalho
    ws.append(["Impressora (Unidade / IP)", "Status (Toner ou Erro)"])
    
    # Adicionar dados
    for i in range(len(impressoras)):
        unidade, setor, marca, ip = impressoras[i]
        ws.append([f"{unidade} ({ip})", resultados[i]])
    
    # Salvar a planilha
    wb.save(arquivo_planilha)
    print(f"Planilha gerada: {arquivo_planilha}")

@app.route('/toner-levels', methods=['GET'])
def toner_levels():
    impressoras = get_impressoras()
    resultados = []

    for unidade, setor, marca, ip in impressoras:
        toner_level = get_toner_level(ip, marca)
        resultados.append(toner_level)
    
    gerar_planilha(impressoras, resultados)
    
    avisos_formatados = "\n".join(avisos)
    enviar_email("Feedback de impressoras", f"{avisos_formatados}")
    return jsonify([{"unidade": unidade, "setor": setor, "marca": marca, "ip": ip, "status": status} 
                    for (unidade, setor, marca, ip), status in zip(impressoras, resultados)])

def requisitar_api():
    print("Executando requisição para API externa...")
    try:
        response = requests.get("http://127.0.0.1:5000/toner-levels")
        print(f"Resposta da API: {response.status_code} - {response.json()}")
    except requests.exceptions.RequestException as e:
        print(f"Erro ao conectar à API: {e}")
        
def agendar_tarefas():
    schedule.every().day.at("08:00").do(requisitar_api)
    schedule.every().day.at("16:30").do(requisitar_api)

    print("Rodando loop...")
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == '__main__':
    initialize_db()
    tarefa_thread = threading.Thread(target=agendar_tarefas)
    tarefa_thread.daemon = True
    tarefa_thread.start()
    app.run(debug=True)
