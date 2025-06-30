from datetime import datetime
import shutil
import os
import pandas as pd
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import pywhatkit as kit
import requests
import sys
import tempfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys

load_dotenv()

# Obter o email e a senha a partir das variáveis de ambiente
email = os.getenv('EMAIL')
senha = os.getenv('SENHA')

def verificar_ultramsg():
# Configuração para rodar o Chrome em modo headless (sem janela)
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Rodar sem interface gráfica
    chrome_options.add_argument('--no-sandbox')  # Para evitar problemas em alguns ambientes
    chrome_options.add_argument('--disable-dev-shm-usage')

# Inicializa o WebDriver para o Chrome
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Acesse o site do UltraMsg
    driver.get('https://user.ultramsg.com/signin.php')

 # Preencher o campo de E-mail
    email_input = driver.find_element(By.CSS_SELECTOR, 'input[name="email"]')  # Identificado pelo 'name="email"'
    email_input.send_keys(email)  # Enviar o email

    # Preencher o campo de Senha
    senha_input = driver.find_element(By.CSS_SELECTOR, 'input[name="password"]')  # Substitua pelo campo correto de senha
    senha_input.send_keys(senha)  # Enviar a senha

    # Clicar no botão de login 
    login_button = driver.find_element(By.CSS_SELECTOR, 'button[name="signin"]')  # Identificado pelo 'name="signin"'
    login_button.click()

    time.sleep(3)

    
    try:
        # Localize o elemento com o texto 'Parada'
        media_body = driver.find_element(By.XPATH, "//div[contains(@class, 'media-body') and contains(text(), 'Parada')]")
        media_body.click()
        print("Clique no elemento 'Parada' realizado com sucesso.")
            
        # Espera a troca de aba e encontra o botão 'Estender Período'
        time.sleep(2)  # Ajuste o tempo conforme necessário
            
        # Troca para a nova aba
        driver.switch_to.window(driver.window_handles[1])  # Assume que a nova aba é a segunda
            
        # Localiza o botão "Estender o Período de Teste" usando o onclick
        estender_button = driver.find_element(By.XPATH, "//button[contains(@onclick, 'extend_trial')]")
        estender_button.click()
        print("Período de teste estendido com sucesso.")

    except Exception as e:
        print(f"Erro ao clicar no elemento 'Parada' ou estender o período: {e}")

    # Fechar o navegador
    driver.quit()

# Execute a função
verificar_ultramsg()    

def resource_path(relative_path):
    """Retorna o caminho absoluto, compatível com PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Função para dividir mensagens grandes em pedaços menores
def dividir_mensagem(mensagem, limite_caracteres=4000):
    """Divide a mensagem em partes menores, com no máximo limite_caracteres cada."""
    return [mensagem[i:i+limite_caracteres] for i in range(0, len(mensagem), limite_caracteres)]

def enviar_whatsapp_ultramsg(mensagem, destinatarios):
    instance_id = "instance123340"
    token = "wo8yary0bdrlo4xy"

    for numero in destinatarios:
        url = f"https://api.ultramsg.com/{instance_id}/messages/chat"
        data = {
            "token": token,
            "to": numero,
            "body": mensagem
        }
        response = requests.post(url, data=data)
        if response.status_code == 200:
            print(f"Mensagem enviada para {numero}: {response.status_code} - {response.text}")
        else:
            print(f"Erro ao enviar para {numero}: {response.status_code} - {response.text}")


# Caminho completo para a planilha original
caminho_original = os.path.join(os.getcwd(), r"R:\Manutenção\PLANILHA COMRAS (MANUTENÇÃO-FERRAMENTARIA)-2025.xlsx")
caminho_temp = os.path.join(tempfile.gettempdir(), "PLANILHA_COMRAS_TEMP.xlsx")

# Verifica se o caminho existe
if not os.path.exists(caminho_original):
    print("Caminho não encontrado:", caminho_original)
    sys.exit()

# Se o caminho existir, prossiga normalmente
shutil.copy(caminho_original, caminho_temp)

# Verifica se o arquivo original possui copia
if os.path.exists(caminho_original):
    shutil.copy(caminho_original, caminho_temp)
    print("Planilha copiada com sucesso.")
else:
    print("Arquivo original nao encontrado")
    sys.exit() 

# Le a planilha temporaria
df = pd.read_excel(caminho_temp, sheet_name="Plan1") 

# Remove quebras de linha e espacos extras nos nomes das colunas
df.columns = df.columns.str.replace("\n", " ").str.strip()

# Converte a coluna de data para tipo datetime
df["DATA ENTREGA"] = pd.to_datetime(df["DATA ENTREGA NA HELPTECH PREVISTA"], errors="coerce")

# Pega a data atual
hoje = pd.to_datetime(datetime.now().date())

# Define o intervalo de vencimento (proximos 5 dias)
data_limite = hoje + pd.Timedelta(days=5)

# Filtra pedidos com vencimentos nos proximos 5 dias
vencendo = df[
    df["STATUS DO PEDIDO"].isin(["REALIZADO", "EM ABERTO"]) &
    (df["DATA ENTREGA"] >= hoje) &
    (df["DATA ENTREGA"] <= data_limite)
]

vencendo_whatsapp = df[  
    df["STATUS DO PEDIDO"].isin(["REALIZADO"]) &  
    (df["DATA ENTREGA"] >= hoje) &
    (df["DATA ENTREGA"] <= data_limite)
]

print(f"\nTotal de pedidos vencendo nos próximos 5 dias: {len(vencendo)}")
print(vencendo[["N° DA SOLICITAÇÃO", "PEDIDO", "DATA ENTREGA", "FORNEDOR DESIGNADO"]])


   # Agrupar os pedidos por fornecedor
for fornecedor, grupo in vencendo_whatsapp.groupby("FORNEDOR DESIGNADO"):
    mensagem_whatsapp = f"Bom dia, {fornecedor}:\n\n Segue abaixo os pedidos com previsão de entrega nos próximos dias, poderia me confirmar se o prazo esta mantido?\n\n"
   
    for _, row in grupo.iterrows():
         mensagem_whatsapp += (
            f"Pedido: {row['PEDIDO']}\n"
            f"Item: {row['DESCRIÇÃO DO ITEM']}\n"
            f"Entrega prevista: {row['DATA ENTREGA'].strftime('%d/%m/%Y')}\n\n"
            )
    mensagens_divididas = dividir_mensagem(mensagem_whatsapp)

    # Enviar cada mensagem dividida
    for msg in mensagens_divididas:
        destinatarios = ["+5519996386684"] 
        # Envia a mensagem via UltraMsg
        enviar_whatsapp_ultramsg(msg, destinatarios)
        print(f"Mensagem enviada com sucesso{fornecedor}.")


# Exibe a mensagem formatada no console
print("\nMensagem WhatsApp:\n")
print(mensagem_whatsapp)

# Filtra os pedidos atrasados com status REALIZADO e EM ABERTO
atrasados = df[
    df["STATUS DO PEDIDO"].isin(["REALIZADO","EM ABERTO"]) &
    (df["DATA ENTREGA"] < hoje)]                          
    
# Exibe os pedidos atrasados
print(atrasados[["N° DA SOLICITAÇÃO", "PEDIDO", "DATA ENTREGA", "FORNEDOR DESIGNADO"]])
# Imprime total de pedidos atrasados
print(f"\nTotal de pedidos atrasados: {len(atrasados)}")

# Calcula dias de atraso
atrasados["DIAS DE ATRASO"] = (hoje - atrasados["DATA ENTREGA"]).dt.days

# Ordena do maior atraso para o menor
atrasados = atrasados.sort_values(by="DIAS DE ATRASO", ascending=False).reset_index(drop=True) 

# Exibe a tabela atualizada
print(atrasados[["N° DA SOLICITAÇÃO", "PEDIDO", "DATA ENTREGA", "DIAS DE ATRASO", "FORNEDOR DESIGNADO"]])

# Gerar PDF com os dados
pdf = FPDF()
pdf.add_page()

# Logo
pdf.image(resource_path("logo_helptech.png"), x=10, y=5, w=80)


# Titulo
pdf.set_font("Arial", "B", 14)
pdf.ln(25)
pdf.cell(0, 10, f"Relatório de Pedidos em Atraso - {hoje.strftime('%d/%m/%Y')}", ln=True )


# Cabecalho da tabela
pdf.set_font("Arial", "B", 10)
pdf.cell(10, 8, "Nº", 1)
pdf.cell(35, 8, "Solicitação", 1)
pdf.cell(25, 8, "Pedido", 1)
pdf.cell(30, 8, "Data Prevista", 1)
pdf.cell(25, 8, "Atraso (dias)", 1)
pdf.cell(65, 8, "Fornecedor", 1)
pdf.ln()

# Remove duplicatas com base nas colunas principais
atrasados = atrasados.drop_duplicates(subset=[
    "N° DA SOLICITAÇÃO", "PEDIDO", "DATA ENTREGA", "DIAS DE ATRASO", "FORNEDOR DESIGNADO"
])

def contar_pedidos(pedido_str):
    return len(str(pedido_str).replace("\n", ", ").split(","))

linha_por_solicitacao ={}

# Conteudo da tabela
pdf.set_font("Arial", "", 8)


for numero, (_, row) in enumerate(atrasados.iterrows(), start=1):
    pedidos = str(row["PEDIDO"]).replace("\n", ", ").split(",")
    pedido_texto = f"{len(pedidos)} pedidos" if len(pedidos) > 1 else pedidos[0].strip()
    
    pdf.cell(10, 6, str(numero), border=1, align='C')
    pdf.cell(35, 6, str(row["N° DA SOLICITAÇÃO"]), border=1, align='C')
    pdf.cell(25, 6, pedido_texto, border=1, align='C')
    pdf.cell(30, 6, row["DATA ENTREGA"].strftime('%d/%m/%Y'), border=1, align='C')
    pdf.cell(25, 6, str(row["DIAS DE ATRASO"]), border=1, align='C')
    pdf.cell(65, 6, str(row["FORNEDOR DESIGNADO"]), border=1)
    pdf.ln()

    linha_por_solicitacao[row["N° DA SOLICITAÇÃO"]] = numero

# Total no rodape
pdf.ln(5)
pdf.set_font("Arial", "B", 11)
pdf.cell(0, 10, f"Total de pedidos em atraso: {len(atrasados)}", ln=True)

# Detalhamento de multiplos pedidos
pdf.add_page()
pdf.set_font("Arial", "B", 11)
pdf.cell(0, 10, "Detalhamento de Pedidos por Solicitação:", ln=True)
pdf.set_font("Arial", "", 8)

agrupado = atrasados.groupby("N° DA SOLICITAÇÃO")

for solicitacao, grupo in agrupado:
    if len(grupo) > 1:
        linha = linha_por_solicitacao.get(solicitacao, "??")
        fornecedores = ", ".join(sorted(set(grupo["FORNEDOR DESIGNADO"].astype(str))))
        pedidos = ", ".join(sorted(set(", ".join(grupo["PEDIDO"].astype(str).str.replace("\n", ", ").str.replace("\r", "").str.strip()).split(", ")))
)
        
        pdf.set_font("Arial", "B", 8)
        pdf.multi_cell(0, 6, f"Linha {linha} - Solicitação {solicitacao} - Fornecedor(es): {fornecedores}")
        pdf.set_font("Arial", "", 8)
        pdf.multi_cell(0, 6, f"Pedidos: {pedidos}")
        pdf.ln(1)


# Salvar PDF temporariamente
pdf_path = os.path.join(tempfile.gettempdir(), "relatorio_atrasos.pdf")
pdf.output(pdf_path)


# Dados direto no código 
remetente = "enviaremails05@gmail.com"
senha = "gfnw gzdi cuqk edkx"
destinatario = "rebeca.dezotti@helptech.ind.br"


# Cria a mensagem
msg = EmailMessage()
msg["Subject"] = "Relatório diário de pedidos em atraso - HelpTech"
msg["From"] = remetente
msg["To"] = destinatario
msg.set_content(
    f"""Prezada Rebeca,

Segue em anexo o relatório automático com os pedidos de compra que estão em atraso, gerado no dia {hoje.strftime('%d/%m/%Y')}.

Atenciosamente,
Automacao Python
"""
)

# Anexa o PDF
with open(pdf_path, "rb") as f:
    msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename="relatorio_atrasos.pdf" )
# Envia o email 
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(remetente, senha)
    smtp.send_message(msg)

print("E-mail enviado com sucesso!")

# Remove arquivos temporarios
if os.path.exists(caminho_temp):
    os.remove(caminho_temp)

if os.path.exists(pdf_path):
    os.remove(pdf_path)