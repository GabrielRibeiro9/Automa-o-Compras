import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dotenv import load_dotenv
import undetected_chromedriver as uc

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Credenciais de login e detalhes da Ultramsg obtidos do .env
ULTRAMSG_EMAIL = os.getenv('EMAIL')
ULTRAMSG_PASSWORD = os.getenv('SENHA')
ULTRAMSG_INSTANCE_ID = os.getenv('ULTRAMSG_INSTANCE') # Apenas o ID da instância (ex: "123340")

def automatizar_extensao_ultramsg():
    chrome_options = uc.ChromeOptions()
    
    # ==================== MUDANÇA PARA MODO SILENCIOSO (HEADLESS) ====================
    chrome_options.add_argument('--headless') # <-- LINHA ATIVADA
    # ==============================================================================
    
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920x1080') # Mantenha isso para evitar problemas de layout
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--incognito')

    driver = None 
    try:
        driver = uc.Chrome(options=chrome_options)

        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            '''
        })

        print("Abrindo página de login da Ultramsg (em modo silencioso)...")
        driver.get("https://user.ultramsg.com/signin.php")

        time.sleep(5)

        if not (ULTRAMSG_EMAIL and ULTRAMSG_PASSWORD):
            print("❌ Credenciais de login (EMAIL/SENHA) não configuradas no .env. Abortando.")
            return

        print("Tentando fazer login...")
        try:
            email_field = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.NAME, "email")))
            password_field = driver.find_element(By.NAME, "password")
            login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")

            email_field.send_keys(ULTRAMSG_EMAIL)
            password_field.send_keys(ULTRAMSG_PASSWORD)
            login_button.click()
            print("Credenciais enviadas. Aguardando redirecionamento para o painel...")

            WebDriverWait(driver, 30).until(EC.url_contains("index.php"))
            print("✅ Login bem-sucedido e redirecionado para o dashboard.")

        except TimeoutException:
            print(f"❌ Erro: Tempo esgotado esperando o redirecionamento para o painel após o login.")
            driver.save_screenshot('screenshot_erro_login.png')
            return
        except Exception as e:
            print(f"❌ Erro inesperado durante o login: {e}")
            driver.save_screenshot('screenshot_erro_login.png')
            return

        print("Navegando para a página de instâncias...")
        driver.get("https://user.ultramsg.com/app/instances/instances.php")
        
        print("Aguardando a tabela de instâncias carregar...")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.TAG_NAME, "table")))
        print("✅ Tabela de instâncias carregada.")
        
        instance_row = None
        try:
            print(f"Procurando pela linha da instância com ID: #{ULTRAMSG_INSTANCE_ID}...")
            
            instance_row_xpath = f"//td[contains(., '{ULTRAMSG_INSTANCE_ID}')]/ancestor::tr"
            instance_row = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, instance_row_xpath))
            )
            print(f"✅ Instância #{ULTRAMSG_INSTANCE_ID} encontrada.")

            status_element_xpath = ".//span[contains(@class, 'badge-danger') and normalize-space()='Parada']"
            extend_button_xpath = f".//button[contains(@onclick, \"extend_trial('{ULTRAMSG_INSTANCE_ID}')\")]"

            try:
                wait_in_row = WebDriverWait(instance_row, 5)
                
                status_parada = wait_in_row.until(EC.visibility_of_element_located((By.XPATH, status_element_xpath)))
                print("Status da instância é 'Parada'.")

                extend_button = wait_in_row.until(EC.element_to_be_clickable((By.XPATH, extend_button_xpath)))
                print("Botão 'Estender o período de testes' encontrado. Clicando...")
                extend_button.click()
                
                try:
                    print("Aguardando o pop-up de confirmação...")
                    confirm_button_xpath = "//button[contains(., 'confirme')]"
                    
                    confirm_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, confirm_button_xpath))
                    )
                    
                    print("Botão 'confirme' encontrado. Clicando...")
                    confirm_button.click()
                    
                    print("🎉 PROCESSO FINALIZADO COM SUCESSO! 🎉")
                    time.sleep(5) 
                    driver.save_screenshot('screenshot_sucesso_final.png')
                    print("Screenshot 'screenshot_sucesso_final.png' salvo para comprovação.")

                except TimeoutException:
                    print("❌ Não foi possível encontrar ou clicar no botão 'confirme' do pop-up.")
                    driver.save_screenshot('screenshot_erro_confirmacao.png')

            except TimeoutException:
                print("✅ A instância não está 'Parada' ou o botão de extensão não está disponível. Nenhuma ação necessária.")

        except TimeoutException:
            print(f"❌ Não foi possível encontrar a linha da instância #{ULTRAMSG_INSTANCE_ID} na tabela.")
            driver.save_screenshot('screenshot_erro_instancia.png')

    except Exception as e:
        print(f"Ocorreu um erro geral e inesperado durante a automação: {e}")
        if driver:
            driver.save_screenshot('screenshot_erro_geral.png')
            print("Screenshot 'screenshot_erro_geral.png' salvo para análise.")
    finally:
        if driver:
            print("Fechando navegador...")
            driver.quit()

if __name__ == "__main__":
    automatizar_extensao_ultramsg()