import os
import time
import logging
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar variáveis de ambiente
load_dotenv()
UNIDADE_MAP = {
    "Saco de 50 Kg": "Saco 50 Quilograma",
    "Saco 50 Kg": "Saco 50 Quilograma",
    "Saco 25 Kg": "Saco 25 Quilograma",
    "Saco de 25 Kg": "Saco 25 Quilograma",
    "KG": "Quilograma",
    "Litro": "Litro",
    "Folha": "Folha",
    "Quilograma": "Quilograma",
    "Embalagem 50 g": "Envelope 50 Grama",
    "Embalagem 100 g": "Envelope 100 Grama",
    "Envelope 50 g": "Envelope 50 Grama",
    "Envelope 5 g": "Envelope 5 Grama",
    "Unidade": "Unidade",
    "unidade": "Unidade",
    "Envelope 2000 sementes": "Embalagem 2000 Unidade",
    "Envelope 1000 sementes": "Envelope 1000 Unidade",
    "Envelope 400 mg": "Miligrama",
    "Pacote 500 g": "Envelope 500 Grama",
    "Embalagem 1 kg": "Envelope 1 Quilograma",
    "Embalagem 500 g": "Envelope 500 Grama",
    "Embalagem 250 g": "Envelope 250 Grama",
    "Caixa com 100 unidades": "Caixa 100 Unidade",
    "Caixa 1000 unidades": "Caixa 1000 Unidade",
    "Rolo 50 metros": "Rolo 50 Metro",
    "Metro cúbico": "Metro Cúbico",
    "Galão 3,6 litro": "Galão 3.6 Litro",
    "Rolo 500 metros": "Rolo 500 Metro",
    "Tubo 6 metros": "Tubo 6 Metro",
    "Embalagem 100 gramas": "Embalagem 100 Grama",
    "Caixa com 100 unidades": "Caixa 100 Unidade",
    "1 L": "Litro",
    "Caixa com 50 unidades": "Caixa 50 Unidade",
    "GRAMA": "Grama",

}

# Configurações do Google Sheets
GOOGLE_SHEETS_CREDENTIALS_PATH = os.getenv('GOOGLE_SHEETS_CREDENTIALS_PATH')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SHEET_NAME = os.getenv('SHEET_NAME')

# Configurações do site
SITE_USERNAME = os.getenv('SITE_USERNAME')
SITE_PASSWORD = os.getenv('SITE_PASSWORD')
SITE_URL = os.getenv('SITE_URL')
BUSCA_URL = os.getenv('BUSCA_URL')

# Configurações de processamento
START_ITEM = 1  # Item inicial (linha START_ITEM da planilha)
END_ITEM = 3   # Item final (None para processar até o fim)

def setup_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(GOOGLE_SHEETS_CREDENTIALS_PATH, scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
        logging.info("Conectado ao Google Sheets com sucesso.")
        return sheet
    except Exception as e:
        logging.error(f"Erro ao conectar ao Google Sheets: {e}")
        raise

def setup_selenium():
    service = EdgeService(executable_path='/home/tiago/Documentos/Github/automatizador-tiago/edgedriver_linux64/msedgedriver')
    driver = webdriver.Edge(service=service)
    return driver

def login_site(driver):
    driver.get(SITE_URL)
    # Clicar no link "Efetue o login novamente" se estiver presente
    try:
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.LINK_TEXT, "Efetue o login novamente")))
        driver.find_element(By.LINK_TEXT, "Efetue o login novamente").click()
        logging.info("Clicado no link 'Efetue o login novamente'.")
    except Exception as e:
        logging.warning(f"Link 'Efetue o login novamente' não encontrado ou não clicável: {e}")
    
    # Preencher login e senha automaticamente
    try:
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "txtLogin")))
        login_field = driver.find_element(By.NAME, "txtLogin")
        login_field.clear()
        login_field.send_keys(SITE_USERNAME)
        logging.info("Login preenchido.")
        
        senha_field = driver.find_element(By.NAME, "txtSenha")
        senha_field.clear()
        senha_field.send_keys(SITE_PASSWORD)
        logging.info("Senha preenchida.")
        
        # Clicar no botão "Entrar"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.br-button.is-primary')))
        entrar_button = driver.find_element(By.CSS_SELECTOR, 'button.br-button.is-primary')
        entrar_button.click()
        logging.info("Botão 'Entrar' clicado.")
        
        # Aguardar redirecionamento
        time.sleep(5)
        
      
        
     
        
    except Exception as e:
        logging.error(f"Erro ao preencher login/senha ou clicar em Entrar: {e}")
        raise
    
    input("Navegue até a página do IRP com a aba 'Itens' selecionada, clique em 'Incluir Itens' para abrir a nova janela, e pressione Enter para continuar...")
    # Assumir que a nova janela está aberta

def process_data(driver, data):
    logging.info(f"Número de abas/janelas abertas: {len(driver.window_handles)}")
    # Usar a aba atual para busca (assumir que 'Incluir Itens' já abriu a aba de busca)
    driver.get(BUSCA_URL)  # Garantir que estamos na página de busca
    logging.info("Carregada página de busca na aba atual.")

    erros = []  # Lista para coletar itens que falharam

    start_index = START_ITEM - 1
    end_index = END_ITEM if END_ITEM is not None else None
    for i, row in enumerate(data[start_index:end_index], start=START_ITEM):
        try:
            catmat = row.get('CATMAT')  # Coluna confirmada
            unidade_original = row.get('Unidade de Fornecimento')  # Coluna confirmada
            unidade = UNIDADE_MAP.get(unidade_original, unidade_original)  # Usar a unidade original se não mapeada
            logging.info(f"Processando linha {i}: CATMAT={catmat}, Unidade={unidade}")

            # Digitar CATMAT no campo de busca
            logging.info("Inserindo CATMAT...")
            try:
                # Tentar múltiplos seletores para o campo CATMAT
                catmat_field = None
                selectors = [
                    (By.CSS_SELECTOR, 'input[placeholder="Digite aqui o material ou serviço a ser pesquisado"]'),
                    (By.CSS_SELECTOR, 'input.ng-tns-c238-1.p-autocomplete-input.p-inputtext.p-component.ng-star-inserted'),
                    (By.CSS_SELECTOR, 'input[aria-autocomplete="list"]'),
                    (By.XPATH, '//input[@type="text" and contains(@placeholder, "material")]'),
                    (By.CSS_SELECTOR, 'input[type="text"]')  # Último recurso
                ]
                for by, selector in selectors:
                    try:
                        WebDriverWait(driver, 30).until(EC.presence_of_element_located((by, selector)))
                        catmat_field = driver.find_element(by, selector)
                        logging.info(f"Campo CATMAT encontrado com seletor: {by} - {selector}")
                        break
                    except:
                        continue
                if not catmat_field:
                    logging.error("Campo CATMAT não encontrado com nenhum seletor.")
                    raise Exception("Campo CATMAT não encontrado.")
                catmat_field.clear()
                catmat_field.send_keys(catmat)
            except Exception as e:
                logging.error(f"Erro ao inserir CATMAT: {e}")
                raise

            # Pressionar Enter para pesquisar
            logging.info("Pressionando Enter para pesquisar...")
            catmat_field.send_keys(Keys.RETURN)

            # Aguardar a página carregar completamente
            time.sleep(1)
            logging.info("Selecionando unidade...")
            try:
                WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'select.ng-untouched.ng-pristine.ng-valid')))
            except Exception as e:
                logging.warning(f"CATMAT '{catmat}' não encontrado ou dropdown de unidade não apareceu: {e}. Pulando para próxima linha.")
                continue
            unidade_select = Select(driver.find_element(By.CSS_SELECTOR, 'select.ng-untouched.ng-pristine.ng-valid'))
            try:
                unidade_select.select_by_visible_text(unidade)
            except Exception as e:
                logging.error(f"Erro ao selecionar '{unidade}' para linha {i}: {e}")
                opcoes = [opt.text for opt in unidade_select.options if opt.text.strip()]
                logging.info(f"Opções disponíveis no dropdown: {opcoes}")
                erros.append({
                    'linha': i,
                    'catmat': catmat,
                    'unidade_original': unidade_original,
                    'unidade_mapeada': unidade,
                    'opcoes': opcoes
                })
                continue  # Pular para a próxima linha

            # Clicar em adicionar
            logging.info("Clicando em 'Adicionar'...")
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Adicionar"]')))
            add_button = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Adicionar"]')
            add_button.click()

            logging.info(f"Linha {i} processada com sucesso.")
            time.sleep(2)  # Pausa para evitar bloqueios
        except Exception as e:
            logging.error(f"Erro ao processar linha {i}: {e}")
            continue  # Pular para a próxima linha

    # Log dos itens que falharam
    if erros:
        logging.info("Itens que falharam na seleção de unidade:")
        for erro in erros:
            logging.info(f"Linha {erro['linha']}: CATMAT={erro['catmat']}, Unidade original={erro['unidade_original']}, Mapeada={erro['unidade_mapeada']}, Opções disponíveis={erro['opcoes']}")
    else:
        logging.info("Nenhum erro na seleção de unidades.")

    # Não fechar a aba, para conferência manual

def main():
    sheet = setup_google_sheets()
    data = sheet.get_all_records()

    driver = setup_selenium()
    try:
        login_site(driver)
        process_data(driver, data)
        input("Processamento concluído. Verifique os itens adicionados na janela. Pressione Enter para fechar o navegador...")
        driver.quit()
    except Exception as e:
        logging.error(f"Erro no main: {e}")
        input("Erro ocorrido. Pressione Enter para fechar o navegador...")
        driver.quit()

if __name__ == "__main__":
    main()