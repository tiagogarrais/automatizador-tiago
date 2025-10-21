import os
import logging
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from difflib import SequenceMatcher
import time

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do Google Sheets
GOOGLE_SHEETS_CREDENTIALS_PATH = os.getenv('GOOGLE_SHEETS_CREDENTIALS_PATH')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
SHEET_NAME = os.getenv('SHEET_NAME')

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

def parse_relatorio(relatorio_path):
    """Parse o relatório de unidades e retorna um dicionário CATMAT -> lista de opções"""
    catmat_options = {}
    with open(relatorio_path, 'r', encoding='utf-8') as f:
        for line in f:
            if "AS OPÇÕES DE UNIDADE PARA O CATMAT" in line:
                parts = line.split(" SÃO ")
                if len(parts) == 2:
                    catmat_part = parts[0].split("CATMAT ")[1]
                    options_part = parts[1].strip().strip('.')
                    # Remove aspas e divide por vírgula
                    options = [opt.strip().strip('"') for opt in options_part.split(',')]
                    # Remove string vazia se existir
                    options = [opt for opt in options if opt]
                    catmat_options[catmat_part] = options
    return catmat_options

def find_semantic_match(unidade_mapeada, options):
    """Encontra uma sugestão semanticamente compatível usando similaridade de string"""
    if not options:
        return ""

    melhor_match = None
    melhor_score = 0
    for opcao in options:
        score = SequenceMatcher(None, unidade_mapeada.lower(), opcao.lower()).ratio()
        if score > melhor_score:
            melhor_score = score
            melhor_match = opcao

    if melhor_score > 0.6:  # Threshold para considerar uma correspondência
        return melhor_match
    return ""

def main():
    # Conectar ao Google Sheets
    sheet = setup_google_sheets()

    # Parse do relatório
    relatorio_path = 'relatorio_unidades.txt'
    catmat_options = parse_relatorio(relatorio_path)
    logging.info(f"Carregadas opções para {len(catmat_options)} CATMATs")

    # Obter dados da planilha
    data = sheet.get_all_records()
    logging.info(f"Carregadas {len(data)} linhas da planilha")

    # Processar cada linha
    updates = []
    for i, row in enumerate(data, start=2):  # Começar da linha 2 (após cabeçalho)
        catmat = str(row.get('CATMAT', '')).strip()
        unidade_fornecimento = str(row.get('Unidade de Fornecimento', '')).strip()

        if catmat in catmat_options:
            options = catmat_options[catmat]
            if unidade_fornecimento not in options:
                # Encontrar sugestão
                sugestao = find_semantic_match(unidade_fornecimento, options)
                if sugestao:
                    updates.append({
                        'row': i,
                        'col': 'Unidade de Fornecimento sugerida',
                        'value': sugestao
                    })
                    logging.info(f"Linha {i}: CATMAT {catmat} - Sugerindo '{sugestao}' para '{unidade_fornecimento}'")
                else:
                    # Quando não há sugestão, listar as opções disponíveis
                    value = f"Opções disponíveis: {', '.join(options)}"
                    updates.append({
                        'row': i,
                        'col': 'Unidade de Fornecimento sugerida',
                        'value': value
                    })
                    logging.info(f"Linha {i}: CATMAT {catmat} - Nenhuma sugestão encontrada, listando opções para '{unidade_fornecimento}'")
            else:
                logging.info(f"Linha {i}: CATMAT {catmat} - Unidade '{unidade_fornecimento}' já compatível")

    # Aplicar atualizações em lote
    if updates:
        for update in updates:
            try:
                # Encontrar coluna da sugestão
                headers = sheet.row_values(1)
                col_index = headers.index('Unidade de Fornecimento sugerida') + 1  # 1-based
                sheet.update_cell(update['row'], col_index, update['value'])
                logging.info(f"Atualizada linha {update['row']}: {update['value']}")
                time.sleep(0.5)  # Delay de 0.5 segundos entre atualizações para evitar limite de quota
            except Exception as e:
                logging.error(f"Erro ao atualizar linha {update['row']}: {e}")
    else:
        logging.info("Nenhuma atualização necessária")

if __name__ == "__main__":
    main()