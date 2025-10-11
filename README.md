# Automatizador de Preenchimento Web

Este script automatiza o preenchimento de dados de uma planilha Google Sheets em um sistema web de IRP (Intenção de Registro de Preços) usando Selenium e Edge. Inclui automação completa do login e navegação até a página de inclusão de itens.

## Pré-requisitos

- Python 3.8+
- Conta Google com API habilitada
- Arquivo `credentials.json` do Google Cloud Console
- Navegador Microsoft Edge e WebDriver correspondente

## Configuração

1. **Arquivo .env**: Copie `.env.example` para `.env` e preencha com suas credenciais e configurações (não versionado).

   - `GOOGLE_SHEETS_CREDENTIALS_PATH`: Caminho para `credentials.json`
   - `SPREADSHEET_ID`: ID da planilha (do link fornecido)
   - `SHEET_NAME`: Nome da aba da planilha (ex: "Detalhamento por itens")
   - `SITE_USERNAME` e `SITE_PASSWORD`: Credenciais do site
   - `SITE_URL`: URL de login do site
   - `BUSCA_URL`: URL da página de busca de materiais

2. **Credenciais Google**:

   - Crie um projeto no Google Cloud Console.
   - Habilite a Google Sheets API.
   - Crie credenciais de serviço e baixe o `credentials.json`.

3. **Instalação**:
   - Crie um ambiente virtual: `python -m venv .venv`
   - Ative: `source .venv/bin/activate` (Linux/Mac) ou `.venv\Scripts\activate` (Windows)
   - Instale dependências: `pip install -r requirements.txt`

4. **WebDriver**:
   - O script usa o Edge WebDriver localizado em `edgedriver_linux64/msedgedriver`. Certifique-se de que está compatível com sua versão do Edge.

## Como Usar

1. Configure o `.env` e `credentials.json`.
2. Execute: `python automatizador.py`
3. O script fará login automaticamente, navegará para a página de IRP, e aguardará sua interação manual para "Incluir Itens".
4. Após pressionar Enter, processará os itens da planilha, inserindo CATMAT, selecionando unidades e adicionando ao carrinho.

**Fluxo Automático**:
- Login no site
- Clicar no botão da página 2 (se disponível)
- Clicar em "Intenção de Registro de Preços"
- Clicar em "IRP"
- Clicar em "Abrir Intenção de Registro de Preços"
- Aguardar interação manual para abrir a janela de inclusão de itens

**Configurações de Processamento**:
- `START_ITEM`: Linha inicial da planilha (padrão: 1)
- `END_ITEM`: Linha final (None para processar até o fim)

Adicione mapeamentos no `UNIDADE_MAP` conforme necessário para unidades não padrão.

## Logs e Depuração

Os logs são exibidos no console com nível INFO. Em caso de erro, verifique os logs para identificar problemas em seletores ou elementos não encontrados.

## Segurança e Git

- Arquivos sensíveis (`credentials.json`, `.env`) são ignorados pelo `.gitignore`.
- Não compartilhe credenciais.
- Use em conformidade com os termos do site e leis aplicáveis.

## Estrutura do Projeto

- `automatizador.py`: Script principal
- `requirements.txt`: Dependências Python
- `README.md`: Esta documentação
- `.gitignore`: Arquivos ignorados pelo Git
- `edgedriver_linux64/`: WebDriver do Edge

## Ajustes Necessários

- Verifique seletores do site se houver mudanças no DOM.
- Atualize mapeamentos de unidades conforme surgirem novos casos.
