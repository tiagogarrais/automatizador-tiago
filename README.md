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
   - `SPREADSHEET_ID`: ID da planilha (apenas o ID, não o link completo. Ex: se o link é https://docs.google.com/spreadsheets/d/1abc123.../edit, use apenas "1abc123...")
   - `SHEET_NAME`: Nome da aba da planilha (ex: "Detalhamento por itens")
   - `SITE_USERNAME` e `SITE_PASSWORD`: Credenciais do site
   - `SITE_URL`: URL de login do site
   - `BUSCA_URL`: URL da página de busca de materiais

2. **Credenciais Google**:

   - Crie um projeto no [Google Cloud Console](https://console.cloud.google.com/).
   - Habilite a Google Sheets API.
   - Crie credenciais de serviço e baixe o `credentials.json`.

   **Passos detalhados para configurar as credenciais:**

   1. Acesse o [Google Cloud Console](https://console.cloud.google.com/) e faça login com sua conta Google.
   2. Crie um novo projeto ou selecione um existente.
   3. No menu lateral esquerdo, clique em **APIs e Serviços** > **Biblioteca**.
   4. Procure por "Google Sheets API", selecione-a e clique em **Habilitar**.
   5. Volte ao menu lateral, clique em **APIs e Serviços** > **Credenciais**.
   6. Clique em **+ Criar credenciais** > **Conta de serviço**.
   7. Preencha o nome da conta de serviço (ex: "automatizador-sheets"), descrição opcional, e clique em **Criar e continuar**.
   8. Atribua um papel (opcional para este caso, pode pular), clique em **Concluído**.
   9. Na lista de contas de serviço, clique na conta criada.
   10. Vá para a aba **Chaves**, clique em **Adicionar chave** > **Criar nova chave**.
   11. Selecione o tipo **JSON** e clique em **Criar**. O arquivo `credentials.json` será baixado automaticamente.
   12. Salve o arquivo `credentials.json` na raiz do projeto (não versionado).
   13. Abra a planilha Google Sheets que o script irá usar.
   14. Clique em "Compartilhar" (ícone de pessoa com +).
   15. No campo de e-mail, adicione o e-mail da conta de serviço (encontrado no `credentials.json` no campo "client_email", ex: "automatizador-sheets@seu-projeto.iam.gserviceaccount.com").
   16. Defina as permissões como "Editor" e clique em "Enviar" ou "Compartilhar".

3. **Instalação**:

   - Se estiver em Debian/Ubuntu, instale o pacote python3-venv (substitua pela versão do Python, ex: python3.12-venv): `sudo apt install python3-venv`
   - Crie um ambiente virtual: `python3 -m venv .venv`
   - Ative: `source .venv/bin/activate` (Linux/Mac) ou `.venv\Scripts\activate` (Windows)
   - Instale dependências: `pip install -r requirements.txt`

4. **WebDriver**:
   - O script usa o Edge WebDriver localizado em `edgedriver_linux64/msedgedriver`. Certifique-se de que está compatível com sua versão do Edge.
   - Baixe o WebDriver do Edge em: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/

## Como Usar

**Formato da Planilha:**

- A planilha Google Sheets deve conter obrigatoriamente as colunas `CATMAT` e `Unidade de Fornecimento` (escritas exatamente dessa forma, sem variações).

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
