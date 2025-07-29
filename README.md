# Levantamento de Débito Estadual (São Paulo)

#### Descrição Geral 📌

Automação do processo de levantamento de débitos estaduais das empresas do Estado de São Paulo.

A solução integra:
- **Google Sheets**: para leitura das informações das empresas;
- **Google Drive**: para armazenar os arquivos PDF gerados com os débitos;
- **Gmail**: para envio automático das informações por e-mail.

A integração com os serviços do Google é feita via APIs oficiais, e a automação do fluxo é executada com Selenium.


## Ferramentas utilizadas

- ![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
- ![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=for-the-badge&logo=selenium&logoColor=white)
- ![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
- ![Google Drive](https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=google-drive&logoColor=white)
- ![Gmail](https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white)


## Passo a passo da automação

O arquivo `robo_sp.py` é o script principal da automação, responsável por coordenar todo o fluxo de consulta, geração, salvamento e envio dos débitos estaduais das empresas cadastradas.

### 🤖 O que esse script faz?

1.  **Carrega dados da planilha de controle:**
    - A planilha do Google Sheets é acessada via API usando a conta de serviço.
    - Os dados da aba **Estadual (SP)** são carregados, contendo nomes, responsáveis e CNPJs das empresas.

2.  **Acessa o portal da Secretaria da Fazenda do Estado de São Paulo com Selenium:**
    - Para cada empresa:
      - Realiza login no portal, preenche o CNPJ e consulta a situação fiscal.
      - Se houver débitos, gera um PDF da página (salvo localmente com nome estruturado).
      - Se não houver débitos ou acesso, atualiza o status na planilha.

3. **Salva o arquivo PDF no Google Drive:**
   - Cada arquivo é salvo em uma pasta no Google Drive, organizada por mês e responsável.
   - Caso a pasta do mês ainda não exista, ela é criada automaticamente e compartilhada com o responsável via e-mail.

4. **Envia o link da pasta por e-mail:**
   - Para cada responsável com arquivos salvos, é enviado um e-mail com o link da pasta.
   - O e-mail é baseado em um template HTML e preenchido dinamicamente com o nome e link.


## Funções Auxiliares

### 📁 servico_google.py

#### 🔧 acessando_sheets()
Estabelece conexão com a API do Google Sheets utilizando as credenciais de uma conta de serviço.
- Lê o caminho do arquivo .json e escopos do .env.
- Retorna um cliente autenticado do gspread.

#### 🔧 acessando_drive()
Cria uma conexão autenticada com a API do Google Drive.
- Utiliza a conta de serviço e escopos definidos no .env.
- Retorna o serviço do Google Drive pronto.



### 📁 envio_drive.py

#### 📂 compartilhar_pasta(drive_service, pasta_id, seu_email)
Compartilha uma pasta específica do Google Drive com um e-mail, concedendo permissão de escrita.  
Parâmetros:
- `drive_service`: serviço autenticado do Google Drive.
- `pasta_id`: ID da pasta que será compartilhada.
- `seu_email`: e-mail do usuário que receberá acesso.

#### 💾 salvar_drive(caminho_arquivo, resp, nome_arquivo)
Salva um arquivo no Google Drive, organizando-o em uma estrutura de pastas do tipo:  
Levantamento de Débito: MM/yyyy → Responsável → arquivo.pdf  
Parâmetros:
- `caminho_arquivo`: Caminho do arquivo PDF local.
- `resp`: Responsável pela empresa.
- `nome_arquivo`: Nome do arquivo para upload no Drive.


### 📁 envio_email.py

#### 📄 carregar_template(nome, link)
Carrega um template HTML de e-mail e o formata para cada responsável.  
Parâmetros:
- `nome`: Nome do responsável.
- `link`: Link da pasta no Drive que contém os arquivos da pessoa.

#### ✉️ criar_email(destinatario, assunto, nome, link)
Cria um e-mail no formato raw (base64) pronto para ser enviado via API Gmail e preenche o conteúdo HTML com `carregar_template()`.  
Parâmetros:
- `destinatario`: E-mail do destinatário.
- `assunto`: Assunto usado no envio de e-mail.
- `nome`: Nome do responsável.
- `link`: Link da pasta no Drive que contém os arquivos da pessoa.

#### 📬 enviar(destinatario, assunto, nome, link)
Envia o e-mail para o destinatário com o assunto e conteúdo definidos, criando o email com `criar_email()`.  
Parâmetros:
- `destinatario`: E-mail do destinatário.
- `assunto`: Assunto usado no envio de e-mail.
- `nome`: Nome do responsável.
- `link`: Link da pasta no Drive que contém os arquivos da pessoa.

## Autor

- [@RafaelCostrov](https://github.com/RafaelCostrov)
