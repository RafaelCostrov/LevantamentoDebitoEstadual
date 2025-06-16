import datetime
import os
import time
import pyautogui
import base64
from dotenv import load_dotenv
from servico_google import acessando_sheets
from envio_drive import salvar_drive
from envio_email import enviar
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.print_page_options import PrintOptions
from selenium.common.exceptions import NoSuchElementException


load_dotenv()
client = acessando_sheets()
SHEET_LEVANTAMENTO_DEBITO = os.getenv('SHEET_LEVANTAMENTO_DEBITO')
ss = client.open_by_key(SHEET_LEVANTAMENTO_DEBITO)
estadual = ss.worksheet('Estadual (SP)')

coluna_nome = estadual.col_values(2)[1:]
coluna_cnpj = estadual.col_values(3)[1:]
coluna_responsavel = estadual.col_values(4)[1:]

chrome_options = Options()
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--incognito")
chrome_options.add_experimental_option(
    "excludeSwitches", ["enable-automation"])

SERVICE = os.getenv('SERVICE')
servico = Service(SERVICE)

navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.switch_to.window(navegador.window_handles[0])
navegador.get(
    "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx/servicos/pfe/Paginas/Sobre.aspx/cadesp_55.shtm")
navegador.maximize_window()

botao_contabilista = navegador.find_element(
    By.XPATH, '//*[@id="ConteudoPagina_rdoListPerfil"]/label[2]')
botao_contabilista.click()

botao_acesso = navegador.find_element(
    By.XPATH, '//*[@id="ConteudoPagina_ImageButton1"]')
botao_acesso.click()
USER = os.getenv('USER')
SENHA = os.getenv('SENHA')

WebDriverWait(navegador, 30).until(
    EC.presence_of_element_located(
        (By.XPATH, '//*[@id="txtUsuario"]')
    )
)

campo_usuario = navegador.find_element(By.XPATH, '//*[@id="txtUsuario"]')
campo_usuario.send_keys(USER)

campo_senha = navegador.find_element(By.XPATH, '//*[@id="txtSenha"]')
campo_senha.send_keys(SENHA)

time.sleep(2)

iframe = WebDriverWait(navegador, 10).until(
    EC.frame_to_be_available_and_switch_to_it(
        (By.XPATH,
         "/html/body/form/div[3]/div[2]/div/div[3]/div/div/div/iframe")
    )
)

checkbox = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="recaptcha-anchor"]/div[1]'))
)

actions = ActionChains(navegador)
actions.move_to_element(checkbox).pause(3).click().perform()

time.sleep(10)

navegador.switch_to.default_content()
acessar = navegador.find_element(By.XPATH, '//*[@id="btnClaimsIdentity"]')
acessar.click()

time.sleep(2)

WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="contentRow"]/div[1]/h2'))
).click()

conta_fiscal = navegador.find_element(
    By.XPATH, '//*[@id="contentRow"]/ul[1]/li[5]/h4/a')
conta_fiscal.click()

WebDriverWait(navegador, 30).until(
    EC.presence_of_element_located(
        (By.XPATH, '/html/body/form/div[4]/div[2]/h2/span')
    )
)

opcoes_conta = navegador.find_element(
    By.XPATH, '//*[@id="menuBar"]/ul/li[1]/a')
opcoes_conta.click()

time.sleep(2)

opcao_situacao = navegador.find_element(
    By.XPATH, '//*[@id="menuBar:submenu:2"]/li[4]/a')
opcao_situacao.click()

time.sleep(1)

WebDriverWait(navegador, 30).until(
    EC.presence_of_element_located(
        (By.XPATH, '//*[@id="Form1"]/div[4]/div[2]/h2')
    )
)
input_identificacao = navegador.find_element(
    By.XPATH, '//*[@id="MainContent_ddlContribuinte"]')
input_identificacao.click()
time.sleep(1)
opcao_cnpj = navegador.find_element(
    By.XPATH, '//*[@id="MainContent_ddlContribuinte"]/option[2]')
opcao_cnpj.click()
time.sleep(1)

responsaveis_com_pasta = set()

for i, (cnpj, resp, nome) in enumerate(zip(coluna_cnpj, coluna_responsavel, coluna_nome), start=2):
    dia = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta = os.path.join(os.path.expanduser("~"), "Temp")
    os.makedirs(pasta, exist_ok=True)
    caminho_arquivo = os.path.join(pasta, f"{nome} - {cnpj} - {dia}.pdf")
    caminho_arquivo = os.path.normpath(caminho_arquivo)
    nome_arquivo = f"{nome} - {cnpj} - {dia}.pdf"

    input_cnpj = navegador.find_element(By.XPATH,
                                        '//*[@id="MainContent_txtCriterioConsulta"]')
    input_cnpj.send_keys(Keys.CONTROL + "a")
    input_cnpj.send_keys(Keys.DELETE)
    input_cnpj.send_keys(cnpj)
    time.sleep(1)
    botao_consultar = navegador.find_element(By.XPATH,
                                             '//*[@id="MainContent_btnConsultar"]')
    botao_consultar.click()
    time.sleep(1)

    try:
        mensagem_erro = navegador.find_element(
            By.XPATH, '//*[@id="MainContent_lblMensagemDeErro"]')
        if mensagem_erro.text == 'Contabilista não é contador ativo da empresa. Não pode acessar o serviço':
            estadual.update_cell(i, 5, 'Sem procuração')
        else:
            estadual.update_cell(i, 5, mensagem_erro.text)
    except NoSuchElementException:
        status_texto = navegador.find_element(By.XPATH,
                                              '//*[@id="MainContent_lblMsgErroResultado"]')
        if status_texto.text == 'Não há débitos não inscritos para o contribuinte.':
            estadual.update_cell(i, 5, 'Sem débitos')
        else:
            botao_impressao = navegador.find_element(
                By.XPATH, '//*[@id="MainContent_lkbImpressao"]')
            botao_impressao.click()

            time.sleep(1)
            pyautogui.press('delete')
            time.sleep(1)
            pyautogui.write(caminho_arquivo)
            time.sleep(1)
            pyautogui.press('enter')

            time.sleep(1)

            estadual.update_cell(i, 5, 'Com débitos')

            if os.path.exists(caminho_arquivo):
                pasta_id = salvar_drive(caminho_arquivo, resp, nome_arquivo)
                responsaveis_com_pasta.add((resp, pasta_id))
                time.sleep(5)
                os.remove(caminho_arquivo)
            else:
                print("Arquivo não encontrado!")
                estadual.update_cell(i, 5, 'Arquivo não encontrado, erro!')


AMANDA = os.getenv('AMANDA')
AMANDA_O = os.getenv('AMANDA_O')
DANIELA_VIVIANE = os.getenv('DANIELA_VIVIANE')
LENI = os.getenv('LENI')
MARCIA = os.getenv('MARCIA')
TATIANE = os.getenv('TATIANE')
DEFAULT = os.getenv('DEFAULT')

for responsavel, id in responsaveis_com_pasta:
    link_pasta = f'https://drive.google.com/drive/folders/{id}'
    mes_atual = datetime.datetime.now().strftime("%m/%Y")

    def obter_email(responsavel):
        match responsavel:
            case "AMANDA":
                return AMANDA
            case "AMANDA O.":
                return AMANDA_O
            case "DANIELA VIVIANE":
                return DANIELA_VIVIANE
            case "LENI":
                return LENI
            case "MARCIA":
                return MARCIA
            case "TATIANE":
                return TATIANE
            case _:
                return DEFAULT
    enviar(obter_email(responsavel),
           f'Levantamento de Debito Estadual - {mes_atual}', responsavel.title(), link_pasta)


navegador.close()
time.sleep(1)
navegador.quit()
