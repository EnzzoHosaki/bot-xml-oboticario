import subprocess
import os
import time
from datetime import datetime, timedelta
import shutil
import zipfile
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import sched

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Criar diretório para relatórios
os.makedirs("relatórios", exist_ok=True)

# Centraliza configuração das variáveis de ambiente
def configurar_variaveis(tipo_arquivo):
    data_atual = datetime.now()
    data_consulta = (data_atual - timedelta(days=2)).strftime('%d%m%Y')
    os.environ['DATA_CONSULTA'] = data_consulta
    os.environ['TIPO_ARQUIVO'] = tipo_arquivo

# Função principal para o script get-xml.py
def get_xml():
    try:
        tipo_documento = os.getenv('TIPO_ARQUIVO')
        data_consulta_str = os.getenv('DATA_CONSULTA')

        # Configura o WebDriver do Chrome
        driver, wait = configurar_driver()

        # Realiza login
        realizar_login(driver, wait)

        # Navega e preenche o formulário de exportação
        acessar_pagina_exportacao(driver, wait)
        preencher_formulario(driver, wait, tipo_documento, data_consulta_str)

        # Realiza o download do XML
        realizar_download_xml(driver, wait)

    except Exception as e:
        logger.error(f"Erro na execução do script get_xml: {e}")
    finally:
        try:
            driver.quit()
            logger.info("Fechando o navegador...")
        except Exception as e:
            logger.error(f"Erro ao fechar o navegador: {e}")

def configurar_driver():
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.set_page_load_timeout(30)
        wait = WebDriverWait(driver, 30)
        return driver, wait
    except Exception as e:
        logger.error(f"Erro ao configurar o WebDriver: {e}")
        raise

def realizar_login(driver, wait):
    try:
        driver.get("https://cp10307.varejofacil.com/app/#/login")
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        driver.maximize_window()

        username_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username input')))
        username_field.click()
        username_field.send_keys('setor fiscal')

        password_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#password input')))
        password_field.click()
        password_field.send_keys('rps@317309')

        botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/section[2]/div/form/button')))
        botao_entrar.click()
    except TimeoutException as e:
        logger.error(f"Erro de timeout ao fazer login: {e}")
        raise
    except Exception as e:
        logger.error(f"Erro ao fazer login: {e}")
        raise

def acessar_pagina_exportacao(driver, wait):
    try:
        botao_menu_fiscal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#sidemenu-item-3')))
        botao_menu_fiscal.click()
        link_exportacao_de_documentos = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Exportação de documentos eletrônicos')]")))
        link_exportacao_de_documentos.click()

        iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="legadoFrame"]')))
        botao_incluir = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoIncluir"]')))
        botao_incluir.click()
    except TimeoutException as e:
        logger.error(f"Erro de timeout ao acessar a página de exportação: {e}")
        raise
    except Exception as e:
        logger.error(f"Erro ao acessar a página de exportação: {e}")
        raise

def preencher_formulario(driver, wait, tipo_documento, data_consulta):
    try:
        iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameExportacaoDocumento"]')))
        
        selecionar_opcao_select(wait, 'documentos', tipo_documento)
        selecionar_opcao_select(wait, 'emissor', 'PROPRIO')
        selecionar_opcao_select(wait, 'tipoDeOperacao', 'SAIDA')
        selecionar_opcao_select(wait, 'tipoArquivo', 'XML')
        selecionar_opcao_select(wait, 'situacaoNota', 'TODAS')

        # Preenche as datas
        preencher_data(wait, data_consulta)

        # Seleciona todas as lojas
        selecionar_lojas(wait)

        # Clica no botão 'Gravar'
        botao_gravar = wait.until(EC.element_to_be_clickable((By.ID, 'btn_salvar')))
        botao_gravar.click()

        driver.switch_to.default_content()
    except Exception as e:
        logger.error(f"Falha ao preencher o formulário: {e}")
        raise

def selecionar_opcao_select(wait, element_id, valor):
    try:
        select_element = wait.until(EC.visibility_of_element_located((By.ID, element_id)))
        select = Select(select_element)
        select.select_by_value(valor)
    except Exception as e:
        logger.error(f"Erro ao selecionar a opção no elemento {element_id}: {e}")
        raise

def preencher_data(wait, data_consulta):
    try:
        data_inicio = wait.until(EC.presence_of_element_located((By.ID, 'dataInicio')))
        data_inicio.click()
        data_inicio.send_keys(data_consulta)
        time.sleep(0.5)

        data_fim = wait.until(EC.presence_of_element_located((By.ID, 'dataFim')))
        data_fim.send_keys(data_consulta)
        time.sleep(0.5)

    except Exception as e:
        logger.error(f"Erro ao preencher as datas: {e}")
        raise

def selecionar_lojas(wait):
    try:
        element_xpath = '//*[@id="formularioExportacaoDocumentoEletronico"]/div/div[3]/div/div/div/input'
        select_element = wait.until(EC.presence_of_element_located((By.XPATH, element_xpath)))
        select_element.click()
        select_element.send_keys(Keys.ENTER)
        time.sleep(0.5)

        for i in range(22):
            select_element.click()
            select_element.send_keys(Keys.ARROW_DOWN)
            select_element.send_keys(Keys.ENTER)
            time.sleep(0.05)
    except Exception as e:
        logger.error(f"Erro ao selecionar as lojas: {e}")
        raise

def realizar_download_xml(driver, wait):
    try:
        # Verifica a conclusão da exportação e baixa o arquivo
        origem = r'C:\Users\enzzo.maciel\Downloads\documentos_eletronicos.zip'
        
        while True:
            if os.path.exists(origem):
                break
            else:
                logger.error(f"O arquivo {os.path.basename(origem)} não foi baixado ainda. Tentando baixar o arquivo")
                download_xml(driver, wait)
    except Exception as e:
        logger.error(f"Erro ao realizar o download do XML: {e}")
        raise

def download_xml(driver, wait):
    try:
        # Espera até que o status da primeira linha seja "Concluído"
        logger.info("Aguardando concluir a exportação")
        time.sleep(10)

        # Dar um refresh na página
        logger.info("Dando refresh na página")
        driver.refresh()

        # Espera até que a página esteja carregada completamente
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # Localize e clique no botão 'Pesquisar' dentro do iframe
        iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="legadoFrame"]')))
        botao_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoPesquisar"]')))
        botao_pesquisar.click()
        time.sleep(1)

        logger.info("Procurando o checkbox")
        tempo_maximo = 10
        inicio_tempo = time.time()

        while True:
            try:
                checkbox = driver.find_element(By.XPATH, '//*[@id="gridExportacaoDocumento"]/div[3]/table/tbody/tr/td[1]')
                logger.info("Checkbox encontrado")
                break
            except NoSuchElementException:
                logger.info("Tentando achar a checkbox")
                time.sleep(1)
                if time.time() - inicio_tempo > tempo_maximo:
                    logger.error("Checkbox não encontrado em 10 segundos. Saindo do loop.")
                    break

        checkbox.click()
        logger.info("Checkbox marcado")

        # Localize e clique no botão 'Donwload' dentro do iframe
        botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="baixarDocumentos"]')))
        botao_download.click()

        # Aguardando um possível download
        time.sleep(10)

    except TimeoutException as e:
        logger.error(f"Erro de timeout ao realizar o download: {e}")
        raise
    except Exception as e:
        logger.error(f"Erro ao realizar o download: {e}")
        raise

# Função principal para o script move-xml-directory.py
def move_xml_directory():
    try:
        data_consulta_str = os.getenv('DATA_CONSULTA')
        tipo_arquivo = os.getenv('TIPO_ARQUIVO')
        data_consulta = datetime.strptime(data_consulta_str, '%d%m%Y')

        destino_dia_mes_ano = preparar_diretorios(data_consulta, tipo_arquivo)

        # Movendo o arquivo para o novo diretório
        origem = r'C:\Users\enzzo.maciel\Downloads\documentos_eletronicos.zip'
        shutil.move(origem, destino_dia_mes_ano)
        arquivo_zip = os.path.join(destino_dia_mes_ano, 'documentos_eletronicos.zip')

        # Descompactando arquivos
        descompactar_arquivo(destino_dia_mes_ano)

        remover_arquivos_zip(arquivo_zip)

        # Descompactando arquivos
        descompactar_arquivo(destino_dia_mes_ano)
        processar_arquivos_internos(destino_dia_mes_ano)

        logger.info(f"Arquivo movido, descompactado e processado com sucesso para {destino_dia_mes_ano}")
    except Exception as e:
        logger.error(f"Erro ao mover e processar o arquivo XML: {e}")
        raise

def preparar_diretorios(data_consulta, tipo_arquivo):
    mes_ano = data_consulta.strftime("%m%Y")
    dia_mes_ano = data_consulta.strftime("%d%m%Y")

    destino_base = f'R:/XML_ASINCRONIZAR/GRUPO OBOTICARIO/{tipo_arquivo}'
    destino_mes_ano = os.path.join(destino_base, mes_ano)
    destino_dia_mes_ano = os.path.join(destino_mes_ano, dia_mes_ano)

    if not os.path.exists(destino_mes_ano):
        os.makedirs(destino_mes_ano)

    if not os.path.exists(destino_dia_mes_ano):
        os.makedirs(destino_dia_mes_ano)

    return destino_dia_mes_ano

def descompactar_arquivo(extract_to):
    try:
        for root, dirs, files in os.walk(extract_to):
            for file in files:
                if file.endswith('.zip'):
                    zip_path = os.path.join(root, file)
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_to)

    except Exception as e:
        logger.error(f"Erro ao descompactar o arquivo {zip_path}: {e}")
        raise

def processar_arquivos_internos(destino_dia_mes_ano):
    try:
        remover_arquivos_zip(destino_dia_mes_ano)

        pasta_docs_fiscais = os.path.join(destino_dia_mes_ano, 'docs-fiscais')
        if os.path.exists(pasta_docs_fiscais):
            for item in os.listdir(pasta_docs_fiscais):
                src_path = os.path.join(pasta_docs_fiscais, item)
                dst_path = os.path.join(destino_dia_mes_ano, item)
                if os.path.exists(dst_path):
                    if os.path.isdir(dst_path):
                        shutil.rmtree(dst_path)
                    else:
                        os.remove(dst_path)
                shutil.move(src_path, dst_path)

            os.rmdir(pasta_docs_fiscais)
    except Exception as e:
        logger.error(f"Erro ao processar arquivos internos: {e}")
        raise

def remover_arquivos_zip(diretorio):
    try:
        for root, dirs, files in os.walk(diretorio):
            for file in files:
                if file.endswith('.zip'):
                    os.remove(os.path.join(root, file))

    except Exception as e:
        logger.error(f"Erro ao remover arquivos zip: {e}")
        raise

def agendar_tarefas(horario):
    try:
        scheduler = sched.scheduler(time.time, time.sleep)
        hora, minuto = map(int, horario.split(':'))
        agora = datetime.now()
        proxima_execucao = agora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
        if proxima_execucao < agora:
            proxima_execucao += timedelta(days=1)

        delay = (proxima_execucao - agora).total_seconds()

        def agendar_execucao():
            for tipo_arquivo in ['NFE', 'NFCE']:
                configurar_variaveis(tipo_arquivo)
                get_xml()
                move_xml_directory()

            feedback_tkinter(f"Execução concluída às {datetime.now().strftime('%H:%M:%S')}")

            scheduler.enter(86400, 1, agendar_execucao)

        scheduler.enter(delay, 1, agendar_execucao)
        logger.info(f"Script agendado para ser executado diariamente às {horario}.")
        scheduler.run()
    except Exception as e:
        logger.error(f"Erro ao agendar tarefas: {e}")
        raise

def feedback_tkinter(mensagem):
    try:
        feedback_root = tk.Tk()
        feedback_root.title("Feedback do Bot XML")
        feedback_root.minsize(300, 100)

        label_feedback = tk.Label(feedback_root, text=mensagem, padx=10, pady=10)
        label_feedback.pack()

        ok_button = tk.Button(feedback_root, text="OK", command=feedback_root.destroy)
        ok_button.pack(pady=10)

        feedback_root.mainloop()
    except Exception as e:
        logger.error(f"Erro ao exibir feedback: {e}")
        raise

def iniciar_agendamento():
    def submit():
        horario = f"{combobox_horas.get()}:{combobox_minutos.get()}"
        root.destroy()
        threading.Thread(target=agendar_tarefas, args=(horario,)).start()
        feedback_tkinter(f"Script agendado para executar diariamente às {horario}.")

    try:
        root = tk.Tk()
        root.title("Bot XML Oboticário")
        root.minsize(200, 100)
        root.iconbitmap('C:/Users/enzzo.maciel/Documents/bot-xml-oboticario/img/bot-icon.png')

        frame = ttk.Frame(root, padding="10")
        frame.grid(row=0, column=0, sticky="nsew")

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)

        tk.Label(frame, text="Hora:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        combobox_horas = ttk.Combobox(frame, values=[f"{i:02d}" for i in range(24)], width=5)
        combobox_horas.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        combobox_horas.set("00")

        tk.Label(frame, text="Minutos:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        combobox_minutos = ttk.Combobox(frame, values=[f"{i:02d}" for i in range(60)], width=5)
        combobox_minutos.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        combobox_minutos.set("00")

        submit_button = tk.Button(frame, text="Agendar", command=submit)
        submit_button.grid(row=1, column=0, columnspan=4, pady=10)

        root.mainloop()
    except Exception as e:
        logger.error(f"Erro ao iniciar agendamento: {e}")
        raise

if __name__ == "__main__":
    iniciar_agendamento()
