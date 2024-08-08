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
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import sched

# Definir a data atual
data_atual = datetime.now()

# Calcular a data de consulta
data_consulta = (data_atual - timedelta(days=2)).strftime('%d%m%Y')

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Criar diretório para relatórios
os.makedirs("relatórios", exist_ok=True)

# Função para executar um script com as variáveis de ambiente
def executar_script(script, tipo_arquivo):
    env = os.environ.copy()
    env['DATA_CONSULTA'] = data_consulta
    env['TIPO_ARQUIVO'] = tipo_arquivo
    result = subprocess.run(['python', script], env=env)
    if result.returncode != 0:
        logger.error(f"Erro ao executar o script {script} com tipo_arquivo={tipo_arquivo}")
    else:
        logger.info(f"Script {script} executado com sucesso com tipo_arquivo={tipo_arquivo}")

# Função para o script get-xml.py
def get_xml():
    tipo_documento = os.getenv('TIPO_ARQUIVO')
    data_consulta_str = os.getenv('DATA_CONSULTA')

    def selecionar_opcao_select(wait, element_id, valor):
        try:
            logger.info(f"Localizando o elemento <select> com ID: {element_id}")
            select_element = wait.until(EC.visibility_of_element_located((By.ID, element_id)))
            logger.info(f"Criando objeto Select para o elemento {element_id}")
            select = Select(select_element)
            logger.info(f"Selecionando a opção com valor: {valor}")
            select.select_by_value(valor)
            logger.info(f"Opção com valor {valor} selecionada para o elemento {element_id}")
        except Exception as e:
            logger.error(f"Erro ao selecionar a opção no elemento {element_id}: {e}")

    def preencher_formulario(driver, wait, tipo_documento, data_consulta):
        try:
            # Localize o iframe e mude o contexto para ele
            iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="frameExportacaoDocumento"]')))

            # Preenche os campos do formulário
            selecionar_opcao_select(wait, 'documentos', tipo_documento)
            selecionar_opcao_select(wait, 'emissor', 'PROPRIO')
            selecionar_opcao_select(wait, 'tipoDeOperacao', 'SAIDA')
            selecionar_opcao_select(wait, 'tipoArquivo', 'XML')
            selecionar_opcao_select(wait, 'situacaoNota', 'TODAS')

            # DATA INICIAL
            logger.info("Selecionando a data inicial")
            data_inicial = wait.until(EC.presence_of_element_located((By.ID, 'dataInicio')))
            data_inicial.click()
            data_inicial.send_keys(data_consulta)
            time.sleep(0.5)

            # DATA FINAL
            logger.info("Selecionando a data final")
            data_final = wait.until(EC.presence_of_element_located((By.ID, 'dataFim')))
            data_final.send_keys(data_consulta)
            time.sleep(0.5)

            # LOJAS
            element_xpath = '//*[@id="formularioExportacaoDocumentoEletronico"]/div/div[3]/div/div/div/input'
            logger.info(f"Localizando o elemento <select> com XPath: {element_xpath}")
            select_element = wait.until(EC.presence_of_element_located((By.XPATH, element_xpath)))
            select_element.click()
            time.sleep(0.5)

            # Enviar teclas para selecionar a opção "Selecionar Tudo"
            select_element.send_keys(Keys.ENTER)  # Selecione a opção
            for i in range(22):
                select_element.click()
                select_element.send_keys(Keys.ARROW_DOWN)  # Navegue para a opção "Selecionar Tudo"
                select_element.send_keys(Keys.ENTER)  # Selecione a opção
                time.sleep(0.05)

            # Clica no botão 'Gravar'
            logger.info("Clicando no botão de salvar do formulário")
            botao_gravar = wait.until(EC.element_to_be_clickable((By.ID, 'btn_salvar')))
            botao_gravar.click()

            # Espera até que a página esteja carregada completamente
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

            logger.info("Voltando para o contexto principal")
            driver.switch_to.default_content()  # Volta para o contexto principal

        except Exception as e:
            logger.error(f"Falha ao preencher o formulário: {e}")

    try:
        # Configura o WebDriver do Chrome e garante que a versão mais recente seja usada
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.set_page_load_timeout(30)
        wait = WebDriverWait(driver, 30)

        # Abre o site
        driver.get("link do ambiente VarejoFacil")

        # Espera até que a página esteja carregada completamente
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        driver.maximize_window()

        # Realiza login
        try:
            username_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#username input')))
            username_field.click()
            username_field.send_keys('login')

            password_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#password input')))
            password_field.click()
            password_field.send_keys('senha')

            botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/section[2]/div/form/button')))
            botao_entrar.click()
        except Exception as e:
            logger.error(f"Erro ao fazer login: {e}")

        # Abre a página "Exportação de documentos eletrônicos"
        try:
            botao_menu_fiscal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#sidemenu-item-3')))
            botao_menu_fiscal.click()
            link_exportacao_de_documentos = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Exportação de documentos eletrônicos')]")))
            link_exportacao_de_documentos.click()
        except Exception as e:
            logger.error(f"Erro ao tentar abrir a página 'Exportação de Documentos Eletrônicos': {e}")

        # Localize e clique no botão 'Incluir' dentro do iframe
        try:
            iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="legadoFrame"]')))
            botao_incluir = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoIncluir"]')))
            botao_incluir.click()
        except Exception as e:
            logger.error(f"Erro ao clicar no botão 'Incluir': {e}")

        # Preenche o formulário de exportação
        preencher_formulario(driver, wait, tipo_documento, data_consulta_str)

        # Volta para o iframe onde a tabela está localizada
        logger.info("Preparando para trocar o iframe")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="legadoFrame"]')))
        logger.info("Iframe trocado")

        # Aguarda até que a tabela esteja presente
        logger.info("Aguardando até que a tabela esteja presente")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[2]/div[3]/table")))
        
        # Aguarda até que o status da primeira linha seja "Concluído"
        logger.info("Aguardando concluir a exportação")
        time.sleep(150)
        logger.info("Continua esperando...")

        # Dar um refresh na página
        logger.info("Dando refresh na página")
        driver.refresh()

        # Espera até que a página esteja carregada completamente
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # Localize e clique no botão 'Pesquisar' dentro do iframe
        try:
            iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="legadoFrame"]')))
            botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoPesquisar"]')))
            botao_download.click()
        except Exception as e:
            logger.error(f"Erro ao clicar no botão 'Pesquisar': {e}")

        time.sleep(1)

        logger.info("Procurando o checkbox")
        # Define o tempo máximo de espera (10 segundos)
        tempo_maximo = 10
        inicio_tempo = time.time()

        while True:
            try:
                checkbox = driver.find_element(By.XPATH, '//*[@id="gridExportacaoDocumento"]/div[3]/table/tbody/tr/td[1]')
                logger.info("Checkbox encontrado")
                break  # Sai do loop se o checkbox for encontrado
            except NoSuchElementException:
                logger.info("Tentando achar a checkbox")
                time.sleep(1)  # Espera 1 segundo antes de tentar novamente
                
                # Verifica se o tempo máximo foi atingido
                if time.time() - inicio_tempo > tempo_maximo:
                    logger.error("Checkbox não encontrado em 10 segundos. Saindo do loop.")
                    break

        # Marcar o checkbox na primeira linha e primeira coluna da tabela
        checkbox.click()
        logger.info("Checkbox marcado")

        # Localize e clique no botão 'Donwload' dentro do iframe
        try:
            botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="baixarDocumentos"]')))
            botao_download.click()
        except Exception as e:
            logger.error(f"Erro ao clicar no botão 'Incluir': {e}")

        time.sleep(10)

    finally:
        # Garante que o driver seja fechado
        try:
            driver.quit()
            logger.info("Fechando o navegador...")
        except Exception as e:
            logger.error(f"Erro ao fechar o navegador: {e}")

# Função para o script move-xml-directory.py
def move_xml_directory():
    data_consulta_str = os.getenv('DATA_CONSULTA')
    tipo_arquivo = os.getenv('TIPO_ARQUIVO')
    data_consulta = datetime.strptime(data_consulta_str, '%d%m%Y')

    # Formatando as datas para criação das pastas
    mes_ano = data_consulta.strftime("%m%Y")
    dia_mes_ano = data_consulta.strftime("%d%m%Y")

    # Diretório de origem e destino
    origem = r'C:\Users\enzzo.maciel\Downloads\documentos_eletronicos.zip'
    destino_base = f'R:/XML_ASINCRONIZAR/GRUPO OBOTICARIO/{tipo_arquivo}'

    # Criando o caminho do novo diretório
    destino_mes_ano = os.path.join(destino_base, mes_ano)
    destino_dia_mes_ano = os.path.join(destino_mes_ano, dia_mes_ano)

    # Função para descompactar arquivos zip
    def descompactar_arquivo(zip_path, extract_to):
        if os.path.exists(zip_path):
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to)

    # Função para remover arquivos zip
    def remover_arquivos_zip(diretorio):
        for root, dirs, files in os.walk(diretorio):
            for file in files:
                if file.endswith('.zip'):
                    os.remove(os.path.join(root, file))

    # Looping infinito para verificar a existência do arquivo
    start_time = time.time()
    while True:
        if os.path.exists(origem):
            break
        elif time.time() - start_time > 120:
            logger.error(f"O arquivo {os.path.basename(origem)} não foi encontrado.")
            exit()

    # Verificando e criando o diretório do mês e ano se não existir
    if not os.path.exists(destino_mes_ano):
        os.makedirs(destino_mes_ano)

    # Verificando e criando o diretório do dia, mês e ano se não existir
    if not os.path.exists(destino_dia_mes_ano):
        os.makedirs(destino_dia_mes_ano)

    # Movendo o arquivo para o novo diretório
    shutil.move(origem, destino_dia_mes_ano)
    arquivo_zip = os.path.join(destino_dia_mes_ano, 'documentos_eletronicos.zip')

    # Descompactando o primeiro arquivo zip
    descompactar_arquivo(arquivo_zip, destino_dia_mes_ano)

    # Localizando e descompactando o segundo arquivo zip
    for root, dirs, files in os.walk(destino_dia_mes_ano):
        for file in files:
            if file.endswith('.zip'):
                arquivo_zip_interno = os.path.join(root, file)
                if os.path.exists(arquivo_zip_interno):
                    # Verifica se a pasta já existe, se sim, remove a antiga
                    pasta_interna = arquivo_zip_interno.replace('.zip', '')
                    if os.path.exists(pasta_interna):
                        shutil.rmtree(pasta_interna)
                        logger.info(f"Diretório duplicado {pasta_interna} foi removido.")
                    descompactar_arquivo(arquivo_zip_interno, root)
                    os.remove(arquivo_zip_interno)

    # Removendo todos os arquivos .zip dos diretórios especificados
    remover_arquivos_zip(destino_mes_ano)
    remover_arquivos_zip(destino_dia_mes_ano)

    # Movendo a pasta docs-fiscais para destino_dia_mes_ano
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

        # Removendo a pasta docs-fiscais vazia
        os.rmdir(pasta_docs_fiscais)

    logger.info(f"Arquivo movido, descompactado e processado com sucesso para {destino_dia_mes_ano}")

def agendar_tarefas(horario):
    scheduler = sched.scheduler(time.time, time.sleep)
    hora, minuto = map(int, horario.split(':'))
    agora = datetime.now()
    proxima_execucao = agora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
    if proxima_execucao < agora:
        proxima_execucao += timedelta(days=1)

    delay = (proxima_execucao - agora).total_seconds()
    
    def agendar_execucao():
        for tipo_arquivo in ['NFE', 'NFCE']:
            os.environ['DATA_CONSULTA'] = data_consulta
            os.environ['TIPO_ARQUIVO'] = tipo_arquivo
            get_xml()
            move_xml_directory()
        logger.info("Execução concluída, re-agendando para o próximo dia.")
        scheduler.enter(86400, 1, agendar_execucao)  # Re-agenda para o próximo dia

    scheduler.enter(delay, 1, agendar_execucao)
    logger.info(f"Script agendado para ser executado diariamente às {horario}.")
    scheduler.run()

def iniciar_agendamento():
    def submit():
        horario = f"{combobox_horas.get()}:{combobox_minutos.get()}"
        root.destroy()  # Fecha a janela do Tkinter
        threading.Thread(target=agendar_tarefas, args=(horario,)).start()
        messagebox.showinfo("Agendamento", f"Script agendado para executar diariamente às {horario}.")

    root = tk.Tk()
    root.title("Bot XML Oboticário")
    root.minsize(200, 100)
    root.iconbitmap('C:/Users/enzzo.maciel/Documents/bot-xml-oboticario-automatico/img/bot-icon.png')

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

if __name__ == "__main__":
    iniciar_agendamento()
