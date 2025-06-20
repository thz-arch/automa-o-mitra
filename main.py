from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import logging
import os
import csv
from pathlib import Path
import unicodedata
import sys
import traceback

# Caminho absoluto para o arquivo de log
LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'automacao_mitra.log')
# Configuração do arquivo de log
logging.basicConfig(filename=LOG_PATH,
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Configurações do Selenium (usando Chrome)
chrome_options = Options()
chrome_options.add_argument('--start-maximized')
# chrome_options.add_argument('--headless')  # Descomente para rodar sem abrir janela

service = Service()  # Use o caminho do chromedriver se necessário

# Dicionário fixo de insumos e códigos, extraído do CSV
INSUMO_CODIGO_FIXO = {
    'CORDA 8MM X 240M': '000013',
    'ETIQUETA ADESIVA COUCHÊ 100X50 ( ROLO COM 650)': '000005',
    'FILME STRECH 500 X 25': '000008',
    'FITA ADESIVA 100X45 TRANSPARENTE': '000010',
    'FITA ADESIVA IMPRESSA - ANTI VIOLAÇÃO 48MM X 50M': '000400',
    'LONA PLASTICA 4 X 100 7KG (FINA)': '000009',
    'PALETE PBR 1.00 X 1.20M': '000439',
    'PAPELÃO ONDULADO 1,00 X 30M': '000425',
    'PAPELÃO ONDULADO 1,00 X 50M': '000015',
    'PLASTICO BOLHA 1.30X100 M': '000738',
    'RIBBON CERA 110X74': '000004',
    'RIBBONS 110 X 450 G': '000012',
    'SEST LACRE 3CM X 1,6CM X 1,9CM': '000011',
}

def normalizar_nome(nome):
    import unicodedata
    nome = nome.strip().upper()
    nome = nome.replace('"', '')
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    nome = nome.replace('-', ' ')
    nome = nome.replace('  ', ' ')
    nome = nome.replace('STRECH', 'STRETCH')
    nome = nome.replace('COUCHE', 'COUCHÊ')
    nome = nome.replace('PRB', 'PBR')
    nome = nome.replace('VIOLACAO', 'VIOLACAO').replace('VIOLAÇÃO', 'VIOLACAO')
    nome = nome.replace(' X ', ' X ')
    nome = nome.replace(' x ', ' X ')
    nome = nome.replace('  ', ' ')
    return ' '.join(nome.split())

def executar_baixa():
    # 1. Navegar até a página de login
    driver.get('https://sistema.mitra.inf.br/')
    time.sleep(2)  # Aguarda o carregamento da página

    # 2. Preencher domínio, usuário e senha
    dominio = driver.find_element(By.ID, "dominio")
    usuario = driver.find_element(By.ID, "usuario")
    senha = driver.find_element(By.ID, "password")

    dominio.send_keys('binho')
    usuario.send_keys('thiago')
    senha.send_keys('178523')

    # Clicar no botão "Entrar"
    entrar_btn = driver.find_element(By.ID, "entrar")
    entrar_btn.click()

    time.sleep(5)  # Aguarda o login

    # 3. Abrir o menu "Almoxarifado"
    almoxarifado = driver.find_element(By.XPATH, "//div[contains(@class, 'sidebar-item-modulo') and contains(@class, 'click-on-enter') and contains(., 'Almoxarifado')]")
    almoxarifado.click()
    time.sleep(1)

    # 4. Abrir o submenu "Saída do Almoxarifado"
    saida_almoxarifado = driver.find_element(By.XPATH, "//div[contains(@class, 'sidebar-item-funcao') and contains(@class, 'click-on-enter') and contains(., 'Saída do Almoxarifado')]")
    saida_almoxarifado.click()
    time.sleep(2)

    # Após abrir "Saída do Almoxarifado"
    # Selecionar Centro de Custo (campo com id="1_5")
    centro_custo = driver.find_element(By.ID, "1_5")
    centro_custo.clear()
    centro_custo.send_keys("SPO0002")
    time.sleep(0.5)

    # Selecionar Estoque (clicar no link e escolher SPO002)
    estoque_link = driver.find_element(By.ID, "1_1")
    estoque_link.click()
    try:
        # Espera até a célula com SPO002 estar visível (até 10s)
        estoque_spo002 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), 'SPO002')]"))
        )
        estoque_spo002.click()
        time.sleep(0.5)
    except Exception as e:
        print("Não foi possível selecionar o estoque SPO002:", e)
        logging.error(f"Não foi possível selecionar o estoque SPO002: {e}")

    # Clicar no botão play (input com id="1_11" e class="button")
    botao_play = driver.find_element(By.ID, "1_11")
    botao_play.click()
    time.sleep(2)

    # Se aparecer a janela de mensagem, clicar em "Sim" (input com id="0" e class="inputb")
    try:
        botao_sim = driver.find_element(By.XPATH, "//input[@id='0' and contains(@class, 'inputb') and @value='Sim']")
        botao_sim.click()
        time.sleep(1)
    except Exception:
        pass  # Se não aparecer, segue normalmente

    # --- Leitura da planilha Google ---
    # Configuração do acesso à planilha (você precisa do arquivo de credenciais JSON do Google)
    SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credenciais.json')
    SPREADSHEET_KEY = '1-1S13HAvTMQkY7aVfsslinV7X5w_Awdp0a2UvepjT1g'
    aba = 'PAINEL'  # Ajustado conforme o nome da aba

    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_KEY).worksheet(aba)

    # Define o intervalo de datas
    inicio_periodo = datetime(2025, 6, 17, 18, 0, 0)  # 17/06/2025 18:00
    fim_periodo = datetime.now()
    datas = sheet.col_values(10)  # Coluna J (DATA ENVIO PARA EXPEDIÇÃO)
    insumos = sheet.col_values(5)  # Coluna E (INSUMO SOLICITADO)
    quantidades = sheet.col_values(11)  # Coluna K (QUANTIDADE ENVIADA)
    colunas_totais = sheet.row_values(1)
    idx_baixa = len(colunas_totais)  # última coluna
    sinais_baixa = sheet.col_values(idx_baixa)
    registros_para_baixa = []
    # Itera todos os registros válidos e armazena índice da linha
    for i in range(1, min(len(datas), len(insumos), len(quantidades), len(sinais_baixa))):
        try:
            data_linha = datetime.strptime(datas[i][:19], '%d/%m/%Y %H:%M:%S')
        except Exception:
            continue
        if (inicio_periodo <= data_linha <= fim_periodo and
            insumos[i].strip() and quantidades[i].strip() and
            sinais_baixa[i].strip().upper() == 'NÃO'):
            registros_para_baixa.append({
                'insumo': insumos[i].strip(),
                'quantidade': quantidades[i],
                'data': datas[i],
                'linha': i+1  # linha real na planilha (1-based)
            })

    # Agrupa por insumo e soma as quantidades, mantendo as linhas
    from collections import defaultdict
    insumos_agrupados = defaultdict(lambda: {'quantidade': 0, 'linhas': [], 'data': None})
    for reg in registros_para_baixa:
        nome = reg['insumo']
        try:
            qtd = float(str(reg['quantidade']).replace(',', '.'))
        except Exception:
            qtd = 0
        insumos_agrupados[nome]['quantidade'] += qtd
        insumos_agrupados[nome]['linhas'].append(reg['linha'])
        insumos_agrupados[nome]['data'] = reg['data']  # última data encontrada

    # Monta o dicionário insumo -> código a partir do CSV
    insumo_codigo_csv = {normalizar_nome(k): v for k, v in INSUMO_CODIGO_FIXO.items()}

    for insumo, info in insumos_agrupados.items():
        quantidade = info['quantidade']
        data_str = info['data']
        linhas_planilha = info['linhas']
        try:
            print(f"Dando baixa em: {insumo}, quantidade: {quantidade}, data: {data_str}")
            logging.info(f"Dando baixa em: {insumo}, quantidade: {quantidade}, data: {data_str}")
            for handler in logging.getLogger().handlers:
                handler.flush()
            codigo_item = insumo_codigo_csv.get(normalizar_nome(insumo))
            print('Insumo da planilha (normalizado):', normalizar_nome(insumo))
            print('Insumos do CSV (normalizados):', list(insumo_codigo_csv.keys()))
            if not codigo_item:
                print(f"Código não encontrado para o insumo '{insumo}'. Verifique o nome na planilha CSV.")
                logging.error(f"Código não encontrado para o insumo '{insumo}'. Verifique o nome na planilha CSV.")
                for handler in logging.getLogger().handlers:
                    handler.flush()
                raise Exception(f"Código não encontrado para o insumo '{insumo}'. Verifique o nome na planilha CSV.")
            campo_codigo = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "2_4"))
            )
            campo_codigo.click()
            time.sleep(0.2)
            campo_codigo.clear()
            time.sleep(0.2)
            campo_codigo.send_keys(codigo_item)
            logging.info(f"Digitado código '{codigo_item}' para o insumo '{insumo}'.")
            for handler in logging.getLogger().handlers:
                handler.flush()
            time.sleep(2)
            time.sleep(1)
            campo_quantidade = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "2_8"))
            )
            campo_quantidade.clear()
            # Ajuste para FILME STRECH 500 X 25: quantidade deve ser vezes 4 (kg)
            if normalizar_nome(insumo) == normalizar_nome('FILME STRECH 500 X 25'):
                try:
                    quantidade = quantidade * 4
                except Exception:
                    pass
            campo_quantidade.send_keys(str(quantidade).replace('.', ','))
            logging.info(f"Quantidade '{quantidade}' inserida para o item '{insumo}'.")
            for handler in logging.getLogger().handlers:
                handler.flush()
            time.sleep(0.5)
            try:
                botao_incluir = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "2_11"))
                )
                botao_incluir.click()
                logging.info(f"Clicado no botão Incluir (id=2_11) para o item '{insumo}'.")
                for handler in logging.getLogger().handlers:
                    handler.flush()
                time.sleep(0.5)
                # Marca todas as linhas do insumo como FINALIZADO
                for linha_planilha in linhas_planilha:
                    sheet.update_cell(linha_planilha, idx_baixa, "FINALIZADO")
                    logging.info(f"Marcado como FINALIZADO na planilha (linha {linha_planilha}, coluna {idx_baixa}) para o insumo '{insumo}'.")
                    for handler in logging.getLogger().handlers:
                        handler.flush()
            except Exception as e:
                print(f"Não foi possível clicar no botão Incluir para o insumo '{insumo}':", e)
                logging.error(f"Não foi possível clicar no botão Incluir para o insumo '{insumo}': {e}")
                for handler in logging.getLogger().handlers:
                    handler.flush()
                raise Exception(f"Falha ao incluir o insumo '{insumo}'. Corrija manualmente e reinicie o script.")
        except Exception as e:
            print(f"Erro ao processar o insumo '{insumo}':", e)
            logging.error(f"Erro ao processar o insumo '{insumo}': {e}")
            for handler in logging.getLogger().handlers:
                handler.flush()
            continue
    # Após processar todos os insumos, clicar em 'Finalizar'
    try:
        print('Tentando clicar no botão Finalizar...')
        time.sleep(1)  # Aguarda atualização da página
        try:
            botao_finalizar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "2_13"))
            )
        except Exception:
            botao_finalizar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[contains(@value, 'Finalizar')]"))
            )
        botao_finalizar.click()
        print('Botão Finalizar clicado.')
        logging.info("Clicado no botão Finalizar.")
        for handler in logging.getLogger().handlers:
            handler.flush()
        # --- NOVO: Preencher data de movimentação e clicar em Continuar ---
        print('Aguardando modal de data de movimentação...')
        campo_data = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@type='text' and @maxlength='6']"))
        )
        if len(insumos_agrupados) > 0:
            # Usa a data do último insumo processado
            data_ultimo = list(insumos_agrupados.values())[-1]['data']
            try:
                data_mov = datetime.strptime(data_ultimo[:10], '%d/%m/%Y').strftime('%y%m%d')
            except Exception:
                data_mov = datetime.now().strftime('%y%m%d')
        else:
            data_mov = datetime.now().strftime('%y%m%d')
        campo_data.clear()
        campo_data.send_keys(data_mov)
        print(f"Data de movimentação '{data_mov}' informada.")
        time.sleep(0.5)
        print('Procurando botão Continuar...')
        botao_continuar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Continuar']"))
        )
        botao_continuar.click()
        print('Botão Continuar clicado.')
        logging.info(f"Data de movimentação '{data_mov}' informada e clicado em Continuar.")
        for handler in logging.getLogger().handlers:
            handler.flush()
    except Exception as e:
        print("Erro ao clicar no botão Finalizar ou preencher a data:", e)
        # Salva HTML da tela para depuração
        with open('erro_finalizar.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        logging.error(f"Não foi possível clicar no botão Finalizar ou preencher a data: {e}")
        for handler in logging.getLogger().handlers:
            handler.flush()
    logging.info('Processo concluído!')
    for handler in logging.getLogger().handlers:
        handler.flush()
    logging.shutdown()

if __name__ == "__main__":
    while True:
        driver = None
        try:
            driver = webdriver.Chrome(service=service, options=chrome_options)
            executar_baixa()
        except Exception as e:
            print('Erro na execução principal:', e)
            traceback.print_exc()
            logging.error(f'Erro na execução principal: {e}')
            for handler in logging.getLogger().handlers:
                handler.flush()
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
        print('Execução finalizada. Aguardando 24h para próxima execução...')
        time.sleep(60*60*24)  # Aguarda 24 horas para rodar novamente
