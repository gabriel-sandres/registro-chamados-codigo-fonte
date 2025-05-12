#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
from datetime import datetime
import pandas as pd
import os
import time
import logging
import traceback
from typing import Optional, Tuple
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# === CONFIGURAÇÃO DE LOGGING ===
def setup_logging():
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"registro_chamados_{timestamp}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

# === CONFIGURAÇÕES GERAIS ===
BASE_URL = "https://portal.sisbr.coop.br/visao360/consult"
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(os.path.dirname(__file__), "planilha_registro.xlsm"))
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", os.path.join(os.path.dirname(__file__), "chromedriver.exe"))
dotenv_path = os.path.join(os.path.dirname(__file__), "login.env")

SERVICOS_VALIDOS = {
    "dúvida negocial": "Dúvida Negocial",
    "duvida negocial": "Dúvida Negocial",
    "duvida negociacao": "Dúvida Negocial",
    "dúvida negociacao": "Dúvida Negocial",
    "duvida de negocio": "Dúvida Negocial",
    "duvida negocio": "Dúvida Negocial",
    "dúvida técnica": "Dúvida Técnica",
    "duvida tecnica": "Dúvida Técnica",
    "duvida de tecnica": "Dúvida Técnica",
    "ambiente de testes": "Ambiente de testes",
    "ambiente testes": "Ambiente de testes",
    "ambiente de teste": "Ambiente de testes",
    "ambiente teste": "Ambiente de testes",
    "erro de documentação": "Erro De Documentação",
    "erro de documentacao": "Erro De Documentação",
    "erro documentacao": "Erro De Documentação",
    "erro documentação": "Erro De Documentação",
    "integração imcompleta": "Integração Imcompleta",
    "integracao imcompleta": "Integração Imcompleta",
    "integracao incompleta": "Integração Imcompleta",
    "integração incompleta": "Integração Imcompleta",
    "sugestão de melhoria": "Sugestão De Melhoria",
    "sugestao de melhoria": "Sugestão De Melhoria",
    "sugestao melhoria": "Sugestão De Melhoria",
    "sugestão melhoria": "Sugestão De Melhoria",
}

def normalizar_servico(servico):
    if not isinstance(servico, str):
        return servico
    chave = (servico.strip().lower()
        .replace("á", "a").replace("à", "a").replace("ã", "a").replace("â", "a")
        .replace("é", "e").replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o").replace("ô", "o").replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c"))
    return SERVICOS_VALIDOS.get(chave, servico)

class RegistroChamadoError(Exception):
    pass

class LoginError(RegistroChamadoError):
    pass

class FormularioError(RegistroChamadoError):
    pass

class FinalizacaoError(RegistroChamadoError):
    pass

def log_error(error: Exception, context: str, index: Optional[int] = None, df: Optional[pd.DataFrame] = None) -> None:
    error_msg = f"[{'Linha ' + str(index) if index is not None else 'Geral'}] ❌ ERRO em {context}: {str(error)}"
    logger.error(error_msg)
    logger.error("Stack trace:", exc_info=True)
    
    if index is not None and df is not None:
        df.at[index, 'Observação'] = f"Erro em {context}: {str(error)}"
        df.to_excel(EXCEL_PATH, index=False)

def setup_driver(download_dir: str) -> webdriver.Chrome:
    options = Options()
    options.add_experimental_option('prefs', {'download.default_directory': download_dir})
    service = ChromeService(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver

def load_credentials():
    if not os.path.exists(dotenv_path):
        raise FileNotFoundError(f"Arquivo {dotenv_path} não encontrado")
    load_dotenv(dotenv_path)
    username = os.getenv("LOGIN_USERNAME")
    password = os.getenv("LOGIN_PASSWORD")
    if not username or not password:
        raise ValueError("Credenciais não encontradas no arquivo .env")
    return username, password

def load_excel_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(
        file_path,
        dtype={'Documento do cooperado': str}
    )
    return df

def login(driver: webdriver.Chrome, username: str, password: str):
    try:
        logger.info("🔄 Iniciando processo de login...")
        driver.get(BASE_URL)
        
        logger.info("Preenchendo credenciais...")
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'username'))).send_keys(username)
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'password'))).send_keys(password)
        
        logger.info("Clicando no botão de login...")
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'kc-login'))).click()
        
        logger.info("Aguardando QR code desaparecer...")
        WebDriverWait(driver, 300).until(EC.invisibility_of_element_located((By.ID, "qr-code")))
        logger.info("✅ Login realizado com sucesso!")
        
    except TimeoutException as e:
        log_error(e, "processo de login")
        raise LoginError(f"Timeout durante o login: {str(e)}")
    except NoSuchElementException as e:
        log_error(e, "processo de login")
        raise LoginError(f"Elemento não encontrado durante o login: {str(e)}")
    except Exception as e:
        log_error(e, "processo de login")
        raise LoginError(f"Falha no login: {str(e)}")

def limpar_e_preencher(campo, valor):
    campo.click()
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    campo.send_keys(valor)

# Modificado para lidar com cliques interceptados e garantir interação
def preencher_com_sugestao(campo, valor, driver):
    try:
        # Garantir que o campo esteja clicável
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, campo.get_attribute("id"))))
        # Usar ActionChains para clicar
        actions = ActionChains(driver)
        actions.move_to_element(campo).click().perform()
        
        campo.clear()  # Limpar qualquer valor pré-existente
        # Digita os primeiros caracteres para acionar a lista
        campo.send_keys(valor[:3])
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//option[contains(text(), '{valor}')] | //li[contains(text(), '{valor}')]"))
        )
        # Simula navegação pela lista
        campo.send_keys(Keys.ARROW_DOWN)
        campo.send_keys(Keys.ENTER)
    except TimeoutException as e:
        raise FormularioError(f"Timeout ao localizar sugestão para '{valor}': {str(e)}")
    except NoSuchElementException as e:
        raise FormularioError(f"Sugestão para '{valor}' não encontrada: {str(e)}")
    except ElementClickInterceptedException as e:
        logger.warning(f"Clique interceptado ao preencher '{valor}': {str(e)}")
        # Tenta clicar via JavaScript
        driver.execute_script("arguments[0].click();", campo)
        campo.clear()
        campo.send_keys(valor[:3])
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//option[contains(text(), '{valor}')] | //li[contains(text(), '{valor}')]"))
        )
        campo.send_keys(Keys.ARROW_DOWN)
        campo.send_keys(Keys.ENTER)
    except Exception as e:
        raise FormularioError(f"Erro ao preencher sugestão para '{valor}': {str(e)}")

def preencher_com_datalist(campo, valor):
    campo.click()
    campo.clear()
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    campo.click()
    for char in valor:
        campo.send_keys(char)
    campo.send_keys(Keys.TAB)

def preencher_campo_com_js(driver, campo_xpath, valor):
    try:
        print(f"Preenchendo campo com valor: {valor}")
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo_xpath))
        )
        preencher_com_sugestao(campo, valor, driver)
        print(f"Campo preenchido com: {valor}")
        
    except TimeoutException as e:
        print(f"Timeout ao localizar campo: {e}")
        raise FormularioError(f"Timeout ao localizar campo: {str(e)}")
    except NoSuchElementException as e:
        print(f"Campo não encontrado: {e}")
        raise FormularioError(f"Campo não encontrado: {str(e)}")
    except Exception as e:
        print(f"Erro ao preencher campo: {e}")
        raise FormularioError(f"Erro ao preencher campo: {str(e)}")

def selecionar_opcao(driver, campo_xpath, opcao_xpath):
    try:
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo_xpath))
        )
        campo.click()
        
        opcao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, opcao_xpath))
        )
        valor = opcao.get_attribute("value")
        
        primeiros_chars = valor[:3]
        campo.clear()
        campo.send_keys(primeiros_chars)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//option[@value='{valor}']")))
        
        for _ in range(10):
            campo.send_keys(Keys.ARROW_DOWN)
            texto_atual = campo.get_attribute("value")
            if texto_atual and texto_atual.lower() == valor.lower():
                campo.send_keys(Keys.ENTER)
                return
        
        driver.execute_script("arguments[0].click();", opcao)
        
    except TimeoutException as e:
        print(f"Timeout ao selecionar opção: {e}")
        raise FormularioError(f"Timeout ao selecionar opção: {str(e)}")
    except NoSuchElementException as e:
        print(f"Opção não encontrada: {e}")
        raise FormularioError(f"Opção não encontrada: {str(e)}")
    except Exception as e:
        print(f"Erro ao selecionar opção: {e}")
        try:
            campo.clear()
            campo.send_keys(valor)
            campo.send_keys(Keys.TAB)
        except:
            raise FormularioError(f"Erro ao selecionar opção: {str(e)}")

def selecionar_opcao_select(driver, select_xpath, valor):
    try:
        print(f"Selecionando opção '{valor}' no select...")
        select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, select_xpath))
        )
        
        select = Select(select_element)
        try:
            select.select_by_visible_text(valor.title())
        except NoSuchElementException:
            select.select_by_value(valor.lower())
        
        print(f"Opção '{valor}' selecionada no select")
        
    except TimeoutException as e:
        print(f"Timeout ao localizar select: {e}")
        raise FormularioError(f"Timeout ao localizar select: {str(e)}")
    except NoSuchElementException as e:
        print(f"Select ou opção não encontrada: {e}")
        try:
            options = select_element.find_elements(By.TAG_NAME, "option")
            for option in options:
                if option.text.lower() == valor.lower() or option.get_attribute("value").lower() == valor.lower():
                    driver.execute_script("arguments[0].selected = true;", option)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", select_element)
                    print(f"Opção '{valor}' selecionada via JavaScript")
                    return
            raise FormularioError(f"Opção '{valor}' não encontrada no select: {str(e)}")
        except Exception as e2:
            print(f"Erro na abordagem alternativa: {e2}")
            raise FormularioError(f"Erro ao selecionar opção no select: {str(e2)}")
    except Exception as e:
        print(f"Erro ao selecionar opção no select: {e}")
        raise FormularioError(f"Erro ao selecionar opção no select: {str(e)}")

def selecionar_conta_por_cooperativa(driver, cooperativa, index):
    try:
        print(f"[Linha {index}] Selecionando conta para cooperativa {cooperativa}...")
        select_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
        
        # Espera o spinner desaparecer antes de tentar selecionar a conta
        esperar_spinner_desaparecer(driver, index)
        
        # Espera o select estar presente e clicável
        select_element = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, select_xpath))
        )
        
        # Rola até o elemento para garantir que está visível
        driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
        time.sleep(1)  # Pequena pausa para garantir que a rolagem terminou
        
        # Tenta clicar no select primeiro
        try:
            select_element.click()
        except:
            actions = ActionChains(driver)
            actions.move_to_element(select_element).click().perform()
        
        # Espera as opções aparecerem
        time.sleep(1)
        
        options = select_element.find_elements(By.TAG_NAME, 'option')
        
        conta_encontrada = False
        for option in options:
            texto_opcao = option.text.strip()
            # Verifica se o texto da opção contém a cooperativa
            if f"Coop: {cooperativa}" in texto_opcao:
                print(f"[Linha {index}] Conta encontrada: {texto_opcao}")
                try:
                    option.click()
                except:
                    try:
                        driver.execute_script("arguments[0].click();", option)
                    except:
                        actions = ActionChains(driver)
                        actions.move_to_element(option).click().perform()
                conta_encontrada = True
                break
        
        if not conta_encontrada:
            print(f"[Linha {index}] ⚠️ ATENÇÃO: Nenhuma conta encontrada para cooperativa {cooperativa}")
            return False
            
        # Verifica se a conta foi realmente selecionada
        time.sleep(1)  # Pequena pausa para garantir que a seleção foi processada
        try:
            selected_option = select_element.find_element(By.XPATH, "./option[@selected]")
            texto_selecionado = selected_option.text.strip()
            print(f"[Linha {index}] Conta selecionada: {texto_selecionado}")
            
            # Verifica se o texto selecionado contém a cooperativa
            if f"Coop: {cooperativa}" not in texto_selecionado:
                print(f"[Linha {index}] ⚠️ Conta selecionada não corresponde à cooperativa {cooperativa}")
                return False
        except NoSuchElementException:
            print(f"[Linha {index}] ⚠️ Não foi possível verificar a conta selecionada")
            return False
            
        # Espera o spinner desaparecer após a seleção
        esperar_spinner_desaparecer(driver, index)
        
        # Aguarda a tela mudar para a tela de formulário
        try:
            print(f"[Linha {index}] Aguardando mudança para tela de formulário...")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//form"))
            )
            print(f"[Linha {index}] ✅ Tela de formulário carregada")
            return True
        except TimeoutException:
            print(f"[Linha {index}] ⚠️ Timeout ao aguardar tela de formulário")
            return False
            
    except TimeoutException as e:
        print(f"[Linha {index}] Timeout ao selecionar conta: {e}")
        return False
    except NoSuchElementException as e:
        print(f"[Linha {index}] Select de conta não encontrado: {e}")
        return False
    except Exception as e:
        print(f"[Linha {index}] Erro ao selecionar conta: {e}")
        return False

def verificar_pessoa_nao_encontrada(driver, index):
    try:
        erro_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[1]/h6'
        
        elementos_erro = driver.find_elements(By.XPATH, erro_xpath)
        if elementos_erro:
            mensagem_erro = elementos_erro[0].text.strip()
            if "Pessoa não identificada como cooperada!" in mensagem_erro:
                print(f"[Linha {index}] ⚠️ ERRO: {mensagem_erro}")
                return True
        return False
        
    except Exception as e:
        print(f"[Linha {index}] Erro ao verificar pessoa não encontrada: {e}")
        return False

def formatar_documento(documento):
    numeros = ''.join(filter(str.isdigit, str(documento)))
    if len(numeros) == 11:
        numeros = numeros.zfill(11)
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    elif len(numeros) == 14:
        numeros = numeros.zfill(14)
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    else:
        logger.warning(f"Documento inválido: {documento}")
        return documento

def esperar_modal_desaparecer(driver, index, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "modal"))
        )
        return True
    except TimeoutException:
        logger.warning(f"[Linha {index}] Modal ainda presente após {timeout} segundos")
        return False

def esperar_spinner_desaparecer(driver, index, timeout=30):
    try:
        spinner_xpath = "//div[contains(@class, 'ngx-spinner-overlay')]"
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.XPATH, spinner_xpath))
        )
        return True
    except TimeoutException:
        print(f"[Linha {index}] Timeout ao esperar spinner desaparecer")
        return False
    except Exception as e:
        print(f"[Linha {index}] Erro ao esperar spinner desaparecer: {e}")
        return False

def clicar_botao_consulta(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão consultar...")
        botao_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'
        
        botao = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, botao_xpath))
        )
        
        tentativas = 0
        max_tentativas = 3
        while tentativas < max_tentativas:
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                actions = ActionChains(driver)
                actions.move_to_element(botao).click().perform()
                print(f"[Linha {index}] Botão consultar clicado com sucesso")
                return True
            except ElementClickInterceptedException:
                try:
                    driver.execute_script("arguments[0].click();", botao)
                    print(f"[Linha {index}] Botão consultar clicado via JavaScript")
                    return True
                except:
                    tentativas += 1
                    if tentativas < max_tentativas:
                        print(f"[Linha {index}] Tentativa {tentativas} falhou, tentando novamente...")
                        time.sleep(1)
                    else:
                        print(f"[Linha {index}] ❌ Não foi possível clicar no botão após {max_tentativas} tentativas")
                        return False
        return False
    except TimeoutException as e:
        print(f"[Linha {index}] Timeout ao localizar botão consultar: {str(e)}")
        return False
    except NoSuchElementException as e:
        print(f"[Linha {index}] Botão consultar não encontrado: {str(e)}")
        return False
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao tentar clicar no botão consultar: {str(e)}")
        return False

# Nova função para verificar a tela atual
def verificar_tela_atual(driver, index):
    try:
        # Verificar se está na tela de consulta (campo de documento presente)
        campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
            )
            print(f"[Linha {index}] Tela atual: Consulta")
            return "consulta"
        except TimeoutException:
            pass

        # Verificar se está na tela de seleção de conta
        select_conta_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, select_conta_xpath))
            )
            print(f"[Linha {index}] Tela atual: Seleção de conta")
            return "selecao_conta"
        except TimeoutException:
            pass

        # Verificar se está na tela de formulário
        form_xpath = "//form"
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, form_xpath))
            )
            print(f"[Linha {index}] Tela atual: Formulário")
            return "formulario"
        except TimeoutException:
            pass

        print(f"[Linha {index}] Tela atual desconhecida")
        return "desconhecida"
    except Exception as e:
        print(f"[Linha {index}] Erro ao verificar tela atual: {e}")
        return "desconhecida"

def preencher_formulario(driver, actions, row, index, df: pd.DataFrame, tentativa=0):
    try:
        if tentativa >= 3:  # Limita o número de tentativas
            print(f"[Linha {index}] ❌ Número máximo de tentativas excedido")
            df.at[index, 'Observação'] = "Número máximo de tentativas excedido"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        logger.info(f"\n[Linha {index}] Iniciando preenchimento do formulário... (Tentativa {tentativa + 1})")
        
        # Espera o spinner desaparecer antes de começar
        esperar_spinner_desaparecer(driver, index)
        
        # Verifica se está na tela correta
        tela_atual = verificar_tela_atual(driver, index)
        
        if tela_atual == "formulario":
            print(f"[Linha {index}] Já está na tela de formulário")
        elif tela_atual == "selecao_conta":
            print(f"[Linha {index}] ⚠️ Está na tela de seleção de conta. Tentando selecionar conta...")
            if not selecionar_conta_por_cooperativa(driver, row['Cooperativa'], index):
                df.at[index, 'Observação'] = "Falha ao selecionar conta"
                df.to_excel(EXCEL_PATH, index=False)
                return None
        elif tela_atual == "consulta":
    print(f"[Linha {index}] Está na tela de consulta. Preenchendo documento...")

    campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
    try:
        campo_documento = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", campo_documento)
        time.sleep(1)
        campo_documento.clear()
        campo_documento.click()

        doc_original = str(row['Documento do cooperado']).strip()
        numeros = ''.join(filter(str.isdigit, doc_original))
        print(f"[Linha {index}] Preenchendo documento: {numeros}")

        for digito in numeros:
            campo_documento.send_keys(digito)
            time.sleep(0.1)

        valor_preenchido = campo_documento.get_attribute('value')
        if valor_preenchido and numeros in valor_preenchido:
            print(f"[Linha {index}] ✅ Documento preenchido com sucesso: {valor_preenchido}")
        else:
            print(f"[Linha {index}] ❌ Documento não preenchido corretamente")
            df.at[index, 'Observação'] = "Falha ao preencher documento"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        if not clicar_botao_consulta(driver, index):
            df.at[index, 'Observação'] = "Falha ao clicar no botão consultar"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        try:
            WebDriverWait(driver, 10).until(
                lambda d: verificar_tela_atual(d, index) != "consulta"
            )
            print(f"[Linha {index}] ✅ Tela mudou após clique em consultar")
        except TimeoutException:
            print(f"[Linha {index}] ❌ Tela não mudou após clique em consultar")
            df.at[index, 'Observação'] = "Tela não avançou após clique"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        return preencher_formulario(driver, actions, row, index, df, tentativa + 1)

    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao preencher documento: {e}")
        df.at[index, 'Observação'] = f"Erro ao preencher documento: {str(e)}"
        df.to_excel(EXCEL_PATH, index=False)
        return None
            
            # Clica no botão consultar
            if not clicar_botao_consulta(driver, index):
                df.at[index, 'Observação'] = "Falha ao clicar no botão consultar"
                df.to_excel(EXCEL_PATH, index=False)
                return None
                
            # Aguarda a tela mudar
            time.sleep(2)
            return preencher_formulario(driver, actions, row, index, df, tentativa + 1)
        else:
            print(f"[Linha {index}] ⚠️ Tela desconhecida. Tentando voltar...")
            driver.get(BASE_URL)
            time.sleep(2)
            esperar_spinner_desaparecer(driver, index)
            return preencher_formulario(driver, actions, row, index, df, tentativa + 1)
        
        # Verifica se está realmente na tela de formulário
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//form"))
            )
        except TimeoutException:
            print(f"[Linha {index}] ⚠️ Não foi possível confirmar tela de formulário")
            df.at[index, 'Observação'] = "Falha ao carregar formulário"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        required_fields = ['Documento do cooperado', 'Protocolo PLAD', 'Categoria', 'Serviço', 'Cooperativa']
        for field in required_fields:
            if pd.isna(row[field]) or not str(row[field]).strip():
                error_msg = f"Campo '{field}' inválido ou ausente"
                logger.error(f"[Linha {index}] {error_msg}")
                df.at[index, 'Observação'] = error_msg
                df.to_excel(EXCEL_PATH, index=False)
                return None

        doc_original = str(row['Documento do cooperado']).strip()
        doc_formatado = formatar_documento(doc_original)
        logger.info(f"[Linha {index}] Documento original: {doc_original}")
        logger.info(f"[Linha {index}] Documento formatado: {doc_formatado}")
        
        protocolo_plad = str(row['Protocolo PLAD']).strip()
        categoria = str(row['Categoria']).strip()
        servico = normalizar_servico(str(row['Serviço']).strip())
        cooperativa = str(row['Cooperativa']).strip()
        
        MENSAGEM_PADRAO = "Registro de atendimento realizado na Plataforma de Atendimento Digital via automação"
        
        observacao = str(row.get('Observação', '')).strip()
        if (pd.isna(row.get('Observação')) or 
            observacao.lower() == 'nan' or 
            not observacao or 
            len(observacao) < 10):
            descricao = MENSAGEM_PADRAO
            if observacao and len(observacao) < 10:
                print(f"[Linha {index}] Observação '{observacao}' tem menos de 10 caracteres. Usando mensagem padrão.")
        else:
            descricao = observacao

        print(f"[Linha {index}] Aguardando campo de documento...")
        campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
        
        # Tenta encontrar o campo de documento com retry
        max_tentativas = 3
        for tentativa in range(max_tentativas):
            try:
                campo_documento = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
                )
                print(f"[Linha {index}] Campo de documento encontrado")
                break
            except TimeoutException:
                if tentativa < max_tentativas - 1:
                    print(f"[Linha {index}] Tentativa {tentativa + 1} de encontrar campo de documento falhou, tentando novamente...")
                    driver.refresh()
                    time.sleep(2)
                    esperar_spinner_desaparecer(driver, index)
                else:
                    raise

        campo_documento.clear()
        numeros = ''.join(filter(str.isdigit, doc_original))
        for digito in numeros:
            campo_documento.send_keys(digito)
        campo_documento.send_keys(Keys.TAB)

        print(f"[Linha {index}] Documento preenchido: {doc_formatado}")

        print(f"[Linha {index}] Aguardando botão de consulta...")
        if not clicar_botao_consulta(driver, index):
            df.at[index, 'Observação'] = "Falha ao clicar no botão consultar"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        if verificar_pessoa_nao_encontrada(driver, index):
            df.at[index, 'Observação'] = "Pessoa não identificada como cooperada!"
            df.to_excel(EXCEL_PATH, index=False)
            print(f"[Linha {index}] ℹ️ Observação atualizada na planilha")
            return None

        print(f"[Linha {index}] Aguardando botão de seleção de conta...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[2]/sc-button/button'))
        ).click()
        print(f"[Linha {index}] Botão de seleção de conta clicado")

        if not selecionar_conta_por_cooperativa(driver, cooperativa, index):
            print(f"[Linha {index}] Não foi possível continuar sem a conta correta")
            df.at[index, 'Observação'] = "Conta não encontrada para a cooperativa"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        print(f"[Linha {index}] Aguardando botão de categoria...")
        xpath_categoria = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[4]/div[1]/sc-card/div/div/div/div'
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath_categoria)))
        actions.move_to_element(driver.find_element(By.XPATH, xpath_categoria)).click().perform()
        print(f"[Linha {index}] Botão de categoria clicado")

        print(f"[Linha {index}] Aguardando botão de registro de chamado...")
        registro_xpath = '/html/body/div[1]/sc-app/sc-register-ticket-button/div/div/div/button'
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, registro_xpath)))
        actions.move_to_element(driver.find_element(By.XPATH, registro_xpath)).click().perform()
        print(f"[Linha {index}] Botão de registro de chamado clicado")

        # Aguardar o formulário estar completamente carregado
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//form"))
        )
        # Garantir que qualquer spinner tenha desaparecido
        esperar_spinner_desaparecer(driver, index)

        campos = {
            'tipo': {
                'xpath': '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[3]/sc-form-field/div/input',
                'valor': 'Chat Receptivo'
            },
            'categoria': {
                'xpath': '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[4]/sc-form-field/div/input',
                'valor': categoria
            },
            'subcategoria': {
                'xpath': '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[5]/sc-form-field/div/input',
                'valor': 'Api Sicoob'
            },
            'servico': {
                'xpath': '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[6]/sc-form-field/div/input',
                'valor': servico
            }
        }

        for campo_nome, campo_info in campos.items():
            print(f"[Linha {index}] Preenchendo {campo_nome}...")
            preencher_campo_com_js(driver, campo_info['xpath'], campo_info['valor'])
            print(f"[Linha {index}] {campo_nome} preenchido com: {campo_info['valor']}")

        print(f"[Linha {index}] Preenchendo Canal de autoatendimento...")
        canal_autoatendimento_xpath = "//sc-form-field[div/label[contains(text(), 'Canal de autoatendimento')]]/div/select"
        selecionar_opcao_select(driver, canal_autoatendimento_xpath, "Não se aplica")
        print(f"[Linha {index}] Canal de autoatendimento selecionado")

        print(f"[Linha {index}] Preenchendo Protocolo...")
        protocolo_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[8]/sc-additional-service-data/form/div/div[2]/sc-form-field/div/input'
        campo_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        campo_protocolo.clear()
        campo_protocolo.send_keys(protocolo_plad)
        campo_protocolo.send_keys(Keys.TAB)
        print(f"[Linha {index}] Protocolo preenchido: {protocolo_plad}")

        print(f"[Linha {index}] Preenchendo Descrição...")
        descricao_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[9]/sc-form-field/div/textarea'
        try:
            campo_descricao = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, descricao_xpath))
            )
            
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_descricao)
            
            campo_descricao.clear()
            campo_descricao.send_keys(descricao)
            print(f"[Linha {index}] Descrição preenchida: {descricao[:50]}..." if len(descricao) > 50 else f"[Linha {index}] Descrição preenchida: {descricao}")
                
        except TimeoutException as e:
            print(f"[Linha {index}] Timeout ao encontrar campo de descrição: {str(e)}")
            raise FormularioError(f"Timeout ao encontrar campo de descrição: {str(e)}")
        except NoSuchElementException as e:
            print(f"[Linha {index}] Campo de descrição não encontrado: {str(e)}")
            raise FormularioError(f"Campo de descrição não encontrado: {str(e)}")
        except Exception as e:
            print(f"[Linha {index}] Erro ao encontrar campo de descrição: {str(e)}")
            raise FormularioError(f"Erro ao encontrar campo de descrição: {str(e)}")

        print(f"[Linha {index}] Aguardando botão Registrar ficar habilitado...")
        registrar_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[20]/sc-button/button'
        WebDriverWait(driver, 30).until(
            lambda d: d.find_element(By.XPATH, registrar_xpath).is_enabled()
        )
        botao_registrar = driver.find_element(By.XPATH, registrar_xpath)
        actions.move_to_element(botao_registrar).click().perform()
        print(f"[Linha {index}] Botão Registrar clicado")

        print(f"[Linha {index}] Aguardando botão Confirmar...")
        confirmar_xpath = '/html/body/div[3]/div[4]/div/sc-register-ticket-modal/sc-modal/div/div/sc-modal-footer/div/div/div[2]/sc-button/button'
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        actions.move_to_element(botao_confirmar).click().perform()
        print(f"[Linha {index}] Botão Confirmar clicado")

        print(f"[Linha {index}] Capturando número do protocolo...")
        protocolo_xpath = '/html/body/div[3]/div[4]/div/sc-view-ticket-data/sc-actionbar/div/div/div[2]/form/div/div[2]/sc-card/div/sc-card-content/div/div/div[1]/h5'
        elemento_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        numero_protocolo = elemento_protocolo.text.strip()
        logger.info(f"[Linha {index}] Protocolo capturado: {numero_protocolo}")

        try:
            df.at[index, 'Protocolo Visão'] = numero_protocolo
            df.to_excel(EXCEL_PATH, index=False)
            logger.info(f"[Linha {index}] Protocolo salvo na planilha com sucesso!")
        except Exception as e:
            logger.error(f"[Linha {index}] Erro ao salvar protocolo na planilha: {e}")
        
        return numero_protocolo

    except TimeoutException as e:
        log_error(e, "preencher formulário", index, df)
        raise FormularioError(f"Timeout durante preenchimento do formulário: {str(e)}")
    except NoSuchElementException as e:
        log_error(e, "preencher formulário", index, df)
        raise FormularioError(f"Elemento não encontrado durante preenchimento: {str(e)}")
    except Exception as e:
        log_error(e, "preencher formulário", index, df)
        raise FormularioError(f"Erro ao preencher formulário: {str(e)}")

def tentar_preencher_formulario(driver, actions, row, index, df, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            if tentativa > 0:
                print(f"[Linha {index}] 🔄 Tentativa {tentativa + 1} de {max_tentativas}")
                driver.refresh()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            return preencher_formulario(driver, actions, row, index, df)
            
        except FormularioError as e:
            print(f"[Linha {index}] ❌ Erro na tentativa {tentativa + 1}: {str(e)}")
            if tentativa == max_tentativas - 1:
                print(f"[Linha {index}] ❌ Todas as tentativas falharam")
                df.at[index, 'Observação'] = f"Erro após {max_tentativas} tentativas: {str(e)}"
                df.to_excel(EXCEL_PATH, index=False)
                return None
    return None

def finalizar_atendimento(driver, index, df: pd.DataFrame):
    try:
        logger.info(f"[Linha {index}] 🔄 Iniciando finalização do atendimento...")
        
        if not esperar_modal_desaparecer(driver, index):
            raise FinalizacaoError("Modal não desapareceu a tempo")
        
        logger.info(f"[Linha {index}] Clicando no botão 'Finalizar atendimento'...")
        finalizar_xpath = '/html/body/div[3]/div[4]/div/sc-view-ticket-data/sc-actionbar/div/div/div[2]/form/div/div[5]/sc-button/button'
        
        botao_finalizar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, finalizar_xpath))
        )
        actions = ActionChains(driver)
        actions.move_to_element(botao_finalizar).click().perform()
        
        logger.info(f"[Linha {index}] Aguardando modal de confirmação...")
        confirmar_xpath = '/html/body/div[3]/div[2]/div/sc-end-service-modal/sc-modal/div/div/main/div/div[4]/button'
        
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        actions.move_to_element(botao_confirmar).click().perform()
        
        logger.info(f"[Linha {index}] Aguardando retorno à tela inicial...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        logger.info(f"[Linha {index}] ✅ Atendimento finalizado com sucesso!")
        return True
        
    except TimeoutException as e:
        log_error(e, "finalização do atendimento", index, df)
        raise FinalizacaoError(f"Timeout durante finalização: {str(e)}")
    except NoSuchElementException as e:
        log_error(e, "finalização do atendimento", index, df)
        raise FinalizacaoError(f"Elemento não encontrado durante finalização: {str(e)}")
    except Exception as e:
        log_error(e, "finalização do atendimento", index, df)
        raise FinalizacaoError(f"Falha ao finalizar atendimento: {str(e)}")

def main():
    try:
        logger.info("🚀 Iniciando sistema de registro de chamados...")
        
        logger.info("Carregando credenciais...")
        username, password = load_credentials()
        
        download_dir = os.path.dirname(os.path.abspath(__file__))
        
        logger.info("Inicializando navegador...")
        driver = setup_driver(download_dir)
        actions = ActionChains(driver)
        
        try:
            login(driver, username, password)
            
            logger.info("Carregando dados da planilha...")
            df = load_excel_data(EXCEL_PATH)
            required_columns = ['Documento do cooperado', 'Protocolo PLAD', 'Categoria', 'Serviço', 'Cooperativa']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Colunas ausentes no Excel: {missing_columns}")
            
            total_registros = len(df)
            logger.info(f"📊 Total de registros a processar: {total_registros}")
            
            registros_processados = 0
            registros_com_erro = 0
            
            for index, row in df.iterrows():
                try:
                    logger.info(f"\n{'='*50}")
                    logger.info(f"[Linha {index}] 📝 Iniciando processamento do registro {index + 1}/{total_registros}")
                    
                    if tentar_preencher_formulario(driver, actions, row, index, df):
                        if finalizar_atendimento(driver, index, df):
                            registros_processados += 1
                            logger.info(f"[Linha {index}] ✅ Registro processado com sucesso!")
                        else:
                            registros_com_erro += 1
                            logger.error(f"[Linha {index}] ❌ Erro ao finalizar atendimento")
                    else:
                        registros_com_erro += 1
                        logger.error(f"[Linha {index}] ❌ Erro ao preencher formulário")
                    
                except Exception as e:
                    registros_com_erro += 1
                    log_error(e, "processamento do registro", index, df)
                    continue
            
            logger.info("\n" + "="*50)
            logger.info("📊 RELATÓRIO FINAL:")
            logger.info(f"Total de registros: {total_registros}")
            logger.info(f"Registros processados com sucesso: {registros_processados}")
            logger.info(f"Registros com erro: {registros_com_erro}")
            logger.info("="*50)
            
        finally:
            logger.info("Fechando navegador...")
            try:
                time.sleep(1)
                driver.quit()
            except Exception as e:
                logger.warning(f"Erro ao fechar o navegador: {str(e)}")
            
    except Exception as e:
        log_error(e, "execução geral do sistema")
        if 'driver' in locals():
            try:
                driver.quit()
            except Exception as e:
                logger.warning(f"Erro ao fechar o navegador: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical("❌ Sistema encerrado com erro crítico!", exc_info=True)
        raise