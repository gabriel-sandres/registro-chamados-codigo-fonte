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
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(os.path.dirname(__file__), "planilha_registro.xlsx"))
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

def login(driver: webdriver.Chrome, username: str, password: str, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            logger.info(f"🔄 Iniciando processo de login... (Tentativa {tentativa + 1}/{max_tentativas})")
            driver.get(BASE_URL)
            
            # Espera a página carregar completamente
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # Verifica se está realmente logado
            try:
                elementos_logado = [
                    "//sc-sidebar-container",
                    "//sc-app",
                    "//sc-template"
                ]
                
                for elemento in elementos_logado:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, elemento))
                        )
                    except TimeoutException:
                        raise NoSuchElementException(f"Elemento {elemento} não encontrado")
                
                logger.info("✅ Verificação de login bem-sucedida")
                return True
                
            except NoSuchElementException:
                logger.info("⚠️ Não está logado. Iniciando processo de login...")
                
                # Verifica se o campo de login está presente
                try:
                    campo_username = WebDriverWait(driver, 30).until(
                        EC.visibility_of_element_located((By.ID, 'username'))
                    )
                    campo_username.clear()
                    campo_username.send_keys(username)
                    
                    campo_password = WebDriverWait(driver, 30).until(
                        EC.visibility_of_element_located((By.ID, 'password'))
                    )
                    campo_password.clear()
                    campo_password.send_keys(password)
                    
                    logger.info("Clicando no botão de login...")
                    botao_login = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.ID, 'kc-login'))
                    )
                    botao_login.click()
                    
                    # Aguarda o QR code desaparecer
                    try:
                        logger.info("Aguardando QR code desaparecer...")
                        WebDriverWait(driver, 300).until(
                            EC.invisibility_of_element_located((By.ID, "qr-code"))
                        )
                        
                        # Verifica novamente se está logado
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, elementos_logado[0]))
                        )
                        
                        logger.info("✅ Login realizado com sucesso!")
                        return True
                        
                    except TimeoutException:
                        logger.warning("QR code não desapareceu a tempo")
                        if tentativa < max_tentativas - 1:
                            logger.info("Tentando novamente...")
                            continue
                        else:
                            raise LoginError("QR code não desapareceu após várias tentativas")
                            
                except TimeoutException as e:
                    logger.error(f"Timeout durante o login: {str(e)}")
                    if tentativa < max_tentativas - 1:
                        logger.info("Tentando novamente...")
                        continue
                    else:
                        raise LoginError(f"Timeout durante o login após {max_tentativas} tentativas: {str(e)}")
                except NoSuchElementException as e:
                    logger.error(f"Elemento não encontrado durante o login: {str(e)}")
                    if tentativa < max_tentativas - 1:
                        logger.info("Tentando novamente...")
                        continue
                    else:
                        raise LoginError(f"Elemento não encontrado durante o login após {max_tentativas} tentativas: {str(e)}")
                except Exception as e:
                    logger.error(f"Erro inesperado durante o login: {str(e)}")
                    if tentativa < max_tentativas - 1:
                        logger.info("Tentando novamente...")
                        continue
                    else:
                        raise LoginError(f"Falha no login após {max_tentativas} tentativas: {str(e)}")
            
        except Exception as e:
            logger.error(f"Erro geral durante o login: {str(e)}")
            if tentativa < max_tentativas - 1:
                logger.info("Tentando novamente...")
                continue
            else:
                raise LoginError(f"Falha no login após {max_tentativas} tentativas: {str(e)}")
    
    raise LoginError(f"Falha no login após {max_tentativas} tentativas")

def limpar_e_preencher(campo, valor):
    campo.click()
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    campo.send_keys(valor)

def preencher_com_sugestao(campo, valor, driver):
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, campo.get_attribute("id"))))
        actions = ActionChains(driver)
        actions.move_to_element(campo).click().perform()
        
        campo.clear()
        campo.send_keys(valor[:3])
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//option[contains(text(), '{valor}')] | //li[contains(text(), '{valor}')]"))
        )
        campo.send_keys(Keys.ARROW_DOWN)
        campo.send_keys(Keys.ENTER)
        
        # Verifica se o campo foi preenchido corretamente (ng-valid)
        WebDriverWait(driver, 10).until(
            lambda d: 'ng-valid' in d.find_element(By.ID, campo.get_attribute("id")).get_attribute('class').split()
        )
    except TimeoutException as e:
        raise FormularioError(f"Timeout ao localizar sugestão para '{valor}' ou validar preenchimento: {str(e)}")
    except NoSuchElementException as e:
        raise FormularioError(f"Sugestão para '{valor}' não encontrada: {str(e)}")
    except ElementClickInterceptedException as e:
        logger.warning(f"Clique interceptado ao preencher '{valor}': {str(e)}")
        driver.execute_script("arguments[0].click();", campo)
        campo.clear()
        campo.send_keys(valor[:3])
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//option[contains(text(), '{valor}')] | //li[contains(text(), '{valor}')]"))
        )
        campo.send_keys(Keys.ARROW_DOWN)
        campo.send_keys(Keys.ENTER)
        WebDriverWait(driver, 10).until(
            lambda d: 'ng-valid' in d.find_element(By.ID, campo.get_attribute("id")).get_attribute('class').split()
        )
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
        
        # Verifica se a opção foi selecionada corretamente
        WebDriverWait(driver, 10).until(
            lambda d: 'ng-valid' in d.find_element(By.XPATH, select_xpath).get_attribute('class').split()
        )
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
                    WebDriverWait(driver, 10).until(
                        lambda d: 'ng-valid' in d.find_element(By.XPATH, select_xpath).get_attribute('class').split()
                    )
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
        
        if not esperar_spinner_desaparecer(driver, index):
            print(f"[Linha {index}] ⚠️ Spinner não desapareceu a tempo")
            return False
        
        select_element = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, select_xpath))
        )
        
        driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
        
        try:
            select_element.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", select_element)
        
        options = select_element.find_elements(By.TAG_NAME, 'option')
        conta_encontrada = False
        for option in options:
            texto_opcao = option.text.strip()
            if f"Coop: {cooperativa}" in texto_opcao:
                print(f"[Linha {index}] Conta encontrada: {texto_opcao}")
                select = Select(select_element)
                select.select_by_visible_text(texto_opcao)
                WebDriverWait(driver, 10).until(
                    lambda d: d.find_element(By.XPATH, select_xpath).get_attribute('value')
                )
                print(f"[Linha {index}] ✅ Conta selecionada com sucesso")
                return True
        if not conta_encontrada:
            print(f"[Linha {index}] ⚠️ ATENÇÃO: Nenhuma conta encontrada para cooperativa {cooperativa}")
            return False
            
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao selecionar conta: {str(e)}")
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

def esperar_spinner_desaparecer(driver, index, timeout=30, check_interval=1):
    try:
        spinner_xpath = "//div[contains(@class, 'ngx-spinner-overlay')]"
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                spinner = driver.find_element(By.XPATH, spinner_xpath)
                if not spinner.is_displayed():
                    print(f"[Linha {index}] ✅ Spinner desapareceu")
                    return True
            except NoSuchElementException:
                print(f"[Linha {index}] ✅ Spinner não encontrado")
                return True
            time.sleep(check_interval)
        
        print(f"[Linha {index}] ⚠️ Timeout ao esperar spinner desaparecer após {timeout} segundos")
        return False
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao esperar spinner desaparecer: {str(e)}")
        return False

def clicar_botao_consulta(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão consultar...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'
        
        botao = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, botao_xpath))
        )
        
        driver.execute_script("arguments[0].scrollIntoView(true);", botao)
        driver.execute_script("arguments[0].click();", botao)
        print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
        return True
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao tentar clicar no botão consultar: {str(e)}")
        return False

def verificar_tela_atual(driver, index):
    try:
        campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
            )
            print(f"[Linha {index}] Tela atual: Consulta")
            return "consulta"
        except TimeoutException:
            pass

        select_conta_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, select_conta_xpath))
            )
            print(f"[Linha {index}] Tela atual: Seleção de conta")
            return "selecao_conta"
        except TimeoutException:
            pass

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

def clicar_botao_abrir(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão Abrir...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[2]/sc-button/button'
        
        botao = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, botao_xpath))
        )
        
        driver.execute_script("arguments[0].scrollIntoView(true);", botao)
        driver.execute_script("arguments[0].click();", botao)
        print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
        return True
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao tentar clicar no botão Abrir: {str(e)}")
        return False

def clicar_menu_cobranca(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no menu 'Cobrança'...")
        cobranca_xpaths = [
            '//*[@id="products"]/div[10]/sc-card/div/div/div/div',
            '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[4]/div[10]/sc-card/div/div/div/div',
            "//h6[contains(text(), 'Cobrança')]",
            "//div[contains(@class, 'title-products')]//h6[contains(text(), 'Cobrança')]"
        ]
        
        menu_cobranca = None
        for xpath in cobranca_xpaths:
            try:
                menu_cobranca = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                if menu_cobranca:
                    print(f"[Linha {index}] Menu 'Cobrança' encontrado usando XPath: {xpath}")
                    break
            except:
                continue
        
        if not menu_cobranca:
            raise NoSuchElementException("Menu 'Cobrança' não encontrado com nenhum dos XPaths")
        
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_cobranca)
        driver.execute_script("arguments[0].click();", menu_cobranca)
        print(f"[Linha {index}] ✅ Menu 'Cobrança' clicado via JavaScript")
        return True
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao clicar no menu 'Cobrança': {str(e)}")
        return False

def clicar_botao_registro_chamado(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão de registro de chamado...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-register-ticket-button/div/div/div/button'
        
        botao = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, botao_xpath))
        )
        
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
        driver.execute_script("arguments[0].click();", botao)
        print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
        return True
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao clicar no botão de registro de chamado: {str(e)}")
        return False

def preencher_campos_formulario(driver, actions, row, index, df: pd.DataFrame) -> Optional[str]:
    """Preenche os campos do formulário de registro de chamado."""
    try:
        print(f"[Linha {index}] Preenchendo campos do formulário...")

        def preencher_campo_com_validacao(campo_id, valor, max_tentativas=3):
            for tentativa in range(max_tentativas):
                try:
                    campo = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, campo_id))
                    )
                    driver.execute_script("""
                        arguments[0].value = '';
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """, campo, valor)
                    campo.send_keys(Keys.ENTER)
                    WebDriverWait(driver, 10).until(
                        lambda d: 'ng-valid' in d.find_element(By.ID, campo_id).get_attribute('class').split()
                    )
                    print(f"[Linha {index}] Campo {campo_id} preenchido com: {valor}")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] Tentativa {tentativa + 1} falhou para {campo_id}: {str(e)}")
                    if tentativa < max_tentativas - 1:
                        time.sleep(1)
                        continue
                    raise FormularioError(f"Falha ao preencher {campo_id} após {max_tentativas} tentativas: {str(e)}")

        # Tipo de atendimento
        preencher_campo_com_validacao("serviceTypeId", "Chat Receptivo")

        # Categoria
        preencher_campo_com_validacao("categoryId", row['Categoria'])

        # Subcategoria
        preencher_campo_com_validacao("subCategoryId", "Api Sicoob")

        # Serviço
        valor_servico = normalizar_servico(row['Serviço'])
        preencher_campo_com_validacao("serviceId", valor_servico)

        # Canal de autoatendimento
        print(f"[Linha {index}] Preenchendo Canal de autoatendimento...")
        try:
            selecionar_opcao_select(driver, '//*[@id="Canal De Autoatendimento"]', "não se aplica")
        except Exception as e:
            print(f"[Linha {index}] ⚠️ Campo Canal de autoatendimento não encontrado: {str(e)}")

        # Protocolo PLAD
        print(f"[Linha {index}] Preenchendo Protocolo PLAD...")
        campo_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Protocolo Plad"))
        )
        campo_protocolo.clear()
        campo_protocolo.send_keys(str(row['Protocolo PLAD']))
        WebDriverWait(driver, 10).until(
            lambda d: 'ng-valid' in d.find_element(By.ID, "Protocolo Plad").get_attribute('class').split()
        )
        print(f"[Linha {index}] Protocolo PLAD preenchido: {row['Protocolo PLAD']}")

        # Descrição
        print(f"[Linha {index}] Preenchendo Descrição...")
        MENSAGEM_PADRAO = "Chamado da Plataforma de atendimento digital registrado via automação"
        observacao = str(row.get('Observação', '')).strip()
        descricao = MENSAGEM_PADRAO if (pd.isna(row.get('Observação')) or observacao.lower() == 'nan' or not observacao or len(observacao) < 10) else observacao
        if observacao and len(observacao) < 10:
            print(f"[Linha {index}] Observação '{observacao}' tem menos de 10 caracteres. Usando mensagem padrão.")
        
        campo_descricao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "description"))
        )
        campo_descricao.clear()
        campo_descricao.send_keys(descricao)
        WebDriverWait(driver, 10).until(
            lambda d: 'ng-valid' in d.find_element(By.ID, "description").get_attribute('class').split()
        )
        print(f"[Linha {index}] Descrição preenchida: {descricao[:50]}..." if len(descricao) > 50 else f"[Linha {index}] Descrição preenchida: {descricao}")

        # Aguarda o botão Registrar ficar habilitado
        print(f"[Linha {index}] Aguardando botão Registrar ficar habilitado...")
        registrar_xpath = '//*[@id="actionbar hide"]/div/div[2]/form/div/div[20]/sc-button/button'
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, registrar_xpath))
        )
        botao_registrar = driver.find_element(By.XPATH, registrar_xpath)
        driver.execute_script("arguments[0].click();", botao_registrar)
        print(f"[Linha {index}] Botão Registrar clicado")

        # Aguarda e clica no botão Confirmar
        print(f"[Linha {index}] Aguardando botão Confirmar...")
        confirmar_xpath = '//*[@id="modal"]/div/sc-modal-footer/div/div/div[2]/sc-button/button'
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        driver.execute_script("arguments[0].click();", botao_confirmar)
        print(f"[Linha {index}] Botão Confirmar clicado")

        # Captura o número do protocolo
        print(f"[Linha {index}] Capturando número do protocolo...")
        protocolo_xpath = '//*[@id="actionbar hide"]/div/div[2]/form/div/div[2]/sc-card/div/sc-card-content/div/div/div[1]/h5'
        elemento_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        numero_protocolo = elemento_protocolo.text.strip()
        print(f"[Linha {index}] Protocolo capturado: {numero_protocolo}")

        # Salva o protocolo na coluna F
        df.at[index, 'Protocolo PLAD'] = numero_protocolo
        df.to_excel(EXCEL_PATH, index=False)
        print(f"[Linha {index}] Protocolo salvo na planilha: {numero_protocolo}")

        return numero_protocolo

    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao preencher campos do formulário: {str(e)}")
        df.at[index, 'Observação'] = f"Erro ao preencher campos do formulário: {str(e)}"
        df.to_excel(EXCEL_PATH, index=False)
        return None

def preencher_formulario(driver, actions, row, index, df: pd.DataFrame, tentativa=0, max_tentativas_por_tela=3):
    try:
        if tentativa >= max_tentativas_por_tela:
            print(f"[Linha {index}] ❌ Número máximo de tentativas excedido para esta tela")
            df.at[index, 'Observação'] = "Número máximo de tentativas excedido para esta tela"
            df.to_excel(EXCEL_PATH, index=False)
            return None

        logger.info(f"\n[Linha {index}] Iniciando preenchimento do formulário... (Tentativa {tentativa + 1})")

        esperar_spinner_desaparecer(driver, index)
        tela_atual = verificar_tela_atual(driver, index)

        if tela_atual == "formulario":
            print(f"[Linha {index}] Já está na tela de formulário")
            return preencher_campos_formulario(driver, actions, row, index, df)

        elif tela_atual == "selecao_conta":
            print(f"[Linha {index}] ⚠️ Está na tela de seleção de conta. Tentando selecionar conta...")
            if not selecionar_conta_por_cooperativa(driver, row['Cooperativa'], index):
                df.at[index, 'Observação'] = "Falha ao selecionar conta"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            select_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
            try:
                select_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, select_xpath))
                )
                valor_selecionado = select_element.get_attribute('value')
                if not valor_selecionado:
                    print(f"[Linha {index}] ⚠️ Conta não foi selecionada corretamente")
                    df.at[index, 'Observação'] = "Conta não foi selecionada corretamente"
                    df.to_excel(EXCEL_PATH, index=False)
                    return None
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro ao verificar seleção da conta: {str(e)}")
                df.at[index, 'Observação'] = f"Erro ao verificar seleção da conta: {str(e)}"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            if not clicar_menu_cobranca(driver, index):
                df.at[index, 'Observação'] = "Falha ao clicar no menu Cobrança"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            if not clicar_botao_registro_chamado(driver, index):
                df.at[index, 'Observação'] = "Falha ao clicar no botão de registro de chamado"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            try:
                categoria_xpath = '//*[@id="categoryId"]'
                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, categoria_xpath))
                )
                print(f"[Linha {index}] ✅ Formulário aberto e pronto para preenchimento")
                return preencher_campos_formulario(driver, actions, row, index, df)
            except Exception as e:
                print(f"[Linha {index}] ❌ Formulário não abriu corretamente ou campos não apareceram: {str(e)}")
                df.at[index, 'Observação'] = "Formulário não abriu corretamente ou campos não apareceram"
                df.to_excel(EXCEL_PATH, index=False)
                return None

        elif tela_atual == "consulta":
            print(f"[Linha {index}] Está na tela de consulta. Preenchendo documento...")
            campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
            try:
                campo_documento = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, campo_documento_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", campo_documento)
                campo_documento.clear()
                campo_documento.click()
                
                doc_original = str(row['Documento do cooperado']).strip()
                doc_formatado = formatar_documento(doc_original)
                print(f"[Linha {index}] Preenchendo documento: {doc_formatado}")
                
                for digito in doc_formatado:
                    campo_documento.send_keys(digito)
                
                valor_preenchido = campo_documento.get_attribute('value')
                if not valor_preenchido:
                    driver.execute_script(f"arguments[0].value = '{doc_formatado}';", campo_documento)
                    valor_preenchido = campo_documento.get_attribute('value')
                
                valor_preenchido_numeros = ''.join(filter(str.isdigit, valor_preenchido))
                numeros = ''.join(filter(str.isdigit, doc_original))
                if valor_preenchido_numeros == numeros:
                    print(f"[Linha {index}] ✅ Documento preenchido com sucesso: {valor_preenchido}")
                    
                    if not clicar_botao_consulta(driver, index):
                        df.at[index, 'Observação'] = "Falha ao clicar no botão consultar"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None
                    
                    if verificar_pessoa_nao_encontrada(driver, index):
                        df.at[index, 'Observação'] = "Pessoa não encontrada"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None
                    
                    if not clicar_botao_abrir(driver, index):
                        df.at[index, 'Observação'] = "Falha ao clicar no botão Abrir"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None
                    
                    tela_atual = verificar_tela_atual(driver, index)
                    if tela_atual == "selecao_conta":
                        print(f"[Linha {index}] ✅ Tela mudou para seleção de conta")
                        return preencher_formulario(driver, actions, row, index, df)
                    else:
                        df.at[index, 'Observação'] = "Tela não mudou para seleção de conta após clicar em Abrir"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None
                else:
                    print(f"[Linha {index}] ❌ Documento não preenchido corretamente")
                    df.at[index, 'Observação'] = "Falha ao preencher documento"
                    df.to_excel(EXCEL_PATH, index=False)
                    return None
            except Exception as e:
                print(f"[Linha {index}] ❌ Erro ao preencher documento: {str(e)}")
                df.at[index, 'Observação'] = f"Erro ao preencher documento: {str(e)}"
                df.to_excel(EXCEL_PATH, index=False)
                return None
        else:
            print(f"[Linha {index}] ⚠️ Tela desconhecida. Tentando voltar...")
            driver.get(BASE_URL)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            return preencher_formulario(driver, actions, row, index, df, tentativa + 1)

    except Exception as e:
        print(f"[Linha {index}] ❌ Erro geral na função preencher_formulario: {str(e)}")
        df.at[index, 'Observação'] = f"Erro geral na função preencher_formulario: {str(e)}"
        df.to_excel(EXCEL_PATH, index=False)
        return None

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
        driver.execute_script("arguments[0].click();", botao_finalizar)
        
        logger.info(f"[Linha {index}] Aguardando modal de confirmação...")
        confirmar_xpath = '/html/body/div[3]/div[2]/div/sc-end-service-modal/sc-modal/div/div/main/div/div[4]/button'
        
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        driver.execute_script("arguments[0].click();", botao_confirmar)
        
        logger.info(f"[Linha {index}] Aguardando retorno à tela inicial...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        logger.info(f"[Linha {index}] ✅ Atendimento finalizado com sucesso!")
        return True
        
    except Exception as e:
        log_error(e, "finalização do atendimento", index, df)
        return False

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
                    
                    protocolo = tentar_preencher_formulario(driver, actions, row, index, df)
                    if protocolo:
                        if finalizar_atendimento(driver, index, df):
                            registros_processados += 1
                            logger.info(f"[Linha {index}] ✅ Registro processado com sucesso!")
                        else:
                            registros_com_erro += 1
                            logger.error(f"[Linha {index}] ❌ Erro ao finalizar atendimento")
                    else:
                        registros_com_erro += 1
                        logger.error(f"[Linha {index}] ❌ Erro ao preencher formulário")
                    
                    # Reinicia o processo carregando a URL inicial
                    logger.info(f"[Linha {index}] Reiniciando processo com URL inicial...")
                    driver.get(BASE_URL)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    esperar_spinner_desaparecer(driver, index)
                    
                except Exception as e:
                    registros_com_erro += 1
                    log_error(e, "processamento do registro", index, df)
                    # Tenta reiniciar a URL mesmo em caso de erro
                    try:
                        driver.get(BASE_URL)
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        esperar_spinner_desaparecer(driver, index)
                    except Exception as e2:
                        logger.error(f"[Linha {index}] Erro ao reiniciar URL: {str(e2)}")
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
        logger.critical(f"❌ Sistema encerrado com erro crítico! {str(e)}", exc_info=True)
        raise