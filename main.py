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

# Tempo de espera padrão entre ações de preenchimento de campos
FIELD_DELAY = 2
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
                # Tenta encontrar elementos que só existem quando logado
                elementos_logado = [
                    "//sc-sidebar-container",  # Sidebar do sistema
                    "//sc-app",  # Container principal do app
                    "//sc-template"  # Template do sistema
                ]
                
                for elemento in elementos_logado:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, elemento))
                        )
                    except TimeoutException:
                        raise NoSuchElementException(f"Elemento {elemento} não encontrado")
                
                # Se chegou aqui, está logado
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
                        time.sleep(5)  # Aguarda um pouco para garantir que a página carregou
                        for elemento in elementos_logado:
                            WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, elemento))
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
        if not esperar_spinner_desaparecer(driver, index):
            print(f"[Linha {index}] ⚠️ Spinner não desapareceu a tempo")
            return False
        
        # Espera o select estar presente e clicável
        try:
            select_element = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, select_xpath))
            )
        except TimeoutException:
            print(f"[Linha {index}] ⚠️ Timeout ao aguardar select de conta")
            return False
        
        # Rola até o elemento para garantir que está visível
        driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
        time.sleep(1)
        
        # Tenta clicar no select primeiro
        try:
            select_element.click()
        except ElementClickInterceptedException:
            try:
                driver.execute_script("arguments[0].click();", select_element)
            except:
                actions = ActionChains(driver)
                actions.move_to_element(select_element).click().perform()
        
        # Espera as opções aparecerem
        time.sleep(1)
        
        try:
            options = select_element.find_elements(By.TAG_NAME, 'option')
        except NoSuchElementException:
            print(f"[Linha {index}] ⚠️ Não foi possível encontrar as opções do select")
            return False
        
        conta_encontrada = False
        for option in options:
            try:
                texto_opcao = option.text.strip()
                # Verifica se o texto da opção contém a cooperativa
                if f"Coop: {cooperativa}" in texto_opcao:
                    print(f"[Linha {index}] Conta encontrada: {texto_opcao}")
                    try:
                        # Tenta selecionar via Select
                        select = Select(select_element)
                        select.select_by_visible_text(texto_opcao)
                        time.sleep(2)
                        
                        # Verifica se a seleção foi bem sucedida
                        valor_selecionado = select_element.get_attribute('value')
                        if valor_selecionado:
                            print(f"[Linha {index}] ✅ Conta selecionada com sucesso")
                            return True
                    except:
                        # Se falhar, tenta via JavaScript
                        try:
                            driver.execute_script(f"arguments[0].value = '{option.get_attribute('value')}';", select_element)
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", select_element)
                            time.sleep(2)
                            
                            # Verifica se a seleção foi bem sucedida
                            valor_selecionado = select_element.get_attribute('value')
                            if valor_selecionado:
                                print(f"[Linha {index}] ✅ Conta selecionada com sucesso via JavaScript")
                                return True
                        except:
                            print(f"[Linha {index}] ⚠️ Falha ao selecionar conta via JavaScript")
                    
                    conta_encontrada = True
                    break
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro ao processar opção: {str(e)}")
                continue
        
        if not conta_encontrada:
            print(f"[Linha {index}] ⚠️ ATENÇÃO: Nenhuma conta encontrada para cooperativa {cooperativa}")
            return False
            
        # Aguarda um momento para garantir que a seleção foi processada
        time.sleep(2)
        
        # Verifica se a conta foi realmente selecionada
        try:
            valor_selecionado = select_element.get_attribute('value')
            if valor_selecionado:
                print(f"[Linha {index}] ✅ Conta selecionada com sucesso")
                return True
            else:
                print(f"[Linha {index}] ⚠️ Conta não foi selecionada corretamente")
                return False
        except Exception as e:
            print(f"[Linha {index}] ⚠️ Erro ao verificar seleção da conta: {str(e)}")
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
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro ao verificar spinner: {str(e)}")
            
            time.sleep(check_interval)
        
        print(f"[Linha {index}] ⚠️ Timeout ao esperar spinner desaparecer após {timeout} segundos")
        return False

    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao esperar spinner desaparecer: {str(e)}")
        return False

def aguardar_campo_valido(driver, elemento, index, timeout=10):
    """Espera o campo possuir a classe 'ng-valid' antes de prosseguir."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: "ng-valid" in elemento.get_attribute("class")
        )
        print(f"[Linha {index}] Campo validado com 'ng-valid'")
        return True
    except TimeoutException:
        print(
            f"[Linha {index}] ⚠️ Campo não ficou válido após {timeout} segundos"
        )
        return False

def clicar_botao_consulta(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão consultar...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'
        
        # Espera o botão estar presente e clicável
        try:
            botao = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, botao_xpath))
            )
            print(f"[Linha {index}] ✅ Botão consultar encontrado")
            
            # Verifica se o botão está desabilitado
            if botao.get_attribute("disabled"):
                print(f"[Linha {index}] ⚠️ Botão está desabilitado, aguardando habilitação...")
                # Aguarda até que o botão seja habilitado
                WebDriverWait(driver, 10).until(
                    lambda d: not d.find_element(By.XPATH, botao_xpath).get_attribute("disabled")
                )
                print(f"[Linha {index}] ✅ Botão foi habilitado")
            
        except TimeoutException:
            print(f"[Linha {index}] ❌ Timeout ao localizar botão consultar")
            return False
        
        # Rola até o botão
        driver.execute_script("arguments[0].scrollIntoView(true);", botao)
        time.sleep(1)
        
        # Tenta diferentes métodos de clique
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                print(f"[Linha {index}] Tentativa {tentativas + 1} de clicar no botão...")
                
                # Tenta clicar via JavaScript primeiro (mais confiável neste caso)
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                    driver.execute_script("arguments[0].click();", botao)
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via JavaScript: {str(e)}")
                
                # Tenta clicar normalmente
                try:
                    botao.click()
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado com sucesso")
                    return True
                except ElementClickInterceptedException:
                    print(f"[Linha {index}] ⚠️ Clique interceptado, tentando via ActionChains...")
                
                # Tenta clicar via ActionChains
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(botao).click().perform()
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado via ActionChains")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via ActionChains: {str(e)}")
                
                # Se chegou aqui, nenhum método funcionou
                tentativas += 1
                if tentativas < max_tentativas:
                    print(f"[Linha {index}] ⚠️ Tentando novamente em 1 segundo...")
                    time.sleep(1)
                else:
                    print(f"[Linha {index}] ❌ Todas as tentativas de clique falharam")
                    return False
                
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro durante tentativa de clique: {str(e)}")
                tentativas += 1
                if tentativas < max_tentativas:
                    time.sleep(1)
                else:
                    return False
        
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

def clicar_botao_abrir(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão Abrir...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[2]/sc-button/button'
        
        # Espera o botão estar presente e clicável
        try:
            botao = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, botao_xpath))
            )
            print(f"[Linha {index}] ✅ Botão Abrir encontrado")
        except TimeoutException:
            print(f"[Linha {index}] ❌ Timeout ao localizar botão Abrir")
            return False
        
        # Rola até o botão
        driver.execute_script("arguments[0].scrollIntoView(true);", botao)
        time.sleep(1)
        
        # Tenta diferentes métodos de clique
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                print(f"[Linha {index}] Tentativa {tentativas + 1} de clicar no botão...")
                
                # Tenta clicar via JavaScript primeiro
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                    driver.execute_script("arguments[0].click();", botao)
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via JavaScript: {str(e)}")
                
                # Tenta clicar normalmente
                try:
                    botao.click()
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado com sucesso")
                    return True
                except ElementClickInterceptedException:
                    print(f"[Linha {index}] ⚠️ Clique interceptado, tentando via ActionChains...")
                
                # Tenta clicar via ActionChains
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(botao).click().perform()
                    time.sleep(2)  # Aguarda efeito do clique
                    print(f"[Linha {index}] ✅ Botão clicado via ActionChains")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via ActionChains: {str(e)}")
                
                # Se chegou aqui, nenhum método funcionou
                tentativas += 1
                if tentativas < max_tentativas:
                    print(f"[Linha {index}] ⚠️ Tentando novamente em 1 segundo...")
                    time.sleep(1)
                else:
                    print(f"[Linha {index}] ❌ Todas as tentativas de clique falharam")
                    return False
                
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro durante tentativa de clique: {str(e)}")
                tentativas += 1
                if tentativas < max_tentativas:
                    time.sleep(1)
                else:
                    return False
        
        return False
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao tentar clicar no botão Abrir: {str(e)}")
        return False

def clicar_menu_cobranca(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no menu 'Cobrança'...")
        
        # Tenta diferentes XPaths para encontrar o menu Cobrança
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
        
        # Rola até o elemento
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_cobranca)
        time.sleep(1)
        
        # Tenta diferentes métodos de clique
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                print(f"[Linha {index}] Tentativa {tentativas + 1} de clicar no menu 'Cobrança'...")
                
                # Tenta clicar via JavaScript primeiro
                try:
                    driver.execute_script("arguments[0].click();", menu_cobranca)
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Menu 'Cobrança' clicado via JavaScript")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via JavaScript: {str(e)}")
                
                # Tenta clicar via ActionChains
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(menu_cobranca).pause(1).click().perform()
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Menu 'Cobrança' clicado via ActionChains")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via ActionChains: {str(e)}")
                
                # Tenta clicar normalmente
                try:
                    menu_cobranca.click()
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Menu 'Cobrança' clicado com sucesso")
                    return True
                except ElementClickInterceptedException:
                    print(f"[Linha {index}] ⚠️ Clique interceptado, tentando remover elemento interceptador...")
                    try:
                        # Tenta remover o elemento que está interceptando o clique
                        elemento_interceptador = driver.find_element(By.XPATH, "//div[contains(@class, 'col-offset-start-6')]")
                        driver.execute_script("arguments[0].remove();", elemento_interceptador)
                        time.sleep(1)
                        menu_cobranca.click()
                        time.sleep(2)
                        print(f"[Linha {index}] ✅ Menu 'Cobrança' clicado após remover elemento interceptador")
                        return True
                    except Exception as e:
                        print(f"[Linha {index}] ⚠️ Falha ao remover elemento interceptador: {str(e)}")
                
                tentativas += 1
                if tentativas < max_tentativas:
                    print(f"[Linha {index}] ⚠️ Tentando novamente em 2 segundos...")
                    time.sleep(2)
                else:
                    print(f"[Linha {index}] ❌ Todas as tentativas de clique falharam")
                    return False
                
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro durante tentativa de clique: {str(e)}")
                tentativas += 1
                if tentativas < max_tentativas:
                    time.sleep(2)
                else:
                    return False
        
        return False
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao clicar no menu 'Cobrança': {str(e)}")
        return False

def clicar_botao_registro_chamado(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão de registro de chamado...")
        botao_xpath = '/html/body/div[1]/sc-app/sc-register-ticket-button/div/div/div/button'
        
        # Espera o botão estar presente e clicável
        try:
            botao = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, botao_xpath))
            )
            print(f"[Linha {index}] ✅ Botão de registro de chamado encontrado")
        except TimeoutException:
            print(f"[Linha {index}] ❌ Timeout ao localizar botão de registro de chamado")
            return False
        
        # Rola até o elemento e aguarda um momento
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
        time.sleep(2)
        
        # Tenta diferentes métodos de clique
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                print(f"[Linha {index}] Tentativa {tentativas + 1} de clicar no botão...")
                
                # Tenta clicar via JavaScript primeiro
                try:
                    driver.execute_script("arguments[0].click();", botao)
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Botão clicado via JavaScript")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via JavaScript: {str(e)}")
                
                # Tenta clicar via ActionChains
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(botao).pause(1).click().perform()
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Botão clicado via ActionChains")
                    return True
                except Exception as e:
                    print(f"[Linha {index}] ⚠️ Falha ao clicar via ActionChains: {str(e)}")
                
                # Tenta clicar normalmente
                try:
                    botao.click()
                    time.sleep(2)
                    print(f"[Linha {index}] ✅ Botão clicado com sucesso")
                    return True
                except ElementClickInterceptedException:
                    print(f"[Linha {index}] ⚠️ Clique interceptado, tentando remover elemento interceptador...")
                    try:
                        # Tenta remover o elemento que está interceptando o clique
                        elemento_interceptador = driver.find_element(By.XPATH, "//div[contains(@class, 'col-offset-start-6')]")
                        driver.execute_script("arguments[0].remove();", elemento_interceptador)
                        time.sleep(1)
                        botao.click()
                        time.sleep(2)
                        print(f"[Linha {index}] ✅ Botão clicado após remover elemento interceptador")
                        return True
                    except Exception as e:
                        print(f"[Linha {index}] ⚠️ Falha ao remover elemento interceptador: {str(e)}")
                
                tentativas += 1
                if tentativas < max_tentativas:
                    print(f"[Linha {index}] ⚠️ Tentando novamente em 2 segundos...")
                    time.sleep(2)
                else:
                    print(f"[Linha {index}] ❌ Todas as tentativas de clique falharam")
                    return False
                
            except Exception as e:
                print(f"[Linha {index}] ⚠️ Erro durante tentativa de clique: {str(e)}")
                tentativas += 1
                if tentativas < max_tentativas:
                    time.sleep(2)
                else:
                    return False
        
        return False
        
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao clicar no botão de registro de chamado: {str(e)}")
        return False

def preencher_campos_formulario(driver, actions, row, index, df: pd.DataFrame) -> Optional[str]:
    """Preenche os campos do formulário de registro de chamado."""
    try:
        print(f"[Linha {index}] Preenchendo campos do formulário...")

        # Tipo de atendimento
        print(f"[Linha {index}] Preenchendo Tipo de atendimento...")
        tipo_xpath = '//*[@id="serviceTypeId"]'
        campo_tipo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, tipo_xpath))
        )
        # Preenche o valor diretamente via JavaScript e aciona eventos
        valor_tipo = "Chat Receptivo"
        for tentativa in range(3):
            driver.execute_script(
                """
                arguments[0].value = '';
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """,
                campo_tipo,
                valor_tipo,
            )
            time.sleep(FIELD_DELAY)
            campo_tipo.send_keys(Keys.ENTER)
            if aguardar_campo_valido(driver, campo_tipo, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Tipo de atendimento inválido")
        print(f"[Linha {index}] Tipo de atendimento preenchido: {valor_tipo}")

        # Categoria
        print(f"[Linha {index}] Preenchendo Categoria...")
        categoria_xpath = '//*[@id="categoryId"]'
        campo_categoria = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, categoria_xpath))
        )
        # Preenche o valor diretamente via JavaScript e aciona eventos
        valor_categoria = row['Categoria']
        for tentativa in range(3):
            driver.execute_script(
                """
                arguments[0].value = '';
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """,
                campo_categoria,
                valor_categoria,
            )
            time.sleep(FIELD_DELAY)
            campo_categoria.send_keys(Keys.ENTER)
            if aguardar_campo_valido(driver, campo_categoria, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Categoria inválida")
        print(f"[Linha {index}] Categoria preenchida: {valor_categoria}")

        # Subcategoria
        print(f"[Linha {index}] Preenchendo Subcategoria...")
        subcategoria_xpath = '//*[@id="subCategoryId"]'
        campo_subcategoria = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, subcategoria_xpath))
        )
        # Preenche o valor diretamente via JavaScript e aciona eventos
        valor_subcategoria = "Api Sicoob"
        for tentativa in range(3):
            driver.execute_script(
                """
                arguments[0].value = '';
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """,
                campo_subcategoria,
                valor_subcategoria,
            )
            time.sleep(FIELD_DELAY)
            campo_subcategoria.send_keys(Keys.ENTER)
            if aguardar_campo_valido(driver, campo_subcategoria, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Subcategoria inválida")
        print(f"[Linha {index}] Subcategoria preenchida: {valor_subcategoria}")

        # Aguarda a tela carregar antes de preencher o campo Serviço
        time.sleep(10)

        # Serviço
        print(f"[Linha {index}] Preenchendo Serviço...")
        servico_xpath = '//*[@id="serviceId"]'
        campo_servico = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, servico_xpath))
        )
        valor_servico = normalizar_servico(row['Serviço'])

        # Tenta preencher o campo até que seja validado
        for tentativa in range(3):
            driver.execute_script(
                """
                arguments[0].value = '';
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """,
                campo_servico,
                valor_servico,
            )
            time.sleep(FIELD_DELAY)
            campo_servico.send_keys(Keys.ENTER)
            if aguardar_campo_valido(driver, campo_servico, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Serviço inválido")
        print(f"[Linha {index}] Serviço preenchido: {valor_servico}")

        # Canal de autoatendimento
        print(f"[Linha {index}] Preenchendo Canal de autoatendimento...")
        canal_xpath = '//*[@id="Canal De Autoatendimento"]'
        try:
            select_canal = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, canal_xpath))
            )
            # Preenche o valor diretamente via JavaScript e aciona eventos
            valor_canal = "não se aplica"
            for tentativa in range(3):
                driver.execute_script(
                    """
                    arguments[0].value = '';
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """,
                    select_canal,
                    valor_canal,
                )
                time.sleep(FIELD_DELAY)
                select_canal.send_keys(Keys.ENTER)
                if aguardar_campo_valido(driver, select_canal, index):
                    break
                if tentativa < 2:
                    time.sleep(FIELD_DELAY)
            else:
                raise FormularioError("Canal de autoatendimento inválido")
            print(f"[Linha {index}] Canal de autoatendimento selecionado: {valor_canal}")
        except Exception as e:
            print(f"[Linha {index}] ⚠️ Campo Canal de autoatendimento não encontrado: {str(e)}")
            # Continua mesmo se não encontrar o campo

        # Protocolo PLAD
        print(f"[Linha {index}] Preenchendo Protocolo PLAD...")
        protocolo_xpath = '//*[@id="Protocolo Plad"]'
        campo_protocolo = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, protocolo_xpath))
        )
        # Este campo não parece ter autocomplete complexo, mantendo send_keys
        for tentativa in range(3):
            campo_protocolo.clear()
            campo_protocolo.click()
            campo_protocolo.send_keys(str(row['Protocolo PLAD']))
            if aguardar_campo_valido(driver, campo_protocolo, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Protocolo PLAD inválido")
        print(f"[Linha {index}] Protocolo PLAD preenchido: {row['Protocolo PLAD']}")

        # Descrição
        print(f"[Linha {index}] Preenchendo Descrição...")
        descricao_xpath = '//*[@id="description"]'
        campo_descricao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, descricao_xpath))
        )

        # Mensagem padrão para descrição
        MENSAGEM_PADRAO = "Chamado registrado via automação"

        # Verifica se existe observação válida na coluna G
        observacao = str(row.get('Observação', '')).strip()
        # Define a descrição, tratando casos de nan, valores vazios e tamanho mínimo
        if (pd.isna(row.get('Observação')) or
            observacao.lower() == 'nan' or
            not observacao or
            len(observacao) < 10):
            descricao = MENSAGEM_PADRAO
            if observacao and len(observacao) < 10:
                print(f"[Linha {index}] Observação '{observacao}' tem menos de 10 caracteres. Usando mensagem padrão.")
        else:
            descricao = observacao

        # Limpa o campo e preenche a descrição
        # Este campo não parece ter autocomplete complexo, mantendo send_keys
        for tentativa in range(3):
            campo_descricao.clear()
            campo_descricao.click()
            campo_descricao.send_keys(descricao)
            if aguardar_campo_valido(driver, campo_descricao, index):
                break
            if tentativa < 2:
                time.sleep(FIELD_DELAY)
        else:
            raise FormularioError("Descrição inválida")
        print(f"[Linha {index}] Descrição preenchida: {descricao[:50]}..." if len(descricao) > 50 else f"[Linha {index}] Descrição preenchida: {descricao}")

        # Aguarda o botão Registrar ficar habilitado e clica nele
        print(f"[Linha {index}] Aguardando botão Registrar ficar habilitado...")
        registrar_xpath = '//*[@id="actionbar hide"]/div/div[2]/form/div/div[20]/sc-button/button'
        try:
            botao_registrar = WebDriverWait(driver, 30).until(
                lambda d: d.find_element(By.XPATH, registrar_xpath)
            )
            WebDriverWait(driver, 30).until(
                lambda d: not botao_registrar.get_attribute("disabled")
            )
        except TimeoutException:
            raise FormularioError("Botão Registrar desabilitado")
        botao_registrar.click()
        print(f"[Linha {index}] Botão Registrar clicado")
        time.sleep(2)

        # Aguarda e clica no botão Confirmar
        print(f"[Linha {index}] Aguardando botão Confirmar...")
        confirmar_xpath = '//*[@id="modal"]/div/main/div/div[4]/button'
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        botao_confirmar.click()
        print(f"[Linha {index}] Botão Confirmar clicado")
        time.sleep(2)

        # Captura o número do protocolo
        print(f"[Linha {index}] Capturando número do protocolo...")
        protocolo_xpath = '//*[@id="actionbar hide"]/div/div[2]/form/div/div[2]/sc-card/div/sc-card-content/div/div/div[1]/h5'
        elemento_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        numero_protocolo = elemento_protocolo.text.strip()
        print(f"[Linha {index}] Protocolo capturado: {numero_protocolo}")

        # Salva o protocolo na planilha
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

        # Espera o spinner desaparecer antes de começar
        esperar_spinner_desaparecer(driver, index)

        # Verifica se está na tela correta
        tela_atual = verificar_tela_atual(driver, index)

        if tela_atual == "formulario":
            print(f"[Linha {index}] Já está na tela de formulário")
            return preencher_campos_formulario(driver, actions, row, index, df)

        elif tela_atual == "selecao_conta":
            print(f"[Linha {index}] ⚠️ Está na tela de seleção de conta. Tentando selecionar conta...")
            # Primeiro seleciona a conta
            if not selecionar_conta_por_cooperativa(driver, row['Cooperativa'], index):
                df.at[index, 'Observação'] = "Falha ao selecionar conta"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            # Aguarda um momento para garantir que a conta foi selecionada
            time.sleep(2)

            # Verifica se a conta foi realmente selecionada
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

            # Após confirmar que a conta foi selecionada, clica no menu Cobrança
            if not clicar_menu_cobranca(driver, index):
                df.at[index, 'Observação'] = "Falha ao clicar no menu Cobrança"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            # Aguarda um momento para o menu carregar
            time.sleep(2)

            # Clica no botão de registro de chamado
            if not clicar_botao_registro_chamado(driver, index):
                df.at[index, 'Observação'] = "Falha ao clicar no botão de registro de chamado"
                df.to_excel(EXCEL_PATH, index=False)
                return None

            # Aguarda o campo de categoria do formulário ficar visível/clicável
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
                # Espera o campo estar presente e clicável
                campo_documento = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, campo_documento_xpath))
                )

                # Rola até o elemento
                driver.execute_script("arguments[0].scrollIntoView(true);", campo_documento)
                time.sleep(1)

                # Limpa o campo
                campo_documento.clear()
                time.sleep(0.5)

                # Clica no campo
                campo_documento.click()
                time.sleep(0.5)

                # Obtém o documento e formata
                doc_original = str(row['Documento do cooperado']).strip()
                numeros = ''.join(filter(str.isdigit, doc_original))
                doc_formatado = formatar_documento(numeros)
                print(f"[Linha {index}] Preenchendo documento: {doc_formatado}")

                # Preenche o documento caractere por caractere
                for digito in doc_formatado:
                    campo_documento.send_keys(digito)
                    time.sleep(0.1)

                # Aguarda um momento para garantir que o valor foi preenchido
                time.sleep(1)

                # Verifica se o valor foi preenchido corretamente
                valor_preenchido = campo_documento.get_attribute('value')
                print(f"[Linha {index}] Valor preenchido no campo: {valor_preenchido}")

                if not valor_preenchido:
                    print(f"[Linha {index}] ⚠️ Campo está vazio após preenchimento")
                    # Tenta preencher novamente usando JavaScript
                    driver.execute_script(f"arguments[0].value = '{doc_formatado}';", campo_documento)
                    time.sleep(1)
                    valor_preenchido = campo_documento.get_attribute('value')
                    print(f"[Linha {index}] Valor após tentativa JavaScript: {valor_preenchido}")

                # Verifica se o valor foi preenchido corretamente (ignorando formatação)
                valor_preenchido_numeros = ''.join(filter(str.isdigit, valor_preenchido))
                if valor_preenchido_numeros == numeros:
                    print(f"[Linha {index}] ✅ Documento preenchido com sucesso: {valor_preenchido}")

                    # Tenta clicar no botão consultar
                    if not clicar_botao_consulta(driver, index):
                        print(f"[Linha {index}] ❌ Falha ao clicar no botão consultar")
                        df.at[index, 'Observação'] = "Falha ao clicar no botão consultar"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None

                    # Aguarda um momento para a consulta ser processada
                    time.sleep(2)

                    # Verifica se a pessoa foi encontrada
                    if verificar_pessoa_nao_encontrada(driver, index):
                        print(f"[Linha {index}] ❌ Pessoa não encontrada")
                        df.at[index, 'Observação'] = "Pessoa não encontrada"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None

                    # Tenta clicar no botão Abrir
                    if not clicar_botao_abrir(driver, index):
                        print(f"[Linha {index}] ❌ Falha ao clicar no botão Abrir")
                        df.at[index, 'Observação'] = "Falha ao clicar no botão Abrir"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None

                    # Aguarda um momento para a ação ser processada
                    time.sleep(2)

                    # Verifica se mudou para a tela de seleção de conta
                    tela_atual = verificar_tela_atual(driver, index)
                    if tela_atual == "selecao_conta":
                        print(f"[Linha {index}] ✅ Tela mudou para seleção de conta")
                        # Chama a função novamente para processar a tela de seleção de conta
                        return preencher_formulario(driver, actions, row, index, df)
                    else:
                        print(f"[Linha {index}] ❌ Tela não mudou para seleção de conta após clicar em Abrir")
                        df.at[index, 'Observação'] = "Tela não mudou para seleção de conta após clicar em Abrir"
                        df.to_excel(EXCEL_PATH, index=False)
                        return None
                else:
                    print(f"[Linha {index}] ❌ Documento não preenchido corretamente. Valor esperado: {numeros}, Valor obtido: {valor_preenchido_numeros}")
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
            time.sleep(2)
            esperar_spinner_desaparecer(driver, index)
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
        actions = ActionChains(driver)
        actions.move_to_element(botao_finalizar).click().perform()
        
        logger.info(f"[Linha {index}] Aguardando modal de confirmação...")
        confirmar_xpath = '//*[@id="modal"]/div/main/div/div[4]/button'

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
        logger.critical(f"❌ Sistema encerrado com erro crítico! {str(e)}", exc_info=True)
        raise