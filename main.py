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
    # Cria o diretório de logs se não existir
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # Configura o nome do arquivo de log com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"registro_chamados_{timestamp}.log")
    
    # Configura o logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

# Inicializa o logger
logger = setup_logging()

# === CONFIGURAÇÕES GERAIS ===
BASE_URL = "https://portal.sisbr.coop.br/visao360/consult"
# Ponto 7: Caminhos relativos configuráveis via variáveis de ambiente
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(os.path.dirname(__file__), "planilha_registro.xlsm"))
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH", os.path.join(os.path.dirname(__file__), "chromedriver.exe"))
dotenv_path = os.path.join(os.path.dirname(__file__), "login.env")

# Dicionário de mapeamento para o campo 'Serviço' com variações comuns
SERVICOS_VALIDOS = {
    # Dúvida Negocial
    "dúvida negocial": "Dúvida Negocial",
    "duvida negocial": "Dúvida Negocial",
    "duvida negociacao": "Dúvida Negocial",
    "dúvida negociacao": "Dúvida Negocial",
    "duvida de negocio": "Dúvida Negocial",
    "duvida negocio": "Dúvida Negocial",
    # Dúvida Técnica
    "dúvida técnica": "Dúvida Técnica",
    "duvida tecnica": "Dúvida Técnica",
    "duvida tecnica": "Dúvida Técnica",
    "duvida de tecnica": "Dúvida Técnica",
    # Ambiente de testes
    "ambiente de testes": "Ambiente de testes",
    "ambiente testes": "Ambiente de testes",
    "ambiente de teste": "Ambiente de testes",
    "ambiente teste": "Ambiente de testes",
    # Erro De Documentação
    "erro de documentação": "Erro De Documentação",
    "erro de documentacao": "Erro De Documentação",
    "erro documentacao": "Erro De Documentação",
    "erro documentação": "Erro De Documentação",
    # Integração Imcompleta
    "integração imcompleta": "Integração Imcompleta",
    "integracao imcompleta": "Integração Imcompleta",
    "integracao incompleta": "Integração Imcompleta",
    "integração incompleta": "Integração Imcompleta",
    # Sugestão De Melhoria
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
    """Classe base para exceções específicas do sistema de registro de chamados"""
    pass

class LoginError(RegistroChamadoError):
    """Erro durante o processo de login"""
    pass

class FormularioError(RegistroChamadoError):
    """Erro durante o preenchimento do formulário"""
    pass

class FinalizacaoError(RegistroChamadoError):
    """Erro durante a finalização do atendimento"""
    pass

# Ponto 2: Passar df como parâmetro em vez de usar globalmente
def log_error(error: Exception, context: str, index: Optional[int] = None, df: Optional[pd.DataFrame] = None) -> None:
    """Função auxiliar para logar erros de forma padronizada"""
    error_msg = f"[{'Linha ' + str(index) if index is not None else 'Geral'}] ❌ ERRO em {context}: {str(error)}"
    logger.error(error_msg)
    logger.error("Stack trace:", exc_info=True)
    
    # Adiciona o erro ao DataFrame se houver um índice e df for fornecido
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

# Ponto 9: Verificar existência do arquivo .env
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
    # Lê o Excel especificando que a coluna 'Documento do cooperado' deve ser tratada como texto
    df = pd.read_excel(
        file_path,
        dtype={'Documento do cooperado': str}  # Força a coluna a ser lida como texto
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
    # Ponto 4: Remover time.sleep
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    campo.send_keys(valor)

def preencher_com_sugestao(campo, valor):
    campo.click()
    # Ponto 4: Remover time.sleep
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    campo.send_keys(valor[:3])
    # Ponto 4: Substituir time.sleep por espera explícita
    WebDriverWait(campo.parent, 10).until(EC.presence_of_element_located((By.XPATH, f"//option[contains(text(), '{valor}')]")))
    campo.send_keys(Keys.ARROW_DOWN)
    campo.send_keys(Keys.ENTER)

def preencher_com_datalist(campo, valor):
    campo.click()
    # Ponto 4: Remover time.sleep
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
        
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo, valor)
        
        campo.send_keys(Keys.ENTER)
        
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
        # Ponto 4: Substituir time.sleep
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
        select.select_by_value(valor.lower())
        
    except TimeoutException as e:
        print(f"Timeout ao localizar select: {e}")
        raise FormularioError(f"Timeout ao localizar select: {str(e)}")
    except NoSuchElementException as e:
        print(f"Select não encontrado: {e}")
        raise FormularioError(f"Select não encontrado: {str(e)}")
    except Exception as e:
        print(f"Erro ao selecionar opção no select: {e}")
        try:
            driver.execute_script("""
                var select = arguments[0];
                var value = arguments[1];
                select.value = value;
                select.dispatchEvent(new Event('change', { bubbles: true }));
            """, select_element, valor.lower())
        except Exception as e2:
            print(f"Erro na abordagem alternativa: {e2}")
            raise FormularioError(f"Erro ao selecionar opção no select: {str(e2)}")

def selecionar_conta_por_cooperativa(driver, cooperativa, index):
    try:
        print(f"[Linha {index}] Selecionando conta para cooperativa {cooperativa}...")
        select_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
        
        select_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, select_xpath))
        )
        
        options = select_element.find_elements(By.TAG_NAME, 'option')
        
        conta_encontrada = False
        for option in options:
            texto_opcao = option.text.strip()
            if texto_opcao.startswith(f"Coop: {cooperativa}"):
                print(f"[Linha {index}] Conta encontrada: {texto_opcao}")
                option.click()
                conta_encontrada = True
                break
        
        if not conta_encontrada:
            print(f"[Linha {index}] ⚠️ ATENÇÃO: Nenhuma conta encontrada para cooperativa {cooperativa}")
            return False
            
        return True

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

# Ponto 8: Adicionar validação para documentos
def formatar_documento(documento):
    numeros = ''.join(filter(str.isdigit, str(documento)))
    if len(numeros) == 11:  # CPF
        numeros = numeros.zfill(11)
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    elif len(numeros) == 14:  # CNPJ
        numeros = numeros.zfill(14)
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    else:
        logger.warning(f"Documento inválido: {documento}")
        return documento

# Ponto 6: Função para esperar modal desaparecer
def esperar_modal_desaparecer(driver, index, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "modal"))
        )
        return True
    except TimeoutException:
        logger.warning(f"[Linha {index}] Modal ainda presente após {timeout} segundos")
        return False

def esperar_spinner_desaparecer(driver, timeout=30):
    try:
        spinner_xpath = "//div[contains(@class, 'ngx-spinner-overlay')]"
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.XPATH, spinner_xpath))
        )
        return True
    except TimeoutException:
        print(f"Timeout ao esperar spinner desaparecer")
        return False
    except Exception as e:
        print(f"Erro ao esperar spinner desaparecer: {e}")
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
                botao.click()
                print(f"[Linha {index}] Botão consultar clicado com sucesso")
                return True
            except ElementClickInterceptedException:
                try:
                    driver.execute_script("arguments[0].click();", botao)
                    print(f"[Linha {index}] Botão consultar clicado via JavaScript")
                    return True
                except:
                    try:
                        actions = ActionChains(driver)
                        actions.move_to_element(botao).click().perform()
                        print(f"[Linha {index}] Botão consultar clicado via Actions")
                        return True
                    except:
                        tentativas += 1
                        if tentativas < max_tentativas:
                            print(f"[Linha {index}] Tentativa {tentativas} falhou, tentando novamente...")
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

# Ponto 2 e 5: Passar df como parâmetro e adicionar validação de dados
def preencher_formulario(driver, actions, row, index, df: pd.DataFrame):
    try:
        logger.info(f"\n[Linha {index}] Iniciando preenchimento do formulário...")
        # Ponto 5: Validar dados da linha
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
        
        campo_documento = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
        )
        print(f"[Linha {index}] Campo de documento encontrado")

        campo_documento.clear()
        numeros = ''.join(filter(str.isdigit, doc_original))
        for digito in numeros:
            campo_documento.send_keys(digito)
            # Ponto 4: Remover time.sleep
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
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath_categoria))).click()
        print(f"[Linha {index}] Botão de categoria clicado")

        print(f"[Linha {index}] Aguardando botão de registro de chamado...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/sc-app/sc-register-ticket-button/div/div/div/button'))
        ).click()
        print(f"[Linha {index}] Botão de registro de chamado clicado")

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
        canal_autoatendimento_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[7]/sc-additional-category-data/form/div/div[2]/sc-form-field/div/select'
        selecionar_opcao_select(driver, canal_autoatendimento_xpath, "não se aplica")
        print(f"[Linha {index}] Canal de autoatendimento selecionado")

        print(f"[Linha {index}] Preenchendo Protocolo...")
        protocolo_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[8]/sc-additional-service-data/form/div/div[2]/sc-form-field/div/input'
        campo_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo_protocolo, protocolo_plad)
        print(f"[Linha {index}] Protocolo preenchido: {protocolo_plad}")

        print(f"[Linha {index}] Preenchendo Descrição...")
        descricao_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[9]/sc-form-field/div/textarea'
        try:
            campo_descricao = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, descricao_xpath))
            )
            
            driver.execute_script("arguments[0].scrollIntoView(true);", campo_descricao)
            
            try:
                driver.execute_script("""
                    arguments[0].value = '';
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, campo_descricao, descricao)
                
                valor_preenchido = driver.execute_script("return arguments[0].value;", campo_descricao)
                if not valor_preenchido:
                    campo_descricao.clear()
                    campo_descricao.send_keys(descricao)
                    valor_preenchido = campo_descricao.get_attribute('value')
                    if not valor_preenchido:
                        actions = ActionChains(driver)
                        actions.move_to_element(campo_descricao).click().perform()
                        actions.send_keys(descricao).perform()
                
                print(f"[Linha {index}] Descrição preenchida: {descricao[:50]}..." if len(descricao) > 50 else f"[Linha {index}] Descrição preenchida: {descricao}")
                
            except Exception as e:
                print(f"[Linha {index}] Erro ao preencher descrição: {str(e)}")
                raise
                
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
        botao_registrar.click()
        print(f"[Linha {index}] Botão Registrar clicado")

        print(f"[Linha {index}] Aguardando botão Confirmar...")
        confirmar_xpath = '/html/body/div[3]/div[4]/div/sc-register-ticket-modal/sc-modal/div/div/sc-modal-footer/div/div/div[2]/sc-button/button'
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        botao_confirmar.click()
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

# Ponto 2: Passar df como parâmetro
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

# Ponto 6: Usar função esperar_modal_desaparecer
def finalizar_atendimento(driver, index, df: pd.DataFrame):
    try:
        logger.info(f"[Linha {index}] 🔄 Iniciando finalização do atendimento...")
        
        if not esperar_modal_desaparecer(driver, index):
            raise FinalizacaoError("Modal não desapareceu a tempo")
        
        logger.info(f"[Linha {index}] Clicando no botão 'Finalizar atendimento'...")
        finalizar_xpath = '/html/body/div[3]/div[4]/div/sc-view-ticket-data/sc-actionbar/div/div/div[2]/form/div/div[5]/sc-button/button'
        
        try:
            botao_finalizar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, finalizar_xpath))
            )
            botao_finalizar.click()
        except ElementClickInterceptedException:
            try:
                driver.execute_script("arguments[0].click();", botao_finalizar)
            except:
                actions = ActionChains(driver)
                actions.move_to_element(botao_finalizar).click().perform()
        
        logger.info(f"[Linha {index}] Aguardando modal de confirmação...")
        confirmar_xpath = '/html/body/div[3]/div[2]/div/sc-end-service-modal/sc-modal/div/div/main/div/div[4]/button'
        
        try:
            botao_confirmar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
            )
            botao_confirmar.click()
        except ElementClickInterceptedException:
            try:
                driver.execute_script("arguments[0].click();", botao_confirmar)
            except:
                actions = ActionChains(driver)
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
            # Ponto 5: Validar colunas do Excel
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
            driver.quit()
            
    except Exception as e:
        log_error(e, "execução geral do sistema")
        if 'driver' in locals():
            driver.quit()
        raise

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical("❌ Sistema encerrado com erro crítico!", exc_info=True)
        raise