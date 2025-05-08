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
EXCEL_PATH = r"C:\Users\gabriel.sandres\OneDrive - Sicoob\Área de Trabalho\cod_fonte_registro\registro-chamados-codigo-fonte\planilha_registro.xlsm"
CHROMEDRIVER_PATH = "chromedriver.exe"
dotenv_path = "login.env"

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

def log_error(error: Exception, context: str, index: Optional[int] = None) -> None:
    """Função auxiliar para logar erros de forma padronizada"""
    error_msg = f"[{'Linha ' + str(index) if index is not None else 'Geral'}] ❌ ERRO em {context}: {str(error)}"
    logger.error(error_msg)
    logger.error("Stack trace:", exc_info=True)
    
    # Adiciona o erro ao DataFrame se houver um índice
    if index is not None and 'df' in globals():
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
        
    except Exception as e:
        log_error(e, "processo de login")
        raise LoginError(f"Falha no login: {str(e)}")


def limpar_e_preencher(campo, valor):
    campo.click()
    time.sleep(0.5)
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    time.sleep(0.3)
    campo.send_keys(valor)


def preencher_com_sugestao(campo, valor):
    campo.click()
    time.sleep(0.5)
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    time.sleep(0.3)
    campo.send_keys(valor[:3])
    time.sleep(1)
    campo.send_keys(Keys.ARROW_DOWN)
    campo.send_keys(Keys.ENTER)


def preencher_com_datalist(campo, valor):
    # Primeiro clique para focar o campo
    campo.click()
    time.sleep(0.5)
    
    # Limpa o campo de forma mais robusta
    campo.clear()
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.DELETE)
    time.sleep(0.5)
    
    # Clica novamente para garantir o foco
    campo.click()
    time.sleep(0.5)
    
    # Envia o valor caractere por caractere
    for char in valor:
        campo.send_keys(char)
        time.sleep(0.1)
    
    time.sleep(0.5)
    campo.send_keys(Keys.TAB)


def preencher_campo_com_js(driver, campo_xpath, valor):
    try:
        print(f"Preenchendo campo com valor: {valor}")
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo_xpath))
        )
        
        # Limpa o campo e define o valor usando JavaScript
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo, valor)
        
        time.sleep(1)
        
        # Simula pressionar Enter
        campo.send_keys(Keys.ENTER)
        time.sleep(0.5)
        
    except Exception as e:
        print(f"Erro ao preencher campo: {e}")
        raise e


def selecionar_opcao(driver, campo_xpath, opcao_xpath):
    try:
        # Clica no campo para focar
        campo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo_xpath))
        )
        campo.click()
        time.sleep(0.5)
        
        # Pega o texto da opção
        opcao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, opcao_xpath))
        )
        valor = opcao.get_attribute("value")
        
        # Digita os primeiros caracteres
        primeiros_chars = valor[:3]
        campo.clear()
        campo.send_keys(primeiros_chars)
        time.sleep(1)
        
        # Pressiona seta para baixo até encontrar a opção correta
        for _ in range(10):  # tenta no máximo 10 vezes
            campo.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.2)
            # Verifica se a opção atual é a desejada
            texto_atual = campo.get_attribute("value")
            if texto_atual and texto_atual.lower() == valor.lower():
                campo.send_keys(Keys.ENTER)
                time.sleep(0.5)
                return
        
        # Se não encontrou, tenta clicar diretamente
        driver.execute_script("arguments[0].click();", opcao)
        time.sleep(0.5)
        
    except Exception as e:
        print(f"Erro ao selecionar opção: {e}")
        # Tenta abordagem alternativa
        try:
            campo.clear()
            campo.send_keys(valor)
            time.sleep(0.3)
            campo.send_keys(Keys.TAB)
        except:
            raise e


def selecionar_opcao_select(driver, select_xpath, valor):
    try:
        print(f"Selecionando opção '{valor}' no select...")
        # Espera o select estar presente e clicável
        select_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, select_xpath))
        )
        
        # Cria um objeto Select
        select = Select(select_element)
        
        # Seleciona pelo valor
        select.select_by_value(valor.lower())
        time.sleep(0.5)
        
    except Exception as e:
        print(f"Erro ao selecionar opção no select: {e}")
        # Tenta abordagem alternativa com JavaScript
        try:
            driver.execute_script("""
                var select = arguments[0];
                var value = arguments[1];
                select.value = value;
                select.dispatchEvent(new Event('change', { bubbles: true }));
            """, select_element, valor.lower())
        except Exception as e2:
            print(f"Erro na abordagem alternativa: {e2}")
            raise e2


def selecionar_conta_por_cooperativa(driver, cooperativa, index):
    try:
        print(f"[Linha {index}] Selecionando conta para cooperativa {cooperativa}...")
        select_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[2]/div[1]/div/form/div/select'
        
        # Espera o select estar presente
        select_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, select_xpath))
        )
        
        # Pega todas as opções
        options = select_element.find_elements(By.TAG_NAME, 'option')
        
        # Procura a opção com a cooperativa correta
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
            
        time.sleep(2)
        return True

    except Exception as e:
        print(f"[Linha {index}] Erro ao selecionar conta: {e}")
        return False


def verificar_pessoa_nao_encontrada(driver, index):
    try:
        # Aguarda um momento para a mensagem aparecer, se existir
        time.sleep(2)
        
        # Tenta encontrar o elemento h6 com a mensagem de "erro"
        erro_xpath = '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[1]/h6'
        
        # Verifica se o elemento existe
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
    # Remove todos os caracteres não numéricos
    numeros = ''.join(filter(str.isdigit, str(documento)))
    
    # Garante que o número tenha 11 dígitos para CPF ou 14 para CNPJ
    if len(numeros) <= 11:  # CPF
        numeros = numeros.zfill(11)  # Preenche com zeros à esquerda se necessário
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    elif len(numeros) <= 14:  # CNPJ
        numeros = numeros.zfill(14)  # Preenche com zeros à esquerda se necessário
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    else:
        return documento  # Retorna como está se não for CPF nem CNPJ


def esperar_spinner_desaparecer(driver, timeout=30):
    try:
        # Espera o spinner desaparecer
        spinner_xpath = "//div[contains(@class, 'ngx-spinner-overlay')]"
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.XPATH, spinner_xpath))
        )
        return True
    except Exception as e:
        print(f"Erro ao esperar spinner desaparecer: {e}")
        return False


def clicar_botao_consulta(driver, index):
    try:
        print(f"[Linha {index}] Tentando clicar no botão consultar...")
        
        # XPath do botão
        botao_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'
        
        # Espera o botão estar presente e visível
        botao = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, botao_xpath))
        )
        
        # Tenta diferentes abordagens para clicar no botão
        tentativas = 0
        max_tentativas = 3
        while tentativas < max_tentativas:
            try:
                # Rola até o botão
                driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                time.sleep(1)
                
                # Tenta clicar normalmente
                botao.click()
                print(f"[Linha {index}] Botão consultar clicado com sucesso")
                return True
            except:
                try:
                    # Tenta clicar com JavaScript
                    driver.execute_script("arguments[0].click();", botao)
                    print(f"[Linha {index}] Botão consultar clicado via JavaScript")
                    return True
                except:
                    try:
                        # Tenta com Actions
                        actions = ActionChains(driver)
                        actions.move_to_element(botao).click().perform()
                        print(f"[Linha {index}] Botão consultar clicado via Actions")
                        return True
                    except:
                        tentativas += 1
                        if tentativas < max_tentativas:
                            print(f"[Linha {index}] Tentativa {tentativas} falhou, tentando novamente...")
                            time.sleep(2)
                        else:
                            print(f"[Linha {index}] ❌ Não foi possível clicar no botão após {max_tentativas} tentativas")
                            return False
        return False
    except Exception as e:
        print(f"[Linha {index}] ❌ Erro ao tentar clicar no botão consultar: {str(e)}")
        return False


def preencher_formulario(driver, actions, row, index):
    try:
        logger.info(f"\n[Linha {index}] Iniciando preenchimento do formulário...")
        # Pega o documento original e formatado
        doc_original = str(row['Documento do cooperado']).strip()
        doc_formatado = formatar_documento(doc_original)
        logger.info(f"[Linha {index}] Documento original: {doc_original}")
        logger.info(f"[Linha {index}] Documento formatado: {doc_formatado}")
        
        protocolo_plad = str(row['Protocolo PLAD']).strip()
        categoria = str(row['Categoria']).strip()
        servico = normalizar_servico(str(row['Serviço']).strip())
        cooperativa = str(row['Cooperativa']).strip()
        
        # Mensagem padrão para descrição
        MENSAGEM_PADRAO = "Registro de atendimento realizado na Plataforma de Atendimento Digital via automação"
        
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

        print(f"[Linha {index}] Aguardando campo de documento...")
        campo_documento_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[2]/sc-form-field/div/input'
        
        # Espera o campo estar presente e interativo
        campo_documento = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, campo_documento_xpath))
        )
        print(f"[Linha {index}] Campo de documento encontrado")

        # Limpa o campo e aguarda um momento
        campo_documento.clear()
        time.sleep(0.5)
        
        # Simula digitação humana do documento não formatado
        numeros = ''.join(filter(str.isdigit, doc_original))
        for digito in numeros:
            campo_documento.send_keys(digito)
            time.sleep(0.1)
        
        # Pressiona Tab para permitir a formatação automática
        campo_documento.send_keys(Keys.TAB)
        time.sleep(1)

        print(f"[Linha {index}] Documento preenchido: {doc_formatado}")

        print(f"[Linha {index}] Aguardando botão de consulta...")
        botao_xpath = '/html/body/div/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'
        
        # Espera o botão estar presente e visível
        botao = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, botao_xpath))
        )
        
        # Rola até o botão e clica
        driver.execute_script("arguments[0].scrollIntoView(true);", botao)
        time.sleep(1)
        botao.click()
        print(f"[Linha {index}] Botão de consulta clicado")
        time.sleep(2)  # Aguarda um pouco após clicar

        # Verifica se a pessoa não foi encontrada
        if verificar_pessoa_nao_encontrada(driver, index):
            # Atualiza a observação na planilha
            df.at[index, 'Observação'] = "Pessoa não identificada como cooperada!"
            df.to_excel(EXCEL_PATH, index=False)
            print(f"[Linha {index}] ℹ️ Observação atualizada na planilha")
            return None

        print(f"[Linha {index}] Aguardando botão de seleção de conta...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/sc-app/sc-template/sc-root/main/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[2]/sc-button/button'))
        ).click()
        print(f"[Linha {index}] Botão de seleção de conta clicado")
        time.sleep(2)

        # Seleciona a conta com base na cooperativa
        if not selecionar_conta_por_cooperativa(driver, cooperativa, index):
            print(f"[Linha {index}] Não foi possível continuar sem a conta correta")
            return None

        time.sleep(2)

        print(f"[Linha {index}] Aguardando botão de categoria...")
        xpath_categoria = '/html/body/div[1]/sc-app/sc-template/sc-root/main/aside/sc-sidebar-container/aside/sc-sidebar/div[4]/div[1]/sc-card/div/div/div/div'
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, xpath_categoria))).click()
        print(f"[Linha {index}] Botão de categoria clicado")
        time.sleep(1)

        print(f"[Linha {index}] Aguardando botão de registro de chamado...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/sc-app/sc-register-ticket-button/div/div/div/button'))
        ).click()
        print(f"[Linha {index}] Botão de registro de chamado clicado")
        time.sleep(2)

        # Campos do formulário
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

        # Preenchendo os campos na ordem
        for campo_nome, campo_info in campos.items():
            print(f"[Linha {index}] Preenchendo {campo_nome}...")
            preencher_campo_com_js(driver, campo_info['xpath'], campo_info['valor'])
            print(f"[Linha {index}] {campo_nome} preenchido com: {campo_info['valor']}")
            time.sleep(1)

        # Canal de autoatendimento
        print(f"[Linha {index}] Preenchendo Canal de autoatendimento...")
        canal_autoatendimento_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[7]/sc-additional-category-data/form/div/div[2]/sc-form-field/div/select'
        selecionar_opcao_select(driver, canal_autoatendimento_xpath, "não se aplica")
        print(f"[Linha {index}] Canal de autoatendimento selecionado")
        time.sleep(1)

        # Protocolo
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
        time.sleep(1)

        # Descrição
        print(f"[Linha {index}] Preenchendo Descrição...")
        descricao_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[9]/sc-form-field/div/textarea'
        campo_descricao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, descricao_xpath))
        )
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo_descricao, descricao)
        print(f"[Linha {index}] Descrição preenchida: {descricao[:50]}..." if len(descricao) > 50 else f"[Linha {index}] Descrição preenchida: {descricao}")
        time.sleep(1)

        # Aguarda o botão Registrar ficar habilitado e clica nele
        print(f"[Linha {index}] Aguardando botão Registrar ficar habilitado...")
        registrar_xpath = '/html/body/div[3]/div[2]/div/sc-register-ticket/sc-actionbar/div/div/div[2]/form/div/div[20]/sc-button/button'
        # Espera até o botão ficar clicável (não estar disabled)
        WebDriverWait(driver, 30).until(
            lambda d: d.find_element(By.XPATH, registrar_xpath).is_enabled()
        )
        botao_registrar = driver.find_element(By.XPATH, registrar_xpath)
        botao_registrar.click()
        print(f"[Linha {index}] Botão Registrar clicado")
        time.sleep(2)

        # Aguarda e clica no botão Confirmar
        print(f"[Linha {index}] Aguardando botão Confirmar...")
        confirmar_xpath = '/html/body/div[3]/div[4]/div/sc-register-ticket-modal/sc-modal/div/div/sc-modal-footer/div/div/div[2]/sc-button/button'
        botao_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        )
        botao_confirmar.click()
        print(f"[Linha {index}] Botão Confirmar clicado")
        time.sleep(2)

        # Captura o número do protocolo
        print(f"[Linha {index}] Capturando número do protocolo...")
        protocolo_xpath = '/html/body/div[3]/div[4]/div/sc-view-ticket-data/sc-actionbar/div/div/div[2]/form/div/div[2]/sc-card/div/sc-card-content/div/div/div[1]/h5'
        elemento_protocolo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, protocolo_xpath))
        )
        # Extrai o texto e remove espaços em branco
        numero_protocolo = elemento_protocolo.text.strip()
        print(f"[Linha {index}] Protocolo capturado: {numero_protocolo}")

        return numero_protocolo

    except Exception as e:
        log_error(e, "preencher formulário", index)
        return None


def tentar_preencher_formulario(driver, actions, row, index, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            if tentativa > 0:
                print(f"[Linha {index}] 🔄 Tentativa {tentativa + 1} de {max_tentativas}")
                # Recarrega a página para tentar novamente
                driver.refresh()
                time.sleep(5)  # Aguarda a página recarregar
            
            return preencher_formulario(driver, actions, row, index)
            
        except Exception as e:
            print(f"[Linha {index}] ❌ Erro na tentativa {tentativa + 1}:")
            print(str(e))
            if tentativa == max_tentativas - 1:  # Se for a última tentativa
                print(f"[Linha {index}] ❌ Todas as tentativas falharam")
                # Atualiza a observação na planilha
                df.at[index, 'Observação'] = f"Erro após {max_tentativas} tentativas: {str(e)}"
                df.to_excel(EXCEL_PATH, index=False)
                return None
            time.sleep(3)  # Aguarda um pouco antes da próxima tentativa
    return None


def finalizar_atendimento(driver, index):
    try:
        logger.info(f"[Linha {index}] 🔄 Iniciando finalização do atendimento...")
        
        # Clica no botão "Finalizar atendimento"
        logger.info(f"[Linha {index}] Clicando no botão 'Finalizar atendimento'...")
        finalizar_xpath = '/html/body/div[3]/div[4]/div/sc-view-ticket-data/sc-actionbar/div/div/div[2]/form/div/div[5]/sc-button/button'
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, finalizar_xpath))
        ).click()
        
        # Aguarda e clica no botão de confirmação
        logger.info(f"[Linha {index}] Confirmando finalização...")
        confirmar_xpath = '/html/body/div[3]/div[2]/div/sc-end-service-modal/sc-modal/div/div/main/div/div[4]/button'
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, confirmar_xpath))
        ).click()
        
        # Aguarda a tela inicial carregar
        logger.info(f"[Linha {index}] Aguardando retorno à tela inicial...")
        time.sleep(3)
        
        logger.info(f"[Linha {index}] ✅ Atendimento finalizado com sucesso!")
        return True
        
    except Exception as e:
        log_error(e, "finalização do atendimento", index)
        raise FinalizacaoError(f"Falha ao finalizar atendimento: {str(e)}")


def main():
    try:
        logger.info("🚀 Iniciando sistema de registro de chamados...")
        
        # Carrega as credenciais
        logger.info("Carregando credenciais...")
        username, password = load_credentials()
        
        # Configura o diretório de download
        download_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Inicializa o driver
        logger.info("Inicializando navegador...")
        driver = setup_driver(download_dir)
        actions = ActionChains(driver)
        
        try:
            # Faz login
            login(driver, username, password)
            
            # Carrega os dados do Excel
            logger.info("Carregando dados da planilha...")
            df = load_excel_data(EXCEL_PATH)
            total_registros = len(df)
            logger.info(f"📊 Total de registros a processar: {total_registros}")
            
            # Processa cada linha do Excel
            registros_processados = 0
            registros_com_erro = 0
            
            for index, row in df.iterrows():
                try:
                    logger.info(f"\n{'='*50}")
                    logger.info(f"[Linha {index}] 📝 Iniciando processamento do registro {index + 1}/{total_registros}")
                    
                    # Tenta preencher o formulário
                    if tentar_preencher_formulario(driver, actions, row, index):
                        # Se o preenchimento foi bem sucedido, finaliza o atendimento
                        if finalizar_atendimento(driver, index):
                            registros_processados += 1
                            logger.info(f"[Linha {index}] ✅ Registro processado com sucesso!")
                        else:
                            registros_com_erro += 1
                            logger.error(f"[Linha {index}] ❌ Erro ao finalizar atendimento")
                    else:
                        registros_com_erro += 1
                        logger.error(f"[Linha {index}] ❌ Erro ao preencher formulário")
                    
                    # Aguarda um momento antes de processar o próximo registro
                    time.sleep(2)
                    
                except Exception as e:
                    registros_com_erro += 1
                    log_error(e, "processamento do registro", index)
                    continue
            
            # Relatório final
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
