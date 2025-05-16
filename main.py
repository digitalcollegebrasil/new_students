import os
import time
import shutil
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import traceback

load_dotenv()

head_office = os.getenv("HEAD_OFFICE")
email_address = os.getenv("SPONTE_EMAIL")
password_value = os.getenv("SPONTE_PASSWORD")

current_dir = os.path.dirname(os.path.abspath(__file__))
download_dir = current_dir
base_target_dir = os.path.join(current_dir, 'target')

if not os.path.exists(base_target_dir):
    os.makedirs(base_target_dir)

def remove_value_attribute(driver, element):
    driver.execute_script("arguments[0].removeAttribute('value')", element)

def set_input_value(driver, element, value):
    driver.execute_script("arguments[0].value = arguments[1]", element, value)

def move_downloaded_file(download_dir, target_dir, start_date_range):
    filename = f"Relatorio_{start_date_range.strftime('%d_%m_%Y')}.xls"
    target_path = os.path.join(target_dir, filename)
    downloaded_files = [f for f in os.listdir(download_dir) if f.endswith('.xls')]

    if downloaded_files:
        latest_file = max(downloaded_files, key=lambda f: os.path.getctime(os.path.join(download_dir, f)))
        shutil.move(os.path.join(download_dir, latest_file), target_path)
        print(f"Movido XLS para {start_date_range.strftime('%d/%m/%Y')} em {target_path}")
        return target_path
    else:
        print(f"Nenhum arquivo XLS encontrado para mover para {start_date_range.strftime('%d/%m/%Y')}.")
        return None

def click_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView();", element)
    driver.execute_script("arguments[0].click();", element)

def select_turma_by_name(driver, turma_name):
    turma_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "select2-ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder2_tab_tabFiltro_cmbTurma-container"))
    )
    turma_dropdown.click()
    time.sleep(1)

    try:
        turma_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{turma_name}')]"))
        )
        turma_option.click()
        print(f"Turma {turma_name} selecionado.")
    except TimeoutException:
        print(f"Turma {turma_name} não encontrado no dropdown.")
        turma_dropdown.click()
        return False
    
    time.sleep(2)
    return True

def carregar_turmas(caminho_arquivo):
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as file:
            turmas = [linha.strip() for linha in file if linha.strip()]
        print(f"{len(turmas)} turmas carregadas a partir de {caminho_arquivo}.")
        return turmas
    except FileNotFoundError:
        print(f"Arquivo {caminho_arquivo} não encontrado.")
        return []
    except Exception as e:
        print(f"Erro ao ler o arquivo {caminho_arquivo}: {str(e)}")
        return []

turmas = carregar_turmas(os.path.join(current_dir, 'turmas.txt'))

start_date_range = datetime.strptime("01/01/2025", "%d/%m/%Y")
end_date_range = datetime.strptime("31/01/2025", "%d/%m/%Y")

chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www.sponteeducacional.net.br/Home.aspx")

def clicar_checkbox(driver, checkbox_id):
    checkbox = driver.find_element(By.ID, checkbox_id)
    print(checkbox_id)

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, checkbox_id))
    )

    action = ActionChains(driver)
    action.move_to_element(checkbox).click().perform()

    time.sleep(3)

try:
    email = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "txtLogin"))
    )
    email.send_keys(email_address)
    print("Email inserido.")

    password = driver.find_element(By.ID, "txtSenha")
    password.send_keys(password_value)
    print("Senha inserida.")

    login_button = driver.find_element(By.ID, "btnok")
    login_button.click()
    print("Botão de login clicado.")
    time.sleep(5)

    enterprise = driver.find_element(By.ID, "ctl00_spnNomeEmpresa").get_attribute("innerText").strip().replace(" ", "")
    print(enterprise)

    combinacoes = {
        ("Aldeota", "DIGITALCOLLEGESUL-74070"): (1, "Acessando a sede Aldeota."),
        ("Aldeota", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (1, "Acessando a sede Aldeota."),
        ("Sul", "DIGITALCOLLEGEALDEOTA-72546"): (3, "Acessando a sede Sul."),
        ("Sul", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (3, "Acessando a sede Sul."),
        ("Bezerra", "DIGITALCOLLEGEALDEOTA-72546"): (4, "Acessando a sede Bezerra."),
        ("Bezerra", "DIGITALCOLLEGESUL-74070"): (4, "Acessando a sede Bezerra."),
        ("Aldeota", "DIGITALCOLLEGEALDEOTA-72546"): (None, "O script já está na Aldeota."),
        ("Sul", "DIGITALCOLLEGESUL-74070"): (None, "O script já está no Sul."),
        ("Bezerra", "DIGITALCOLLEGEBEZERRADEMENEZES-488365"): (None, "O script já está na Bezerra."),
    }

    resultado = combinacoes.get((head_office, enterprise), (None, "Ação não realizada: combinação não reconhecida."))
    val, message = resultado

    print(message)

    if val is not None:
        driver.execute_script(f"$('#ctl00_hdnEmpresa').val({val});javascript:__doPostBack('ctl00$lnkChange','');")
        time.sleep(3)

    driver.get("https://www.sponteeducacional.net.br/SPRel/Alunos/DadosCadastro.aspx")
    time.sleep(5)

    combined_data = []

    primeira_turma = True

    for turma in turmas:
        try:
            print(f"\nProcessando turma: {turma}")

            if not select_turma_by_name(driver, turma):
                print(f"Pular processamento para a turma {turma}, pois ele não foi encontrado.")
                continue

            if primeira_turma:
                export_checkbox = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ContentPlaceHolder1_chkExportar"))
                )
                click_element(driver, export_checkbox)
                print("Checkbox de exportação clicado.")
                time.sleep(1)

                select2_span = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='select2-ctl00_ctl00_ContentPlaceHolder1_cmbTipoExportacao-container']"))
                )

                select2_span.click()
                print("Dropdown de formato de exportação clicado.")
                time.sleep(1)

                option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Excel Sem Formatação')]"))
                )
                option.click()
                print("Opção 'Excel Sem Formatação' selecionada.")
                primeira_turma = False

            emit_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ContentPlaceHolder1_btnGerar_div"))
            )
            click_element(driver, emit_button)
            print("Botão 'Emitir' clicado. Relatório emitido, aguardando download...")

            time.sleep(20)

            xls_path = move_downloaded_file(download_dir, base_target_dir, start_date_range)
            if xls_path:
                df = pd.read_excel(xls_path, skiprows=3)
                print(f"Dados capturados para {turma}:\n{df}\n")

                combined_data.append(df)
                print(f"Relatório para {turma} processado com sucesso.")
            else:
                print(f"Nenhum arquivo encontrado para a turma {turma}.")

        except Exception as e:
                print(f"Erro ao processar turma {turma}: {str(e)}")
                traceback.print_exc()

    if combined_data:
        combined_df = pd.concat(combined_data, ignore_index=True)
        output_path = f"{base_target_dir}/Relatorio_Geral_{head_office}.xlsx"

        combined_df.to_excel(output_path, index=False)
        print(f"Arquivo Excel consolidado salvo com sucesso em {output_path}.")
    else:
        print("Nenhum dado para consolidar.")

except Exception as e:
    print(f"Ocorreu um erro durante o processo: {str(e)}")
    traceback.print_exc()

finally:
    driver.quit()

script_dir = os.path.dirname(__file__)
path = os.path.join(script_dir, f'{base_target_dir}/Relatorio_Geral_{head_office}.xlsx')

print(f"Usando o caminho: {path}")

dias_semana_map = {
    "Monday": "Segunda-feira",
    "Tuesday": "Terça-feira",
    "Wednesday": "Quarta-feira",
    "Thursday": "Quinta-feira",
    "Friday": "Sexta-feira",
    "Saturday": "Sábado",
    "Sunday": "Domingo"
}

def extrair_data(data_val):
    try:
        if pd.notnull(data_val):
            return datetime.strptime(str(data_val), '%d/%m/%Y %H:%M:%S').date()
        return None
    except ValueError:
        return None

if not os.path.exists(path):
    print("Arquivo não encontrado no caminho especificado.")
else:
    try:
        df = pd.read_excel(path, usecols="A:H")

        print(f"Colunas disponíveis no DataFrame: {df.columns.tolist()}")

        print(f"Conteúdo do DataFrame antes da extração de datas: \n{df.head()}")

        if 'Data' in df.columns:
            df['Data'] = df['Data'].apply(extrair_data)

            df['Dia da Semana'] = df['Data'].apply(lambda x: dias_semana_map[x.strftime('%A')] if pd.notnull(x) else 'Data Inválida')

            print(f"Conteúdo do dataframe após a extração de datas: \n{df.head()}")

            output_path = os.path.join(script_dir, 'arquivo_personalizado.xlsx')

            df.to_excel(output_path, index=False)
            print(f"Arquivo personalizado salvo com sucesso em: {output_path}")
        else:
            print("Coluna 'Data' não encontrada no DataFrame.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
