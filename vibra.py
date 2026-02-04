import time
import re
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURAÇÕES ---
data_hoje = datetime.now().strftime("%d-%m") 
# Usando o ID da planilha oficial para garantir a gravação
ID_PLANILHA_OFICIAL = "1bUvHLoUAmT4vcFy7J_qN5SendCl72Ih7ZLnF6UGd8VI" 
NOME_ABA_HOJE = f"Preço {data_hoje}"
ABA_MODELO = "Preço 20-01" 
ARQUIVO_JSON_GOOGLE = "dados-google.json"

BASES_VIBRA = {
    "TERESINA": {
        "usuario": "6074", "senha": "DERIVADOS02", 
        "celula_s10": "E13", "celula_s500": "I13"
    },
    "ACAILANDIA": {
        "usuario": "1772103", "senha": "DERIVADOS02", 
        "celula_s10_1": "E30", "celula_s500_1": "I30",
        "celula_s10_2": "E55", "celula_s500_2": "I55"
    },
    "PORTO_NACIONAL": {
        "usuario": "1775785", "senha": "DERIVADOS02", 
        "celula_s10": "E66", "celula_s500": "I66"
    },
    "LUIS_EDUARDO": { 
        "usuario": "1845988", "senha": "DERIVADOS02",
        "celula_s10_1": "E120", "celula_s500_1": "I120",
        "celula_s10_2": "E106", "celula_s500_2": "I106",
        "celula_s10_3": "E190"
    },
    "BELEM": {
        "usuario": "1846285", "senha": "DERIVADOS02",
        "celula_s10_1": "E212", "celula_s500_1": "I212",
        "celula_s10_2": "E227", "celula_s500_2": "I227"
    }
}

def aceitar_todos_cookies_vibra(driver):
    """Limpa informativos, banners de cookies e modais da Vibra."""
    # Termos comuns em botões de fechar/aceitar
    termos = ["Aceitar", "Entendi", "OK", "Fechar", "Prosseguir", "Concordo"]
    
    # 1. Tenta fechar o botão específico de informativo da Vibra
    try:
        elementos_f = driver.find_elements(By.CSS_SELECTOR, ".btn-fecha-informativo, .close, [data-dismiss='modal']")
        for el in elementos_f:
            if el.is_displayed():
                driver.execute_script("arguments[0].click();", el)
    except: pass

    # 2. Busca por textos comuns em botões
    for texto in termos:
        try:
            botoes = driver.find_elements(By.XPATH, f"//button[contains(text(), '{texto}')] | //a[contains(text(), '{texto}')]")
            for btn in botoes:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
        except: continue

def obter_aba_planilha():
    escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(ARQUIVO_JSON_GOOGLE, escopo)
    client = gspread.authorize(creds)
    planilha = client.open_by_key(ID_PLANILHA_OFICIAL)
    try:
        return planilha.worksheet(NOME_ABA_HOJE)
    except gspread.exceptions.WorksheetNotFound:
        # Se não achar a de hoje, tenta duplicar a última aba existente
        abas = planilha.worksheets()
        return planilha.duplicate_sheet(abas[-1].id, new_sheet_name=NOME_ABA_HOJE)

def extrair_apenas_numeros(texto):
    if not texto: return ""
    return re.sub(r'[^0-9,]', '', texto.replace('.', ','))

def salvar_no_google_direto(celula, valor):
    if not celula or not valor: return
    try:
        aba = obter_aba_planilha()
        aba.update_acell(celula, extrair_apenas_numeros(valor))
        print(f"📊 Salvo na {celula}: {valor}")
    except Exception as e:
        print(f"❌ Erro ao salvar {celula}: {e}")

def rodar_coleta(base_id):
    conf = BASES_VIBRA[base_id]
    print(f"\n--- 🛰️  INICIANDO: {base_id} ---")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    wait = WebDriverWait(driver, 25)

    try:
        driver.get("https://cn.vibraenergia.com.br/login")
        
        # 1. LOGIN
        aceitar_todos_cookies_vibra(driver)
        wait.until(EC.element_to_be_clickable((By.ID, "usuario"))).send_keys(conf['usuario'])
        driver.find_element(By.ID, "senha").send_keys(conf['senha'])
        driver.find_element(By.ID, "btn-acessar").click()

        # 2. PÓS-LOGIN (Limpeza de avisos)
        time.sleep(6)
        aceitar_todos_cookies_vibra(driver)

        # 3. NAVEGAÇÃO PARA PREÇOS
        print("🖱️ Clicando em CRIAR...")
        btn_criar = wait.until(EC.presence_of_element_located((By.LINK_TEXT, "CRIAR")))
        driver.execute_script("arguments[0].click();", btn_criar)
        
        time.sleep(8) 
        aceitar_todos_cookies_vibra(driver) # Pode aparecer aviso na central de pedidos

        print("🛒 Abrindo Carrinho de Preços...")
        carrinho_btn = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "i.icone-carrinho, .icone-carrinho")))
        driver.execute_script("arguments[0].click();", carrinho_btn)
        
        # 4. CAPTURA DE VALORES
        print(f"⏳ Aguardando lista de preços de {base_id}...")
        time.sleep(10) 
        aceitar_todos_cookies_vibra(driver) # Limpa se algo abrir sobre a lista
        
        driver.execute_script("window.scrollBy(0, 400);")
        
        itens = driver.find_elements(By.CLASS_NAME, "accordion-item")
        if not itens:
            itens = driver.find_elements(By.CSS_SELECTOR, "div.item-produto")

        count_s10 = 0
        count_s500 = 0

        for item in itens:
            try:
                # Localiza o valor dentro do item
                valor_el = item.find_element(By.XPATH, ".//span[@class='valor-unidade']/strong[contains(text(), ',')]")
                texto_item = item.text.upper()
                valor_texto = valor_el.text

                if "S10" in texto_item:
                    count_s10 += 1
                    celula = conf.get(f"celula_s10_{count_s10}", conf.get("celula_s10") if count_s10 == 1 else None)
                    if celula: salvar_no_google_direto(celula, valor_texto)

                elif "S500" in texto_item:
                    count_s500 += 1
                    celula = conf.get(f"celula_s500_{count_s500}", conf.get("celula_s500") if count_s500 == 1 else None)
                    if celula: salvar_no_google_direto(celula, valor_texto)
            except: continue
        
        if count_s10 == 0 and count_s500 == 0:
            print(f"⚠️ Atenção: Preços não localizados para {base_id}.")

    except Exception as e:
        print(f"❌ Erro na {base_id}: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    for b in BASES_VIBRA.keys():
        rodar_coleta(b)
    print(f"\n🚀 PROCESSO VIBRA FINALIZADO!")