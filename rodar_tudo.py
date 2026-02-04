import time
import gspread
import subprocess
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÕES ---
ARQUIVO_JSON_GOOGLE = "dados-google.json"
ID_PLANILHA_OFICIAL = "1bUvHLoUAmT4vcFy7J_qN5SendCl72Ih7ZLnF6UGd8VI" 

def preparar_aba():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(ARQUIVO_JSON_GOOGLE, scope)
    client = gspread.authorize(creds)
    
    try:
        ss = client.open_by_key(ID_PLANILHA_OFICIAL)
        print(f"✅ Conectado à planilha!")
    except Exception as e:
        print(f"❌ Erro ao abrir a planilha: {e}")
        return

    # 1. DEFINIÇÃO DAS DATAS
    agora = datetime.now()
    hoje_str = agora.strftime("%d-%m")
    
    lista_abas = ss.worksheets()
    nomes_abas = [aba.title for aba in lista_abas]

    # LÓGICA DE DECISÃO: Criar a de hoje ou a de amanhã?
    aba_hoje_existe = any(hoje_str in nome for nome in nomes_abas)

    if not aba_hoje_existe:
        data_alvo_dt = agora
        print(f"📅 Aba de hoje ({hoje_str}) não encontrada. Criando para HOJE.")
    else:
        data_alvo_dt = agora + timedelta(days=1)
        # Pular fim de semana
        if agora.weekday() == 4: data_alvo_dt = agora + timedelta(days=3)
        elif data_alvo_dt.weekday() == 5: data_alvo_dt += timedelta(days=2)
        elif data_alvo_dt.weekday() == 6: data_alvo_dt += timedelta(days=1)
        print(f"📅 Aba de hoje já existe. Preparando para o PRÓXIMO DIA ÚTIL.")

    novo_nome_aba = data_alvo_dt.strftime("Preço %d-%m")
    data_alvo_texto = data_alvo_dt.strftime("%d/%m/%Y")
    
    # 2. IDENTIFICAR ABA BASE E POSIÇÃO DE INSERÇÃO
    # Definimos a aba base (fallback para a primeira)
    aba_base = lista_abas[0]
    for i in range(1, 10):
        busca_retroativa = (data_alvo_dt - timedelta(days=i)).strftime("%d-%m")
        for aba in lista_abas:
            if busca_retroativa in aba.title:
                aba_base = aba
                break
        if aba_base and busca_retroativa in aba_base.title: break

    # LÓGICA PARA FICAR DEPOIS DO DIA 03-02
    aba_referencia = "Preço 03-02"
    if aba_referencia in nomes_abas:
        # A nova aba entra no índice da referência + 1 (logo à direita)
        novo_indice = nomes_abas.index(aba_referencia) + 1
    else:
        # Se não achar a 03-02, coloca no final para garantir que não fique no início
        novo_indice = len(nomes_abas)

    data_ontem_texto = (data_alvo_dt - timedelta(days=1)).strftime("%d/%m/%Y")
    if data_alvo_dt.weekday() == 0: 
        data_ontem_texto = (data_alvo_dt - timedelta(days=3)).strftime("%d/%m/%Y")

    print(f"✅ Aba base identificada: {aba_base.title}")

    try:
        # 3. VERIFICAR SE O ALVO JÁ EXISTE E CRIAR
        if novo_nome_aba in nomes_abas:
            print(f"⚠️ A aba {novo_nome_aba} já existe. Partindo para os coletores.")
        else:
            print(f"🔄 Criando aba {novo_nome_aba} na posição {novo_indice} (após {aba_referencia})...")
            # Inserção na posição calculada para não ficar no início
            nova_aba = ss.duplicate_sheet(aba_base.id, new_sheet_name=novo_nome_aba, insert_sheet_index=novo_indice)
            
            dados = nova_aba.get_all_values(value_render_option='FORMULA')
            
            # Atualizar datas nos cabeçalhos
            indices_ontem, indices_hoje = [3, 7, 11, 15], [4, 8, 12, 16]
            for col in indices_ontem: 
                if len(dados[10]) > col: dados[10][col] = data_ontem_texto
            for col in indices_hoje: 
                if len(dados[10]) > col: dados[10][col] = data_alvo_texto
            dados[3][2] = data_alvo_texto 

            # Limpeza e Movimentação de dados
            mapeamento = [(4, 3), (8, 7), (12, 11), (16, 15)]
            intervalos = [(11, 21), (28, 45), (53, 56), (64, 76), (85, 86), (104, 114), (118, 133), (188, 190), (210, 217), (225, 227)]
            
            formatos_limpeza = []
            for col_hoje, col_ontem in mapeamento:
                col_dif = col_hoje + 1 
                for inicio, fim in intervalos:
                    intervalo_a1 = f"{gspread.utils.rowcol_to_a1(inicio + 1, col_hoje + 1)}:{gspread.utils.rowcol_to_a1(fim + 1, col_dif + 1)}"
                    formatos_limpeza.append({
                        "range": intervalo_a1,
                        "format": {
                            "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                            "textFormat": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "bold": False}
                        }
                    })
                    for r in range(inicio, fim + 1):
                        dados[r][col_ontem] = dados[r][col_hoje]
                        if not str(dados[r][col_hoje]).startswith('='): dados[r][col_hoje] = ""

            nova_aba.update(dados, "A1", value_input_option='USER_ENTERED')
            nova_aba.batch_format(formatos_limpeza)
            print(f"✨ Aba {novo_nome_aba} pronta e posicionada!")

        # 4. RODAR COLETORES
        print("🚀 Rodando coletores...")
        for script in ["vibra.py", "shell.py", "ipiranga.py"]:
            try:
                subprocess.run(["python", script], check=True)
            except Exception as e:
                print(f"⚠️ Erro ao rodar {script}: {e}")

    except Exception as e:
        print(f"❌ Erro na automação: {e}")

if __name__ == "__main__":
    preparar_aba()