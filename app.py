import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
import requests
import unicodedata
from difflib import SequenceMatcher
import os
import json
import io

# =============================================================================
# CONFIGURAÇÕES GERAIS E CONSTANTES
# =============================================================================
st.set_page_config(page_title="Ferramentas Logísticas", page_icon="📦", layout="wide")

# =============================================================================
# INJEÇÃO CSS - LIQUID GLASSMORPHISM (ESTILO iOS)
# =============================================================================
st.markdown("""
<style>
/* 1. Fundo principal com gradiente animado (necessário para ver o desfoque) */
.stApp {
    background: linear-gradient(-45deg, #1e1e2f, #2a2a40, #1f3b4d, #252538);
    background-size: 400% 400%;
    animation: gradient 15s ease infinite;
}

@keyframes gradient {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

/* 2. Barra Lateral (Sidebar) com efeito de vidro */
[data-testid="stSidebar"] {
    background: rgba(255, 255, 255, 0.05) !important;
    backdrop-filter: blur(15px) !important;
    -webkit-backdrop-filter: blur(15px) !important;
    border-right: 1px solid rgba(255, 255, 255, 0.1);
}

/* 3. Estilização das Abas (Tabs) */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(255, 255, 255, 0.08);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border-radius: 15px;
    padding: 6px;
    border: 1px solid rgba(255, 255, 255, 0.15);
    gap: 8px;
}

.stTabs [data-baseweb="tab"] {
    background: transparent;
    border-radius: 10px;
    color: #e0e0e0 !important;
    border: none !important;
}

.stTabs [aria-selected="true"] {
    background: rgba(255, 255, 255, 0.2) !important;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
    font-weight: bold;
}

/* 4. Botões Translúcidos e Arredondados */
.stButton > button {
    background: rgba(255, 255, 255, 0.1) !important;
    backdrop-filter: blur(10px) !important;
    -webkit-backdrop-filter: blur(10px) !important;
    border: 1px solid rgba(255, 255, 255, 0.2) !important;
    border-radius: 20px !important;
    color: white !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
}

.stButton > button:hover {
    background: rgba(255, 255, 255, 0.25) !important;
    transform: translateY(-2px);
    box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2) !important;
}

/* 5. Área de Upload de Arquivos */
[data-testid="stFileUploadDropzone"] {
    background: rgba(255, 255, 255, 0.05) !important;
    backdrop-filter: blur(10px) !important;
    -webkit-backdrop-filter: blur(10px) !important;
    border: 1px dashed rgba(255, 255, 255, 0.3) !important;
    border-radius: 20px !important;
}

/* 6. Inputs de Texto e Caixas de Seleção */
.stTextInput > div > div > input, .stSelectbox > div > div > div {
    background: rgba(255, 255, 255, 0.1) !important;
    backdrop-filter: blur(10px) !important;
    border: 1px solid rgba(255, 255, 255, 0.15) !important;
    border-radius: 12px !important;
    color: white !important;
}

/* Remove a cor sólida de fundo de blocos internos para manter o vidro visível */
[data-testid="stVerticalBlock"] > div {
    background-color: transparent !important;
}
</style>
""", unsafe_allow_html=True)


CAMINHO_CACHE_IBGE = 'municipios_ibge_cache.json'
ARQUIVO_MODELO_REGIAO = 'Modelo Região.xlsx'
ARQUIVO_MODELO_ROTA = 'Modelo Rota.xlsx'

NOME_ABA = 'Base'
COL_CIDADE = 'Destino'
COL_UF = 'UF Destino'
COL_PRAZO = 'Prazo'
COL_IBGE = 'Codigo IBGE'

# =============================================================================
# FUNÇÕES DE APOIO E GERAÇÃO DE MODELO
# =============================================================================
def normalizar(texto):
    if pd.isna(texto): return ""
    texto = str(texto).strip().upper()
    try:
        texto_normalizado = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    except:
        return texto 
    texto_normalizado = texto_normalizado.replace("'", " ").replace(".", " ").replace("-", " ")
    while "  " in texto_normalizado:
        texto_normalizado = texto_normalizado.replace("  ", " ")
    return texto_normalizado.strip()

def API_Atualizar_Cache_IBGE():
    try:
        r = requests.get("https://servicodados.ibge.gov.br/api/v1/localidades/municipios", timeout=30, verify=False)
        municipios_api = r.json()

        mapa_final = {}
        for m in municipios_api:
            if not isinstance(m, dict): continue
            nome = normalizar(m.get('nome', ''))
            micro = m.get("microrregiao") or {}
            meso = micro.get("mesorregiao") or {}
            uf_obj = meso.get("UF") or {}
            uf = normalizar(uf_obj.get("sigla", ""))
            
            if nome and uf: 
                mapa_final[(nome, uf)] = {'nome': nome, 'uf': uf, 'id': m.get('id')}
        
        lista_dados = list(mapa_final.values())
        with open(CAMINHO_CACHE_IBGE, 'w', encoding='utf-8') as f:
            json.dump(lista_dados, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        st.error(f"Falha ao processar dados IBGE: {e}")
        return False

def gerar_modelo_base_vazio():
    """Gera o modelo 'Base' apenas com os cabeçalhos, sempre vazio."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Base"
    
    headers = [
        "Nome da Região", "Destino", "UF Destino", "Prazo", "Codigo IBGE", 
        "DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", 
        "SEXTA", "SABADO", "FREQUENCIA"
    ]
    ws.append(headers)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =============================================================================
# LÓGICA DE NEGÓCIO
# =============================================================================

def processar_ibge(file):
    if not os.path.exists(CAMINHO_CACHE_IBGE):
        API_Atualizar_Cache_IBGE()

    with open(CAMINHO_CACHE_IBGE, 'r', encoding='utf-8') as f:
        municipios_lista = json.load(f)

    db_ibge_por_uf = {}
    mapa_exato = {} 
    for m in municipios_lista:
        nome_norm = normalizar(m['nome'])
        uf_norm = normalizar(m['uf'])
        id_ibge = m['id']
        mapa_exato[(nome_norm, uf_norm)] = id_ibge
        if uf_norm not in db_ibge_por_uf: db_ibge_por_uf[uf_norm] = []
        db_ibge_por_uf[uf_norm].append({'nome_norm': nome_norm, 'id': id_ibge, 'nome_real': m['nome']})

    df = pd.read_excel(file, sheet_name=NOME_ABA)
    file.seek(0) 
    
    if COL_IBGE not in df.columns: df[COL_IBGE] = ""
    wb = load_workbook(file)
    ws = wb[NOME_ABA]
    col_ibge_num = df.columns.get_loc(COL_IBGE) + 1 
    
    count_exato, count_aprox = 0, 0
    nao_encontrados = []

    for index, row in df.iterrows():
        cidade_excel_raw = str(row[COL_CIDADE])
        uf_excel_raw = str(row[COL_UF])
        cidade_norm = normalizar(cidade_excel_raw)
        uf_norm = normalizar(uf_excel_raw)
        ibge_encontrado = mapa_exato.get((cidade_norm, uf_norm))
        
        if ibge_encontrado:
            count_exato += 1
        else:
            if uf_norm in db_ibge_por_uf:
                candidatos = db_ibge_por_uf[uf_norm]
                melhor_ratio = 0.0
                melhor_candidato = None
                for item in candidatos:
                    ratio = SequenceMatcher(None, cidade_norm, item['nome_norm']).ratio()
                    if ratio > melhor_ratio:
                        melhor_ratio = ratio
                        melhor_candidato = item
                if melhor_ratio >= 0.80 and melhor_candidato:
                    ibge_encontrado = melhor_candidato['id']
                    count_aprox += 1
                    
        if ibge_encontrado:
            ws.cell(row=index + 2, column=col_ibge_num).value = ibge_encontrado
        else:
            nao_encontrados.append(f"{cidade_excel_raw} - {uf_excel_raw}")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, count_exato, count_aprox, nao_encontrados

def processar_prazos(file_destino, file_base):
    df_base = pd.read_excel(file_base, sheet_name="Base", dtype=str)
    df_base.columns = [str(col).strip().upper() for col in df_base.columns]
    
    nome_coluna_texto_freq = None
    for col in df_base.columns:
        amostra = df_base[col].head(20).astype(str).tolist()
        for val in amostra:
            if "....." in val or "STQQS" in val:
                nome_coluna_texto_freq = col; break
        if nome_coluna_texto_freq: break
    
    mapa_colunas_base = {}
    padroes_busca = {'Seg': 'SEG', 'Ter': 'TER', 'Qua': 'QUA', 'Qui': 'QUI', 'Sex': 'SEX', 'Sáb': 'SAB', 'Dom': 'DOM'}
    for dia_destino, texto_busca in padroes_busca.items():
        for col_real in df_base.columns:
            if texto_busca in col_real:
                mapa_colunas_base[dia_destino] = col_real; break
    
    col_ibge_nome = next((c for c in df_base.columns if 'IBGE' in c), 'CODIGO IBGE')
    df_base = df_base.dropna(subset=[col_ibge_nome])
    df_base['IBGE_LIMPO'] = df_base[col_ibge_nome].apply(lambda x: str(x).split('.')[0].strip())
    df_base = df_base.drop_duplicates(subset=['IBGE_LIMPO'], keep='first')
    dicionario_base = df_base.set_index('IBGE_LIMPO').to_dict('index')
    
    wb = load_workbook(file_destino)
    nome_aba_destino = "Prazo (localizações)"
    sheet = wb[nome_aba_destino] if nome_aba_destino in wb.sheetnames else wb.active
        
    header_row = 4
    col_map_dest = {}
    if sheet.max_row >= header_row:
        for cell in sheet[header_row]:
            if cell.value:
                col_map_dest[str(cell.value).strip()] = cell.column
    else: raise Exception(f"A planilha de destino não tem dados na linha {header_row}.")
        
    nome_col_ibge_dest = 'Código IBGE da Cidade'
    idx_ibge = col_map_dest.get(nome_col_ibge_dest)
    if not idx_ibge:
        for k, v in col_map_dest.items():
            if "IBGE" in k.upper() and "CIDADE" in k.upper():
                idx_ibge = v; break
    if not idx_ibge: raise Exception("Coluna de IBGE não encontrada no destino.")

    cidades_atualizadas = 0
    for row_index in range(header_row + 1, sheet.max_row + 1):
        cell_ibge = sheet.cell(row=row_index, column=idx_ibge)
        if not cell_ibge.value: continue
        
        ibge_dest_raw = str(cell_ibge.value)
        ibge_chave = ibge_dest_raw.split('.')[0].strip()
        
        if ibge_chave in dicionario_base:
            dados_linha = dicionario_base[ibge_chave]
            if 'Prazo' in col_map_dest and 'PRAZO' in dados_linha:
                sheet.cell(row=row_index, column=col_map_dest['Prazo']).value = dados_linha['PRAZO']
            
            eh_caso_pontinhos = False
            if nome_coluna_texto_freq:
                texto_freq = str(dados_linha.get(nome_coluna_texto_freq, "")).strip()
                if "......" in texto_freq or (len(texto_freq) > 3 and set(texto_freq) == {'.'}): 
                    eh_caso_pontinhos = True
            
            if eh_caso_pontinhos:
                for dia in ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']:
                    if dia in col_map_dest: 
                        sheet.cell(row=row_index, column=col_map_dest[dia]).value = 'VERDADEIRO'
                if 'Dom' in col_map_dest: 
                    sheet.cell(row=row_index, column=col_map_dest['Dom']).value = 'FALSO'
            else:
                for dia_curto, nome_coluna_base in mapa_colunas_base.items():
                    if dia_curto in col_map_dest:
                        valor_bruto = str(dados_linha.get(nome_coluna_base, "")).upper().strip()
                        eh_verdadeiro = valor_bruto in ['VERDADEIRO', 'TRUE', 'S', 'SIM', '1', 'X']
                        col_idx = col_map_dest[dia_curto]
                        sheet.cell(row=row_index, column=col_idx).value = 'VERDADEIRO' if eh_verdadeiro else 'FALSO'
            cidades_atualizadas += 1
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, cidades_atualizadas

def processar_regiao(cnpj, file_base, file_modelo):
    df_prazos = pd.read_excel(file_base, sheet_name='Base')
    
    # Limpa e filtra o campo "Nome da Região"
    df_prazos['Nome da Região'] = df_prazos['Nome da Região'].astype(str).str.strip()
    df_prazos = df_prazos[df_prazos['Nome da Região'].notna() & (df_prazos['Nome da Região'].str.upper() != 'NAN') & (df_prazos['Nome da Região'] != '')]
    df_prazos = df_prazos.drop_duplicates(subset=['Nome da Região', 'Codigo IBGE', 'Prazo'])
    
    # Define o NomeRegiao exatamente como consta na coluna da planilha
    df_prazos['NomeRegiao'] = df_prazos['Nome da Região'].str.upper()
    
    wb_modelo = load_workbook(file_modelo)
    ws_regioes = wb_modelo['regioes']
    ws_localizacoes = wb_modelo['localizacoes_atendidas']
    
    for ws in [ws_regioes, ws_localizacoes]:
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            for cell in row: cell.value = None
    
    for i, nome_regiao in enumerate(df_prazos['NomeRegiao'].unique(), start=5):
        ws_regioes[f'B{i}'] = cnpj
        ws_regioes[f'C{i}'] = nome_regiao
        ws_regioes[f'D{i}'] = "VERDADEIRO"
    
    for i, row in enumerate(df_prazos.iterrows(), start=5):
        ws_localizacoes[f'B{i}'] = row[1]['NomeRegiao']
        ws_localizacoes[f'E{i}'] = row[1]['Codigo IBGE']
    
    output = io.BytesIO()
    wb_modelo.save(output)
    output.seek(0)
    return output

def processar_rotas(escolha_rota, cnpj_transportadora, nome_transportadora, desc_adicional, ibge_origem, file_modelo_regioes, file_template_rota):
    # Lendo as regiões do arquivo gerado anteriormente
    wb_modelo_regioes = load_workbook(file_modelo_regioes)
    ws_regioes = wb_modelo_regioes['regioes']
    regioes_encontradas = []
    
    for row_index in range(5, ws_regioes.max_row + 1):
        nome_regiao = ws_regioes.cell(row=row_index, column=3).value
        if nome_regiao: regioes_encontradas.append(str(nome_regiao))
        
    if not regioes_encontradas:
        raise Exception("Nenhuma região encontrada no modelo de regiões.")
        
    # Carregando o template "rota.xlsx" original do repositório
    wb_rotas = load_workbook(file_template_rota)
    
    if "Rotas" in wb_rotas.sheetnames:
        ws_rotas = wb_rotas["Rotas"]
    else:
        ws_rotas = wb_rotas.active

    # Descobrindo a próxima linha vazia (baseado no formato que preenche a partir da linha 6)
    next_row = 6
    while ws_rotas.cell(row=next_row, column=1).value is not None:
        next_row += 1
    
    # Preenchendo as rotas na aba correta do template
    for regiao_destino in regioes_encontradas:
        ws_rotas.cell(row=next_row, column=1).value = f"{cnpj_transportadora} - {nome_transportadora}"
        desc = f"{desc_adicional} x {regiao_destino}" if desc_adicional else regiao_destino
        ws_rotas.cell(row=next_row, column=2).value = desc
        ws_rotas.cell(row=next_row, column=3).value = ibge_origem
        ws_rotas.cell(row=next_row, column=8).value = regiao_destino
        ws_rotas.cell(row=next_row, column=10).value = "VERDADEIRO"
        ws_rotas.cell(row=next_row, column=11).value = "VERDADEIRO"
        next_row += 1
        
    output = io.BytesIO()
    wb_rotas.save(output)
    output.seek(0)
    return output

def converter_freq(file):
    wb = load_workbook(file)
    ws = wb[NOME_ABA]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=12):
        for cell in row:
            valor = str(cell.value).strip().upper() if cell.value is not None else ""
            if valor == "S": cell.value = "VERDADEIRO"
            elif valor == "N": cell.value = "FALSO"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def converter_freq_txt(file):
    wb = load_workbook(file)
    ws = wb.active
    colunas_destino = [7, 8, 9, 10, 11, 12]
    letras_referencia = ['S', 'T', 'Q', 'Q', 'S', 'S']
    coluna_frequencia_texto = 13 
    for row in ws.iter_rows(min_row=2):
        texto_raw = str(row[coluna_frequencia_texto - 1].value or "").strip().upper()
        for i, col_idx in enumerate(colunas_destino):
            if i < len(texto_raw) and texto_raw[i] == letras_referencia[i]: row[col_idx - 1].value = True
            else: row[col_idx - 1].value = False
        row[5].value = False # Domingo
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =============================================================================
# INTERFACE DO STREAMLIT
# =============================================================================

st.title("📦 Ferramentas Gerais - Logística")

with st.sidebar:
    st.header("⚙️ Configuração")
    
    st.download_button(
        label="📥 Baixar Modelo Base (Vazio)", 
        data=gerar_modelo_base_vazio(), 
        file_name="Base_de_Origem_Template.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.divider()
    
    cnpj_global = st.text_input("CNPJ Transportadora Padrão", help="Usado nas ferramentas de Região e Rotas")
    nome_global = st.text_input("Nome Transportadora Padrão", help="Usado para dar nome automático aos arquivos baixados")
    
    if st.button("Atualizar Cache IBGE", use_container_width=True):
        with st.spinner("Buscando dados da API..."):
            API_Atualizar_Cache_IBGE()
            st.success("Cache atualizado!")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🌍 Preencher IBGE", 
    "⏱️ Prazos/Freq", 
    "🗺️ Criar Região", 
    "📍 Gerar Rotas", 
    "🔄 Conv. S/N", 
    "📅 Conv. STQQS"
])

# --- ABA 1: IBGE ---
with tab1:
    st.markdown("### Preencher Códigos IBGE")
    st.write("Cruza base de dados de cidade/UF com a tabela do IBGE.")
    file_ibge = st.file_uploader("Planilha de Base", type=["xlsx"], key="ibge_file")
    
    if file_ibge and st.button("Processar IBGE"):
        with st.spinner("Processando dados e aplicando Inteligência..."):
            try:
                out_bytes, exatos, aprox, nao_enc = processar_ibge(file_ibge)
                st.session_state['out_ibge'] = out_bytes
                st.success(f"✅ Sucesso! Exatos: {exatos} | IA Aprox: {aprox} | N/E: {len(nao_enc)}")
            except Exception as e:
                st.error(f"Erro: {e}")
                
    if 'out_ibge' in st.session_state:
        st.download_button("📥 Baixar Arquivo IBGE Preenchido", data=st.session_state['out_ibge'], file_name="Base_IBGE_Preenchida.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- ABA 2: Prazos e Frequência ---
with tab2:
    st.markdown("### Preencher Prazos e Frequência")
    col1, col2 = st.columns(2)
    file_destino = col1.file_uploader("Planilha DESTINO", type=["xlsx"], key="prazo_dest")
    file_base = col2.file_uploader("Planilha BASE DE PRAZOS", type=["xlsx"], key="prazo_base")
    
    if file_destino and file_base and st.button("Processar Prazos"):
        with st.spinner("Cruzando dados..."):
            try:
                out_bytes, atualizadas = processar_prazos(file_destino, file_base)
                st.session_state['out_prazos'] = out_bytes
                st.success(f"✅ {atualizadas} cidades atualizadas com sucesso!")
            except Exception as e:
                st.error(f"Erro: {e}")
                
    if 'out_prazos' in st.session_state:
        st.download_button("📥 Baixar Destino Atualizado", data=st.session_state['out_prazos'], file_name="Destino_Prazos_Preenchidos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- ABA 3: Criar Região ---
with tab3:
    st.markdown("### Criar Regiões")
    
    file_base_reg = st.file_uploader("Base de Prazos", type=["xlsx"], key="reg_base")
    
    if file_base_reg and st.button("Criar Regiões"):
        if not cnpj_global:
            st.warning("⚠️ Preencha o CNPJ no menu lateral (Configuração).")
        elif not os.path.exists(ARQUIVO_MODELO_REGIAO):
            st.error(f"⚠️ O arquivo original '{ARQUIVO_MODELO_REGIAO}' não foi encontrado na pasta do projeto!")
        else:
            with st.spinner(f"Extraindo dados do banco ({ARQUIVO_MODELO_REGIAO})..."):
                try:
                    out_bytes = processar_regiao(cnpj_global, file_base_reg, ARQUIVO_MODELO_REGIAO)
                    st.session_state['out_regiao'] = out_bytes
                    
                    # Salva o nome sugerido para o arquivo usando a Transportadora Padrão
                    nome_sugerido = f"Regioes_{nome_global.strip()}.xlsx" if nome_global.strip() else "Modelo_Regioes_Preenchido.xlsx"
                    st.session_state['nome_arq_regiao'] = nome_sugerido
                    
                    st.success("✅ Regiões criadas com sucesso usando o nome padrão da tabela!")
                except Exception as e:
                    st.error(f"Erro: {e}")
                    
    if 'out_regiao' in st.session_state:
        st.download_button(
            label="📥 Baixar Regiões", 
            data=st.session_state['out_regiao'], 
            file_name=st.session_state.get('nome_arq_regiao', 'Modelo_Regioes_Preenchido.xlsx'), 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- ABA 4: Gerar Rotas ---
with tab4:
    st.markdown("### Gerar Rotas")
    
    # Recebe a planilha gerada na etapa 3
    file_modelo_regioes = st.file_uploader("1. Modelo de Região (Já preenchido no passo anterior)", type=["xlsx"], key="rota_mod_reg")
    
    st.divider()
    
    col1, col2 = st.columns(2)
    tipo_rota = col1.selectbox("Dados da Rota", ["1: ROTA - PRAZO", "2: (TRANS) (ORIGEM)"])
    cnpj_rota = col2.text_input("CNPJ Transportadora (se difere do padrão)", value=cnpj_global)
    nome_transp_rota = col1.text_input("Nome Transportadora", value=nome_global) # Puxa do lateral por padrão
    desc_rota = col2.text_input("Desc. Adicional (Opcional)")
    ibge_orig = st.text_input("IBGE Origem")
    
    if file_modelo_regioes and st.button("Gerar Rotas"):
        if not cnpj_rota or not ibge_orig:
            st.warning("⚠️ CNPJ e IBGE Origem são obrigatórios.")
        elif not os.path.exists(ARQUIVO_MODELO_ROTA):
            st.error(f"⚠️ O arquivo original '{ARQUIVO_MODELO_ROTA}' não foi encontrado na pasta do projeto!")
        else:
            with st.spinner(f"Gerando rotas baseadas no banco ({ARQUIVO_MODELO_ROTA})..."):
                try:
                    out_bytes = processar_rotas(tipo_rota.split(":")[0], cnpj_rota, nome_transp_rota.upper(), desc_rota.upper(), ibge_orig, file_modelo_regioes, ARQUIVO_MODELO_ROTA)
                    st.session_state['out_rotas'] = out_bytes
                    
                    # Salva o nome sugerido para o arquivo da Rota
                    nome_sugerido_rota = f"Rotas_{nome_transp_rota.strip()}.xlsx" if nome_transp_rota.strip() else "Rotas_Preenchidas.xlsx"
                    st.session_state['nome_arq_rota'] = nome_sugerido_rota
                    
                    st.success("✅ Rotas estruturadas com sucesso dentro do modelo original!")
                except Exception as e:
                    st.error(f"Erro: {e}")
                    
    if 'out_rotas' in st.session_state:
        st.download_button(
            label="📥 Baixar Planilha de Rotas Preenchida", 
            data=st.session_state['out_rotas'], 
            file_name=st.session_state.get('nome_arq_rota', 'Rotas_Preenchidas.xlsx'), 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- ABA 5: Converter S/N ---
with tab5:
    st.markdown("### Converter Sim/Não")
    st.write("Substitui as letras S/N por VERDADEIRO/FALSO.")
    file_sn = st.file_uploader("Planilha S/N", type=["xlsx"], key="sn_file")
    
    if file_sn and st.button("Converter S/N"):
        with st.spinner("Convertendo..."):
            try:
                out_bytes = converter_freq(file_sn)
                st.session_state['out_sn'] = out_bytes
                st.success("✅ Conversão finalizada!")
            except Exception as e:
                st.error(f"Erro: {e}")
                
    if 'out_sn' in st.session_state:
        st.download_button("📥 Baixar Arquivo S/N", data=st.session_state['out_sn'], file_name="Frequencias_SN_Convertidas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- ABA 6: Converter STQQS ---
with tab6:
    st.markdown("### Converter Formato STQQS")
    st.write("Traduz strings semanais (ex: ST.QS.) para colunas lógicas.")
    file_stqqs = st.file_uploader("Planilha STQQS", type=["xlsx"], key="stqqs_file")
    
    if file_stqqs and st.button("Converter Texto Semanal"):
        with st.spinner("Convertendo..."):
            try:
                out_bytes = converter_freq_txt(file_stqqs)
                st.session_state['out_stqqs'] = out_bytes
                st.success("✅ Conversão finalizada!")
            except Exception as e:
                st.error(f"Erro: {e}")
                
    if 'out_stqqs' in st.session_state:
        st.download_button("📥 Baixar Arquivo STQQS", data=st.session_state['out_stqqs'], file_name="Frequencias_Texto_Convertidas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
