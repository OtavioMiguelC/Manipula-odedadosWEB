import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
import requests
import unicodedata
from difflib import SequenceMatcher
import os
import json
import io
import zipfile
import math
import re

# =============================================================================
# CONFIGURAÇÕES GERAIS E CONSTANTES
# =============================================================================
st.set_page_config(page_title="Ferramentas Logísticas - LINCROS AI", page_icon="📦", layout="wide")

CAMINHO_CACHE_IBGE = 'municipios_ibge_cache.json'
ARQUIVO_MODELO_REGIAO = 'Modelo Região.xlsx'
ARQUIVO_MODELO_ROTA = 'Modelo Rota.xlsx'
ARQUIVO_MODELO_TDE = "Modelo TDE.xlsx"

NOME_ABA = 'Base'
COL_CIDADE = 'Destino'
COL_UF = 'UF Destino'
COL_PRAZO = 'Prazo'
COL_IBGE = 'Codigo IBGE'

# =============================================================================
# FUNÇÕES DE APOIO E CACHE
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
    wb = Workbook()
    ws = wb.active
    ws.title = "Base"
    headers = ["Nome da Região", "Destino", "UF Destino", "Prazo", "Codigo IBGE", "DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SABADO", "FREQUENCIA"]
    ws.append(headers)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def carregar_lista_cidades_ibge():
    """Lê o cache do IBGE e retorna uma lista formatada para o selectbox."""
    if not os.path.exists(CAMINHO_CACHE_IBGE):
        API_Atualizar_Cache_IBGE()
    try:
        with open(CAMINHO_CACHE_IBGE, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        lista = [f"{m['nome']} - {m['uf']} ({m['id']})" for m in dados]
        return sorted(lista)
    except Exception:
        return []

# =============================================================================
# LÓGICA DE NEGÓCIO E PROCESSAMENTO
# =============================================================================

def processar_ibge(file_or_df):
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
    
    if isinstance(file_or_df, pd.DataFrame):
        df = file_or_df.copy()
        is_df = True
    else:
        df = pd.read_excel(file_or_df, sheet_name=NOME_ABA)
        file_or_df.seek(0)
        is_df = False

    if COL_IBGE not in df.columns: df[COL_IBGE] = ""
    count_exato, count_aprox = 0, 0
    nao_encontrados = []
    
    ibge_novos = []
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
            ibge_novos.append(ibge_encontrado)
        else:
            ibge_novos.append(row.get(COL_IBGE, ""))
            nao_encontrados.append(f"{cidade_excel_raw} - {uf_excel_raw}")

    df[COL_IBGE] = ibge_novos

    if is_df:
        return df, count_exato, count_aprox, nao_encontrados

    wb = load_workbook(file_or_df)
    ws = wb[NOME_ABA]
    col_ibge_num = df.columns.get_loc(COL_IBGE) + 1 
    for index, val in enumerate(ibge_novos):
        if val: ws.cell(row=index + 2, column=col_ibge_num).value = val

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, count_exato, count_aprox, nao_encontrados

def processar_prazos(file_destino, file_base):
    df_base = pd.read_excel(file_base, sheet_name="Base", dtype=str).fillna("")
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
            if cell.value: col_map_dest[str(cell.value).strip()] = cell.column
    else: raise Exception(f"A planilha de destino não tem dados na linha {header_row}.")
    idx_ibge = col_map_dest.get('Código IBGE da Cidade')
    if not idx_ibge:
        for k, v in col_map_dest.items():
            if "IBGE" in k.upper() and "CIDADE" in k.upper():
                idx_ibge = v; break
    if not idx_ibge: raise Exception("Coluna de IBGE não encontrada no destino.")
    cidades_atualizadas = 0
    for row_index in range(header_row + 1, sheet.max_row + 1):
        cell_ibge = sheet.cell(row=row_index, column=idx_ibge)
        if not cell_ibge.value: continue
        ibge_chave = str(cell_ibge.value).split('.')[0].strip()
        if ibge_chave in dicionario_base:
            dados_linha = dicionario_base[ibge_chave]
            if 'Prazo' in col_map_dest and 'PRAZO' in dados_linha:
                sheet.cell(row=row_index, column=col_map_dest['Prazo']).value = dados_linha['PRAZO']
            eh_caso_pontinhos = False
            if nome_coluna_texto_freq:
                texto_freq = str(dados_linha.get(nome_coluna_texto_freq, "")).strip()
                if "......" in texto_freq or (len(texto_freq) > 3 and set(texto_freq) == {'.'}): eh_caso_pontinhos = True
            if eh_caso_pontinhos:
                for dia in ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']:
                    if dia in col_map_dest: sheet.cell(row=row_index, column=col_map_dest[dia]).value = 'VERDADEIRO'
                if 'Dom' in col_map_dest: sheet.cell(row=row_index, column=col_map_dest['Dom']).value = 'FALSO'
            else:
                for dia_curto, nome_coluna_base in mapa_colunas_base.items():
                    if dia_curto in col_map_dest:
                        valor_bruto = str(dados_linha.get(nome_coluna_base, "")).upper().strip()
                        eh_verdadeiro = valor_bruto in ['VERDADEIRO', 'TRUE', 'S', 'SIM', '1', 'X']
                        sheet.cell(row=row_index, column=col_map_dest[dia_curto]).value = 'VERDADEIRO' if eh_verdadeiro else 'FALSO'
            cidades_atualizadas += 1
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, cidades_atualizadas

def processar_regiao(cnpj, file_or_df, file_modelo):
    if isinstance(file_or_df, pd.DataFrame):
        df_prazos = file_or_df.copy()
    else:
        df_prazos = pd.read_excel(file_or_df, sheet_name='Base')
        
    df_prazos['Nome da Região'] = df_prazos['Nome da Região'].astype(str).str.strip()
    df_prazos = df_prazos[df_prazos['Nome da Região'].notna() & (df_prazos['Nome da Região'].str.upper() != 'NAN') & (df_prazos['Nome da Região'] != '')]
    
    df_prazos['NomeRegiao'] = df_prazos['Nome da Região'].str.upper()
    wb_modelo = load_workbook(file_modelo)
    ws_regioes = wb_modelo['regioes']; ws_localizacoes = wb_modelo['localizacoes_atendidas']
    for ws in [ws_regioes, ws_localizacoes]:
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            for cell in row: cell.value = None
    for i, nome_regiao in enumerate(df_prazos['NomeRegiao'].unique(), start=5):
        ws_regioes[f'B{i}'] = cnpj; ws_regioes[f'C{i}'] = nome_regiao; ws_regioes[f'D{i}'] = "VERDADEIRO"
    for i, row in enumerate(df_prazos.iterrows(), start=5):
        ws_localizacoes[f'B{i}'] = row[1]['NomeRegiao']; ws_localizacoes[f'E{i}'] = row[1]['Codigo IBGE']
    output = io.BytesIO()
    wb_modelo.save(output)
    output.seek(0)
    return output

def processar_rotas(escolha_rota, cnpj_transportadora, nome_transportadora, desc_adicional, tipo_origem, valor_origem, file_modelo_regioes, file_template_rota):
    wb_modelo_regioes = load_workbook(file_modelo_regioes)
    ws_regioes = wb_modelo_regioes['regioes']
    
    regioes_encontradas = [str(ws_regioes.cell(row=i, column=3).value) for i in range(5, ws_regioes.max_row + 1) if ws_regioes.cell(row=i, column=3).value]
    if not regioes_encontradas: 
        raise Exception("Nenhuma região encontrada no modelo de regiões.")
        
    wb_rotas = load_workbook(file_template_rota)
    ws_rotas = wb_rotas["Rotas"] if "Rotas" in wb_rotas.sheetnames else wb_rotas.active
    
    next_row = 6
    while ws_rotas.cell(row=next_row, column=1).value is not None: 
        next_row += 1
        
    for regiao_destino in regioes_encontradas:
        ws_rotas.cell(row=next_row, column=1).value = f"{cnpj_transportadora} - {nome_transportadora}"
        
        desc = f"{desc_adicional} x {regiao_destino}" if desc_adicional else regiao_destino
        ws_rotas.cell(row=next_row, column=2).value = desc
        
        if tipo_origem == "Cidade (IBGE)": 
            ws_rotas.cell(row=next_row, column=3).value = valor_origem
        else: 
            ws_rotas.cell(row=next_row, column=5).value = valor_origem
            
        ws_rotas.cell(row=next_row, column=8).value = regiao_destino
        ws_rotas.cell(row=next_row, column=10).value = "VERDADEIRO"
        ws_rotas.cell(row=next_row, column=11).value = "VERDADEIRO"
        next_row += 1
        
    output = io.BytesIO()
    wb_rotas.save(output)
    output.seek(0)
    return output

def converter_freq(file):
    wb = load_workbook(file); ws = wb[NOME_ABA]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=12):
        for cell in row:
            valor = str(cell.value).strip().upper() if cell.value is not None else ""
            if valor == "S": cell.value = "VERDADEIRO"
            elif valor == "N": cell.value = "FALSO"
    output = io.BytesIO(); wb.save(output); output.seek(0); return output

def converter_freq_txt(file):
    wb = load_workbook(file); ws = wb.active
    colunas_destino = [7, 8, 9, 10, 11, 12]; letras_referencia = ['S', 'T', 'Q', 'Q', 'S', 'S']
    for row in ws.iter_rows(min_row=2):
        texto_raw = str(row[12].value or "").strip().upper()
        for i, col_idx in enumerate(colunas_destino):
            if i < len(texto_raw) and texto_raw[i] == letras_referencia[i]: row[col_idx - 1].value = True
            else: row[col_idx - 1].value = False
        row[5].value = False
    output = io.BytesIO(); wb.save(output); output.seek(0); return output

def gerar_restricoes_zip(texto_input, template_bytes, limite_linhas, categoria, tipo_f_j, usar_valor):
    linhas = [l for l in texto_input.split('\n') if l.strip()]
    grupos_por_valor = {}

    for l in linhas:
        partes = l.replace('\t', ' ').strip().split()
        if len(partes) < 2: continue 
        
        cnpj_raw = partes[0].strip().replace('.', '').replace('/', '').replace('-', '')
        possivel_valor = partes[-1].replace("R$", "").replace(".", "").replace(",", ".")
        valor_final = "0"
        razao_partes = partes[1:]

        try:
            float(possivel_valor)
            valor_final = possivel_valor
            razao_partes = partes[1:-1]
        except:
            valor_final = "0" 

        razao = " ".join(razao_partes).upper()
        chave_grupo = valor_final if usar_valor else "GERAL"

        if chave_grupo not in grupos_por_valor:
            grupos_por_valor[chave_grupo] = []
        grupos_por_valor[chave_grupo].append((cnpj_raw, razao, valor_final))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for valor_chave, dados_lista in grupos_por_valor.items():
            total_registros = len(dados_lista)
            num_arquivos = math.ceil(total_registros / limite_linhas)
            
            for i in range(num_arquivos):
                inicio = i * limite_linhas
                fim = min((i + 1) * limite_linhas, total_registros)
                lote = dados_lista[inicio:fim]
                
                template_bytes.seek(0)
                wb = load_workbook(template_bytes)
                ws = wb["Pessoa"] if "Pessoa" in wb.sheetnames else wb.active

                row_idx = 5
                for cnpj, razao, v_lin in lote:
                    ws.cell(row=row_idx, column=1).value = str(cnpj)
                    ws.cell(row=row_idx, column=2).value = str(tipo_f_j) 
                    ws.cell(row=row_idx, column=4).value = str(razao)
                    row_idx += 1

                prefixo = f"{categoria} " if categoria != "Outros" else ""
                faixa_linhas = f"{inicio} a {fim} Linhas"
                sufixo_valor = f" R${valor_chave}" if (usar_valor and valor_chave != "0") else ""
                
                nome_arquivo = f"{prefixo}{faixa_linhas}{sufixo_valor}.xlsx".strip()
                
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                zip_file.writestr(nome_arquivo, excel_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# =============================================================================
# NOVAS FUNÇÕES: GERADORES SELETIVOS LINCROS (FRETE & PRAZOS)
# =============================================================================

def gerar_tabela_frete_lincros(dados_json, cnpj="", nome_transp=""):
    """Gera a Tabela de Frete (Preços) no modelo LINCROS com 4 seções empilhadas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tabela de Frete"
    
    cnpj_usar = cnpj or dados_json.get("cnpj", "")
    nome_usar = nome_transp or dados_json.get("nome_transportadora", "")
    nome_tabela = dados_json.get("nome_tabela", "TABELA FRETE FRACIONADO")
    vig_ini = dados_json.get("vigencia_inicio", "01/01/2026")
    vig_fim = dados_json.get("vigencia_fim", "31/12/2026")
    modal = dados_json.get("modal", "ROD")
    taxas = dados_json.get("taxas", {})
    
    # 1. Cabeçalho
    ws.cell(row=1, column=1, value="CNPJ - Nome da transportadora")
    ws.cell(row=1, column=2, value=f"{cnpj_usar} - {nome_usar}")
    
    ws.cell(row=2, column=1, value="Nome da tabela")
    ws.cell(row=2, column=2, value=nome_tabela)
    
    ws.cell(row=3, column=1, value="Inicio da vigência")
    ws.cell(row=3, column=2, value=vig_ini)
    
    ws.cell(row=4, column=1, value="Fim da vigência")
    ws.cell(row=4, column=2, value=vig_fim)
    
    ws.cell(row=5, column=1, value="Modal de transporte")
    ws.cell(row=5, column=2, value=modal)
    
    # 2. Generalidades
    ws.cell(row=7, column=1, value="2. Generalidades")
    ws.cell(row=8, column=1, value="Ad Valorem (%)")
    ws.cell(row=8, column=2, value="Ad Valorem (min)")
    ws.cell(row=8, column=3, value="GRIS (%)")
    ws.cell(row=8, column=4, value="GRIS (min)")
    ws.cell(row=8, column=5, value="ICMS Destacado")
    
    ws.cell(row=9, column=1, value=taxas.get("ad_valorem_pct", 0.40))
    ws.cell(row=9, column=2, value=taxas.get("ad_valorem_min", 5.00))
    ws.cell(row=9, column=3, value=taxas.get("gris_pct", 0.30))
    ws.cell(row=9, column=4, value=taxas.get("gris_min", 6.00))
    ws.cell(row=9, column=5, value="VERDADEIRO" if taxas.get("icms_destacado", True) else "FALSO")
    
    # 3. Rotas e faixas de cálculo
    ws.cell(row=11, column=1, value="3. Rotas e faixas de cálculo")
    
    regioes = dados_json.get("regioes", [])
    faixas_peso = []
    if regioes and "frete_peso" in regioes[0]:
        faixas_peso = [item.get("peso_ate") for item in regioes[0]["frete_peso"]]
    if not faixas_peso:
        faixas_peso = [10, 20, 30, 50, 100]
        
    # Tipo da Faixa
    ws.cell(row=12, column=1, value="Tipo da faixa")
    col_idx = 4
    for _ in faixas_peso:
        ws.cell(row=12, column=col_idx, value="Peso Nominal")
        col_idx += 1
    ws.cell(row=12, column=col_idx, value="Peso Excedente")
    col_idx += 1
    ws.cell(row=12, column=col_idx, value="Valor da Mercadoria")
    
    # Intervalo da Faixa
    ws.cell(row=13, column=1, value="Intervalo da faixa")
    col_idx = 4
    peso_anterior = 0.0
    for p in faixas_peso:
        ws.cell(row=13, column=col_idx, value=f"{peso_anterior + 0.0001:.4f} a {float(p):.4f}")
        peso_anterior = float(p)
        col_idx += 1
    ws.cell(row=13, column=col_idx, value=f"{peso_anterior + 0.0001:.4f} a 9999999999.0000")
    col_idx += 1
    ws.cell(row=13, column=col_idx, value="0.0000 a 9999999999.0000")
    
    # Rota / Componente
    ws.cell(row=14, column=1, value="Rota / Componente")
    ws.cell(row=14, column=2, value="Origem")
    ws.cell(row=14, column=3, value="Destino")
    col_idx = 4
    for _ in faixas_peso:
        ws.cell(row=14, column=col_idx, value="Frete Peso (min)")
        col_idx += 1
    ws.cell(row=14, column=col_idx, value="Frete Peso")
    col_idx += 1
    ws.cell(row=14, column=col_idx, value="Pedágio")
    
    # Dados das regiões
    origem_val = dados_json.get("origem_padrao", {}).get("valor", "ORIGEM")
    row_curr = 15
    for reg in regioes:
        nome_reg = reg.get("nome_regiao", "")
        ws.cell(row=row_curr, column=1, value=f"{origem_val} X {nome_reg}")
        ws.cell(row=row_curr, column=2, value=origem_val)
        ws.cell(row=row_curr, column=3, value=nome_reg)
        
        fp_lista = reg.get("frete_peso", [])
        fp_map = {item.get("peso_ate"): item.get("valor") for item in fp_lista}
        
        c_idx = 4
        for p in faixas_peso:
            ws.cell(row=row_curr, column=c_idx, value=fp_map.get(p, 0.0))
            c_idx += 1
        ws.cell(row=row_curr, column=c_idx, value=reg.get("excedente_por_kg", 0.50))
        c_idx += 1
        ws.cell(row=row_curr, column=c_idx, value=reg.get("pedagio", 0.00))
        row_curr += 1
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def gerar_tabela_prazo_lincros(df_base, cnpj="", nome_transp=""):
    """Gera a Tabela de Prazos (prazo.xlsx) com as abas Prazo (geral) e Prazo (localizações)."""
    wb = Workbook()
    
    # Aba 1: Prazo (geral)
    ws_geral = wb.active
    ws_geral.title = "Prazo (geral)"
    
    header_geral = [
        "Código (sistema)", "Código (planilha)", "CNPJ da Transportadora", "Vigência Inicial", "Vigência Final",
        "Descrição Prazo de Entrega", "Descrição Rota", "Código IBGE da Cidade", "Estado (Sigla)", "Região",
        "CEP Inicial", "CEP Final", "Faixa de CEP", "Modal", "Somente dias úteis", "Prazo", "Tipo",
        "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"
    ]
    ws_geral.append([])
    ws_geral.append([])
    ws_geral.append([])
    ws_geral.append(header_geral)
    
    row_geral = [
        "", 1, cnpj, "01/01/2026", "31/12/2026",
        f"PRAZO {nome_transp}".strip(), f"ROTA {nome_transp}".strip(), "", "", "",
        "", "", "", "ROD", "VERDADEIRO", 1, "DIAS",
        "FALSO", "VERDADEIRO", "VERDADEIRO", "VERDADEIRO", "VERDADEIRO", "VERDADEIRO", "FALSO"
    ]
    ws_geral.append(row_geral)
    
    # Aba 2: Prazo (localizações)
    ws_loc = wb.create_sheet(title="Prazo (localizações)")
    header_loc = [
        "Código (sistema)", "Código (planilha)", "Código IBGE da Cidade", "Nome da cidade - UF",
        "Estado", "Região", "Faixa de CEP", "Prazo", "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"
    ]
    ws_loc.append([])
    ws_loc.append([])
    ws_loc.append([])
    ws_loc.append(header_loc)
    
    for idx, row in df_base.iterrows():
        ibge_val = str(row.get("Codigo IBGE", "")).split('.')[0].strip()
        cid_uf = f"{row.get('Destino', '')} - {row.get('UF Destino', '')}"
        prazo_val = row.get("Prazo", 1)
        regiao_val = row.get("Nome da Região", "")
        
        dom = "VERDADEIRO" if str(row.get("DOMINGO", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        seg = "VERDADEIRO" if str(row.get("SEGUNDA", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        ter = "VERDADEIRO" if str(row.get("TERÇA", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        qua = "VERDADEIRO" if str(row.get("QUARTA", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        qui = "VERDADEIRO" if str(row.get("QUINTA", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        sex = "VERDADEIRO" if str(row.get("SEXTA", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        sab = "VERDADEIRO" if str(row.get("SABADO", "")).upper() in ["VERDADEIRO", "TRUE", "S", "1"] else "FALSO"
        
        ws_loc.append([
            "", 1, ibge_val, cid_uf,
            row.get("UF Destino", ""), regiao_val, "", prazo_val,
            dom, seg, ter, qua, qui, sex, sab
        ])
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =============================================================================
# MÓDULO DE EXTRAÇÃO INTELIGENTE VIA GEMINI API
# =============================================================================

def extrair_dados_tabela_ia(arquivo_bytes, nome_arquivo, api_key=None):
    """Utiliza a API do Gemini para ler o arquivo da transportadora (PDF, Excel, Imagem ou Texto)
    e extrair os dados estruturados de CNPJ, Praças, Preços, Taxas e Prazos."""
    secrets_key = st.secrets.get("GEMINI_API_KEY") if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets else None
    key = api_key or os.environ.get("GEMINI_API_KEY") or secrets_key
    
    if not key:
        st.warning("⚠️ Chave GEMINI_API_KEY não detectada. Usando parser heurístico estruturado.")
        return extrair_dados_tabela_heuristico(arquivo_bytes, nome_arquivo)
    
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=key)
    except Exception as e:
        st.warning(f"Biblioteca google-genai não iniciou: {e}. Usando parser fallback.")
        return extrair_dados_tabela_heuristico(arquivo_bytes, nome_arquivo)

    prompt = """
Você é um especialista em logística e tabelas de frete. Analise o documento/tabela fornecido da transportadora e extraia todos os dados de frete, regras, prazos e localidades em formato JSON estritamente conforme a estrutura solicitada.

ATENÇÃO CRÍTICA SOBRE AS REGRAS DE NEGÓCIO:
1. MANTENHA OS NOMES DAS PRAÇAS/REGIÕES EXATAMENTE COMO A TRANSPORTADORA DEFINIU (ex: "CAPITAL SP", "GRANDE SP", "INTERIOR SUL", "ROTA 101"). NÃO crie nomes genéricos se a tabela possui nomes específicos.
2. Associe cada cidade/localidade à sua respectiva Praça/Região exatamente como a tabela indica.
3. Extraia o CNPJ, Nome/Razão Social da transportadora, vigências e modal (padrão "ROD").
4. Extraia a tabela de frete peso (faixas de peso em kg e seus respectivos valores).
5. Extraia taxas adicionais como Ad Valorem (%), GRIS (%), Pedágio, Taxa de Coleta, etc.
6. Extraia os prazos de entrega (em dias) e dias de atendimento (Segunda a Domingo).

Formato JSON esperado de resposta:
{
  "cnpj": "12345678000199",
  "nome_transportadora": "TRANSPORTADORA TESTE",
  "nome_tabela": "TABELA FRETE FRACIONADO 2026",
  "vigencia_inicio": "01/01/2026",
  "vigencia_fim": "31/12/2026",
  "modal": "ROD",
  "taxas": {
    "ad_valorem_pct": 0.4,
    "ad_valorem_min": 5.0,
    "gris_pct": 0.3,
    "gris_min": 6.0,
    "icms_destacado": true
  },
  "origem_padrao": {
    "tipo": "Cidade (IBGE)",
    "valor": "SAO PAULO"
  },
  "regioes": [
    {
      "nome_regiao": "CAPITAL-SP",
      "cidades": [
        {"cidade": "SAO PAULO", "uf": "SP", "prazo": 1},
        {"cidade": "GUARULHOS", "uf": "SP", "prazo": 1}
      ],
      "frete_peso": [
        {"peso_ate": 10, "valor": 30.0},
        {"peso_ate": 20, "valor": 45.0},
        {"peso_ate": 50, "valor": 80.0}
      ],
      "excedente_por_kg": 0.5,
      "pedagio": 4.0,
      "dias_atendimento": {"seg": true, "ter": true, "qua": true, "qui": true, "sex": true, "sab": false, "dom": false}
    }
  ]
}
Retorne APENAS o JSON puro, sem textos adicionais.
"""

    ext = nome_arquivo.split('.')[-1].lower()
    if ext in ['pdf', 'png', 'jpg', 'jpeg']:
        mime_map = {'pdf': 'application/pdf', 'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg'}
        bytes_data = arquivo_bytes.getvalue() if hasattr(arquivo_bytes, 'getvalue') else arquivo_bytes
        content_part = types.Part.from_bytes(data=bytes_data, mime_type=mime_map[ext])
        contents = [content_part, prompt]
    else:
        try:
            if ext in ['xlsx', 'xls']:
                xls = pd.ExcelFile(arquivo_bytes)
                sheets_summary = []
                for sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet).head(150)
                    sheets_summary.append(f"--- Aba: {sheet} ---\n" + df_sheet.to_string())
                texto_tabela = "\n\n".join(sheets_summary)
            else:
                texto_tabela = arquivo_bytes.getvalue().decode('utf-8', errors='ignore')
        except Exception:
            texto_tabela = str(arquivo_bytes)
        
        contents = [f"Tabela da transportadora para extrair:\n\n{texto_tabela[:30000]}", prompt]

    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        dados = json.loads(response.text)
        return dados
    except Exception as e:
        st.error(f"Erro na chamada Gemini API: {e}. Alternando para extração heurística.")
        return extrair_dados_tabela_heuristico(arquivo_bytes, nome_arquivo)

def extrair_dados_tabela_heuristico(arquivo_bytes, nome_arquivo):
    """Fallback heurístico para ler arquivos Excel/CSV caso a API key da IA não esteja configurada."""
    try:
        ext = nome_arquivo.split('.')[-1].lower()
        if ext in ['xlsx', 'xls']:
            df = pd.read_excel(arquivo_bytes)
        elif ext == 'csv':
            df = pd.read_csv(arquivo_bytes)
        else:
            raise Exception("Formato não suportado no modo heurístico sem chave Gemini.")
            
        cidades = []
        for idx, row in df.iterrows():
            col_cid = next((c for c in df.columns if 'CIDADE' in str(c).upper() or 'DESTINO' in str(c).upper()), df.columns[0])
            col_uf = next((c for c in df.columns if 'UF' in str(c).upper() or 'ESTADO' in str(c).upper()), df.columns[1] if len(df.columns) > 1 else df.columns[0])
            col_reg = next((c for c in df.columns if 'REGIAO' in str(c).upper() or 'PRAÇA' in str(c).upper() or 'PRACA' in str(c).upper() or 'ROTA' in str(c).upper()), None)
            col_prazo = next((c for c in df.columns if 'PRAZO' in str(c).upper()), None)
            
            nome_reg = str(row[col_reg]).upper() if col_reg else f"REGIAO-{str(row[col_uf]).upper()}"
            prazo_val = int(row[col_prazo]) if col_prazo and str(row[col_prazo]).isdigit() else 2
            
            cidades.append({
                "cidade": str(row[col_cid]),
                "uf": str(row[col_uf]),
                "prazo": prazo_val,
                "regiao": nome_reg
            })
            
        regioes_map = {}
        for c in cidades:
            reg_nome = c["regiao"]
            if reg_nome not in regioes_map:
                regioes_map[reg_nome] = []
            regioes_map[reg_nome].append(c)
            
        regioes_lista = []
        for reg_nome, c_list in regioes_map.items():
            regioes_lista.append({
                "nome_regiao": reg_nome,
                "cidades": c_list,
                "frete_peso": [
                    {"peso_ate": 10, "valor": 35.0},
                    {"peso_ate": 20, "valor": 50.0},
                    {"peso_ate": 50, "valor": 85.0}
                ],
                "excedente_por_kg": 0.60,
                "pedagio": 5.0,
                "dias_atendimento": {"seg": True, "ter": True, "qua": True, "qui": True, "sex": True, "sab": False, "dom": False}
            })
            
        return {
            "cnpj": "12345678000100",
            "nome_transportadora": "TRANSPORTADORA PROCESSADA",
            "nome_tabela": "TABELA FRETE 2026",
            "vigencia_inicio": "01/01/2026",
            "vigencia_fim": "31/12/2026",
            "modal": "ROD",
            "taxas": {"ad_valorem_pct": 0.4, "ad_valorem_min": 5.0, "gris_pct": 0.3, "gris_min": 6.0, "icms_destacado": True},
            "origem_padrao": {"tipo": "Cidade (IBGE)", "valor": "SAO PAULO"},
            "regioes": regioes_lista
        }
    except Exception as e:
        st.error(f"Erro no processamento heurístico: {e}")
        return {"cnpj": "", "nome_transportadora": "", "regioes": []}

def construir_df_base_do_json(dados_json):
    """Converte o JSON extraído pela IA ou heurística em um DataFrame padronizado com IBGE."""
    linhas = []
    for reg in dados_json.get("regioes", []):
        nome_reg = reg.get("nome_regiao", "").strip().upper()
        dias = reg.get("dias_atendimento", {})
        
        for cid_obj in reg.get("cidades", []):
            cid_nome = cid_obj.get("cidade", "")
            uf_nome = cid_obj.get("uf", "")
            prazo_val = cid_obj.get("prazo", 1)
            
            linhas.append({
                "Nome da Região": nome_reg,
                "Destino": cid_nome,
                "UF Destino": uf_nome,
                "Prazo": prazo_val,
                "Codigo IBGE": "",
                "DOMINGO": "VERDADEIRO" if dias.get("dom") else "FALSO",
                "SEGUNDA": "VERDADEIRO" if dias.get("seg", True) else "FALSO",
                "TERÇA": "VERDADEIRO" if dias.get("ter", True) else "FALSO",
                "QUARTA": "VERDADEIRO" if dias.get("qua", True) else "FALSO",
                "QUINTA": "VERDADEIRO" if dias.get("qui", True) else "FALSO",
                "SEXTA": "VERDADEIRO" if dias.get("sex", True) else "FALSO",
                "SABADO": "VERDADEIRO" if dias.get("sab") else "FALSO",
                "FREQUENCIA": "STQQS"
            })
            
    df_base = pd.DataFrame(linhas)
    if not df_base.empty:
        df_base, ex, ap, ne = processar_ibge(df_base)
    return df_base

# =============================================================================
# INTERFACE DO STREAMLIT
# =============================================================================

st.title("📦 Ferramentas Logísticas & IA LINCROS")

with st.sidebar:
    st.header("⚙️ Configurações Padrão")
    st.download_button(label="📥 Baixar Modelo Base (Vazio)", data=gerar_modelo_base_vazio(), file_name="Base_de_Origem_Template.xlsx", use_container_width=True)
    st.divider()
    cnpj_global = st.text_input("CNPJ Transportadora Padrão", value="12345678000100")
    nome_global = st.text_input("Nome Transportadora Padrão", value="TRANSPORTADORA LOG")
    gemini_key_input = st.text_input("Chave Gemini API (Opcional)", type="password", help="Cole sua chave para leitura de PDF/Imagens via IA")
    if gemini_key_input:
        os.environ["GEMINI_API_KEY"] = gemini_key_input
        
    st.divider()
    if st.button("Atualizar Cache IBGE", use_container_width=True):
        API_Atualizar_Cache_IBGE()
        st.success("Cache IBGE atualizado com sucesso!")

# Abas do aplicativo
tab_ia, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🤖 Processamento Inteligente (IA)",
    "🌍 Preencher IBGE", 
    "⏱️ Prazos/Freq", 
    "🗺️ Criar Região", 
    "📍 Gerar Rotas", 
    "🔄 Conv. S/N", 
    "📅 Conv. STQQS", 
    "👥 Restrições Por Pessoas"
])

# --- ABA INTELIGENTE COM IA ---
with tab_ia:
    st.markdown("### 🤖 Processamento Autônomo por IA")
    st.info("Faça upload de qualquer tabela da transportadora (PDF, Excel, Imagem ou CSV). A IA lerá as **praças da transportadora**, cruzará os IBGEs e gerará os arquivos do LINCROS no modo selecionado.")
    
    col_up, col_opt = st.columns([2, 1])
    file_ia = col_up.file_uploader("Arquivo da Transportadora (PDF, XLSX, CSV, PNG, JPG)", type=["pdf", "xlsx", "xls", "csv", "png", "jpg", "jpeg"])
    
    modo_exportacao = col_opt.radio(
        "Selecione o que deseja gerar:",
        [
            "📦 Kit Completo (Todas as Planilhas em ZIP)",
            "💰 Apenas Tabela de Frete/Preços (frete.xlsx)",
            "⏱️ Apenas Tabela de Prazos (prazo.xlsx)",
            "🗺️ Apenas Regiões + Rotas (regiao.xlsx + rota.xlsx)"
        ]
    )
    
    col_orig1, col_orig2 = st.columns(2)
    tipo_orig_ia = col_orig1.selectbox("Origem Padrão das Rotas", ["Cidade (IBGE)", "Região"])
    if tipo_orig_ia == "Cidade (IBGE)":
        lista_cidades = carregar_lista_cidades_ibge()
        cid_sel = col_orig2.selectbox("Selecione a Cidade de Origem", options=[""] + lista_cidades, key="cid_orig_ia")
        origem_val_ia = cid_sel.split("(")[-1].replace(")", "").strip() if cid_sel else "3550308"
    else:
        origem_val_ia = col_orig2.text_input("Nome da Região de Origem", value="SAO PAULO")

    if file_ia and st.button("🚀 PROCESSAR COM IA & GERAR ARQUIVOS", use_container_width=True):
        with st.spinner("Extraindo dados com IA, cruzando IBGE e montando estrutura LINCROS..."):
            try:
                dados_json = extrair_dados_tabela_ia(file_ia, file_ia.name, api_key=gemini_key_input)
                cnpj_final = cnpj_global or dados_json.get("cnpj", "")
                nome_final = (nome_global or dados_json.get("nome_transportadora", "")).upper()
                
                df_base_ia = construir_df_base_do_json(dados_json)
                st.session_state['df_base_ia'] = df_base_ia
                st.session_state['dados_json_ia'] = dados_json
                
                # Geração de arquivos
                out_regiao = processar_regiao(cnpj_final, df_base_ia, ARQUIVO_MODELO_REGIAO)
                out_rota = processar_rotas("1", cnpj_final, nome_final, "", tipo_orig_ia, origem_val_ia, out_regiao, ARQUIVO_MODELO_ROTA)
                out_frete = gerar_tabela_frete_lincros(dados_json, cnpj=cnpj_final, nome_transp=nome_final)
                out_prazo = gerar_tabela_prazo_lincros(df_base_ia, cnpj=cnpj_final, nome_transp=nome_final)
                
                if "Kit Completo" in modo_exportacao:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                        zf.writestr(f"regiao_{nome_final}.xlsx", out_regiao.getvalue())
                        zf.writestr(f"rota_{nome_final}.xlsx", out_rota.getvalue())
                        zf.writestr(f"frete_{nome_final}.xlsx", out_frete.getvalue())
                        zf.writestr(f"prazo_{nome_final}.xlsx", out_prazo.getvalue())
                    zip_buf.seek(0)
                    st.session_state['out_ia_result'] = zip_buf
                    st.session_state['out_ia_ext'] = "zip"
                    st.session_state['out_ia_name'] = f"Kit_LINCROS_{nome_final}.zip"
                    
                elif "Apenas Tabela de Frete" in modo_exportacao:
                    st.session_state['out_ia_result'] = out_frete
                    st.session_state['out_ia_ext'] = "xlsx"
                    st.session_state['out_ia_name'] = f"Frete_{nome_final}.xlsx"
                    
                elif "Apenas Tabela de Prazos" in modo_exportacao:
                    st.session_state['out_ia_result'] = out_prazo
                    st.session_state['out_ia_ext'] = "xlsx"
                    st.session_state['out_ia_name'] = f"Prazos_{nome_final}.xlsx"
                    
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                        zf.writestr(f"regiao_{nome_final}.xlsx", out_regiao.getvalue())
                        zf.writestr(f"rota_{nome_final}.xlsx", out_rota.getvalue())
                    zip_buf.seek(0)
                    st.session_state['out_ia_result'] = zip_buf
                    st.session_state['out_ia_ext'] = "zip"
                    st.session_state['out_ia_name'] = f"Regioes_e_Rotas_{nome_final}.zip"
                    
                st.success(f"✅ Processamento concluído! Praças encontradas: {len(dados_json.get('regioes', []))} | Cidades cruzadas: {len(df_base_ia)}")
            except Exception as e:
                st.error(f"Erro durante o processamento: {e}")

    if 'out_ia_result' in st.session_state:
        st.download_button(
            label=f"📥 BAIXAR {st.session_state['out_ia_name']}",
            data=st.session_state['out_ia_result'],
            file_name=st.session_state['out_ia_name'],
            mime="application/zip" if st.session_state['out_ia_ext'] == "zip" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# --- ABA 1: IBGE ---
with tab1:
    st.markdown("### Preencher Códigos IBGE")
    file_ibge = st.file_uploader("Planilha de Base", type=["xlsx"], key="ibge_file")
    if file_ibge and st.button("Processar IBGE"):
        out_bytes, exatos, aprox, nao_enc = processar_ibge(file_ibge)
        st.session_state['out_ibge'] = out_bytes
        st.success(f"✅ Sucesso! Exatos: {exatos} | IA: {aprox}")
    if 'out_ibge' in st.session_state:
        st.download_button("📥 Baixar Arquivo IBGE", data=st.session_state['out_ibge'], file_name="Base_IBGE_Preenchida.xlsx")

# --- ABA 2: Prazos e Frequência ---
with tab2:
    st.markdown("### Preencher Prazos e Frequência")
    c1, c2 = st.columns(2); f_dest = c1.file_uploader("Planilha DESTINO", type=["xlsx"]); f_base = c2.file_uploader("BASE", type=["xlsx"])
    if f_dest and f_base and st.button("Processar Prazos"):
        out, at = processar_prazos(f_dest, f_base)
        st.session_state['out_prazos'] = out; st.success(f"✅ {at} cidades atualizadas!")
    if 'out_prazos' in st.session_state:
        st.download_button("📥 Baixar Destino", data=st.session_state['out_prazos'], file_name="Destino_Prazos.xlsx")

# --- ABA 3: Criar Regiões ---
with tab3:
    st.markdown("### Criar Regiões")
    f_reg = st.file_uploader("Base de Prazos", type=["xlsx"], key="reg_up")
    if f_reg and st.button("Criar Regiões"):
        if not cnpj_global: st.warning("Preencha o CNPJ lateral")
        else:
            out = processar_regiao(cnpj_global, f_reg, ARQUIVO_MODELO_REGIAO)
            st.session_state['out_regiao'] = out
            st.success("✅ Regiões criadas!")
    if 'out_regiao' in st.session_state:
        st.download_button("📥 Baixar Regiões", data=st.session_state['out_regiao'], file_name=f"Regioes_{nome_global}.xlsx")

# --- ABA 4: Gerar Rotas ---
with tab4:
    st.markdown("### Gerar Rotas")
    
    file_modelo_regioes = st.file_uploader("1. Modelo de Região (Já preenchido no passo anterior)", type=["xlsx"], key="rota_mod_reg")
    
    st.divider()
    
    col1, col2 = st.columns(2)
    tipo_rota = col1.selectbox("Dados da Rota", ["1: ROTA - PRAZO", "2: (TRANS) (ORIGEM)"])
    cnpj_rota = col2.text_input("CNPJ Transportadora (se difere do padrão)", value=cnpj_global)
    nome_transp_rota = col1.text_input("Nome Transportadora", value=nome_global) 
    
    desc_rota = col2.text_input("Desc. Adicional (Opcional)")
    
    st.markdown("#### Dados de Origem")
    col3, col4 = st.columns(2)
    tipo_origem = col3.radio("Definir origem por:", ["Cidade (IBGE)", "Região"])
    
    if tipo_origem == "Cidade (IBGE)":
        lista_cidades = carregar_lista_cidades_ibge()
        cidade_selecionada = col4.selectbox("Selecione ou digite a Cidade de Origem", options=[""] + lista_cidades)
        
        if cidade_selecionada:
            valor_origem = cidade_selecionada.split("(")[-1].replace(")", "").strip()
        else:
            valor_origem = ""
    else:
        valor_origem = col4.text_input("Nome da Região de Origem")
    
    if file_modelo_regioes and st.button("Gerar Rotas"):
        if not cnpj_rota or not valor_origem:
            st.warning("⚠️ CNPJ e Origem são obrigatórios.")
        elif not os.path.exists(ARQUIVO_MODELO_ROTA):
            st.error(f"⚠️ O arquivo original '{ARQUIVO_MODELO_ROTA}' não foi encontrado na pasta do projeto!")
        else:
            with st.spinner(f"Gerando rotas baseadas no banco ({ARQUIVO_MODELO_ROTA})..."):
                try:
                    out_bytes = processar_rotas(
                        tipo_rota.split(":")[0], 
                        cnpj_rota, 
                        nome_transp_rota.upper(), 
                        desc_rota.upper(), 
                        tipo_origem, 
                        valor_origem, 
                        file_modelo_regioes, 
                        ARQUIVO_MODELO_ROTA
                    )
                    st.session_state['out_rotas'] = out_bytes
                    
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
    f_sn = st.file_uploader("Planilha S/N", type=["xlsx"])
    if f_sn and st.button("Converter S/N"):
        st.session_state['out_sn'] = converter_freq(f_sn); st.success("Convertido!")
    if 'out_sn' in st.session_state:
        st.download_button("📥 Baixar S/N", data=st.session_state['out_sn'], file_name="S_N_Convertido.xlsx")

# --- ABA 6: Converter STQQS ---
with tab6:
    f_st = st.file_uploader("Planilha STQQS", type=["xlsx"])
    if f_st and st.button("Converter STQQS"):
        st.session_state['out_stqqs'] = converter_freq_txt(f_st); st.success("Convertido!")
    if 'out_stqqs' in st.session_state:
        st.download_button("📥 Baixar STQQS", data=st.session_state['out_stqqs'], file_name="STQQS_Convertido.xlsx")

# --- ABA 7: RESTRIÇÕES POR PESSOAS ---
with tab7:
    st.markdown("### 👥 Restrições Por Pessoas")
    st.info("Gera arquivos de cadastro de pessoas separando por valor e limite de linhas.")

    col_btn1, col_btn2, col_btn3 = st.columns(3)
    categoria_fleg = col_btn1.radio("Selecione a Categoria:", ["TDE", "TAE", "Outros"], horizontal=True)
    tipo_pessoa_fleg = col_btn2.radio("Tipo de Pessoa:", ["J", "F"], horizontal=True)
    usar_valor_fleg = col_btn3.checkbox("Filtrar/Incluir Valor no nome?", value=True)

    limite_linhas_rest = st.number_input("Linhas por arquivo:", min_value=1, value=500, step=100)
    
    st.markdown("**📋 Cole abaixo os dados (Formato: CNPJ | Razão Social | Valor Opcional):**")
    texto_rest = st.text_area("Dados:", height=250, placeholder="Ex: 12345678000100 EMPRESA LTDA 250,00", key="txt_rest")

    if st.button("🚀 PROCESSAR RESTRIÇÕES POR PESSOAS", use_container_width=True):
        if not os.path.exists(ARQUIVO_MODELO_TDE):
            st.error(f"Arquivo '{ARQUIVO_MODELO_TDE}' não encontrado na pasta do projeto!")
        elif not texto_rest.strip():
            st.warning("Cole os dados para processar.")
        else:
            with st.spinner("Estruturando dados e gerando arquivos..."):
                try:
                    with open(ARQUIVO_MODELO_TDE, "rb") as f:
                        template_bytes = io.BytesIO(f.read())
                    
                    zip_out = gerar_restricoes_zip(
                        texto_rest, 
                        template_bytes, 
                        int(limite_linhas_rest), 
                        categoria_fleg, 
                        tipo_pessoa_fleg, 
                        usar_valor_fleg
                    )
                    st.session_state['zip_rest'] = zip_out
                    st.success("✅ Arquivos gerados e compactados com sucesso!")
                except Exception as e:
                    st.error(f"Erro: {e}")

    if 'zip_rest' in st.session_state:
        st.download_button(
            label="📥 BAIXAR ARQUIVOS GERADOS (ZIP)",
            data=st.session_state['zip_rest'],
            file_name=f"Restricoes_{categoria_fleg}.zip",
            mime="application/zip",
            use_container_width=True
        )
