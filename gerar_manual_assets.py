import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import matplotlib.pyplot as plt
import pandas as pd
import os
import io

def render_table_image(df, title, filename, highlight_cols=None, bg_header='#1a365d'):
    fig, ax = plt.subplots(figsize=(8.5, len(df)*0.55 + 1.2))
    ax.axis('tight')
    ax.axis('off')
    
    plt.title(title, fontsize=12, fontweight='bold', pad=12, color='#1a365d')
    
    tb = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
    tb.auto_set_font_size(False)
    tb.set_fontsize(9.5)
    
    for k, cell in tb.get_celld().items():
        cell.set_edgecolor('#cbd5e0')
        row_idx, col_idx = k
        if row_idx == 0:
            cell.set_facecolor(bg_header)
            cell.set_text_props(color='white', weight='bold')
        else:
            col_name = df.columns[col_idx]
            if highlight_cols and col_name in highlight_cols:
                cell.set_facecolor('#fef08a') # Marca texto amarelo
                cell.set_text_props(color='#854d0e', weight='bold')
            else:
                cell.set_facecolor('#f7fafc' if row_idx % 2 == 0 else '#ffffff')
                
    plt.tight_layout()
    plt.savefig(filename, bbox_inches='tight', dpi=180)
    plt.close()

# Gerar Diagramas
os.makedirs('manual_assets', exist_ok=True)

# 1. Cadastro CEP
df_cep_in = pd.DataFrame({'CEP Inicial (Entrada Analista)': ['80010-000', '97700000', '84500000'], 'CEP Final (Entrada Analista)': ['80010-000', '97700000', '84500000']})
df_cep_out = pd.DataFrame({'ID Localização': ['', '', ''], 'CEP Inicial': ['80010000', '97700000', '84500000'], 'CEP Final': ['80010000', '97700000', '84500000'], 'Nome (Retornado Sistema)': ['Curitiba - PR', 'Santiago - RS', 'Irati - PR'], 'Ativo': ['VERDADEIRO', 'VERDADEIRO', 'VERDADEIRO']})
render_table_image(df_cep_in, "1. ENTRADA: Dados de CEP colados pelo Analista", "manual_assets/cep_in.png", bg_header='#2b6cb0')
render_table_image(df_cep_out, "2. SAÍDA: Planilha Modelo CEP Retornada pelo Sistema", "manual_assets/cep_out.png", highlight_cols=['Nome (Retornado Sistema)', 'Ativo'], bg_header='#2f855a')

# 2. IBGE
df_ibge_in = pd.DataFrame({'Destino (Cidade)': ['Curitiba', 'São Paulo', 'Porto Alegre'], 'UF Destino': ['PR', 'SP', 'RS'], 'Codigo IBGE': ['', '', '']})
df_ibge_out = pd.DataFrame({'Destino (Cidade)': ['Curitiba', 'São Paulo', 'Porto Alegre'], 'UF Destino': ['PR', 'SP', 'RS'], 'Codigo IBGE (Gerado Sistema)': ['4106902', '3550308', '4314902']})
render_table_image(df_ibge_in, "1. ENTRADA: Planilha Base de Origem sem IBGE", "manual_assets/ibge_in.png", bg_header='#2b6cb0')
render_table_image(df_ibge_out, "2. SAÍDA: Planilha Base com Códigos IBGE Preenchidos", "manual_assets/ibge_out.png", highlight_cols=['Codigo IBGE (Gerado Sistema)'], bg_header='#2f855a')

# 3. Regiões
df_reg_in = pd.DataFrame({'Nome da Região': ['REG_SUL_1', 'REG_SUL_2'], 'CEP Inicial': ['80010000', '84500000'], 'CEP Final': ['80010000', '84500000'], 'Codigo IBGE': ['', '4110706']})
df_reg_out = pd.DataFrame({'Região (Aba localizacoes_atendidas)': ['REG_SUL_1', 'REG_SUL_2'], 'CEP Inicial': ['80010000', '84500000'], 'CEP Final': ['80010000', '84500000'], 'Código IBGE da Cidade': ['', '4110706']})
render_table_image(df_reg_in, "1. ENTRADA: Planilha Base com Faixas de CEP e/ou IBGE", "manual_assets/reg_in.png", bg_header='#2b6cb0')
render_table_image(df_reg_out, "2. SAÍDA: Modelo Região Lincros Estruturado", "manual_assets/reg_out.png", highlight_cols=['Região (Aba localizacoes_atendidas)', 'CEP Inicial', 'CEP Final'], bg_header='#2f855a')

print('Diagramas visuais gerados com sucesso!')
