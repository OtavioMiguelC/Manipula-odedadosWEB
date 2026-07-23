import matplotlib.pyplot as plt
import os

os.makedirs('manual_assets', exist_ok=True)

def render_excel_sheet(title, headers, data, filename, highlights={}, title_color='#107c41'):
    fig, ax = plt.subplots(figsize=(11, len(data)*0.55 + 1.8))
    ax.axis('off')
    
    plt.title(title, fontsize=13, fontweight='bold', pad=15, color=title_color)
    
    num_cols = len(headers)
    col_labels = [chr(65 + i) for i in range(num_cols)]
    
    full_table = [headers] + data
    row_labels = [str(r + 1) for r in range(len(full_table))]
    
    table = ax.table(
        cellText=full_table,
        colLabels=col_labels,
        rowLabels=row_labels,
        loc='center',
        cellLoc='center'
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(9.5)
    table.scale(1.15, 1.45)
    
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#cbd5e1')
        
        if r == 0 and c == -1:
            cell.set_facecolor('#e2e8f0')
            cell.get_text().set_text('')
        elif r == 0:
            cell.set_facecolor('#e2e8f0')
            cell.get_text().set_weight('bold')
            cell.get_text().set_color('#334155')
        elif c == -1:
            cell.set_facecolor('#e2e8f0')
            cell.get_text().set_weight('bold')
            cell.get_text().set_color('#334155')
        elif r == 1:
            cell.set_facecolor(title_color)
            cell.get_text().set_color('white')
            cell.get_text().set_weight('bold')
        else:
            data_r = r - 2
            data_c = c
            if (data_r, data_c) in highlights:
                cell.set_facecolor(highlights[(data_r, data_c)])
                cell.get_text().set_weight('bold')
                cell.get_text().set_color('#854d0e' if highlights[(data_r, data_c)]=='#fef08a' else '#064e3b')
            else:
                cell.set_facecolor('#f8fafc' if data_r % 2 == 1 else '#ffffff')

    plt.tight_layout()
    plt.savefig(filename, bbox_inches='tight', dpi=200)
    plt.close()

# 1. Modelo Base (Analista)
render_excel_sheet(
    "FIGURA 1: Modelo Base Preenchido pelo Analista (Base_de_Origem_Template.xlsx)",
    ["Nome da Região", "Destino", "UF Destino", "CEP Inicial", "CEP Final", "Prazo", "Codigo IBGE", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"],
    [
        ["REG_SUL_CURITIBA", "Curitiba", "PR", "80010-000", "80010-000", "2", "4106902", "S", "S", "S", "S", "S", "N"],
        ["REG_SUL_IRATI", "Irati", "PR", "84500000", "84500000", "3", "4110706", "S", "N", "S", "N", "S", "N"],
        ["REG_RS_SANTIAGO", "Santiago", "RS", "97700000", "97700000", "4", "4317400", "S", "S", "N", "S", "N", "N"]
    ],
    "manual_assets/fig01_base_analista.png",
    highlights={(0,3): '#fef08a', (0,4): '#fef08a', (1,3): '#fef08a', (1,4): '#fef08a', (2,3): '#fef08a', (2,4): '#fef08a'},
    title_color='#1e3a8a'
)

# 2. Cadastro CEP - Entrada Analista
render_excel_sheet(
    "FIGURA 2: Entrada de Dados pelo Analista na Aba Cadastro CEP",
    ["CEP Inicial (Colado / Planilha)", "CEP Final (Opcional)"],
    [
        ["80010-000", "80010-000"],
        ["97700000", "97700000"],
        ["84500000", "84500000"],
        ["01001-000", "01001-000"]
    ],
    "manual_assets/fig02_cep_entrada.png",
    title_color='#1e3a8a'
)

# 3. Cadastro CEP - Retorno Sistema Lincros
render_excel_sheet(
    "FIGURA 3: Retorno do Sistema - Modelo CEP Lincros (Modelo CEP Preenchido.xlsx)",
    ["ID Localização", "CEP Inicial", "CEP Final", "Nome (Cidade - UF)", "Ativo"],
    [
        ["", "80010000", "80010000", "Curitiba - PR", "VERDADEIRO"],
        ["", "97700000", "97700000", "Santiago - RS", "VERDADEIRO"],
        ["", "84500000", "84500000", "Irati - PR", "VERDADEIRO"],
        ["", "01001000", "01001000", "São Paulo - SP", "VERDADEIRO"]
    ],
    "manual_assets/fig03_cep_retorno.png",
    highlights={(0,3): '#fef08a', (0,4): '#a7f3d0', (1,3): '#fef08a', (1,4): '#a7f3d0', (2,3): '#fef08a', (2,4): '#a7f3d0', (3,3): '#fef08a', (3,4): '#a7f3d0'},
    title_color='#065f46'
)

# 4. IBGE - Entrada Analista
render_excel_sheet(
    "FIGURA 4: Planilha Base de Entrada Sem Códigos IBGE",
    ["Destino (Cidade)", "UF Destino", "Codigo IBGE"],
    [
        ["Curitiba", "PR", ""],
        ["São Paulo", "SP", ""],
        ["Porto Alegre", "RS", ""]
    ],
    "manual_assets/fig04_ibge_entrada.png",
    title_color='#1e3a8a'
)

# 5. IBGE - Retorno Sistema
render_excel_sheet(
    "FIGURA 5: Planilha Base Retornada pelo Sistema com Códigos IBGE Preenchidos",
    ["Destino (Cidade)", "UF Destino", "Codigo IBGE (Gerado)"],
    [
        ["Curitiba", "PR", "4106902"],
        ["São Paulo", "SP", "3550308"],
        ["Porto Alegre", "RS", "4314902"]
    ],
    "manual_assets/fig05_ibge_retorno.png",
    highlights={(0,2): '#fef08a', (1,2): '#fef08a', (2,2): '#fef08a'},
    title_color='#065f46'
)

# 6. Regiões - Aba localizacoes_atendidas
render_excel_sheet(
    "FIGURA 6: Retorno do Sistema - Modelo Região Lincros (Aba: localizacoes_atendidas)",
    ["ID Localização", "Região", "CEP Inicial", "CEP Final", "Código IBGE da Cidade", "Sigla da UF"],
    [
        ["", "REG_SUL_CURITIBA", "80010000", "80010000", "", ""],
        ["", "REG_SUL_IRATI", "84500000", "84500000", "4110706", ""],
        ["", "REG_RS_SANTIAGO", "97700000", "97700000", "4317400", ""]
    ],
    "manual_assets/fig06_regiao_retorno.png",
    highlights={(0,1): '#bae6fd', (0,2): '#fef08a', (0,3): '#fef08a', (1,1): '#bae6fd', (1,2): '#fef08a', (1,3): '#fef08a', (1,4): '#fef08a'},
    title_color='#065f46'
)

# 7. Rotas - Aba Rotas
render_excel_sheet(
    "FIGURA 7: Retorno do Sistema - Matriz de Rotas Lincros (Aba: Rotas)",
    ["Transportadora", "Descrição da Rota", "Origem Cidade", "Origem Região", "Destino Região", "Ativo"],
    [
        ["12345678000100 - LOGISTICA", "REG_SUL_CURITIBA", "4106902", "", "REG_SUL_CURITIBA", "VERDADEIRO"],
        ["12345678000100 - LOGISTICA", "REG_SUL_IRATI", "4106902", "", "REG_SUL_IRATI", "VERDADEIRO"]
    ],
    "manual_assets/fig07_rotas_retorno.png",
    highlights={(0,1): '#fef08a', (0,4): '#fef08a', (1,1): '#fef08a', (1,4): '#fef08a'},
    title_color='#065f46'
)

print('Todas as 7 figuras Excel em alta resolução foram renderizadas com sucesso!')
