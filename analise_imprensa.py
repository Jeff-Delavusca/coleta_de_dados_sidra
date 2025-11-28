# ==========================================================
# 1. Importações
# ==========================================================

import pandas as pd
from pathlib import Path
import xlsxwriter 
from datetime import datetime
import locale
locale.setlocale(locale.LC_TIME, "pt_BR.UTF8")
import matplotlib.pyplot as plt
import numpy as np

# Caminho para ler os dados processados (Ajustar conforme o seu diretório)
PATH_PIM_COMPLETO = Path(r"C:\Projetos\pim_producao_industrial\Dados\Processados\[Boletim] Produção Industrial - Coleta dos Dados.xlsx")

# Caminho para salvar o arquivo Excel com os principais indicadores para imprensa (Ajustar conforme o seu diretório)
PATH_ARQUIVO_IMPRENSA = Path(r"C:\Projetos\pim_producao_industrial\Relatórios\Relatório Imprensa.xlsx")

pim = pd.read_excel(PATH_PIM_COMPLETO, sheet_name="PIM")
pim_sazonal = pd.read_excel(PATH_PIM_COMPLETO, sheet_name="PIM (Sazonal)")
pim_estados = pd.read_excel(PATH_PIM_COMPLETO, sheet_name="PIM Estados")
categorias_econ = pd.read_excel(PATH_PIM_COMPLETO, sheet_name="Categorias Econômicas")
categorias_econ_sazonal = pd.read_excel(PATH_PIM_COMPLETO, sheet_name="Categorias Econômicas (Sazonal)")

# ==========================================================
# 2. Construindo função para os principais indicadores da imprensa
# ==========================================================

def calcular_variacoes(df, coluna_data='Mês', ajustada=False):
    """
    Calcula variações econômicas para todas as colunas numéricas:
      - Var. Mensal (%): apenas se ajustada=True
      - Var. Interanual (%)
      - Var. 12 Meses (%)
      - Var. Acumulada no Ano (%)
      - Var. Trimestre Móvel Interanual (%)
    Retorna o último valor disponível de cada variação.
    """

    df = df.copy()
    df[coluna_data] = pd.to_datetime(df[coluna_data])
    df = df.sort_values(by=coluna_data)
    
    colunas_numericas = df.select_dtypes(include='number').columns
    
    resultados = []
    ultima_data = df[coluna_data].iloc[-1] 
    ano_atual = ultima_data.year

    for col in colunas_numericas:
        # serie = df[col]
        serie = df[col].copy()

        # Apenas se série tiver mais de 1 observação
        var_mensal = None
        if ajustada and len(serie.dropna()) > 1:
            var_mensal = ((serie.iloc[-1] / serie.iloc[-2]) - 1)
        
        # Variação interanual (%)
        var_interanual = None
        if len(serie.dropna()) > 12:
            var_interanual = ((serie.iloc[-1] / serie.shift(12).iloc[-1]) - 1)
        
        # Acumulada em 12 meses (%)
        var_12m = None
        if len(serie.dropna()) > 24:
            soma_atual = serie.iloc[-12:].sum()
            soma_anterior = serie.iloc[-24:-12].sum()
            var_12m = ((soma_atual / soma_anterior) - 1)
        
        # Acumulada no ano (%) (YTD)
        var_ytd = None
        serie_ano = df[df[coluna_data].dt.year == ano_atual][col]
        serie_ano_ant = df[df[coluna_data].dt.year == ano_atual - 1][col]
        if len(serie_ano) > 0 and len(serie_ano_ant) > 0:
            soma_ano = serie_ano.sum()
            soma_ano_ant = serie_ano_ant.iloc[:len(serie_ano)].sum()
            var_ytd = ((soma_ano / soma_ano_ant) - 1)

        # Trimstre Móvel Internaual (%)
        var_tmi = None
        if len(serie.dropna()) > 15: # suficientes para 3 meses + 12 defasagem
            media_3m = serie.rolling(3).mean()
            var_tmi = ((media_3m.iloc[-1] / media_3m.shift(12).iloc[-1]) -1)
        
        resultados.append({
            'Atividades industriais': col,
            'Data atual': ultima_data.strftime('%Y-%m'),
            'Var. Mensal (Sazonal) (%)': var_mensal if ajustada else '—',
            'Var. Interanual (%)': var_interanual,
            'Var. Acumulada no Ano (%)': var_ytd,
            'Var. 12 Meses (%)': var_12m,
            'Var. Trimestre Móvel interanual (%)': var_tmi
        })
    
    df_resultados = pd.DataFrame(resultados)
    return df_resultados

# ==========================================================
# 2.1 Construindo função para os principais indicadores -gráficos
# ==========================================================

def calcular_variacoes_series(df, coluna_data='Mês'):
    df = df.copy()
    df[coluna_data] = pd.to_datetime(df[coluna_data])
    df = df.sort_values(by=coluna_data)

    colunas = [
        'Brasil - Indústria geral',
        'Santa Catarina - Indústria geral'
    ]

    novas_colunas = [coluna_data]

    for col in colunas:
        serie = df[col]

        # Var. Interanual (%)
        var_interanual = f'{col} - Var. Interanual (%)' 
        df[var_interanual] = (serie / serie.shift(12) - 1) *100
        novas_colunas.append(var_interanual)
        
        # Var. 12 Meses (%)
        var_12m = f'{col} - Var. 12 Meses (%)'
        df[var_12m] = (
            serie.rolling(window=12).sum() /
            serie.shift(12).rolling(window=12).sum() - 1
        ) * 100
        novas_colunas.append(var_12m)

        # Var. Acumulada no Ano (%)
        var_ytd = f'{col} - Var. Acumulada no Ano (%)'
        df[var_ytd] = (
            df.groupby(df[coluna_data].dt.year)[col].cumsum()/
            df[col].shift(12).groupby(df[coluna_data].dt.year).cumsum() - 1
        ) * 100

        df[var_ytd] = df[var_ytd].replace([np.inf, -np.inf], np.nan)
        novas_colunas.append(var_ytd)

        # Var. Trimesre Móvel Interanual (%)
        var_tmi = f'{col} - Var. Trimestre Móvel Interanual (%)'
        media_3m = serie.rolling(window=3).mean()
        df[var_tmi] = (media_3m / media_3m.shift(12) - 1) * 100
        novas_colunas.append(var_tmi)

    return df[novas_colunas]

# ==========================================================
# 3. Calculando os principais indicadores da imprensa
# ==========================================================

# Séries sem ajuste sazonal
pim_variacoes = calcular_variacoes(pim, ajustada=False)
categorias_variacoes = calcular_variacoes(categorias_econ, ajustada=False)
pim_estados_variacoes = calcular_variacoes(pim_estados, ajustada=False)
pim_estados_variacoes = pim_estados_variacoes.drop(columns=["Var. Mensal (Sazonal) (%)"]) # coluna em branco

# Séries com ajuste sazonal
pim_sazonal_variacoes = calcular_variacoes(pim_sazonal, ajustada=True)
categorias_sazonal_variacoes = calcular_variacoes(categorias_econ_sazonal, ajustada=True)


# Séries para gráficos
df_graficos = calcular_variacoes_series(pim)
# df_graficos_brasil = calcular_variacoes_series(pim_brasil)

# Excluir as colunas com o mesmo nome nas duas bases antes de fazer o merge
pim_variacoes = pim_variacoes.drop(columns=["Var. Mensal (Sazonal) (%)"])
categorias_variacoes = categorias_variacoes.drop(columns=["Var. Mensal (Sazonal) (%)"])

# Faz o merge na base da PIM para trazer a variação mensal com ajuste sazonal
pim_variacoes = pd.merge(
    pim_variacoes,
    pim_sazonal_variacoes[['Atividades industriais','Var. Mensal (Sazonal) (%)']],
    on = 'Atividades industriais',
    how = 'left'
)

# Alterando a ordem das colunas
pim_variacoes = pim_variacoes[['Atividades industriais', 'Data atual', 
                                     'Var. Mensal (Sazonal) (%)',
                                     'Var. Interanual (%)',
                                     'Var. Acumulada no Ano (%)', 
                                     'Var. 12 Meses (%)',
                                     'Var. Trimestre Móvel interanual (%)']]

# Faz o merge na base das Categorias Econômicas para trazer a variação mensal com ajuste sazonal
categorias_variacoes = pd.merge(
    categorias_variacoes,
    categorias_sazonal_variacoes[['Atividades industriais','Var. Mensal (Sazonal) (%)']],
    on = 'Atividades industriais',
    how = 'left'
)

# Alterando a ordem das colunas
categorias_variacoes = categorias_variacoes[['Atividades industriais', 'Data atual', 
                                     'Var. Mensal (Sazonal) (%)',
                                     'Var. Interanual (%)',
                                     'Var. Acumulada no Ano (%)', 
                                     'Var. 12 Meses (%)',
                                     'Var. Trimestre Móvel interanual (%)']]



# ==========================================================
# 4. Construindo o principais arquivo enviado a imprensa
# ==========================================================

with pd.ExcelWriter(PATH_ARQUIVO_IMPRENSA, engine="xlsxwriter") as writer:

    workbook = writer.book
    
    # ============================ Início da criação e formatação da aba Sumário ==========================================================================

    # Cria uma aba chamada "Sumário"
    worksheet_sumario = workbook.add_worksheet("Sumário")

    # Define a orientação da página e margens do Sumário
    worksheet_sumario.set_landscape()   # formato papel na horizontal
    worksheet_sumario.set_paper(9)      # tipo de formato de folha A4
    worksheet_sumario.set_margins(left = 0.5, right = 0.5, top = 0.5, bottom = 0.5)     # Definição de margens

    # Ajusta o zoom da aba Sumário
    worksheet_sumario.set_zoom(150)

    # Formato para título principal do Sumário
    formato_titulo_sumario = workbook.add_format({
        'bold': True,
        'font_size': 22,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': 'black'
    })

    # Formato para subtítulos do Sumário
    formato_subtitulo = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': 'black'
    })

    # Formato para cabeçalhos do Sumário
    formato_cabecalho_sumario = workbook.add_format({
        'bold': True,
        'font_size': 15,
        'align': 'left',        # alinha à esquerda
        'valign': 'vcenter',
        'font_color': 'black',
        'bottom': 1             # adiciona borda inferior e tipo de borda (1)
    })

    # Formato para texto comum do Sumário
    formato_texto = workbook.add_format({
        'text_wrap': True,      # Permite quebra de linha
        'font_size': 12,
        'text_wrap': True,
        'valign': 'vcenter',    # centraliza verticalmente
        'align': 'left',
        'bottom': 1 
    })

    # Formato cinza clarinho do Sumário
    formato_cinza = workbook.add_format({'bg_color': '#D9D9D9'})

    # Preenchendo o restante da aba com cinza, exceto a parte branca onde ficarão os textos
    for row in range(0, 50):
        for col in range(0, 15):
            # evita sobrescrever o bloco branco
            if not (1 <= row <= 12 and 1 <= col <= 7):
                worksheet_sumario.write_blank(row, col, None, formato_cinza)

    # Dscreve o título do sumário
    worksheet_sumario.write('C3', 'Produção Industrial Mensal (PIM)', formato_titulo_sumario)

    # Descreve o subtítulo  do sumário com o último mês de referência do arquivo
    mes_referencia = pd.to_datetime(pim_variacoes['Data atual']).max().strftime('%B/%Y').capitalize()
    titulo = f"Mês de referência - {mes_referencia}"
    worksheet_sumario.write('C4', titulo, formato_subtitulo)
    
    # Descreve os cabeçalhos do sumário
    worksheet_sumario.write('C6', 'Planilhas', formato_cabecalho_sumario)
    worksheet_sumario.write('D6', 'Descrição', formato_cabecalho_sumario)

    # Texto explicativo do sumário
    # Define largura das colunas
    worksheet_sumario.set_column('C:C', 30)  # Nome da planilha
    worksheet_sumario.set_column('D:D', 65)  # Texto descritivo

    # Primeira linha de conteúdo
    worksheet_sumario.write('C7', 'PIM', formato_texto)
    worksheet_sumario.write('D7',
        'Dados referentes à PIM de Santa Catarina divididos por CNAE. '
        'Na planilha, encontram-se os dados históricos disponíveis no '
        'Instituto Brasileiro de Geografia e Estatística (IBGE), com '
        'variações que permitem a análise do desempenho da produção '
        'industrial setorial em diferentes períodos de tempo.',
        formato_texto
    )

    # Segunda linha de conteúdo
    worksheet_sumario.write('C8', 'Categorias Econômicas', formato_texto)
    worksheet_sumario.write('D8',
        'Dados referentes às categorias econômicas: Bens de Consumo, '
        'Bens Intermediários e Bens de Capital. É possível encontrar a '
        'série histórica, com variações que permitem a análise do desempenho '
        'das grandes categorias econômicas ao longo do tempo.',
        formato_texto
    )
    
    # Inseri as imagens na aba Sumário
    worksheet_sumario.insert_image(
        'E3', 
        r"C:\Projetos\pim_producao_industrial\Dados\Brutos\logo_iel.jpg",  
        {'x_scale': 1, 'y_scale': 0.6})
    
    worksheet_sumario.insert_image(
        'F3', 
        r"C:\Projetos\pim_producao_industrial\Dados\Brutos\logo_fiesc.png",  
        {'x_scale': 1, 'y_scale': 0.6, 'x_offset': 30 })
    
    worksheet_sumario.insert_image(
        'E4', 
        r"C:\Projetos\pim_producao_industrial\Dados\Brutos\logo_observatorio.jpg",  
        {'x_scale': 0.8, 'y_scale': 0.8, 'x_offset': 30, 'y_offset': 10})
    
    # ============================ Fim da formatação do Sumário ==========================================================================

    # ================ Início das formatações das abas PIM e Categorias Econômicas =======================================================
    
    # Escreve os nomes das abas
    pim_variacoes.to_excel(writer, sheet_name="PIM", index=False, startrow=3)
    categorias_variacoes.to_excel(writer, sheet_name="Categorias Econômicas", index=False, startrow=3)
    pim_estados_variacoes.to_excel(writer, sheet_name="PIM Estados", index=False, startrow=3)
    #df_graficos.to_excel(writer, sheet_name="Séries para Gráficos", index=False, startrow=3)

    # Recupera as abas criadas
    worksheet_pim = writer.sheets["PIM"]
    worksheet_cat = writer.sheets["Categorias Econômicas"]
    worksheet_pim_estados = writer.sheets["PIM Estados"]
    #worksheet_graficos = writer.sheets["Séries para Gráficos"]

    # Define o zoom individualmente para as abas PIM e Categorias Econômicas
    worksheet_pim.set_zoom(110)   
    worksheet_cat.set_zoom(130)
    worksheet_pim_estados.set_zoom(120)
    #worksheet_graficos.set_zoom(110)  


    # Retira a linha de grade de cada aba do arquivo
    for aba in writer.sheets:
        ws = writer.sheets[aba]
        ws.hide_gridlines(2)

    # Formatos
    formato_titulo = workbook.add_format({
        'bold': True,
        'font_size': 18,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#1F497D',
        'font_color': 'white'
    })

    formato_cabecalho = workbook.add_format({
        'bold': True, 'font_size': 12 ,'bg_color': '#0C769E', 'font_color': 'white',
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    formato_percentual = workbook.add_format({
        'num_format': '0.0%', 'align': 'center', 'valign': 'vcenter'
    })

    formato_data = workbook.add_format({
        'num_format': 'yyyy-mm', 'align': 'center', 'valign': 'vcenter'
    })

    # Função para formatar cada aba
    def formatar_aba(df, aba_nome):
        worksheet = writer.sheets[aba_nome]
        num_colunas = len(df.columns)
        
        # Título
        mes_ano = pd.to_datetime(df['Data atual']).max().strftime('%B/%Y').capitalize()
        titulo = f"{aba_nome} - Resultados de ({mes_ano})"
        worksheet.merge_range(2, 0, 0, num_colunas-1, titulo, formato_titulo)
        
        # Cabeçalho
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, formato_cabecalho)
        
        # Aumenta a altura da linha de cabeçalho
        worksheet.set_row(3, 25)

        # Largura + formatos
        for i, col in enumerate(df.columns):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 4
            if "(%)" in col:
                worksheet.set_column(i, i, width, formato_percentual)
            elif "Data" in col:
                worksheet.set_column(i, i, width, formato_data)
            else:
                worksheet.set_column(i, i, width)

        # Congela cabeçalho
        worksheet.freeze_panes(4, 0)

        # Formatação condicional colunas %
        for col_num, col_name in enumerate(df.columns):
            if "(%)" in col_name:
                start_row = 4
                end_row = len(df) + 4
                col_letter = chr(65 + col_num)
                cell_range = f"{col_letter}{start_row + 1}:{col_letter}{end_row}"

                # positivo → verde
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0,
                    'format': workbook.add_format({'bg_color': '#C6EFCE', 
                                                   'font_color': '#000000', 
                                                   'num_format': '0.0%'})
                })
                # negativo → vermelho
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0,
                    'format': workbook.add_format({'bg_color': '#FFC7CE',
                                                   'font_color': '#000000',
                                                   'num_format': '0.0%'})
                })

                # Borda inferir após o final da tabela
                linha_final = len(df) + 4 # última linha da tabela
                formato_borda_inferior = workbook.add_format({'bottom': 1})
                for col_name in range(num_colunas):
                    worksheet.write_blank(linha_final, col_name, None, formato_borda_inferior)

                # Legenda explicativa
                # Legenda explicativa
                worksheet.write(len(df) + 5, 0, "Legenda:", workbook.add_format({'bold': True, 'font_size': 9}))
                worksheet.write(len(df) + 6, 0, "Var. Mensal (%): compara o mês atual com o mês anterior, com ajuste sazonal.", workbook.add_format({'font_size': 8}))
                worksheet.write(len(df) + 7, 0, "Var. Interanual (%): compara o mês atual com o mesmo mês do ano anterior.", workbook.add_format({'font_size': 8}))
                worksheet.write(len(df) + 8, 0, "Var. 12 Meses (%): mostra a variação dos últimos 12 meses em relação aos 12 meses anteriores.", workbook.add_format({'font_size': 8}))
                worksheet.write(len(df) + 9, 0, "Var. Acumulada no Ano (%): mostra a variação acumulada desde janeiro até o mês atual, em relação ao mesmo período do ano anterior.", workbook.add_format({'font_size': 8}))
                worksheet.write(len(df) + 10, 0, "Var. Trimestre Móvel interanual (%): mostra a variação dos últimos três meses em comparação com os mesmos três meses do ano anterior.", workbook.add_format({'font_size': 8}))


    # Aplica em ambas as abas
    formatar_aba(pim_variacoes, "PIM")
    formatar_aba(categorias_variacoes, "Categorias Econômicas")
    formatar_aba(pim_estados_variacoes, "PIM Estados")

    # ============================ Início da criação e formatação da aba Gráficos ==========================================================================

    worksheet_graficos = writer.book.add_worksheet("Gráficos Indústria Geral - SC")

    titulo = "Séries de Variações Percentuais - Indústria Geral"
    worksheet_graficos.merge_range(0, 0, 2, 21, titulo, formato_titulo)
    worksheet_graficos.set_zoom(95)
    worksheet_graficos.freeze_panes(3, 0) # congela cabeçalho


    # Gráfico 1 - Var. Interanual (%)
    fig1, ax1 = plt.subplots()
    #ax1.plot(df_graficos['Mês'], df_graficos['Brasil - Indústria geral - Var. Interanual (%)'], label = 'Brasil')
    ax1.plot(df_graficos['Mês'], df_graficos['Santa Catarina - Indústria geral - Var. Interanual (%)'], label = 'Santa Catarina')
    ax1.set_title('Variação Interanual (%) - Indústria Geral - Santa Catarina')
    ax1.legend()
    fig1.tight_layout()
    fig1.savefig('grafico_interanual.png') # salvar imagem temporária
    plt.close(fig1)

    # Inserir no Excel
    worksheet_graficos.insert_image('A5', 'grafico_interanual.png', {'x_scale': 1, 'y_scale': 0.7})
    

    # Gráfico 2 - Var. 12 Meses (%)
    fig2, ax2 = plt.subplots()
    ax2.plot(df_graficos['Mês'], df_graficos['Santa Catarina - Indústria geral - Var. 12 Meses (%)'], label = 'Santa Catarina')
    ax2.set_title("Variação 12 meses (%) - Indústria Geral - Santa Catarina")
    ax2.legend()
    fig2.tight_layout()
    fig2.savefig('grafico_12_meses.png') # salvar imagem temporária
    plt.close(fig2)

    # Inserir no Excel
    worksheet_graficos.insert_image('A23', 'grafico_12_meses.png', {'x_scale': 1, 'y_scale': 0.7})
   

    # Gráfico 3 - Var. Acumulada no Ano (%)
    fig3, ax3 = plt.subplots()
    ax3.plot(df_graficos['Mês'], df_graficos['Santa Catarina - Indústria geral - Var. Acumulada no Ano (%)'], label = 'Santa Catarina')
    ax3.set_title("Variação Acumulada no Ano (%) - Indústria Geral - Santa Catarina")
    ax3.legend()
    fig3.tight_layout()
    fig3.savefig('grafico_acumulada_ano.png') # salvar imagem temporária
    plt.close(fig3)
    
    # Inserir no Excel
    worksheet_graficos.insert_image('L5', 'grafico_acumulada_ano.png', {'x_scale': 1, 'y_scale': 0.7})
    
    # Gráfico 4 - Var. Trimestre Móvel Interanual (%)
    fig4, ax4 = plt.subplots()
    ax4.plot(df_graficos['Mês'], df_graficos['Santa Catarina - Indústria geral - Var. Trimestre Móvel Interanual (%)'], label = 'Santa Catarina')
    ax4.set_title("Variação Trim. Móvel Interanual (%) - Indústria Geral - Santa Catarina")
    ax4.legend()
    fig4.tight_layout()
    fig4.savefig('grafico_trimestre_movel_interanual.png') # salvar imagem temporária
    plt.close(fig4)

    # Inserir no Excel
    worksheet_graficos.insert_image('L23', 'grafico_trimestre_movel_interanual.png', {'x_scale': 1, 'y_scale': 0.7})

    # Retira a linha de grade da aba de gráficos
    worksheet_graficos.hide_gridlines(2)


    # ================================ Início da criação da aba de Ranking dos Estados ========================================================
 
    # ================================ Ranking da variação interanual dos estados
    
    # Criar uma aba de ranking dos Estados 
    worksheet_ranking = workbook.add_worksheet("Ranking Estados")

    # Adiciona o título, o zoom, congela o cabeçalho e retira a linha de grade
    titulo = "Ranking dos Estados - Indústria Geral"
    worksheet_ranking.merge_range(0, 0, 2, 10, titulo, formato_titulo)
    worksheet_ranking.set_zoom(100)
    worksheet_ranking.freeze_panes(3, 0) # congela cabeçalho
    worksheet_ranking.hide_gridlines(2) # retira a linha de grade da aba de ranking

    # Preparar os dataframes para o ranking
    df_ranking_interanual = pim_estados_variacoes[['Atividades industriais', 'Var. Interanual (%)']].copy()
    df_ranking_acm_12_meses = pim_estados_variacoes[['Atividades industriais', 'Var. 12 Meses (%)']].copy()
    df_ranking_acm_ano = pim_estados_variacoes[['Atividades industriais', 'Var. Acumulada no Ano (%)']].copy()

    # Calcular o ranking para cada variação
    df_ranking_interanual['Posição'] = df_ranking_interanual['Var. Interanual (%)'].rank(ascending=False, method='min').astype(int)
    df_ranking_acm_12_meses['Posição'] = df_ranking_acm_12_meses['Var. 12 Meses (%)'].rank(ascending=False, method='min').astype(int)
    df_ranking_acm_ano['Posição'] = df_ranking_acm_ano['Var. Acumulada no Ano (%)'].rank(ascending=False, method='min').astype(int)

    # Ordenar os dataframes pelo ranking
    df_ranking_interanual = df_ranking_interanual.sort_values('Posição')
    df_ranking_acm_12_meses = df_ranking_acm_12_meses.sort_values('Posição')
    df_ranking_acm_ano = df_ranking_acm_ano.sort_values('Posição')

    # Fromatação do cabeçalho
    cabecalho_formatado = workbook.add_format({
        'bold': True,
        'bg_color': '#0C769E',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Escreve os cabeçalhos das tabelas de ranking
    cabecalho = ['Estado', 'Var. Interanual (%)', 'Posição']
    for col, titulo in enumerate(cabecalho):
        worksheet_ranking.write(3, col, titulo, cabecalho_formatado)
    
    cabecalho2 = ['Estado', 'Var. 12 Meses (%)', 'Posição']
    for col, titulo in enumerate(cabecalho2):
        worksheet_ranking.write(3, col + 4, titulo, cabecalho_formatado)

    cabecalho3 = ['Estado', 'Var. Acumulada no Ano (%)', 'Posição']
    for col, titulo in enumerate(cabecalho3):
        worksheet_ranking.write(3, col + 8, titulo, cabecalho_formatado)

    # Define a largura das colunas e aplica os formatos
    for col in [1, 5, 9]:
        worksheet_ranking.set_column(col, col, 25, formato_percentual)  # Percentual

    for col in [0, 4, 8]:
        worksheet_ranking.set_column(col, col, 35)  # Estado

    for col in [2, 6, 10]:
        worksheet_ranking.set_column(col, col, 12, workbook.add_format({'align': 'center'}))  # Ranking

    # Preenche as tabelas de ranking com os dados
    for row_num, row in enumerate(df_ranking_interanual.itertuples(), start=4):
        worksheet_ranking.write(row_num, 0, row[1])  # Estado
        worksheet_ranking.write(row_num, 1, row[2])  # Valor percentual
        worksheet_ranking.write(row_num, 2, row[3])  # Ranking

    for row_num, row in enumerate(df_ranking_acm_12_meses.itertuples(), start=4):
        worksheet_ranking.write(row_num, 4, row[1])  # Estado
        worksheet_ranking.write(row_num, 5, row[2])  # Valor percentual
        worksheet_ranking.write(row_num, 6, row[3])  # Ranking

    for row_num, row in enumerate(df_ranking_acm_ano.itertuples(), start=4):
        worksheet_ranking.write(row_num, 8, row[1])  # Estado
        worksheet_ranking.write(row_num, 9, row[2])  # Valor percentual
        worksheet_ranking.write(row_num, 10, row[3])  # Ranking

    # Destacar linha da tabela onde o estado é Santa Catarina - Indústria geral
    formato_sc = workbook.add_format({
        'bg_color': 'A9ECEF',  
        'bold': True
    })

    # Nome alvo
    estado_sc = "Santa Catarina - Indústria geral"

    # Total de linhas em cada rank
    num_linhas = len(df_ranking_interanual)

    # Tabela interanual
    worksheet_ranking.conditional_format(
        f"A5:C{num_linhas+4}",
        {
            'type': 'text',
            'criteria': 'containing',
            'value': estado_sc,
            'format': formato_sc
        }
    )

    # Tabela 12 meses
    worksheet_ranking.conditional_format(
        f"E5:G{num_linhas+4}",
        {
            'type': 'text',
            'criteria': 'containing',
            'value': estado_sc,
            'format': formato_sc
        }
    )

    # Tabela acumulada no ano
    worksheet_ranking.conditional_format(
        f"I5:K{num_linhas+4}",
        {
            'type': 'text',
            'criteria': 'containing',
            'value': estado_sc,
            'format': formato_sc
        }
    )
 