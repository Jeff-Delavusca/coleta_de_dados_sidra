# ==========================================================
# 1. Importações
# ==========================================================

import pandas as pd
import sidrapy
from pathlib import Path

# Caminho para os arquivos de dados brutos coletados (Ajustar conforme necessário)
PATH_PIM = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\df_pim_brutos.xlsx") 
PATH_PIM_ESTADOS = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\df_pim_brutos.xlsx") 
PATH_PIM_BRASIL = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\df_pim_brutos.xlsx") 
PATH_PIM_BRASIL_SA = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\df_pim_brutos.xlsx") 
PESOS_PATH = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\Categorias Econômicas - Pesos PIM.xlsx")


# Caminho para salvar o arquivo Excel com os dados processados (Ajustar conforme necessário)
PATH_SALVAR = Path(r"C:\Projetos\pim_producao_industrial\Dados\Processados\[Boletim] Produção Industrial - Coleta dos Dados.xlsx")

df_pim = pd.read_excel(PATH_PIM, sheet_name = 'PIM', index_col = 0, parse_dates = True)
df_pim_estados = pd.read_excel(PATH_PIM_ESTADOS, sheet_name = 'PIM Estados', index_col = 0, parse_dates = True)
df_pim_brasil = pd.read_excel(PATH_PIM_BRASIL, sheet_name = 'PIM Brasil', index_col = 0, parse_dates = True)
df_pim_brasil_sa = pd.read_excel(PATH_PIM_BRASIL, sheet_name = 'PIM Brasil (Sazonal)', index_col = 0, parse_dates = True)


# ==========================================================
# 2. Função auxiliar para converter um valor em número
# ==========================================================
def is_number(x):
    """Verifica se o valor pode ser convertido em número."""
    try:
        float(x)
        return True
    except ValueError:
        return False

# ==========================================================
# 3. Conversão para númerico da base PIM de Brasil, SC e estados
# ==========================================================
df_pim = (
    df_pim[df_pim.applymap(is_number)]
    .astype(float)
    .dropna(axis=1, how="all")
)
print("   DataFrame finalizado com shape:", df_pim.shape)

df_pim_estados = (
    df_pim_estados[df_pim_estados.applymap(is_number)]
    .astype(float)
    .dropna(axis=1, how="all")
)
print("   DataFrame finalizado com shape:", df_pim_estados.shape)

df_pim_brasil = (
    df_pim_brasil[df_pim_brasil.applymap(is_number)]
    .astype(float)
    .dropna(axis=1, how="all")
)
print("   DataFrame finalizado com shape:", df_pim_brasil.shape)

df_pim_brasil_sa = (
    df_pim_brasil_sa[df_pim_brasil_sa.applymap(is_number)]
    .astype(float)
    .dropna(axis=1, how="all")
)
print("   DataFrame finalizado com shape:", df_pim_brasil_sa.shape)


# Selecionar somente as colunas desejadas do DataFrame df_pim_brasil

colunas_desejadas = ['Brasil - 10 Fabricação de produtos alimentícios',
       'Brasil - 13 Fabricação de produtos têxteis',
       'Brasil - 14 Confecção de artigos do vestuário e acessórios',
       'Brasil - 16 Fabricação de produtos de madeira',
       'Brasil - 17 Fabricação de celulose, papel e produtos de papel',
       'Brasil - 20 Fabricação de produtos químicos',
       'Brasil - 22 Fabricação de produtos de borracha e de material plástico',
       'Brasil - 23 Fabricação de produtos de minerais não metálicos',
       'Brasil - 24 Metalurgia',
       'Brasil - 25 Fabricação de produtos de metal, exceto máquinas e equipamentos',
       'Brasil - 27 Fabricação de máquinas, aparelhos e materiais elétricos',
       'Brasil - 28 Fabricação de máquinas e equipamentos',
       'Brasil - 29 Fabricação de veículos automotores, reboques e carrocerias',
       'Brasil - 31 Fabricação de móveis']

df_pim_brasil = df_pim_brasil[colunas_desejadas]
df_pim_brasil_sa = df_pim_brasil_sa[colunas_desejadas]


# ========================================================
# 4. Coleta dados Brasil e SC com Ajuste Sazonal (já vem pronto do SIDRA)
# ========================================================

df_pim_as = pd.DataFrame()

# Dicionário com os parâmetros para coleta de dados do SIDRA: Brasil e Santa Catarina
DIC = {
    "Brasil": {"level": "1", "code": "all", "classifications": "129314,129315,129316"},
    "Santa Catarina": {"level": "3", "code": "42", "classifications": "129314"},
}

for region in DIC.keys():
    print(f"\nColetando dados de {region} (ajuste sazonal)...")

    sidra_pim = sidrapy.get_table(
        table_code="8888",
        territorial_level=DIC[region]["level"],
        ibge_territorial_code=DIC[region]["code"],
        variable="12607",  # índice ajustado sazonalmente
        period="all",
        classifications={"544": DIC[region]["classifications"]},
        header="n",
    )

    sections = list(sidra_pim["D4N"].drop_duplicates())

    for section in sections:
        serie = pd.DataFrame(
            sidra_pim.loc[sidra_pim["D4N"] == section]["V"].reset_index(drop=True)
        )
        serie.index = pd.date_range("2002-01-01", periods=len(serie), freq="MS")
        # column_name = f"{region} - {section[2:]} (Ajuste Sazonal)"
        column_name = f"{region} - {section[2:]}"
        df_pim_as[column_name] = serie
        print("Adicionada:", column_name)

df_pim_as.index.name = "Mês"

df_pim_as = df_pim_as[df_pim_as.applymap(is_number)].astype(float).dropna(axis=1, how="all")

# =========================================================
# 5. Dessazonalizando os setores induastriais do PIM de SC
# ========================================================

import os
import warnings
from statsmodels.tsa.x13 import x13_arima_analysis, X13Warning

os.environ["X13PATH"] = r"C:\Projetos\pim_producao_industrial\Dados\Brutos\x13as"
warnings.filterwarnings("ignore", category=X13Warning)

df_pim_as_completa = pd.DataFrame()

series_as = list(df_pim.columns)[4:]

for serie in series_as:
    try:
        ajuste = x13_arima_analysis(endog=df_pim[serie].dropna(), trading=True)
    except Exception as e:
        print(f"Caught an exception: {type(e).__name__}, {e}")
        continue

    # df_pim_as[serie + ' (Ajuste Sazonal)'] = pd.DataFrame(ajuste.seasadj)
    df_pim_as[serie] = pd.DataFrame(ajuste.seasadj)
    df_pim_as_completa[serie + ' (Irregular)'] = pd.DataFrame(ajuste.irregular)
    df_pim_as_completa[serie + ' (Tendência)'] = pd.DataFrame(ajuste.trend)
    if round(df_pim_as_completa[serie + ' (Irregular)'].mean(),0) == 1:
        # df_pim_as_completa[serie + ' (Sazonal)'] = df_pim[serie]/df_pim_as[serie + ' (Ajuste Sazonal)']
        df_pim_as_completa[serie] = df_pim[serie]/df_pim_as[serie]
    else:
        # df_pim_as_completa[serie + ' (Sazonal)'] = df_pim[serie] - df_pim_as[serie + ' (Ajuste Sazonal)']
        df_pim_as_completa[serie ] = df_pim[serie] - df_pim_as[serie]
    print(serie)

# ==========================================================
# 6. Cálculo das categorias econômicas (com pesos)
# ==========================================================
print("\nCalculando categorias econômicas...")

# Lê o arquivo de pesos
categorias_ecn = pd.read_excel(PESOS_PATH).set_index("Categoria Econômica")

# Filtra o df_pim para pegar só as colunas de setores
df_pim_filtered = df_pim.iloc[:, 4:]

# Ajusta o nome das colunas dos pesos para coincidir com df_pim_filtered
categorias_ecn.columns = list(df_pim_filtered.columns)

# Cria DataFrame vazio com mesmo índice
df_cat_ecn = pd.DataFrame(index=df_pim_filtered.index)

# Lista das categorias
colunas = [
    "Bens de Consumo Duráveis",
    "Bens de Consumo Não Duráveis",
    "Bens de Consumo",
    "Bens Intermediários",
    "Bens de Capital",
]

# Cálculo vetorizado
for col in colunas:
    df_cat_ecn[col] = df_pim_filtered.mul(categorias_ecn.loc[col], axis=1).sum(axis=1)

print("Categorias econômicas calculadas com sucesso!")

# ==========================================================
# 7. Dessazonalizando as grandes categorias econômicas
# =========================================================

df_cat_ecn_as = pd.DataFrame()

series_as = list(df_cat_ecn.columns)

for serie in series_as:
    try:
        ajuste = x13_arima_analysis(endog=df_cat_ecn[serie].dropna(), trading=True)
    except Exception as e:
        print(f"Caught an exception: {type(e).__name__}, {e}")
        continue
    # df_cat_ecn_as[serie + ' (Ajuste Sazonal)'] = pd.DataFrame(ajuste.seasadj)
    df_cat_ecn_as[serie] = pd.DataFrame(ajuste.seasadj)
    print(serie)

# ==========================================================
# 8. Salvamento dos resultados
# ==========================================================

with pd.ExcelWriter(PATH_SALVAR) as write:
    df_pim.to_excel(write, sheet_name = 'PIM')
    df_pim_as.to_excel(write, sheet_name = 'PIM (Sazonal)')
    df_cat_ecn.to_excel(write, sheet_name = 'Categorias Econômicas')
    df_cat_ecn_as.to_excel(write, sheet_name = 'Categorias Econômicas (Sazonal)')
    df_pim_estados.to_excel(write, sheet_name = 'PIM Estados')
    df_pim_brasil.to_excel(write, sheet_name = 'PIM Brasil')
    df_pim_brasil_sa.to_excel(write, sheet_name = 'PIM Brasil (Sazonal)')