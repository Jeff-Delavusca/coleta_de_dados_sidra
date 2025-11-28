# ==========================================================
# 1. Importações
# ==========================================================
import pandas as pd
import sidrapy
from pathlib import Path

# ==========================================================
# 2. Parâmetros gerais
# ==========================================================

# Caminho para salvar o arquivo Excel com os dados coletados (Sempre ajustar conforme necessário)
PATH_SALVAR = Path(r"C:\Projetos\pim_producao_industrial\Dados\Brutos\df_pim_brutos.xlsx")

# Dicionário com os parâmetros para coleta de dados do SIDRA: Brasil e Santa Catarina
DIC = {
    "Brasil": {"level": "1", "code": "all", "classifications": "129314,129315,129316"},
    "Santa Catarina": {"level": "3", "code": "42", "classifications": "129314"},
}

# Dicionário com os parâmetros para coleta de dados do SIDRA: somente Brasil
DIC_BR = {
    "Brasil": {"level": "1", "code": "all", "classifications": "129314"}
}

# Dicionário com os códigos IBGE dos estados brasileiros
ESTADOS_IBGE = {
    "Rondônia": "11", "Acre": "12", "Amazonas": "13", "Roraima": "14", "Pará": "15", "Amapá": "16", "Tocantins": "17", "Maranhão": "21",
    "Piauí": "22", "Ceará": "23", "Rio Grande do Norte": "24", "Paraíba": "25", "Pernambuco": "26", "Alagoas": "27", "Sergipe": "28",
    "Bahia": "29", "Minas Gerais": "31", "Espírito Santo": "32", "Rio de Janeiro": "33", "São Paulo": "35", "Paraná": "41",
    "Santa Catarina": "42", "Rio Grande do Sul": "43", "Mato Grosso do Sul": "50", "Mato Grosso": "51", "Goiás": "52", "Distrito Federal": "53",
}

# Dicionário para coleta de dados do SIDRA para todos os estados brasileiros
DIC2 = {}

# População do DIC2 com os estados e seus respectivos códigos IBGE
for estado, codigo in ESTADOS_IBGE.items():
    DIC2[estado] = {"level": "3", "code": codigo, "classifications": "129314"}

# ==========================================================
# 3. Coleta de dados do SIDRA somente de Brasil e SC (sem classificações)
# ==========================================================
df_pim = pd.DataFrame()

for region in ["Brasil", "Santa Catarina"]:
    print(f"\nColetando dados de {region}...")

    sidra_pim = sidrapy.get_table(
        table_code="8888",
        territorial_level=DIC[region]["level"],
        ibge_territorial_code=DIC[region]["code"],
        variable="12606",
        period="all",
        classifications={"544": DIC[region]["classifications"]},
        header="n",
    )

    for section in sidra_pim["D4N"].drop_duplicates():
        ind = pd.DataFrame(
            sidra_pim.loc[sidra_pim["D4N"] == section]["V"].reset_index(drop=True)
        )
        ind.index = pd.date_range("2002-01-01", periods=len(ind), freq="MS")
        column_name = f"{region} - {section[2:]}"
        df_pim[column_name] = ind
        print("Adicionada:", column_name)

df_pim.index.name = "Mês"

# ==========================================================
# 4. Coleta de dados do SIDRA de todos os estado (sem classificações)
# ==========================================================

df_pim_estados = pd.DataFrame()

for estados in DIC2.keys():
    print(f"\nColetando dados de {estados}...")

    sidra_pim = sidrapy.get_table(
        table_code="8888",
        territorial_level=DIC2[estados]["level"],
        ibge_territorial_code=DIC2[estados]["code"],
        variable="12606",
        period="all",
        classifications={"544": DIC2[estados]["classifications"]},
        header="n",
    )

    for section in sidra_pim["D4N"].drop_duplicates():
        ind = pd.DataFrame(
            sidra_pim.loc[sidra_pim["D4N"] == section]["V"].reset_index(drop=True)
        )
        ind.index = pd.date_range("2002-01-01", periods=len(ind), freq="MS")
        column_name = f"{estados} - {section[2:]}"
        df_pim_estados[column_name] = ind
        print("Adicionada:", column_name)

df_pim_estados.index.name = "Mês"

# ----------------------------------------------------------
# 5. Coleta detalhada de dados das classificações de SC
# ----------------------------------------------------------

print("\nColetando dados detalhados de Santa Catarina (todas as classificações)...")

sidra_pim_sc = sidrapy.get_table(
    table_code="8888",
    territorial_level="3",
    ibge_territorial_code="42",
    variable="12606",
    period="all",
    classifications={"544": "all"},
    header="n",
)

series = list(sidra_pim_sc["D4N"].drop_duplicates())[3:]

for serie in series:
    pim = pd.DataFrame(
        sidra_pim_sc.loc[sidra_pim_sc["D4N"] == serie]["V"].reset_index(drop=True)
    )
    pim.index = df_pim.index
    df_pim[serie] = pim
    print("Adicionada série extra:", serie)

# ----------------------------------------------------------
# 6. Coleta detalhada de dados das classificações de Brasil (Sazonal)
# ----------------------------------------------------------

df_pim_br_as = pd.DataFrame()

for item in DIC_BR.keys():
    print(f"\nColetando dados de {item}...")

    sidra_pim = sidrapy.get_table(
        table_code="8888",
        territorial_level=DIC_BR[item]["level"],
        ibge_territorial_code=DIC_BR[item]["code"],
        variable="12607",
        period="all",
        classifications={"544": "all"},
        header="n",
    )

    for section in sidra_pim["D4N"].drop_duplicates():
        ind = pd.DataFrame(
            sidra_pim.loc[sidra_pim["D4N"] == section]["V"].reset_index(drop=True)
        )
        ind.index = pd.date_range("2002-01-01", periods=len(ind), freq="MS")
        column_name = f"{item} - {section[2:]}"
        df_pim_br_as[column_name] = ind
        print("Adicionada:", column_name)

df_pim_br_as.index.name = "Mês"

# ----------------------------------------------------------
# 7. Coleta detalhada de dados das classificações de Brasil
# ----------------------------------------------------------

df_pim_br = pd.DataFrame()

for item in DIC_BR.keys():
    print(f"\nColetando dados de {item}...")

    sidra_pim = sidrapy.get_table(
        table_code="8888",
        territorial_level=DIC_BR[item]["level"],
        ibge_territorial_code=DIC_BR[item]["code"],
        variable="12606",
        period="all",
        classifications={"544": "all"},
        header="n",
    )

    for section in sidra_pim["D4N"].drop_duplicates():
        ind = pd.DataFrame(
            sidra_pim.loc[sidra_pim["D4N"] == section]["V"].reset_index(drop=True)
        )
        ind.index = pd.date_range("2002-01-01", periods=len(ind), freq="MS")
        column_name = f"{item} - {section[2:]}"
        df_pim_br[column_name] = ind
        print("Adicionada:", column_name)

df_pim_br.index.name = "Mês"

# ----------------------------------------------------------
# 8. Salvamento dos dados coletados
# ----------------------------------------------------------

with pd.ExcelWriter(PATH_SALVAR) as write:
    df_pim.to_excel(write, sheet_name = 'PIM')
    df_pim_estados.to_excel(write, sheet_name = 'PIM Estados')
    df_pim_br_as.to_excel(write, sheet_name = 'PIM Brasil (Sazonal)')
    df_pim_br.to_excel(write, sheet_name = 'PIM Brasil')