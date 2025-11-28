# Produção Industrial (PIM/IBGE) — Pipeline de Coleta, Tratamento e Análise.

Este repositório apresenta um pipeline completo para **coleta, tratamento, dessazonalização e análise dos dados da Produção Física Industrial (PIM-PF)** do IBGE, com foco no Brasil, Santa Catarina e nos principais setores industriais.

O projeto está estruturado em três scripts principais:
1. Coleta de dados - coleta de dados
2. Tratamento de dados - impeza, padronização, dessazonalização e construção das categorias econômicas
3. Criação de relatório - cálculo das variações econômicas (MoM, YoY, acumulado, 12 meses, trimestre móvel)

**Objetivo Geral**

Este projeto automatiza todo o processo de preparação dos dados da produção industrial, permitindo:
- Coleta atualizada via SIDRA/IBGE
- Padronização de bases históricas
- Dessazonalização via X13-ARIMA-SEATS
- Construção das categorias econômicas oficiais
- Cálculo dos principais indicadores econômicos
- Geração de tabelas e séries para análises

**Estrutura dos Scripts**
**Script 1 — Coleta dos Dados**

Responsável por:
- Coletar dados atualizados via sidrapy (Brasil e Santa Catarina)

**Script 2 — Tratamento dos Dados**

Responsável por:
- Importar bases brutas
- Limpar colunas, identificar valores numéricos e converter formatos
- Dessazonalizar séries usando X13-ARIMA-SEATS
- Calcular categorias econômicas com pesos oficiais
- Salvar todas as séries tratadas em um único arquivo Excel

**Principais etapas**:

1. **Função auxiliar**: Identifica valores realmente numéricos para conversão.

2. **Limpeza e filtragem**: Remove strings, espaços, colunas vazias e converte para **float**.

3. **Dessazonalização (X13-ARIMA-SEATS)**: A dessazonalização é aplicada nas séries: PIM SC, Setores industriais e Categorias Econômicas.

4. **Cálculo das categorias econômicas**

Com base nos pesos oficiais do IBGE:
- Bens de Consumo, Intermediários e de capital.

Por fim, o script salva todos os dados gerados num arquivo excel com as seguintes abas:
- PIM
- PIM (Sazonal)
- Categorias Econômicas
- Categorias Econômicas (Sazonal)
- PIM Estados
- PIM Brasil
- PIM Brasil (Sazonal)

**Script 3 — Análise dos Dados**

Responsável por:
- Calcular as variações econômicas:
    - Mensal (MoM)
    - Interanual (YoY)
    - Acumulado no ano (YTD)
    - Acumulado em 12 meses
    - Trimestre Móvel Interanual
- Gerar tabelas
- Preparar séries para gráficos (Brasil x SC)

**Função principal**

**calcular_variacoes()**
- Identifica colunas numéricas automaticamente
- Organiza a série temporal por data
- Calcula todas as variações
- Retorna somente o ponto mais recente

**Função gráfica**

**calcular_variacoes_series()**

