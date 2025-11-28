# Coleta automatizada de dados do Sistema IBGE de Recuperação Automática (SIDRA).
Nesse exemplo, estarei coletando dados da Produção Física Industrial Geral (PIM) de Santa Catarina, bem como dos subsetores que a compõem.
Para isso, o material está organizando em três scripts:
1. Coleta de dados
2. Tratamento de dados
3. Criação de relatório

**Coleta de dados**

O script de coleta de dados define, inicialmente, os parâmetros de local para o armazenamento dos dados coletados (que devem ser alterados conforme o usuário), seguido dos parâmetros para a coleta dos dados separados por três dicionários, sendo: um para coletar os dados para o Brasil e Santa Catarina, dois para coletar os dados somente para o Brasil e três para coletar os dados para todos as Unidades Federativas.

Para realizar a integração com o sistema do IBGE, foi utilizado o pacote sidrapy.

Dentro do primeiro dicionário, tem-se duas chaves, a primeira é referente ao **Brasil** com os seguintes parâmetros:
- **level: 1**, para dados a nível nacional (Brasil)
- **code: all**, (para todos os estados)
- **classifications: 129314** (Indústria geral), **129315** (Indústrias extrativas), **129316** (Indústrias de transformação), são as atividades industriais.

A segunda chave refere-se ao **Santa Catarina**, tem os seguintes parâmetros:
- **level: 3**, para definir o nível estadual
- **code: 42**, define o código do estado
- **classifications: 129314** (Indústria geral), define a atividade industrial

No segundo dicionário, tem-se os parâmetros para o **Brasil**.

No terceiro dicionário, tem os parâmetros para identificar cada estado na coleta de dados.

Após a identificação, são feitas alguns loops e laços de repetição para coletar os dados, armazená-los em um dicionário vazio e transformar esse dicionário em um dataframe.

Por fim, o scritp de coleta salvo esses dados em um arquivo em excel (xlsx) separados por aba.
