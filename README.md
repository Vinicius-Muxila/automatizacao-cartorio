## 📊 Automatização de Faturamento Anual de Empresas no Cartório

Este projeto tem como objetivo automatizar a análise de faturamento anual das empresas mensalistas atendidas pelo cartório onde trabalho. Através de um script Python, os dados mensais são consolidados em uma única planilha anual, facilitando o acompanhamento, análise e apresentação dos dados.

## ✨ Funcionalidades

- 📥 Leitura automática de 12 planilhas mensais (`planilha_jan.xlsx`, `planilha_fev.xlsx`, ..., `planilha_dez.xlsx`)
- 📊 Consolidação do faturamento mensal e anual de cada cliente
- 📈 Cálculo de:
  - Faturamento total anual por cliente
  - Faturamento médio mensal
  - Quantidade de meses com movimentação
- 📑 Geração de estatísticas gerais:
  - Total de empresas
  - Quantas empresas estão acima da média de R$ 300,00 por mês
  - Percentual de empresas acima/abaixo da média
  - Faturamento total geral
- 💾 Exportação para planilha Excel com duas abas:
  - **Estatísticas**
  - **Faturamento Consolidado**
- 🎨 Aplicação de formatação profissional no Excel, incluindo cabeçalho com cor, bordas, alinhamento e formatação monetária

## 🛠 Tecnologias utilizadas

- [Python 3.x](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/) – manipulação de dados
- [OpenPyXL](https://openpyxl.readthedocs.io/) – leitura, escrita e formatação de arquivos Excel

## 📁 Estrutura de entrada

O script espera encontrar 12 arquivos `.xlsx` no mesmo diretório, com os seguintes nomes:
planilha_jan.xlsx
planilha_fev.xlsx e assim, por diante.

cada planilha deve conter ao menos duas colunas:

- `CLIENTE` (nome da empresa, em letras maiúsculas)
- `VALOR` (valor faturado no mês)

## 📤 Saída

Um arquivo Excel chamado `faturamento_anual.xlsx`, contendo:

- **Aba "Faturamento Consolidado"**: Dados por cliente com valores mensais, totais, média e quantidade de meses com faturamento.
- **Aba "Estatísticas"**: Informações agregadas e indicadores úteis para análise.

## 🧠 Motivação

No cartório onde atuo, o controle de faturamento das empresas mensalistas era feito de forma manual e repetitiva, com grande chance de erros e retrabalho. Com esta automação, ganhamos tempo, reduzimos erros e temos uma visão muito mais clara dos dados ao final de cada ano.




