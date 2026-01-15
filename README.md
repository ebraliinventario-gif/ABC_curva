# Dashboard de Análise de Curva ABC

Este é um dashboard interativo para análise de Curva ABC desenvolvido em Python usando Streamlit e Plotly.

## Pré-requisitos

- Python 3.7 ou superior
- pip (gerenciador de pacotes do Python)

## Instalação

1. Clone este repositório ou faça o download dos arquivos
2. Instale as dependências executando:
   ```
   pip install -r requirements.txt
   ```

## Como usar

1. Prepare seu arquivo Excel com as seguintes colunas:
   - `descrição`: Nome ou descrição do item
   - `KG`: Quantidade em KG
   - `%individual`: Percentual individual
   - `tipo item`: Tipo do item
   - `%acumulado`: Percentual acumulado

2. Execute o dashboard:
   ```
   streamlit run dashboard_abc.py
   ```

3. Acesse o dashboard no navegador (geralmente abre automaticamente)

4. Faça o upload do seu arquivo Excel usando o uploader na página

## Funcionalidades

- Visualização interativa da Curva ABC
- Classificação automática em categorias A, B e C baseada no percentual acumulado
- Gráfico de barras para percentual individual
- Gráfico de Pareto mostrando percentuais individual e acumulado
- Gráfico de pizza para distribuição das classes ABC
- Filtro por tipo de item
- Estatísticas gerais (total KG, contagem por classe)

## Personalização

Você pode personalizar o dashboard editando o arquivo `dashboard_abc.py`:
- Cores dos gráficos
- Títulos e rótulos
- Limiares das classes A, B e C
- Adicionar mais filtros ou gráficos

## Exemplo de Estrutura do Arquivo Excel

| descrição | KG  | %individual | tipo item | %acumulado |
|-----------|-----|-------------|-----------|------------|
| Item1     | 100 | 20.0        | TipoA     | 20.0       |
| Item2     | 150 | 30.0        | TipoB     | 50.0       |
| ...       | ... | ...         | ...       | ...        |

## Observações

- O dashboard classifica automaticamente os itens em A (até 80%), B (80-95%) e C (acima de 95%) do percentual acumulado
- O arquivo Excel deve ter a extensão .xlsx
- Certifique-se de que os nomes das colunas estejam exatamente como especificado
