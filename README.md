# AnaliseVendas
 Análise de Vendas com Python
Script Python que lê dados de vendas em CSV, realiza análises com Pandas e gera automaticamente um relatório Excel formatado com múltiplas abas e indicadores executivos.
 Motivação
Empresas lidam diariamente com planilhas de vendas que precisam ser consolidadas e analisadas. Esse projeto automatiza esse processo: com um único comando, os dados brutos viram um relatório profissional pronto para apresentação.

⚙️ Funcionalidades

Leitura e processamento de dados CSV com Pandas
Cálculo automático de faturamento total por linha de venda
Análises agrupadas por vendedor, categoria, mês e região
Geração de relatório .xlsx com 6 abas organizadas
Formatação automática: cabeçalhos coloridos, linhas alternadas, bordas e largura de coluna ajustada
Aba de Resumo Executivo com os principais KPIs (indicadores)


 Estrutura do projeto
analise-vendas/
│
├── vendas.csv           # Base de dados de entrada
├── analise_vendas.py    # Script principal
├── README.md            # Este arquivo
│
└── relatorio_vendas.xlsx  # Gerado ao rodar o script
    ├── Resumo           ← KPIs executivos
    ├── Dados Brutos     ← Todos os registros
    ├── Por Vendedor     ← Ranking + ticket médio
    ├── Por Categoria    ← Faturamento por produto
    ├── Por Mes          ← Evolução mensal
    └── Por Regiao       ← Desempenho regional

 Como usar
1. Clone o repositório
bashgit clone https://github.com/Aramaki-nicolas/AnaliseVendas.git
cd AnaliseVendas
2. Instale as dependências
bashpip install pandas openpyxl
3. Execute o script
bashpython analise_vendas.py
O arquivo relatorio_vendas.xlsx será gerado na mesma pasta.

Formato esperado do CSV
O arquivo vendas.csv deve conter as seguintes colunas:
ColunaTipoExemplodatadata2024-01-15vendedortextoAna LimaprodutotextoNotebookcategoriatextoEletronicosquantidadeinteiro2preco_unitariodecimal3500.00regiaotextoNordeste

Para usar com seus próprios dados, substitua o vendas.csv mantendo essas colunas.


Tecnologias utilizadas
BibliotecaUsopandasLeitura do CSV, agrupamentos e análisesopenpyxlFormatação visual do Excel gerado

Conceitos aplicados

Manipulação de DataFrames com Pandas (groupby, agg, sort_values)
Criação de arquivos Excel com múltiplas abas (pd.ExcelWriter)
Estilização de planilhas com openpyxl (cores, fontes, bordas)
Boas práticas: funções reutilizáveis, código comentado, separação de responsabilidades
