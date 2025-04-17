# 👩🏻‍💻 Emanuelly Araújo

**`Analista de Dados`**

Sou uma profissional apaixonada por transformar dados em insights estratégicos, com uma trajetória que une formações e experiências complementares. Minha base acadêmica em Ciências Econômicas e Administração, obtida na UFPR, aliada a um MBA em Finanças e à especialização em Data Science, me proporciona uma visão única e integrada dos negócios. Essa combinação não convencional me permite entender, de forma holística, tanto os desafios operacionais quanto as oportunidades estratégicas que os dados oferecem. Ao longo da minha carreira, tenho atuado na transformação digital de processos, integrando sistemas e criando soluções inovadoras que conectam áreas diversas e promovem uma cultura orientada por dados.

---

### 🤖 Linguagens e Tecnologias

<img 
    align="left" 
    alt="mySQL"
    title="mySQL" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/mysql/mysql-original-wordmark.svg" 
  />
          

<img 
    align="left" 
    alt="SQL"
    title="SQL" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/azuresqldatabase/azuresqldatabase-original.svg" />

<img 
    align="left" 
    alt="Postman"
    title="Postman" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/postman/postman-original.svg"        
  />

<img 
    align="left" 
    alt="dbeaver"
    title="dbeaver" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/dbeaver/dbeaver-original.svg" 
  />
          
  
<img 
    align="left" 
    alt="AWS"
    title="AWS" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/amazonwebservices/amazonwebservices-original-wordmark.svg" 
  />
          

<img 
    align="left" 
    alt="Python" 
    title="Python"
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/python/python-original.svg" 
/>

<img 
    align="left" 
    alt="anaconda"
    title="anaconda" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/anaconda/anaconda-original-wordmark.svg" 
  />
          

<br/>
<br/>


# Projetos Desenvolvidos

# 📊 Power BI - Análise de Vendas e Previsão de Demanda
🚀 Projeto desenvolvido para análise de vendas, estoque e previsão de demanda de uma empresa do ramo de Alimentos.

## 📌 Objetivo do Projeto
Criar um dashboard interativo no Power BI para:
- ✔ Analisar vendas e desempenho de produtos.
- ✔ Monitorar giro de estoque e prever demandas futuras.
- ✔ Gerar insights estratégicos para otimização de supply chain.

## 📂 Estrutura do Projeto
- 📁 dados/ → Base de dados em Excel utilizada no projeto.
- 📁 scripts/ → Códigos Python e DAX para manipulação e análise.
- 📁 dashboard/ → Arquivo .pbix com o relatório final no Power BI.
- 📄 README.md → Documentação completa do projeto.

🐍 Código Python
Codigo utilizado para criar uma base de dados fictícia estruturada para 4 anos, contendo as tabelas “Fatos_Vendas”, “Produtos”, “Calendário” e “Canais_Regionais”. Esses dados foram elaborados para simular um ambiente real e possibilitar análises robustas no Power BI.

```python
import pandas as pd
from datetime import datetime, timedelta

# Create sample data for the sheets

# 1. Fatos_Vendas DataFrame
data_fatos = [
    {"Data": "05/01/2022", "Ano": 2022, "Mês": "Janeiro", "Produto_ID": 439, "Produto": "Leite em Pó Integral 200g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "B2B", "Região": "Curitiba", "Volume_Vendas": 1000, "Receita (R$)": 50000, "Custo (R$)": 35000, "CMV (%)": 70, "EBITDA (R$)": 15000, "Novo_Lançamento": "Não"},
    {"Data": "15/02/2022", "Ano": 2022, "Mês": "Fevereiro", "Produto_ID": 956, "Produto": "Leite em Pó Inst. 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "Varejo", "Região": "Porto Alegre", "Volume_Vendas": 800, "Receita (R$)": 40000, "Custo (R$)": 28000, "CMV (%)": 70, "EBITDA (R$)": 12000, "Novo_Lançamento": "Sim"},
    {"Data": "10/03/2022", "Ano": 2022, "Mês": "Março", "Produto_ID": 397, "Produto": "Leite em Pó Inst. Vitaminado 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "FoodService", "Região": "São José do Rio", "Volume_Vendas": 650, "Receita (R$)": 32500, "Custo (R$)": 22750, "CMV (%)": 70, "EBITDA (R$)": 9750, "Novo_Lançamento": "Não"},
    {"Data": "22/06/2022", "Ano": 2022, "Mês": "Junho", "Produto_ID": 364, "Produto": "Queijo Azul (Capa Azul 2,5Kg)", "Categoria": "Queijos", "Subcategoria": "Queijo Azul", "Canal": "Atacado", "Região": "Curitiba", "Volume_Vendas": 120, "Receita (R$)": 36000, "Custo (R$)": 25200, "CMV (%)": 70, "EBITDA (R$)": 8400, "Novo_Lançamento": "Sim"},
    {"Data": "05/09/2022", "Ano": 2022, "Mês": "Setembro", "Produto_ID": 978, "Produto": "Queijo Parmesão Arroyo Cabral", "Categoria": "Queijos", "Subcategoria": "Parmesão", "Canal": "B2B", "Região": "Foz do Iguaçu", "Volume_Vendas": 90, "Receita (R$)": 45000, "Custo (R$)": 31500, "CMV (%)": 70, "EBITDA (R$)": 10500, "Novo_Lançamento": "Não"},
    {"Data": "15/11/2023", "Ano": 2023, "Mês": "Novembro", "Produto_ID": 979, "Produto": "Queijo Cremont 7Kg", "Categoria": "Queijos", "Subcategoria": "Cremont", "Canal": "Varejo", "Região": "Curitiba", "Volume_Vendas": 75, "Receita (R$)": 37500, "Custo (R$)": 26250, "CMV (%)": 70, "EBITDA (R$)": 8750, "Novo_Lançamento": "Sim"},
    {"Data": "08/01/2023", "Ano": 2023, "Mês": "Janeiro", "Produto_ID": 956, "Produto": "Leite em Pó Inst. 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "B2B", "Região": "Porto Alegre", "Volume_Vendas": 1200, "Receita (R$)": 60000, "Custo (R$)": 42000, "CMV (%)": 70, "EBITDA (R$)": 18000, "Novo_Lançamento": "Não"},
    {"Data": "20/04/2023", "Ano": 2023, "Mês": "Abril", "Produto_ID": 364, "Produto": "Queijo Azul (Capa Azul 2,5Kg)", "Categoria": "Queijos", "Subcategoria": "Queijo Azul", "Canal": "FoodService", "Região": "São José do Rio", "Volume_Vendas": 110, "Receita (R$)": 33000, "Custo (R$)": 23100, "CMV (%)": 70, "EBITDA (R$)": 8700, "Novo_Lançamento": "Não"},
    {"Data": "12/07/2024", "Ano": 2024, "Mês": "Julho", "Produto_ID": 397, "Produto": "Leite em Pó Inst. Vitaminado 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "Atacado", "Região": "Curitiba", "Volume_Vendas": 950, "Receita (R$)": 47500, "Custo (R$)": 33250, "CMV (%)": 70, "EBITDA (R$)": 11750, "Novo_Lançamento": "Sim"},
    {"Data": "28/10/2024", "Ano": 2024, "Mês": "Outubro", "Produto_ID": 978, "Produto": "Queijo Parmesão Arroyo Cabral", "Categoria": "Queijos", "Subcategoria": "Parmesão", "Canal": "Varejo", "Região": "Foz do Iguaçu", "Volume_Vendas": 85, "Receita (R$)": 42500, "Custo (R$)": 29750, "CMV (%)": 70, "EBITDA (R$)": 10750, "Novo_Lançamento": "Não"},
    {"Data": "03/03/2025", "Ano": 2025, "Mês": "Março", "Produto_ID": 979, "Produto": "Queijo Cremont 7Kg", "Categoria": "Queijos", "Subcategoria": "Cremont", "Canal": "B2B", "Região": "Curitiba", "Volume_Vendas": 100, "Receita (R$)": 50000, "Custo (R$)": 35000, "CMV (%)": 70, "EBITDA (R$)": 12000, "Novo_Lançamento": "Sim"},
    {"Data": "18/05/2025", "Ano": 2025, "Mês": "Maio", "Produto_ID": 956, "Produto": "Leite em Pó Inst. 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Canal": "FoodService", "Região": "Porto Alegre", "Volume_Vendas": 1100, "Receita (R$)": 55000, "Custo (R$)": 38500, "CMV (%)": 70, "EBITDA (R$)": 14500, "Novo_Lançamento": "Não"},
    {"Data": "30/08/2025", "Ano": 2025, "Mês": "Agosto", "Produto_ID": 364, "Produto": "Queijo Azul (Capa Azul 2,5Kg)", "Categoria": "Queijos", "Subcategoria": "Queijo Azul", "Canal": "Atacado", "Região": "São José do Rio", "Volume_Vendas": 130, "Receita (R$)": 39000, "Custo (R$)": 27300, "CMV (%)": 70, "EBITDA (R$)": 11700, "Novo_Lançamento": "Sim"},
]
df_fatos = pd.DataFrame(data_fatos)

# 2. Produtos DataFrame
data_produtos = [
    {"Produto_ID": 439, "Produto": "Leite em Pó Integral 200g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Data_Lancamento": "2020-03-15", "Marca": "Hecke", "Preço_Unitário (R$)": 0.05},
    {"Produto_ID": 956, "Produto": "Leite em Pó Inst. 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Data_Lancamento": "2021-07-10", "Marca": "Hecke", "Preço_Unitário (R$)": 0.06},
    {"Produto_ID": 397, "Produto": "Leite em Pó Inst. Vitaminado 350g", "Categoria": "Laticínios", "Subcategoria": "Leite em Pó", "Data_Lancamento": "2022-01-20", "Marca": "Hecke", "Preço_Unitário (R$)": 0.07},
    {"Produto_ID": 364, "Produto": "Queijo Azul (Capa Azul 2,5Kg)", "Categoria": "Queijos", "Subcategoria": "Queijo Azul", "Data_Lancamento": "2021-05-05", "Marca": "Hecke", "Preço_Unitário (R$)": 30.00},
    {"Produto_ID": 978, "Produto": "Queijo Parmesão Arroyo Cabral", "Categoria": "Queijos", "Subcategoria": "Parmesão", "Data_Lancamento": "2020-09-25", "Marca": "Hecke", "Preço_Unitário (R$)": 50.00},
    {"Produto_ID": 979, "Produto": "Queijo Cremont 7Kg", "Categoria": "Queijos", "Subcategoria": "Cremont", "Data_Lancamento": "2022-11-15", "Marca": "Hecke", "Preço_Unitário (R$)": 70.00},
]
df_produtos = pd.DataFrame(data_produtos)

# 3. Calendário DataFrame
# Creating a date range from 01/01/2022 to 31/12/2025
date_range = pd.date_range(start="2022-01-01", end="2025-12-31", freq="D")
calendario = pd.DataFrame({"Data": date_range})
calendario["Ano"] = calendario["Data"].dt.year
calendario["Mês"] = calendario["Data"].dt.month
calendario["Nome_Mês"] = calendario["Data"].dt.strftime("%B")
calendario["Trimestre"] = calendario["Data"].dt.to_period("Q").astype(str)

# 4. Canais_Regionais DataFrame
data_canais = [
    {"Canal": "B2B", "Região": "Curitiba", "Meta_Vendas (unid)": 1200, "Responsável Regional": "João da Silva"},
    {"Canal": "Varejo", "Região": "Porto Alegre", "Meta_Vendas (unid)": 1000, "Responsável Regional": "Maria Oliveira"},
    {"Canal": "FoodService", "Região": "São José do Rio", "Meta_Vendas (unid)": 900, "Responsável Regional": "Carlos Pereira"},
    {"Canal": "Atacado", "Região": "Foz do Iguaçu", "Meta_Vendas (unid)": 800, "Responsável Regional": "Fernanda Souza"},
]
df_canais = pd.DataFrame(data_canais)

# Write to Excel file with multiple sheets
output_file = "BaseHeckeAlimentos.xlsx"
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df_fatos.to_excel(writer, sheet_name="Fatos_Vendas", index=False)
    df_produtos.to_excel(writer, sheet_name="Produtos", index=False)
    calendario.to_excel(writer, sheet_name="Calendário", index=False)
    df_canais.to_excel(writer, sheet_name="Canais_Regionais", index=False)

output_file
````
📊 Fórmulas DAX Utilizadas

1️⃣ Receita bruta de vendas
```DAX
ReceitaBruta = 
SUMX(
    'vendas',
    'vendas'[Quantidade_Vendida] * 'vendas'[Preco_Venda(R$)]
)

````
2️⃣ Margem de lucro

```DAX
Margem Lucro = 
AVERAGEX(
    'vendas',
    IF(
        [Preco_Venda(R$)] > 0,
        ( [Preco_Venda(R$)] - [Preco_Custo(R$)] - [Custo_Logistico(R$)] ) / [Preco_Venda(R$)] * 100,
        BLANK()
    )
)
````

3️⃣ Volume de Vendas

```DAX
Volume Vendas = SUM(vendas[Quantidade_Vendida])
````

4️⃣ CMV

```DAX
CMV = 
SUM(vendas[Preco_Custo(R$)]) + SUM(vendas[Custo_Logistico(R$)])
````

5️⃣ Custo Logistico Total

```DAX
CustoLogísticoTotal = SUMX('vendas', 'vendas'[Quantidade_Vendida] * 'vendas'[Custo_Logistico(R$)])
````

6️⃣ Custo Logistico Total

```DAX
CustoTotal = 
SUMX(
    'vendas',
    ('vendas'[Quantidade_Vendida] * 'vendas'[Preco_Custo(R$)] + 
    ('vendas'[Quantidade_Vendida] * 'vendas'[Custo_Logistico(R$)]
)))
````
7️⃣ Vendas por cidade

```DAX
Vendas por Região = SUM(vendas[Quantidade_Vendida])
````

8️⃣ Estoque Inicial
- Calcula o estoque inicial de cada SKU baseado na média de vendas, com um fator de ajuste.
```DAX
Estoque Inicial = 
VAR MediaVendas = AVERAGEX(FILTER(ALL('TabelaVendas'), 'TabelaVendas'[SKU] = EARLIER('TabelaVendas'[SKU])), 'TabelaVendas'[Quantidade_Vendida])
VAR FatorAjuste = RAND() * (3 - 1.5) + 1.5
RETURN ROUND(MediaVendas * FatorAjuste, 0)

````
9️⃣ Previsão de Demanda
Calcula a demanda prevista com base na média de vendas, ajustada entre 90% e 120%.

```DAX
Previsao Demanda = 
VAR MediaVendas = AVERAGEX(FILTER(ALL('TabelaVendas'), 'TabelaVendas'[SKU] = EARLIER('TabelaVendas'[SKU])), 'TabelaVendas'[Quantidade_Vendida])
VAR FatorAjuste = RAND() * (1.2 - 0.9) + 0.9
RETURN ROUND(MediaVendas * FatorAjuste, 0)
````

# 📊 Dashboard no Power BI 
- 🔹 Indicadores-Chave (KPIs):
- ✔ Faturamento Total
- ✔ Ticket Médio por Cliente
- ✔ CMV (Custo da Mercadoria Vendida)
- ✔ Margem de Lucro Bruta (%)
- ✔ Giro de Estoque

## 🔹 Principais Análises:
- 📌 Vendas por Categoria e SKU
- 📌 Histórico de Vendas por Mês
- 📌 Projeção de Demanda para os Próximos 6 Meses
- 📌 Análise de Supply Chain (Estoque vs. Demanda)

## 📢 Conclusão
Este projeto fornece uma visão estratégica e acionável para melhorar a gestão de vendas e estoque da Hecke Alimentos.

## 💡 Próximos Passos:
- ✅ Refinar previsões usando Machine Learning (ARIMA, Prophet).
- ✅ Integrar com dados financeiros para calcular EBITDA e CMV detalhado.
- ✅ Implementar automação (RPA) para atualização de dados.

--
# Transformando Dados em Estratégia: Diagnóstico Operacional🚴‍♂️📊

### 🎯 Objetivo do Projeto
Este projeto teve como objetivo realizar uma **análise estratégica baseada em dados operacionais reais de uma empresa do setor esportivo com presença global**.

Desenvolvi um diagnóstico completo com foco em:
- **Identificar gargalos operacionais**
- **Aumentar a rentabilidade da empresa**
- **Gerar planos de ação estratégicos baseados em dados reais**

---

## 📌 Contexto do Case
Recebi o desafio de apresentar um plano estratégico para a diretoria, utilizando dados históricos de vendas, custos, prazos logísticos e performance por produto, categoria e região.

---

## 🧠 Como Abordei o Problema (Minha Metodologia)

### 1. **Entendimento do Negócio**
- Analisei os principais KPIs: Vendas Totais, Margem Bruta R$, Margem %, Custo Total e Prazo Médio de Entrega.
- Relacionei métricas operacionais a metas de negócio: **eficiência, lucratividade e escalabilidade.**

### 2. **Criação de Dashboard Interativo**
- Desenvolvi um dashboard (Power BI) com foco em:
  - Evolução de vendas por período
  - Rentabilidade por produto e categoria
  - Análise de prazos logísticos por país
  - Performance por região geográfica

### 3. **Extração de Insights com Foco em Negócio**
- Identifiquei mais de **R$ 14 milhões** em produtos não rentáveis
- Observei que produtos com **maior margem estão sendo subexplorados**
- Detectei **ineficiências logísticas** que afetam a recompra e a experiência do cliente

---

## 📈 Impacto Estratégico (Resultados do Diagnóstico)

| Área                     | Problema Encontrado                                | Solução Proposta                                             |
|--------------------------|----------------------------------------------------|--------------------------------------------------------------|
| Rentabilidade            | Margem Bruta de apenas 11%                         | Reestruturação do portfólio e política de preços             |
| Produtos Deficitários    | Produtos com margens negativas em circulação       | Descontinuidade ou reprecificação baseada em rentabilidade   |
| Logística Internacional  | Prazo médio de entrega de até 40 dias              | Implementação de hubs logísticos e renegociação com parceiros|
| Oportunidade Comercial   | Alta dependência de categorias de baixa margem     | Cross-selling com produtos de maior margem e menor giro      |
| Regionalização           | Vendas concentradas na América do Norte            | Estratégia de expansão geográfica para Europa e Oceania      |

---

## 📊 Ferramentas Utilizadas

- **Power BI** – Criação de dashboards e visualizações interativas
- **Excel** – Tratamento e organização de dados brutos
- **Storytelling com Dados** – Estrutura de apresentação executiva com foco em tomada de decisão
- **Modelagem de KPI** – Margem Bruta, Prazo Médio, Custo Total, Rentabilidade por Produto

---

## 📘 Entregáveis

- 📎 [Relatório Estratégico Completo - PDF](link-aqui)
- 📊 [Dashboard Interativo (imagem)](link-aqui)
- 📂 Diagnóstico com storytelling para diretoria
- 📈 Plano de ação com foco em impacto financeiro e escalabilidade

---

## 🚀 Aprendizados e Soft Skills Demonstradas

- **Pensamento Estratégico**: Conectei dados operacionais a decisões de alto impacto.
- **Orientação a Resultados**: O foco sempre foi gerar valor para o negócio.
- **Comunicação Executiva**: Entreguei storytelling claro para públicos não técnicos.
- **Domínio de KPIs**: Trabalhei com os indicadores que realmente importam para o negócio.

---

## 💼 Como esse projeto gera valor para empresas?

Este projeto mostra como posso:

- **Apoiar a liderança na tomada de decisão com base em dados reais**
- **Criar relatórios estratégicos acionáveis, não apenas informativos**
- **Elevar a performance de times comerciais, de produto e logística**
- **Fazer diagnósticos críticos e gerar planos estratégicos com ROI positivo**

--
# 📊 Planilha de Controle Financeiro Multimoeda

> Projeto desenvolvido com foco em automação de processos financeiros e aplicação de **Excel avançado** para controle de contratos internacionais com variação cambial.

---

## 💼 Sobre o Projeto

Esta planilha foi construída com o objetivo de automatizar o controle contábil de contratos em moeda estrangeira, otimizando o processo de conversão cambial, geração de lançamentos e análise de saldos com base em regras de negócio pré-definidas.

🔍 **Destaques Técnicos**:
- Fórmulas complexas (ÍNDICE, CORRESP, PROCV, SEERRO, SE, MATRIZ)
- Lógica condicional avançada
- Busca dinâmica por data e tipo de lançamento
- Cálculo de saldos, diferenças e registros contábeis
- Estrutura modular, fácil de escalar e manter

---

## 🧠 Habilidades Demonstradas

- ✅ **Excel Avançado**: fórmulas aninhadas, funções matriciais, manipulação de dados dinâmicos por data
- ✅ **Automação Contábil**: geração de lançamentos com lógica baseada em contas contábeis e variação cambial
- ✅ **Modelagem de Dados**: estruturação das abas como "tabelas relacionais"
- ✅ **Pensamento analítico**: identificação de regras de negócio e conversão em lógica técnica
- ✅ **Documentação Técnica Profissional**

---

## 🗂️ Estrutura da Planilha

### 📌 Abas principais

| Aba             | Função Principal                                           |
|------------------|------------------------------------------------------------|
| `Dólar`          | Taxas de câmbio (compra e venda) por mês                  |
| `Conta Contábil` | Mapeia contratos às contas contábeis                      |
| `Histórico`      | Armazena dados brutos dos contratos                       |
| `Lançamento`     | Geração automática de lançamentos com base na data        |
| `Identificador`  | (Em desenvolvimento) controle de contratos já processados |

---

## 🔢 Fórmulas e Automação (Excel Avançado)

### 🎯 Identificação Dinâmica

```excel
=SEERRO(
   ÍNDICE(Historico!B:B;
     MENOR(
       SE(
         (TEXTO(Historico!A:A;"mm/aaaa") = TEXTO($G$2;"mm/aaaa")) *
         (ÉERROS(Historico!P:P) + (Historico!P:P = ""));
         LIN(Historico!B:B) - LIN(Historico!B$1) + 1;
         ""
       );
       LIN(1:1)
     )
   );
   ""
````
🔍 Busca dinamicamente os contratos ainda não liquidados no mês de referência ($G$2)

🔄 Classificação de Ativo ou Passivo
````
=SE(N2="ADIANTAMENTO";"ATIVO";SE(OU(N2="FORNECEDOR";N2="FINIMP");"PASSIVO";""))
````

🧾 Geração de Histórico Descritivo
excel
````
="FORNECEDOR " & PROCV(B7;Historico!B:Q;8;FALSO) & 
" CONTRATO " & PROCV(B7;Historico!B:Q;3;FALSO) & 
" INVOICE " & PROCV(B7;Historico!B:Q;4;FALSO) & 
" REF USD " & C7 & " TAXA USD " & F7 & 
" = R$ " & ARRED(G7;2)
````

🧾 Cálculo de Contas Contábeis (Débito / Crédito)
````
=SE(
  H7="ADIANTAMENTO";
  SE(F7>D7;
     PROCV(B7;'Conta Contabil'!A:B;2;FALSO);
     2356
  );
  SE(
     OU(H7="FORNECEDOR"; H7="FINIMP");
     SE(F7>D7;
        2356;
        PROCV(B7;'Conta Contabil'!A:B;2;FALSO)
     );
     ""
  )
)

````

## 📈 Performance e Escalabilidade
A planilha processa os dados em até 15 segundos ao mudar o mês de referência (alta carga de fórmulas matriciais)

Estrutura preparada para migração futura para Power Query ou Power BI

Pode ser adaptada para VBA/Macro caso seja necessário integrar com sistemas ou automatizar geração de arquivos


## ✅ Conclusão
Este projeto demonstra a capacidade de aplicar lógica de negócios em Excel avançado, utilizando técnicas de automação, modelagem de dados e otimização de processos.

📎 Ideal para ambientes que ainda utilizam Excel como ferramenta principal de gestão financeira ou contábil, mas com visão de evolução para BI ou ERP.

--

# 📊 Automação Contábil para Importações

Este projeto consiste em macros em VBA que automatizam tarefas contábeis no Excel, otimizando o processo de importação e conciliação de dados contábeis. Através da centralização das informações e da automação de processos, o projeto elimina a fragmentação de dados e reduz significativamente o tempo gasto com tarefas manuais e retrabalhos.

## 🚨 Problema
A contabilização de importações enfrentava diversos desafios, tais como:

- Fragmentação de Dados: Informações dispersas em 71 abas, distribuídas em 4 planilhas distintas, dificultando a localização de dados específicos.
- Processos Manuais Demorados:
- Atualização da taxa de câmbio demandava, em média, 4 horas mensais.
- A busca por aquisições antigas consumia cerca de 20 minutos por operação.
- Retrabalho Acumulado: De janeiro a agosto, foram registradas 14 horas de retrabalho.
- Ausência de Regras Documentadas: Os lançamentos contábeis eram realizados manualmente sem regras formalizadas, o que aumentava o risco de erros.

## ✅ Solução
O projeto implementa melhorias significativas por meio da automação com VBA:
Centralização de Informações: Consolidação dos dados em uma única planilha, permitindo a localização imediata dos registros por meio de filtros.

## 🔧 Automação de Processos:

Macro 1: Automatiza a criação de lançamentos atualizados, realizando a identificação e processamento de registros, além de atualizar automaticamente as taxas de câmbio.

Macro 2: Realiza a conciliação dos lançamentos e gera relatórios prontos para importação no sistema contábil, integrando cálculos, histórico e tratamento de exceções.

Eliminação do Retrabalho: Com a centralização e a automação, todas as atualizações e validações são realizadas de forma rápida e precisa, eliminando a necessidade de processos manuais.

### Macro 1: Criação de lançamentos atualizados automáticos

```vba
Sub CriarLinhas()
    Dim wsMov As Worksheet
    Dim wsPTAX As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim ultimaData As Date
    Dim novaLinha As Long
    Dim identificador As String
    Dim tipoLancamento As String
    Dim existeLiquidado As Boolean
    Dim lançamentosAdicionados As Long
    Dim identificadoresProcessados As Object
    Dim dataExistente As Boolean
    Dim taxaPTAX As Double
    Dim dataBusca As Date
    Dim ptaxUltimaLinha As Long
    Dim encontrado As Boolean
    Dim j As Long
    Dim mesAnoFiltro As String
    Dim mesFiltro As Integer
    Dim anoFiltro As Integer
    Dim mesAnoIdentificador As String

    ' Inicializa o contador de lançamentos adicionados
    lançamentosAdicionados = 0

    ' Definindo as referências para as planilhas
    Set wsMov = ThisWorkbook.Sheets("Movimentações")
    Set wsPTAX = ThisWorkbook.Sheets("PTAX")
    
    ' Criando um dicionário para armazenar identificadores processados
    Set identificadoresProcessados = CreateObject("Scripting.Dictionary")

    ' Obter mês/ano da célula C3
    mesAnoFiltro = wsMov.Range("C3").Value
    If mesAnoFiltro = "" Then
        MsgBox "Por favor, insira um mês/ano válido na célula C3 (MM/YYYY).", vbExclamation
        Exit Sub
    End If
    
    ' Validar formato da data
    If Not IsDate(mesAnoFiltro) Then
        MsgBox "Formato inválido em C3. Use MM/YYYY.", vbExclamation
        Exit Sub
    End If
    
    ' Extrair mês e ano
    mesFiltro = Month(CDate(mesAnoFiltro))
    anoFiltro = Year(CDate(mesAnoFiltro))
    
    ' Calcular último dia do mês informado
    ultimaData = DateSerial(anoFiltro, mesFiltro + 1, 0)

    ' Última linha da planilha Movimentações
    ultimaLinha = wsMov.Cells(wsMov.Rows.Count, 1).End(xlUp).Row

    ' Última linha da planilha PTAX
    ptaxUltimaLinha = wsPTAX.Cells(wsPTAX.Rows.Count, 1).End(xlUp).Row

    ' Loop para verificar os identificadores na coluna A
    For i = 7 To ultimaLinha ' Começar da linha 7
        identificador = wsMov.Cells(i, 1).Value
        
        If Not identificadoresProcessados.exists(identificador) Then
            existeLiquidado = False
            dataExistente = False
            
            ' Verificar ocorrências do identificador
            For j = 7 To ultimaLinha
                If wsMov.Cells(j, 1).Value = identificador Then
                    tipoLancamento = wsMov.Cells(j, 13).Value ' Coluna M
                    
                    ' Verificar liquidação
                    If tipoLancamento = "LIQUIDADO" Then
                        existeLiquidado = True
                        wsMov.Cells(j, 17).Value = "Liquidado"
                    End If
                    
                    ' Verificar se já existe linha para o mês/ano informado
                    mesAnoIdentificador = Format(wsMov.Cells(j, 2).Value, "mm/yyyy")
                    If mesAnoIdentificador = Format(ultimaData, "mm/yyyy") Then
                        dataExistente = True
                    End If
                End If
            Next j
            
            If existeLiquidado Then
                ' Marcar todas as linhas do identificador como Liquidado
                For j = 7 To ultimaLinha
                    If wsMov.Cells(j, 1).Value = identificador Then
                        wsMov.Cells(j, 17).Value = "Liquidado"
                    End If
                Next j
            ElseIf Not dataExistente Then
                ' Criar nova linha
                novaLinha = ultimaLinha + 1
                
                ' Copiar identificador e data
                wsMov.Cells(novaLinha, 1).Value = identificador
                wsMov.Cells(novaLinha, 2).Value = ultimaData
                
                ' Buscar PTAX para a data
                dataBusca = ultimaData
                encontrado = False
                For j = 2 To ptaxUltimaLinha
                    If wsPTAX.Cells(j, 1).Value = dataBusca Then
                        taxaPTAX = wsPTAX.Cells(j, 2).Value
                        encontrado = True
                        Exit For
                    End If
                Next j
                
                ' Preencher taxa PTAX ou mensagem de erro
                wsMov.Cells(novaLinha, 10).Value = IIf(encontrado, taxaPTAX, "Data não encontrada")
                
                ' Copiar demais dados
                wsMov.Range("D" & i & ":I" & i).Copy wsMov.Range("D" & novaLinha)
                wsMov.Cells(novaLinha, 11).Value = wsMov.Cells(i, 11).Value
                wsMov.Cells(novaLinha, 13).Value = wsMov.Cells(i, 13).Value
                wsMov.Cells(novaLinha, 14).Value = wsMov.Cells(i, 14).Value
                
                ultimaLinha = ultimaLinha + 1
                lançamentosAdicionados = lançamentosAdicionados + 1
            End If

            identificadoresProcessados.Add identificador, True
        End If
    Next i

    ' Mensagem final
    If lançamentosAdicionados > 0 Then
        MsgBox lançamentosAdicionados & " lançamento(s) adicionado(s) para " & Format(ultimaData, "MM/yyyy") & "!", vbInformation
    Else
        MsgBox "Nenhum lançamento novo necessário para " & Format(ultimaData, "MM/yyyy") & ".", vbInformation
    End If
End Sub

````

## Macro 2: Conciliação e criação de relatório automático

```vba
Sub FiltrarLancamentosAtualizado()
    Dim wsMov As Worksheet
    Dim wsPTAX As Worksheet
    Dim wsContabil As Worksheet
    Dim wsLanc As Worksheet
    Dim mesAnoFiltro As String
    Dim ultimaLinhaMov As Long
    Dim linhaLanc As Long
    Dim linhaMov As Long
    Dim dataMov As Date
    Dim mesFiltro As Integer
    Dim anoFiltro As Integer
    Dim valorCalculado As Double
    Dim identificador As String
    Dim tipoMov As String
    Dim passivoAtivo As String
    Dim ultimaDataPTAX As Date
    Dim valorDolarAnterior As Double
    Dim valorDolarAtual As Double
    Dim estrutura As String
    Dim tipoLiquidado As String
    Dim historico As String
    Dim somaValorUSD As Double
    Dim taxaCambio As Double
    Dim valorReais As Double
    Dim varDolar As String
    Dim valorContaContabil As String
    Dim contaCadastrada As Boolean
    Dim ultimaLinhaContabil As Long
    Dim lookupRange As Range
    Dim resultadoContabil As Variant

    ' Definir as abas
    Set wsMov = ThisWorkbook.Sheets("Movimentações")
    Set wsLanc = ThisWorkbook.Sheets("Lançamento")
    Set wsPTAX = ThisWorkbook.Sheets("PTAX")
    Set wsContabil = ThisWorkbook.Sheets("Conta Contabil")
    
    ' Obter o mês e ano do filtro na célula H3
    mesAnoFiltro = wsLanc.Range("H3").Value
    If mesAnoFiltro = "" Then
        MsgBox "Por favor, insira um mês e ano válido na célula H3.", vbExclamation
        Exit Sub
    End If
    
    ' Validar formato da data
    If Not IsDate(mesAnoFiltro) Then
        MsgBox "Data no formato incorreto. Use MM/YYYY.", vbExclamation
        Exit Sub
    End If
    
    ' Extrair mês e ano
    mesFiltro = Month(CDate(mesAnoFiltro))
    anoFiltro = Year(CDate(mesAnoFiltro))
    
    ' Identificar última linha da aba Movimentações
    ultimaLinhaMov = wsMov.Cells(wsMov.Rows.Count, "A").End(xlUp).Row
    
    ' Identificar intervalo dinâmico na Conta Contabil
    ultimaLinhaContabil = wsContabil.Cells(wsContabil.Rows.Count, "A").End(xlUp).Row
    Set lookupRange = wsContabil.Range("A2:B" & ultimaLinhaContabil) ' Assume que a linha 1 é cabeçalho
    
    ' Limpar dados existentes
    wsLanc.Rows("7:" & wsLanc.Rows.Count).ClearContents
    
    ' Iniciar preenchimento
    linhaLanc = 7

    For linhaMov = 7 To ultimaLinhaMov
        If IsDate(wsMov.Cells(linhaMov, "B").Value) Then
            dataMov = wsMov.Cells(linhaMov, "B").Value
            
            If Month(dataMov) = mesFiltro And Year(dataMov) = anoFiltro Then
                ' Verificar se está liquidado em M ou Q
                If UCase(Trim(wsMov.Cells(linhaMov, "M").Value)) = "LIQUIDADO" Or _
                   UCase(Trim(wsMov.Cells(linhaMov, "Q").Value)) = "LIQUIDADO" Then
                    GoTo ContinueLoop ' Pula linhas liquidadas
                End If
                
                ' Cálculo do valor
                valorCalculado = (wsMov.Cells(linhaMov, "K").Value + wsMov.Cells(linhaMov, "L").Value) * wsMov.Cells(linhaMov, "J").Value
                
                ' IDENTIFICADOR CORRETO (COLUNA A DA MOVIMENTAÇÕES)
                identificador = Trim(CStr(wsMov.Cells(linhaMov, "A").Value))
                tipoMov = wsMov.Cells(linhaMov, "M").Value
                
                ' Determinar Ativo/Passivo (COLUNA N)
                passivoAtivo = UCase(Trim(wsMov.Cells(linhaMov, "N").Value))
                If passivoAtivo <> "ATIVO" And passivoAtivo <> "PASSIVO" Then
                    Debug.Print "Valor inválido na linha " & linhaMov & ": " & passivoAtivo
                End If
                
                ' Depuração reforçada
                Debug.Print "Processando linha " & linhaMov & _
                    " | ID: " & identificador & _
                    " | Tipo: " & passivoAtivo & _
                    " | Mov: " & tipoMov
                
                ' Cálculo PTAX
                ultimaDataPTAX = WorksheetFunction.EoMonth(CDate(dataMov), -1)
                
                ' Busca valores dólar
                On Error Resume Next
                If passivoAtivo = "ATIVO" Then
                    valorDolarAnterior = Application.VLookup(ultimaDataPTAX, wsPTAX.Range("A:B"), 2, False)
                Else
                    valorDolarAnterior = Application.VLookup(ultimaDataPTAX, wsPTAX.Range("A:C"), 3, False)
                End If
                On Error GoTo 0
                
                ' Tratamento de erros PTAX
                If IsError(valorDolarAnterior) Then
                    valorDolarAnterior = 0
                    Debug.Print "Erro PTAX na linha " & linhaMov
                End If
                
                valorDolarAtual = wsMov.Cells(linhaMov, "J").Value
                If valorDolarAtual = 0 Then Debug.Print "PTAX zero na linha " & linhaMov
                
                ' Determinar variação
                varDolar = IIf(valorDolarAtual > valorDolarAnterior, "aumentou", _
                             IIf(valorDolarAtual < valorDolarAnterior, "diminuiu", "estável"))
                
                ' Definir estrutura
                estrutura = Switch(passivoAtivo = "ATIVO", "132", passivoAtivo = "PASSIVO", "133")
                
                ' Preencher dados básicos
                With wsLanc
                    .Cells(linhaLanc, "A").Value = identificador
                    .Cells(linhaLanc, "B").Value = dataMov
                    .Cells(linhaLanc, "E").Value = valorCalculado
                    .Cells(linhaLanc, "F").Value = varDolar
                    .Cells(linhaLanc, "G").Value = estrutura
                    .Cells(linhaLanc, "H").Value = IIf(estrutura = "133", _
                        "VLR VARIAÇÃO CAMBIAL PASSIVA - PROVISÃO", _
                        "VLR VARIAÇÃO CAMBIAL ATIVA - PROVISÃO")
                End With
                
                ' Montar histórico
                somaValorUSD = wsMov.Cells(linhaMov, "K").Value + wsMov.Cells(linhaMov, "L").Value
                taxaCambio = valorDolarAtual
                valorReais = somaValorUSD * taxaCambio
                
                historico = ""
                With wsMov
                    If .Cells(linhaMov, "I").Value <> "" Then historico = historico & "FORNECEDOR " & .Cells(linhaMov, "I").Value & " "
                    If .Cells(linhaMov, "D").Value <> "" Then historico = historico & "CONTRATO " & .Cells(linhaMov, "D").Value & " "
                    If .Cells(linhaMov, "E").Value <> "" Then historico = historico & "INVOICE " & .Cells(linhaMov, "E").Value & " "
                    If .Cells(linhaMov, "F").Value <> "" Then historico = historico & "DI " & .Cells(linhaMov, "F").Value & " "
                    If .Cells(linhaMov, "G").Value <> "" Then historico = historico & "NF " & .Cells(linhaMov, "G").Value & " "
                    If .Cells(linhaMov, "H").Value <> "" Then historico = historico & "FINIMP " & .Cells(linhaMov, "H").Value & " "
                End With
                
                historico = historico & "REF USD " & Format(somaValorUSD, "#,##0.00") & _
                            " TAXA USD " & Format(taxaCambio, "#,##0.0000") & _
                            " = R$ " & Format(valorReais, "#,##0.00")
                
                wsLanc.Cells(linhaLanc, "I").Value = historico
                
                ' PREENCHIMENTO DAS COLUNAS C E D (BUSCA PELO ID CORRETO)
                contaCadastrada = False
                valorContaContabil = ""
                
                ' Nova lógica de busca com tratamento robusto
                On Error Resume Next ' Ignorar erros temporariamente
                
                ' Verificar se o identificador é numérico
                If IsNumeric(identificador) Then
                    ' Buscar como número
                    resultadoContabil = Application.VLookup(CLng(identificador), lookupRange, 2, False)
                Else
                    ' Buscar como texto
                    resultadoContabil = Application.VLookup(identificador, lookupRange, 2, False)
                End If
                
                On Error GoTo 0 ' Restaurar tratamento de erros
                
                If Not IsError(resultadoContabil) Then
                    valorContaContabil = CStr(resultadoContabil)
                    contaCadastrada = True
                    Debug.Print "Linha " & linhaMov & ": Conta encontrada - " & valorContaContabil
                Else
                    Debug.Print "Linha " & linhaMov & ": Conta NÃO encontrada para ID " & identificador & _
                             " (Tipo: " & TypeName(identificador) & ")"
                End If
                
                ' Lógica complexa para Débito/Crédito
                Select Case passivoAtivo
                    Case "ATIVO"
                        If tipoMov = "ADIANTAMENTO" Then
                            If varDolar = "aumentou" Then
                                wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                wsLanc.Cells(linhaLanc, "D").Value = 2666
                            Else
                                wsLanc.Cells(linhaLanc, "C").Value = 2356
                                wsLanc.Cells(linhaLanc, "D").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                            End If
                        End If
                        
                    Case "PASSIVO"
                        Select Case tipoMov
                            Case "FORNECEDOR"
                                If varDolar = "aumentou" Then
                                    wsLanc.Cells(linhaLanc, "C").Value = 2356
                                    wsLanc.Cells(linhaLanc, "D").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                Else
                                    wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                    wsLanc.Cells(linhaLanc, "D").Value = 2666
                                End If
                                
                            Case "FINIMP"
                                If varDolar = "aumentou" Then
                                    wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                    wsLanc.Cells(linhaLanc, "D").Value = valorContaContabil
                                Else
                                    wsLanc.Cells(linhaLanc, "C").Value = 2666
                                    wsLanc.Cells(linhaLanc, "D").Value = valorContaContabil
                                End If
                        End Select
                End Select
                
                ' Atualizar Movimentações
                wsMov.Cells(linhaMov, "O").Value = wsLanc.Cells(linhaLanc, "C").Value
                wsMov.Cells(linhaMov, "P").Value = wsLanc.Cells(linhaLanc, "D").Value
                
                linhaLanc = linhaLanc + 1
            End If
        End If
ContinueLoop:
    Next linhaMov
    
    MsgBox "Processamento concluído!", vbInformation

    ' CHAMAR A MACRO PARA GERAR O ARQUIVO EXCEL
    Call CriarArquivoExcel
End Sub

````

# 📊 Conciliação Fiscal: Tributário de ICMS e IPI

## 💡 Visão Geral
Este projeto automatiza a conciliação fiscal dos impostos ICMS e IPI no Excel, utilizando macros em VBA para otimizar o processo de importação e verificação de dados contábeis. A automação reduz significativamente o tempo gasto com tarefas manuais e minimiza erros, garantindo maior precisão e controle.

## 🚨 O Problema
Antes da automação, a conciliação fiscal apresentava diversos desafios:

- Processo manual demorado: A conciliação demandava, em média, 2 horas mensais por empresa.
- Retrabalho constante: Para evitar erros e multas, o processo precisava ser revisado por pelo menos duas pessoas, tornando-o moroso e custoso.
- Falta de padronização: Lánçamentos contábeis eram realizados sem regras formalizadas, aumentando o risco de inconsistências.
- Multas e perdas financeiras: Erros frequentes geravam penalizações e custos adicionais.

## ✅ A Solução
Foi desenvolvida uma macro em VBA que automatiza e padroniza a conciliação fiscal, garantindo eficiência e segurança. O código:
- Executa formatação e organização das planilhas de ICMS e IPI, eliminando linhas desnecessárias e ajustando os formatos.
- Consolida e calcula automaticamente os valores de ICMS e IPI na planilha MEMÓRIA, garantindo precisão nos dados.
- Elimina retrabalho ao automatizar validações e atualizações, reduzindo a necessidade de revisões manuais.
- Gera relatórios detalhados, facilitando auditorias e análises.

##🔧 Tecnologias Utilizadas

- VBA (Visual Basic for Applications) para automação no Excel.
- Estruturas de controle e tratamento de erros para garantir a integridade dos dados.
- Proteção de planilhas para evitar edições indevidas.

### Macro : 

```vba
Sub ExecutarTodasMacros()
    '======================================================================
    ' PASSO 1 - Processamento da planilha IPI
    '======================================================================
    Dim ws As Worksheet
    Dim cell As Range
    Dim novoValor As String
    Dim partes() As String
    
    ' Processamento da planilha IPI
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("IPI")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "A planilha 'IPI' não foi encontrada!", vbExclamation
        Exit Sub
    End If
    
    ws.Rows(1).Delete
    
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230)
    End With
    
    For Each cell In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            novoValor = Replace(cell.Value, ".", "")
            
            If InStr(novoValor, "-") > 0 Then
                partes = Split(novoValor, "-")
                
                If UBound(partes) = 1 Then
                    If Len(Trim(partes(1))) = 1 Then
                        novoValor = Trim(partes(0)) & " - 0" & Trim(partes(1))
                    Else
                        novoValor = Trim(partes(0)) & " - " & Trim(partes(1))
                    End If
                End If
            End If
            
            cell.Value = novoValor
            
            If UCase(Trim(cell.Value)) = "TOTAL" Then
                cell.Font.Bold = True
                ws.Range(cell.Offset(0, 1), cell.Offset(0, 5)).Font.Bold = True
            End If
        End If
    Next cell
    
    ws.Columns.ColumnWidth = 15

    '======================================================================
    ' PASSO 2 - Processamento da planilha ICMS
    '======================================================================
    Set ws = ThisWorkbook.Sheets("ICMS")
    ws.Rows(1).Delete
    
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230)
    End With
    
    For Each cell In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            novoValor = Replace(cell.Value, ".", "")
            
            If InStr(novoValor, "-") > 0 Then
                partes = Split(novoValor, "-")
                
                If UBound(partes) = 1 Then
                    If Len(Trim(partes(1))) = 1 Then
                        novoValor = Trim(partes(0)) & " - 0" & Trim(partes(1))
                    Else
                        novoValor = Trim(partes(0)) & " - " & Trim(partes(1))
                    End If
                End If
            End If
            
            cell.Value = novoValor
            
            If UCase(Trim(cell.Value)) = "TOTAL" Then
                cell.Font.Bold = True
                ws.Range(cell.Offset(0, 1), cell.Offset(0, 5)).Font.Bold = True
            End If
        End If
    Next cell
    
    ws.Columns.ColumnWidth = 15

    '======================================================================
    ' PASSO 3 - Processamento da planilha MEMORIA
    '======================================================================
    Dim wsMemoria As Worksheet
    Dim wsICMS As Worksheet
    Dim wsIPI As Worksheet
    Dim wsRef As Worksheet
    Dim lastRowMemoria As Long
    Dim lastRowICMS As Long
    Dim lastRowIPI As Long
    Dim lastRowRef As Long
    Dim i As Long
    Dim j As Long
    Dim valorICMS As Variant
    Dim valorIPI As Variant
    Dim igualEncontrado As Boolean
    Dim palavrasPermitidas As Variant
    Dim cfopParaVerificar As Variant
    Dim valorRef As Variant
    Dim lastRow As Long

    Set wsMemoria = ThisWorkbook.Sheets("MEMORIA")
    Set wsICMS = ThisWorkbook.Sheets("ICMS")
    Set wsIPI = ThisWorkbook.Sheets("IPI")
    Set wsRef = ThisWorkbook.Sheets("Ref")

    lastRowMemoria = wsMemoria.Cells(wsMemoria.Rows.Count, 1).End(xlUp).Row
    lastRowICMS = wsICMS.Cells(wsICMS.Rows.Count, 1).End(xlUp).Row
    lastRowIPI = wsIPI.Cells(wsIPI.Rows.Count, 1).End(xlUp).Row
    lastRowRef = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row

    With wsMemoria.Range("A1:K1")
        .Value = Array("CFOP", "Descrição", "Valor Contábil", "Base de Cálculo", "Diferença", _
                      "Valor ICMS", "Valor IPI", "ICMS - IPI", "Diferença Ajustada", "Observação", "Observação 2")
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230)
    End With

    If wsMemoria.Cells(3, 1).Value <> "Saídas" Then
        wsMemoria.Rows("2:3").Delete
    Else
        wsMemoria.Rows(2).Delete
    End If

    lastRow = wsMemoria.Cells(wsMemoria.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If wsMemoria.Cells(i, 1).Value = "Entradas" Then
            wsMemoria.Rows(i).Insert
            lastRow = lastRow + 1
            wsMemoria.Cells(i, 1).Font.Bold = True
            Exit For
        End If
    Next i

    For i = 4 To lastRow
        If wsMemoria.Cells(i, 1).Value = "Bases Extras Tributáveis" Then
            wsMemoria.Rows(i & ":" & lastRow).Delete
            Exit For
        End If
    Next i

    palavrasPermitidas = Array("Entradas", "Base Tributo Entrada", "Base Tributo", "Base Tributo Saída", "Saídas")
    cfopParaVerificar = Array("1253 - 01", "1407 - 01", "1556 - 01", "1556 - 09")

    For i = 4 To wsMemoria.Cells(wsMemoria.Rows.Count, 1).End(xlUp).Row
        igualEncontrado = False
        
        For j = 4 To lastRow
            If wsMemoria.Cells(i, 2).Value = wsMemoria.Cells(j, 2).Value And i <> j Then
                igualEncontrado = True
                Exit For
            End If
        Next j

        If IsError(Application.Match(wsMemoria.Cells(i, 1).Value, palavrasPermitidas, 0)) Then
            If Not igualEncontrado Then
                If Not wsMemoria.Cells(i, 1).Value Like "#### - ##" And _
                   Not wsMemoria.Cells(i, 1).Value Like "# - ##" And _
                   Not wsMemoria.Cells(i, 1).Value Like "#### - #" Then
                    wsMemoria.Cells(i, 1).ClearContents
                End If
            End If
        Else
            If wsMemoria.Cells(i, 1).Value = "Entradas" Or wsMemoria.Cells(i, 1).Value = "Saídas" Then
                wsMemoria.Cells(i, 1).Font.Bold = True
            End If
        End If
    Next i

    For i = 4 To lastRow
        If wsMemoria.Cells(i, 1).Value = "Saídas" Then
            wsMemoria.Cells(i, 1).Font.Bold = True
        End If
    Next i

    For i = 4 To lastRow
        If Not (wsMemoria.Cells(i, 1).Value = "Base Tributo Entrada" Or _
                wsMemoria.Cells(i, 1).Value = "Base Tributo" Or _
                wsMemoria.Cells(i, 1).Value = "Base Tributo Saída" Or _
                wsMemoria.Cells(i, 1).Value = "Saídas") Then
            
            If IsNumeric(Trim(wsMemoria.Cells(i, 2).Value)) Then
                wsMemoria.Cells(i - 1, 3).Value = wsMemoria.Cells(i, 2).Value
                wsMemoria.Cells(i, 2).ClearContents
                If IsNumeric(Trim(wsMemoria.Cells(i, 3).Value)) Then
                    wsMemoria.Cells(i - 1, 4).Value = wsMemoria.Cells(i, 3).Value
                    wsMemoria.Cells(i, 3).ClearContents
                End If
            End If
        End If
    Next i

    For i = 4 To lastRow
        If wsMemoria.Cells(i, 1).Value = "Base Tributo Entrada" Or _
           wsMemoria.Cells(i, 1).Value = "Base Tributo" Or _
           wsMemoria.Cells(i, 1).Value = "Base Tributo Saída" Or _
           wsMemoria.Cells(i, 1).Value = "Saídas" Then
            wsMemoria.Cells(i, 1).Font.Bold = True
            wsMemoria.Cells(i, 2).Font.Bold = True
        End If
    Next i

    For i = 3 To lastRow
        If IsNumeric(wsMemoria.Cells(i, 3).Value) And IsNumeric(wsMemoria.Cells(i, 4).Value) Then
            wsMemoria.Cells(i, 5).Value = wsMemoria.Cells(i, 3).Value - wsMemoria.Cells(i, 4).Value
        Else
            wsMemoria.Cells(i, 5).ClearContents
        End If
    Next i

    For i = 3 To lastRow
        If wsMemoria.Cells(i, 5).Value <> 0 Then
            wsMemoria.Cells(i, 5).Interior.Color = RGB(255, 255, 224)
        Else
            wsMemoria.Cells(i, 5).Interior.ColorIndex = xlNone
        End If
    Next i

    For i = 2 To lastRowICMS
        wsICMS.Cells(i, 1).Value = Replace(wsICMS.Cells(i, 1).Value, ".", "")
    Next i

    For i = 3 To lastRow
        If wsMemoria.Cells(i, 1).Value <> "" Then
            valorICMS = Application.VLookup(wsMemoria.Cells(i, 1).Value, wsICMS.Range("A2:D" & lastRowICMS), 4, False)
            If Not IsError(valorICMS) Then
                wsMemoria.Cells(i, 6).Value = valorICMS
            End If
        End If
    Next i

    For i = 2 To lastRowIPI
        wsIPI.Cells(i, 1).Value = Replace(wsIPI.Cells(i, 1).Value, ".", "")
    Next i

    For i = 3 To lastRow
        If wsMemoria.Cells(i, 1).Value <> "" Then
            valorIPI = Application.VLookup(wsMemoria.Cells(i, 1).Value, wsIPI.Range("A2:D" & lastRowIPI), 4, False)
            If Not IsError(valorIPI) Then
                wsMemoria.Cells(i, 7).Value = valorIPI
            End If
        End If
    Next i

    For i = 3 To lastRow
        If IsNumeric(wsMemoria.Cells(i, 6).Value) And IsNumeric(wsMemoria.Cells(i, 7).Value) Then
            wsMemoria.Cells(i, 8).Value = wsMemoria.Cells(i, 6).Value - wsMemoria.Cells(i, 7).Value
        Else
            wsMemoria.Cells(i, 8).ClearContents
        End If
    Next i

    For i = 3 To lastRow
        If IsNumeric(wsMemoria.Cells(i, 8).Value) And IsNumeric(wsMemoria.Cells(i, 5).Value) Then
            wsMemoria.Cells(i, 9).Value = wsMemoria.Cells(i, 8).Value - wsMemoria.Cells(i, 5).Value
        Else
            wsMemoria.Cells(i, 9).ClearContents
        End If
    Next i

    For i = 3 To lastRow
        If Not IsError(Application.Match(wsMemoria.Cells(i, 1).Value, cfopParaVerificar, 0)) Then
            wsMemoria.Cells(i, 10).Value = "Verificar as notas registradas com esta CFOP e variação"
        End If
    Next i

    For i = 3 To lastRow
        If wsMemoria.Cells(i, 9).Value <> 0 Then
            wsMemoria.Cells(i, 11).Value = "Verificar DIFAL. Caso não seja o DIFAL, extraia o relatorio de conferencia dos itens da nota"
        End If
    Next i

    For i = 3 To lastRow
        valorRef = Application.VLookup(wsMemoria.Cells(i, 1).Value, wsRef.Range("A2:B" & lastRowRef), 2, False)
        
        If Not IsError(valorRef) Then
            If valorRef = "não" And wsMemoria.Cells(i, 4).Value <> 0 Then
                wsMemoria.Cells(i, 4).Interior.Color = RGB(255, 182, 193)
            Else
                wsMemoria.Cells(i, 4).Interior.ColorIndex = xlNone
            End If
        End If
    Next i

    For i = lastRow To 3 Step -1
        If wsMemoria.Cells(i, 1).Value = "" Or wsMemoria.Cells(i, 1).Value = 0 Then
            wsMemoria.Rows(i).Delete
        End If
    Next i

    wsMemoria.Range("A1:K1").AutoFilter
    wsMemoria.Columns("A").ColumnWidth = 19
    wsMemoria.Columns("B").ColumnWidth = 56
    wsMemoria.Columns("C:K").ColumnWidth = 19
    wsMemoria.Columns("J:K").ColumnWidth = 56
End Sub




````

# 📊 Automação de Atualização Massiva de Planilhas

## 💡 Visão Geral
Este projeto automatiza a atualização de múltiplas planilhas do Excel utilizando macros em VBA, eliminando a necessidade de abrir manualmente mais de 30 arquivos para inserir novas informações.

## 🚨 O Problema
Antes da implementação da automação, o processo de atualização de planilhas era:

- Extremamente repetitivo e demorado: Cada planilha precisava ser aberta e editada manualmente, consumindo tempo e aumentando o risco de erros.
- Propenso a falhas humanas: Alterações incorretas ou esquecidas poderiam comprometer a integridade dos dados.
- Pouco eficiente: A equipe gastava um tempo considerável com tarefas manuais que poderiam ser automatizadas.

## ✅ A Solução
Foi desenvolvida uma macro em VBA que permite a atualização massiva de planilhas de forma rápida e eficiente, com um único clique. O código:
- Abre automaticamente todas as planilhas necessárias em diretórios específicos.
- Atualiza as informações das abas "Empresas", "Serviços" e "Colaboradores" com base em um arquivo matriz.
- Garante a segurança dos dados ao proteger as planilhas após as alterações.
- Exibe relatórios sobre possíveis falhas no processo, listando arquivos que não puderam ser atualizados.

## 🔧 Tecnologias Utilizadas
- VBA (Visual Basic for Applications) para automação no Excel.
- Dicionários e tratamento de erros para identificar e relatar falhas na atualização.
- Proteção de planilhas com senha para garantir a integridade dos dados.

### Macro
```vba
Sub AtualizarPlanilhas()
    Dim wbMãe As Workbook
    Dim pastaFiscal As String
    Dim pastaContabilidade As String
    Dim pastaPessoal As String
    Dim pastaAdministrativo As String
    Dim pastaDiretoria As String ' Nova variável para a pasta Diretoria
    Dim planilhasNaoAtualizadas As String
    Dim nomesPlanilhasFiscal As Variant
    Dim nomesPlanilhasContabilidade As Variant
    Dim nomesPlanilhasPessoal As Variant
    Dim nomesPlanilhasAdministrativo As Variant
    Dim nomesPlanilhasDiretoria As Variant ' Novo array para as planilhas de Diretoria
    Dim senha As String
    Dim dictErros As Object ' Dicionário para armazenar planilhas não atualizadas

    ' Inicializa o dicionário
    Set dictErros = CreateObject("Scripting.Dictionary")

    ' Caminhos de destino
    pastaFiscal = "G:\EMPRESAS GRACIOSA\001 - Escritorio\Timesheet\Fiscal\"
    pastaContabilidade = "G:\EMPRESAS GRACIOSA\001 - Escritorio\Timesheet\Contabilidade\"
    pastaPessoal = "G:\EMPRESAS GRACIOSA\001 - Escritorio\Timesheet\Pessoal\"
    pastaAdministrativo = "G:\EMPRESAS GRACIOSA\001 - Escritorio\Timesheet\Administrativo\"
    pastaDiretoria = "G:\EMPRESAS GRACIOSA\001 - Escritorio\Timesheet\Diretoria\" ' Caminho da pasta Diretoria

    ' Planilha matriz
    Set wbMãe = ThisWorkbook

    ' Arquivos de planilhas para cada pasta
    nomesPlanilhasFiscal = Array("timesheet - Ana Maria.xlsx", "timesheet - Anderson.xlsx", "timesheet - Ariane.xlsx", _
                                 "timesheet - Camily.xlsx", "timesheet - Daniele.xlsx", "timesheet - Diane.xlsx", _
                                 "timesheet - Ana Paula.xlsx", "timesheet - Edina.xlsx")

    nomesPlanilhasContabilidade = Array("timesheet - Débora.xlsx", "timesheet - Eloisa.xlsx", _
                                        "timesheet - Maria.xlsx", "timesheet - Marili.xlsx", "timesheet - Mielke.xlsx", _
                                        "timesheet - Marcelo.xlsx", "timesheet - Nathally.xlsx")

    nomesPlanilhasPessoal = Array("timesheet - Sossela.xlsx", "timesheet - Francieli.xlsx", _
                                  "timesheet - Gabrielly.xlsx", "timesheet - Geisa.xlsx")

    nomesPlanilhasAdministrativo = Array("timesheet - Bruna.xlsx", "timesheet - Cauane.xlsx", "timesheet - Danielle.xlsx")

    nomesPlanilhasDiretoria = Array("timesheet - Ana Carolina.xlsx", "timesheet - Andre.xlsx", "timesheet - Girelli.xlsx") ' Novos nomes das planilhas da pasta Diretoria

    senha = "1234" ' Senha para desbloquear as planilhas

    ' Atualiza as planilhas de todas as pastas
    AtualizaPlanilhasPorPasta wbMãe, nomesPlanilhasFiscal, pastaFiscal, dictErros
    AtualizaPlanilhasPorPasta wbMãe, nomesPlanilhasContabilidade, pastaContabilidade, dictErros
    AtualizaPlanilhasPorPasta wbMãe, nomesPlanilhasPessoal, pastaPessoal, dictErros
    AtualizaPlanilhasPorPasta wbMãe, nomesPlanilhasAdministrativo, pastaAdministrativo, dictErros
    AtualizaPlanilhasPorPasta wbMãe, nomesPlanilhasDiretoria, pastaDiretoria, dictErros ' Novo trecho para atualizar as planilhas de Diretoria

    ' Exibe a mensagem final sobre as planilhas que não puderam ser atualizadas
    If dictErros.Count > 0 Then
        Dim chunkSize As Integer
        chunkSize = 1000 ' Tamanho máximo de cada parte a ser exibida
        Dim currentPosition As Integer
        currentPosition = 1

        ' Concatena a lista de planilhas não atualizadas
        planilhasNaoAtualizadas = Join(dictErros.Keys, vbCrLf)

        ' Exibe a lista de planilhas não atualizadas em partes
        Do While currentPosition <= Len(planilhasNaoAtualizadas)
            MsgBox Mid(planilhasNaoAtualizadas, currentPosition, chunkSize), vbExclamation
            currentPosition = currentPosition + chunkSize
        Loop
    Else
        MsgBox "Atualização concluída! Todas as planilhas foram atualizadas.", vbInformation
    End If
End Sub

Sub AtualizaPlanilhasPorPasta(wbMãe As Workbook, nomesPlanilhas As Variant, pastaDestino As String, dictErros As Object)
    Dim wbDestino As Workbook
    Dim abaOrigem As Worksheet
    Dim abaDestino As Worksheet
    Dim nomeArquivo As String
    Dim i As Integer
    Dim senha As String

    senha = "1234" ' Senha para desbloquear as planilhas

    ' Atualiza a aba 'empresas'
    Set abaOrigem = wbMãe.Worksheets("empresas")
    For i = LBound(nomesPlanilhas) To UBound(nomesPlanilhas)
        nomeArquivo = pastaDestino & nomesPlanilhas(i)

        If Not IsFileOpen(nomeArquivo) Then
            On Error Resume Next
            Set wbDestino = Workbooks.Open(nomeArquivo)
            On Error GoTo 0

            If Not wbDestino Is Nothing Then
                ' ABA DESTINO
                Set abaDestino = wbDestino.Worksheets("empresas")
                abaDestino.Unprotect Password:=senha
                abaDestino.Range("A1:D1000").ClearContents
                abaOrigem.Range("A1:D1000").Copy Destination:=abaDestino.Range("A1")
                abaDestino.Rows(1).AutoFilter
                abaDestino.Protect Password:=senha, AllowSorting:=True, AllowFiltering:=True
                wbDestino.Close SaveChanges:=True
            Else
                If Not dictErros.Exists(nomeArquivo) Then
                    dictErros.Add nomeArquivo, 1
                End If
            End If
        Else
            If Not dictErros.Exists(nomeArquivo) Then
                dictErros.Add nomeArquivo, 1
            End If
        End If
    Next i

    ' Atualiza a aba 'serviços'
    Set abaOrigem = wbMãe.Worksheets("serviços")
    For i = LBound(nomesPlanilhas) To UBound(nomesPlanilhas)
        nomeArquivo = pastaDestino & nomesPlanilhas(i)

        If Not IsFileOpen(nomeArquivo) Then
            On Error Resume Next
            Set wbDestino = Workbooks.Open(nomeArquivo)
            On Error GoTo 0

            If Not wbDestino Is Nothing Then
                ' ABA DESTINO
                Set abaDestino = wbDestino.Worksheets("serviços")
                abaDestino.Unprotect Password:=senha
                abaDestino.Range("A1:D1000").ClearContents
                abaOrigem.Range("A1:D1000").Copy Destination:=abaDestino.Range("A1")
                abaDestino.Rows(1).AutoFilter
                abaDestino.Protect Password:=senha, AllowSorting:=True, AllowFiltering:=True
                wbDestino.Close SaveChanges:=True
            Else
                If Not dictErros.Exists(nomeArquivo) Then
                    dictErros.Add nomeArquivo, 1
                End If
            End If
        Else
            If Not dictErros.Exists(nomeArquivo) Then
                dictErros.Add nomeArquivo, 1
            End If
        End If
    Next i

    ' Atualiza a aba 'colaboradores'
    Set abaOrigem = wbMãe.Worksheets("colaboradores")
    For i = LBound(nomesPlanilhas) To UBound(nomesPlanilhas)
        nomeArquivo = pastaDestino & nomesPlanilhas(i)

        If Not IsFileOpen(nomeArquivo) Then
            On Error Resume Next
            Set wbDestino = Workbooks.Open(nomeArquivo)
            On Error GoTo 0

            If Not wbDestino Is Nothing Then
                ' ABA DESTINO
                Set abaDestino = wbDestino.Worksheets("colaboradores")
                abaDestino.Unprotect Password:=senha
                abaDestino.Range("A1:E1000").ClearContents
                abaOrigem.Range("A1:E1000").Copy Destination:=abaDestino.Range("A1")
                abaDestino.Rows(1).AutoFilter
                abaDestino.Protect Password:=senha, AllowSorting:=True, AllowFiltering:=True
                wbDestino.Close SaveChanges:=True
            Else
                If Not dictErros.Exists(nomeArquivo) Then
                    dictErros.Add nomeArquivo, 1
                End If
            End If
        Else
            If Not dictErros.Exists(nomeArquivo) Then
                dictErros.Add nomeArquivo, 1
            End If
        End If
    Next i
End Sub

Function IsFileOpen(filePath As String) As Boolean
    Dim fileNum As Integer
    On Error Resume Next
    fileNum = FreeFile()
    Open filePath For Input Lock Read As #fileNum
    IsFileOpen = (Err.Number <> 0)
    Close #fileNum
    On Error GoTo 0
End Function
```

