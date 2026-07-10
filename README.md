
# 📊 DataVision AI

Uma aplicação web desenvolvida em Python para transformar arquivos de dados em dashboards modernos, interativos e inteligentes.

O sistema permite importar arquivos Excel e CSV, analisar informações, comparar diferentes tabelas, criar gráficos personalizados e fazer perguntas aos dados utilizando Inteligência Artificial.

## 🎯 Objetivo do projeto

O objetivo do **DataVision AI** é facilitar a análise de dados para pessoas que não possuem conhecimentos avançados em programação, estatística ou ferramentas de Business Intelligence.

Com poucos cliques, o usuário poderá transformar planilhas em dashboards e obter respostas claras sobre os dados importados.

## ✨ Principais funcionalidades

* Importação de arquivos Excel nos formatos `.xlsx` e `.xls`;
* Importação de arquivos CSV;
* Visualização dos dados em formato de tabela;
* Limpeza e tratamento básico dos dados;
* Identificação automática das colunas;
* Criação de dashboards interativos;
* Criação de gráficos de barras, linhas, pizza, área e dispersão;
* Aplicação de filtros por período, categoria, produto, jogador, time ou outras informações;
* Comparação entre diferentes arquivos e tabelas;
* Comparação entre períodos, como meses, anos ou temporadas;
* Exportação dos resultados analisados;
* Interface moderna, responsiva e intuitiva;
* Integração com Inteligência Artificial para responder perguntas sobre os dados.

## 🤖 Inteligência Artificial

A aplicação contará com um assistente de Inteligência Artificial capaz de interpretar os arquivos importados e responder perguntas em linguagem natural.

### Exemplos de perguntas

```text
Qual produto vendeu mais comparando os meses de dezembro e janeiro?

Qual foi o mês com maior faturamento?

Quais são os cinco produtos mais vendidos?

Qual vendedor teve o melhor desempenho?

Quem possui mais gols na Ligue 1 e na Serie A?

Qual time possui a maior média de gols?

Qual categoria apresentou maior crescimento?

Compare os resultados das duas planilhas.
```

A IA analisará os dados disponíveis e apresentará uma resposta acompanhada de tabelas, indicadores ou gráficos sempre que possível.

## 📈 Tipos de gráficos

O usuário poderá escolher diferentes formas de visualizar os dados:

* Gráfico de barras;
* Gráfico de linhas;
* Gráfico de pizza;
* Gráfico de área;
* Gráfico de dispersão;
* Histograma;
* Ranking;
* Indicadores numéricos;
* Tabelas interativas;
* Comparação entre períodos;
* Comparação entre diferentes arquivos.

## 🗂️ Exemplos de utilização

O projeto poderá ser utilizado em diferentes áreas.

### Vendas

* Comparar vendas entre meses;
* Identificar os produtos mais vendidos;
* Analisar faturamento;
* Verificar o desempenho dos vendedores;
* Comparar diferentes lojas ou filiais.

### Futebol

* Comparar jogadores de diferentes ligas;
* Identificar os maiores artilheiros;
* Comparar gols, partidas, assistências e médias;
* Criar rankings de jogadores e times;
* Comparar campeonatos e temporadas.

### Financeiro

* Comparar receitas e despesas;
* Identificar os maiores gastos;
* Analisar resultados mensais;
* Acompanhar metas financeiras;
* Criar projeções com base nos dados.

### Estoque

* Identificar produtos com baixo estoque;
* Encontrar produtos com maior saída;
* Comparar entradas e saídas;
* Acompanhar movimentações por período.

## 🛠️ Tecnologias

O projeto poderá utilizar as seguintes tecnologias:

* Python;
* Pandas;
* NumPy;
* Streamlit;
* Plotly;
* Matplotlib;
* OpenPyXL;
* Scikit-learn;
* API de Inteligência Artificial;
* Git e GitHub.

## 📁 Estrutura sugerida do projeto

```text
datavision-ai/
│
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
│
├── data/
│   └── arquivos_importados/
│
├── pages/
│   ├── dashboard.py
│   ├── comparacao.py
│   └── assistente_ia.py
│
├── services/
│   ├── file_reader.py
│   ├── data_processing.py
│   ├── chart_generator.py
│   └── ai_service.py
│
├── utils/
│   └── helpers.py
│
└── assets/
    ├── images/
    └── styles/
```

## 🚀 Como executar o projeto

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/datavision-ai.git
```

### 2. Entre na pasta do projeto

```bash
cd datavision-ai
```

### 3. Crie um ambiente virtual

No Windows:

```bash
python -m venv venv
```

### 4. Ative o ambiente virtual

No Prompt de Comando:

```bash
venv\Scripts\activate
```

No PowerShell:

```bash
.\venv\Scripts\Activate.ps1
```

### 5. Instale as dependências

```bash
pip install -r requirements.txt
```

### 6. Execute a aplicação

```bash
streamlit run app.py
```

Após executar o comando, a aplicação deverá abrir automaticamente no navegador.

Caso isso não aconteça, acesse:

```text
http://localhost:8501
```

## 📦 Dependências sugeridas

Exemplo de conteúdo para o arquivo `requirements.txt`:

```text
streamlit
pandas
numpy
plotly
matplotlib
openpyxl
scikit-learn
python-dotenv
```

A biblioteca utilizada para integração com a Inteligência Artificial deverá ser adicionada de acordo com a API escolhida.

## 🔐 Variáveis de ambiente

As chaves de API não devem ser publicadas diretamente no código.

Crie um arquivo chamado `.env`:

```text
AI_API_KEY=sua_chave_de_api
```

Adicione o arquivo `.env` ao `.gitignore`:

```text
.env
venv/
__pycache__/
*.pyc
```

## 🧭 Fluxo da aplicação

1. O usuário acessa a aplicação;
2. Importa um ou mais arquivos Excel ou CSV;
3. O sistema identifica as colunas e os tipos de dados;
4. O usuário seleciona as informações que deseja analisar;
5. A aplicação cria tabelas, indicadores e gráficos;
6. O usuário pode comparar diferentes arquivos ou períodos;
7. O usuário faz perguntas para a Inteligência Artificial;
8. A aplicação analisa os dados e apresenta a resposta.

## 🎨 Interface

A interface será desenvolvida com foco em:

* Visual moderno e inovador;
* Navegação simples;
* Layout responsivo;
* Cards com indicadores;
* Gráficos interativos;
* Menu lateral;
* Área para importação de arquivos;
* Área de perguntas para a Inteligência Artificial;
* Modo claro e modo escuro;
* Boa experiência em computadores, tablets e celulares.

## 🗺️ Planejamento do projeto

### Primeira versão

* [ ] Criar a estrutura inicial do projeto;
* [ ] Criar a interface com Streamlit;
* [ ] Permitir o envio de arquivos CSV;
* [ ] Permitir o envio de arquivos Excel;
* [ ] Exibir os dados em uma tabela;
* [ ] Criar gráficos básicos.

### Segunda versão

* [ ] Adicionar filtros;
* [ ] Permitir a escolha das colunas;
* [ ] Comparar diferentes períodos;
* [ ] Comparar dois ou mais arquivos;
* [ ] Criar indicadores automáticos;
* [ ] Permitir a exportação dos resultados.

### Terceira versão

* [ ] Integrar uma Inteligência Artificial;
* [ ] Permitir perguntas em linguagem natural;
* [ ] Gerar gráficos com base nas perguntas;
* [ ] Criar resumos automáticos;
* [ ] Identificar tendências e padrões;
* [ ] Publicar a aplicação na internet.

## 💡 Diferenciais

O principal diferencial do projeto será a combinação de:

* Análise de planilhas;
* Dashboards interativos;
* Comparação entre diferentes tabelas;
* Gráficos personalizados;
* Inteligência Artificial;
* Interface simples e moderna.

Dessa forma, o usuário não precisará conhecer comandos de programação ou fórmulas avançadas para analisar seus dados.

## 📌 Status do projeto

```text
🚧 Projeto em desenvolvimento
```

Novas funcionalidades e melhorias serão adicionadas durante o desenvolvimento.

## 🤝 Contribuições

Contribuições são bem-vindas.

Para contribuir:

1. Faça um fork do projeto;
2. Crie uma branch para a alteração;
3. Realize as modificações;
4. Faça o commit;
5. Envie as alterações;
6. Abra um Pull Request.

## 👨‍💻 Autor

Desenvolvido por **Thiago Volpolini**.

* GitHub: `@thiagovolpolini-art`
* Portfólio: `https://thiagovolpolini.com/`

## 📄 Licença

Este projeto poderá ser distribuído sob a licença MIT.

Consulte o arquivo `LICENSE` para mais informações.

  