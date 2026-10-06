# Dashboard de Fluxo de Caixa

Painel financeiro em HTML/JavaScript puro que lê planilhas `.xlsx` diretamente no navegador e monta indicadores de **entradas, saídas, fluxo de caixa e gastos por setor**. Projeto de portfólio: todos os dados incluídos são **fictícios**.

> **Aviso:** os arquivos de `bd_exemplo/` foram gerados por `gerar_dados_exemplo.py` com valores, setores, fornecedores e beneficiários inventados. Não representam nenhuma instituição real.

## Funcionalidades

- Visão geral executiva com KPIs, tendência e resultado operacional acumulado.
- Abas de **Entradas**, **Saídas** e **Fluxo de Caixa** (séries mensais, médias móveis, seletor de período 12m / 24m / 5 anos / tudo).
- Análise por **setor**: roscas do último mês e dos anos comparados, tabela de variação e tabela mensal por setor com filtro e exportação CSV.
- **Histórico por setor** com KPIs, gráfico mensal e média anual.
- Leitura 100% local: as planilhas são lidas pelo navegador (File System Access API, com alternativa por `<input webkitdirectory>`); nada é enviado a servidores.
- Cache incremental: só as planilhas alteradas são relidas ao clicar em "Atualizar Dashboard".
- Tema claro e tema neon escuro.
- Avisos de qualidade da base (datas/valores inválidos, arquivos repetidos, setores novos) na própria tela e no console.

## Tecnologias

- HTML, CSS e JavaScript (sem framework e sem build)
- [Chart.js 4.4.1](https://www.chartjs.org/) e [chartjs-plugin-datalabels 2.2.0](https://chartjs-plugin-datalabels.netlify.app/)
- [SheetJS (xlsx) 0.18.5](https://sheetjs.com/) para leitura de Excel
- Python 3 + [openpyxl](https://openpyxl.readthedocs.io/) apenas para gerar os dados de exemplo

As bibliotecas são carregadas por CDN (cdnjs), portanto **é necessário ter conexão com a internet** para abrir o painel.

## Como executar

1. (Opcional) Regenerar os dados de exemplo:
   ```bash
   pip install openpyxl
   python gerar_dados_exemplo.py
   ```
2. Servir a pasta do projeto por HTTP (necessário para o navegador carregar os scripts):
   ```bash
   python -m http.server 8813
   ```
3. Abrir `http://localhost:8813/index.html` no Chrome ou Edge.
4. Clicar em **Selecionar pasta da base** e escolher a pasta `bd_exemplo` (que contém `Entradas` e `Saídas`). O painel é montado na hora.

Em navegadores sem File System Access API (ex.: Firefox) o painel usa o seletor de pastas clássico; nesse caso é preciso reselecionar a pasta para reler as planilhas.

## Formato das planilhas

```
bd_exemplo/
  Entradas/Entradas_AAAA.xlsx   colunas: Data | Entradas do dia  (aceita também "Valor" ou "Total")
  Saídas/Saidas_AAAA.xlsx       colunas: Data | Departamento | Valor
```

Colunas extras são ignoradas. Um arquivo por ano; o ano é reconhecido pelo nome do arquivo ou pela pasta. Para usar seus próprios dados, aponte o painel para uma pasta com esse formato.

## Estrutura

```
.
├── index.html               painel (layout, gráficos, tabelas)
├── base-local.js            leitura e consolidação das planilhas .xlsx
├── gerar_dados_exemplo.py   gerador de dados fictícios (seed fixa)
├── bd_exemplo/              planilhas fictícias 2023–2026
│   ├── Entradas/
│   └── Saídas/
├── README.md
├── LICENSE
└── .gitignore
```

Opcionalmente, um script `semente.js` que defina `window.SEMENTE_HISTORICO` pode preencher meses anteriores às planilhas; o painel funciona sem ele.

## Geração dos dados de exemplo

`gerar_dados_exemplo.py` usa seed fixa (`20260101`), gera planilhas de 2023 a set/2026 com 12 setores genéricos (Administrativo, Operações, Tecnologia, Eventos, RH etc.), crescimento anual e sazonalidade leves. Cada arquivo tem menos de 150 KB.

## Privacidade

O repositório não contém dados reais, senhas, chaves ou identidade de organizações. Se você adaptar o painel para dados reais, **não publique as planilhas** (o `.gitignore` já bloqueia `*.xlsx` fora de `bd_exemplo/`) e controle o acesso ao arquivo no servidor ou na rede.

## Licença

[MIT](LICENSE) — Copyright (c) 2026 Alan Moura
