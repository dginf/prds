# Painel de Indicadores dos PRDs

Painel dos indicadores dos Planos Regionais de Desenvolvimento 2024–2027 (PRDA/SUDAM, PRDCO/SUDECO e PRDNE/SUDENE), no mesmo padrão visual e de filtros do [Painel PNDR](https://github.com/dginf/painel-pndr). SDR / DGINF — MIDR.

## Estrutura do repositório

```
painel-prd/
├── index.html          ← painel (HTML + Chart.js + SheetJS via cdnjs)
├── gerar_dados.py      ← gera dados/prd.json a partir da planilha-base
└── dados/
    ├── prd.json        ← indicadores, metas, séries, qualidade e metodologia (~330 KB)
    └── Base_Indicadores_PRDs_Series_Observadas_2026.xlsx  ← base completa (download na aba Metodologia)
```

Nenhum arquivo passa de 1 MB, então não é preciso Git LFS.

## Abas

| Aba | Conteúdo |
|---|---|
| Principal | Busca de indicador, contagem por situação, cartões por plano e gráfico de situação por plano |
| Painel de Metas | Tabela com linha de base, valor mais recente, meta e progresso de cada indicador |
| Evolução do Indicador | Seletor por eixo, série observada com metas anuais, ficha técnica, limitações e fontes |
| Por Eixo | Situação por eixo, tabela eixo × plano e ranking de progresso rumo à meta |
| Qualidade dos Dados | Completude das linhas de base, origem e defasagem das séries, ressalvas documentais |
| Metodologia | Critérios de classificação, decisões metodológicas e documentos-fonte |

Filtros laterais: plano, eixo, situação, origem do dado, busca textual, linha de base com valor, ressalva de comparabilidade, recorte por bioma/UF e polaridade. Todas as tabelas exportam CSV e XLSX.

## Publicar no GitHub Pages (primeira vez)

```bash
git clone https://github.com/dginf/painel-prd
cd painel-prd
# copie index.html, gerar_dados.py, README.md e a pasta dados/ para cá
git add .
git commit -m "Painel PRDs - versão inicial"
git push origin main
```

Depois, em **Settings → Pages**, escolha *Deploy from a branch*, branch `main`, pasta `/ (root)`. O painel fica em `https://dginf.github.io/painel-prd/`.

## Atualizar dados

```bash
pip install pandas openpyxl
python gerar_dados.py dados/Base_Indicadores_PRDs_Series_Observadas_2026.xlsx
git add dados/
git commit -m "Atualização de dados - AAAA-MM-DD"
git push origin main
```

A planilha precisa manter as abas `Indicadores`, `Metas_anuais`, `Series_observadas`, `Cobertura_calculo`, `Qualidade`, `Metodologia_series`, `Fontes` e `Resumo` com os mesmos cabeçalhos.

## Teste local

O `index.html` carrega `dados/prd.json` via `fetch`, que não funciona abrindo o arquivo direto do disco. Use um servidor local:

```bash
python -m http.server 8000
# abrir http://localhost:8000
```

## Critérios de situação frente à meta

- **PRDCO e PRDNE**: meta 2027 comparada ao valor mais recente na mesma unidade, respeitando a polaridade; direção desde a linha de base conforme a aba `Cobertura_calculo`. Sem avaliação quando a observação mais recente não é comparável em nível ou só existe o ano-base. IDH do PRDCO usa a meta de 2030.
- **PRDA**: metas de variação relativa sem valor de linha de base publicado; o painel calcula a variação acumulada entre o valor observado no ano-base e o mais recente e compara com a meta quadrienal.
- Progresso = (valor mais recente − linha de base) ÷ (meta − linha de base).
