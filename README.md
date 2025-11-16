# Projeto de automatização e manipulação de dados

Neste projeto é utilizado Python como linguagem principal para implementar de forma eficiente e prática a automatização de processos antes repetitivos e manuais.

## Resumo prático

- **Descrição:** Automatiza o download e a agregação de relatórios (SAF, OSM, SSP, OSP) do sistema interno e gera planilhas com contagens e detalhes de falhas.
- **Linguagem:** Python

**Principais scripts:**
- `baixador.py`: Automação com Playwright para efetuar login e baixar relatórios (SAF, OSM, SSP, OSP) dentro do intervalo de datas definido; pede usuário e senha interativos.
- `mining_combined.py`: Processa CSVs (SAF/OSM) — filtra por grupos/veículos/datas, conta ocorrências por nível (A/B/C) por dia e atualiza `dados_resumo.xlsx`. Também extrai falhas de equipamento em uma aba detalhada.
- `data_mining_ssp.py`: Analisa `SspCompleta.csv` para contar manutenções preventivas (diária/semanal) por veículo e gera abas de resumo e detalhes em `dados_resumo.xlsx`.

**Entradas esperadas:**
- Arquivos CSV gerados pelo sistema: `SafCompleta.csv`, `OsmCompleta.csv`, `SspCompleta.csv`, `OspCompleta.csv` (normalmente baixados em `~/Downloads/Documentos - Farol`).
- Separador `;` e codificação `latin-1` (os scripts tentam alternativas quando necessário).

**Saídas produzidas:**
- `dados_resumo.xlsx` — planilha com abas para contagens diárias, registros sem-falhas, falhas detalhadas e resumo/detalhes das manutenções SSP.

**Uso rápido:**
- Baixar relatórios (interativo): `python baixador.py`
- Processar SAF/OSM e atualizar planilhas: `python mining_combined.py`
- Gerar contagem de manutenções SSP: `python data_mining_ssp.py`

**Notas / cuidados:**
- Verifique/ajuste o caminho padrão `~/Downloads/Documentos - Farol` conforme seu ambiente.
- As regras de filtragem (grupos, veículos, agentes) são específicas e podem ser alteradas diretamente nos scripts quando necessário.
- Confirme que os CSVs possuem colunas esperadas (`Data de Abertura Saf`, `Grupo`, `Veiculo`, `Nome Status`, `Agente Causador`, `Nível`) ou normalize-as antes.

---

