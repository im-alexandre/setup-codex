# Reparo ABNT/UERJ do artigo

- Documento: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\smoke-main.docx`
- Lock: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\repair-article.lock`
- Autor da etapa: `Jarvis-1`
- Utilitario/codigo usado: `docx-utils`, comando `repair-article-abnt-layout` (.NET + Open XML)
- Gerado em UTC: `2026-05-15T13:37:55.4051382Z`

## Aplicado
- settings: removed `w:updateFields` to avoid Word prompt for fields that may reference external files
- figure: Figura smoke inserida
- tables: applied `tabelauerj`, paragraph style `dados`, centered width 100% and autofit to 2 remaining table(s)
- sources: applied style `legenda` to all source paragraphs below figures/tables
- paragraphs: indentation set to zero on 3 figure caption/image/source paragraph(s)
- references: moved FONTANA before GNEITING to restore alphabetical order

## Nao aplicado / revisar
- table `SÃ­ntese da base de PLD utilizada`: caption not found
- table `ConfiguraÃ§Ã£o do experimento de geraÃ§Ã£o de cenÃ¡rios`: caption not found
- table `MÃ©tricas de validaÃ§Ã£o dos cenÃ¡rios de PLD em 2025`: caption not found
- table `CritÃ©rios de aceite registrados no experimento`: caption not found

## Observacoes
- As figuras foram movidas para o paragrafo da legenda, com estilo `Figura`, centralizacao e wrap `topBottom`.
- As tabelas duplicadas anteriores foram removidas mantendo a segunda tabela de cada par, conforme a revisao visual do artigo.
- Tabelas remanescentes usam estilo `tabelauerj`, largura 100% e `autofit`; paragrafos internos usam estilo `dados`.
- As fontes permanecem em paragrafos proprios abaixo de figuras/tabelas, com estilo `legenda`.
