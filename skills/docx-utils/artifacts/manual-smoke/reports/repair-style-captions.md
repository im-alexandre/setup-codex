# Reparo de legendas numeradas por estilo

- Documento: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\smoke-main.docx`
- Plano: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\plans\crossrefs.json`
- Lock: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\repair-style-captions.lock`
- Autor das revisoes: `Brainiac-1`
- Gerado em UTC: `2026-05-15T13:36:48.6224139Z`

## Padrao aplicado

- Legendas de figuras e tabelas permanecem em paragrafos com estilos `Figura` e `Tabela`.
- A numeracao visivel vem da numeracao automatica do estilo, nao de prefixo textual manual nem de campo `SEQ` inserido na legenda.
- Bookmarks `xref_fig_*` e `xref_tab_*` foram reposicionados no paragrafo de legenda numerado.
- Campos `REF` existentes foram ajustados com `\r` para referenciar o numero do paragrafo numerado.
- O documento foi marcado com `w:updateFields`, para recalcular listas automaticas e referencias ao abrir no Word.

## Aplicado
- updated REF fields with paragraph-number switch: 0

## Nao aplicado / revisar
- caption-1: style numbering not found after repair
