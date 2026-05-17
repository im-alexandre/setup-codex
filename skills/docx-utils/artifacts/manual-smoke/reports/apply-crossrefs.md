# RelatÃ³rio de aplicaÃ§Ã£o de referÃªncias cruzadas

- VersÃ£o de entrada: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\smoke-main.docx`
- VersÃ£o de saÃ­da: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\smoke-main.docx`
- UtilitÃ¡rio/cÃ³digo usado: `docx-utils`, comando `apply-crossrefs` (.NET + Open XML)
- Plano usado: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\plans\crossrefs.json`
- Lock usado: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\apply-crossrefs.lock`
- Autor das revisÃµes: `Ultron-1`
- Gerado em UTC: `2026-05-15T13:36:47.7039808Z`

## Chamadas convertidas e alvos criados
- caption-1: caption numbered/bookmarked as Figura 1
- xref-call-1: converted 1 call(s)

## Chamadas puladas por ambiguidade ou validaÃ§Ã£o
- Nenhuma.

## ValidaÃ§Ã£o executada
- Executar apÃ³s a aplicaÃ§Ã£o: `docx-utils validate <saida.docx>`.
- Executar apÃ³s a aplicaÃ§Ã£o: `docx-utils structure-audit <saida.docx> --out <json>`.

## Riscos remanescentes
- O arquivo de entrada nÃ£o continha campos `SEQ` em legendas de tabela/figura; o utilitÃ¡rio inseriu campos novos apenas nas legendas planejadas e revalidadas.
- Campos `SEQ`/`REF` foram gravados com resultado atual no XML; o Word pode recalcular visualmente os campos ao atualizar o documento.
- A Lista de Tabelas e a Lista de IlustraÃ§Ãµes nÃ£o foram atualizadas nesta rodada.
