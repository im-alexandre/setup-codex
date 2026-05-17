# Reparo de referÃªncias cruzadas numÃ©ricas

- Documento: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\smoke-main.docx`
- Lock: `C:\Users\imale\.codex\skills\docx-utils\artifacts\manual-smoke\repair-ref-number.lock`
- Autor das revisoes: `Brainiac-1`
- Gerado em UTC: `2026-05-15T13:36:49.0195983Z`

## Aplicado
- updated REF fields from label-returning switches to number-only switch: 0
- updated cached REF results to numeric text: 0

## ObservaÃ§Ã£o
- Campos `REF xref_fig_*` e `REF xref_tab_*` devem retornar apenas o nÃºmero da legenda, porque o texto corrido jÃ¡ contÃ©m `figura`, `figuras`, `tabela` ou `tabelas`.
