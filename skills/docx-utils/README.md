# docx-utils

Skill para inspecao, validacao e mutacao auditavel de arquivos DOCX/Open XML com a ferramenta .NET publicada em `bin/docx-utils`.

## Uso rapido

- Use o binario publicado por padrao:
  - `C:\Users\imale\.codex\skills\docx-utils\bin\docx-utils\docx-utils.exe <comando> <docx> [opcoes]`
- `accept-revisions` aceita insercoes/delecoes rastreadas, repara a validacao antes de salvar e preserva `--disable-track` quando solicitado.
- Consulte `SKILL.md` e `docx-utils --help` para a superficie operacional completa.

## Inventario De Superficie

- JS/Node: vazio nesta skill. Nao ha `package.json`, fontes `.js/.ts/.tsx` ou comandos publicados nesse ramo.
- .NET: a superficie publicada real e a mesma exposta por `docx-utils --help`.

## Comandos Publicados

- Entrada e contratos: `help`, `--help`, `-h`, `/?`, `plan-contracts`, `plan-contract`, `validate-plan`.
- Criacao e template: `create-article`, `create-docx`, `inspect-template`, `validate-template-profile`, `apply-template`, `audit-template-application`.
- Inspecao e auditoria: `paragraphs`, `paragraph-detail`, `structure-audit`, `layout-audit`, `equations-audit`, `math-audit`, `math-text-audit`, `linear-equation-plan-preview`, `revisions`, `comments`, `comment-anchors`, `next-author`, `validate`.
- Estilos e ajustes: `export-used-styles`, `ensure-canonical-styles`, `sync-styles-from-docx`, `style-running-text`, `ensure-style-fonts`, `format-equation-paragraphs`, `normalize-figure-indent`, `apply-table-design-style`.
- Configuracao de abertura: `enable-update-fields-on-open`, `disable-update-fields-on-open`.
- Edicao e revisoes: `insert-tracked`, `insert-blocks`, `replace-blocks`, `edit-paragraphs`, `append-paragraphs`, `accept-revisions`.
- Figuras, formulas e referencias: `insert-figures`, `replace-figures-from-plan`, `rewrite-equation-blocks`, `replace-formulas-with-linear-equations`, `replace-formulas-with-mathml-omml`, `convert-text-formulas-to-omath`, `apply-crossrefs`, `add-bookmarks`, `rewrite-ref-fields`, `repair-style-captions`.
- Reparos academicos: `repair-article-abnt-layout`, `format-abnt-reference-titles`, `repair-layout-pendencies`, `repair-ref-number-only`.
- Comentarios: `insert-comments`, `reanchor-comments`, `answer-comments`, `reply-comments`, `remove-comments`.
- Comandos mutadores exigem `--lock <lockfile>`; em mutacoes, `--author` continua opcional para a thread principal.
- `validate-plan` valida apenas `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
- `plan-contracts`/`plan-contract` expõem os contratos JSON/Markdown para `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
