# Backlog docx-utils

Registro curto de lacunas e falhas observadas durante uso real da skill. Use este arquivo quando uma operação DOCX necessária ainda nao tiver comando publicado ou quando um comando publicado precisar de endurecimento.

## Observado no smoketest de 2026-05-15

- `math-text-audit`: corrigir excecao `RegexParseException` no padrao de letras gregas.
- `export-used-styles`: tratar documentos sem `StyleDefinitionsPart` com erro de contrato legivel, sem excecao nao tratada.
- `ensure-style-fonts`: tratar documentos sem `StyleDefinitionsPart` com diagnostico acionavel.
- `style-running-text`: revisar validacoes que deixam o comando aplicar mudancas e retornar `VALIDATION_ERRORS`.
- `reply-comments`: documentar ou corrigir o caso em que o comentario pai nao possui `paraId`.
- `validate`: investigar XML final com `w:tabs` em posicao invalida e atributos `w15:paraId`/`w15:textId` sem namespace declarado em DOCX sintetico.
