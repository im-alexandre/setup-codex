---
name: pptx-utils
description: Editar, inspecionar, validar e renderizar apresentacoes PPTX com foco em templates existentes, usando o binario .NET/Open XML publicado como caminho padrao.
---

# PPTX Utils

Use esta skill para qualquer tarefa envolvendo `.pptx`.

## Regra De Execucao

1. No primeiro uso ou depois de alterar a fonte, execute:

   `C:\Users\imale\.codex\skills\pptx-utils\scripts\install-pptx-utils.ps1`

   No Linux/WSL, use `scripts/install-pptx-utils.sh`.

1. Para edicao de apresentacoes existentes ou templates, use o binario publicado:

   `C:\Users\imale\.codex\skills\pptx-utils\bin\pptx-utils\pptx-utils.exe <comando> <pptx> [opcoes]`

1. Quando o shim estiver no `PATH`, use:

   `pptx-utils <comando> <pptx> [opcoes]`

1. Nao edite XML de slides manualmente salvo fallback, depuracao ou lacuna registrada em `BACKLOG.md`.
1. Nao gere JavaScript/PptxGenJS para editar templates existentes.
1. Use planos JSON para mutacoes.
1. Depois de mutar um PPTX, rode `validate`; para entrega visual, use `scripts/thumbnail.py` ou `scripts/office/soffice.py` enquanto `render-preview` nao estiver implementado.

## Comandos Disponiveis

- `inspect`: lista estrutura, slides, shapes e textos.
- `text-map`: gera mapa de textos editaveis.
- `replace-text`: substitui textos por plano JSON.
- `validate`: valida integridade Open XML.

## Backlog De Comandos

- `duplicate-slide`: duplica slide existente.
- `delete-slide`: remove slides.
- `reorder-slides`: reordena slides.
- `replace-image`: substitui imagens por plano JSON.
- `render-preview`: renderiza previews usando LibreOffice/Poppler.

## Recursos Herdados

- `scripts/thumbnail.py`: analise visual rapida de templates.
- `scripts/office/soffice.py`: conversao headless para PDF.
- `scripts/office/unpack.py`, `pack.py`, `validate.py`: fallback e depuracao.
- `references/plan-contracts.md`: contratos de planos.

## Fallback Para Criacao Do Zero

PptxGenJS pode ser usado apenas quando nao houver template e a tarefa for criar um deck novo do zero. Esse caminho e secundario e nao deve substituir o fluxo de edicao confiavel.
