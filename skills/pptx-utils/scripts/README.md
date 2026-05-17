# Scripts herdados

Estes scripts foram herdados da skill `pptx` original para preservar utilidades de apoio enquanto o fluxo principal de mutacao passa a ser o binario `.NET` de `pptx-utils`.

Uso previsto:

- `thumbnail.py`: analise visual rapida de templates e decks existentes.
- `office/soffice.py`: conversao headless para PDF.
- `office/unpack.py`, `office/pack.py`, `office/validate.py`: fallback e QA.

Esses scripts nao sao o caminho padrao de mutacao. Quando existir comando equivalente em `pptx-utils`, prefira o comando da skill nova e reserve estes scripts para fallback, depuracao e verificacao.
