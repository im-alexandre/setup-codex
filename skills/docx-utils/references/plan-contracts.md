# Contratos de planos

Fonte operacional:

- `docx-utils plan-contracts [comando] [--format markdown|json]`
- `references/plan-contracts.json`

Este arquivo resume os contratos publicados para `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.

## `create-docx`

### Contrato

- Sem `--plan`, cria um DOCX vazio.
- Com `--plan`, a raiz JSON deve ser um objeto.
- `title` e `paragraphs` sao obrigatorios quando o plano contem conteudo.
- `paragraphs` deve conter ao menos um paragrafo nao vazio.
- `subtitles`, `sections` e `references` sao opcionais.
- `sections`, quando presente, e uma lista de objetos com `heading`, `level` e `paragraphs`.

### Exemplo minimo

```json
{
  "title": "Titulo do Documento",
  "paragraphs": [
    "Primeiro paragrafo.",
    "Segundo paragrafo."
  ],
  "subtitles": [
    "Subtitulo opcional"
  ],
  "sections": [
    {
      "heading": "Secao 1",
      "level": 1,
      "paragraphs": [
        "Texto da secao."
      ]
    }
  ]
}
```

### Consultas uteis

```powershell
docx-utils create-docx novo.docx
docx-utils create-docx tese.docx --plan documento.json
docx-utils validate-plan create-docx --plan documento.json
docx-utils plan-contracts create-docx --format markdown
docx-utils plan-contracts create-docx --format json
```

## `insert-blocks`

### Contrato

- Raiz JSON: objeto com `blocks`.
- `blocks` nao pode estar vazio.
- Cada bloco precisa de `id`, `afterPrefix`, `beforePrefix` e `items`.
- `styleSource`, quando presente, deve ser `after` ou `before`.
- Cada item precisa de `kind` valido: `paragraph` ou `table`.
- Itens `paragraph` precisam de `text` ou `latex`.
- Itens `table` precisam de `rows` com ao menos uma linha e uma celula por linha.

### Exemplo minimo

```json
{
  "blocks": [
    {
      "id": "bloco-1",
      "afterPrefix": "Texto antes",
      "beforePrefix": "Texto depois",
      "items": [
        {
          "kind": "paragraph",
          "text": "Paragrafo inserido",
          "styleId": "CorpoTexto"
        },
        {
          "kind": "table",
          "tableStyleId": "tabelauerj",
          "cellStyleId": "dados",
          "rows": [
            ["A1", "A2"],
            ["B1", "B2"]
          ]
        }
      ]
    }
  ]
}
```

### Consultas uteis

```powershell
docx-utils validate-plan insert-blocks --plan blocos.json
docx-utils plan-contracts insert-blocks --format markdown
docx-utils plan-contracts insert-blocks --format json
```

## `replace-blocks`

### Contrato

- Raiz JSON: objeto com `blocks`.
- `blocks` nao pode estar vazio.
- Cada bloco precisa de `id`, `afterPrefix`, `beforePrefix` e `items`.
- `styleSource`, quando presente, deve ser `after` ou `before`.
- Cada item precisa de `kind` valido: `paragraph` ou `table`.
- Itens `paragraph` precisam de `text` ou `latex`.
- Itens `table` precisam de `rows` com ao menos uma linha e uma celula por linha.
- O comando remove o conteudo existente entre os marcadores antes de inserir os novos blocos.

### Exemplo minimo

```json
{
  "blocks": [
    {
      "id": "intervalo-1",
      "afterPrefix": "Introducao",
      "beforePrefix": "Conclusao",
      "items": [
        {
          "kind": "table",
          "tableStyleId": "tabelauerj",
          "cellStyleId": "dados",
          "rows": [
            ["A1", "A2"],
            ["B1", "B2"]
          ]
        }
      ]
    }
  ]
}
```

### Consultas uteis

```powershell
docx-utils validate-plan replace-blocks --plan blocos.json
docx-utils plan-contracts replace-blocks --format markdown
docx-utils plan-contracts replace-blocks --format json
```

## `replace-table`

### Contrato

- Raiz JSON: objeto com `tables`.
- `tables` nao pode estar vazio.
- Cada item precisa de `id`.
- Pelo menos um seletor deve estar presente: `ordinal`, `block`, `blockIndex`, `firstCellText`, `previousParagraphPrefix` ou `nextParagraphPrefix`.
- `ordinal`, `block` e `blockIndex`, quando presentes, devem ser inteiros positivos.
- `rows` e obrigatorio e nao pode estar vazio.
- Cada linha em `rows` precisa conter ao menos uma celula.

### Exemplo minimo

```json
{
  "tables": [
    {
      "id": "tabela-1",
      "ordinal": 2,
      "rows": [
        ["Novo A1", "Novo A2"],
        ["Novo B1", "Novo B2"]
      ]
    }
  ]
}
```

### Consultas uteis

```powershell
docx-utils validate-plan replace-table --plan tabela.json
docx-utils plan-contracts replace-table --format markdown
docx-utils plan-contracts replace-table --format json
```
