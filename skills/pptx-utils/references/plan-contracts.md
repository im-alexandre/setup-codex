# Contratos De Planos pptx-utils

## replace-text

```json
{
  "replacements": [
    {
      "slideIndex": 1,
      "shapeIndex": 2,
      "find": "Texto antigo",
      "replace": "Texto novo",
      "mode": "exact"
    }
  ]
}
```

`slideIndex` e 1-based. `shapeIndex` e 1-based dentro do slide. `mode` aceita `exact`.

## replace-image

```json
{
  "replacements": [
    {
      "slideIndex": 1,
      "shapeIndex": 4,
      "imagePath": "imagens/nova.png"
    }
  ]
}
```
