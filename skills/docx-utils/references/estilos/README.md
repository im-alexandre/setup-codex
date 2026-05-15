# Referencia de estilos canonicos

Este diretório guarda estilos extraidos de `dissertacao-20260426-final.docx` para uso operacional da skill `docx-utils`. Para aplicar ou sincronizar estilos em documentos, use primeiro o binario publicado:

```powershell
docx-utils ensure-canonical-styles <docx>
docx-utils sync-styles-from-docx <docx> --source <modelo.docx>
docx-utils apply-table-design-style <docx> --style-id tabelauerj
```

Use `src/StyleXmlExporter` apenas em manutencao da skill, quando for necessario regenerar estes arquivos a partir de um DOCX fonte.

Conteudo:

- `styles.xml`: copia integral de `word/styles.xml`.
- `docDefaults.xml`: definicoes default de documento, quando presentes.
- `latentStyles.xml`: definicoes de estilos latentes, quando presentes.
- `manifest.json`: mapa de `styleId`, nome visual, tipo e arquivo XML individual.
- `styles/*.xml`: um arquivo XML por elemento `w:style`.

Estilos recorrentes nas regras e nos contratos da skill:

- `Figura`: `styles/0061_paragraph_Figura_Figura.xml`
- `Tabela`: `styles/0078_paragraph_Tabela_Tabela.xml`
- `dados`: `styles/0056_paragraph_dados_dados.xml`
- `tabelauerj`: `styles/0080_table_tabelauerj_tabela_uerj.xml`
- estilo visual `legenda`: `styles/0065_paragraph_legenda0_legenda.xml`

Observacao: o estilo visual `legenda` usa `styleId="legenda0"` no DOCX base. Para aplicar estilo via Open XML, use o `styleId` do XML/manifesto, nao apenas o nome exibido no Word.

Para tabelas UERJ, preserve `tableStyleId: "tabelauerj"` nos planos `insert-blocks`, `replace-blocks` e `replace-table`.
