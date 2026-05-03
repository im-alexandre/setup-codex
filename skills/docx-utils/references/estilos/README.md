# Estilos extraidos da dissertacao

Estilos extraidos de `dissertacao-20260426-final.docx` com `.NET` + Open XML pelo utilitario:

```powershell
dotnet run --project .codex\docx_utils\StyleXmlExporter\StyleXmlExporter.csproj -- dissertacao-20260426-final.docx .codex\docx_utils\estilos
```

Conteudo:

- `styles.xml`: copia integral de `word/styles.xml`.
- `docDefaults.xml`: definicoes default de documento, quando presentes.
- `latentStyles.xml`: definicoes de estilos latentes, quando presentes.
- `manifest.json`: mapa de `styleId`, nome visual, tipo e arquivo XML individual.
- `styles/*.xml`: um arquivo XML por elemento `w:style`.

Estilos recorrentes nas regras do projeto:

- `Figura`: `styles/0061_paragraph_Figura_Figura.xml`
- `Tabela`: `styles/0078_paragraph_Tabela_Tabela.xml`
- `dados`: `styles/0056_paragraph_dados_dados.xml`
- `tabelauerj`: `styles/0080_table_tabelauerj_tabela_uerj.xml`
- estilo visual `legenda`: `styles/0065_paragraph_legenda0_legenda.xml`

Observacao: o estilo visual `legenda` usa `styleId="legenda0"` no DOCX base. Para aplicar estilo via Open XML, use o `styleId` do XML/manifesto, nao apenas o nome exibido no Word.
