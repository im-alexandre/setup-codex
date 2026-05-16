# Docx Utils Template Profile

Data: 2026-05-16

## Contexto

A `docx-utils` ja trabalha como utilitario .NET/Open XML para criar, inspecionar, validar e mutar arquivos DOCX com contratos CLI explicitos. O proximo passo e evoluir a skill para adaptar documentos a templates reais de Word, especialmente templates de artigo, dissertacao, relatorio ou material didatico que podem vir preenchidos com conteudo exemplo, semi-vazios ou vazios.

O caso principal nao deve depender de estilos semanticos do Word. Em muitos templates, titulo, secoes, autores, resumo, referencias e legendas aparecem todos com estilo `Normal`, diferenciados apenas por formatacao direta como negrito, italico, caixa alta, alinhamento, recuo, espacamento e numeracao manual. Por isso, a ferramenta nao deve assumir que `Heading 1`, `Title` ou `Caption` existem.

O recurso sera operado pelo Codex, nao manualmente pelo usuario. A extracao pesada do template deve acontecer uma unica vez. Depois que o perfil canonico do template for validado, aplicacoes futuras devem usar somente essa definicao canonica, salvo quando o template mudar ou o usuario pedir nova extracao.

## Objetivos

- Usar .NET/Open XML como motor deterministico para extrair estrutura e formatacao de templates DOCX.
- Permitir que o agente faca a curadoria semantica dos candidatos extraidos e grave um perfil canonico reutilizavel.
- Adaptar documentos existentes para templates preenchidos, semi-vazios ou vazios.
- Preservar a formatacao real do template alvo, inclusive quando tudo estiver em estilo `Normal`.
- Guardar perfis canonicos no projeto consumidor, nao dentro da skill global.
- Validar hash do template antes de aplicar um perfil canonico.
- Produzir relatorios auditaveis de extracao, curadoria, aplicacao e pendencias.
- Tratar referencias bibliograficas com granularidade suficiente para preservar formato e formatacao interna.

## Fora De Escopo

- Implementar uma classificacao semantica automatica no binario .NET.
- Depender de regex como fonte final para decidir se um bloco e titulo, resumo, secao ou referencia.
- Substituir a curadoria do agente por heuristicas opacas.
- Reextrair o template em toda aplicacao.
- Armazenar perfis de projeto dentro de `C:\Users\imale\.codex\skills\docx-utils`.
- Modificar arquivos originais usados como fixtures de teste, como `D:\dissertacao\normas` ou `D:\dissertacao\dissertacao_versoes`.

## Arquitetura Do Fluxo

A solucao tem tres papeis separados:

- `docx-utils`: extrai estrutura, formatacao, runs, regioes, tabelas, imagens, cabecalhos, rodapes e referencias; aplica transformacoes deterministicas; valida contratos e gera relatorios.
- Codex/agente: le os candidatos extraidos, faz a interpretacao semantica e cria ou corrige o perfil canonico.
- Projeto consumidor: guarda template, candidatos, perfil canonico, relatorios e artefatos de aplicacao.

Estrutura recomendada no projeto consumidor:

```text
.codex/docx-templates/
  nome-do-template/
    template.docx
    profile.candidates.json
    profile.candidates.md
    profile.canonical.json
    profile.canonical.md
    applications/
      2026-05-16-documento-report.md
```

Fluxo operacional:

```powershell
docx-utils inspect-template template.docx --out .codex/docx-templates/nome/profile.candidates.json --report .codex/docx-templates/nome/profile.candidates.md
docx-utils validate-template-profile .codex/docx-templates/nome/profile.canonical.json
docx-utils apply-template --template .codex/docx-templates/nome/template.docx --source artigo.docx --profile .codex/docx-templates/nome/profile.canonical.json --out artigo-adaptado.docx
docx-utils audit-template-application artigo-adaptado.docx --profile .codex/docx-templates/nome/profile.canonical.json --report .codex/docx-templates/nome/applications/aplicacao.md
```

`inspect-template` so deve ser executado novamente se o hash do template mudar, se o perfil canonico for descartado ou se o usuario pedir explicitamente nova leitura.

## Extracao Do Template

`inspect-template` deve produzir um retrato tecnico do template, sem decidir semanticamente o papel final de cada elemento.

Elementos a extrair:

- paragrafos do corpo;
- tabelas e celulas;
- imagens e relacoes com legendas proximas;
- cabecalhos e rodapes;
- quebras de pagina e de secao;
- campos Word;
- comentarios ou instrucoes embutidas no template;
- referencias bibliograficas em nivel de paragrafo e de run.

Para cada bloco, a extracao deve registrar:

- `id` estavel dentro da extracao;
- texto;
- parte do documento (`body`, `header`, `footer`, tabela, celula);
- ordinal;
- estilo Word declarado;
- formatacao direta dos runs;
- alinhamento, espacamento, recuo e propriedades de paragrafo;
- caixa alta;
- sinais estruturais, como numeracao manual ou texto curto destacado;
- vizinhos anterior e posterior;
- relacao com tabela, imagem, campo ou comentario.

Exemplo conceitual:

```json
{
  "id": "p-0012",
  "text": "1 INTRODUCAO",
  "location": {
    "part": "body",
    "ordinal": 12
  },
  "paragraphFormat": {
    "styleId": "Normal",
    "alignment": "left",
    "spacingBefore": 240,
    "spacingAfter": 120
  },
  "runs": [
    {
      "text": "1 INTRODUCAO",
      "bold": true,
      "italic": false,
      "fontSize": 12,
      "allCaps": true
    }
  ],
  "structuralHints": {
    "manualNumbering": "1",
    "shortHighlightedParagraph": true
  },
  "neighbors": {
    "previous": "p-0011",
    "next": "p-0013"
  }
}
```

`profile.candidates.md` deve ser otimizado para leitura pelo agente: tabelas compactas, agrupamento por regioes e destaque de formatacoes relevantes.

## Curadoria Semantica Pelo Agente

A semantica nao pertence ao binario .NET nesta primeira versao. O agente le os candidatos e cria a definicao canonica:

- identifica titulo, autores, afiliacoes, resumo, palavras-chave, abstract, keywords, secoes, subsecoes, figuras, tabelas, referencias e anexos;
- decide quais partes do template sao conteudo exemplo removivel;
- decide quais partes sao estrutura fixa a preservar;
- define de onde viram os dados no documento fonte;
- registra pendencias e regioes ambiguas;
- documenta as decisoes em `profile.canonical.md`.

O perfil canonico e o contrato entre a interpretacao do agente e a execucao deterministica da ferramenta.

## Perfil Canonico

`profile.canonical.json` deve conter:

- identificacao do template;
- caminho relativo esperado do template;
- `sha256` do DOCX usado na extracao;
- versao do perfil;
- data da extracao;
- regioes substituiveis;
- regioes preservadas;
- regras de formatacao;
- regras de aplicacao para conteudo fonte;
- regras especificas para referencias;
- pendencias conhecidas;
- origem dos blocos usados na curadoria.

Exemplo conceitual:

```json
{
  "template": {
    "name": "simposio-x",
    "sourceFile": "template.docx",
    "sha256": "...",
    "profileVersion": 1
  },
  "regions": [
    {
      "role": "title",
      "templateBlockId": "p-0004",
      "replaceWith": "source.title",
      "preserveFormatting": true
    },
    {
      "role": "abstract",
      "startBlockId": "p-0012",
      "endBlockId": "p-0015",
      "replaceWith": "source.abstract",
      "preserveFormattingFrom": "p-0012"
    },
    {
      "role": "references",
      "startBlockId": "p-0080",
      "endBlockId": "p-0098",
      "replaceWith": "source.references",
      "referenceFormattingProfile": "refs-main"
    }
  ]
}
```

## Referencias Bibliograficas

Referencias bibliograficas exigem tratamento especial. A ferramenta deve extrair a secao de referencias em nivel de paragrafo e run, pois a formatacao de cada elemento pode importar:

- autores;
- ano;
- titulo;
- periodico, evento ou editora;
- volume, numero e paginas;
- DOI, URL ou hyperlink;
- pontuacao;
- italico;
- negrito;
- caixa alta;
- recuo pendente;
- espacamento entre referencias.

Exemplo conceitual de candidato:

```json
{
  "id": "ref-0003",
  "paragraphId": "p-0088",
  "text": "SILVA, J. A. Titulo do artigo. Revista X, v. 10, n. 2, p. 1-10, 2024.",
  "paragraphFormat": {
    "hangingIndent": true,
    "spacingAfter": 120
  },
  "runs": [
    {
      "text": "SILVA, J. A.",
      "bold": false,
      "italic": false,
      "allCaps": true
    },
    {
      "text": "Titulo do artigo.",
      "bold": true,
      "italic": false
    },
    {
      "text": "Revista X",
      "bold": false,
      "italic": true
    }
  ]
}
```

O agente define no perfil canonico como autores, titulo, fonte, DOI e links devem ser preservados ou recriados. A aplicacao deve usar esse perfil para formatar as referencias do documento fonte no padrao do template alvo.

## Aplicacao Do Template

`apply-template` deve:

- validar o perfil canonico antes de mutar;
- calcular o hash do template informado;
- bloquear ou alertar quando o hash nao bater com o perfil;
- copiar o template para o arquivo de saida;
- remover ou substituir apenas regioes declaradas no perfil;
- preservar cabecalhos, rodapes, margens, quebras, estilos, numeracao e formatacao direta do template;
- inserir conteudo do documento fonte usando as regras canonicas;
- gerar relatorio com regioes aplicadas, regioes preservadas, pendencias e avisos.

O documento fonte pode variar a cada aplicacao. A leitura do documento fonte pode ocorrer por execucao, porque o conteudo e novo. A restricao de extracao unica se aplica ao template apos o perfil canonico ser validado.

## Auditoria

`audit-template-application` deve verificar:

- se o DOCX final e valido em Open XML;
- se ainda ha regioes de exemplo que deveriam ter sido removidas;
- se regioes obrigatorias foram preenchidas;
- se referencias receberam a formatacao esperada;
- se cabecalho, rodape, margens e quebras principais vieram do template;
- se ha comentarios ou revisoes pendentes;
- se o relatorio de aplicacao contem pendencias.

## Cenario De Teste Real

Para validar o recurso com documentos reais, usar os arquivos locais de dissertacao somente como fonte read-only:

- `D:\dissertacao\normas`: contem normas e um template de dissertacao.
- `D:\dissertacao\dissertacao_versoes`: contem versoes DOCX da dissertacao.

O teste deve criar um diretorio separado, por exemplo:

```text
C:\Users\imale\.codex\skills\docx-utils\artifacts\template-profile-dissertacao-test\
```

O fluxo de teste deve copiar para esse diretorio os elementos necessarios de `D:\dissertacao`, sem modificar nada no diretorio original ou seus subdiretorios.

Objetivo do teste real:

1. Gerar candidatos de template com base no modelo de dissertacao em `normas`.
2. Usar tambem os documentos normativos em `normas` como contexto para curadoria do perfil canonico, quando necessario.
3. Criar o perfil canonico do modelo.
4. Transplantar uma versao `.docx` de `dissertacao_versoes` para o modelo recem perfilado.
5. Gerar relatorio de aplicacao e auditoria.
6. Confirmar que nenhum arquivo em `D:\dissertacao` foi alterado.

## Testes Automatizados

A primeira implementacao deve seguir TDD e usar documentos sinteticos pequenos antes do teste real:

- template preenchido com estilo `Normal` e formatacoes diretas diferentes;
- template vazio ou com placeholders;
- template com referencias contendo runs diferentes;
- perfil canonico valido;
- perfil canonico invalido;
- aplicacao com hash correto;
- bloqueio ou aviso com hash divergente;
- aplicacao preservando formatacao do template;
- auditoria detectando regiao obrigatoria nao preenchida;
- referencias formatadas conforme perfil.

## Criterios De Aceite

- `inspect-template` gera JSON e Markdown com estrutura, formatacao e runs relevantes.
- O perfil canonico e validavel por comando dedicado.
- `apply-template` usa perfil canonico e nao reexecuta a extracao do template.
- O hash do template e conferido antes da aplicacao.
- Regioes nao declaradas no perfil sao preservadas.
- Referencias bibliograficas preservam formato e formatacao interna conforme perfil.
- Os perfis ficam no projeto consumidor.
- O teste real copia arquivos de `D:\dissertacao` para area separada e nao altera os originais.
- A implementacao mantem o contrato published-binary-first da `docx-utils`.

