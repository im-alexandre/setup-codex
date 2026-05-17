# pptx-utils: skill unica para edicao confiavel de PPTX

Data: 2026-05-16

## Contexto

A skill `pptx` atual e util para leitura, criacao visual e QA de apresentacoes, mas seu caminho de criacao do zero depende fortemente de `PptxGenJS` e incentiva geracao de codigo JavaScript longo. Para edicao de apresentacoes existentes, a skill usa um fluxo de unpack, edicao XML manual, limpeza e pack. Esse fluxo funciona como fallback, mas consome contexto, aumenta risco de corromper relacoes/namespace e exige que o agente leia e edite XML de slides diretamente.

O projeto `docx-utils` ja estabeleceu um padrao operacional melhor para documentos Office: skill com binario publicado primeiro, contratos JSON, validacao, testes xUnit e manipulacao Open XML encapsulada em .NET. A nova solucao deve aplicar esse padrao a PPTX, sem criar duas skills concorrentes para a mesma tarefa.

## Decisao De Produto

A entrada operacional sera uma skill unica chamada `$pptx-utils`.

A skill atual `pptx` deve ser incorporada/migrada para `C:\Users\imale\.codex\skills\pptx-utils`, e o ponto de entrada principal passa a ser o comando `pptx-utils`. A skill antiga nao deve continuar como caminho concorrente para a mesma tarefa.

## Objetivo

Criar uma skill local `pptx-utils` focada em edicao confiavel de apresentacoes `.pptx` existentes e templates, usando .NET/Open XML como nucleo mutador e reaproveitando os recursos uteis da skill `pptx` atual para leitura, thumbnails, renderizacao e QA visual.

## Nao Objetivos

- Criar uma segunda skill user-facing para PPTX.
- Reescrever renderizacao visual em .NET.
- Depender de PowerPoint instalado.
- Usar `PptxGenJS` como caminho principal para edicao de templates.
- Editar XML manualmente como fluxo normal.
- Reimplementar todos os recursos de criacao visual do zero no primeiro MVP.

## Arquitetura Alvo

A pasta alvo sera:

```text
C:\Users\imale\.codex\skills\pptx-utils\
  SKILL.md
  BACKLOG.md
  bin\pptx-utils\
  docs\superpowers\specs\
  docs\superpowers\plans\
  references\plan-contracts.md
  references\plan-contracts.json
  scripts\install-pptx-utils.ps1
  scripts\install-pptx-utils.sh
  scripts\thumbnail.py
  scripts\office\soffice.py
  scripts\office\unpack.py
  scripts\office\pack.py
  scripts\office\validate.py
  src\PptxOpenXmlTools\PptxOpenXmlTools.csproj
  src\PptxOpenXmlTools.Tests\PptxOpenXmlTools.Tests.csproj
```

O binario publicado `bin\pptx-utils\pptx-utils.exe` sera o caminho operacional padrao no Windows. No Linux/WSL, o equivalente publicado sera `bin/pptx-utils/pptx-utils`.

## Dependencias

- `DocumentFormat.OpenXml`: base para abrir, validar e manipular pacotes PresentationML.
- `ShapeCrawler`: camada produtiva para trabalhar com slides, shapes, texto, imagens e tabelas com menos XML manual.
- `Spectre.Console`: saida CLI legivel, seguindo o estilo de `docx-utils` quando fizer sentido.
- `System.Text.Json`: contratos JSON declarativos e saidas estruturadas.
- Scripts reaproveitados da skill `pptx` atual:
  - `thumbnail.py`;
  - `scripts/office/soffice.py`;
  - `scripts/office/unpack.py`, `pack.py` e `validate.py` como fallback/debug;
  - guias de QA visual e cuidados de template.

## Politica De Uso

O `SKILL.md` de `pptx-utils` deve orientar:

1. Para `.pptx` existente ou template, use `pptx-utils` primeiro.
2. Use planos JSON para mutacoes.
3. Nao edite XML de slide manualmente salvo fallback, depuracao ou lacuna registrada.
4. Nao gere JavaScript/PptxGenJS para editar template.
5. Sempre valide depois de mutar.
6. Para entregas visuais, renderize preview com LibreOffice/Poppler e faca QA por imagem.
7. Registre lacunas em `BACKLOG.md` quando faltar comando.

## Superficie MVP

### `inspect`

Entrada:

```powershell
pptx-utils inspect entrada.pptx --format json
```

Saida JSON deve incluir:

- dimensoes do deck;
- quantidade de slides;
- para cada slide: indice, id interno, layout quando disponivel, contagem de shapes, textos encontrados, imagens e tabelas quando detectaveis.

### `text-map`

Entrada:

```powershell
pptx-utils text-map entrada.pptx --format json
```

Saida JSON deve mapear texto editavel por slide e shape. Cada item deve ter identificador estavel o suficiente para ser usado por `replace-text`, preferencialmente com:

- `slideIndex`;
- `shapeIndex`;
- `shapeName` quando existir;
- `text`;
- `paragraphCount`;
- `runCount`;
- flags simples como `hasBullets` e `isPlaceholder` quando detectaveis.

### `replace-text`

Entrada:

```powershell
pptx-utils replace-text entrada.pptx --plan plano.json --output saida.pptx
```

Contrato inicial do plano:

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

O comando deve preservar estilo de texto sempre que a substituicao couber em um shape existente. Quando nao puder preservar com seguranca, deve falhar com erro claro e sem escrever saida parcial.

### Operacoes Estruturais

Comandos MVP:

```powershell
pptx-utils duplicate-slide entrada.pptx --slide 2 --output saida.pptx
pptx-utils delete-slide entrada.pptx --slides 3,5 --output saida.pptx
pptx-utils reorder-slides entrada.pptx --order 1,3,2 --output saida.pptx
```

Esses comandos devem preservar relacoes, notas e recursos associados ao slide quando a biblioteca suportar. Se uma operacao tiver risco de perda, deve falhar explicitamente ou cair para uma implementacao Open XML controlada com testes.

### `replace-image`

Entrada:

```powershell
pptx-utils replace-image entrada.pptx --plan plano.json --output saida.pptx
```

Contrato inicial:

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

### `validate`

Entrada:

```powershell
pptx-utils validate entrada.pptx --format json
```

Deve executar validacao Open XML e checks proprios de integridade minima:

- arquivo abre como PresentationDocument;
- ha `PresentationPart`;
- ha lista de slides;
- relacoes de slides apontam para partes existentes;
- saida JSON informa `openXmlValidationErrors`.

### `render-preview`

Entrada:

```powershell
pptx-utils render-preview entrada.pptx --out previews
```

Esse comando pode delegar para os scripts atuais `soffice.py` e `pdftoppm`. O objetivo e padronizar a chamada dentro da skill, nao reescrever renderizacao.

## Reaproveitamento Da Skill Atual

Conteudo que deve ser incorporado:

- Fluxo de QA visual do `SKILL.md`.
- `thumbnail.py` para visao geral rapida de templates.
- `soffice.py` para conversao headless.
- `unpack.py`, `pack.py` e `validate.py` como fallback/debug.
- Trechos de `editing.md` sobre escolha de layouts, variacao visual, remocao de placeholders e verificacao de overflow.
- `pptxgenjs.md` como referencia secundaria para criacao do zero.

Conteudo que deve ser rebaixado de prioridade:

- Instrucoes que mandam editar `slide{N}.xml` manualmente como fluxo normal.
- Geracao de muito JS para qualquer tarefa que possa ser resolvida com `pptx-utils`.

## Testes

A implementacao deve seguir TDD:

- cada comando novo nasce com teste de CLI;
- comandos mutadores devem provar que a saida abre novamente;
- comandos mutadores devem rodar `validate` no arquivo produzido;
- testes devem criar decks fixture pequenos programaticamente quando possivel;
- fixtures binarias so devem entrar quando forem necessarias para reproduzir template real.

## Criterios De Aceite

- Existe apenas uma skill user-facing para PPTX: `$pptx-utils`.
- `SKILL.md` orienta o binario publicado como caminho padrao.
- A skill atual `pptx` foi incorporada ou substituida sem manter dois pontos de entrada concorrentes.
- `pptx-utils inspect`, `text-map`, `replace-text` e `validate` funcionam pelo binario publicado.
- Pelo menos uma mutacao real em PPTX e testada end-to-end.
- `render-preview` reutiliza a infraestrutura atual de QA visual.
- `pptxgenjs.md` permanece disponivel apenas como fallback para criacao do zero.
- O plano JSON e documentado em `references/plan-contracts.md` e `references/plan-contracts.json`.
- A validacao final roda testes .NET e valida a skill com `quick_validate.py` quando disponivel.

