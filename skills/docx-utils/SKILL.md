---
name: docx-utils
description: Automatiza inspeção, validação, criação e edições auditáveis em arquivos DOCX com .NET/Open XML, incluindo revisões rastreadas, comentários, tabelas, figuras, equações e estilos canônicos.
---

# Docx Utils

## Visão Geral

Use esta skill quando precisar trabalhar com utilitários .NET para inspeção, mutação e validação de documentos DOCX/Open XML, incluindo geração de estilos canônicos e automação de tarefas de dissertação.

## Instalação E Uso

### Regra de execução

1. Use preferencialmente o binário publicado da skill:

   `C:\Users\imale\.codex\skills\docx-utils\bin\docx-utils\docx-utils.exe <comando> <docx> [opções]`

1. Quando o shim estiver disponível no `PATH`, use a forma curta:

   `docx-utils <comando> <docx> [opções]`

1. Não use `dotnet run --project` como caminho padrão de execução; use-o apenas para desenvolvimento, depuração ou quando o binário publicado estiver ausente/quebrado.
1. Se o recurso necessário não funcionar, execute `docx-utils --help` para listar comandos, formas de chamada e exemplos.
1. Se não existir comando para a operação necessária, registre a lacuna em `BACKLOG.md` nesta skill para futura implementação pelo mantenedor/agente mantenedor da própria skill.

### Primeiro uso

1. Na pasta da skill, execute `scripts/install-docx-utils.ps1` no Windows/PowerShell ou `scripts/install-docx-utils.sh` no Linux/WSL.
1. O script valida o SDK do .NET, descobre os projetos `.csproj`, garante as referências NuGet esperadas, faz `restore` e `build`.
1. Se houver testes automatizados, eles são executados por padrão.
1. Se a skill global `skill-creator` estiver instalada, a validação rápida de skill também é executada por padrão.

### Depois de alterar a fonte

1. Reexecute `scripts/install-docx-utils.ps1` ou `scripts/install-docx-utils.sh` na fonte atualizada.
1. Rode `~/.codex/scripts/install-docx-utils-global.ps1` ou `~/.codex/scripts/install-docx-utils-global.sh` para copiar a fonte para `~/.codex/skills/docx-utils` e instalar/restaurar o destino global.
1. Use `-Clean` no instalador global quando quiser remover resíduos da instalação global anterior antes da cópia.
1. Quando necessário, use `-SkipTests`, `-SkipSkillValidation` ou `-NoPackageMutation` para controlar o que o instalador faz.

## Recursos

- `scripts/install-docx-utils.ps1` / `scripts/install-docx-utils.sh`: prepara a skill local, restaura projetos e valida a instalação.
- `~/.codex/scripts/install-docx-utils-global.ps1` / `~/.codex/scripts/install-docx-utils-global.sh`: copia a skill para `~/.codex/skills/docx-utils` e executa a instalação no destino.
- `bin/docx-utils/docx-utils.exe` ou `bin/docx-utils/docx-utils`: binário publicado preferencial para execução operacional.
- `scripts/detect-codex-surface.ps1` / `scripts/detect-codex-surface.sh`: detecta se a sessão atual parece `cli` ou `app`.
- `BACKLOG.md`: registro de lacunas de comandos/recursos para implementação futura.
- `agents/mantenedor-dotnet-docx.md`: papel especializado para manutenção .NET/Open XML com TDD, xUnit, contratos CLI e validação por binário publicado.
- `agents/dotnet-docx-maintainer.yaml`: manifesto de descoberta do agente mantenedor .NET DOCX.
- `references/plan-contracts.md`: contratos minimos e exemplos JSON para `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
- `references/plan-contracts.json`: fonte machine-readable dos contratos operacionais.
- `references/estilos/README.md`: referencia dos estilos canônicos extraidos da dissertação, incluindo `tabelauerj`, `Tabela`, `Figura`, `dados` e `legenda0`.
- Inventário JS/Node: vazio nesta skill. Não há base publicada em `package.json` ou fontes `.js/.ts/.tsx`.

## Agentes

- Regra obrigatória e estrita: a thread principal jamais deve editar código .NET ou JavaScript da própria skill `docx-utils`, incluindo código-fonte, testes, projetos, scripts de instalação/publicação e contratos executáveis. Qualquer alteração desse tipo deve ser implementada exclusivamente por meio do subagent mantenedor `dotnet-docx-maintainer`: a thread principal deve fazer o spawn do subagent e passar a necessidade de implementação para ele, sem tocar diretamente nos arquivos de código. Se o spawn estiver indisponível, registre o bloqueio e não faça a alteração de código.
- A thread principal também não deve criar binários, scripts, projetos auxiliares ou executáveis locais como workaround para contornar uma lacuna da skill. Quando uma capacidade operacional exigir novo código ou novo binário, aguarde o `dotnet-docx-maintainer` implementar, testar, documentar e publicar a mudança pelo fluxo normal da skill.
- Toda implementação feita pelo `dotnet-docx-maintainer` deve manter sincronizados, na mesma rodada, `SKILL.md`, o `README.md` relevante da skill e o texto de `docx-utils --help`/help embutido, além de contratos `plan-contracts` quando aplicável.
- Use `agents/mantenedor-dotnet-docx.md` quando a tarefa envolver implementação, depuração, testes, revisão ou manutenção da fonte .NET em `src/`.
- O agente mantenedor deve combinar as skills `dotnet-cli`, `csharp`, `xunit-tdd`, `openxml-sdk`, `docx-cli-contracts`, `nuget-msbuild`, `published-binary-first` e `cross-platform-installers`.
- Para tarefas operacionais em DOCX, continue seguindo a regra principal desta skill: use o binário publicado `docx-utils` por padrão.

## Superfície Operacional

Use `docx-utils --help` como fonte final da CLI. A superficie publicada real e a mesma exposta pelo binario publicado e pela tabela de dispatch em `Program.cs`.

### Inventário JS/Node

- Vazio nesta skill.
- Nao ha `package.json`, lockfile ou fontes `.js/.ts/.tsx` publicadas para a CLI.

### Entrada e contratos

- `help`, `--help`, `-h`, `/?`: mostram a ajuda da CLI.
- `plan-contracts` e `plan-contract`: expõem os contratos de `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
- `validate-plan`: valida o JSON antes de mutar DOCX; suporta apenas `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
- `create-article`: delega ao comportamento do binario `ArticleDocxBuilder`.
- `create-docx`: cria DOCX vazio ou renderiza um plano JSON.

### Template profiles

- `inspect-template <docx> --out <json> [--report <md>]`: extrai candidatos tecnicos do template.
- `validate-template-profile <profile.json>`: valida o perfil canonico e o hash do template.
- `apply-template --template <docx> --source <docx> --profile <json> --out <docx> [--report <md>]`: aplica um perfil canonico ao documento fonte.
- `audit-template-application <docx> --profile <json> [--report <md>]`: audita o DOCX aplicado contra o perfil canonico.

### Inspeção e auditoria

- `paragraphs <docx> [--start N] [--count N] [--contains TEXT] [--all true|false]`
- `paragraph-detail <docx> --index N [--all true|false]`
- `structure-audit <docx> [--out json]`
- `layout-audit <docx> [--out json] [--report md]`
- `equations-audit <docx> [--out json]`
- `math-audit <docx> [--out json]`
- `math-text-audit <docx> [--out json]`
- `linear-equation-plan-preview <docx> --plan json --out html`
- `revisions <docx> [--author TEXT]`
- `comments <docx> [--author TEXT] [--format auto|table|json|markdown|raw]`
- `comment-anchors <docx>`
- `next-author <docx>`
- `validate <docx>`

### Estilos e tabelas

- `export-used-styles <docx> [--out dir]`
- `ensure-canonical-styles <docx> [--author NAME] --lock <lockfile> [--source dir] [--report md]`
- `sync-styles-from-docx <target.docx> --source-docx <source.docx> [--author NAME] --lock <lockfile> [--report md]`
- `style-running-text <docx> [--author NAME] --lock <lockfile> [--report md]`
- `ensure-style-fonts <docx> [--author NAME] --lock <lockfile> [--font NAME] [--report md]`
- `format-equation-paragraphs <docx> [--author NAME] --lock <lockfile> [--style-id ID] [--seq-name NAME] [--report md]`
- `normalize-figure-indent <docx> [--author NAME] --lock <lockfile> [--report md]`
- `apply-table-design-style <docx> [--author NAME] --lock <lockfile> --style-id ID [--style-name NAME] [--report md]`
- `enable-update-fields-on-open <docx> [--author NAME] --lock <lockfile> [--report md]`
- `disable-update-fields-on-open <docx> [--author NAME] --lock <lockfile> [--report md]`

### Edição com revisões rastreadas

- `insert-tracked <docx> --plan <json> [--author NAME] [--lock <lockfile>] [--report md]`
- `insert-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `replace-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `edit-paragraphs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `append-paragraphs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `accept-revisions <docx> --lock <lockfile> [--disable-track true|false] [--report md]`
- `accept-revisions` aceita insercoes/delecoes rastreadas, repara a validacao antes de salvar e preserva `--disable-track` quando solicitado.
- Em mutacoes, `--author` continua opcional para a thread principal; a omissao ativa o autor automatico.

### Figuras, formulas e referencias

- `insert-figures <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `replace-figures-from-plan <docx> --plan json [--author NAME] --lock <lockfile> [--report md]`
- `rewrite-equation-blocks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `replace-formulas-with-linear-equations <docx> --plan json [--author NAME] --lock <lockfile> [--keep-linear true|false] [--report md]`
- `replace-formulas-with-mathml-omml <docx> --plan json [--author NAME] --lock <lockfile> [--xsl MML2OMML.XSL] [--report md]`
- `convert-text-formulas-to-omath <docx> --plan json [--author NAME] --lock <lockfile> [--report md]`
- `apply-crossrefs <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `add-bookmarks <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `rewrite-ref-fields <docx> [--author NAME] --lock <lockfile> --bookmark-prefixes CSV --template TEXT [--report md]`

### Comentarios

- `insert-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `reanchor-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `answer-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `reply-comments <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `remove-comments <docx> --ids CSV|all [--author NAME] --lock <lockfile> [--report md]`

### Reparos e finalizacao

- `repair-article-abnt-layout <docx> [--author NAME] --lock <lockfile> [--report md]`
- `format-abnt-reference-titles <docx> [--author NAME] --lock <lockfile> [--target publication|article|both] [--emphasis italic|bold] [--report md]`
- `repair-style-captions <docx> --plan <json> [--author NAME] --lock <lockfile> [--report md]`
- `repair-layout-pendencies <docx> [--author NAME] --lock <lockfile> [--report md]`
- `repair-ref-number-only <docx> [--author NAME] --lock <lockfile> [--report md]`

### Observacoes

- Comandos mutadores exigem `--lock <lockfile>` para escrita exclusiva.
- `comments` aceita `--format auto|table|json|markdown|raw`; sem `--format`, a saida se adapta ao ambiente CLI/app.
- `plan-contracts` e `validate-plan` sao os contratos publicados para planos JSON; nao use `Program.cs` como fonte do formato.
- A criacao de linhas, celulas e tabelas OpenXML e responsabilidade da skill; em uso operacional, declare o plano e nao monte `w:tr`/`w:tc` manualmente.

## Perfis De Template

- `inspect-template <docx> --out <json> [--report <md>]` não classifica semanticamente o documento de forma final; ele gera candidatos técnicos para curadoria pelo agente.
- Use `inspect-template` quando precisar avaliar um template ou documento por sinais de formatação, inclusive possíveis seções e subseções.
- O JSON/relatório de candidatos inclui, por parágrafo: `id`, texto, parte do documento (`body`, `table-cell`, `header`, `footer`, `footnote`, `endnote`), ordinal, estilo Word, alinhamento, espaçamento antes/depois, recuo pendente, runs com negrito/itálico/caixa alta/tamanho de fonte, numeração manual, vizinhos e pistas estruturais.
- As pistas estruturais publicadas atualmente são `manualNumbering`, `shortHighlightedParagraph` e `looksLikeReference`.
- Para seções e subseções, trate como candidatos fortes os parágrafos curtos com numeração manual (`1`, `1.1`, `1.1.1`), negrito, caixa alta, itálico, tamanho de fonte diferenciado, espaçamento destacado ou estilo de título. A decisão final sobre `section`, `subsection`, `title`, `abstract`, `references` etc. é responsabilidade do agente ao montar o `profile.canonical.json`.
- `apply-template --template <docx> --source <docx> --profile <json> --out <docx> [--report <md>]` aplica apenas as regiões declaradas no profile canônico. Na versão atual, esse fluxo é útil para regiões curadas como título, resumo e referências, mas não substitui automaticamente o corpo inteiro de um DOCX fonte dentro de um template.

## Planos de blocos e tabelas

- `docx-utils plan-contracts [comando] [--format markdown|json]` expõe os contratos operacionais sem depender do código-fonte.
- `docx-utils validate-plan <comando> --plan <json>` valida o contrato antes de mutar o DOCX.
- `create-article` delega ao comportamento exato do binario `ArticleDocxBuilder`.
- `create-docx` cria um DOCX vazio quando chamado sem plano e renderiza um plano JSON quando informado.
- `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table` devem ser consultados via `plan-contracts`, `references/plan-contracts.md` ou `references/plan-contracts.json`; não dependa de `Program.cs` para descobrir o formato.
- `insert-blocks` insere blocos entre `afterPrefix` e `beforePrefix` sem remover o conteúdo existente entre eles.
- `replace-blocks` remove o intervalo entre as âncoras e insere os blocos declarativos no lugar.
- `create-docx` exige `title` e `paragraphs` quando o plano tem conteúdo; `subtitles`, `sections` e `references` são opcionais.
- `insert-blocks` e `replace-blocks` exigem `blocks[]`, `afterPrefix`, `beforePrefix`, `items[]`, `kind` valido e `rows` quando o item for tabela.
- `replace-table` exige `tables[]`, um seletor valido e `rows` nao vazio.
- Criar linhas, celulas e tabelas OpenXML e responsabilidade da skill; em uso operacional, o Codex deve declarar o plano e nao montar `w:tr`/`w:tc` manualmente.

## Detecção Codex CLI/App

- Antes de escolher formatos voltados à interface, detecte a superfície atual, porque a mesma conversa pode alternar entre Codex CLI e Codex app.
- Use este comando quando houver dúvida:

  `powershell -ExecutionPolicy Bypass -File C:\Users\imale\.codex\skills\docx-utils\scripts\detect-codex-surface.ps1`

- No Linux/WSL, use:

  `~/.codex/skills/docx-utils/scripts/detect-codex-surface.sh`

- A detecção usa apenas `CODEX_MANAGED_BY_NPM`: quando `CODEX_MANAGED_BY_NPM=1`, assume `cli`; caso contrário, assume `app`.
- Se o usuário informar explicitamente que a sessão atual é CLI ou app, trate essa informação como override para a rodada atual.

## Comentários DOCX

- `docx-utils comments <docx>` sem `--format` usa detecção automática:
  - com `CODEX_MANAGED_BY_NPM=1`, retorna tabela de terminal para leitura humana;
  - sem `CODEX_MANAGED_BY_NPM=1`, retorna tabela Markdown.
- `docx-utils comments <docx> --format auto` segue a mesma regra e deve ser tratado como formato forçado pela superfície: `cli` gera `table`; `app` gera `markdown`.
- Ao repassar a saída ao usuário, preserve exatamente a saída textual do executável; não resuma, não converta, não reordene, não renderize de outro modo e não edite a tabela.
- Use `docx-utils comments <docx> --format json` quando a tarefa pedir JSON/dados estruturados, automação, parsing, depuração da saída bruta ou validação do contrato do binário.
- Use `--format raw` apenas quando precisar da saída textual legada.
- Quando o usuário enviar de volta uma tabela editada de comentários, trate a coluna `orientacao` como instrução operacional para cada `id` de comentário:
  - `Resolver`, `resolva`, `corrigir`, ou texto equivalente: aplique a correção solicitada pelo conteúdo do comentário ao DOCX, depois responda/remova/marque conforme a tarefa pedir.
  - `apagar este comentário`, `remover comentário`, ou texto equivalente: remova apenas o comentário indicado, sem alterar o texto do documento salvo se a orientação disser isso explicitamente.
  - Texto livre na coluna `orientacao`: siga essa orientação específica para o comentário daquele `id`.
- Antes de agir sobre uma tabela editada, releia o DOCX no disco com `docx-utils comments <docx> --format json` e confira se os `id` ainda existem; se algum `id` não existir mais, informe o conflito.

## Autoria Em Mutações DOCX

- Na thread principal, omita `--author` em comandos mutadores; o `docx-utils` escolhe automaticamente o próximo autor disponível no DOCX.
- A lista automática usada pela thread principal é: `Ultron`, `Brainiac`, `Jarvis`, `Vision`, `HumanTorch`, `Friday`, `C3PO`, `R2D2`.
- Se todos os nomes-base já existirem, o utilitário tenta a mesma lista com sufixos numéricos (`Ultron-1`, `Brainiac-1`, etc.).
- Use `docx-utils next-author <docx>` quando quiser apenas descobrir, sem mutar, qual será o próximo autor livre.
- Em subagents, informe sempre `--author` com o nome atribuído ao subagent.
