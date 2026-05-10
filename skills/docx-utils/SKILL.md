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
- `references/plan-contracts.md`: contratos minimos e exemplos JSON para `create-docx`, `insert-blocks`, `replace-blocks` e `replace-table`.
- `references/plan-contracts.json`: fonte machine-readable dos contratos operacionais.

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
- Em subagents, informe sempre `--author` com o nome atribuído ao subagent.
