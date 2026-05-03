---
name: docx-utils
description: Instala e valida utilitários .NET/Open XML para inspeção, edição auditável e exportação de estilos de DOCX.
---

# Docx Utils

## Visão Geral

Use esta skill quando precisar trabalhar com utilitários .NET para inspeção, mutação e validação de documentos DOCX/Open XML, incluindo geração de estilos canônicos e automação de tarefas de dissertação.

## Instalação E Uso

### Primeiro uso

1. Na pasta da skill, execute `scripts/install-docx-utils.ps1`.
1. O script valida o SDK do .NET, descobre os projetos `.csproj`, garante as referências NuGet esperadas, faz `restore` e `build`.
1. Se houver testes automatizados, eles são executados por padrão.
1. Se a skill global `skill-creator` estiver instalada, a validação rápida de skill também é executada por padrão.

### Depois de alterar a fonte

1. Reexecute `scripts/install-docx-utils.ps1` na fonte atualizada.
1. Rode `~/.codex/scripts/install-docx-utils-global.ps1` para copiar a fonte para `~/.codex/skills/docx-utils` e instalar/restaurar o destino global.
1. Use `-Clean` no instalador global quando quiser remover resíduos da instalação global anterior antes da cópia.
1. Quando necessário, use `-SkipTests`, `-SkipSkillValidation` ou `-NoPackageMutation` para controlar o que o instalador faz.

## Recursos

- `scripts/install-docx-utils.ps1`: prepara a skill local, restaura projetos e valida a instalação.
- `~/.codex/scripts/install-docx-utils-global.ps1`: copia a skill para `~/.codex/skills/docx-utils` e executa a instalação no destino.
