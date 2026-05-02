# codex-setup

Repositório-fonte para instalar a base local do Codex em `%USERPROFILE%\.codex`.

Este repo não é o runtime. A fonte versionada fica em `files/`; o runtime instalado fica em `%USERPROFILE%\.codex`.

## Estrutura

```text
install_codex.ps1
files/
  AGENTS.md
  config.template.toml
  docker-compose.yml
  agents/
  hooks/
  mcp/
  presets/
    base-project.toml
    mcp-services.json
    skill-services.json
    mcp/
      mendeley.template.toml
      openai-docs.template.toml
      pdf-indexer.template.toml
      zotero.template.toml
  scripts/
    bootstrap-codex-project.ps1
    resolve-codex-services.ps1
    start-codex-services.ps1
    start-mendeley-mcp.ps1
    start-pdf-indexer-mcp.ps1
  skills/
```

## Instalação

```powershell
.\install_codex.ps1
```

O instalador:

- descobre o root pelo caminho do próprio `install_codex.ps1`;
- copia arquivo a arquivo de `files/` para `%USERPROFILE%\.codex`;
- cria backup `.bak-YYYYMMDD-HHMMSS` antes de sobrescrever qualquer arquivo existente;
- converte cada `*.template.toml` instalado para `.toml`;
- substitui `{{CODEX_HOME}}` por um caminho absoluto com `/` nos templates TOML;
- ignora caches, bancos locais, logs, ambientes virtuais e artefatos temporários;
- ajusta `presets/mcp-services.json` e `presets/skill-services.json` para `ollama-cpu` quando não há GPU NVIDIA detectada;
- mantém `ollama-gpu` como padrão quando a máquina tem NVIDIA.

## MCPs e skills

MCPs disponíveis por preset:

- `mendeley`
- `openai-docs`
- `pdf-indexer`
- `zotero`

`scopus-search` não é MCP. Ele é instalado e usado como skill.

Mapas de serviços:

- `files/presets/mcp-services.json`: serviços Docker usados pelos MCPs.
- `files/presets/skill-services.json`: serviços Docker usados pelas skills.

`scopus-search` aparece apenas em `skill-services.json`, e o resolver trata esse nome como skill, não como MCP.

Os dois mapas podem apontar para os mesmos serviços Docker, mas a origem do nome é diferente: o bootstrap recebe os MCPs selecionados, enquanto `resolve-codex-services.ps1` e `start-codex-services.ps1` cuidam dos mapas de serviço. Nesse fluxo, `mcp-services.json` cobre MCPs e `skill-services.json` cobre skills.

## Serviços

Resolver serviços:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\scripts\resolve-codex-services.ps1" -SkillNames scopus-search
```

Subir serviços:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\scripts\start-codex-services.ps1" -SkillNames scopus-search
```

O compose instalado fica em `%USERPROFILE%\.codex\docker-compose.yml`.

## Bootstrap de projeto

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\scripts\bootstrap-codex-project.ps1" -ProjectPath "D:\algum-projeto"
```

O bootstrap:

- recebe os MCPs selecionados pelo usuário;
- monta ou atualiza `.codex\config.toml` no projeto;
- cria backup local de `config.toml` antes de sobrescrever;
- resolve os serviços exigidos pelos MCPs selecionados;
- sobe automaticamente os serviços Docker necessários quando há serviços resolvidos;
- não edita `%USERPROFILE%\.codex\config.toml`.

## Validação

```powershell
git diff --check
powershell -NoProfile -Command "[System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Get-Location) 'install_codex.ps1'),[ref]$null,[ref]$null).Errors.Count"
powershell -NoProfile -Command "Get-Content (Join-Path (Get-Location) 'files/presets/mcp-services.json') -Raw | ConvertFrom-Json | Out-Null"
powershell -NoProfile -Command "Get-Content (Join-Path (Get-Location) 'files/presets/skill-services.json') -Raw | ConvertFrom-Json | Out-Null"
docker compose -f files\docker-compose.yml config --quiet
```
