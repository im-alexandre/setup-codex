# Codex Home

Este diretório é o `CODEX_HOME` versionável desta máquina.

## O Que Fica Versionado

- `AGENTS.md`
- `config.toml`
- `agents/`
- `hooks/`
- `mcp/`
- `overlays/`
- `preset/`
- `presets/`
- `rules/`
- `scripts/`
- `skills/`
- `docker-compose.yml`
- `mcps.toml`

## O Que Não Deve Ser Versionado

O `.gitignore` exclui estado local, segredos, sessões, logs, bancos SQLite, caches, plugins baixados, ambientes virtuais, artefatos de build e backups.

Arquivos como `auth.json`, `installation_id`, `logs_2.sqlite*`, `state_5.sqlite*`, `sessions/`, `plugins/`, `.tmp/` e `cache/` são runtime local do Codex e devem ficar fora do Git.

## Validação

```powershell
python -c "import pathlib, tomllib; tomllib.loads(pathlib.Path('config.toml').read_text(encoding='utf-8')); print('config.toml ok')"
git status --ignored --short
```

## Observação Sobre Arquivos Bloqueados

Enquanto o Codex estiver aberto, alguns bancos SQLite podem ficar bloqueados e não podem ser movidos ou removidos. Para limpar fisicamente esses arquivos, feche todas as sessões Codex e remova:

```powershell
Remove-Item -LiteralPath "$env:USERPROFILE\.codex\logs_2.sqlite*" -Force
Remove-Item -LiteralPath "$env:USERPROFILE\.codex\state_5.sqlite*" -Force
Remove-Item -LiteralPath "$env:USERPROFILE\.codex\sqlite" -Recurse -Force
```
