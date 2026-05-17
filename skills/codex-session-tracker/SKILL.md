---
name: codex-session-tracker
description: Rastreia sessoes Codex locais, processos vivos e workers/subagents usando hook SessionStart, SQLite em ~/.codex e leitura dos rollout-*.jsonl. Use quando o usuario quiser listar sessoes em execucao, descobrir parent_thread_id, agent_nickname, workers, resumir/inspecionar threads, retomar uma thread com codex resume, ou configurar um hook para registrar sessoes e subagents automaticamente.
---

# Codex Session Tracker

Use esta skill para consultar o estado local das threads Codex e retomar uma thread principal ou worker pelo `session_id`.

O recurso principal fica em `scripts/script.py`. Ele cria e atualiza o banco `~/.codex/codex-session-tracker.sqlite`.

## Fluxo Padrao

1. Sincronizar rollouts quando precisar reconstruir historico:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" sync
```

2. Listar sessoes principais:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" list --sync
```

3. Listar apenas sessoes vivas, excluindo a sessao Codex atual por padrao:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" list --sync --live
```

4. Ver workers/subagents de uma thread:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" workers <session_id> --sync
```

5. Retomar uma thread principal ou worker:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" resume <session_id>
```

Para conferir antes de executar:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" resume <session_id> --print-only
```

## Hook SessionStart

O hook grava a sessao no SQLite no momento em que ela nasce. Se o evento trouxer dados de subagent, ele registra:

- `session_id`
- `cwd`
- `project_name`
- `parent_thread_id`
- `agent_nickname`
- `agent_role`
- PID do hook e, quando detectavel, PID do processo `codex.exe`/`codex` ancestral

Gerar o bloco TOML:

```powershell
python "$env:USERPROFILE\.codex\skills\codex-session-tracker\scripts\script.py" hook-snippet
```

O bloco deve ser adicionado a `~/.codex/config.toml` junto aos demais hooks `StartSession`.

## Como o Rastreamento Funciona

- O hook `session-start` e a tabela SQLite sao a fonte primaria para novas sessoes.
- `sync` le `~/.codex/sessions/**/rollout-*.jsonl` e completa dados que nao vieram no hook.
- Workers sao encontrados por `session_meta.payload.source.subagent.thread_spawn.parent_thread_id`.
- A lista de sessoes vivas cruza os registros com processos reais chamados `codex.exe` ou `codex`.
- A sessao Codex atual e excluida da lista viva por padrao; use `--include-current` quando precisar ve-la.

## Comandos

```powershell
python scripts/script.py sync [--limit N]
python scripts/script.py list [--sync] [--live] [--all] [--include-current]
python scripts/script.py workers <session_id> [--sync]
python scripts/script.py show <session_id> [--sync]
python scripts/script.py resume <session_id> [--print-only]
python scripts/script.py hook session-start
python scripts/script.py hook-snippet
```

## Limites

- Retomar worker significa abrir a thread do worker com `codex resume <worker_session_id>`.
- O hook depende do payload entregue pela versao atual do Codex; quando `parent_thread_id` nao vier no hook, `sync` reconstrui a relacao pelos rollouts.
- Sem o hook instalado, a skill ainda funciona usando `sync`, mas so depois que os rollouts existirem.
