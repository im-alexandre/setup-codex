#!/usr/bin/env python3
"""Rastreador local de sessoes Codex e subagents.

Mantem um indice SQLite em ~/.codex/codex-session-tracker.sqlite e usa os
rollouts como fonte auditavel para completar dados que o hook nao recebeu.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import platform
import re
import sqlite3
import subprocess
import sys
from pathlib import Path
from typing import Any


CODEX_HOME = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex")).expanduser()
DB_PADRAO = CODEX_HOME / "codex-session-tracker.sqlite"
SESSIONS_DIR = CODEX_HOME / "sessions"
SESSION_RE = re.compile(r"rollout-.+-(?P<id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\.jsonl$", re.I)


def agora_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def parse_data(valor: Any) -> dt.datetime | None:
    if not valor:
        return None
    if isinstance(valor, (int, float)):
        return dt.datetime.fromtimestamp(valor, tz=dt.timezone.utc)
    texto = str(valor).strip()
    if not texto:
        return None
    if texto.endswith("Z"):
        texto = texto[:-1] + "+00:00"
    try:
        data = dt.datetime.fromisoformat(texto)
    except ValueError:
        return None
    if data.tzinfo is None:
        data = data.replace(tzinfo=dt.datetime.now().astimezone().tzinfo)
    return data.astimezone(dt.timezone.utc)


def iso_utc(valor: Any) -> str | None:
    data = parse_data(valor)
    return data.isoformat() if data else None


def projeto_de(cwd: str | None) -> str:
    if not cwd:
        return ""
    caminho = Path(cwd)
    return caminho.name or str(caminho)


def conectar(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sessions (
            session_id TEXT PRIMARY KEY,
            cwd TEXT,
            project_name TEXT,
            transcript_path TEXT,
            rollout_path TEXT,
            model TEXT,
            source_json TEXT,
            thread_source TEXT,
            parent_thread_id TEXT,
            agent_nickname TEXT,
            agent_role TEXT,
            is_subagent INTEGER NOT NULL DEFAULT 0,
            started_at TEXT,
            last_seen_at TEXT,
            hook_pid INTEGER,
            codex_pid INTEGER,
            codex_started_at TEXT,
            platform TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sessions_parent ON sessions(parent_thread_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_sessions_cwd ON sessions(cwd)")
    return conn


def upsert_session(conn: sqlite3.Connection, dados: dict[str, Any]) -> None:
    session_id = dados.get("session_id")
    if not session_id:
        return
    existente = conn.execute("SELECT * FROM sessions WHERE session_id = ?", (session_id,)).fetchone()
    agora = agora_iso()
    base = dict(existente) if existente else {}

    def escolher(chave: str) -> Any:
        valor = dados.get(chave)
        if valor is None or valor == "":
            return base.get(chave)
        return valor

    row = {
        "session_id": session_id,
        "cwd": escolher("cwd"),
        "project_name": escolher("project_name") or projeto_de(escolher("cwd")),
        "transcript_path": escolher("transcript_path"),
        "rollout_path": escolher("rollout_path"),
        "model": escolher("model"),
        "source_json": escolher("source_json"),
        "thread_source": escolher("thread_source"),
        "parent_thread_id": escolher("parent_thread_id"),
        "agent_nickname": escolher("agent_nickname"),
        "agent_role": escolher("agent_role"),
        "is_subagent": 1 if escolher("parent_thread_id") else int(escolher("is_subagent") or 0),
        "started_at": escolher("started_at"),
        "last_seen_at": escolher("last_seen_at") or agora,
        "hook_pid": escolher("hook_pid"),
        "codex_pid": escolher("codex_pid"),
        "codex_started_at": escolher("codex_started_at"),
        "platform": escolher("platform") or platform.system(),
        "created_at": base.get("created_at") or agora,
        "updated_at": agora,
    }
    conn.execute(
        """
        INSERT INTO sessions (
            session_id, cwd, project_name, transcript_path, rollout_path, model,
            source_json, thread_source, parent_thread_id, agent_nickname,
            agent_role, is_subagent, started_at, last_seen_at, hook_pid,
            codex_pid, codex_started_at, platform, created_at, updated_at
        ) VALUES (
            :session_id, :cwd, :project_name, :transcript_path, :rollout_path, :model,
            :source_json, :thread_source, :parent_thread_id, :agent_nickname,
            :agent_role, :is_subagent, :started_at, :last_seen_at, :hook_pid,
            :codex_pid, :codex_started_at, :platform, :created_at, :updated_at
        )
        ON CONFLICT(session_id) DO UPDATE SET
            cwd=excluded.cwd,
            project_name=excluded.project_name,
            transcript_path=excluded.transcript_path,
            rollout_path=excluded.rollout_path,
            model=excluded.model,
            source_json=excluded.source_json,
            thread_source=excluded.thread_source,
            parent_thread_id=excluded.parent_thread_id,
            agent_nickname=excluded.agent_nickname,
            agent_role=excluded.agent_role,
            is_subagent=excluded.is_subagent,
            started_at=excluded.started_at,
            last_seen_at=excluded.last_seen_at,
            hook_pid=excluded.hook_pid,
            codex_pid=excluded.codex_pid,
            codex_started_at=excluded.codex_started_at,
            platform=excluded.platform,
            updated_at=excluded.updated_at
        """,
        row,
    )


def comando_json_powershell(script: str) -> Any:
    resultado = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        text=True,
        capture_output=True,
        check=False,
    )
    if resultado.returncode != 0 or not resultado.stdout.strip():
        return []
    try:
        return json.loads(resultado.stdout)
    except json.JSONDecodeError:
        return []


def processos_windows() -> list[dict[str, Any]]:
    script = r"""
$rows = Get-CimInstance Win32_Process | ForEach-Object {
  [pscustomobject]@{
    pid = [int]$_.ProcessId
    ppid = [int]$_.ParentProcessId
    name = [string]$_.Name
    executable = [string]$_.ExecutablePath
    command_line = [string]$_.CommandLine
    created_at = if ($_.CreationDate) { $_.CreationDate.ToUniversalTime().ToString("o") } else { $null }
  }
}
$rows | ConvertTo-Json -Compress
"""
    dados = comando_json_powershell(script)
    if isinstance(dados, dict):
        return [dados]
    return dados if isinstance(dados, list) else []


def processos_linux() -> list[dict[str, Any]]:
    proc = Path("/proc")
    if not proc.exists():
        return []
    try:
        btime = 0
        for linha in Path("/proc/stat").read_text(encoding="utf-8", errors="ignore").splitlines():
            if linha.startswith("btime "):
                btime = int(linha.split()[1])
                break
        ticks = os.sysconf(os.sysconf_names["SC_CLK_TCK"])
    except Exception:
        btime = 0
        ticks = 100
    rows: list[dict[str, Any]] = []
    for item in proc.iterdir():
        if not item.name.isdigit():
            continue
        try:
            stat = (item / "stat").read_text(encoding="utf-8", errors="ignore")
            cmd = (item / "cmdline").read_bytes().replace(b"\x00", b" ").decode("utf-8", errors="ignore").strip()
            exe = os.readlink(item / "exe") if (item / "exe").exists() else ""
            partes = stat.split()
            ppid = int(partes[3])
            inicio = int(partes[21])
            criado = dt.datetime.fromtimestamp(btime + inicio / ticks, tz=dt.timezone.utc).isoformat() if btime else None
            rows.append(
                {
                    "pid": int(item.name),
                    "ppid": ppid,
                    "name": Path(exe).name or partes[1].strip("()"),
                    "executable": exe,
                    "command_line": cmd,
                    "created_at": criado,
                }
            )
        except Exception:
            continue
    return rows


def todos_processos() -> list[dict[str, Any]]:
    if platform.system().lower().startswith("win"):
        return processos_windows()
    return processos_linux()


def eh_codex_bin(proc: dict[str, Any]) -> bool:
    nome = (proc.get("name") or "").lower()
    exe = (proc.get("executable") or "").replace("\\", "/").lower()
    return nome in {"codex.exe", "codex"} or exe.endswith("/codex.exe") or exe.endswith("/codex")


def ancestrais(pid: int, processos: list[dict[str, Any]]) -> set[int]:
    por_pid = {int(p["pid"]): p for p in processos if p.get("pid") is not None}
    vistos: set[int] = set()
    atual = pid
    while atual and atual not in vistos and atual in por_pid:
        vistos.add(atual)
        atual = int(por_pid[atual].get("ppid") or 0)
    return vistos


def processos_codex(include_current: bool = False) -> list[dict[str, Any]]:
    processos = todos_processos()
    current_ancestors = ancestrais(os.getpid(), processos)
    rows = [p for p in processos if eh_codex_bin(p)]
    if not include_current:
        rows = [p for p in rows if int(p.get("pid") or 0) not in current_ancestors]
    return rows


def codex_ancestral_atual() -> dict[str, Any] | None:
    processos = todos_processos()
    por_pid = {int(p["pid"]): p for p in processos if p.get("pid") is not None}
    for pid in ancestrais(os.getpid(), processos):
        proc = por_pid.get(pid)
        if proc and eh_codex_bin(proc):
            return proc
    return None


def extrair_parent_source(source: Any) -> dict[str, Any]:
    if not isinstance(source, dict):
        return {}
    spawn = source.get("subagent", {}).get("thread_spawn", {})
    if not isinstance(spawn, dict):
        return {}
    return {
        "parent_thread_id": spawn.get("parent_thread_id"),
        "agent_nickname": spawn.get("agent_nickname"),
        "agent_role": spawn.get("agent_role"),
    }


def dados_de_evento_hook(evento: dict[str, Any]) -> dict[str, Any]:
    payload = evento.get("payload") if isinstance(evento.get("payload"), dict) else evento
    source = payload.get("source")
    parent = extrair_parent_source(source)
    session_id = payload.get("session_id") or payload.get("id")
    cwd = payload.get("cwd")
    proc = codex_ancestral_atual()
    return {
        "session_id": session_id,
        "cwd": cwd,
        "project_name": projeto_de(cwd),
        "transcript_path": payload.get("transcript_path"),
        "model": payload.get("model") or payload.get("model_provider"),
        "source_json": json.dumps(source, ensure_ascii=False) if source is not None else None,
        "thread_source": payload.get("thread_source"),
        "parent_thread_id": parent.get("parent_thread_id") or payload.get("parent_thread_id"),
        "agent_nickname": parent.get("agent_nickname") or payload.get("agent_nickname"),
        "agent_role": parent.get("agent_role") or payload.get("agent_role"),
        "is_subagent": 1 if parent.get("parent_thread_id") or payload.get("parent_thread_id") else 0,
        "started_at": iso_utc(payload.get("timestamp") or payload.get("started_at")),
        "last_seen_at": agora_iso(),
        "hook_pid": os.getpid(),
        "codex_pid": int(proc["pid"]) if proc else None,
        "codex_started_at": iso_utc(proc.get("created_at")) if proc else None,
        "platform": platform.system(),
    }


def iter_rollouts() -> list[Path]:
    if not SESSIONS_DIR.exists():
        return []
    return sorted(SESSIONS_DIR.rglob("rollout-*.jsonl"), key=lambda p: p.stat().st_mtime, reverse=True)


def resumo_rollout(path: Path) -> dict[str, Any] | None:
    session_id = None
    meta: dict[str, Any] = {}
    completed = False
    last_event_at = None
    spawn_calls = 0
    wait_calls = 0
    try:
        with path.open("r", encoding="utf-8", errors="replace") as fh:
            for linha in fh:
                try:
                    obj = json.loads(linha)
                except json.JSONDecodeError:
                    continue
                ts = obj.get("timestamp")
                if ts:
                    last_event_at = iso_utc(ts) or last_event_at
                tipo = obj.get("type")
                payload = obj.get("payload") if isinstance(obj.get("payload"), dict) else {}
                if tipo == "session_meta":
                    meta = payload
                    session_id = payload.get("id")
                elif tipo == "event_msg" and payload.get("type") == "task_complete":
                    completed = True
                elif tipo == "response_item":
                    nome = payload.get("name")
                    if nome == "spawn_agent":
                        spawn_calls += 1
                    elif nome == "wait_agent":
                        wait_calls += 1
    except OSError:
        return None
    if not session_id:
        match = SESSION_RE.search(path.name)
        session_id = match.group("id") if match else None
    if not session_id:
        return None
    parent = extrair_parent_source(meta.get("source"))
    return {
        "session_id": session_id,
        "cwd": meta.get("cwd"),
        "project_name": projeto_de(meta.get("cwd")),
        "rollout_path": str(path),
        "model": meta.get("model_provider"),
        "source_json": json.dumps(meta.get("source"), ensure_ascii=False) if meta.get("source") is not None else None,
        "thread_source": meta.get("thread_source"),
        "parent_thread_id": parent.get("parent_thread_id"),
        "agent_nickname": parent.get("agent_nickname") or meta.get("agent_nickname"),
        "agent_role": parent.get("agent_role") or meta.get("agent_role"),
        "is_subagent": 1 if parent.get("parent_thread_id") else 0,
        "started_at": iso_utc(meta.get("timestamp")),
        "last_seen_at": last_event_at or dt.datetime.fromtimestamp(path.stat().st_mtime, tz=dt.timezone.utc).isoformat(),
        "task_complete": completed,
        "spawn_calls": spawn_calls,
        "wait_calls": wait_calls,
    }


def sincronizar_rollouts(conn: sqlite3.Connection, limite: int | None = None) -> int:
    total = 0
    for path in iter_rollouts()[:limite]:
        resumo = resumo_rollout(path)
        if not resumo:
            continue
        upsert_session(conn, resumo)
        total += 1
    conn.commit()
    return total


def carregar_sessoes(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = conn.execute("SELECT * FROM sessions ORDER BY COALESCE(last_seen_at, started_at) DESC").fetchall()
    return [dict(r) for r in rows]


def correlacionar_live(conn: sqlite3.Connection, include_current: bool = False) -> dict[str, dict[str, Any]]:
    sessoes = carregar_sessoes(conn)
    mapa: dict[str, dict[str, Any]] = {}
    for proc in processos_codex(include_current=include_current):
        proc_dt = parse_data(proc.get("created_at"))
        escolhido = None
        melhor = 10**9
        for sessao in sessoes:
            if sessao.get("parent_thread_id"):
                continue
            if sessao.get("codex_pid") and int(sessao["codex_pid"]) == int(proc["pid"]):
                escolhido = sessao
                break
            inicio = parse_data(sessao.get("started_at"))
            if proc_dt and inicio:
                delta = abs((proc_dt - inicio).total_seconds())
                if delta < melhor:
                    melhor = delta
                    escolhido = sessao
        if escolhido and melhor <= 180:
            mapa[escolhido["session_id"]] = proc
    return mapa


def largura(texto: Any, n: int) -> str:
    valor = "" if texto is None else str(texto)
    return valor if len(valor) <= n else valor[: n - 3] + "..."


def imprimir_tabela(rows: list[dict[str, Any]], colunas: list[tuple[str, str, int]]) -> None:
    if not rows:
        print("Nenhum registro encontrado.")
        return
    cab = "  ".join(largura(titulo, largura_col).ljust(largura_col) for _, titulo, largura_col in colunas)
    print(cab)
    print("  ".join("-" * largura_col for _, _, largura_col in colunas))
    for row in rows:
        print("  ".join(largura(row.get(chave), largura_col).ljust(largura_col) for chave, _, largura_col in colunas))


def cmd_hook(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    raw = sys.stdin.read().strip()
    if not raw:
        return 0
    try:
        evento = json.loads(raw)
    except json.JSONDecodeError:
        return 0
    dados = dados_de_evento_hook(evento)
    upsert_session(conn, dados)
    conn.commit()
    if args.print_response:
        print(json.dumps({"hookSpecificOutput": {"hookEventName": "SessionStart"}}))
    return 0


def cmd_sync(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    total = sincronizar_rollouts(conn, args.limit)
    print(f"Sincronizados {total} rollouts em {args.db}")
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    if args.sync:
        sincronizar_rollouts(conn, args.limit)
    live = correlacionar_live(conn, include_current=args.include_current)
    rows = []
    for sessao in carregar_sessoes(conn):
        if sessao.get("parent_thread_id") and not args.all:
            continue
        if args.live and sessao["session_id"] not in live:
            continue
        proc = live.get(sessao["session_id"])
        workers = conn.execute("SELECT COUNT(*) AS n FROM sessions WHERE parent_thread_id = ?", (sessao["session_id"],)).fetchone()["n"]
        rows.append(
            {
                "status": "viva" if proc else "historico",
                "pid": proc.get("pid") if proc else "",
                "project": sessao.get("project_name"),
                "session": sessao.get("session_id"),
                "workers": workers,
                "cwd": sessao.get("cwd"),
                "last_seen": sessao.get("last_seen_at") or sessao.get("started_at"),
            }
        )
    imprimir_tabela(
        rows,
        [
            ("status", "status", 9),
            ("pid", "pid", 7),
            ("project", "projeto", 18),
            ("workers", "workers", 7),
            ("session", "thread", 38),
            ("last_seen", "ultimo", 24),
            ("cwd", "cwd", 42),
        ],
    )
    return 0


def cmd_workers(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    if args.sync:
        sincronizar_rollouts(conn, args.limit)
    rows = []
    for row in conn.execute(
        "SELECT * FROM sessions WHERE parent_thread_id = ? ORDER BY COALESCE(last_seen_at, started_at) DESC",
        (args.session_id,),
    ):
        sessao = dict(row)
        resumo = resumo_rollout(Path(sessao["rollout_path"])) if sessao.get("rollout_path") else None
        rows.append(
            {
                "worker": sessao.get("agent_nickname") or "",
                "role": sessao.get("agent_role") or "",
                "status": "complete" if resumo and resumo.get("task_complete") else "aberto",
                "thread": sessao.get("session_id"),
                "last_seen": sessao.get("last_seen_at") or sessao.get("started_at"),
            }
        )
    imprimir_tabela(
        rows,
        [
            ("worker", "worker", 18),
            ("role", "role", 12),
            ("status", "status", 10),
            ("thread", "thread", 38),
            ("last_seen", "ultimo", 24),
        ],
    )
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    if args.sync:
        sincronizar_rollouts(conn, args.limit)
    row = conn.execute("SELECT * FROM sessions WHERE session_id = ?", (args.session_id,)).fetchone()
    if not row:
        print(f"Thread nao encontrada: {args.session_id}", file=sys.stderr)
        return 1
    sessao = dict(row)
    print(json.dumps(sessao, indent=2, ensure_ascii=False))
    return 0


def cmd_resume(args: argparse.Namespace) -> int:
    conn = conectar(Path(args.db))
    row = conn.execute("SELECT * FROM sessions WHERE session_id = ?", (args.session_id,)).fetchone()
    if not row:
        print(f"Thread nao encontrada: {args.session_id}", file=sys.stderr)
        return 1
    sessao = dict(row)
    cwd = sessao.get("cwd") or str(Path.home())
    cmd = ["codex", "resume", args.session_id]
    if args.print_only:
        print(f"cd {cwd}")
        print(" ".join(cmd))
        return 0
    subprocess.run(cmd, cwd=cwd, check=False)
    return 0


def cmd_hook_snippet(args: argparse.Namespace) -> int:
    script = Path(__file__).resolve().as_posix()
    print(
        f"""
[[hooks.StartSession.hooks]]
type = "command"
command = "python {script} hook session-start"
timeout = 5
statusMessage = "Indexando sessao Codex"
""".strip()
    )
    return 0


def criar_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Rastreia sessoes Codex e workers em SQLite.")
    parser.add_argument("--db", default=str(DB_PADRAO), help=f"Banco SQLite. Padrao: {DB_PADRAO}")
    sub = parser.add_subparsers(dest="cmd", required=True)

    hook = sub.add_parser("hook", help="Comandos para hooks do Codex.")
    hook_sub = hook.add_subparsers(dest="hook_cmd", required=True)
    session_start = hook_sub.add_parser("session-start", help="Ler evento SessionStart do stdin e gravar no SQLite.")
    session_start.add_argument("--print-response", action="store_true", help="Imprime resposta JSON minima para debug.")
    session_start.set_defaults(func=cmd_hook)

    sync = sub.add_parser("sync", help="Sincronizar o SQLite a partir dos rollout-*.jsonl.")
    sync.add_argument("--limit", type=int, default=None, help="Limitar quantidade de rollouts lidos.")
    sync.set_defaults(func=cmd_sync)

    listar = sub.add_parser("list", help="Listar sessoes principais.")
    listar.add_argument("--sync", action="store_true", help="Sincronizar rollouts antes de listar.")
    listar.add_argument("--limit", type=int, default=None, help="Limite de rollouts quando usar --sync.")
    listar.add_argument("--live", action="store_true", help="Mostrar apenas sessoes com processo codex vivo.")
    listar.add_argument("--all", action="store_true", help="Incluir subagents na lista principal.")
    listar.add_argument("--include-current", action="store_true", help="Nao excluir a sessao Codex atual.")
    listar.set_defaults(func=cmd_list)

    workers = sub.add_parser("workers", help="Listar workers/subagents de uma thread principal.")
    workers.add_argument("session_id")
    workers.add_argument("--sync", action="store_true", help="Sincronizar rollouts antes de listar.")
    workers.add_argument("--limit", type=int, default=None, help="Limite de rollouts quando usar --sync.")
    workers.set_defaults(func=cmd_workers)

    show = sub.add_parser("show", help="Mostrar registro completo de uma thread.")
    show.add_argument("session_id")
    show.add_argument("--sync", action="store_true", help="Sincronizar rollouts antes de mostrar.")
    show.add_argument("--limit", type=int, default=None)
    show.set_defaults(func=cmd_show)

    resume = sub.add_parser("resume", help="Retomar uma thread no cwd registrado.")
    resume.add_argument("session_id")
    resume.add_argument("--print-only", action="store_true", help="Mostrar o comando sem executar.")
    resume.set_defaults(func=cmd_resume)

    snippet = sub.add_parser("hook-snippet", help="Imprimir bloco TOML para ativar o hook.")
    snippet.set_defaults(func=cmd_hook_snippet)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = criar_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
