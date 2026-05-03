# Global Subagent Overlays

Default rule:

- any non-specialized subagent must run with `zero MCP`
- use `../overlays/no-mcp.toml` unless a specialized workflow explicitly requires a narrower MCP-enabled overlay

## Available overlays

- `../overlays/no-mcp.toml`
  - global default for generic subagents
  - disables every MCP currently configured in the user environment

## Typical use

- coding workers
- shell and filesystem tasks
- Docker and container diagnostics
- GPU benchmarks
- local performance probes
- log inspection
- generic parallel workers

If new MCPs are added to the base config later, extend `../overlays/no-mcp.toml` so the default remains exhaustive.
