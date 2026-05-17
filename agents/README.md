# Global Subagent Overlays

Default rule:

- any non-specialized subagent must run with `zero MCP`
- use `../overlays/no-mcp.toml` unless a specialized workflow explicitly requires a narrower MCP-enabled overlay

## Available overlays

- `../overlays/no-mcp.toml`
  - global default for generic subagents
  - disables every MCP currently configured in the user environment

## Specialized implementation agents

- `dotnet-docx-maintainer.toml`
  - .NET/Open XML implementation and review for the `docx-utils` skill
  - strict xUnit red-green-refactor workflow around CLI behavior and plan contracts
  - use for scoped changes in `C:\Users\imale\.codex\skills\docx-utils\src`
  - loads the local `docx-utils` skill instructions and `agents/mantenedor-dotnet-docx.md`
- `backend-django-drf-tdd.toml`
  - Django + Django REST Framework implementation
  - strict pytest/pytest-django red-green-refactor workflow
  - use for backend slices with serializers, viewsets/views, permissions, services/selectors, models, migrations, and API tests
  - receives only backend skills: `django-tdd`, `drf`, and `senior-django-developer`
- `frontend-react-vite-tdd.toml`
  - React + Vite implementation
  - strict Vitest + React Testing Library red-green-refactor workflow
  - use for component, hook, route, state, validation, loading/error/empty-state, and build/test config slices
  - receives only frontend skills: `react-vite`, `frontend-testing`, and `testing-library`
- `node-javascript-tdd.toml`
  - Node.js and modern JavaScript/TypeScript implementation
  - strict red-green-refactor workflow using the repo's existing test runner, Node test runner, or Vitest when appropriate
  - use for scoped Node libraries, JS utilities, TypeScript scripts, CLI helpers, async/stream code, and JS test infrastructure
  - receives only JavaScript/Node skills: `node`, `modern-javascript-patterns`, and `javascript-testing-patterns`
- `golang-tdd-maintainer.toml`
  - Go implementation and review for backend, CLI, daemon, IPC, TUI, Wails backend, concurrency, and cross-platform build-tag work
  - strict Go red-green-refactor workflow with table-driven tests, targeted regression coverage, gofmt, and Go validation commands
  - use for scoped Go patches such as AgentHub daemon/socket/client changes
  - receives only Go skills: `golang-pro` and `golang-testing`
- `pptx-js-fallback-tdd.toml`
  - JavaScript/PptxGenJS fallback implementation for `pptx-utils`
  - use only for creation-from-scratch fallback, generated visual assets, icons, or JS fixtures explicitly assigned by the main thread
  - does not replace the planned .NET/Open XML `pptx-utils` path for reliable PPTX/template editing
  - receives JavaScript/Node skills and inspects the current `pptx` skill fallback guidance when needed
- `tdd-implementation-coordinator.toml`
  - splits implementation plans into parallel backend/frontend TDD work packages
  - marks dependencies, file ownership, validation commands, and integration order
  - receives only orchestration/TDD skills from `superpowers`
- `tdd-quality-reviewer.toml`
  - read-only reviewer for TDD compliance and Django/DRF/React/Vite quality
  - use after worker completion and before final integration
  - always receives the TDD review skill and loads only the reviewed stack's review skills

## Typical use

- coding workers
- shell and filesystem tasks
- Docker and container diagnostics
- GPU benchmarks
- local performance probes
- log inspection
- generic parallel workers

If new MCPs are added to the base config later, extend `../overlays/no-mcp.toml` so the default remains exhaustive.
