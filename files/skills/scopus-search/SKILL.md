---
name: scopus-search
description: Use local Scopus semantic search/indexing over PostgreSQL + pgvector collections. Trigger when the user says `$scopus-search`, asks to list Scopus collections, index one or more Scopus CSV/RIS exports, index a directory of Scopus exports, inspect collection stats, or search the local Scopus vector index.
---

# Scopus Search

Use this skill for local bibliographic work over Scopus exports indexed in the PostgreSQL container with pgvector.

The deterministic command wrapper lives at:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py <subcommand> ...
```

It writes to and searches the PostgreSQL/pgvector schema used by the local RAG stack:

```text
container: rag_pgvector
dsn: postgresql://rag:rag@localhost:5432/ragdb
tables: rag_documents, rag_chunks
```

## Natural Language Interface

The user may call this skill naturally, for example:

```text
$scopus-search lista as collections no postgres
$scopus-search busca "hierarchical forecasting" na collection scopus
$scopus-search indexa os documentos no diretório D:\fontes\scopus na collection revisao
$scopus-search indexa D:\fontes\a.csv e D:\fontes\b.ris na nova collection revisao
```

Map those requests to the CLI subcommands below.

## Query Translation Rule

For `search`, inspect the user's query language first.

- If the user already wrote the query in English, use it as-is.
- If the user wrote the query in another language, translate the search intent to English before calling the CLI.
- Preserve the user's meaning; do not over-expand or silently add unrelated constraints.
- When answering, always show the English query text that was actually used for the search.

Recommended reply shape for search:

- original user query
- English query used
- collection searched
- top results

When invoking the CLI for translated searches, pass both:

- `--query "<english query used for retrieval>"`
- `--original-query "<user query in the original language>"`

If the original query was already in English, you may pass the same text in both fields or omit `--original-query`.

## Commands

List available PostgreSQL/pgvector collections:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections
```

Show collection stats:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py stats --collection scopus
```

Search a collection:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py search --collection scopus --query "hierarchical forecasting" --top-k 5
```

Search a collection when the user asked in another language:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py search --collection scopus --query "hierarchical forecasting" --original-query "previsão hierárquica" --top-k 5
```

Index one CSV export:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py index-csv --collection scopus --csv-path D:\fontes\scopus.csv
```

Index one RIS export:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py index-ris --collection scopus --ris-path D:\fontes\scopus.ris
```

Index one or more CSV/RIS exports:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py index-files --collection scopus D:\fontes\a.csv D:\fontes\b.ris
```

Index every `.csv` and `.ris` file under a directory:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py index-dir --collection scopus --dir D:\fontes\scopus
```

## Required Collection Rule

There is no default collection.

For `search`, `stats`, `index-csv`, `index-ris`, `index-files`, and `index-dir`, the user must provide `collection` explicitly.

For search and stats, the collection must already exist in PostgreSQL.

For indexing one or more files, the user must choose either:

- one existing PostgreSQL collection listed by `list-collections`; or
- one explicit new collection name to be created by the indexing command.

If the user did not provide a collection:

1. Run `list-collections`.
2. Tell the user which PostgreSQL collections exist.
3. Ask which collection to use.
4. For indexing only, mention that they may provide the name of a new collection.

Do not guess a collection and do not use an environment variable as a hidden fallback.

## Indexing Notes

- `index-files` accepts one or more explicit `.csv`/`.ris` paths.
- `index-dir` recursively finds `.csv` and `.ris` files, ordered by path.
- CSV files are sent to `index_csv`; RIS files are sent to `index_ris`.
- Unsupported file types are ignored.
- Scopus CSV exports should contain `Title`, `Abstract`, `Author Keywords`, `Authors`, `Year`, `Source title`, and `DOI`.
- The wrapper handles long abstracts with chunking; do not manually split Scopus CSVs.

## Runtime Options

The wrapper keeps the current defaults:

- `--postgres-dsn postgresql://rag:rag@localhost:5432/ragdb`
- `--ollama-url http://localhost:11434`
- `--model embeddinggemma`

Only override these when the user explicitly asks or the local runtime requires it.
