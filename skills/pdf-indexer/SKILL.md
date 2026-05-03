---
name: pdf-indexer
description: Index, deduplicate, store, and semantically search local PDF files with pgvector or Chroma compatibility. Use when the user says `$pdf-indexer`, wants to index one PDF or a directory of PDFs, search PDF content by collection, inspect PDF collection stats, list available PDF collections, or fetch PDF structure/sections. For pgvector, this skill stores PDF binaries by default and keeps PDF collections separate from non-PDF collections by redirecting writes to a related `_pdf` collection when needed.
---

# Pdf Indexer

## Overview

Use this skill for local PDF ingestion and retrieval over pgvector or Chroma. It wraps a deterministic CLI at:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py <subcommand> ...
```

## Commands

Bare `$pdf-indexer`:

- Run the script with no subcommand to list commands, defaults, and examples.

Download a PDF:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py download https://host/paper.pdf --dir D:\papers
```

Preview chunking:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py chunk D:\papers\paper.pdf --method header
```

Index one PDF:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py index-pdf D:\papers\paper.pdf --collection pdf --backend pgvector
```

Index a directory of PDFs:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py index-dir --dir D:\papers --recursive --collection pdf
```

List collections:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py list-collections --backend pgvector
```

Show stats:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py stats --collection pdf
```

Search:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py search --collection pdf --query "markov chain" --top-k 5
```

Show structure or retrieve a section:

```powershell
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py structure my-paper.pdf --collection pdf
python C:\Users\imale\.codex\skills\pdf-indexer\scripts\pdf_indexer.py section my-paper.pdf --collection pdf --pages 3-5
```

## Collection Rules

- Default collection name is `pdf`.
- For read operations, if the target collection does not exist, tell the user and list available collections.
- For write operations, if the user omitted the collection and the default `pdf` is about to be used, ask before executing.
- If the user asks to write into an existing non-PDF collection such as `embedding_rag`, the skill should write into `embedding_rag_pdf` instead and record the relationship.
- Keep PDF collections separate from Scopus collections. Do not index PDF chunks directly into the original Scopus collection.

## Backend Rules

- Default backend is `pgvector`.
- For `pgvector`, `store-pdf` defaults to true and the binary PDF is stored for audit use only.
- For `chroma`, do not store PDF binaries; store only the original path and metadata.
- Deduplicate by SHA-256 before extraction, chunking, or embedding, regardless of collection.

## Chunking Rules

- `header` groups text by detected section headers and is the default.
- `s2` uses a heavier spatial-semantic segmentation path and is better for PDFs with weak structure.

## Linking Rules

- If a PDF DOI matches a Scopus DOI, relate the PDF document to the Scopus record.
- Prefer the related Scopus collection first; only fall back to all collections if no match exists there.
