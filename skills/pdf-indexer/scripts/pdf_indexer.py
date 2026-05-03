#!/usr/bin/env python3
"""Deterministic CLI for indexing and searching local PDFs."""

from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import queue
import re
import sys
import threading
from collections import Counter
from pathlib import Path
from typing import Any

import fitz
import psycopg
import requests
import tomllib
from pgvector import Vector
from pgvector.psycopg import register_vector
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb

try:
    import chromadb
except ImportError:
    chromadb = None

DEFAULT_COLLECTION = "pdf"
DEFAULT_PG_USER = "rag"
DEFAULT_PG_PASSWORD = "rag"
DEFAULT_PG_HOST = "localhost"
DEFAULT_PG_PORT = 5432
DEFAULT_PG_DATABASE = "ragdb"
DEFAULT_OLLAMA_URL = "http://localhost:11434"
DEFAULT_MODEL = "nomic-embed-text"
DEFAULT_TIMEOUT = 180
DEFAULT_WORKERS = 4
DEFAULT_TOP_K = 5
DEFAULT_CONTEXT_WINDOW = 1
DEFAULT_CHROMA_DIR = Path(__file__).resolve().parent.parent / "data" / "chroma"
MAX_EMBED_TEXT_LEN = 3000
DOI_PATTERN = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.IGNORECASE)

_S2_MODEL = None


def _json_print(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))


def _strip_nul(value: str) -> str:
    return value.replace("\x00", "")


def _normalize_for_embedding(text: str) -> str:
    """Reduce PDF extraction noise without flattening paragraph structure."""
    normalized = _strip_nul(text).replace("\r\n", "\n").replace("\r", "\n")
    normalized = re.sub(r"\.{4,}", " ... ", normalized)
    normalized = re.sub(r"[_-]{4,}", " ", normalized)
    lines = [re.sub(r"[ \u00A0]+", " ", line).strip() for line in normalized.split("\n")]
    normalized = "\n".join(lines)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.strip()


def _sanitize_value(value: Any) -> Any:
    if isinstance(value, str):
        return _strip_nul(value)
    if isinstance(value, dict):
        return {key: _sanitize_value(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_sanitize_value(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_sanitize_value(item) for item in value)
    return value


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", _strip_nul(str(value or ""))).strip()


def _clean_or_none(value: Any) -> str | None:
    text = _clean(value)
    return text or None


def _normalize_collection(collection: str | None) -> str:
    name = _clean(collection) or DEFAULT_COLLECTION
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", name):
        raise ValueError(
            "Collection invalida. Use apenas letras, numeros, ponto, hifen ou underscore."
        )
    return name


def _content_hash(text: str) -> str:
    return hashlib.sha256(_strip_nul(text).encode("utf-8")).hexdigest()


def _file_sha256(path: Path) -> tuple[str, bytes]:
    data = path.read_bytes()
    return hashlib.sha256(data).hexdigest(), data


def _looks_like_pdf_collection(name: str) -> bool:
    return name.casefold() == DEFAULT_COLLECTION or name.casefold().endswith("_pdf")


def _config_candidates(start: Path) -> list[Path]:
    candidates: list[Path] = []
    for parent in [start, *start.parents]:
        for rel_path in (Path(".codex") / "config.toml", Path("config.toml")):
            candidate = parent / rel_path
            if candidate.exists():
                candidates.append(candidate)
    seen: set[Path] = set()
    ordered: list[Path] = []
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved not in seen:
            seen.add(resolved)
            ordered.append(resolved)
    return ordered


def _read_project_config() -> dict[str, Any]:
    cwd = Path.cwd()
    for candidate in _config_candidates(cwd):
        try:
            with candidate.open("rb") as handle:
                return tomllib.load(handle)
        except Exception:
            continue
    return {}


def _config_lookup(config: dict[str, Any], *keys: str) -> Any:
    current: Any = config
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            return None
        current = current[key]
    return current


def _resolve_setting(
    explicit: Any,
    env_names: list[str],
    config_values: list[Any],
    default: Any,
) -> Any:
    if explicit not in (None, ""):
        return explicit
    for env_name in env_names:
        value = os.environ.get(env_name)
        if value not in (None, ""):
            return value
    for value in config_values:
        if value not in (None, ""):
            return value
    return default


def _bool_flag(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = _clean(value).casefold()
    return text in {"1", "true", "yes", "on", "sim"}


def _chroma_metadata_value(value: Any) -> str | int | float | bool:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return _clean(value)


def _sanitize_chroma_metadata(metadata: dict[str, Any]) -> dict[str, str | int | float | bool]:
    return {key: _chroma_metadata_value(value) for key, value in metadata.items()}


def _effective_store_pdf(args: argparse.Namespace) -> bool:
    if args.backend == "pgvector":
        if args.store_pdf is None:
            return True
        return _bool_flag(args.store_pdf)
    return False


def _requires_collection_confirmation(args: argparse.Namespace) -> bool:
    return args.command in {"index-pdf", "index-files", "index-dir"} and not _clean_or_none(
        args.collection
    )


def _ollama_embed(text: str, model: str, ollama_url: str, timeout: int) -> list[float]:
    prepared = _normalize_for_embedding(text)
    if len(prepared) > MAX_EMBED_TEXT_LEN:
        half = MAX_EMBED_TEXT_LEN // 2
        prepared = prepared[:half].strip() + "\n...\n" + prepared[-half:].strip()
    base = ollama_url.rstrip("/")
    response = requests.post(
        f"{base}/api/embed",
        json={"model": model, "input": prepared},
        timeout=timeout,
    )
    if response.status_code == 404:
        response = requests.post(
            f"{base}/api/embeddings",
            json={"model": model, "prompt": prepared},
            timeout=timeout,
        )
    response.raise_for_status()
    data = response.json()
    embedding = None
    if isinstance(data.get("embeddings"), list) and data["embeddings"]:
        embedding = data["embeddings"][0]
    elif isinstance(data.get("embedding"), list):
        embedding = data["embedding"]
    if not isinstance(embedding, list) or not embedding:
        raise RuntimeError(f"Resposta Ollama sem embedding: {data}")
    return [float(value) for value in embedding]


def _embed_texts_parallel(
    texts: list[str],
    model: str,
    ollama_url: str,
    timeout: int,
    workers: int,
) -> list[list[float]]:
    if not texts:
        return []
    task_queue: queue.Queue[tuple[int, str] | None] = queue.Queue()
    results: list[list[float] | None] = [None] * len(texts)
    errors: list[Exception] = []
    error_lock = threading.Lock()

    def worker() -> None:
        while True:
            item = task_queue.get()
            if item is None:
                task_queue.task_done()
                break
            index, text = item
            try:
                results[index] = _ollama_embed(text, model, ollama_url, timeout)
            except Exception as exc:
                with error_lock:
                    errors.append(exc)
            finally:
                task_queue.task_done()

    threads = [
        threading.Thread(target=worker, daemon=True, name=f"pdf-indexer-{idx}")
        for idx in range(max(1, workers))
    ]
    for thread in threads:
        thread.start()
    for index, text in enumerate(texts):
        task_queue.put((index, text))
    for _ in threads:
        task_queue.put(None)
    task_queue.join()
    if errors:
        raise errors[0]
    if any(item is None for item in results):
        raise RuntimeError("Falha ao gerar embeddings para todos os chunks.")
    return [item for item in results if item is not None]


def _calculate_common_font_size(doc: fitz.Document) -> float:
    sizes: list[float] = []
    for page_index in range(len(doc)):
        for block in doc[page_index].get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    size = span.get("size")
                    if size:
                        sizes.append(float(size))
    if not sizes:
        return 12.0
    return Counter(sizes).most_common(1)[0][0]


def _is_likely_header(
    text: str,
    font_size: float,
    font_name: str,
    flags: int,
    common_font_size: float,
    x_position: float,
) -> tuple[bool, int]:
    stripped = text.strip()
    if len(stripped) < 3:
        return False, 0
    patterns = [
        r"^(\d+\.?\s+)?(Abstract|Introduction|Related Work|Background|Methodology|Methods|Results|Discussion|Conclusion|References|Bibliography)",
        r"^(\d+\.?\s+)?[A-Z][a-z]+\s+",
        r"^\d+\.\s+[A-Z]",
        r"^[A-Z][A-Z\s]+$",
    ]
    is_section_like = any(re.match(pattern, stripped, re.IGNORECASE) for pattern in patterns)
    size_ratio = font_size / common_font_size if common_font_size > 0 else 1.0
    is_bold = (flags & 16) == 16 or "BX" in font_name.upper() or "Bold" in font_name
    is_left_aligned = x_position < 100
    level = 0
    if is_section_like:
        if size_ratio > 1.5 or (size_ratio > 1.2 and is_bold):
            level = 1
        elif size_ratio > 1.1 or is_bold:
            level = 2
        elif size_ratio >= 1.0 or (is_bold and is_left_aligned):
            level = 3
    is_header = (
        (size_ratio > 1.1 and is_bold)
        or (is_section_like and (size_ratio > 1.05 or is_bold))
        or (is_bold and is_left_aligned and len(stripped) < 100)
    )
    return is_header, level


def _extract_segments(doc: fitz.Document) -> list[dict[str, Any]]:
    segments: list[dict[str, Any]] = []
    common_font_size = _calculate_common_font_size(doc)
    for page_index in range(len(doc)):
        page = doc[page_index]
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            block_text_parts: list[str] = []
            block_is_header = False
            block_header_level = 0
            block_font_size = common_font_size
            block_font_name = ""
            block_bbox = block.get("bbox", [0, 0, 0, 0])
            for line in block.get("lines", []):
                line_text_parts: list[str] = []
                for span in line.get("spans", []):
                    span_text = span.get("text", "").strip()
                    if not span_text:
                        continue
                    font_size = float(span.get("size", common_font_size))
                    font_name = span.get("font", "")
                    flags = int(span.get("flags", 0))
                    x_position = span.get("bbox", [0, 0, 0, 0])[0]
                    is_header, header_level = _is_likely_header(
                        span_text,
                        font_size,
                        font_name,
                        flags,
                        common_font_size,
                        x_position,
                    )
                    line_text_parts.append(span_text)
                    if is_header:
                        block_is_header = True
                        block_header_level = max(block_header_level, header_level)
                        block_font_size = font_size
                        block_font_name = font_name
                if line_text_parts:
                    block_text_parts.append(" ".join(line_text_parts))
            if block_text_parts:
                block_text = "\n".join(block_text_parts).strip()
                if block_text:
                    segments.append(
                        {
                            "text": block_text,
                            "page": page_index,
                            "bbox": tuple(block_bbox),
                            "is_header": block_is_header,
                            "header_level": block_header_level,
                            "font_size": block_font_size,
                            "font_name": block_font_name,
                        }
                    )
    return segments


def _build_header_path(segments: list[dict[str, Any]], current_index: int) -> str:
    stack: list[tuple[int, str]] = []
    for index in range(current_index + 1):
        segment = segments[index]
        if not segment["is_header"]:
            continue
        level = int(segment["header_level"])
        text = re.sub(r"^\d+\.\s*", "", segment["text"].strip())
        while stack and stack[-1][0] >= level:
            stack.pop()
        stack.append((level, text))
    return " > ".join(text for _, text in stack)


def _chunk_by_headers(segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
    chunks: list[dict[str, Any]] = []
    current_segments: list[dict[str, Any]] = []
    current_header_path = ""
    current_header_level = 0
    current_page_start = 0
    current_page_end = 0
    chunk_index = 0
    for index, segment in enumerate(segments):
        is_header = bool(segment["is_header"])
        header_level = int(segment["header_level"])
        page = int(segment["page"])
        if is_header and header_level > 0:
            if current_segments:
                chunks.append(
                    {
                        "chunk_index": chunk_index,
                        "text": "\n\n".join(item["text"] for item in current_segments).strip(),
                        "header_path": current_header_path,
                        "page_start": current_page_start,
                        "page_end": current_page_end,
                        "header_level": current_header_level,
                    }
                )
                chunk_index += 1
            current_segments = [segment]
            current_header_path = _build_header_path(segments, index)
            current_header_level = header_level
            current_page_start = page
            current_page_end = page
        else:
            if not current_segments:
                current_page_start = page
            current_segments.append(segment)
            current_page_end = page
    if current_segments:
        chunks.append(
            {
                "chunk_index": chunk_index,
                "text": "\n\n".join(item["text"] for item in current_segments).strip(),
                "header_path": current_header_path,
                "page_start": current_page_start,
                "page_end": current_page_end,
                "header_level": current_header_level,
            }
        )
    return chunks


def _get_s2_model():
    global _S2_MODEL
    if _S2_MODEL is None:
        from sentence_transformers import SentenceTransformer

        _S2_MODEL = SentenceTransformer("all-MiniLM-L6-v2")
    return _S2_MODEL


def _spatial_distance(bbox1: tuple[float, ...], bbox2: tuple[float, ...]) -> float:
    x1 = (bbox1[0] + bbox1[2]) / 2
    y1 = (bbox1[1] + bbox1[3]) / 2
    x2 = (bbox2[0] + bbox2[2]) / 2
    y2 = (bbox2[1] + bbox2[3]) / 2
    return ((x1 - x2) ** 2 + (y1 - y2) ** 2) ** 0.5


def _split_long_chunks(chunks: list[dict[str, Any]], max_token_length: int) -> list[dict[str, Any]]:
    max_chars = max_token_length * 4
    final_chunks: list[dict[str, Any]] = []
    next_index = 0
    for chunk in chunks:
        text = chunk["text"]
        if len(text) <= max_chars:
            chunk["chunk_index"] = next_index
            final_chunks.append(chunk)
            next_index += 1
            continue
        paragraphs = text.split("\n\n")
        current = ""
        for paragraph in paragraphs:
            piece = paragraph.strip()
            if not piece:
                continue
            if len(current) + len(piece) + 2 <= max_chars:
                current = f"{current}\n\n{piece}".strip()
                continue
            if current:
                final_chunks.append(
                    {
                        "chunk_index": next_index,
                        "text": current,
                        "header_path": chunk.get("header_path", ""),
                        "page_start": chunk["page_start"],
                        "page_end": chunk["page_end"],
                        "header_level": chunk.get("header_level", 0),
                    }
                )
                next_index += 1
            current = piece
        if current:
            final_chunks.append(
                {
                    "chunk_index": next_index,
                    "text": current,
                    "header_path": chunk.get("header_path", ""),
                    "page_start": chunk["page_start"],
                    "page_end": chunk["page_end"],
                    "header_level": chunk.get("header_level", 0),
                }
            )
            next_index += 1
    return final_chunks


def _chunk_by_s2(segments: list[dict[str, Any]], max_token_length: int = 512) -> list[dict[str, Any]]:
    if not segments:
        return []
    if len(segments) == 1:
        segment = segments[0]
        return [
            {
                "chunk_index": 0,
                "text": segment["text"],
                "header_path": "",
                "page_start": int(segment["page"]),
                "page_end": int(segment["page"]),
                "header_level": 0,
            }
        ]
    import numpy as np
    from sklearn.cluster import SpectralClustering

    model = _get_s2_model()
    texts = [segment["text"] for segment in segments]
    embeddings = model.encode(texts, normalize_embeddings=True)
    affinity = np.zeros((len(segments), len(segments)))
    for left in range(len(segments)):
        affinity[left, left] = 1.0
        for right in range(left + 1, len(segments)):
            spatial_weight = 1.0 / (1.0 + _spatial_distance(segments[left]["bbox"], segments[right]["bbox"]))
            semantic_weight = max(0.0, float(np.dot(embeddings[left], embeddings[right])))
            combined = (spatial_weight + semantic_weight) / 2.0
            affinity[left, right] = combined
            affinity[right, left] = combined
    total_chars = sum(len(text) for text in texts)
    estimated_tokens = max(1, total_chars // 4)
    cluster_count = max(1, min(len(segments), (estimated_tokens + max_token_length - 1) // max_token_length))
    clustering = SpectralClustering(
        n_clusters=cluster_count,
        affinity="precomputed",
        assign_labels="kmeans",
        random_state=42,
    )
    labels = clustering.fit_predict(affinity)
    grouped: dict[int, list[dict[str, Any]]] = {}
    for index, label in enumerate(labels):
        grouped.setdefault(int(label), []).append(segments[index])
    chunks: list[dict[str, Any]] = []
    for label in sorted(grouped):
        group = grouped[label]
        pages = [int(segment["page"]) for segment in group]
        chunks.append(
            {
                "chunk_index": label,
                "text": "\n\n".join(segment["text"] for segment in group).strip(),
                "header_path": "",
                "page_start": min(pages),
                "page_end": max(pages),
                "header_level": 0,
            }
        )
    return _split_long_chunks(chunks, max_token_length)


def _extract_doi(doc: fitz.Document, segments: list[dict[str, Any]]) -> str | None:
    metadata = doc.metadata or {}
    candidate_fields = [
        metadata.get("title"),
        metadata.get("subject"),
        metadata.get("keywords"),
        metadata.get("author"),
    ]
    for candidate in candidate_fields:
        if candidate:
            match = DOI_PATTERN.search(candidate)
            if match:
                return match.group(0).rstrip(".")
    sample = "\n".join(segment["text"] for segment in segments[:25])
    match = DOI_PATTERN.search(sample)
    if match:
        return match.group(0).rstrip(".")
    return None


def _extract_title(path: Path, doc: fitz.Document, chunks: list[dict[str, Any]]) -> str:
    metadata_title = _clean_or_none((doc.metadata or {}).get("title"))
    if metadata_title:
        return metadata_title
    for chunk in chunks:
        header_path = _clean_or_none(chunk.get("header_path"))
        if header_path:
            return header_path.split(" > ")[0]
    return path.stem


def _prepare_pdf_chunks(path: Path, method: str) -> dict[str, Any]:
    doc = fitz.open(path)
    try:
        segments = _extract_segments(doc)
        if method == "header":
            chunks = _chunk_by_headers(segments)
        elif method == "s2":
            chunks = _chunk_by_s2(segments)
        else:
            raise ValueError("Metodo invalido. Use header ou s2.")
        for index, chunk in enumerate(chunks):
            chunk["prev_chunk_index"] = index - 1 if index > 0 else None
            chunk["next_chunk_index"] = index + 1 if index + 1 < len(chunks) else None
            chunk["chunk_index"] = index
            chunk["chunk_count"] = len(chunks)
        title = _extract_title(path, doc, chunks)
        doi = _extract_doi(doc, segments)
        return {
            "title": title,
            "doi": doi,
            "num_pages": len(doc),
            "segments": segments,
            "chunks": chunks,
        }
    finally:
        doc.close()


def _connect_pg(args: argparse.Namespace, config: dict[str, Any]) -> tuple[psycopg.Connection, dict[str, Any]]:
    pg_user = _resolve_setting(
        args.pg_user,
        ["PDF_INDEXER_PG_USER", "PGUSER"],
        [
            _config_lookup(config, "pdf_indexer", "pgvector", "user"),
            _config_lookup(config, "tool", "pdf_indexer", "pgvector", "user"),
        ],
        DEFAULT_PG_USER,
    )
    pg_password = _resolve_setting(
        args.pg_password,
        ["PDF_INDEXER_PG_PASSWORD", "PGPASSWORD"],
        [
            _config_lookup(config, "pdf_indexer", "pgvector", "password"),
            _config_lookup(config, "tool", "pdf_indexer", "pgvector", "password"),
        ],
        DEFAULT_PG_PASSWORD,
    )
    pg_host = _resolve_setting(
        args.pg_host,
        ["PDF_INDEXER_PG_HOST", "PGHOST"],
        [
            _config_lookup(config, "pdf_indexer", "pgvector", "host"),
            _config_lookup(config, "tool", "pdf_indexer", "pgvector", "host"),
        ],
        DEFAULT_PG_HOST,
    )
    pg_port = int(
        _resolve_setting(
            args.pg_port,
            ["PDF_INDEXER_PG_PORT", "PGPORT"],
            [
                _config_lookup(config, "pdf_indexer", "pgvector", "port"),
                _config_lookup(config, "tool", "pdf_indexer", "pgvector", "port"),
            ],
            DEFAULT_PG_PORT,
        )
    )
    pg_database = _resolve_setting(
        args.pg_database,
        ["PDF_INDEXER_PG_DATABASE", "PGDATABASE"],
        [
            _config_lookup(config, "pdf_indexer", "pgvector", "database"),
            _config_lookup(config, "tool", "pdf_indexer", "pgvector", "database"),
        ],
        DEFAULT_PG_DATABASE,
    )
    dsn = (
        f"postgresql://{pg_user}:{pg_password}@{pg_host}:{pg_port}/{pg_database}"
    )
    conn = psycopg.connect(dsn, row_factory=dict_row, connect_timeout=10)
    register_vector(conn)
    return conn, {
        "dsn": dsn,
        "user": pg_user,
        "host": pg_host,
        "port": pg_port,
        "database": pg_database,
    }


def _ensure_pg_schema(conn: psycopg.Connection) -> None:
    with conn.cursor() as cur:
        cur.execute("CREATE EXTENSION IF NOT EXISTS vector")
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS rag_documents (
                id BIGSERIAL PRIMARY KEY,
                source TEXT NOT NULL,
                external_id TEXT,
                doi TEXT,
                eid TEXT,
                title TEXT NOT NULL,
                authors TEXT,
                year INTEGER,
                source_title TEXT,
                document_type TEXT,
                language TEXT DEFAULT 'en',
                abstract TEXT,
                author_keywords TEXT,
                index_keywords TEXT,
                search_batches TEXT,
                method_family TEXT,
                method TEXT,
                domain TEXT,
                subdomain TEXT,
                application TEXT,
                source_file TEXT,
                raw_metadata JSONB DEFAULT '{}'::jsonb,
                document_key TEXT NOT NULL UNIQUE,
                collection TEXT,
                embedding_model TEXT,
                embedding_dim INTEGER,
                content_hash TEXT,
                created_at TIMESTAMPTZ DEFAULT now(),
                updated_at TIMESTAMPTZ DEFAULT now()
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS rag_chunks (
                id BIGSERIAL PRIMARY KEY,
                document_id BIGINT NOT NULL REFERENCES rag_documents(id) ON DELETE CASCADE,
                chunk_id TEXT NOT NULL UNIQUE,
                chunk_index INTEGER NOT NULL DEFAULT 0,
                chunk_count INTEGER NOT NULL DEFAULT 1,
                text_kind TEXT,
                content TEXT NOT NULL,
                embedding VECTOR NOT NULL,
                raw_metadata JSONB DEFAULT '{}'::jsonb,
                embedding_model TEXT,
                embedding_dim INTEGER,
                content_hash TEXT,
                created_at TIMESTAMPTZ DEFAULT now(),
                updated_at TIMESTAMPTZ DEFAULT now()
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS rag_pdf_files (
                id BIGSERIAL PRIMARY KEY,
                sha256 TEXT NOT NULL UNIQUE,
                filename TEXT NOT NULL,
                original_path TEXT,
                mime_type TEXT,
                size_bytes BIGINT NOT NULL,
                file_data BYTEA,
                created_at TIMESTAMPTZ DEFAULT now(),
                updated_at TIMESTAMPTZ DEFAULT now()
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS rag_collections (
                name TEXT PRIMARY KEY,
                kind TEXT NOT NULL,
                backend TEXT NOT NULL DEFAULT 'pgvector',
                related_collection TEXT,
                metadata JSONB DEFAULT '{}'::jsonb,
                created_at TIMESTAMPTZ DEFAULT now(),
                updated_at TIMESTAMPTZ DEFAULT now()
            )
            """
        )
        cur.execute(
            """
            ALTER TABLE rag_documents
            ADD COLUMN IF NOT EXISTS pdf_file_id BIGINT REFERENCES rag_pdf_files(id)
            """
        )
        cur.execute(
            """
            ALTER TABLE rag_documents
            ADD COLUMN IF NOT EXISTS related_scopus_document_id BIGINT REFERENCES rag_documents(id)
            """
        )
    conn.commit()


def _list_pg_collections(conn: psycopg.Connection) -> list[dict[str, Any]]:
    _ensure_pg_schema(conn)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            WITH doc_counts AS (
                SELECT
                    collection,
                    count(DISTINCT id) AS documents,
                    count(*) FILTER (WHERE source = 'pdf') AS pdf_documents,
                    count(*) FILTER (WHERE source <> 'pdf') AS non_pdf_documents
                FROM rag_documents
                WHERE collection IS NOT NULL AND collection <> ''
                GROUP BY collection
            ),
            chunk_counts AS (
                SELECT d.collection, count(c.id) AS chunks
                FROM rag_documents d
                LEFT JOIN rag_chunks c ON c.document_id = d.id
                WHERE d.collection IS NOT NULL AND d.collection <> ''
                GROUP BY d.collection
            ),
            model_dims AS (
                SELECT
                    d.collection,
                    COALESCE(c.embedding_model, d.embedding_model) AS embedding_model,
                    COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS embedding_dim
                FROM rag_documents d
                LEFT JOIN rag_chunks c ON c.document_id = d.id
                WHERE d.collection IS NOT NULL AND d.collection <> ''
                GROUP BY d.collection, 2, 3
            ),
            names AS (
                SELECT name AS collection FROM rag_collections
                UNION
                SELECT collection FROM rag_documents WHERE collection IS NOT NULL AND collection <> ''
            )
            SELECT
                n.collection,
                COALESCE(rc.kind,
                    CASE
                        WHEN COALESCE(dc.pdf_documents, 0) > 0 AND COALESCE(dc.non_pdf_documents, 0) = 0 THEN 'pdf'
                        WHEN COALESCE(dc.non_pdf_documents, 0) > 0 THEN 'external'
                        WHEN n.collection LIKE '%_pdf' THEN 'pdf'
                        ELSE 'unknown'
                    END
                ) AS kind,
                COALESCE(rc.related_collection, '') AS related_collection,
                COALESCE(dc.documents, 0) AS documents,
                COALESCE(cc.chunks, 0) AS chunks,
                max(md.embedding_model) AS embedding_model,
                max(md.embedding_dim) AS embedding_dim
            FROM names n
            LEFT JOIN rag_collections rc ON rc.name = n.collection
            LEFT JOIN doc_counts dc ON dc.collection = n.collection
            LEFT JOIN chunk_counts cc ON cc.collection = n.collection
            LEFT JOIN model_dims md ON md.collection = n.collection
            GROUP BY n.collection, 2, 3, 4, 5
            ORDER BY n.collection
            """
        )
        return [dict(row) for row in cur.fetchall()]


def _pg_collection_exists(conn: psycopg.Connection, collection: str) -> bool:
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT EXISTS (
                SELECT 1 FROM rag_documents WHERE collection = %s
                UNION ALL
                SELECT 1 FROM rag_collections WHERE name = %s
            )
            """,
            (collection, collection),
        )
        row = cur.fetchone()
        return bool(row["exists"]) if row else False


def _pg_collection_kind(conn: psycopg.Connection, collection: str) -> str | None:
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT kind FROM rag_collections WHERE name = %s
            """,
            (collection,),
        )
        row = cur.fetchone()
        if row:
            return row["kind"]
        cur.execute(
            """
            SELECT
                CASE
                    WHEN count(*) FILTER (WHERE source = 'pdf') > 0
                         AND count(*) FILTER (WHERE source <> 'pdf') = 0 THEN 'pdf'
                    WHEN count(*) FILTER (WHERE source <> 'pdf') > 0 THEN 'external'
                    ELSE NULL
                END AS kind
            FROM rag_documents
            WHERE collection = %s
            """,
            (collection,),
        )
        row = cur.fetchone()
        return row["kind"] if row else None


def _ensure_collection_registration(
    conn: psycopg.Connection,
    name: str,
    kind: str,
    related_collection: str | None,
    metadata: dict[str, Any],
) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO rag_collections (name, kind, backend, related_collection, metadata, updated_at)
            VALUES (%s, %s, 'pgvector', %s, %s, now())
            ON CONFLICT (name) DO UPDATE SET
                kind = EXCLUDED.kind,
                backend = EXCLUDED.backend,
                related_collection = EXCLUDED.related_collection,
                metadata = COALESCE(rag_collections.metadata, '{}'::jsonb) || EXCLUDED.metadata,
                updated_at = now()
            """,
            (name, kind, related_collection, Jsonb(metadata)),
        )


def _resolve_pg_target_collection(
    conn: psycopg.Connection,
    requested_collection: str,
) -> dict[str, Any]:
    requested = _normalize_collection(requested_collection)
    kind = _pg_collection_kind(conn, requested)
    if kind and kind != "pdf" and not _looks_like_pdf_collection(requested):
        effective = f"{requested}_pdf"
        related_collection = requested
    else:
        effective = requested
        related_collection = requested[:-4] if requested.endswith("_pdf") else None
        if related_collection and not _pg_collection_exists(conn, related_collection):
            related_collection = None
    _ensure_collection_registration(
        conn,
        effective,
        "pdf",
        related_collection,
        {
            "requested_collection": requested,
            "effective_collection": effective,
            "related_collection": related_collection,
        },
    )
    conn.commit()
    return {
        "requested_collection": requested,
        "effective_collection": effective,
        "related_collection": related_collection,
        "requested_kind": kind,
    }


def _upsert_pdf_file(
    conn: psycopg.Connection,
    path: Path,
    sha256: str,
    data: bytes,
    store_pdf: bool,
) -> dict[str, Any]:
    mime_type = mimetypes.guess_type(path.name)[0] or "application/pdf"
    safe_name = _strip_nul(path.name)
    safe_path = _strip_nul(str(path))
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            INSERT INTO rag_pdf_files (sha256, filename, original_path, mime_type, size_bytes, file_data, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, now())
            ON CONFLICT (sha256) DO UPDATE SET
                filename = EXCLUDED.filename,
                original_path = EXCLUDED.original_path,
                mime_type = EXCLUDED.mime_type,
                size_bytes = EXCLUDED.size_bytes,
                file_data = CASE
                    WHEN EXCLUDED.file_data IS NOT NULL THEN EXCLUDED.file_data
                    ELSE rag_pdf_files.file_data
                END,
                updated_at = now()
            RETURNING id, sha256, filename, original_path, mime_type, size_bytes, file_data IS NOT NULL AS has_binary
            """,
            (
                sha256,
                safe_name,
                safe_path,
                mime_type,
                len(data),
                data if store_pdf else None,
            ),
        )
        return dict(cur.fetchone())


def _find_existing_pdf_doc(
    conn: psycopg.Connection,
    pdf_file_id: int,
    collection: str | None = None,
) -> dict[str, Any] | None:
    query = """
        SELECT id, collection, document_key, raw_metadata, related_scopus_document_id
        FROM rag_documents
        WHERE source = 'pdf' AND pdf_file_id = %s
    """
    params: list[Any] = [pdf_file_id]
    if collection:
        query += " AND collection = %s"
        params.append(collection)
    query += " ORDER BY updated_at DESC NULLS LAST LIMIT 1"
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(query, params)
        row = cur.fetchone()
        return dict(row) if row else None


def _find_any_existing_pdf_doc(conn: psycopg.Connection, pdf_file_id: int) -> dict[str, Any] | None:
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT id, collection, document_key, raw_metadata, related_scopus_document_id
            FROM rag_documents
            WHERE source = 'pdf' AND pdf_file_id = %s
            ORDER BY updated_at DESC NULLS LAST
            LIMIT 1
            """,
            (pdf_file_id,),
        )
        row = cur.fetchone()
        return dict(row) if row else None


def _fetch_pdf_document_and_chunks(conn: psycopg.Connection, document_id: int) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute("SELECT * FROM rag_documents WHERE id = %s", (document_id,))
        document = cur.fetchone()
        if not document:
            raise RuntimeError(f"Documento nao encontrado: {document_id}")
        cur.execute(
            """
            SELECT chunk_id, chunk_index, chunk_count, text_kind, content, embedding, raw_metadata,
                   embedding_model, embedding_dim, content_hash
            FROM rag_chunks
            WHERE document_id = %s
            ORDER BY chunk_index
            """,
            (document_id,),
        )
        chunks = [dict(row) for row in cur.fetchall()]
    return dict(document), chunks


def _upsert_pdf_document(
    conn: psycopg.Connection,
    *,
    collection: str,
    related_collection: str | None,
    pdf_file_id: int,
    pdf_sha256: str,
    source_file: str,
    title: str,
    doi: str | None,
    num_pages: int,
    num_chunks: int,
    method: str,
    model: str,
    embedding_dim: int,
    related_scopus_document_id: int | None,
    extra_metadata: dict[str, Any],
) -> int:
    document_key = f"pdf:{collection}:{pdf_sha256}"
    safe_title = _strip_nul(title)
    safe_source_file = _strip_nul(source_file)
    safe_doi = _strip_nul(doi or "")
    raw_metadata = {
        "source": "pdf",
        "title": safe_title,
        "doi": safe_doi,
        "filename": _strip_nul(Path(source_file).name),
        "source_file": safe_source_file,
        "file_sha256": pdf_sha256,
        "num_pages": num_pages,
        "num_chunks": num_chunks,
        "chunking_method": method,
        "requested_collection": extra_metadata.get("requested_collection"),
        "effective_collection": collection,
        "related_collection": related_collection,
        **extra_metadata,
    }
    raw_metadata = _sanitize_value(raw_metadata)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            INSERT INTO rag_documents (
                source, external_id, doi, title, source_title, document_type, language,
                source_file, raw_metadata, document_key, collection, embedding_model,
                embedding_dim, content_hash, pdf_file_id, related_scopus_document_id, updated_at
            )
            VALUES (
                'pdf', %s, %s, %s, %s, 'PDF', 'und', %s, %s, %s, %s, %s, %s, %s, %s, %s, now()
            )
            ON CONFLICT (document_key) DO UPDATE SET
                doi = EXCLUDED.doi,
                title = EXCLUDED.title,
                source_title = EXCLUDED.source_title,
                source_file = EXCLUDED.source_file,
                raw_metadata = EXCLUDED.raw_metadata,
                collection = EXCLUDED.collection,
                embedding_model = EXCLUDED.embedding_model,
                embedding_dim = EXCLUDED.embedding_dim,
                content_hash = EXCLUDED.content_hash,
                pdf_file_id = EXCLUDED.pdf_file_id,
                related_scopus_document_id = EXCLUDED.related_scopus_document_id,
                updated_at = now()
            RETURNING id
            """,
            (
                pdf_sha256,
                safe_doi or None,
                safe_title,
                safe_title,
                safe_source_file,
                Jsonb(raw_metadata),
                document_key,
                collection,
                model,
                embedding_dim,
                pdf_sha256,
                pdf_file_id,
                related_scopus_document_id,
            ),
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError("Falha ao persistir o documento PDF.")
        return int(row["id"])


def _replace_pdf_chunks(
    conn: psycopg.Connection,
    document_id: int,
    collection: str,
    pdf_sha256: str,
    source_file: str,
    doi: str | None,
    method: str,
    chunks: list[dict[str, Any]],
    embeddings: list[list[float]],
    model: str,
    related_collection: str | None,
) -> int:
    if len(chunks) != len(embeddings):
        raise RuntimeError("Quantidade de chunks e embeddings divergente.")
    with conn.cursor() as cur:
        cur.execute("DELETE FROM rag_chunks WHERE document_id = %s", (document_id,))
        for chunk, embedding in zip(chunks, embeddings):
            chunk_id = f"pdf:{collection}:{pdf_sha256}:chunk:{chunk['chunk_index']}"
            safe_source_file = _strip_nul(source_file)
            safe_doi = _strip_nul(doi or "")
            safe_text = _strip_nul(chunk["text"])
            metadata = {
                "source": "pdf",
                "collection": collection,
                "related_collection": related_collection,
                "filename": _strip_nul(Path(source_file).name),
                "source_file": safe_source_file,
                "doi": safe_doi,
                "text_kind": "pdf_chunk",
                "chunking_method": method,
                "chunk_index": chunk["chunk_index"],
                "chunk_count": chunk["chunk_count"],
                "page_start": chunk["page_start"],
                "page_end": chunk["page_end"],
                "header_path": chunk.get("header_path", ""),
                "header_level": chunk.get("header_level", 0),
                "prev_chunk_index": chunk.get("prev_chunk_index"),
                "next_chunk_index": chunk.get("next_chunk_index"),
            }
            metadata = _sanitize_value(metadata)
            cur.execute(
                """
                INSERT INTO rag_chunks (
                    document_id, chunk_id, chunk_index, chunk_count, text_kind,
                    content, embedding, raw_metadata, embedding_model, embedding_dim,
                    content_hash, updated_at
                )
                VALUES (
                    %s, %s, %s, %s, 'pdf_chunk', %s, %s, %s, %s, %s, %s, now()
                )
                ON CONFLICT (chunk_id) DO UPDATE SET
                    document_id = EXCLUDED.document_id,
                    chunk_index = EXCLUDED.chunk_index,
                    chunk_count = EXCLUDED.chunk_count,
                    text_kind = EXCLUDED.text_kind,
                    content = EXCLUDED.content,
                    embedding = EXCLUDED.embedding,
                    raw_metadata = EXCLUDED.raw_metadata,
                    embedding_model = EXCLUDED.embedding_model,
                    embedding_dim = EXCLUDED.embedding_dim,
                    content_hash = EXCLUDED.content_hash,
                    updated_at = now()
                """,
                (
                    document_id,
                    chunk_id,
                    chunk["chunk_index"],
                    chunk["chunk_count"],
                    safe_text,
                    Vector(embedding),
                    Jsonb(metadata),
                    model,
                    len(embedding),
                    _content_hash(safe_text),
                ),
            )
    return len(chunks)


def _find_related_scopus_document(
    conn: psycopg.Connection,
    doi: str | None,
    related_collection: str | None,
) -> dict[str, Any] | None:
    if not doi:
        return None
    with conn.cursor(row_factory=dict_row) as cur:
        if related_collection:
            cur.execute(
                """
                SELECT id, collection, source, title
                FROM rag_documents
                WHERE source <> 'pdf' AND doi = %s AND collection = %s
                ORDER BY updated_at DESC NULLS LAST
                LIMIT 1
                """,
                (doi, related_collection),
            )
            row = cur.fetchone()
            if row:
                return dict(row)
        cur.execute(
            """
            SELECT id, collection, source, title
            FROM rag_documents
            WHERE source <> 'pdf' AND doi = %s
            ORDER BY updated_at DESC NULLS LAST
            LIMIT 1
            """,
            (doi,),
        )
        row = cur.fetchone()
        return dict(row) if row else None


def _clone_pdf_document(
    conn: psycopg.Connection,
    source_document_id: int,
    effective_collection: str,
    requested_collection: str,
    related_collection: str | None,
) -> dict[str, Any]:
    source_document, source_chunks = _fetch_pdf_document_and_chunks(conn, source_document_id)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute("SELECT sha256 FROM rag_pdf_files WHERE id = %s", (source_document["pdf_file_id"],))
        pdf_row = cur.fetchone()
        if not pdf_row:
            raise RuntimeError("Arquivo PDF relacionado nao encontrado para clonagem.")
        pdf_sha256 = pdf_row["sha256"]
    related_scopus = _find_related_scopus_document(conn, source_document.get("doi"), related_collection)
    related_scopus_document_id = related_scopus["id"] if related_scopus else source_document.get("related_scopus_document_id")
    doc_id = _upsert_pdf_document(
        conn,
        collection=effective_collection,
        related_collection=related_collection,
        pdf_file_id=source_document["pdf_file_id"],
        pdf_sha256=pdf_sha256,
        source_file=source_document["source_file"],
        title=source_document["title"],
        doi=source_document.get("doi"),
        num_pages=int((source_document.get("raw_metadata") or {}).get("num_pages", 0)),
        num_chunks=int((source_document.get("raw_metadata") or {}).get("num_chunks", len(source_chunks))),
        method=_clean((source_document.get("raw_metadata") or {}).get("chunking_method")) or "header",
        model=_clean(source_document.get("embedding_model")) or DEFAULT_MODEL,
        embedding_dim=int(source_document.get("embedding_dim") or source_chunks[0]["embedding_dim"] or 0),
        related_scopus_document_id=related_scopus_document_id,
        extra_metadata={
            "requested_collection": requested_collection,
            "clone_source_collection": source_document["collection"],
            "dedupe_strategy": "clone_existing_chunks",
        },
    )
    with conn.cursor() as cur:
        cur.execute("DELETE FROM rag_chunks WHERE document_id = %s", (doc_id,))
        for chunk in source_chunks:
            chunk_index = int(chunk["chunk_index"])
            chunk_id = f"pdf:{effective_collection}:{pdf_sha256}:chunk:{chunk_index}"
            metadata = dict(chunk["raw_metadata"] or {})
            metadata["collection"] = effective_collection
            metadata["related_collection"] = related_collection
            metadata["source_file"] = _strip_nul(source_document["source_file"])
            metadata = _sanitize_value(metadata)
            safe_content = _strip_nul(chunk["content"])
            cur.execute(
                """
                INSERT INTO rag_chunks (
                    document_id, chunk_id, chunk_index, chunk_count, text_kind,
                    content, embedding, raw_metadata, embedding_model, embedding_dim,
                    content_hash, updated_at
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, now())
                """,
                (
                    doc_id,
                    chunk_id,
                    chunk_index,
                    int(chunk["chunk_count"]),
                    chunk.get("text_kind") or "pdf_chunk",
                    safe_content,
                    chunk["embedding"],
                    Jsonb(metadata),
                    chunk.get("embedding_model") or source_document.get("embedding_model"),
                    int(chunk.get("embedding_dim") or source_document.get("embedding_dim") or 0),
                    chunk.get("content_hash") or _content_hash(safe_content),
                ),
            )
    conn.commit()
    return {
        "action": "reused_existing_pdf",
        "document_id": doc_id,
        "source_document_id": source_document_id,
        "effective_collection": effective_collection,
        "related_collection": related_collection,
        "related_scopus_document_id": related_scopus_document_id,
        "chunks": len(source_chunks),
    }


def _list_missing_collection_response(conn: psycopg.Connection, collection: str) -> dict[str, Any]:
    return {
        "error": f"Collection nao encontrada: {collection}",
        "collections": _list_pg_collections(conn),
    }


def _locate_pdf_document(
    conn: psycopg.Connection,
    collection: str,
    document: str,
) -> dict[str, Any] | None:
    with conn.cursor(row_factory=dict_row) as cur:
        if document.isdigit():
            cur.execute(
                """
                SELECT * FROM rag_documents
                WHERE id = %s AND collection = %s AND source = 'pdf'
                """,
                (int(document), collection),
            )
            row = cur.fetchone()
            if row:
                return dict(row)
        cur.execute(
            """
            SELECT *
            FROM rag_documents
            WHERE collection = %s AND source = 'pdf'
              AND (
                    source_file = %s
                 OR source_file LIKE %s
                 OR title = %s
              )
            ORDER BY updated_at DESC NULLS LAST
            LIMIT 1
            """,
            (collection, document, f"%\\{document}", document),
        )
        row = cur.fetchone()
        return dict(row) if row else None


def _pg_search(
    conn: psycopg.Connection,
    collection: str,
    query: str,
    top_k: int,
    model: str,
    ollama_url: str,
    timeout: int,
) -> list[dict[str, Any]]:
    query_embedding = _ollama_embed(query, model, ollama_url, timeout)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT
                d.id AS document_id,
                d.collection,
                d.source AS document_source,
                d.document_type,
                d.title,
                d.source_file,
                d.doi,
                d.related_scopus_document_id,
                c.chunk_id,
                c.chunk_index,
                c.chunk_count,
                c.text_kind,
                c.content,
                c.raw_metadata,
                1 - (c.embedding <=> %s) AS score,
                c.embedding <=> %s AS distance,
                COALESCE(c.embedding_model, d.embedding_model) AS embedding_model,
                COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS embedding_dim
            FROM rag_chunks c
            JOIN rag_documents d ON d.id = c.document_id
            WHERE d.collection = %s
              AND COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) = %s
            ORDER BY c.embedding <=> %s
            LIMIT %s
            """,
            (
                Vector(query_embedding),
                Vector(query_embedding),
                collection,
                len(query_embedding),
                Vector(query_embedding),
                top_k,
            ),
        )
        rows = cur.fetchall()
    results: list[dict[str, Any]] = []
    for row in rows:
        metadata = dict(row["raw_metadata"] or {})
        results.append(
            {
                "document_id": row["document_id"],
                "collection": row["collection"],
                "document_source": row["document_source"],
                "document_type": row["document_type"],
                "title": row["title"],
                "source_file": row["source_file"],
                "doi": row["doi"],
                "related_scopus_document_id": row["related_scopus_document_id"],
                "chunk_id": row["chunk_id"],
                "chunk_index": row["chunk_index"],
                "chunk_count": row["chunk_count"],
                "text_kind": row["text_kind"],
                "page_start": metadata.get("page_start"),
                "page_end": metadata.get("page_end"),
                "header_path": metadata.get("header_path"),
                "distance": float(row["distance"]),
                "score": round(float(row["score"]), 4),
                "embedding_model": row["embedding_model"],
                "embedding_dim": row["embedding_dim"],
                "snippet": (row["content"] or "")[:400],
            }
        )
    return results


def _pg_collection_stats(conn: psycopg.Connection, collection: str) -> dict[str, Any]:
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT
                count(DISTINCT d.id) AS documents,
                count(c.id) AS chunks,
                count(DISTINCT d.id) FILTER (WHERE d.source = 'pdf') AS pdf_documents,
                count(DISTINCT d.id) FILTER (WHERE d.source <> 'pdf') AS non_pdf_documents,
                min(d.updated_at) AS oldest_update,
                max(d.updated_at) AS newest_update
            FROM rag_documents d
            LEFT JOIN rag_chunks c ON c.document_id = d.id
            WHERE d.collection = %s
            """,
            (collection,),
        )
        summary = dict(cur.fetchone())
        cur.execute(
            """
            SELECT COALESCE(c.embedding_model, d.embedding_model) AS model,
                   COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS dim,
                   count(c.id) AS chunks
            FROM rag_documents d
            LEFT JOIN rag_chunks c ON c.document_id = d.id
            WHERE d.collection = %s
            GROUP BY 1, 2
            ORDER BY 1, 2
            """,
            (collection,),
        )
        groups = [dict(row) for row in cur.fetchall()]
        cur.execute(
            "SELECT kind, related_collection, metadata FROM rag_collections WHERE name = %s",
            (collection,),
        )
        collection_meta = cur.fetchone()
    return {
        "collection": collection,
        "documents": int(summary["documents"] or 0),
        "chunks": int(summary["chunks"] or 0),
        "pdf_documents": int(summary["pdf_documents"] or 0),
        "non_pdf_documents": int(summary["non_pdf_documents"] or 0),
        "oldest_update": str(summary["oldest_update"]) if summary["oldest_update"] else None,
        "newest_update": str(summary["newest_update"]) if summary["newest_update"] else None,
        "embedding_groups": groups,
        "kind": collection_meta["kind"] if collection_meta else None,
        "related_collection": collection_meta["related_collection"] if collection_meta else None,
        "metadata": dict(collection_meta["metadata"] or {}) if collection_meta else {},
    }


def _pg_structure(conn: psycopg.Connection, collection: str, document: str) -> dict[str, Any]:
    doc = _locate_pdf_document(conn, collection, document)
    if not doc:
        return {"error": f"Documento PDF nao encontrado na collection {collection}: {document}"}
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT chunk_index, chunk_count, raw_metadata
            FROM rag_chunks
            WHERE document_id = %s
            ORDER BY chunk_index
            """,
            (doc["id"],),
        )
        rows = cur.fetchall()
    sections: list[dict[str, Any]] = []
    seen: dict[str, dict[str, Any]] = {}
    for row in rows:
        metadata = dict(row["raw_metadata"] or {})
        header_path = _clean_or_none(metadata.get("header_path"))
        if not header_path:
            continue
        section = seen.get(header_path)
        if section is None:
            section = {
                "header_path": header_path,
                "header_level": metadata.get("header_level", 0),
                "start_chunk_index": row["chunk_index"],
                "end_chunk_index": row["chunk_index"],
                "page_start": metadata.get("page_start"),
                "page_end": metadata.get("page_end"),
            }
            seen[header_path] = section
            sections.append(section)
        else:
            section["end_chunk_index"] = row["chunk_index"]
            section["page_end"] = metadata.get("page_end", section["page_end"])
    return {
        "document_id": doc["id"],
        "collection": collection,
        "title": doc["title"],
        "source_file": doc["source_file"],
        "raw_metadata": dict(doc["raw_metadata"] or {}),
        "sections": sections,
    }


def _pg_section(
    conn: psycopg.Connection,
    collection: str,
    document: str,
    header_path: str | None,
    page_range: tuple[int, int] | None,
) -> dict[str, Any]:
    doc = _locate_pdf_document(conn, collection, document)
    if not doc:
        return {"error": f"Documento PDF nao encontrado na collection {collection}: {document}"}
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT chunk_index, chunk_count, content, raw_metadata
            FROM rag_chunks
            WHERE document_id = %s
            ORDER BY chunk_index
            """,
            (doc["id"],),
        )
        rows = cur.fetchall()
    filtered: list[dict[str, Any]] = []
    for row in rows:
        metadata = dict(row["raw_metadata"] or {})
        if header_path and _clean(metadata.get("header_path")) != _clean(header_path):
            continue
        if page_range:
            start_page, end_page = page_range
            chunk_start = int(metadata.get("page_start", -1))
            chunk_end = int(metadata.get("page_end", -1))
            if chunk_end < start_page or chunk_start > end_page:
                continue
        filtered.append(
            {
                "chunk_index": row["chunk_index"],
                "chunk_count": row["chunk_count"],
                "text": row["content"],
                "header_path": metadata.get("header_path"),
                "header_level": metadata.get("header_level"),
                "page_start": metadata.get("page_start"),
                "page_end": metadata.get("page_end"),
                "prev_chunk_index": metadata.get("prev_chunk_index"),
                "next_chunk_index": metadata.get("next_chunk_index"),
            }
        )
    return {
        "document_id": doc["id"],
        "collection": collection,
        "title": doc["title"],
        "source_file": doc["source_file"],
        "chunks": filtered,
    }


def _get_chroma_client(path: Path):
    if chromadb is None:
        raise RuntimeError("chromadb nao esta instalado neste ambiente.")
    path.mkdir(parents=True, exist_ok=True)
    return chromadb.PersistentClient(path=str(path))


def _list_chroma_collections(client) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    raw_collections = client.list_collections()
    for item in raw_collections:
        name = item.name if hasattr(item, "name") else str(item)
        collection = client.get_collection(name=name)
        metadata = dict(getattr(collection, "metadata", {}) or {})
        result.append(
            {
                "name": name,
                "count": collection.count(),
                "kind": metadata.get("kind"),
                "related_collection": metadata.get("related_collection"),
                "embedding_model": metadata.get("embedding_model"),
                "persist_directory": str(DEFAULT_CHROMA_DIR),
            }
        )
    return sorted(result, key=lambda item: item["name"])


def _chroma_collection_exists(client, name: str) -> bool:
    try:
        client.get_collection(name=name)
        return True
    except Exception:
        return False


def _resolve_chroma_target_collection(client, requested_collection: str) -> dict[str, Any]:
    requested = _normalize_collection(requested_collection)
    effective = requested
    related_collection = None
    if _chroma_collection_exists(client, requested):
        existing = client.get_collection(name=requested)
        metadata = dict(getattr(existing, "metadata", {}) or {})
        if metadata.get("kind") != "pdf" and not _looks_like_pdf_collection(requested):
            effective = f"{requested}_pdf"
            related_collection = requested
    elif requested.endswith("_pdf"):
        base = requested[:-4]
        if _chroma_collection_exists(client, base):
            related_collection = base
    return {
        "requested_collection": requested,
        "effective_collection": effective,
        "related_collection": related_collection,
    }


def _get_or_create_chroma_collection(client, name: str, metadata: dict[str, Any]):
    sanitized = _sanitize_chroma_metadata(metadata)
    collection = client.get_or_create_collection(name=name, metadata=sanitized)
    current = dict(getattr(collection, "metadata", {}) or {})
    merged = _sanitize_chroma_metadata({**current, **sanitized})
    try:
        collection.modify(metadata=merged)
    except Exception:
        pass
    return collection


def _find_chroma_hash(client, sha256: str) -> dict[str, Any] | None:
    for info in _list_chroma_collections(client):
        collection = client.get_collection(name=info["name"])
        result = collection.get(where={"file_sha256": sha256}, include=["metadatas"])
        ids = result.get("ids") or []
        if ids:
            return {"collection": info["name"], "ids": ids}
    return None


def _clone_chroma_file(client, source_collection: str, target_collection: str, sha256: str) -> dict[str, Any]:
    source = client.get_collection(name=source_collection)
    target = client.get_collection(name=target_collection)
    result = source.get(
        where={"file_sha256": sha256},
        include=["documents", "metadatas", "embeddings"],
    )
    ids = result.get("ids") or []
    documents = result.get("documents") or []
    metadatas = result.get("metadatas") or []
    embeddings = result.get("embeddings") or []
    if not ids:
        raise RuntimeError("Nao foi possivel clonar o PDF existente no Chroma.")
    updated_ids: list[str] = []
    updated_metadatas: list[dict[str, Any]] = []
    for index, item_id in enumerate(ids):
        chunk_index = (metadatas[index] or {}).get("chunk_index", index)
        updated_ids.append(f"pdf:{target_collection}:{sha256}:chunk:{chunk_index}")
        metadata = _sanitize_chroma_metadata(dict(metadatas[index] or {}))
        metadata["collection"] = target_collection
        updated_metadatas.append(metadata)
    target.upsert(ids=updated_ids, documents=documents, metadatas=updated_metadatas, embeddings=embeddings)
    return {
        "action": "reused_existing_pdf",
        "source_collection": source_collection,
        "effective_collection": target_collection,
        "chunks": len(updated_ids),
    }


def _index_pdf_pgvector(args: argparse.Namespace, config: dict[str, Any], pdf_path: Path) -> dict[str, Any]:
    conn, pg_info = _connect_pg(args, config)
    try:
        _ensure_pg_schema(conn)
        resolution = _resolve_pg_target_collection(conn, _normalize_collection(args.collection))
        sha256, data = _file_sha256(pdf_path)
        store_pdf = _effective_store_pdf(args)
        pdf_file = _upsert_pdf_file(conn, pdf_path, sha256, data, store_pdf)
        existing_in_target = _find_existing_pdf_doc(conn, pdf_file["id"], resolution["effective_collection"])
        if existing_in_target:
            conn.commit()
            return {
                "backend": "pgvector",
                **resolution,
                "action": "already_indexed",
                "document_id": existing_in_target["id"],
                "pdf_file_id": pdf_file["id"],
                "pdf_sha256": sha256,
                "store_pdf": store_pdf,
                "pg": pg_info,
            }
        existing_any = _find_any_existing_pdf_doc(conn, pdf_file["id"])
        if existing_any:
            cloned = _clone_pdf_document(
                conn,
                existing_any["id"],
                resolution["effective_collection"],
                resolution["requested_collection"],
                resolution["related_collection"],
            )
            cloned["backend"] = "pgvector"
            cloned["pdf_file_id"] = pdf_file["id"]
            cloned["pdf_sha256"] = sha256
            cloned["store_pdf"] = store_pdf
            cloned["pg"] = pg_info
            return cloned
        prepared = _prepare_pdf_chunks(pdf_path, args.method)
        chunks = prepared["chunks"]
        embeddings = _embed_texts_parallel(
            [chunk["text"] for chunk in chunks],
            args.model,
            args.ollama_url,
            args.timeout,
            args.workers,
        )
        dims = sorted({len(item) for item in embeddings})
        if len(dims) != 1:
            raise RuntimeError(f"Dimensoes divergentes retornadas pelo modelo: {dims}")
        related_scopus = _find_related_scopus_document(conn, prepared["doi"], resolution["related_collection"])
        doc_id = _upsert_pdf_document(
            conn,
            collection=resolution["effective_collection"],
            related_collection=resolution["related_collection"],
            pdf_file_id=pdf_file["id"],
            pdf_sha256=sha256,
            source_file=str(pdf_path),
            title=prepared["title"],
            doi=prepared["doi"],
            num_pages=prepared["num_pages"],
            num_chunks=len(chunks),
            method=args.method,
            model=args.model,
            embedding_dim=dims[0],
            related_scopus_document_id=related_scopus["id"] if related_scopus else None,
            extra_metadata={
                "requested_collection": resolution["requested_collection"],
                "store_pdf": store_pdf,
            },
        )
        chunk_count = _replace_pdf_chunks(
            conn,
            doc_id,
            resolution["effective_collection"],
            sha256,
            str(pdf_path),
            prepared["doi"],
            args.method,
            chunks,
            embeddings,
            args.model,
            resolution["related_collection"],
        )
        conn.commit()
        return {
            "backend": "pgvector",
            **resolution,
            "action": "indexed",
            "document_id": doc_id,
            "pdf_file_id": pdf_file["id"],
            "pdf_sha256": sha256,
            "title": prepared["title"],
            "doi": prepared["doi"],
            "chunks": chunk_count,
            "num_pages": prepared["num_pages"],
            "embedding_dim": dims[0],
            "store_pdf": store_pdf,
            "related_scopus_document_id": related_scopus["id"] if related_scopus else None,
            "pg": pg_info,
        }
    finally:
        conn.close()


def _index_pdf_chroma(args: argparse.Namespace, config: dict[str, Any], pdf_path: Path) -> dict[str, Any]:
    chroma_dir = Path(
        _resolve_setting(
            args.chroma_dir,
            ["PDF_INDEXER_CHROMA_DIR"],
            [
                _config_lookup(config, "pdf_indexer", "chroma", "persist_directory"),
                _config_lookup(config, "tool", "pdf_indexer", "chroma", "persist_directory"),
            ],
            str(DEFAULT_CHROMA_DIR),
        )
    )
    client = _get_chroma_client(chroma_dir)
    resolution = _resolve_chroma_target_collection(client, _normalize_collection(args.collection))
    sha256, _ = _file_sha256(pdf_path)
    existing_global = _find_chroma_hash(client, sha256)
    target = _get_or_create_chroma_collection(
        client,
        resolution["effective_collection"],
        {
            "kind": "pdf",
            "related_collection": resolution["related_collection"],
            "requested_collection": resolution["requested_collection"],
            "embedding_model": args.model,
        },
    )
    if existing_global and existing_global["collection"] == resolution["effective_collection"]:
        return {
            "backend": "chroma",
            **resolution,
            "action": "already_indexed",
            "pdf_sha256": sha256,
            "store_pdf": False,
            "persist_directory": str(chroma_dir),
        }
    if existing_global and existing_global["collection"] != resolution["effective_collection"]:
        cloned = _clone_chroma_file(
            client,
            existing_global["collection"],
            resolution["effective_collection"],
            sha256,
        )
        cloned["backend"] = "chroma"
        cloned["pdf_sha256"] = sha256
        cloned["store_pdf"] = False
        cloned["persist_directory"] = str(chroma_dir)
        cloned["requested_collection"] = resolution["requested_collection"]
        cloned["related_collection"] = resolution["related_collection"]
        return cloned
    prepared = _prepare_pdf_chunks(pdf_path, args.method)
    chunks = prepared["chunks"]
    embeddings = _embed_texts_parallel(
        [chunk["text"] for chunk in chunks],
        args.model,
        args.ollama_url,
        args.timeout,
        args.workers,
    )
    ids: list[str] = []
    metadatas: list[dict[str, Any]] = []
    documents: list[str] = []
    for chunk, embedding in zip(chunks, embeddings):
        chunk_id = f"pdf:{resolution['effective_collection']}:{sha256}:chunk:{chunk['chunk_index']}"
        ids.append(chunk_id)
        documents.append(chunk["text"])
        metadatas.append(
            _sanitize_chroma_metadata(
                {
                "kind": "pdf",
                "collection": resolution["effective_collection"],
                "related_collection": resolution["related_collection"],
                "requested_collection": resolution["requested_collection"],
                "file_sha256": sha256,
                "filename": pdf_path.name,
                "source_path": str(pdf_path),
                "title": prepared["title"],
                "doi": prepared["doi"] or "",
                "chunk_index": chunk["chunk_index"],
                "chunk_count": chunk["chunk_count"],
                "page_start": chunk["page_start"],
                "page_end": chunk["page_end"],
                "header_path": chunk.get("header_path", ""),
                "header_level": chunk.get("header_level", 0),
                "chunking_method": args.method,
                "embedding_dim": len(embedding),
                "embedding_model": args.model,
                }
            )
        )
    target.upsert(ids=ids, documents=documents, metadatas=metadatas, embeddings=embeddings)
    return {
        "backend": "chroma",
        **resolution,
        "action": "indexed",
        "pdf_sha256": sha256,
        "title": prepared["title"],
        "doi": prepared["doi"],
        "chunks": len(chunks),
        "num_pages": prepared["num_pages"],
        "embedding_dim": len(embeddings[0]) if embeddings else 0,
        "store_pdf": False,
        "persist_directory": str(chroma_dir),
    }


def _handle_index_pdf(args: argparse.Namespace, config: dict[str, Any]) -> dict[str, Any]:
    pdf_path = Path(args.pdf_path).resolve()
    if not pdf_path.exists() or not pdf_path.is_file():
        return {"error": f"Arquivo PDF nao encontrado: {pdf_path}"}
    if pdf_path.suffix.casefold() != ".pdf":
        return {"error": f"Arquivo nao e PDF: {pdf_path}"}
    if _requires_collection_confirmation(args):
        if args.backend == "chroma":
            client = _get_chroma_client(Path(args.chroma_dir or DEFAULT_CHROMA_DIR))
            collections = _list_chroma_collections(client)
        else:
            conn, _ = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                collections = _list_pg_collections(conn)
            finally:
                conn.close()
        return {
            "error": "Escrita sem collection confirmada.",
            "requires_confirmation": True,
            "recommended_collection": DEFAULT_COLLECTION,
            "collections": collections,
        }
    if args.backend == "pgvector":
        index_args = argparse.Namespace(**{**vars(args), "collection": args.collection or DEFAULT_COLLECTION})
        return _index_pdf_pgvector(index_args, config, pdf_path)
    if args.backend == "chroma":
        index_args = argparse.Namespace(**{**vars(args), "collection": args.collection or DEFAULT_COLLECTION})
        return _index_pdf_chroma(index_args, config, pdf_path)
    index_args = argparse.Namespace(**{**vars(args), "collection": args.collection or DEFAULT_COLLECTION})
    pg_result = _index_pdf_pgvector(index_args, config, pdf_path)
    chroma_result = _index_pdf_chroma(index_args, config, pdf_path)
    return {"backend": "both", "pgvector": pg_result, "chroma": chroma_result}


def _handle_index_files(args: argparse.Namespace, config: dict[str, Any]) -> dict[str, Any]:
    results = []
    for raw_path in args.paths:
        file_args = argparse.Namespace(**{**vars(args), "pdf_path": raw_path})
        try:
            results.append(_handle_index_pdf(file_args, config))
        except requests.HTTPError as exc:
            results.append(
                {
                    "backend": args.backend,
                    "file": str(Path(raw_path).resolve()),
                    "error": f"Erro HTTP: {exc}",
                    "status_code": exc.response.status_code if exc.response else None,
                    "response_text": exc.response.text[:1000] if exc.response is not None else None,
                }
            )
        except Exception as exc:
            results.append(
                {
                    "backend": args.backend,
                    "file": str(Path(raw_path).resolve()),
                    "error": str(exc),
                    "error_type": type(exc).__name__,
                }
            )
    return {"matched_files": len(args.paths), "results": results}


def _handle_index_dir(args: argparse.Namespace, config: dict[str, Any]) -> dict[str, Any]:
    directory = Path(args.dir).resolve()
    if not directory.is_dir():
        return {"error": f"Diretorio nao encontrado: {directory}"}
    paths = sorted(directory.rglob("*.pdf") if args.recursive else directory.glob("*.pdf"))
    results = _handle_index_files(
        argparse.Namespace(**{**vars(args), "paths": [str(path) for path in paths]}),
        config,
    )["results"]
    return {
        "directory": str(directory),
        "matched_files": len(paths),
        "results": results,
    }


def _handle_chunk(args: argparse.Namespace) -> dict[str, Any]:
    pdf_path = Path(args.pdf_path).resolve()
    if not pdf_path.exists() or not pdf_path.is_file():
        return {"error": f"Arquivo PDF nao encontrado: {pdf_path}"}
    prepared = _prepare_pdf_chunks(pdf_path, args.method)
    preview = []
    for chunk in prepared["chunks"][: min(20, len(prepared["chunks"]))]:
        preview.append(
            {
                "chunk_index": chunk["chunk_index"],
                "header_path": chunk.get("header_path", ""),
                "page_start": chunk["page_start"],
                "page_end": chunk["page_end"],
                "text_preview": chunk["text"][:500],
                "text_length": len(chunk["text"]),
            }
        )
    return {
        "file": str(pdf_path),
        "title": prepared["title"],
        "doi": prepared["doi"],
        "num_pages": prepared["num_pages"],
        "num_chunks": len(prepared["chunks"]),
        "method": args.method,
        "chunks": preview,
    }


def _handle_download(args: argparse.Namespace) -> dict[str, Any]:
    response = requests.get(args.url, timeout=args.timeout)
    response.raise_for_status()
    target_dir = Path(args.dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    filename = Path(requests.utils.urlparse(args.url).path).name or "download.pdf"
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    target = target_dir / filename
    target.write_bytes(response.content)
    return {"url": args.url, "file": str(target), "size_bytes": len(response.content)}


def _parse_page_range(raw_value: str | None) -> tuple[int, int] | None:
    text = _clean_or_none(raw_value)
    if not text:
        return None
    if "-" not in text:
        page = int(text)
        return page, page
    left, right = text.split("-", 1)
    return int(left), int(right)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Index and search local PDFs with pgvector or Chroma.",
    )
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--backend", choices=["pgvector", "chroma", "both"], default="pgvector")
    common.add_argument("--collection")
    common.add_argument("--ollama-url", default=DEFAULT_OLLAMA_URL)
    common.add_argument("--model", default=DEFAULT_MODEL)
    common.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    common.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    common.add_argument("--store-pdf", dest="store_pdf", default=None)
    common.add_argument("--pg-user")
    common.add_argument("--pg-password")
    common.add_argument("--pg-host")
    common.add_argument("--pg-port", type=int)
    common.add_argument("--pg-database")
    common.add_argument("--chroma-dir")

    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("help", help="Show commands and examples.", parents=[common])

    download = subparsers.add_parser("download", help="Download a PDF to a local directory.", parents=[common])
    download.add_argument("url")
    download.add_argument("--dir", required=True)

    chunk = subparsers.add_parser("chunk", help="Preview chunks for a PDF without indexing.", parents=[common])
    chunk.add_argument("pdf_path")
    chunk.add_argument("--method", choices=["header", "s2"], default="header")

    index_pdf = subparsers.add_parser("index-pdf", help="Index one PDF.", parents=[common])
    index_pdf.add_argument("pdf_path")
    index_pdf.add_argument("--method", choices=["header", "s2"], default="header")

    index_dir = subparsers.add_parser("index-dir", help="Index every PDF in a directory.", parents=[common])
    index_dir.add_argument("--dir", required=True)
    index_dir.add_argument("--recursive", action="store_true")
    index_dir.add_argument("--method", choices=["header", "s2"], default="header")

    index_files = subparsers.add_parser("index-files", help="Index one or more PDFs.", parents=[common])
    index_files.add_argument("paths", nargs="+")
    index_files.add_argument("--method", choices=["header", "s2"], default="header")

    subparsers.add_parser("list-collections", help="List available collections.", parents=[common])

    stats = subparsers.add_parser("stats", help="Show collection stats.", parents=[common])

    search = subparsers.add_parser("search", help="Run semantic search.", parents=[common])
    search.add_argument("--query", required=True)
    search.add_argument("--top-k", type=int, default=DEFAULT_TOP_K)
    search.add_argument("--context-window", type=int, default=DEFAULT_CONTEXT_WINDOW)

    structure = subparsers.add_parser("structure", help="Show PDF structure for one indexed document.", parents=[common])
    structure.add_argument("document")

    section = subparsers.add_parser("section", help="Fetch a section or page range from one indexed PDF.", parents=[common])
    section.add_argument("document")
    section.add_argument("--header-path")
    section.add_argument("--pages")

    return parser


def _help_payload(parser: argparse.ArgumentParser) -> dict[str, Any]:
    return {
        "skill": "pdf-indexer",
        "default_collection": DEFAULT_COLLECTION,
        "default_backend": "pgvector",
        "default_model": DEFAULT_MODEL,
        "commands": [
            {"name": "help", "example": "python ...\\pdf_indexer.py"},
            {"name": "download", "example": "python ...\\pdf_indexer.py download https://host/paper.pdf --dir D:\\papers"},
            {"name": "chunk", "example": "python ...\\pdf_indexer.py chunk D:\\papers\\paper.pdf --method header"},
            {"name": "index-pdf", "example": "python ...\\pdf_indexer.py index-pdf D:\\papers\\paper.pdf --collection embedding_rag --backend pgvector"},
            {"name": "index-files", "example": "python ...\\pdf_indexer.py index-files D:\\papers\\a.pdf D:\\papers\\b.pdf --collection pdf"},
            {"name": "index-dir", "example": "python ...\\pdf_indexer.py index-dir --dir D:\\papers --recursive --collection pdf"},
            {"name": "list-collections", "example": "python ...\\pdf_indexer.py list-collections --backend pgvector"},
            {"name": "stats", "example": "python ...\\pdf_indexer.py stats --collection pdf"},
            {"name": "search", "example": "python ...\\pdf_indexer.py search --collection pdf --query \"markov chain\" --top-k 5"},
            {"name": "structure", "example": "python ...\\pdf_indexer.py structure my-paper.pdf --collection pdf"},
            {"name": "section", "example": "python ...\\pdf_indexer.py section my-paper.pdf --collection pdf --pages 3-5"},
        ],
        "notes": [
            "Leituras usam a collection informada exatamente como alvo; se nao existir, a skill lista as collections disponiveis.",
            "Escritas em collections nao-PDF existentes devem ser redirecionadas para um alvo com sufixo _pdf.",
            "No backend pgvector, store-pdf assume true por padrao; no Chroma, binarios nao sao armazenados.",
        ],
        "usage": parser.format_help(),
    }


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    config = _read_project_config()

    try:
        if not args.command or args.command == "help":
            _json_print(_help_payload(parser))
            return 0

        if args.command == "download":
            _json_print(_handle_download(args))
            return 0

        if args.command == "chunk":
            _json_print(_handle_chunk(args))
            return 0

        if args.command == "index-pdf":
            _json_print(_handle_index_pdf(args, config))
            return 0

        if args.command == "index-files":
            _json_print(_handle_index_files(args, config))
            return 0

        if args.command == "index-dir":
            _json_print(_handle_index_dir(args, config))
            return 0

        if args.command == "list-collections":
            if args.backend == "both":
                client = _get_chroma_client(Path(args.chroma_dir or DEFAULT_CHROMA_DIR))
                conn, pg_info = _connect_pg(args, config)
                try:
                    _ensure_pg_schema(conn)
                    payload = {
                        "backend": "both",
                        "pgvector": {"pg": pg_info, "collections": _list_pg_collections(conn)},
                        "chroma": {"collections": _list_chroma_collections(client)},
                    }
                finally:
                    conn.close()
                _json_print(payload)
                return 0
            if args.backend == "chroma":
                client = _get_chroma_client(Path(args.chroma_dir or DEFAULT_CHROMA_DIR))
                _json_print({"backend": "chroma", "collections": _list_chroma_collections(client)})
                return 0
            conn, pg_info = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                payload = {"backend": "pgvector", "pg": pg_info, "collections": _list_pg_collections(conn)}
            finally:
                conn.close()
            _json_print(payload)
            return 0

        if args.command == "stats":
            args.collection = args.collection or DEFAULT_COLLECTION
            if args.backend == "chroma":
                client = _get_chroma_client(Path(args.chroma_dir or DEFAULT_CHROMA_DIR))
                if not _chroma_collection_exists(client, args.collection):
                    _json_print({"error": f"Collection nao encontrada: {args.collection}", "collections": _list_chroma_collections(client)})
                    return 2
                collection = client.get_collection(name=args.collection)
                _json_print(
                    {
                        "backend": "chroma",
                        "collection": args.collection,
                        "count": collection.count(),
                        "metadata": dict(getattr(collection, "metadata", {}) or {}),
                    }
                )
                return 0
            conn, _ = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                if not _pg_collection_exists(conn, args.collection):
                    _json_print(_list_missing_collection_response(conn, args.collection))
                    return 2
                _json_print(_pg_collection_stats(conn, args.collection))
                return 0
            finally:
                conn.close()

        if args.command == "search":
            args.collection = args.collection or DEFAULT_COLLECTION
            if args.backend == "chroma":
                client = _get_chroma_client(Path(args.chroma_dir or DEFAULT_CHROMA_DIR))
                if not _chroma_collection_exists(client, args.collection):
                    _json_print({"error": f"Collection nao encontrada: {args.collection}", "collections": _list_chroma_collections(client)})
                    return 2
                collection = client.get_collection(name=args.collection)
                query_embedding = _ollama_embed(args.query, args.model, args.ollama_url, args.timeout)
                result = collection.query(
                    query_embeddings=[query_embedding],
                    n_results=args.top_k,
                    include=["documents", "metadatas", "distances"],
                )
                rows = []
                for index, item_id in enumerate(result.get("ids", [[]])[0]):
                    metadata = (result.get("metadatas", [[]])[0] or [])[index]
                    document = (result.get("documents", [[]])[0] or [])[index]
                    distance = (result.get("distances", [[]])[0] or [])[index]
                    rows.append(
                        {
                            "chunk_id": item_id,
                            "distance": distance,
                            "title": metadata.get("title") if metadata else None,
                            "source_file": metadata.get("source_path") if metadata else None,
                            "page_start": metadata.get("page_start") if metadata else None,
                            "page_end": metadata.get("page_end") if metadata else None,
                            "header_path": metadata.get("header_path") if metadata else None,
                            "snippet": (document or "")[:400],
                        }
                    )
                _json_print({"backend": "chroma", "collection": args.collection, "results": rows})
                return 0
            conn, _ = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                if not _pg_collection_exists(conn, args.collection):
                    _json_print(_list_missing_collection_response(conn, args.collection))
                    return 2
                _json_print(
                    {
                        "backend": "pgvector",
                        "collection": args.collection,
                        "query": args.query,
                        "results": _pg_search(
                            conn,
                            args.collection,
                            args.query,
                            args.top_k,
                            args.model,
                            args.ollama_url,
                            args.timeout,
                        ),
                    }
                )
                return 0
            finally:
                conn.close()

        if args.command == "structure":
            args.collection = args.collection or DEFAULT_COLLECTION
            conn, _ = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                if not _pg_collection_exists(conn, args.collection):
                    _json_print(_list_missing_collection_response(conn, args.collection))
                    return 2
                _json_print(_pg_structure(conn, args.collection, args.document))
                return 0
            finally:
                conn.close()

        if args.command == "section":
            args.collection = args.collection or DEFAULT_COLLECTION
            conn, _ = _connect_pg(args, config)
            try:
                _ensure_pg_schema(conn)
                if not _pg_collection_exists(conn, args.collection):
                    _json_print(_list_missing_collection_response(conn, args.collection))
                    return 2
                if not args.header_path and not args.pages:
                    _json_print({"error": "Informe --header-path ou --pages."})
                    return 2
                _json_print(
                    _pg_section(
                        conn,
                        args.collection,
                        args.document,
                        args.header_path,
                        _parse_page_range(args.pages),
                    )
                )
                return 0
            finally:
                conn.close()

        parser.error(f"Unknown command: {args.command}")
        return 2
    except requests.HTTPError as exc:
        _json_print({"error": f"Erro HTTP: {exc}", "status_code": exc.response.status_code if exc.response else None})
        return 1
    except psycopg.Error as exc:
        _json_print({"error": f"Erro PostgreSQL: {exc}"})
        return 1
    except Exception as exc:
        _json_print({"error": str(exc), "type": exc.__class__.__name__})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
