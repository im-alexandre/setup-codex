"""CLI global para indexacao e busca Scopus em PostgreSQL + pgvector."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from pathlib import Path
from typing import Any

import psycopg
import requests
import rispy
from pgvector import Vector
from pgvector.psycopg import register_vector
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb


DEFAULT_POSTGRES_DSN = "postgresql://rag:rag@localhost:5432/ragdb"
DEFAULT_OLLAMA_URL = "http://localhost:11434"
DEFAULT_MODEL = "embeddinggemma"
DEFAULT_EMBEDDING_DIM = 768
CHUNK_SIZE = 4000
CHUNK_OVERLAP = 200


def _json_print(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _clean_or_none(value: Any) -> str | None:
    text = _clean(value)
    return text or None


def _normalize_doi(value: Any) -> str | None:
    text = _clean(value).casefold()
    if not text:
        return None
    for prefix in ("https://doi.org/", "http://doi.org/", "doi:"):
        if text.startswith(prefix):
            text = text[len(prefix) :]
    return text.rstrip(".") or None


def _as_int(value: Any) -> int | None:
    try:
        text = _clean(value)
        return int(text[:4]) if text else None
    except (TypeError, ValueError):
        return None


def _hash(*parts: Any, length: int = 16) -> str:
    text = "|".join(_clean(part) for part in parts)
    return hashlib.sha1(text.casefold().encode("utf-8")).hexdigest()[:length]


def _content_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _normalize_collection(collection: str | None) -> str:
    name = _clean(collection)
    if not name:
        raise ValueError("Informe explicitamente --collection.")
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", name):
        raise ValueError(
            "Collection invalida. Use apenas letras, numeros, ponto, hifen ou underscore."
        )
    return name


def _connect(dsn: str) -> psycopg.Connection:
    conn = psycopg.connect(dsn)
    register_vector(conn)
    return conn


def _ensure_schema(conn: psycopg.Connection) -> None:
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
    conn.commit()


def _list_collections(conn: psycopg.Connection) -> list[dict[str, Any]]:
    _ensure_schema(conn)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT
                d.collection AS name,
                count(DISTINCT d.id) AS documents,
                count(c.id) AS chunks,
                COALESCE(c.embedding_model, d.embedding_model, 'embeddinggemma') AS embedding_model,
                COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS embedding_dim
            FROM rag_documents d
            LEFT JOIN rag_chunks c ON c.document_id = d.id
            WHERE d.collection IS NOT NULL AND d.collection <> ''
            GROUP BY d.collection, 4, 5
            ORDER BY d.collection, 4, 5
            """
        )
        rows = cur.fetchall()
    return [
        {
            "name": row["name"],
            "documents": int(row["documents"]),
            "chunks": int(row["chunks"]),
            "count": int(row["chunks"]),
            "backend": "pgvector",
            "embedding_model": row["embedding_model"],
            "embedding_dim": row["embedding_dim"],
        }
        for row in rows
    ]


def _collection_exists(conn: psycopg.Connection, collection: str) -> bool:
    _ensure_schema(conn)
    with conn.cursor() as cur:
        cur.execute(
            "SELECT EXISTS (SELECT 1 FROM rag_documents WHERE collection = %s)",
            (collection,),
        )
        return bool(cur.fetchone()[0])


def _missing_collection(conn: psycopg.Connection, for_index: bool) -> dict[str, Any]:
    message = "Informe explicitamente --collection com uma collection/tabela do PostgreSQL."
    if for_index:
        message += " Para indexar, voce tambem pode informar o nome de uma nova collection."
    return {
        "error": message,
        "backend": "pgvector",
        "collections": _list_collections(conn),
    }


def _require_collection(
    args: argparse.Namespace,
    conn: psycopg.Connection,
    *,
    for_index: bool,
) -> str | None:
    try:
        collection = _normalize_collection(args.collection)
    except ValueError:
        _json_print(_missing_collection(conn, for_index=for_index))
        return None
    if not for_index and not _collection_exists(conn, collection):
        _json_print(
            {
                "error": f"Collection nao encontrada no PostgreSQL: {collection}",
                "backend": "pgvector",
                "collections": _list_collections(conn),
            }
        )
        return None
    return collection


def _ollama_embed(text: str, model: str, ollama_url: str, timeout: int) -> list[float]:
    base = ollama_url.rstrip("/")
    response = requests.post(
        f"{base}/api/embed",
        json={"model": model, "input": text},
        timeout=timeout,
    )
    if response.status_code == 404:
        response = requests.post(
            f"{base}/api/embeddings",
            json={"model": model, "prompt": text},
            timeout=timeout,
        )
    response.raise_for_status()
    data = response.json()
    if isinstance(data.get("embeddings"), list) and data["embeddings"]:
        embedding = data["embeddings"][0]
    else:
        embedding = data.get("embedding")
    if not isinstance(embedding, list) or not embedding:
        raise RuntimeError(f"Resposta Ollama sem embedding: {data}")
    return [float(value) for value in embedding]


def _split_chunks(text: str) -> list[str]:
    text = text.strip()
    if not text:
        return []
    if len(text) <= CHUNK_SIZE:
        return [text]
    chunks: list[str] = []
    start = 0
    while start < len(text):
        end = min(start + CHUNK_SIZE, len(text))
        chunk = text[start:end].strip()
        if chunk:
            chunks.append(chunk)
        if end == len(text):
            break
        start = max(0, end - CHUNK_OVERLAP)
    return chunks


def _record_key(record: dict[str, Any]) -> str:
    doi = _normalize_doi(record.get("doi"))
    if doi:
        return f"doi:{doi}"
    eid = _clean_or_none(record.get("eid"))
    if eid:
        return f"eid:{eid}"
    return f"title:{_hash(record.get('title'), record.get('year'), record.get('authors'))}"


def _normalize_record(record: dict[str, Any], source_path: Path, row_number: int) -> dict[str, Any] | None:
    title = _clean(record.get("title"))
    if not title:
        return None
    abstract = _clean(record.get("abstract"))
    keywords = _clean(record.get("keywords") or record.get("author_keywords"))
    full_text = f"{title}. {abstract} {keywords}".strip()
    return {
        "source": "scopus",
        "external_id": _clean_or_none(record.get("external_id")) or f"{source_path.name}:{row_number}",
        "doi": _normalize_doi(record.get("doi")),
        "eid": _clean_or_none(record.get("eid")),
        "title": title,
        "authors": _clean_or_none(record.get("authors")),
        "year": _as_int(record.get("year")),
        "source_title": _clean_or_none(record.get("source_title") or record.get("source")),
        "document_type": _clean_or_none(record.get("document_type")),
        "language": _clean_or_none(record.get("language")) or "en",
        "abstract": abstract or None,
        "author_keywords": keywords or None,
        "index_keywords": _clean_or_none(record.get("index_keywords")),
        "source_file": str(source_path),
        "raw_metadata": dict(record),
        "full_text": full_text,
    }


def _read_csv(path: Path) -> list[dict[str, Any]]:
    csv.field_size_limit(10_000_000)
    records: list[dict[str, Any]] = []
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            records.append(
                {
                    "title": row.get("Title"),
                    "abstract": row.get("Abstract"),
                    "keywords": row.get("Author Keywords") or row.get("Index Keywords"),
                    "authors": row.get("Authors"),
                    "year": row.get("Year"),
                    "source_title": row.get("Source title"),
                    "doi": row.get("DOI"),
                    "eid": row.get("EID"),
                    "document_type": row.get("Document Type"),
                    "language": row.get("Language of Original Document"),
                }
            )
    return records


def _coerce_ris(value: Any) -> str:
    if isinstance(value, list):
        return "; ".join(_clean(item) for item in value if _clean(item))
    return _clean(value)


def _read_ris(path: Path) -> list[dict[str, Any]]:
    with path.open(encoding="utf-8-sig") as handle:
        entries = rispy.load(handle)
    records: list[dict[str, Any]] = []
    for entry in entries:
        records.append(
            {
                "title": _coerce_ris(entry.get("title") or entry.get("primary_title")),
                "abstract": _coerce_ris(entry.get("abstract")),
                "keywords": _coerce_ris(entry.get("keywords")),
                "authors": _coerce_ris(entry.get("authors")),
                "year": _coerce_ris(entry.get("year") or entry.get("publication_year")),
                "source_title": _coerce_ris(
                    entry.get("journal_name")
                    or entry.get("secondary_title")
                    or entry.get("periodical_name")
                ),
                "doi": _coerce_ris(entry.get("doi")),
                "document_type": _coerce_ris(entry.get("type_of_reference")),
            }
        )
    return records


def _upsert_document(
    conn: psycopg.Connection,
    collection: str,
    record: dict[str, Any],
    model: str,
    embedding_dim: int,
    full_text_hash: str,
) -> int:
    key = f"scopus:{collection}:{_record_key(record)}"
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO rag_documents (
                source, external_id, doi, eid, title, authors, year, source_title,
                document_type, language, abstract, author_keywords, index_keywords,
                source_file, raw_metadata, document_key, collection, embedding_model,
                embedding_dim, content_hash, updated_at
            )
            VALUES (
                %(source)s, %(external_id)s, %(doi)s, %(eid)s, %(title)s,
                %(authors)s, %(year)s, %(source_title)s, %(document_type)s,
                %(language)s, %(abstract)s, %(author_keywords)s, %(index_keywords)s,
                %(source_file)s, %(raw_metadata)s, %(document_key)s, %(collection)s,
                %(embedding_model)s, %(embedding_dim)s, %(content_hash)s, now()
            )
            ON CONFLICT (document_key) DO UPDATE SET
                source = EXCLUDED.source,
                external_id = EXCLUDED.external_id,
                doi = EXCLUDED.doi,
                eid = EXCLUDED.eid,
                title = EXCLUDED.title,
                authors = EXCLUDED.authors,
                year = EXCLUDED.year,
                source_title = EXCLUDED.source_title,
                document_type = EXCLUDED.document_type,
                language = EXCLUDED.language,
                abstract = EXCLUDED.abstract,
                author_keywords = EXCLUDED.author_keywords,
                index_keywords = EXCLUDED.index_keywords,
                source_file = EXCLUDED.source_file,
                raw_metadata = EXCLUDED.raw_metadata,
                collection = EXCLUDED.collection,
                embedding_model = EXCLUDED.embedding_model,
                embedding_dim = EXCLUDED.embedding_dim,
                content_hash = EXCLUDED.content_hash,
                updated_at = now()
            RETURNING id
            """,
            {
                **record,
                "raw_metadata": Jsonb(record["raw_metadata"]),
                "document_key": key,
                "collection": collection,
                "embedding_model": model,
                "embedding_dim": embedding_dim,
                "content_hash": full_text_hash,
            },
        )
        return int(cur.fetchone()[0])


def _replace_chunks(
    conn: psycopg.Connection,
    document_id: int,
    collection: str,
    record: dict[str, Any],
    chunks: list[str],
    embeddings: list[list[float]],
    model: str,
    embedding_dim: int,
    full_text_hash: str,
) -> int:
    slug = _hash(collection, _record_key(record), record.get("source_file"), length=24)
    prefix = f"scopus:{collection}:{slug}:chunk"
    with conn.cursor() as cur:
        cur.execute("DELETE FROM rag_chunks WHERE document_id = %s", (document_id,))
        for index, (chunk, embedding) in enumerate(zip(chunks, embeddings)):
            cur.execute(
                """
                INSERT INTO rag_chunks (
                    document_id, chunk_id, chunk_index, chunk_count, text_kind,
                    content, embedding, raw_metadata, embedding_model, embedding_dim,
                    content_hash, updated_at
                )
                VALUES (
                    %(document_id)s, %(chunk_id)s, %(chunk_index)s, %(chunk_count)s,
                    'scopus_record_chunk', %(content)s, %(embedding)s, %(raw_metadata)s,
                    %(embedding_model)s, %(embedding_dim)s, %(content_hash)s, now()
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
                """
                ,
                {
                    "document_id": document_id,
                    "chunk_id": f"{prefix}{index}",
                    "chunk_index": index,
                    "chunk_count": len(chunks),
                    "content": chunk,
                    "embedding": Vector(embedding),
                    "raw_metadata": Jsonb(
                        {
                            **record["raw_metadata"],
                            "collection": collection,
                            "chunk_index": index,
                            "chunk_count": len(chunks),
                        }
                    ),
                    "embedding_model": model,
                    "embedding_dim": embedding_dim,
                    "content_hash": full_text_hash,
                },
            )
    return len(chunks)


def _index_records(
    conn: psycopg.Connection,
    source_path: Path,
    raw_records: list[dict[str, Any]],
    args: argparse.Namespace,
    collection: str,
) -> dict[str, Any]:
    _ensure_schema(conn)
    added = 0
    skipped = 0
    chunks_total = 0
    for row_number, raw_record in enumerate(raw_records, start=1):
        record = _normalize_record(raw_record, source_path, row_number)
        if record is None:
            skipped += 1
            continue
        chunks = _split_chunks(record["full_text"])
        if not chunks:
            skipped += 1
            continue
        embeddings = [
            _ollama_embed(chunk, args.model, args.ollama_url, args.timeout)
            for chunk in chunks
        ]
        dims = {len(embedding) for embedding in embeddings}
        if len(dims) != 1:
            raise RuntimeError(f"Dimensoes divergentes retornadas pelo modelo: {sorted(dims)}")
        embedding_dim = dims.pop()
        digest = _content_hash(record["full_text"])
        document_id = _upsert_document(
            conn, collection, record, args.model, embedding_dim, digest
        )
        chunks_total += _replace_chunks(
            conn,
            document_id,
            collection,
            record,
            chunks,
            embeddings,
            args.model,
            embedding_dim,
            digest,
        )
        added += 1
        if added % 50 == 0:
            conn.commit()
    conn.commit()
    return {
        "file": str(source_path),
        "collection": collection,
        "records_indexed": added,
        "records_skipped": skipped,
        "chunks": chunks_total,
        "backend": "pgvector",
    }


def _index_file(
    conn: psycopg.Connection,
    path: Path,
    args: argparse.Namespace,
    collection: str,
) -> dict[str, Any]:
    if not path.exists():
        return {"file": str(path), "error": "Arquivo nao encontrado"}
    suffix = path.suffix.casefold()
    if suffix == ".csv":
        records = _read_csv(path)
    elif suffix == ".ris":
        records = _read_ris(path)
    else:
        return {"file": str(path), "error": "Tipo nao suportado; use .csv ou .ris"}
    return _index_records(conn, path, records, args, collection)


def _collection_stats(conn: psycopg.Connection, collection: str) -> dict[str, Any]:
    _ensure_schema(conn)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT
                count(DISTINCT d.id) AS documents,
                count(c.id) AS chunks,
                min(d.updated_at) AS oldest_update,
                max(d.updated_at) AS newest_update
            FROM rag_documents d
            LEFT JOIN rag_chunks c ON c.document_id = d.id
            WHERE d.collection = %(collection)s
            """,
            {"collection": collection},
        )
        row = cur.fetchone()
        cur.execute(
            """
            SELECT COALESCE(c.embedding_model, d.embedding_model) AS model,
                   COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS dim,
                   count(c.id) AS chunks
            FROM rag_documents d
            LEFT JOIN rag_chunks c ON c.document_id = d.id
            WHERE d.collection = %(collection)s
            GROUP BY 1, 2
            ORDER BY 1, 2
            """,
            {"collection": collection},
        )
        model_dims = cur.fetchall()
    return {
        "collection": collection,
        "backend": "pgvector",
        "documents": int(row["documents"]),
        "chunks": int(row["chunks"]),
        "count": int(row["chunks"]),
        "oldest_update": str(row["oldest_update"]) if row["oldest_update"] else None,
        "newest_update": str(row["newest_update"]) if row["newest_update"] else None,
        "embedding_groups": [dict(item) for item in model_dims],
    }


def _search(conn: psycopg.Connection, args: argparse.Namespace, collection: str) -> dict[str, Any]:
    query_embedding = _ollama_embed(args.query, args.model, args.ollama_url, args.timeout)
    _ensure_schema(conn)
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(
            """
            SELECT
                1 - (c.embedding <=> %(query_embedding)s) AS score,
                c.embedding <=> %(query_embedding)s AS distance,
                d.title,
                d.authors,
                d.year,
                d.doi,
                d.source_title AS source,
                c.content,
                c.chunk_id,
                COALESCE(c.embedding_model, d.embedding_model) AS embedding_model,
                COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) AS embedding_dim
            FROM rag_chunks c
            JOIN rag_documents d ON d.id = c.document_id
            WHERE d.collection = %(collection)s
              AND COALESCE(c.embedding_dim, d.embedding_dim, vector_dims(c.embedding)) = %(query_dim)s
            ORDER BY c.embedding <=> %(query_embedding)s
            LIMIT %(top_k)s
            """,
            {
                "query_embedding": Vector(query_embedding),
                "query_dim": len(query_embedding),
                "collection": collection,
                "top_k": args.top_k,
            },
        )
        rows = cur.fetchall()
    return {
        "backend": "pgvector",
        "collection": collection,
        "original_query": _clean_or_none(getattr(args, "original_query", None)) or args.query,
        "english_query_used": args.query,
        "top_k": int(args.top_k),
        "results": [
            {
                "score": round(float(row["score"]), 4),
                "distance": float(row["distance"]),
                "title": row["title"],
                "authors": row["authors"],
                "year": row["year"],
                "doi": row["doi"],
                "source": row["source"],
                "snippet": (row["content"] or "")[:300],
                "chunk_id": row["chunk_id"],
                "embedding_model": row["embedding_model"],
                "embedding_dim": row["embedding_dim"],
            }
            for row in rows
        ],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Scopus semantic indexing/search using PostgreSQL + pgvector."
    )
    parser.add_argument("--postgres-dsn", default=DEFAULT_POSTGRES_DSN)
    parser.add_argument("--ollama-url", default=DEFAULT_OLLAMA_URL)
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument("--timeout", type=int, default=180)

    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("list-collections", help="List PostgreSQL/pgvector collections.")

    stats = subparsers.add_parser("stats", help="Show collection stats.")
    stats.add_argument("--collection")

    search = subparsers.add_parser("search", help="Run semantic search.")
    search.add_argument("--collection")
    search.add_argument("--query", required=True)
    search.add_argument("--original-query")
    search.add_argument("--top-k", type=int, default=5)

    index_csv = subparsers.add_parser("index-csv", help="Index one Scopus CSV export.")
    index_csv.add_argument("--collection")
    index_csv.add_argument("--csv-path", required=True)

    index_ris = subparsers.add_parser("index-ris", help="Index one Scopus RIS export.")
    index_ris.add_argument("--collection")
    index_ris.add_argument("--ris-path", required=True)

    index_files = subparsers.add_parser("index-files", help="Index one or more CSV/RIS files.")
    index_files.add_argument("--collection")
    index_files.add_argument("paths", nargs="+")

    index_dir = subparsers.add_parser(
        "index-dir", help="Recursively index all .csv and .ris files in a directory."
    )
    index_dir.add_argument("--collection")
    index_dir.add_argument("--dir", required=True)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    with _connect(args.postgres_dsn) as conn:
        _ensure_schema(conn)

        if args.command == "list-collections":
            _json_print(
                {
                    "backend": "pgvector",
                    "postgres_dsn": args.postgres_dsn,
                    "collections": _list_collections(conn),
                }
            )
            return 0

        if args.command == "stats":
            collection = _require_collection(args, conn, for_index=False)
            if collection is None:
                return 2
            _json_print(_collection_stats(conn, collection))
            return 0

        if args.command == "search":
            collection = _require_collection(args, conn, for_index=False)
            if collection is None:
                return 2
            _json_print(_search(conn, args, collection))
            return 0

        if args.command == "index-csv":
            collection = _require_collection(args, conn, for_index=True)
            if collection is None:
                return 2
            _json_print(_index_file(conn, Path(args.csv_path), args, collection))
            return 0

        if args.command == "index-ris":
            collection = _require_collection(args, conn, for_index=True)
            if collection is None:
                return 2
            _json_print(_index_file(conn, Path(args.ris_path), args, collection))
            return 0

        if args.command == "index-files":
            collection = _require_collection(args, conn, for_index=True)
            if collection is None:
                return 2
            _json_print(
                {
                    "collection": collection,
                    "backend": "pgvector",
                    "results": [
                        _index_file(conn, Path(path), args, collection)
                        for path in args.paths
                    ],
                }
            )
            return 0

        if args.command == "index-dir":
            collection = _require_collection(args, conn, for_index=True)
            if collection is None:
                return 2
            source_dir = Path(args.dir)
            if not source_dir.is_dir():
                _json_print({"error": f"Diretorio nao encontrado: {source_dir}"})
                return 1
            paths = sorted(
                [
                    *source_dir.rglob("*.csv"),
                    *source_dir.rglob("*.ris"),
                ],
                key=lambda item: str(item).casefold(),
            )
            _json_print(
                {
                    "collection": collection,
                    "directory": str(source_dir),
                    "matched_files": len(paths),
                    "backend": "pgvector",
                    "results": [_index_file(conn, path, args, collection) for path in paths],
                }
            )
            return 0

    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
