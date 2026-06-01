"""SQLite FTS5 RAG indexing and retrieval for Research Radar."""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Iterable, Optional

from radar_db import connect, init_db, utc_now
from prompt_templates import render_template


SKILLS_DIR = Path(__file__).resolve().parents[2]
MY_NOTES_DIR = SKILLS_DIR / "my_notes"
IDEA_NOTES_DIR = SKILLS_DIR / "idea_notes"
REVIEW_REPORTS_DIR = SKILLS_DIR / "review_reports"
ROADMAP_REPORTS_DIR = SKILLS_DIR / "roadmap_reports"
PDF_TEXTS_DIR = SKILLS_DIR / "pdf_texts"


def _hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()


def _clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def chunk_text(text: str, size: int = 1400, overlap: int = 180) -> list[str]:
    text = _clean_text(text)
    if not text:
        return []
    chunks = []
    start = 0
    while start < len(text):
        end = min(len(text), start + size)
        chunks.append(text[start:end].strip())
        if end >= len(text):
            break
        start = max(end - overlap, start + 1)
    return [chunk for chunk in chunks if chunk]


def _upsert_document(conn, source_type: str, source_id: str, title: str, profile: str = "", path: str = "", source_url: str = "", content: str = "") -> int:
    now = utc_now()
    content_hash = _hash_text(content)
    conn.execute(
        """
        INSERT INTO rag_documents (source_type, source_id, title, profile, path, source_url, content_hash, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(source_type, source_id, path) DO UPDATE SET
            title=excluded.title,
            profile=excluded.profile,
            source_url=excluded.source_url,
            content_hash=excluded.content_hash,
            updated_at=excluded.updated_at
        """,
        (source_type, source_id or "", title or "", profile or "", path or "", source_url or "", content_hash, now, now),
    )
    row = conn.execute(
        "SELECT id FROM rag_documents WHERE source_type=? AND source_id=? AND path=?",
        (source_type, source_id or "", path or ""),
    ).fetchone()
    return int(row["id"])


def _replace_chunks(conn, document_id: int, source_type: str, title: str, chunks: list[str], source_url: str = "", paper_id: str = "") -> int:
    existing = conn.execute("SELECT id FROM rag_chunks WHERE document_id = ?", (document_id,)).fetchall()
    for row in existing:
        conn.execute("DELETE FROM rag_chunks_fts WHERE rowid = ?", (row["id"],))
    conn.execute("DELETE FROM rag_chunks WHERE document_id = ?", (document_id,))
    now = utc_now()
    count = 0
    for idx, text in enumerate(chunks):
        cursor = conn.execute(
            """
            INSERT INTO rag_chunks (document_id, chunk_index, chunk_text, section_title, token_count, source_url, paper_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (document_id, idx, text, title, len(text), source_url or "", paper_id or "", now),
        )
        chunk_id = int(cursor.lastrowid)
        conn.execute(
            "INSERT INTO rag_chunks_fts(rowid, chunk_text, title, source_type) VALUES (?, ?, ?, ?)",
            (chunk_id, text, title or "", source_type),
        )
        count += 1
    return count


def index_text(source_type: str, source_id: str, title: str, text: str, profile: str = "", path: str = "", source_url: str = "", paper_id: str = "") -> int:
    init_db()
    chunks = chunk_text(text)
    if not chunks:
        return 0
    with connect() as conn:
        doc_id = _upsert_document(conn, source_type, source_id, title, profile, path, source_url, text)
        count = _replace_chunks(conn, doc_id, source_type, title, chunks, source_url, paper_id)
        conn.commit()
    return count


def paper_source_url(paper: dict) -> str:
    for key in ("display_url", "url", "doi_url", "arxiv_url", "pdf_url", "publisher_url", "scholar_url"):
        value = str(paper.get(key) or "").strip()
        if value:
            return value
    doi = str(paper.get("doi") or "").strip()
    return f"https://doi.org/{doi}" if doi else ""


def index_paper(paper_id: str) -> int:
    init_db()
    with connect() as conn:
        paper = conn.execute("SELECT * FROM papers WHERE stable_id = ?", (paper_id,)).fetchone()
        if not paper:
            return 0
        paper = dict(paper)
        note = conn.execute("SELECT * FROM reading_notes WHERE paper_id = ?", (paper_id,)).fetchone()
        note = dict(note) if note else {}
    parts = [
        paper.get("title"),
        paper.get("abstract_original"),
        paper.get("abstract_zh"),
        note.get("paper_topic"),
        note.get("research_background"),
        note.get("research_purpose"),
        note.get("core_method"),
        note.get("main_results"),
        note.get("paper_contribution"),
        note.get("inspiration"),
        note.get("possible_ideas"),
        note.get("user_notes"),
    ]
    return index_text(
        "paper",
        paper_id,
        paper.get("title") or paper_id,
        "\n\n".join(str(part) for part in parts if part),
        paper.get("profile_id") or "",
        "",
        paper_source_url(paper),
        paper_id,
    )


def _iter_markdown_files(*roots: Path) -> Iterable[Path]:
    for root in roots:
        if root.exists():
            yield from root.rglob("*.md")


def index_markdown_file(path: Path, source_type: str, profile: str = "") -> int:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return 0
    rel = str(path)
    title = path.stem
    return index_text(source_type, rel, title, text, profile, rel, rel)


def index_all(profile: str = "") -> dict:
    init_db()
    stats = {"papers": 0, "files": 0, "chunks": 0}
    with connect() as conn:
        if profile:
            paper_rows = conn.execute("SELECT stable_id FROM papers WHERE profile_id = ?", (profile,)).fetchall()
        else:
            paper_rows = conn.execute("SELECT stable_id FROM papers").fetchall()
    for row in paper_rows:
        stats["chunks"] += index_paper(row["stable_id"])
        stats["papers"] += 1
    for path in _iter_markdown_files(MY_NOTES_DIR, IDEA_NOTES_DIR, REVIEW_REPORTS_DIR, ROADMAP_REPORTS_DIR):
        stats["chunks"] += index_markdown_file(path, "markdown", profile)
        stats["files"] += 1
    return stats


def index_pdf_text(paper: dict, pdf_text: str, pdf_path: str = "") -> int:
    return index_text(
        "pdf",
        paper.get("stable_id") or pdf_path,
        paper.get("title") or "PDF",
        pdf_text,
        paper.get("profile_id") or "",
        pdf_path,
        paper.get("pdf_url") or pdf_path,
        paper.get("stable_id") or "",
    )


def _fts_query(question: str) -> str:
    tokens = re.findall(r"[\w\u4e00-\u9fff]{2,}", question or "")
    tokens = tokens[:8]
    if not tokens:
        return '"research"'
    return " OR ".join(f'"{token}"' for token in tokens)


def search(question: str, profile: str = "", source_type: str = "", source_id: str = "", paper_id: str = "", limit: int = 8) -> list[dict]:
    init_db()
    fts = _fts_query(question)
    clauses = ["rag_chunks_fts MATCH ?"]
    params: list = [fts]
    if profile:
        clauses.append("COALESCE(d.profile, '') = ?")
        params.append(profile)
    if source_type:
        clauses.append("d.source_type = ?")
        params.append(source_type)
    if source_id:
        clauses.append("d.source_id = ?")
        params.append(source_id)
    if paper_id:
        clauses.append("(c.paper_id = ? OR d.source_id = ?)")
        params.extend([paper_id, paper_id])
    params.append(limit)
    sql = f"""
        SELECT c.id AS chunk_id, c.chunk_text, c.source_url, c.paper_id,
               d.source_type, d.source_id, d.title, d.profile, d.path,
               bm25(rag_chunks_fts) AS rank
        FROM rag_chunks_fts
        JOIN rag_chunks c ON c.id = rag_chunks_fts.rowid
        JOIN rag_documents d ON d.id = c.document_id
        WHERE {' AND '.join(clauses)}
        ORDER BY rank
        LIMIT ?
    """
    try:
        with connect() as conn:
            results = [dict(row) for row in conn.execute(sql, tuple(params)).fetchall()]
            if results:
                return results
    except Exception:
        pass
    tokens = re.findall(r"[\w\u4e00-\u9fff]{2,}", question or "")[:6]
    like_clauses = []
    like_params: list = []
    for token in tokens or [question[:80]]:
        like_clauses.append("c.chunk_text LIKE ?")
        like_params.append(f"%{token}%")
    extra = []
    if profile:
        extra.append("COALESCE(d.profile, '') = ?")
        like_params.append(profile)
    if source_type:
        extra.append("d.source_type = ?")
        like_params.append(source_type)
    if source_id:
        extra.append("d.source_id = ?")
        like_params.append(source_id)
    if paper_id:
        extra.append("(c.paper_id = ? OR d.source_id = ?)")
        like_params.extend([paper_id, paper_id])
    where = " OR ".join(like_clauses)
    if extra:
        where = f"({where}) AND " + " AND ".join(extra)
    like_params.append(limit)
    with connect() as conn:
        return [
            dict(row)
            for row in conn.execute(
                f"""
                SELECT c.id AS chunk_id, c.chunk_text, c.source_url, c.paper_id,
                       d.source_type, d.source_id, d.title, d.profile, d.path
                FROM rag_chunks c
                JOIN rag_documents d ON d.id = c.document_id
                WHERE {where}
                LIMIT ?
                """,
                tuple(like_params),
            ).fetchall()
        ]


def format_context(chunks: list[dict]) -> str:
    lines = []
    for idx, chunk in enumerate(chunks, 1):
        link = chunk.get("source_url") or chunk.get("path") or chunk.get("source_id") or ""
        lines.append(
            f"[{idx}] {chunk.get('title') or 'Untitled'} | {chunk.get('source_type')} | {link}\n"
            f"{chunk.get('chunk_text')}"
        )
    return "\n\n".join(lines)


def source_list(chunks: list[dict]) -> list[dict]:
    seen = set()
    sources = []
    for chunk in chunks:
        key = (chunk.get("source_type"), chunk.get("source_id"), chunk.get("path"))
        if key in seen:
            continue
        seen.add(key)
        sources.append(
            {
                "source_type": chunk.get("source_type"),
                "source_id": chunk.get("source_id"),
                "paper_id": chunk.get("paper_id"),
                "title": chunk.get("title"),
                "url": chunk.get("source_url") or chunk.get("path"),
            }
        )
    return sources


def save_query(question: str, scope: str, profile: str, source_id: str, answer: str, sources: list[dict]) -> None:
    init_db()
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO rag_queries (question, scope, profile, source_id, answer, sources_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (question, scope, profile, source_id, answer, json.dumps(sources, ensure_ascii=False), utc_now()),
        )
        conn.commit()


def build_answer_prompt(question: str, chunks: list[dict]) -> str:
    fallback = f"""你是 Research Radar 的科研问答助手。只能基于给定的文献库片段、笔记和报告回答；证据不足时请明确说明。

【用户问题】
{question}

【检索到的依据】
{format_context(chunks)}

请用中文输出：
## 结论
## 依据来源
逐条列出来源标题、类型和可追溯链接。
## 可操作建议
"""
    return render_template(
        "rag_answer_prompt",
        {"user_question": question, "rag_context": format_context(chunks)},
        fallback,
    )


def rag_status() -> dict:
    init_db()
    with connect() as conn:
        docs = conn.execute("SELECT COUNT(*) AS c FROM rag_documents").fetchone()["c"]
        chunks = conn.execute("SELECT COUNT(*) AS c FROM rag_chunks").fetchone()["c"]
        queries = conn.execute("SELECT COUNT(*) AS c FROM rag_queries").fetchone()["c"]
        latest = conn.execute("SELECT updated_at FROM rag_documents ORDER BY updated_at DESC LIMIT 1").fetchone()
    return {
        "documents": docs,
        "chunks": chunks,
        "queries": queries,
        "latest_indexed_at": latest["updated_at"] if latest else "",
    }
