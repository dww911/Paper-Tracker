import json
import os
import re
import shutil
import sqlite3
from datetime import datetime
from typing import Dict, Iterable, List, Optional


SKILLS_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
ROOT_DIR = os.path.dirname(os.path.dirname(SKILLS_DIR))
LEGACY_DB_PATH = os.path.join(SKILLS_DIR, "research_radar.db")
ROOT_DB_PATH = os.path.join(ROOT_DIR, "research_radar.db")
DB_PATH = os.environ.get("RESEARCH_RADAR_DB") or ROOT_DB_PATH
FALLBACK_DB_PATH = ROOT_DB_PATH
PROFILES_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "research_profiles.json")


def utc_now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _table_count(db_path: str, table: str) -> int:
    if not os.path.exists(db_path):
        return 0
    try:
        conn = sqlite3.connect(db_path)
        try:
            row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
            return int(row[0] if row else 0)
        finally:
            conn.close()
    except sqlite3.Error:
        return 0


def migrate_legacy_db_if_needed(db_path: str = DB_PATH) -> None:
    if os.environ.get("RESEARCH_RADAR_DB"):
        return
    if os.path.abspath(db_path) != os.path.abspath(ROOT_DB_PATH):
        return
    if not os.path.exists(LEGACY_DB_PATH):
        return
    legacy_count = _table_count(LEGACY_DB_PATH, "papers")
    target_count = _table_count(db_path, "papers")
    if legacy_count <= 0 or target_count > 0:
        return
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    shutil.copy2(LEGACY_DB_PATH, db_path)


def update_paper_abstract_zh(stable_id: str, abstract_zh: str, db_path: str = DB_PATH) -> None:
    if not stable_id or not str(abstract_zh or "").strip():
        return
    with connect(db_path) as conn:
        conn.execute(
            """
            UPDATE papers
            SET abstract_zh = ?,
                abstract_fetch_status = CASE
                    WHEN abstract_is_complete = 1 THEN COALESCE(NULLIF(abstract_fetch_status, ''), 'complete')
                    ELSE COALESCE(abstract_fetch_status, 'translated')
                END,
                updated_at = ?
            WHERE stable_id = ?
            """,
            (str(abstract_zh).strip(), utc_now(), stable_id),
        )


def connect(db_path: str = DB_PATH) -> sqlite3.Connection:
    db_dir = os.path.dirname(db_path)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)
    try:
        conn = sqlite3.connect(db_path)
    except sqlite3.OperationalError:
        if db_path == FALLBACK_DB_PATH:
            raise
        fallback_dir = os.path.dirname(FALLBACK_DB_PATH)
        if fallback_dir:
            os.makedirs(fallback_dir, exist_ok=True)
        conn = sqlite3.connect(FALLBACK_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: str = DB_PATH) -> None:
    migrate_legacy_db_if_needed(db_path)
    with connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS papers (
                stable_id TEXT PRIMARY KEY,
                profile_id TEXT,
                title TEXT NOT NULL,
                authors TEXT,
                year INTEGER,
                journal TEXT,
                doi TEXT,
                issn TEXT,
                url TEXT,
                source TEXT,
                abstract_original TEXT,
                abstract_zh TEXT,
                abstract_source TEXT,
                abstract_is_complete INTEGER DEFAULT 0,
                impact_factor TEXT,
                impact_factor_year TEXT,
                jcr_quartile TEXT,
                cas_quartile TEXT,
                citation_count INTEGER,
                relevance_score INTEGER,
                impact_factor_score INTEGER,
                freshness_score INTEGER,
                citation_score INTEGER,
                final_score INTEGER,
                recommendation_level TEXT,
                publication_year INTEGER,
                system_rating INTEGER,
                user_rating INTEGER,
                rating_reason TEXT,
                reading_status TEXT DEFAULT 'unread',
                favorite INTEGER DEFAULT 0,
                is_favorite INTEGER DEFAULT 0,
                included_in_review INTEGER DEFAULT 0,
                annual_report_year INTEGER,
                pushed_to_wechat INTEGER DEFAULT 0,
                report_path TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS reading_notes (
                paper_id TEXT PRIMARY KEY,
                paper_topic TEXT,
                research_background TEXT,
                research_purpose TEXT,
                core_method TEXT,
                main_results TEXT,
                paper_contribution TEXT,
                inspiration TEXT,
                worth_reading TEXT,
                reason TEXT,
                possible_ideas TEXT,
                user_notes TEXT,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (paper_id) REFERENCES papers(stable_id)
            );

            CREATE TABLE IF NOT EXISTS profiles (
                profile_id TEXT PRIMARY KEY,
                name TEXT,
                display_name TEXT,
                description TEXT,
                config_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile TEXT,
                mode TEXT,
                total_found INTEGER DEFAULT 0,
                kept_after_relevance INTEGER DEFAULT 0,
                new_papers INTEGER DEFAULT 0,
                abstract_completed INTEGER DEFAULT 0,
                if_matched INTEGER DEFAULT 0,
                recommended_count INTEGER DEFAULT 0,
                pushed_count INTEGER DEFAULT 0,
                report_path TEXT,
                run_time TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS review_jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile TEXT,
                topic TEXT,
                review_type TEXT,
                language TEXT,
                paper_count INTEGER DEFAULT 0,
                status TEXT,
                output_path TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_papers_profile ON papers(profile_id);
            CREATE INDEX IF NOT EXISTS idx_papers_level ON papers(recommendation_level);
            CREATE INDEX IF NOT EXISTS idx_papers_year ON papers(year);
            CREATE INDEX IF NOT EXISTS idx_runs_time ON runs(run_time);

            CREATE TABLE IF NOT EXISTS run_candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                stable_id TEXT,
                title TEXT NOT NULL,
                year INTEGER,
                journal TEXT,
                url TEXT,
                relevance_score INTEGER DEFAULT 0,
                status TEXT NOT NULL,
                filter_reason TEXT,
                sort_order INTEGER DEFAULT 0,
                FOREIGN KEY (run_id) REFERENCES runs(id)
            );
            CREATE INDEX IF NOT EXISTS idx_run_candidates_run ON run_candidates(run_id);
            """
        )
        ensure_columns(conn, "papers", {
            "publication_year": "INTEGER",
            "system_rating": "INTEGER",
            "user_rating": "INTEGER",
            "rating_reason": "TEXT",
            "is_favorite": "INTEGER DEFAULT 0",
            "included_in_review": "INTEGER DEFAULT 0",
            "annual_report_year": "INTEGER",
            "arxiv_id": "TEXT",
            "arxiv_url": "TEXT",
            "doi_url": "TEXT",
            "pdf_url": "TEXT",
            "scholar_url": "TEXT",
            "publisher_url": "TEXT",
            "source_url": "TEXT",
            "display_url": "TEXT",
            "is_recommended": "INTEGER DEFAULT 0",
            "is_milestone": "INTEGER DEFAULT 0",
            "milestone_stage": "TEXT",
            "milestone_reason": "TEXT",
            "my_notes_path": "TEXT",
            "last_seen_at": "TEXT",
            "ingestion_tier": "TEXT DEFAULT 'full'",
            "paper_url": "TEXT",
            "eissn": "TEXT",
            "journal_rank_source": "TEXT",
            "journal_matched": "INTEGER DEFAULT 0",
            "journal_match_method": "TEXT",
            "jcr_impact_factor": "TEXT",
            "jcr_year": "TEXT",
            "cas_category": "TEXT",
            "cas_top": "TEXT",
            "cas_warning": "TEXT",
            "cnki_composite_if": "TEXT",
            "cnki_comprehensive_if": "TEXT",
            "core_tags": "TEXT",
            "ccf_rank": "TEXT",
            "journal_quality_score": "INTEGER DEFAULT 0",
            "abstract_fetch_status": "TEXT",
            "last_run_id": "INTEGER",
            "ingest_mode": "TEXT",
            "query_used": "TEXT",
            "filter_reason": "TEXT",
            "is_relevant": "INTEGER DEFAULT 1",
        })
        ensure_columns(conn, "runs", {
            "skipped_low_score": "INTEGER DEFAULT 0",
            "ingested_count": "INTEGER DEFAULT 0",
            "updated_count": "INTEGER DEFAULT 0",
            "skipped_duplicate": "INTEGER DEFAULT 0",
            "star5": "INTEGER DEFAULT 0",
            "star4": "INTEGER DEFAULT 0",
            "star3": "INTEGER DEFAULT 0",
            "star2": "INTEGER DEFAULT 0",
            "star1": "INTEGER DEFAULT 0",
            "run_year": "TEXT",
            "max_results": "INTEGER",
            "data_sources": "TEXT",
            "google_query": "TEXT",
            "ingest_policy": "TEXT",
            "skipped_irrelevant": "INTEGER DEFAULT 0",
            "doi_completed": "INTEGER DEFAULT 0",
            "skipped_cross_profile": "INTEGER DEFAULT 0",
            "filter_stats_json": "TEXT",
        })
        ensure_columns(conn, "reading_notes", {
            "user_notes": "TEXT",
            "my_notes_path": "TEXT",
        })
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS ideas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                body_md TEXT,
                profile_id TEXT,
                status TEXT DEFAULT '想法中',
                tags_json TEXT,
                linked_paper_ids_json TEXT,
                linked_pdf_path TEXT,
                possible_direction TEXT,
                next_tasks TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS citation_collections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                profile_id TEXT,
                format_default TEXT DEFAULT 'gbt7714',
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS citation_items (
                collection_id INTEGER NOT NULL,
                paper_id TEXT NOT NULL,
                sort_order INTEGER DEFAULT 0,
                PRIMARY KEY (collection_id, paper_id)
            );
            CREATE TABLE IF NOT EXISTS progress_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile_id TEXT NOT NULL,
                snapshot_date TEXT NOT NULL,
                paper_count INTEGER DEFAULT 0,
                read_count INTEGER DEFAULT 0,
                todo_count INTEGER DEFAULT 0,
                note_count INTEGER DEFAULT 0,
                idea_count INTEGER DEFAULT 0,
                milestone_total INTEGER DEFAULT 0,
                milestone_read INTEGER DEFAULT 0,
                review_readiness INTEGER DEFAULT 0,
                next_suggestion TEXT,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS prompt_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                task_type TEXT NOT NULL,
                template_path TEXT NOT NULL,
                description TEXT,
                version TEXT DEFAULT '1',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS prompt_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_name TEXT,
                task_type TEXT,
                source_type TEXT,
                source_id TEXT,
                input_summary TEXT,
                output_summary TEXT,
                status TEXT,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS rag_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_type TEXT NOT NULL,
                source_id TEXT,
                title TEXT,
                profile TEXT,
                path TEXT,
                source_url TEXT,
                content_hash TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(source_type, source_id, path)
            );
            CREATE TABLE IF NOT EXISTS rag_chunks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                chunk_index INTEGER NOT NULL,
                chunk_text TEXT NOT NULL,
                section_title TEXT,
                token_count INTEGER DEFAULT 0,
                source_url TEXT,
                paper_id TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES rag_documents(id)
            );
            CREATE TABLE IF NOT EXISTS rag_queries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question TEXT NOT NULL,
                scope TEXT,
                profile TEXT,
                source_id TEXT,
                answer TEXT,
                sources_json TEXT,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS agent_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                agent_name TEXT NOT NULL,
                profile TEXT,
                user_request TEXT,
                plan_json TEXT,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS agent_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                step_index INTEGER NOT NULL,
                name TEXT NOT NULL,
                tool_name TEXT,
                status TEXT NOT NULL,
                message TEXT,
                started_at TEXT,
                finished_at TEXT,
                FOREIGN KEY (run_id) REFERENCES agent_runs(id)
            );
            CREATE TABLE IF NOT EXISTS agent_outputs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                step_id INTEGER,
                output_type TEXT,
                title TEXT,
                content TEXT,
                path TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY (run_id) REFERENCES agent_runs(id),
                FOREIGN KEY (step_id) REFERENCES agent_steps(id)
            );
            CREATE TABLE IF NOT EXISTS roadmap_nodes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile TEXT NOT NULL,
                node_key TEXT NOT NULL,
                title TEXT NOT NULL,
                description TEXT,
                node_type TEXT NOT NULL,
                stage_label TEXT,
                start_year INTEGER,
                end_year INTEGER,
                parent_node_key TEXT,
                importance_level TEXT DEFAULT 'medium',
                reading_order INTEGER DEFAULT 0,
                keywords TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(profile, node_key)
            );
            CREATE TABLE IF NOT EXISTS roadmap_edges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile TEXT NOT NULL,
                from_node_key TEXT NOT NULL,
                to_node_key TEXT NOT NULL,
                edge_type TEXT,
                description TEXT,
                UNIQUE(profile, from_node_key, to_node_key, edge_type)
            );
            CREATE TABLE IF NOT EXISTS roadmap_papers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile TEXT NOT NULL,
                node_key TEXT NOT NULL,
                paper_stable_id TEXT NOT NULL,
                paper_role TEXT DEFAULT 'representative',
                is_must_read INTEGER DEFAULT 0,
                reading_rank INTEGER DEFAULT 0,
                note TEXT,
                UNIQUE(profile, node_key, paper_stable_id)
            );
            """
        )
        conn.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS rag_chunks_fts USING fts5(
                chunk_text,
                title,
                source_type
            )
            """
        )


def ensure_columns(conn: sqlite3.Connection, table: str, columns: Dict[str, str]) -> None:
    existing = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, ddl in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def extract_year(value: str) -> Optional[int]:
    match = re.search(r"(19|20)\d{2}", str(value or ""))
    return int(match.group(0)) if match else None


def as_int(value, default: int = 0) -> int:
    try:
        if value in ("", None):
            return default
        return int(float(str(value)))
    except (TypeError, ValueError):
        return default


def rating_from_score(score: int) -> int:
    if score >= 24:
        return 5
    if score >= 18:
        return 4
    if score >= 12:
        return 3
    if score >= 6:
        return 2
    return 1


def rating_reason(score: int, abstract_complete: bool, impact_factor, citations: int) -> str:
    level = {
        5: "强烈建议精读",
        4: "推荐阅读",
        3: "值得关注",
        2: "可归档备用",
        1: "低优先级",
    }[rating_from_score(score)]
    signals = [f"综合分 {score}/30"]
    signals.append("摘要完整" if abstract_complete else "摘要待补全")
    signals.append("IF 已匹配" if impact_factor not in ("", None, "待补充") else "IF 未匹配")
    if citations:
        signals.append(f"引用 {citations}")
    return f"{level}；" + "，".join(signals)


def paper_stable_id(paper: Dict) -> str:
    """Match tracker.stable_paper_id for consistent DB/history keys."""
    for key in ("doi", "arxiv_id", "semantic_scholar_id"):
        value = paper.get(key)
        if value:
            return f"{key}:{str(value).lower()}"
    pid = paper.get("stable_id")
    if pid:
        return str(pid)
    pid = paper.get("id")
    if pid and not str(pid).startswith("scholar_"):
        return str(pid)
    title = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", str(paper.get("title", "")).lower())
    title = re.sub(r"\s+", " ", title).strip()
    if not title:
        title = str(paper.get("link") or paper.get("url") or "unknown")
    import hashlib
    return "title:" + hashlib.sha256(title.encode("utf-8")).hexdigest()[:16]


def normalize_authors(authors) -> str:
    if isinstance(authors, list):
        return json.dumps(authors, ensure_ascii=False)
    return str(authors or "")


def _journal_rank_db_values(paper: Dict) -> tuple:
    impact = paper.get("jcr_impact_factor") or paper.get("impact_factor", "")
    year = paper.get("jcr_year") or paper.get("impact_factor_year", "")
    return (
        paper.get("eissn", ""),
        paper.get("journal_rank_source", ""),
        1 if paper.get("journal_matched") or paper.get("matched") else 0,
        paper.get("journal_match_method", ""),
        impact,
        year,
        paper.get("cas_category", "") or paper.get("category", ""),
        paper.get("cas_top", ""),
        paper.get("cas_warning", ""),
        paper.get("cnki_composite_if", ""),
        paper.get("cnki_comprehensive_if", ""),
        paper.get("core_tags", ""),
        paper.get("ccf_rank", ""),
        as_int(paper.get("journal_quality_score")),
    )


def _resolve_links_for_paper(paper: Dict) -> Dict:
    try:
        from metadata_enricher import resolve_paper_links
        return resolve_paper_links(paper)
    except Exception:
        url = paper.get("link") or paper.get("url") or ""
        return {
            "arxiv_id": paper.get("arxiv_id", ""),
            "arxiv_url": paper.get("arxiv_url", ""),
            "doi_url": paper.get("doi_url", ""),
            "pdf_url": paper.get("pdf_url", ""),
            "scholar_url": paper.get("scholar_url", ""),
            "publisher_url": paper.get("publisher_url", ""),
            "source_url": url,
            "display_url": url,
        }


def upsert_papers(papers: Iterable[Dict], profile_id: str, mode: str,
                  report_path: str = "", pushed_ids: Optional[Iterable[str]] = None,
                  db_path: str = DB_PATH) -> Dict[str, int]:
    init_db(db_path)
    pushed = set(pushed_ids or [])
    now = utc_now()
    inserted = 0
    updated = 0
    skipped_cross_profile = 0
    star_counts = {5: 0, 4: 0, 3: 0, 2: 0, 1: 0}
    with connect(db_path) as conn:
        for paper in papers:
            stable_id = paper_stable_id(paper)
            existing_row = conn.execute(
                "SELECT profile_id FROM papers WHERE stable_id = ?", (stable_id,)
            ).fetchone()
            if existing_row and existing_row[0] and existing_row[0] != profile_id:
                skipped_cross_profile += 1
                continue
            existed = existing_row is not None
            links = _resolve_links_for_paper(paper)
            pushed_to_wechat = 1 if stable_id in pushed or paper.get("pushed_to_wechat") else 0
            publication_year = extract_year(paper.get("published_time") or paper.get("publication_date")) or as_int(paper.get("year"), 0)
            final_score = as_int(paper.get("final_score") or paper.get("relevance_score"))
            citations = as_int(paper.get("citation_count"))
            abstract_complete = bool(paper.get("abstract_is_complete"))
            abstract_status = (
                paper.get("abstract_fetch_status")
                or ("complete" if abstract_complete else "snippet" if paper.get("abstract_source") else "missing")
            )
            system_rating = as_int(paper.get("system_rating")) or rating_from_score(final_score)
            display_url = links.get("display_url") or paper.get("display_url") or paper.get("link") or paper.get("url", "")
            star_counts[min(5, max(1, system_rating))] = star_counts.get(min(5, max(1, system_rating)), 0) + 1
            conn.execute(
                """
                INSERT INTO papers (
                    stable_id, profile_id, title, authors, year, journal, doi, issn, url, source,
                    abstract_original, abstract_zh, abstract_source, abstract_is_complete,
                    impact_factor, impact_factor_year, jcr_quartile, cas_quartile, citation_count,
                    relevance_score, impact_factor_score, freshness_score, citation_score, final_score,
                    recommendation_level, publication_year, system_rating, rating_reason,
                    reading_status, is_favorite, included_in_review, annual_report_year,
                    pushed_to_wechat, report_path, arxiv_id, arxiv_url, doi_url, pdf_url,
                    scholar_url, publisher_url, source_url, display_url, is_recommended,
                    is_milestone, my_notes_path, last_seen_at, ingestion_tier, paper_url,
                    eissn, journal_rank_source, journal_matched, journal_match_method,
                    jcr_impact_factor, jcr_year, cas_category, cas_top, cas_warning,
                    cnki_composite_if, cnki_comprehensive_if, core_tags, ccf_rank,
                    journal_quality_score, abstract_fetch_status, last_run_id, ingest_mode,
                    query_used, filter_reason, is_relevant, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(stable_id) DO UPDATE SET
                    title=excluded.title,
                    authors=excluded.authors,
                    year=excluded.year,
                    journal=excluded.journal,
                    doi=excluded.doi,
                    issn=excluded.issn,
                    url=excluded.url,
                    source=excluded.source,
                    abstract_original=excluded.abstract_original,
                    abstract_zh=CASE WHEN excluded.abstract_zh != '' THEN excluded.abstract_zh ELSE papers.abstract_zh END,
                    abstract_source=excluded.abstract_source,
                    abstract_is_complete=excluded.abstract_is_complete,
                    impact_factor=excluded.impact_factor,
                    impact_factor_year=excluded.impact_factor_year,
                    jcr_quartile=excluded.jcr_quartile,
                    cas_quartile=excluded.cas_quartile,
                    citation_count=excluded.citation_count,
                    relevance_score=excluded.relevance_score,
                    impact_factor_score=excluded.impact_factor_score,
                    freshness_score=excluded.freshness_score,
                    citation_score=excluded.citation_score,
                    final_score=excluded.final_score,
                    recommendation_level=excluded.recommendation_level,
                    publication_year=excluded.publication_year,
                    system_rating=excluded.system_rating,
                    rating_reason=excluded.rating_reason,
                    pushed_to_wechat=MAX(papers.pushed_to_wechat, excluded.pushed_to_wechat),
                    report_path=COALESCE(NULLIF(excluded.report_path, ''), papers.report_path),
                    arxiv_id=excluded.arxiv_id,
                    arxiv_url=excluded.arxiv_url,
                    doi_url=excluded.doi_url,
                    pdf_url=excluded.pdf_url,
                    scholar_url=excluded.scholar_url,
                    publisher_url=excluded.publisher_url,
                    source_url=excluded.source_url,
                    display_url=excluded.display_url,
                    is_recommended=excluded.is_recommended,
                    is_milestone=COALESCE(papers.is_milestone, excluded.is_milestone),
                    last_seen_at=excluded.last_seen_at,
                    ingestion_tier=excluded.ingestion_tier,
                    paper_url=excluded.paper_url,
                    eissn=excluded.eissn,
                    journal_rank_source=excluded.journal_rank_source,
                    journal_matched=excluded.journal_matched,
                    journal_match_method=excluded.journal_match_method,
                    jcr_impact_factor=excluded.jcr_impact_factor,
                    jcr_year=excluded.jcr_year,
                    cas_category=excluded.cas_category,
                    cas_top=excluded.cas_top,
                    cas_warning=excluded.cas_warning,
                    cnki_composite_if=excluded.cnki_composite_if,
                    cnki_comprehensive_if=excluded.cnki_comprehensive_if,
                    core_tags=excluded.core_tags,
                    ccf_rank=excluded.ccf_rank,
                    journal_quality_score=excluded.journal_quality_score,
                    abstract_fetch_status=excluded.abstract_fetch_status,
                    last_run_id=excluded.last_run_id,
                    ingest_mode=excluded.ingest_mode,
                    query_used=excluded.query_used,
                    filter_reason=excluded.filter_reason,
                    is_relevant=excluded.is_relevant,
                    updated_at=excluded.updated_at
                """,
                (
                    stable_id,
                    profile_id,
                    paper.get("title", ""),
                    normalize_authors(paper.get("authors")),
                    extract_year(paper.get("published_time") or paper.get("publication_date")),
                    paper.get("journal_name") or paper.get("journal", ""),
                    paper.get("doi", ""),
                    paper.get("issn", ""),
                    display_url,
                    paper.get("metadata_source") or paper.get("source") or mode,
                    paper.get("abstract_original") or paper.get("abstract", ""),
                    paper.get("摘要中文翻译", "") or paper.get("abstract_zh", ""),
                    paper.get("abstract_source", ""),
                    1 if paper.get("abstract_is_complete") else 0,
                    paper.get("jcr_impact_factor") or paper.get("impact_factor", ""),
                    paper.get("jcr_year") or paper.get("impact_factor_year", ""),
                    paper.get("jcr_quartile", ""),
                    paper.get("cas_quartile", ""),
                    as_int(paper.get("citation_count")),
                    as_int(paper.get("relevance_score")),
                    as_int(paper.get("impact_factor_score")),
                    as_int(paper.get("freshness_score")),
                    as_int(paper.get("citation_score")),
                    final_score,
                    paper.get("recommendation_level", ""),
                    publication_year,
                    system_rating,
                    paper.get("rating_reason") or rating_reason(final_score, abstract_complete, paper.get("impact_factor", ""), citations),
                    paper.get("reading_status", "unread"),
                    1 if paper.get("favorite") or paper.get("is_favorite") else 0,
                    1 if paper.get("included_in_review") else 0,
                    publication_year,
                    pushed_to_wechat,
                    report_path,
                    links.get("arxiv_id", ""),
                    links.get("arxiv_url", ""),
                    links.get("doi_url", ""),
                    links.get("pdf_url", ""),
                    links.get("scholar_url", ""),
                    links.get("publisher_url", ""),
                    links.get("source_url", ""),
                    display_url,
                    1 if paper.get("is_recommended") else 0,
                    1 if paper.get("is_milestone") else 0,
                    paper.get("my_notes_path", ""),
                    now,
                    paper.get("ingestion_tier", "full"),
                    links.get("paper_url") or display_url,
                    *_journal_rank_db_values(paper),
                    abstract_status,
                    as_int(paper.get("last_run_id")),
                    paper.get("ingest_mode") or mode,
                    paper.get("query_used", ""),
                    paper.get("filter_reason", ""),
                    1 if paper.get("is_relevant", 1) else 0,
                    now,
                    now,
                ),
            )
            upsert_reading_note(conn, stable_id, paper)
            if existed:
                updated += 1
            else:
                inserted += 1
    return {
        "ingested": inserted + updated,
        "inserted": inserted,
        "updated": updated,
        "skipped_cross_profile": skipped_cross_profile,
        "star5": star_counts[5],
        "star4": star_counts[4],
        "star3": star_counts[3],
        "star2": star_counts[2],
        "star1": star_counts[1],
    }


def upsert_reading_note(conn: sqlite3.Connection, paper_id: str, paper: Dict) -> None:
    now = utc_now()
    conn.execute(
        """
        INSERT INTO reading_notes (
            paper_id, paper_topic, research_background, research_purpose, core_method,
            main_results, paper_contribution, inspiration, worth_reading, reason,
            possible_ideas, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(paper_id) DO UPDATE SET
            paper_topic=excluded.paper_topic,
            research_background=excluded.research_background,
            research_purpose=excluded.research_purpose,
            core_method=excluded.core_method,
            main_results=excluded.main_results,
            paper_contribution=excluded.paper_contribution,
            inspiration=excluded.inspiration,
            worth_reading=excluded.worth_reading,
            reason=excluded.reason,
            possible_ideas=excluded.possible_ideas,
            updated_at=excluded.updated_at
        """,
        (
            paper_id,
            paper.get("paper_topic") or paper.get("论文主题") or paper.get("title", ""),
            paper.get("research_background") or paper.get("研究背景", ""),
            paper.get("research_purpose") or paper.get("研究目的", ""),
            paper.get("core_method") or paper.get("核心方法", ""),
            paper.get("main_results") or paper.get("实验结果", ""),
            paper.get("paper_contribution") or paper.get("论文创新点", ""),
            paper.get("inspiration") or paper.get("对我的启发", "") or paper.get("未来展望", ""),
            paper.get("worth_reading") or paper.get("recommendation_level", ""),
            paper.get("reason") or paper.get("总结", ""),
            paper.get("possible_ideas") or paper.get("可创新点", ""),
            now,
        ),
    )


def save_reading_note_for_paper(paper_id: str, paper: Dict, db_path: str = DB_PATH) -> None:
    init_db(db_path)
    with connect(db_path) as conn:
        upsert_reading_note(conn, paper_id, paper)


def stamp_paper_lineage(run_id: int, stable_ids: Iterable[str], db_path: str = DB_PATH) -> None:
    ids = [sid for sid in stable_ids if sid]
    if not run_id or not ids:
        return
    init_db(db_path)
    now = utc_now()
    with connect(db_path) as conn:
        for stable_id in ids:
            conn.execute(
                "UPDATE papers SET last_run_id = ?, updated_at = ? WHERE stable_id = ?",
                (run_id, now, stable_id),
            )


def record_run(profile: str, mode: str, stats: Dict, report_path: str = "",
               pushed_count: int = 0, db_path: str = DB_PATH) -> int:
    init_db(db_path)
    with connect(db_path) as conn:
        cursor = conn.execute(
            """
            INSERT INTO runs (
                profile, mode, total_found, kept_after_relevance, new_papers,
                abstract_completed, if_matched, recommended_count, pushed_count,
                report_path, run_time, ingested_count, updated_count, skipped_duplicate,
                star5, star4, star3, star2, star1, skipped_low_score,
                run_year, max_results, data_sources, google_query, ingest_policy,
                skipped_irrelevant, doi_completed, skipped_cross_profile, filter_stats_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                profile,
                mode,
                as_int(stats.get("retrieved")),
                as_int(stats.get("kept_after_relevance")),
                as_int(stats.get("new_papers")),
                as_int(stats.get("abstract_completed")),
                as_int(stats.get("if_matched")),
                as_int(stats.get("push_count") or stats.get("recommended_count")),
                pushed_count or as_int(stats.get("push_count")),
                report_path,
                utc_now(),
                as_int(stats.get("ingested_count")),
                as_int(stats.get("updated_count")),
                as_int(stats.get("skipped_duplicate") or stats.get("already_seen")),
                as_int(stats.get("star5")),
                as_int(stats.get("star4")),
                as_int(stats.get("star3")),
                as_int(stats.get("star2")),
                as_int(stats.get("star1")),
                as_int(stats.get("skipped_low_score")),
                str(stats.get("run_year") or ""),
                as_int(stats.get("max_results")),
                str(stats.get("data_sources") or ""),
                str(stats.get("google_query") or ""),
                str(stats.get("ingest_policy") or ""),
                as_int(stats.get("skipped_irrelevant")),
                as_int(stats.get("doi_completed")),
                as_int(stats.get("skipped_cross_profile")),
                stats.get("filter_stats_json") if isinstance(stats.get("filter_stats_json"), str)
                else json.dumps(stats.get("filter_stats_json") or {}, ensure_ascii=False),
            ),
        )
        return int(cursor.lastrowid)


def sync_profiles(profiles_doc: Optional[Dict] = None, db_path: str = DB_PATH) -> None:
    init_db(db_path)
    if profiles_doc is None:
        with open(PROFILES_PATH, "r", encoding="utf-8") as f:
            profiles_doc = json.load(f)
    now = utc_now()
    with connect(db_path) as conn:
        for profile_id, profile in profiles_doc.get("profiles", {}).items():
            conn.execute(
                """
                INSERT INTO profiles (profile_id, name, display_name, description, config_json, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(profile_id) DO UPDATE SET
                    name=excluded.name,
                    display_name=excluded.display_name,
                    description=excluded.description,
                    config_json=excluded.config_json,
                    updated_at=excluded.updated_at
                """,
                (
                    profile_id,
                    profile.get("name", profile_id),
                    profile.get("display_name", profile.get("name", profile_id)),
                    profile.get("description", ""),
                    json.dumps(profile, ensure_ascii=False),
                    now,
                ),
            )


def latest_run(db_path: str = DB_PATH) -> Optional[Dict]:
    init_db(db_path)
    with connect(db_path) as conn:
        row = conn.execute("SELECT * FROM runs ORDER BY run_time DESC, id DESC LIMIT 1").fetchone()
        return dict(row) if row else None


RATING_SQL = """
COALESCE(
    user_rating,
    system_rating,
    CASE
        WHEN COALESCE(final_score, relevance_score, 0) >= 24 THEN 5
        WHEN COALESCE(final_score, relevance_score, 0) >= 18 THEN 4
        WHEN COALESCE(final_score, relevance_score, 0) >= 12 THEN 3
        WHEN COALESCE(final_score, relevance_score, 0) >= 6 THEN 2
        ELSE 1
    END
)
"""


def top_papers(limit: int = 3, db_path: str = DB_PATH) -> List[Dict]:
    init_db(db_path)
    with connect(db_path) as conn:
        result = conn.execute(
            f"""
            SELECT *, {RATING_SQL.strip()} AS display_rating
            FROM papers
            ORDER BY is_recommended DESC, display_rating DESC, final_score DESC, updated_at DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in result]


def get_run(run_id: int, db_path: str = DB_PATH) -> Optional[Dict]:
    init_db(db_path)
    with connect(db_path) as conn:
        row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        return dict(row) if row else None


def record_run_candidates(run_id: int, candidates: List[Dict], db_path: str = DB_PATH) -> None:
    if not run_id or not candidates:
        return
    init_db(db_path)
    with connect(db_path) as conn:
        conn.execute("DELETE FROM run_candidates WHERE run_id = ?", (run_id,))
        conn.executemany(
            """
            INSERT INTO run_candidates (
                run_id, stable_id, title, year, journal, url,
                relevance_score, status, filter_reason, sort_order
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    run_id,
                    item.get("stable_id") or item.get("id") or "",
                    item.get("title") or "",
                    as_int(item.get("year")),
                    item.get("journal") or "",
                    item.get("url") or item.get("paper_url") or "",
                    as_int(item.get("relevance_score")),
                    item.get("status") or "filtered",
                    item.get("filter_reason") or "",
                    idx,
                )
                for idx, item in enumerate(candidates)
            ],
        )


def get_run_candidates(run_id: int, db_path: str = DB_PATH) -> List[Dict]:
    init_db(db_path)
    with connect(db_path) as conn:
        rows = conn.execute(
            """
            SELECT stable_id, title, year, journal, url, relevance_score, status, filter_reason
            FROM run_candidates
            WHERE run_id = ?
            ORDER BY sort_order, id
            """,
            (run_id,),
        ).fetchall()
        return [dict(row) for row in rows]


def recommendation_level_from_score(score: int) -> str:
    if score >= 20:
        return "A+ 必读"
    if score >= 15:
        return "A 推荐精读"
    if score >= 10:
        return "B 值得关注"
    return "C 仅归档"


def rematch_journal_ranks(metrics_path: str = "", db_path: str = DB_PATH) -> Dict[str, int]:
    """Re-apply journal_metrics.csv to all papers in SQLite."""
    from journal_rank_enhancer import apply_journal_rank_to_paper, load_journal_metrics

    init_db(db_path)
    metrics = load_journal_metrics(metrics_path or None, quiet=True)
    matched = 0
    total = 0
    with connect(db_path) as conn:
        paper_rows = conn.execute("SELECT * FROM papers").fetchall()
        for row in paper_rows:
            total += 1
            paper = dict(row)
            paper["relevance_score"] = as_int(paper.get("relevance_score"))
            paper["freshness_score"] = as_int(paper.get("freshness_score"))
            paper["citation_score"] = as_int(paper.get("citation_score"))
            apply_journal_rank_to_paper(paper, metrics)
            if paper.get("journal_matched"):
                matched += 1
            final_score = as_int(paper.get("final_score"))
            system_rating = rating_from_score(final_score)
            conn.execute(
                """
                UPDATE papers SET
                    impact_factor=?, impact_factor_year=?, jcr_quartile=?, cas_quartile=?,
                    eissn=?, journal_rank_source=?, journal_matched=?, journal_match_method=?,
                    jcr_impact_factor=?, jcr_year=?, cas_category=?, cas_top=?, cas_warning=?,
                    cnki_composite_if=?, cnki_comprehensive_if=?, core_tags=?, ccf_rank=?,
                    journal_quality_score=?, impact_factor_score=?, final_score=?,
                    system_rating=?, recommendation_level=?, updated_at=datetime('now')
                WHERE stable_id=?
                """,
                (
                    paper.get("jcr_impact_factor") or paper.get("impact_factor", ""),
                    paper.get("jcr_year") or paper.get("impact_factor_year", ""),
                    paper.get("jcr_quartile", ""),
                    paper.get("cas_quartile", ""),
                    paper.get("eissn", ""),
                    paper.get("journal_rank_source", ""),
                    1 if paper.get("journal_matched") else 0,
                    paper.get("journal_match_method", ""),
                    paper.get("jcr_impact_factor", ""),
                    paper.get("jcr_year", ""),
                    paper.get("cas_category", ""),
                    paper.get("cas_top", ""),
                    paper.get("cas_warning", ""),
                    paper.get("cnki_composite_if", ""),
                    paper.get("cnki_comprehensive_if", ""),
                    paper.get("core_tags", ""),
                    paper.get("ccf_rank", ""),
                    as_int(paper.get("journal_quality_score")),
                    as_int(paper.get("impact_factor_score")),
                    final_score,
                    system_rating,
                    recommendation_level_from_score(final_score),
                    paper["stable_id"],
                ),
            )
        conn.commit()
    return {"total": total, "matched": matched, "unmatched": total - matched}
