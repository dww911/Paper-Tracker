"""Idea Lab: local RAG + multi-source external search + structured idea polishing."""

from __future__ import annotations

import os
import re
from difflib import SequenceMatcher
from typing import Optional
from urllib.parse import quote_plus

try:
    import requests
except ImportError:
    requests = None

try:
    import feedparser
except ImportError:
    feedparser = None

from rag_indexer import format_context as format_rag_context, search as rag_search
from prompt_templates import render_template

ROADMAP_BRANCH_DEFS = [
    ("branch_phase_retrieval", "Phase retrieval 基础", "相位恢复、成像反演和算法基础。", "phase retrieval,ptychography,coherent diffraction"),
    ("branch_pie_epie", "PIE / ePIE", "核心迭代算法、探针恢复和重建稳定性。", "PIE,ePIE,iterative,probe"),
    ("branch_electron", "Electron Ptychography", "电子显微 ptychography 的主干方向。", "electron ptychography,electron microscopy,TEM,STEM"),
    ("branch_4dstem", "4D-STEM acquisition", "4D-STEM 数据采集、扫描策略和探测器数据。", "4D-STEM,4D STEM,detector,scan"),
    ("branch_low_dose", "Low-dose reconstruction", "低剂量、噪声鲁棒性和剂量效率。", "low dose,low-dose,dose,noise"),
    ("branch_multislice", "Multislice / thick samples", "多切片、厚样品和多重散射建模。", "multislice,thick,multiple scattering"),
    ("branch_wdd", "WDD / initialization", "WDD、SSB、初始化和直接方法。", "WDD,SSB,single sideband,initialization,direct"),
    ("branch_applications", "Materials / semiconductor applications", "材料、半导体、缺陷和应变等应用。", "semiconductor,materials,strain,defect,device"),
    ("branch_ai", "AI / automation", "深度学习、自动化重建和智能分析。", "AI,deep learning,machine learning,automation,neural"),
]

SECTION_HEADERS = [
    "1. 初步判断",
    "2. 更清晰的研究问题",
    "3. 已有工作基础",
    "4. 可能创新点",
    "5. 风险与重复性判断",
    "6. 推荐阅读论文",
    "7. 下一步验证方案",
]

USER_AGENT = "ResearchRadar/1.0 (mailto:research@local)"


def branch_meta(branch_key: str) -> dict:
    for key, title, desc, keywords in ROADMAP_BRANCH_DEFS:
        if key == branch_key:
            return {"node_key": key, "title": title, "description": desc, "keywords": keywords}
    return {"node_key": branch_key or "", "title": branch_key or "全方向", "description": "", "keywords": ""}


def title_similarity(a: str, b: str) -> float:
    a = re.sub(r"\s+", " ", (a or "").lower().strip())
    b = re.sub(r"\s+", " ", (b or "").lower().strip())
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def normalize_hit(item: dict, source_type: str) -> dict:
    return {
        "title": item.get("title") or "无标题",
        "year": item.get("year") or item.get("publication_year") or "",
        "journal": item.get("journal") or item.get("venue") or item.get("journal_name") or "",
        "doi": str(item.get("doi") or "").replace("https://doi.org/", ""),
        "url": item.get("url") or item.get("link") or item.get("display_url") or "",
        "abstract": (item.get("abstract") or item.get("abstract_original") or "")[:500],
        "citation_count": item.get("citation_count") or 0,
        "source_type": source_type,
        "relevance_reason": item.get("relevance_reason") or f"来自 {source_type}",
        "stable_id": item.get("stable_id") or "",
        "display_rating": item.get("display_rating") or item.get("system_rating") or "",
    }


def dedupe_hits(hits: list[dict], limit: int = 20) -> list[dict]:
    seen: list[dict] = []
    for hit in hits:
        title = hit.get("title") or ""
        doi = (hit.get("doi") or "").lower()
        duplicate = False
        for existing in seen:
            if doi and doi == (existing.get("doi") or "").lower():
                duplicate = True
                break
            if title_similarity(title, existing.get("title") or "") >= 0.88:
                duplicate = True
                break
        if not duplicate:
            seen.append(hit)
        if len(seen) >= limit:
            break
    return seen


def search_openalex_keywords(query: str, limit: int = 8) -> list[dict]:
    if requests is None or not query.strip():
        return []
    try:
        response = requests.get(
            "https://api.openalex.org/works",
            params={
                "search": query[:300],
                "per-page": min(limit, 25),
                "select": "id,title,doi,publication_year,primary_location,cited_by_count,abstract_inverted_index",
            },
            timeout=25,
        )
        if response.status_code != 200:
            return []
        results = []
        for item in response.json().get("results", [])[:limit]:
            source = ((item.get("primary_location") or {}).get("source") or {})
            abstract = ""
            inv = item.get("abstract_inverted_index") or {}
            if inv:
                positions = sorted((pos, w) for w, ps in inv.items() for pos in ps)
                abstract = " ".join(w for _, w in positions)[:600]
            results.append(
                normalize_hit(
                    {
                        "title": item.get("title"),
                        "year": item.get("publication_year"),
                        "journal": source.get("display_name"),
                        "doi": item.get("doi"),
                        "url": item.get("id"),
                        "abstract": abstract,
                        "citation_count": item.get("cited_by_count"),
                        "relevance_reason": "OpenAlex 关键词匹配",
                    },
                    "openalex",
                )
            )
        return results
    except Exception:
        return []


def search_crossref_keywords(query: str, limit: int = 8) -> list[dict]:
    if requests is None or not query.strip():
        return []
    try:
        response = requests.get(
            "https://api.crossref.org/works",
            params={"query": query[:200], "rows": min(limit, 20)},
            headers={"User-Agent": USER_AGENT},
            timeout=25,
        )
        if response.status_code != 200:
            return []
        results = []
        for item in response.json().get("message", {}).get("items", [])[:limit]:
            title = (item.get("title") or [""])[0]
            journal = ((item.get("container-title") or [""])[0]) if item.get("container-title") else ""
            year = (item.get("published-print") or item.get("published-online") or {}).get("date-parts", [[None]])[0][0]
            results.append(
                normalize_hit(
                    {
                        "title": title,
                        "year": year,
                        "journal": journal,
                        "doi": item.get("DOI"),
                        "url": item.get("URL") or (f"https://doi.org/{item['DOI']}" if item.get("DOI") else ""),
                        "citation_count": item.get("is-referenced-by-count"),
                        "relevance_reason": "Crossref 关键词匹配",
                    },
                    "crossref",
                )
            )
        return results
    except Exception:
        return []


def search_arxiv_keywords(query: str, limit: int = 8) -> list[dict]:
    if requests is None or feedparser is None or not query.strip():
        return []
    try:
        q = f"all:{quote_plus(query[:120])}"
        response = requests.get(
            "http://export.arxiv.org/api/query",
            params={"search_query": q, "start": 0, "max_results": min(limit, 15)},
            timeout=25,
        )
        response.raise_for_status()
        feed = feedparser.parse(response.text)
        results = []
        for entry in feed.entries[:limit]:
            results.append(
                normalize_hit(
                    {
                        "title": entry.get("title"),
                        "year": (entry.get("published") or "")[:4],
                        "journal": "arXiv",
                        "url": entry.get("link"),
                        "abstract": entry.get("summary"),
                        "relevance_reason": "arXiv 关键词匹配",
                    },
                    "arxiv",
                )
            )
        return results
    except Exception:
        return []


def search_semantic_scholar_keywords(query: str, limit: int = 8) -> list[dict]:
    if requests is None or not query.strip():
        return []
    api_key = os.environ.get("SEMANTIC_SCHOLAR_API_KEY", "")
    headers = {"x-api-key": api_key} if api_key else {}
    try:
        response = requests.get(
            "https://api.semanticscholar.org/graph/v1/paper/search",
            headers=headers,
            params={
                "query": query[:200],
                "limit": min(limit, 20),
                "fields": "title,year,venue,url,abstract,citationCount,externalIds",
            },
            timeout=25,
        )
        if response.status_code != 200:
            return []
        results = []
        for item in response.json().get("data", [])[:limit]:
            results.append(
                normalize_hit(
                    {
                        "title": item.get("title"),
                        "year": item.get("year"),
                        "journal": item.get("venue"),
                        "url": item.get("url"),
                        "abstract": item.get("abstract"),
                        "citation_count": item.get("citationCount"),
                        "doi": (item.get("externalIds") or {}).get("DOI"),
                        "relevance_reason": "Semantic Scholar 关键词匹配",
                    },
                    "semantic_scholar",
                )
            )
        return results
    except Exception:
        return []


def search_local_evidence(
    query: str,
    profile_id: str,
    branch_key: str = "",
    limit: int = 10,
    paper_id: str = "",
) -> tuple[list[dict], list[dict]]:
    chunks = rag_search(query, profile=profile_id, paper_id=paper_id, limit=limit)
    hits = []
    for chunk in chunks:
        hits.append(
            normalize_hit(
                {
                    "title": chunk.get("title") or "本地片段",
                    "year": "",
                    "journal": chunk.get("source_type") or "",
                    "url": chunk.get("source_url") or "",
                    "abstract": chunk.get("chunk_text") or "",
                    "stable_id": chunk.get("paper_id") or "",
                    "relevance_reason": "本地文献库 / 笔记 RAG",
                },
                "local",
            )
        )
    if branch_key and profile_id:
        from radar_db import connect, init_db

        init_db()
        with connect() as conn:
            rows = conn.execute(
                """
                SELECT p.* FROM roadmap_papers rp
                JOIN papers p ON p.stable_id = rp.paper_stable_id
                WHERE rp.profile = ? AND rp.node_key = ?
                ORDER BY rp.reading_rank LIMIT ?
                """,
                (profile_id, branch_key, limit),
            ).fetchall()
        for row in rows:
            paper = dict(row)
            hits.append(
                normalize_hit(
                    {
                        **paper,
                        "relevance_reason": f"当前分支「{branch_meta(branch_key).get('title')}」关联论文",
                    },
                    "local",
                )
            )
    return dedupe_hits(hits, limit), chunks


def search_external_evidence(
    query: str,
    sources: Optional[list[str]] = None,
    limit_per_source: int = 6,
) -> list[dict]:
    sources = sources or ["openalex", "crossref", "arxiv", "semantic_scholar"]
    hits: list[dict] = []
    if "openalex" in sources:
        hits.extend(search_openalex_keywords(query, limit_per_source))
    if "crossref" in sources:
        hits.extend(search_crossref_keywords(query, limit_per_source))
    if "arxiv" in sources:
        hits.extend(search_arxiv_keywords(query, limit_per_source))
    if "semantic_scholar" in sources:
        hits.extend(search_semantic_scholar_keywords(query, limit_per_source))
    return dedupe_hits(hits, limit_per_source * 2)


def parse_structured_sections(markdown: str) -> dict:
    sections = {h: "" for h in SECTION_HEADERS}
    if not markdown:
        return sections
    current = None
    for line in markdown.splitlines():
        stripped = line.strip()
        matched = None
        for header in SECTION_HEADERS:
            if stripped.startswith(f"## {header}") or stripped.startswith(f"##{header}"):
                matched = header
                break
        if matched:
            current = matched
            continue
        if current:
            sections[current] += line + "\n"
    for key in sections:
        sections[key] = sections[key].strip()
    return sections


def extract_fetch_keywords(analysis_markdown: str, raw_idea: str = "") -> list[str]:
    keywords: list[str] = []
    for line in (analysis_markdown or "").splitlines():
        if "补充检索关键词" in line:
            part = line.split("：", 1)[-1].split(":", 1)[-1]
            keywords.extend([k.strip() for k in re.split(r"[,，;；]", part) if k.strip()])
    if not keywords and raw_idea:
        tokens = re.findall(r"[A-Za-z][A-Za-z0-9\-+/ ]{2,}|[\u4e00-\u9fff]{2,8}", raw_idea)
        keywords = [t.strip() for t in tokens[:6] if len(t.strip()) >= 3]
    seen = set()
    out = []
    for kw in keywords:
        low = kw.lower()
        if low not in seen:
            seen.add(low)
            out.append(kw)
    return out[:8]


def polish_idea_structured(
    raw_idea: str,
    profile: dict,
    branch_key: str,
    local_hits: list[dict],
    external_hits: list[dict],
    call_llm_fn,
    ai_configured: bool,
    rag_chunks: Optional[list[dict]] = None,
) -> dict:
    local_context = format_rag_context(rag_chunks or [])
    if not local_context.strip() and local_hits:
        local_context = "\n".join(
            f"- {h.get('title')} ({h.get('year') or 'n.d.'}): {(h.get('abstract') or '')[:200]}"
            for h in local_hits[:8]
        )
    external_context = "\n".join(
        f"- {h.get('title')} ({h.get('year') or 'n.d.'}) [{h.get('source_type')}]: {(h.get('abstract') or '')[:200]}"
        for h in external_hits[:10]
    ) or "（未启用或未检索到公开学术源结果）"
    branch = branch_meta(branch_key)
    fallback = f"""请打磨以下科研想法。

研究方向：{profile.get('display_name') or profile.get('name') or ''}
分支：{branch.get('title')}
想法：{raw_idea}

本地依据：
{local_context}

公开源：
{external_context}
"""
    prompt = render_template(
        "idea_polish_prompt",
        {
            "profile_name": profile.get("display_name") or profile.get("name") or "",
            "branch_name": branch.get("title") or "全方向",
            "raw_idea": raw_idea,
            "local_context": local_context,
            "external_context": external_context,
        },
        fallback,
    )
    raw_md = call_llm_fn(prompt) if ai_configured else ""
    if not raw_md:
        raw_md = (
            f"## 1. 初步判断\n想法「{raw_idea[:120]}」值得进一步检索验证；当前本地依据 {len(local_hits)} 条，"
            f"公开源 {len(external_hits)} 条。\n\n"
            "## 2. 更清晰的研究问题\n（请配置 AI Key 后重新打磨，或手动整理。）\n\n"
            "## 7. 下一步验证方案\n补充检索关键词：" + ", ".join(extract_fetch_keywords("", raw_idea))
        )
    sections = parse_structured_sections(raw_md)
    return {
        "raw_markdown": raw_md,
        "sections": sections,
        "fetch_keywords": extract_fetch_keywords(raw_md, raw_idea),
    }


def group_evidence(local_hits: list[dict], external_hits: list[dict]) -> dict:
    supporting = [h for h in local_hits if h.get("source_type") == "local"]
    external = external_hits or []
    duplicate_risk = external[:2]
    must_read = supporting[:4]
    gaps = []
    if len(external) < 3:
        gaps.append("公开学术源结果偏少，建议生成补充检索任务")
    if len(supporting) < 2:
        gaps.append("本地分支论文不足，建议先获取更多文献")
    return {
        "supporting": supporting,
        "duplicate_risk": duplicate_risk,
        "must_read": must_read,
        "gaps": gaps,
        "external": external,
    }
