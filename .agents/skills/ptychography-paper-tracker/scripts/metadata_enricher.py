import html
import re
import time
from typing import Dict, List, Optional
from urllib.parse import quote_plus

try:
    import feedparser
except ImportError:
    feedparser = None

try:
    import requests
except ImportError:
    requests = None


USER_AGENT = "Ptychography-Paper-Tracker/1.0 (metadata enrichment)"


def clean_text(value: str) -> str:
    value = html.unescape(str(value or ""))
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_title(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def title_similarity(left: str, right: str) -> float:
    left_tokens = set(normalize_title(left).split())
    right_tokens = set(normalize_title(right).split())
    if not left_tokens or not right_tokens:
        return 0.0
    return len(left_tokens & right_tokens) / max(len(left_tokens), len(right_tokens))


def first_external_id(external_ids: Dict, key: str) -> str:
    value = (external_ids or {}).get(key, "")
    if isinstance(value, list):
        return str(value[0]) if value else ""
    return str(value or "")


def search_semantic_scholar(paper: Dict, api_key: str = "") -> Dict:
    if requests is None:
        return {}
    title = paper.get("title", "")
    if not title:
        return {}
    headers = {"User-Agent": USER_AGENT}
    if api_key:
        headers["x-api-key"] = api_key
    params = {
        "query": title,
        "limit": 3,
        "fields": "title,abstract,venue,publicationDate,year,citationCount,externalIds,journal,url,authors",
    }
    try:
        response = requests.get(
            "https://api.semanticscholar.org/graph/v1/paper/search",
            params=params,
            headers=headers,
            timeout=20,
        )
        response.raise_for_status()
        candidates = response.json().get("data", [])
    except Exception:
        return {}

    best = None
    best_score = 0.0
    for candidate in candidates:
        score = title_similarity(title, candidate.get("title", ""))
        if score > best_score:
            best = candidate
            best_score = score
    if not best or best_score < 0.55:
        return {}

    external_ids = best.get("externalIds", {}) or {}
    journal = best.get("journal") or {}
    return {
        "abstract_original": clean_text(best.get("abstract", "")),
        "abstract_source": "Semantic Scholar" if best.get("abstract") else "",
        "doi": first_external_id(external_ids, "DOI"),
        "issn": first_external_id(external_ids, "ISSN"),
        "arxiv_id": first_external_id(external_ids, "ArXiv"),
        "journal_name": journal.get("name") or best.get("venue", ""),
        "publication_date": best.get("publicationDate") or str(best.get("year", "") or ""),
        "citation_count": best.get("citationCount", ""),
        "metadata_source": "Semantic Scholar",
        "url": best.get("url", ""),
    }


def reconstruct_openalex_abstract(index: Dict) -> str:
    if not index:
        return ""
    positions = []
    for word, word_positions in index.items():
        for pos in word_positions:
            positions.append((pos, word))
    return clean_text(" ".join(word for _, word in sorted(positions)))


def search_openalex(paper: Dict) -> Dict:
    if requests is None:
        return {}
    title = paper.get("title", "")
    if not title:
        return {}
    params = {
        "search": title,
        "per-page": 3,
        "select": "id,title,doi,publication_year,publication_date,authorships,primary_location,abstract_inverted_index,cited_by_count",
    }
    try:
        response = requests.get("https://api.openalex.org/works", params=params, timeout=20)
        response.raise_for_status()
        candidates = response.json().get("results", [])
    except Exception:
        return {}

    best = None
    best_score = 0.0
    for candidate in candidates:
        score = title_similarity(title, candidate.get("title", ""))
        if score > best_score:
            best = candidate
            best_score = score
    if not best or best_score < 0.55:
        return {}

    source = ((best.get("primary_location") or {}).get("source") or {})
    issn_list = source.get("issn") or []
    abstract = reconstruct_openalex_abstract(best.get("abstract_inverted_index") or {})
    return {
        "abstract_original": abstract,
        "abstract_source": "OpenAlex" if abstract else "",
        "doi": str(best.get("doi", "") or "").replace("https://doi.org/", ""),
        "issn": issn_list[0] if issn_list else "",
        "journal_name": source.get("display_name", ""),
        "publication_date": best.get("publication_date") or str(best.get("publication_year", "") or ""),
        "citation_count": best.get("cited_by_count", ""),
        "metadata_source": "OpenAlex",
        "url": best.get("id", ""),
    }


def search_arxiv(paper: Dict) -> Dict:
    if requests is None or feedparser is None:
        return {}
    title = paper.get("title", "")
    if not title:
        return {}
    query = f"ti:{quote_plus(title)}"
    try:
        response = requests.get(
            "http://export.arxiv.org/api/query",
            params={"search_query": query, "start": 0, "max_results": 3},
            timeout=20,
        )
        response.raise_for_status()
        feed = feedparser.parse(response.text)
    except Exception:
        return {}

    best = None
    best_score = 0.0
    for entry in feed.entries:
        score = title_similarity(title, entry.get("title", ""))
        if score > best_score:
            best = entry
            best_score = score
    if not best or best_score < 0.55:
        return {}
    return {
        "abstract_original": clean_text(best.get("summary", "")),
        "abstract_source": "arXiv" if best.get("summary") else "",
        "doi": str(best.get("arxiv_doi", "") or ""),
        "issn": "",
        "arxiv_id": str(best.get("id", "")).rsplit("/", 1)[-1],
        "journal_name": "arXiv",
        "publication_date": best.get("published", ""),
        "citation_count": "",
        "metadata_source": "arXiv",
        "url": best.get("link", ""),
    }


def search_crossref(paper: Dict) -> Dict:
    if requests is None:
        return {}
    title = paper.get("title", "")
    if not title:
        return {}
    try:
        response = requests.get(
            "https://api.crossref.org/works",
            params={"query.title": title, "rows": 3},
            headers={"User-Agent": USER_AGENT},
            timeout=20,
        )
        response.raise_for_status()
        candidates = response.json().get("message", {}).get("items", [])
    except Exception:
        return {}

    best = None
    best_score = 0.0
    for candidate in candidates:
        candidate_title = " ".join(candidate.get("title", [])[:1])
        score = title_similarity(title, candidate_title)
        if score > best_score:
            best = candidate
            best_score = score
    if not best or best_score < 0.55:
        return {}

    issn_list = best.get("ISSN", []) or []
    published = best.get("published-print") or best.get("published-online") or best.get("created") or {}
    date_parts = published.get("date-parts", [[]])[0]
    return {
        "abstract_original": clean_text(best.get("abstract", "")),
        "abstract_source": "Crossref" if best.get("abstract") else "",
        "doi": best.get("DOI", ""),
        "issn": issn_list[0] if issn_list else "",
        "journal_name": " ".join(best.get("container-title", [])[:1]),
        "publication_date": "-".join(str(part) for part in date_parts) if date_parts else "",
        "citation_count": "",
        "metadata_source": "Crossref",
        "url": best.get("URL", ""),
    }


def merge_metadata(base: Dict, update: Dict) -> Dict:
    merged = dict(base)
    for key, value in update.items():
        if value not in ("", None, [], {}):
            merged[key] = value
    return merged


def enrich_paper_metadata(paper: Dict, semantic_scholar_key: str = "") -> Dict:
    enriched = {
        **paper,
        "abstract_original": paper.get("abstract", ""),
        "abstract_source": "Google Scholar snippet" if paper.get("abstract") else "missing",
        "abstract_is_complete": False,
        "doi": paper.get("doi", ""),
        "issn": paper.get("issn", ""),
        "journal_name": paper.get("journal", ""),
        "citation_count": paper.get("citation_count", ""),
        "metadata_source": "Google Scholar",
    }

    for getter in (
        lambda p: search_semantic_scholar(p, semantic_scholar_key),
        search_openalex,
        search_arxiv,
        search_crossref,
    ):
        update = getter(enriched)
        if not update:
            continue
        enriched = merge_metadata(enriched, update)
        if update.get("abstract_original"):
            enriched["abstract_is_complete"] = True
            enriched["abstract_source"] = update.get("abstract_source", enriched["abstract_source"])
            break
        time.sleep(0.2)

    if not enriched.get("abstract_original"):
        enriched["abstract_original"] = paper.get("abstract", "")
        enriched["abstract_source"] = "Google Scholar snippet" if paper.get("abstract") else "missing"
        enriched["abstract_is_complete"] = False

    if enriched.get("abstract_is_complete"):
        enriched["abstract_fetch_status"] = "complete"
    elif enriched.get("abstract_original"):
        src = str(enriched.get("abstract_source") or "")
        enriched["abstract_fetch_status"] = "snippet" if "snippet" in src.lower() else "complete"
    else:
        enriched["abstract_fetch_status"] = "missing"
    if enriched.get("abstract_original"):
        enriched["abstract"] = enriched["abstract_original"]
    if enriched.get("journal_name"):
        enriched["journal"] = enriched["journal_name"]
    if enriched.get("publication_date"):
        enriched["published_time"] = enriched["publication_date"]
    try:
        from abstract_translate import ensure_paper_chinese_abstract
        enriched = ensure_paper_chinese_abstract(enriched)
    except Exception:
        pass
    return {**enriched, **resolve_paper_links(enriched)}


def resolve_paper_links(paper: Dict) -> Dict:
    """Derive canonical URLs for web display and export."""
    url = str(paper.get("link") or paper.get("url") or "").strip()
    doi = str(paper.get("doi") or "").strip()
    arxiv_id = str(paper.get("arxiv_id") or "").strip()
    if not arxiv_id and "arxiv.org/abs/" in url:
        arxiv_id = url.split("arxiv.org/abs/")[-1].split("v")[0].strip("/")
    if not arxiv_id and url.startswith("http") and "arxiv" in url:
        arxiv_id = url.rstrip("/").split("/")[-1].split("v")[0]

    arxiv_url = f"https://arxiv.org/abs/{arxiv_id}" if arxiv_id else ""
    pdf_url = f"https://arxiv.org/pdf/{arxiv_id}.pdf" if arxiv_id else ""
    doi_url = f"https://doi.org/{doi.lstrip('https://doi.org/').lstrip('http://doi.org/')}" if doi else ""
    scholar_url = ""
    title = paper.get("title", "")
    if title:
        scholar_url = "https://scholar.google.com/scholar?q=" + quote_plus(title)

    publisher_url = url if url and "arxiv.org" not in url else ""
    source_url = url or arxiv_url or doi_url
    paper_url = url or publisher_url or arxiv_url or doi_url or scholar_url or source_url
    display_url = paper_url or publisher_url or arxiv_url or doi_url or scholar_url or source_url

    return {
        "arxiv_id": arxiv_id,
        "arxiv_url": arxiv_url,
        "doi_url": doi_url,
        "pdf_url": pdf_url,
        "scholar_url": scholar_url,
        "publisher_url": publisher_url,
        "source_url": source_url,
        "paper_url": paper_url,
        "display_url": display_url,
        "url": display_url,
        "link": display_url,
    }


def enrich_papers_metadata(papers: List[Dict], semantic_scholar_key: str = "") -> List[Dict]:
    enriched = []
    for idx, paper in enumerate(papers, 1):
        print(f"补全元数据 {idx}/{len(papers)}：{paper.get('title', '无标题')}")
        enriched.append(enrich_paper_metadata(paper, semantic_scholar_key))
        time.sleep(0.4)
    return enriched
