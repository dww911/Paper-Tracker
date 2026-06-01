"""Journal Rank Enhancer — match papers to local journal_metrics.csv."""

from __future__ import annotations

import csv
import os
import re
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple

METRIC_FIELDS = [
    "journal_name",
    "issn",
    "eissn",
    "jcr_impact_factor",
    "jcr_year",
    "jcr_quartile",
    "cas_quartile",
    "cas_category",
    "cas_top",
    "cas_warning",
    "cnki_composite_if",
    "cnki_comprehensive_if",
    "core_tags",
    "ccf_rank",
    "source",
    # legacy aliases
    "impact_factor",
    "impact_factor_year",
    "category",
]


def metrics_path() -> str:
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), "journal_metrics.csv")


def example_metrics_path() -> str:
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), "journal_metrics.example.csv")


def normalize_issn(value: str) -> str:
    return re.sub(r"[^0-9Xx]", "", str(value or "")).upper()


def normalize_journal_name(value: str) -> str:
    value = str(value or "").strip().lower()
    value = value.replace("：", ":").replace("，", ",")
    value = re.sub(r"&", " and ", value)
    value = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", value)
    stopwords = {"the", "journal", "of", "and"}
    tokens = [t for t in value.split() if t and t not in stopwords]
    return " ".join(tokens).strip()


def ensure_journal_metrics_file(path: Optional[str] = None) -> str:
    """Create journal_metrics.csv from example template when missing."""
    import shutil

    path = path or metrics_path()
    if os.path.exists(path):
        return path
    example = example_metrics_path()
    if os.path.exists(example):
        shutil.copyfile(example, path)
    return path


def load_journal_metrics(path: Optional[str] = None, quiet: bool = False) -> List[Dict]:
    path = path or metrics_path()
    if not os.path.exists(path):
        ensure_journal_metrics_file(path)
    if not os.path.exists(path):
        if not quiet:
            print(
                "⚠️ 未找到 journal_metrics.csv，已跳过期刊等级匹配。"
                f"请在设置中心上传期刊指标表。样例：{example_metrics_path()}"
            )
        return []
    if not quiet and path == metrics_path():
        print(f"📊 已加载期刊指标表：{path}")

    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows: List[Dict] = []
        for raw in reader:
            row = {field: str(raw.get(field, "") or "").strip() for field in METRIC_FIELDS}
            # legacy column names
            if not row["jcr_impact_factor"] and row.get("impact_factor"):
                row["jcr_impact_factor"] = row["impact_factor"]
            if not row["jcr_year"] and row.get("impact_factor_year"):
                row["jcr_year"] = row["impact_factor_year"]
            if not row["cas_category"] and row.get("category"):
                row["cas_category"] = row["category"]
            row["_name_key"] = normalize_journal_name(row["journal_name"])
            row["_issn_key"] = normalize_issn(row["issn"])
            row["_eissn_key"] = normalize_issn(row["eissn"])
            rows.append(row)
        return rows


def _empty_rank() -> Dict:
    return {
        "matched": False,
        "journal_matched": 0,
        "journal_match_method": "",
        "journal_rank_source": "",
        "journal_name": "",
        "issn": "",
        "eissn": "",
        "jcr_impact_factor": "",
        "jcr_year": "",
        "jcr_quartile": "",
        "cas_quartile": "",
        "cas_category": "",
        "cas_top": "",
        "cas_warning": "",
        "cnki_composite_if": "",
        "cnki_comprehensive_if": "",
        "core_tags": "",
        "ccf_rank": "",
        "journal_quality_score": 0,
        "impact_factor": "",
        "impact_factor_year": "",
        "category": "",
    }


def fetch_journal_from_doi(doi: str) -> Dict[str, str]:
    """Resolve journal name and ISSN from DOI via Crossref; never raises."""
    doi = str(doi or "").strip()
    if not doi:
        return {}
    doi = doi.replace("https://doi.org/", "").replace("http://doi.org/", "").strip()
    try:
        import requests
    except ImportError:
        return {}
    try:
        response = requests.get(
            f"https://api.crossref.org/works/{doi}",
            headers={"User-Agent": "ResearchRadar/1.0 (mailto:research-radar@local)"},
            timeout=15,
        )
        if response.status_code != 200:
            return {}
        item = response.json().get("message", {}) or {}
        issn_list = item.get("ISSN", []) or []
        return {
            "journal_name": " ".join(item.get("container-title", [])[:1]),
            "issn": issn_list[0] if issn_list else "",
            "eissn": issn_list[1] if len(issn_list) > 1 else "",
        }
    except Exception:
        return {}


def match_journal_rank(
    journal_name: str = "",
    issn: str = "",
    eissn: str = "",
    doi: str = "",
    metrics: Optional[List[Dict]] = None,
) -> Dict:
    """Match journal metrics; never raises."""
    metrics = metrics if metrics is not None else load_journal_metrics(quiet=True)
    if not metrics:
        return _empty_rank()

    issn_key = normalize_issn(issn)
    eissn_key = normalize_issn(eissn)
    name_key = normalize_journal_name(journal_name)

    if doi and (not issn_key or not name_key):
        doi_meta = fetch_journal_from_doi(doi)
        if doi_meta.get("issn") and not issn_key:
            issn_key = normalize_issn(doi_meta["issn"])
            issn = doi_meta["issn"]
        if doi_meta.get("eissn") and not eissn_key:
            eissn_key = normalize_issn(doi_meta["eissn"])
            eissn = doi_meta["eissn"]
        if doi_meta.get("journal_name") and not name_key:
            journal_name = doi_meta["journal_name"]
            name_key = normalize_journal_name(journal_name)

    def _build(row: Dict, method: str) -> Dict:
        out = _empty_rank()
        out.update(
            {
                "matched": True,
                "journal_matched": 1,
                "journal_match_method": method,
                "journal_rank_source": row.get("source") or "journal_metrics.csv",
                "journal_name": row.get("journal_name", ""),
                "issn": row.get("issn", ""),
                "eissn": row.get("eissn", ""),
                "jcr_impact_factor": row.get("jcr_impact_factor", ""),
                "jcr_year": row.get("jcr_year", ""),
                "jcr_quartile": row.get("jcr_quartile", ""),
                "cas_quartile": row.get("cas_quartile", ""),
                "cas_category": row.get("cas_category", ""),
                "cas_top": row.get("cas_top", ""),
                "cas_warning": row.get("cas_warning", ""),
                "cnki_composite_if": row.get("cnki_composite_if", ""),
                "cnki_comprehensive_if": row.get("cnki_comprehensive_if", ""),
                "core_tags": row.get("core_tags", ""),
                "ccf_rank": row.get("ccf_rank", ""),
                "impact_factor": row.get("jcr_impact_factor", ""),
                "impact_factor_year": row.get("jcr_year", ""),
                "category": row.get("cas_category", ""),
            }
        )
        out["journal_quality_score"] = compute_journal_quality_score(out)
        return out

    for row in metrics:
        if issn_key and issn_key in {row.get("_issn_key"), row.get("_eissn_key")}:
            return _build(row, "issn")
    for row in metrics:
        if eissn_key and eissn_key in {row.get("_issn_key"), row.get("_eissn_key")}:
            return _build(row, "eissn")
    for row in metrics:
        if name_key and name_key == row.get("_name_key"):
            return _build(row, "exact_name" if not doi else "doi_journal_name")

    best = None
    best_ratio = 0.0
    for row in metrics:
        candidate = row.get("_name_key", "")
        if not name_key or not candidate:
            continue
        ratio = SequenceMatcher(None, name_key, candidate).ratio()
        if ratio > best_ratio:
            best = row
            best_ratio = ratio
    if best and best_ratio >= 0.88:
        return _build(best, f"fuzzy_name:{best_ratio:.2f}")
    return _empty_rank()


def compute_journal_quality_score(metrics: Dict) -> int:
    """Journal tier score; relevance should remain primary in final_score."""
    if not metrics.get("matched"):
        return 0
    score = 0
    q = str(metrics.get("jcr_quartile", "")).upper()
    if q == "Q1":
        score += 4
    elif q == "Q2":
        score += 2
    cas = str(metrics.get("cas_quartile", ""))
    if "1区" in cas or cas.startswith("1"):
        score += 4
    elif "2区" in cas or cas.startswith("2"):
        score += 2
    top = str(metrics.get("cas_top", "")).lower()
    if top in ("yes", "y", "true", "1", "是"):
        score += 2
    warn = str(metrics.get("cas_warning", "")).lower()
    if warn in ("yes", "y", "true", "1", "是", "warning"):
        score -= 5
    try:
        impact = float(str(metrics.get("jcr_impact_factor") or metrics.get("impact_factor") or "0"))
    except (TypeError, ValueError):
        impact = 0.0
    if impact >= 10:
        score += 5
    elif impact >= 5:
        score += 3
    elif impact >= 3:
        score += 2
    elif impact >= 1:
        score += 1
    tags = str(metrics.get("core_tags", "")).lower()
    if any(tag in tags for tag in ("北大核心", "cscd", "cssci", "南大核心", "核心")):
        score += 2
    return score


def apply_journal_rank_to_paper(paper: Dict, metrics: Optional[List[Dict]] = None) -> Dict:
    """Merge journal rank fields into paper dict and refresh scores."""
    rank = match_journal_rank(
        paper.get("journal_name") or paper.get("journal", ""),
        paper.get("issn", ""),
        paper.get("eissn", ""),
        paper.get("doi", ""),
        metrics,
    )
    paper.update(rank)
    relevance = int(paper.get("relevance_score") or 0)
    jq = int(rank.get("journal_quality_score") or 0)
    # Relevance-first: cap journal contribution
    journal_part = min(jq, max(6, relevance // 2 + 2)) if relevance else min(jq, 6)
    paper["journal_quality_score"] = jq
    paper["impact_factor_score"] = journal_part
    fresh = int(paper.get("freshness_score") or 0)
    cite = int(paper.get("citation_score") or 0)
    paper["final_score"] = relevance + journal_part + fresh + cite
    return paper


def journal_rank_summary(paper: Dict) -> str:
    if not paper.get("journal_matched") and not paper.get("matched"):
        return "未匹配"
    parts = []
    impact = paper.get("jcr_impact_factor") or paper.get("impact_factor")
    if impact not in ("", None, "待补充", "暂无影响因子数据", "谷歌学术不提供"):
        parts.append(f"IF {impact}")
    if paper.get("jcr_quartile"):
        parts.append(f"JCR {paper['jcr_quartile']}")
    if paper.get("cas_quartile"):
        parts.append(f"中科院 {paper['cas_quartile']}")
    warn = str(paper.get("cas_warning", "")).lower()
    if warn in ("yes", "y", "true", "1", "是", "warning"):
        parts.append("预警")
    if paper.get("core_tags"):
        parts.append(str(paper["core_tags"])[:24])
    return "｜".join(parts) if parts else "已匹配"


def paper_metrics_line(paper: Dict) -> str:
    """Compact list-line: IF | JCR | CAS | citations | abstract status."""
    parts = []
    rank = journal_rank_summary(paper)
    if rank and rank != "未匹配":
        parts.append(rank)
    elif not paper.get("journal_matched"):
        parts.append("未匹配")
    cite = paper.get("citation_count")
    if cite not in ("", None):
        parts.append(f"引用 {cite}")
    if paper.get("abstract_fetch_status") == "complete" or paper.get("abstract_is_complete"):
        parts.append("摘要完整")
    elif paper.get("abstract_fetch_status") == "snippet":
        parts.append("摘要片段")
    elif paper.get("abstract_fetch_status") == "translated":
        parts.append("摘要已译")
    else:
        parts.append("摘要未补全")
    return "｜".join(parts)


def export_unmatched_journals(papers: List[Dict]) -> List[Dict]:
    """Aggregate unmatched journal names from paper list."""
    buckets: Dict[str, Dict] = {}
    for paper in papers:
        if paper.get("journal_matched") or paper.get("matched"):
            continue
        journal = (paper.get("journal") or paper.get("journal_name") or "").strip()
        if not journal:
            continue
        key = normalize_journal_name(journal) or journal.lower()
        bucket = buckets.setdefault(
            key,
            {
                "journal_name": journal,
                "issn": paper.get("issn", ""),
                "eissn": paper.get("eissn", ""),
                "paper_count": 0,
                "example_title": paper.get("title", ""),
            },
        )
        bucket["paper_count"] += 1
    return sorted(buckets.values(), key=lambda x: (-x["paper_count"], x["journal_name"]))


def metrics_file_stats(path: Optional[str] = None) -> Dict:
    path = path or metrics_path()
    if not os.path.exists(path):
        return {"exists": False, "path": path, "count": 0, "mtime": ""}
    mtime = os.path.getmtime(path)
    from datetime import datetime

    metrics = load_journal_metrics(path, quiet=True)
    return {
        "exists": True,
        "path": path,
        "count": len(metrics),
        "mtime": datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M"),
    }
