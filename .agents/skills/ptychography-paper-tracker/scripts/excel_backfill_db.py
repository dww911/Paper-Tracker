#!/usr/bin/env python3
"""Import papers from Ptychography_论文全量库.xlsx into SQLite."""
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import pandas as pd  # noqa: E402

from ptychography_tracker import BASE_CONFIG  # noqa: E402
from radar_db import init_db, record_run, upsert_papers  # noqa: E402


def excel_row_to_paper(row) -> dict:
    title = str(row.get("论文名字", "") or "").strip()
    link = str(row.get("网址", "") or "").strip()
    journal = str(row.get("期刊", "") or "").strip()
    impact = str(row.get("影响因子", "") or "").strip()
    pub = str(row.get("发布时间", "") or "").strip()
    year_s = pub[:4] if len(pub) >= 4 else ""
    abstract_zh = str(row.get("摘要中文翻译", "") or "").strip()
    return {
        "title": title,
        "link": link,
        "url": link,
        "journal": journal,
        "journal_name": journal,
        "impact_factor": impact,
        "published_time": pub,
        "year": int(year_s) if year_s.isdigit() else None,
        "abstract_zh": abstract_zh,
        "摘要中文翻译": abstract_zh,
        "研究背景": str(row.get("研究背景", "") or ""),
        "论文创新点": str(row.get("论文创新点", "") or ""),
        "实验结果": str(row.get("实验结果", "") or ""),
        "总结": str(row.get("总结", "") or ""),
        "未来展望": str(row.get("未来展望", "") or ""),
        "可创新点": str(row.get("可创新点", "") or ""),
        "abstract_is_complete": bool(abstract_zh),
        "ingestion_tier": "full",
        "is_relevant": 1,
    }


def main() -> int:
    path = BASE_CONFIG["EXCEL_SAVE_PATH"]
    if not os.path.exists(path):
        print(f"Excel not found: {path}")
        return 1
    init_db()
    papers = []
    xl = pd.ExcelFile(path)
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        for _, row in df.iterrows():
            title = str(row.get("论文名字", "") or "").strip()
            if not title or title == "论文名字":
                continue
            papers.append(excel_row_to_paper(row))
    if not papers:
        print("No paper rows found in Excel.")
        return 1
    profile = os.environ.get("BACKFILL_PROFILE", "electron_ptychography")
    print(f"Importing {len(papers)} rows from {path} -> profile {profile}")
    stats = upsert_papers(papers, profile, "annual_summary")
    run_stats = {
        "retrieved": len(papers),
        "kept_after_relevance": len(papers),
        "new_papers": stats.get("inserted", 0),
        "ingested_count": stats.get("ingested", 0),
        "updated_count": stats.get("updated", 0),
        "run_year": "2024-2026",
        "max_results": 100,
        "data_sources": "Excel backfill",
        "ingest_policy": "all",
        **stats,
    }
    run_id = record_run(profile, "annual_summary", run_stats)
    print("ingest:", stats)
    print("run_id:", run_id)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
