"""Prompt template loading for Research Radar AI tasks."""

from __future__ import annotations

import os
from pathlib import Path
from string import Formatter
from typing import Dict

from radar_db import connect, init_db, utc_now


SKILL_DIR = Path(__file__).resolve().parents[1]
PROMPTS_DIR = SKILL_DIR / "prompts"


PROMPT_REGISTRY = {
    "abstract_reading_prompt": ("abstract_reading", "摘要阅读与结构化精读"),
    "full_paper_reading_prompt": ("full_paper_reading", "PDF 全文精读"),
    "paper_rating_prompt": ("paper_rating", "论文星级评分"),
    "idea_mining_prompt": ("idea_mining", "灵感挖掘"),
    "idea_polish_prompt": ("idea_polish", "想法打磨"),
    "review_writing_prompt": ("review_writing", "综述写作"),
    "weekly_report_prompt": ("weekly_report", "周报生成"),
    "roadmap_prompt": ("roadmap", "研究脉络分析"),
    "progress_prompt": ("progress", "研究进展分析"),
    "citation_prompt": ("citation", "引用解释与整理"),
    "rag_answer_prompt": ("rag_answer", "基于文献库问答"),
}


class SafeDict(dict):
    def __missing__(self, key: str) -> str:
        return ""


def safe_format(template: str, values: Dict) -> str:
    safe_values = SafeDict({key: "" if value is None else value for key, value in values.items()})
    formatter = Formatter()
    parts = []
    for literal, field_name, format_spec, conversion in formatter.parse(template):
        parts.append(literal)
        if field_name is None:
            continue
        value = safe_values[field_name]
        if conversion == "r":
            value = repr(value)
        elif conversion == "s":
            value = str(value)
        if format_spec:
            try:
                value = format(value, format_spec)
            except (ValueError, TypeError):
                value = str(value)
        parts.append(str(value))
    return "".join(parts)


def template_path(name: str) -> Path:
    return PROMPTS_DIR / f"{name}.md"


def load_template(name: str, fallback: str = "") -> str:
    path = template_path(name)
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        text = fallback
    return text or fallback


def render_template(name: str, values: Dict, fallback: str = "") -> str:
    rendered = safe_format(load_template(name, fallback), values)
    log_prompt_run(
        name,
        source_type=str(values.get("source_type") or ""),
        source_id=str(values.get("source_id") or ""),
        input_summary=str(values.get("user_question") or values.get("title") or values.get("topic") or "")[:500],
        output_summary=f"rendered {len(rendered)} chars",
        status="rendered",
    )
    return rendered


def log_prompt_run(
    template_name: str,
    task_type: str = "",
    source_type: str = "",
    source_id: str = "",
    input_summary: str = "",
    output_summary: str = "",
    status: str = "rendered",
) -> None:
    try:
        init_db()
        registry_task = PROMPT_REGISTRY.get(template_name, ("", ""))[0]
        with connect() as conn:
            conn.execute(
                """
                INSERT INTO prompt_runs (
                    template_name, task_type, source_type, source_id,
                    input_summary, output_summary, status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    template_name,
                    task_type or registry_task,
                    source_type,
                    source_id,
                    input_summary,
                    output_summary,
                    status,
                    utc_now(),
                ),
            )
            conn.commit()
    except Exception:
        pass


def ensure_prompt_template_records() -> None:
    init_db()
    now = utc_now()
    with connect() as conn:
        for name, (task_type, description) in PROMPT_REGISTRY.items():
            path = str(template_path(name))
            conn.execute(
                """
                INSERT INTO prompt_templates (name, task_type, template_path, description, version, created_at, updated_at)
                VALUES (?, ?, ?, ?, '1', ?, ?)
                ON CONFLICT(name) DO UPDATE SET
                    task_type=excluded.task_type,
                    template_path=excluded.template_path,
                    description=excluded.description,
                    updated_at=excluded.updated_at
                """,
                (name, task_type, path, description, now, now),
            )
        conn.commit()


def prompt_status() -> list[dict]:
    ensure_prompt_template_records()
    rows = []
    with connect() as conn:
        for row in conn.execute("SELECT * FROM prompt_templates ORDER BY task_type, name").fetchall():
            item = dict(row)
            item["exists"] = os.path.exists(item["template_path"])
            rows.append(item)
    return rows


def recent_prompt_runs(limit: int = 20) -> list[dict]:
    init_db()
    with connect() as conn:
        return [
            dict(row)
            for row in conn.execute(
                "SELECT * FROM prompt_runs ORDER BY created_at DESC, id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        ]
