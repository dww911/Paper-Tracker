"""Fixed, auditable research task agents for Research Radar."""

from __future__ import annotations

import json
from typing import Callable, Optional

from radar_db import connect, init_db, utc_now


AGENT_SPECS = {
    "LiteratureFetchAgent": {
        "label": "文献获取 Agent",
        "steps": [
            ("检索文献", "fetch_papers", "按当前 profile 运行 daily/high_quality 检索。"),
            ("补全元数据", "enrich_metadata", "补摘要、DOI、IF、引用数等字段。"),
            ("写入主库", "save_to_sqlite", "归档到 SQLite，并保持 Excel/Markdown 兼容。"),
            ("生成报告", "generate_daily_report", "生成日报或年度资料包。"),
        ],
    },
    "PaperReadingAgent": {
        "label": "文献精读 Agent",
        "steps": [
            ("读取论文上下文", "load_paper", "读取论文、PDF 文本、AI 笔记和用户笔记。"),
            ("调用精读 Prompt", "generate_paper_note", "按 full_paper_reading_prompt 输出结构化精读。"),
            ("保存笔记", "save_notes", "写入 reading_notes 与 paper_notes Markdown。"),
            ("更新 RAG 索引", "rag_index", "把新笔记加入文献库索引。"),
        ],
    },
    "IdeaMiningAgent": {
        "label": "灵感挖掘 Agent",
        "steps": [
            ("RAG 检索证据", "rag_search", "从论文、笔记和报告里找支持依据。"),
            ("生成灵感草稿", "mine_idea", "按 idea_mining_prompt 输出创新点和实验方案。"),
            ("关联论文", "link_papers", "把支持论文写入灵感关联字段。"),
            ("保存灵感", "save_idea", "保存到 ideas 和 idea_notes。"),
        ],
    },
    "WeeklyReviewAgent": {
        "label": "周报/综述 Agent",
        "steps": [
            ("选择文献", "select_papers", "选择本周新增或综述资料包文献。"),
            ("组织证据", "build_context", "整理论文摘要、AI 笔记和关键贡献。"),
            ("生成报告", "generate_review", "生成周报、综述大纲或综述初稿。"),
            ("保存输出", "save_markdown", "保存 Markdown 和文献表。"),
        ],
    },
    "ProgressAgent": {
        "label": "研究进展 Agent",
        "steps": [
            ("统计进展", "progress_metrics", "统计文献、精读、笔记、灵感和里程碑。"),
            ("分析缺口", "progress_prompt", "判断关键词覆盖和写作准备度。"),
            ("生成建议", "next_actions", "生成下一步阅读和写作建议。"),
            ("保存快照", "save_snapshot", "记录 progress_snapshots 和报告。"),
        ],
    },
    "ResearchRoadmapAgent": {
        "label": "研究方向地图 Agent",
        "steps": [
            ("筛选候选里程碑", "roadmap_candidates", "按年份、引用、星级、综述标记和里程碑标记筛选候选论文。"),
            ("生成阶段与分支", "roadmap_build", "生成时间轴阶段、脉络树节点和精读路线。"),
            ("保存方向地图", "roadmap_save", "写入 roadmap_nodes / roadmap_edges / roadmap_papers。"),
            ("生成地图报告", "roadmap_report", "导出 Markdown 版方向地图报告。"),
        ],
    },
    "CitationAgent": {
        "label": "引用整理 Agent",
        "steps": [
            ("读取引用篮", "load_citation_collection", "读取引用篮中的论文。"),
            ("格式化引用", "format_citations", "生成 GB/T 7714、APA、IEEE、BibTeX 等格式。"),
            ("解释用途", "citation_prompt", "说明每篇文献可支撑的论点。"),
            ("保存导出", "save_citations", "保存引用输出或显示预览。"),
        ],
    },
}


def list_agents() -> list[dict]:
    return [{"name": name, **spec} for name, spec in AGENT_SPECS.items()]


def create_plan(agent_name: str, profile: str = "", user_request: str = "") -> int:
    init_db()
    spec = AGENT_SPECS.get(agent_name)
    if not spec:
        raise ValueError(f"Unknown agent: {agent_name}")
    plan = [
        {"index": idx, "name": name, "tool": tool, "description": description}
        for idx, (name, tool, description) in enumerate(spec["steps"], 1)
    ]
    now = utc_now()
    with connect() as conn:
        cursor = conn.execute(
            """
            INSERT INTO agent_runs (agent_name, profile, user_request, plan_json, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, 'planned', ?, ?)
            """,
            (agent_name, profile or "", user_request or spec["label"], json.dumps(plan, ensure_ascii=False), now, now),
        )
        run_id = int(cursor.lastrowid)
        for item in plan:
            conn.execute(
                """
                INSERT INTO agent_steps (run_id, step_index, name, tool_name, status, message)
                VALUES (?, ?, ?, ?, 'planned', ?)
                """,
                (run_id, item["index"], item["name"], item["tool"], item["description"]),
            )
        conn.commit()
    return run_id


def get_run(run_id: int) -> Optional[dict]:
    init_db()
    with connect() as conn:
        run = conn.execute("SELECT * FROM agent_runs WHERE id = ?", (run_id,)).fetchone()
        if not run:
            return None
        data = dict(run)
        data["plan"] = json.loads(data.get("plan_json") or "[]")
        data["steps"] = [
            dict(row)
            for row in conn.execute("SELECT * FROM agent_steps WHERE run_id = ? ORDER BY step_index", (run_id,)).fetchall()
        ]
        data["outputs"] = [
            dict(row)
            for row in conn.execute("SELECT * FROM agent_outputs WHERE run_id = ? ORDER BY id", (run_id,)).fetchall()
        ]
    return data


def recent_runs(limit: int = 20) -> list[dict]:
    init_db()
    with connect() as conn:
        return [dict(row) for row in conn.execute("SELECT * FROM agent_runs ORDER BY updated_at DESC, id DESC LIMIT ?", (limit,)).fetchall()]


def execute_run(run_id: int, tool_executor: Optional[Callable[[str, dict], dict]] = None) -> dict:
    init_db()
    run = get_run(run_id)
    if not run:
        raise ValueError(f"Unknown agent run: {run_id}")
    now = utc_now()
    with connect() as conn:
        conn.execute("UPDATE agent_runs SET status='running', updated_at=? WHERE id=?", (now, run_id))
        conn.commit()
    outputs = []
    for step in run["steps"]:
        step_id = int(step["id"])
        start = utc_now()
        with connect() as conn:
            conn.execute("UPDATE agent_steps SET status='running', started_at=? WHERE id=?", (start, step_id))
            conn.commit()
        try:
            result = tool_executor(step["tool_name"], run) if tool_executor else {}
            message = result.get("message") or step.get("message") or "完成"
            content = result.get("content") or message
            path = result.get("path") or ""
            output_type = result.get("output_type") or "text"
            finish = utc_now()
            with connect() as conn:
                conn.execute(
                    "UPDATE agent_steps SET status='done', message=?, finished_at=? WHERE id=?",
                    (message, finish, step_id),
                )
                conn.execute(
                    """
                    INSERT INTO agent_outputs (run_id, step_id, output_type, title, content, path, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (run_id, step_id, output_type, step["name"], content, path, finish),
                )
                conn.commit()
            outputs.append({"step": step["name"], "message": message, "path": path})
        except Exception as exc:
            finish = utc_now()
            with connect() as conn:
                conn.execute(
                    "UPDATE agent_steps SET status='failed', message=?, finished_at=? WHERE id=?",
                    (str(exc), finish, step_id),
                )
                conn.execute("UPDATE agent_runs SET status='failed', updated_at=? WHERE id=?", (finish, run_id))
                conn.commit()
            return get_run(run_id) or {}
    finish = utc_now()
    with connect() as conn:
        conn.execute("UPDATE agent_runs SET status='done', updated_at=? WHERE id=?", (finish, run_id))
        conn.commit()
    return get_run(run_id) or {}
