"""Call Kimi / OpenAI / Anthropic to parse papers into structured Chinese notes."""

from __future__ import annotations

import json
import os
import re
from typing import Dict, List, Optional

try:
    import requests
except ImportError:
    requests = None

PARSE_FIELD_LABELS = [
    "摘要中文翻译",
    "研究背景",
    "研究目的",
    "核心方法",
    "论文创新点",
    "实验结果",
    "总结",
    "未来展望",
    "可创新点",
    "对我的启发",
    "是否值得精读",
]

FAILED_NOTE_MARKERS = ("解析失败", "待生成", "待判断")


def ai_parse_configured() -> bool:
    return bool(
        os.environ.get("KIMI_API_KEY", "").strip()
        or os.environ.get("MOONSHOT_API_KEY", "").strip()
        or os.environ.get("OPENAI_API_KEY", "").strip()
        or os.environ.get("ANTHROPIC_API_KEY", "").strip()
    )


def kimi_api_key() -> str:
    return os.environ.get("KIMI_API_KEY", "").strip() or os.environ.get("MOONSHOT_API_KEY", "").strip()


def kimi_api_base() -> str:
    return os.environ.get("KIMI_API_BASE", "https://api.moonshot.cn/v1").rstrip("/")


def normalize_paper_for_parse(paper: Dict) -> Dict:
    authors = paper.get("authors", [])
    if isinstance(authors, str):
        try:
            parsed = json.loads(authors)
            authors = parsed if isinstance(parsed, list) else [authors]
        except json.JSONDecodeError:
            authors = [item.strip() for item in authors.split(",") if item.strip()]
    if not authors:
        authors = ["未知作者"]

    abstract = (
        paper.get("abstract")
        or paper.get("abstract_original")
        or paper.get("abstract_zh")
        or ""
    )
    return {
        "title": paper.get("title", ""),
        "authors": authors,
        "published_time": str(paper.get("published_time") or paper.get("year") or "未知"),
        "link": paper.get("link") or paper.get("url", ""),
        "abstract": abstract,
        "relevance_score": paper.get("relevance_score", paper.get("final_score", "未评分")),
    }


def build_parse_prompt(paper: Dict, profile: Optional[Dict] = None) -> str:
    profile = profile or {}
    focus_lines = "\n".join(f"- {item}" for item in profile.get("parse_focus", []))
    focus_block = (
        f"\n【当前研究方向】{profile.get('name', profile.get('display_name', '科研方向'))}\n{focus_lines}\n"
        if focus_lines
        else ""
    )
    authors = paper.get("authors", [])
    if isinstance(authors, list):
        author_text = ", ".join(str(item) for item in authors)
    else:
        author_text = str(authors)

    return f"""你是当前研究方向的资深科研顾问。请基于下列论文信息完成结构化拆解。
要求：禁止编造；信息不足时明确写“摘要未提供，需阅读原文确认”；语言简洁专业；每字段单独成段。
{focus_block}
【论文基础信息】
标题：{paper.get('title', '')}
作者：{author_text}
发表时间：{paper.get('published_time', '')}
原文链接：{paper.get('link', '')}
摘要：{paper.get('abstract', '')}
相关性评分：{paper.get('relevance_score', '未评分')}

请严格按以下字段输出，每段以【字段名】：开头：
【摘要中文翻译】：
【研究背景】：
【研究目的】：
【核心方法】：
【论文创新点】：
【实验结果】：
【总结】：
【未来展望】：
【可创新点】：
【对我的启发】：
【是否值得精读】：
"""


def parse_structured_fields(text: str) -> Dict[str, str]:
    result: Dict[str, str] = {}
    if not text:
        return result
    for label in PARSE_FIELD_LABELS:
        pattern = rf"【{re.escape(label)}】\s*[:：]\s*(.*?)(?=\n\s*【|\Z)"
        match = re.search(pattern, text, flags=re.S)
        if match:
            value = re.sub(r"\s+\n", "\n", match.group(1).strip())
            if value:
                result[label] = value
    return result


def call_kimi(prompt: str) -> Optional[str]:
    api_key = kimi_api_key()
    if not api_key or requests is None:
        return None
    try:
        response = requests.post(
            f"{kimi_api_base()}/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": os.environ.get("KIMI_PARSE_MODEL", "moonshot-v1-32k"),
                "temperature": 0.2,
                "messages": [
                    {"role": "system", "content": "你是学术论文结构化解析助手，只输出中文字段内容。"},
                    {"role": "user", "content": prompt},
                ],
            },
            timeout=120,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception:
        return None


def call_openai(prompt: str) -> Optional[str]:
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key or requests is None:
        return None
    try:
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": os.environ.get("OPENAI_PARSE_MODEL", "gpt-4o-mini"),
                "temperature": 0.2,
                "messages": [
                    {"role": "system", "content": "你是学术论文结构化解析助手，只输出中文字段内容。"},
                    {"role": "user", "content": prompt},
                ],
            },
            timeout=120,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception:
        return None


def call_anthropic(prompt: str) -> Optional[str]:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key or requests is None:
        return None
    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": os.environ.get("ANTHROPIC_PARSE_MODEL", "claude-3-5-haiku-20241022"),
                "max_tokens": 4096,
                "temperature": 0.2,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=120,
        )
        response.raise_for_status()
        blocks = response.json().get("content", [])
        parts = [block.get("text", "") for block in blocks if block.get("type") == "text"]
        return "\n".join(part for part in parts if part).strip() or None
    except Exception:
        return None


def call_llm(prompt: str) -> Optional[str]:
    provider = os.environ.get("PAPER_AI_PROVIDER", "").strip().lower()
    if provider == "kimi" or provider == "moonshot":
        order = ["kimi", "openai", "anthropic"]
    elif provider == "anthropic":
        order = ["anthropic", "kimi", "openai"]
    elif provider == "openai":
        order = ["openai", "kimi", "anthropic"]
    elif kimi_api_key():
        order = ["kimi", "openai", "anthropic"]
    else:
        order = ["openai", "anthropic", "kimi"]
    callers = {"kimi": call_kimi, "openai": call_openai, "anthropic": call_anthropic}
    for name in order:
        content = callers[name](prompt)
        if content:
            return content
    return None


def parse_paper_with_ai(paper: Dict, profile: Optional[Dict] = None) -> Dict[str, str]:
    if not ai_parse_configured():
        return {}
    normalized = normalize_paper_for_parse(paper)
    if not normalized.get("abstract") and not normalized.get("title"):
        return {}
    prompt = build_parse_prompt(normalized, profile)
    raw = call_llm(prompt)
    if not raw:
        return {}
    return parse_structured_fields(raw)


def note_needs_generation(note: Optional[Dict]) -> bool:
    if not note:
        return True
    required = (
        "research_background",
        "research_purpose",
        "core_method",
        "main_results",
        "paper_contribution",
        "possible_ideas",
    )
    for key in required:
        value = str(note.get(key) or "").strip()
        if not value:
            return True
        if any(marker in value for marker in FAILED_NOTE_MARKERS):
            return True
    return False


def parsed_to_reading_note_payload(paper: Dict, parsed: Dict[str, str]) -> Dict:
    return {
        **paper,
        **parsed,
        "paper_topic": paper.get("title", ""),
        "research_background": parsed.get("研究背景", ""),
        "research_purpose": parsed.get("研究目的", ""),
        "core_method": parsed.get("核心方法", ""),
        "main_results": parsed.get("实验结果", ""),
        "paper_contribution": parsed.get("论文创新点", ""),
        "inspiration": parsed.get("对我的启发", ""),
        "worth_reading": parsed.get("是否值得精读", ""),
        "reason": parsed.get("总结", ""),
        "possible_ideas": parsed.get("可创新点", ""),
        "摘要中文翻译": parsed.get("摘要中文翻译", ""),
    }
