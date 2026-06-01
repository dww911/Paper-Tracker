"""Translate English abstracts to Chinese for display and archival."""

from __future__ import annotations

import os
import re
from typing import Dict, Optional

try:
    import requests
except ImportError:
    requests = None

MISSING_ZH_VALUES = {
    "",
    "无",
    "无摘要",
    "暂无",
    "暂无摘要",
    "none",
    "n/a",
    "null",
    "等待文献摘要 skill 生成。",
    "等待文献摘要 skill 生成",
}


def is_missing_zh(text: Optional[str]) -> bool:
    if text is None:
        return True
    cleaned = str(text).strip()
    if not cleaned:
        return True
    if cleaned.lower() in {value.lower() for value in MISSING_ZH_VALUES if value}:
        return True
    return cleaned.startswith("等待文献摘要")


def _chunk_text(text: str, max_len: int = 3500) -> list[str]:
    text = text.strip()
    if len(text) <= max_len:
        return [text]
    chunks: list[str] = []
    start = 0
    while start < len(text):
        end = min(start + max_len, len(text))
        if end < len(text):
            split_at = text.rfind(" ", start, end)
            if split_at > start + max_len // 2:
                end = split_at
        chunks.append(text[start:end].strip())
        start = end
    return [chunk for chunk in chunks if chunk]


def _looks_english(text: str) -> bool:
    letters = re.findall(r"[A-Za-z]", text)
    if not letters:
        return False
    cjk = re.findall(r"[\u4e00-\u9fff]", text)
    return len(letters) >= max(20, len(cjk) * 2)


def translate_with_openai(text: str) -> Optional[str]:
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
                "model": os.environ.get("OPENAI_TRANSLATE_MODEL", "gpt-4o-mini"),
                "temperature": 0.2,
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "你是学术翻译助手。将英文学术论文摘要翻译成流畅、准确的中文，"
                            "保留专业术语，只输出译文，不要加标题或解释。"
                        ),
                    },
                    {"role": "user", "content": text[:12000]},
                ],
            },
            timeout=90,
        )
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"]
        return str(content or "").strip() or None
    except Exception:
        return None


def translate_with_anthropic(text: str) -> Optional[str]:
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
                "model": os.environ.get("ANTHROPIC_TRANSLATE_MODEL", "claude-3-5-haiku-20241022"),
                "max_tokens": 4096,
                "temperature": 0.2,
                "messages": [
                    {
                        "role": "user",
                        "content": (
                            "请将以下英文学术摘要翻译成流畅、准确的中文，保留专业术语，"
                            "只输出译文：\n\n" + text[:12000]
                        ),
                    }
                ],
            },
            timeout=90,
        )
        response.raise_for_status()
        blocks = response.json().get("content", [])
        parts = [block.get("text", "") for block in blocks if block.get("type") == "text"]
        content = "\n".join(part for part in parts if part).strip()
        return content or None
    except Exception:
        return None


def translate_with_deep_translator(text: str) -> Optional[str]:
    try:
        from deep_translator import GoogleTranslator
    except ImportError:
        return None
    try:
        translator = GoogleTranslator(source="en", target="zh-CN")
        translated_parts = []
        for chunk in _chunk_text(text, 4500):
            translated_parts.append(translator.translate(chunk))
        return "".join(part for part in translated_parts if part).strip() or None
    except Exception:
        return None


def translate_with_mymemory(text: str) -> Optional[str]:
    if requests is None:
        return None
    translated_parts = []
    try:
        for chunk in _chunk_text(text, 450):
            response = requests.get(
                "https://api.mymemory.translated.net/get",
                params={"q": chunk, "langpair": "en|zh-CN"},
                timeout=30,
            )
            response.raise_for_status()
            translated = response.json().get("responseData", {}).get("translatedText", "")
            if translated:
                translated_parts.append(str(translated).strip())
        return " ".join(translated_parts).strip() or None
    except Exception:
        return None


def translate_en_to_zh(text: str) -> str:
    text = str(text or "").strip()
    if not text or text in {"无摘要", "无"}:
        return ""
    if not _looks_english(text):
        return text

    for provider in (
        translate_with_openai,
        translate_with_anthropic,
        translate_with_deep_translator,
        translate_with_mymemory,
    ):
        translated = provider(text)
        if translated and not is_missing_zh(translated):
            return translated.strip()
    return ""


def ensure_paper_chinese_abstract(paper: Dict) -> Dict:
    updated = dict(paper)
    original = (
        updated.get("abstract_original")
        or updated.get("abstract")
        or ""
    ).strip()
    chinese = updated.get("abstract_zh") or updated.get("摘要中文翻译") or ""

    if not is_missing_zh(chinese):
        updated["abstract_zh"] = str(chinese).strip()
        updated["摘要中文翻译"] = updated["abstract_zh"]
        return updated

    if not original or original in {"无摘要", "无"}:
        return updated

    translated = translate_en_to_zh(original)
    if translated:
        updated["abstract_zh"] = translated
        updated["摘要中文翻译"] = translated
        if not updated.get("abstract_original"):
            updated["abstract_original"] = original
        if not updated.get("abstract_is_complete"):
            updated["abstract_fetch_status"] = "translated"
    return updated
