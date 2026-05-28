"""Extract text from uploaded PDFs using PyMuPDF."""
import re
from pathlib import Path
from typing import Dict, Optional


def extract_pdf_text(path: Path, max_pages: int = 30) -> Dict[str, str]:
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return {"error": "未安装 PyMuPDF，请运行: pip install pymupdf", "text": ""}

    if not path.exists():
        return {"error": "文件不存在", "text": ""}

    try:
        doc = fitz.open(str(path))
    except Exception as exc:
        return {"error": f"无法打开 PDF: {exc}", "text": ""}

    parts = []
    for page_num in range(min(len(doc), max_pages)):
        page = doc[page_num]
        parts.append(page.get_text("text"))
    doc.close()
    text = "\n".join(parts).strip()
    if len(text) < 80:
        return {
            "error": "提取文本过短，可能是扫描版 PDF（暂不支持 OCR）",
            "text": text,
        }
    return {"text": text, "error": ""}


def guess_metadata_from_text(text: str) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    doi_match = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", text, re.I)
    if doi_match:
        meta["doi"] = doi_match.group(0)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if lines:
        meta["title_guess"] = lines[0][:300]
    abstract_idx = text.lower().find("abstract")
    if abstract_idx >= 0:
        meta["abstract_guess"] = text[abstract_idx: abstract_idx + 2000]
    return meta
