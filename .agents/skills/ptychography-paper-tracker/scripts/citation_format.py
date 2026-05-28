"""Format paper dicts as bibliographic citations."""
from typing import Dict, List


def _authors_text(paper: Dict) -> str:
    authors = paper.get("authors") or ""
    if isinstance(authors, str) and authors.startswith("["):
        try:
            import json
            authors = ", ".join(json.loads(authors))
        except Exception:
            pass
    return str(authors or "佚名")


def format_gbt7714(paper: Dict) -> str:
    authors = _authors_text(paper)
    title = paper.get("title") or "无标题"
    journal = paper.get("journal") or ""
    year = paper.get("year") or paper.get("publication_year") or "n.d."
    doi = paper.get("doi") or ""
    url = paper.get("display_url") or paper.get("url") or ""
    line = f"{authors}. {title}[J]. {journal}, {year}."
    if doi:
        line += f" DOI:{doi}."
    elif url:
        line += f" {url}"
    return line


def format_apa(paper: Dict) -> str:
    authors = _authors_text(paper)
    year = paper.get("year") or "n.d."
    title = paper.get("title") or "Untitled"
    journal = paper.get("journal") or ""
    return f"{authors} ({year}). {title}. *{journal}*."


def format_ieee(paper: Dict) -> str:
    authors = _authors_text(paper)
    title = paper.get("title") or "Untitled"
    journal = paper.get("journal") or ""
    year = paper.get("year") or "n.d."
    return f"{authors}, \"{title},\" *{journal}*, {year}."


def format_bibtex(paper: Dict, key: str = "paper") -> str:
    title = (paper.get("title") or "untitled").replace("{", "").replace("}", "")
    journal = (paper.get("journal") or "").replace("{", "").replace("}", "")
    year = paper.get("year") or ""
    doi = paper.get("doi") or ""
    lines = [
        f"@article{{{key},",
        f"  title = {{{title}}},",
        f"  journal = {{{journal}}},",
        f"  year = {{{year}}},",
    ]
    if doi:
        lines.append(f"  doi = {{{doi}}},")
    lines.append("}")
    return "\n".join(lines)


def format_markdown_list(papers: List[Dict]) -> str:
    lines = []
    for idx, paper in enumerate(papers, 1):
        url = paper.get("display_url") or paper.get("url") or ""
        title = paper.get("title") or "Untitled"
        if url:
            lines.append(f"{idx}. [{title}]({url})")
        else:
            lines.append(f"{idx}. {title}")
    return "\n".join(lines)


FORMATTERS = {
    "gbt7714": format_gbt7714,
    "apa": format_apa,
    "ieee": format_ieee,
    "bibtex": format_bibtex,
    "markdown": lambda p: format_markdown_list([p]),
}


def format_citations(papers: List[Dict], style: str = "gbt7714") -> str:
    style = (style or "gbt7714").lower()
    if style == "bibtex":
        return "\n\n".join(format_bibtex(p, f"p{i}") for i, p in enumerate(papers, 1))
    if style == "markdown":
        return format_markdown_list(papers)
    formatter = FORMATTERS.get(style, format_gbt7714)
    return "\n".join(formatter(p) for p in papers)
