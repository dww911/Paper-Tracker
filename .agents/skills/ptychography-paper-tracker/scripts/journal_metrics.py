"""Backward-compatible facade; use journal_rank_enhancer for new code."""

from journal_rank_enhancer import (  # noqa: F401
    METRIC_FIELDS,
    compute_journal_quality_score,
    example_metrics_path,
    export_unmatched_journals,
    journal_rank_summary,
    load_journal_metrics,
    match_journal_rank,
    metrics_path,
    normalize_issn,
    normalize_journal_name,
    apply_journal_rank_to_paper,
)


def impact_factor_score(value: str) -> int:
    """Legacy IF-only score; prefer compute_journal_quality_score."""
    from journal_rank_enhancer import compute_journal_quality_score

    return compute_journal_quality_score(
        {"matched": True, "jcr_impact_factor": value, "jcr_quartile": "", "cas_quartile": "", "cas_warning": "", "core_tags": ""}
    )


def match_journal_metrics(journal_name: str = "", issn: str = "", eissn: str = "", metrics=None) -> dict:
    rank = match_journal_rank(journal_name, issn, eissn, metrics)
    if not rank.get("matched"):
        return {
            "impact_factor": "",
            "impact_factor_year": "",
            "jcr_quartile": "",
            "cas_quartile": "",
            "category": "",
            "journal_metrics_match": "missing",
        }
    method = rank.get("journal_match_method", "name")
    return {
        **rank,
        "category": rank.get("cas_category", ""),
        "journal_metrics_match": method,
    }
