---
name: ptychography-paper-tracker
description: Configurable academic paper tracking and research brief generation for Ptychography, electron microscopy, X-ray ptychography, medical imaging AI, and custom research directions. Use when Codex needs to find recent papers, run daily or annual literature tracking, configure a research profile before searching, deduplicate papers, score relevance, append structured results to Excel, or generate Chinese research summaries and innovation ideas.
---

# Ptychography Paper Tracker

Use this skill as a configurable literature tracking assistant. Always select or confirm a research profile before running a search. The default profile is `electron_ptychography`.

## Quick Start

List available research directions:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --list_profiles
```

Create a new research direction without editing JSON:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --create_profile
```

Check the local environment:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --doctor
```

Preview recent papers without writing files:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 7 --max_results 20 --dry_run
```

Run and append results to Excel:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20
```

Push a short digest through Server Chan Turbo:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode daily --notify serverchan
```

Test Server Chan Turbo:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --test_notify --notify serverchan
```

Generate annual summary:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode annual_summary --start_year 2024 --end_year 2026 --max_papers_per_year 50
```

Preview high-quality literature radar:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode high_quality --max_results 30 --dry_run
```

Run high-quality radar and push A+/A papers:

```bash
python .agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py --profile electron_ptychography --mode high_quality --max_results 30 --notify serverchan
```

## Research Profiles

Profiles live in `research_profiles.json`. Each profile controls:

- `include_keywords`: positive search terms
- `exclude_keywords`: terms that reduce relevance or filter unrelated areas
- `arxiv_categories`: arXiv category scope
- `sources`: intended data sources
- `must_have_any`: at least one of these terms must be found when provided
- `research_focus`: extra domain signals used for scoring and report context
- `score_rules`: profile-specific weights and `min_score`
- `score_threshold`: legacy minimum relevance score fallback
- `parse_focus`: domain-specific instructions for structured analysis
- `output_fields`: Excel/report fields

Current built-in profiles:

- `electron_ptychography`: electron microscopy, 4D-STEM, electron ptychography, WDD, SSB, phase retrieval
- `xray_ptychography`: X-ray ptychography, CDI, synchrotron, nanotomography
- `medical_ai`: medical imaging AI, diagnosis, segmentation, clinical evaluation

When adapting to a new field, add a new profile instead of editing the Python search logic.

## Workflow

1. Choose a profile with `--profile`.
2. Run `--dry_run` first for new directions to inspect relevance and avoid polluting history.
3. If results look correct, rerun without `--dry_run`.
4. The script searches papers, applies relevance scoring, deduplicates by DOI/arXiv/Semantic Scholar/title hash, prints structured parse prompts, appends rows to Excel, and updates history.

## Output Files

- Excel library: `.agents/skills/Ptychography_论文全量库.xlsx`
- Markdown summary: `.agents/skills/Ptychography论文解析汇总.md`
- Daily reports: `.agents/skills/daily_reports/YYYY-MM-DD_profile.md`
- High-quality reports: `.agents/skills/daily_reports/YYYY-MM-DD_profile_high_quality.md`
- History file: `.agents/skills/ptychography-paper-tracker/paper_history.json`
- Profile config: `.agents/skills/ptychography-paper-tracker/research_profiles.json`
- Journal metrics example: `.agents/skills/ptychography-paper-tracker/journal_metrics.example.csv`

## High-Quality Mode

`--mode high_quality` treats Google Scholar as a discovery source only. It then enriches candidate papers through Semantic Scholar, OpenAlex, arXiv, and Crossref, matches local journal metrics from `journal_metrics.csv`, calculates final recommendation scores, archives all scored papers, and pushes only A+/A papers.

Journal metrics are local and user-provided. Copy `journal_metrics.example.csv` to `journal_metrics.csv` and fill:

```csv
journal_name,issn,eissn,impact_factor,impact_factor_year,jcr_quartile,cas_quartile,category
```

## Notes For Codex

- Do not hard-code API keys. Prefer command-line arguments or environment variables.
- Use `--dry_run` when testing a new profile, changing keywords, or checking network results.
- Use `--doctor` before debugging failures; it checks Python packages, profile JSON, archive writability, and push keys.
- If Google Scholar mode is requested, ensure `google-search-results` is installed and pass `--serp_api_key`.
- For Server Chan Turbo push, set `SCT_KEY` in the environment and use `--notify serverchan`.
- Treat `parse_paper_full_fields()` output as a prompt handoff unless a real model integration is added later.
- If unrelated papers appear, tune `include_keywords`, `must_have_any`, `exclude_keywords`, `research_focus`, and `score_rules.min_score` in `research_profiles.json`.
