# Research Radar 中文使用教程

## MVP 怎么用（推荐路径）

```text
配置研究方向 → 获取文献（/fetch，可先试跑预览）→ 文献库筛选 → 打开原文 / 导出 Excel / 生成日报
```

- **当前方向**：由 `research_profiles.json` 的 `default_profile` 决定；首页、文献库、报告导出默认只显示该方向，不会用全库凑数。
- **试跑**：获取文献页点击「试跑预览」，或 CLI `--dry_run`（写入运行记录与检索列表，不入库）。
- **入库策略**：网页默认「按研究方向筛选」；需全量归档时选「全部入库」。
- **高级功能**：`app_settings.json` 中 `advanced_features_enabled` 默认为 `false`；在设置页开启后显示侧栏扩展入口。
- **研究方向地图**：首页卡片始终可见 → `/roadmap`（时间轴 / 分支 / 精读路线），无需开启高级功能。

---

本文档面向第一次使用或准备把项目公开到 GitHub 的用户，按“安装、配置、运行、发布”的顺序说明。

## 1. 项目结构

核心文件如下：

```text
F:\Ptychography-Paper-Tracker
├─ requirements.txt
├─ README.md
├─ .gitignore
├─ docs/
│  └─ USAGE_ZH.md
└─ .agents/
   └─ skills/
      └─ ptychography-paper-tracker/
         ├─ SKILL.md
         ├─ research_profiles.json
         ├─ paper_history.json
         └─ scripts/
            └─ ptychography_tracker.py
```

最重要的是：

- `ptychography_tracker.py`：主程序。
- `radar_db.py`：SQLite 数据层，负责建表、写入论文、运行记录和阅读笔记。
- `research_profiles.json`：研究方向配置。
- `paper_history.json`：去重历史，公开仓库时通常不提交真实运行记录。
- `requirements.txt`：Python 依赖。
- `README.md`：GitHub 首页说明。
- `web/`：FastAPI + Jinja2 可视化界面。

## 2. 安装环境

建议使用 Python 3.10 或更新版本。

在 PowerShell 中进入项目目录：

```powershell
cd F:\Ptychography-Paper-Tracker
```

创建并启用虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

安装依赖：

```powershell
pip install -r requirements.txt
```

检查环境：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --doctor
```

如果 `--doctor` 显示缺少依赖，重新运行 `pip install -r requirements.txt`。

## 3. 查看研究方向

列出当前可用方向：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --list_profiles
```

默认内置：

- `electron_ptychography`：电子显微、4D-STEM、电子 ptychography、相位恢复。
- `xray_ptychography`：X-ray ptychography、CDI、同步辐射。
- `medical_ai`：医学影像 AI、诊断、分割、临床验证。

## 4. 每日文献追踪

第一次运行建议使用 `--dry_run`，只预览，不写入 Excel 和历史记录：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 7 --max_results 20 --dry_run
```

确认结果相关后，正式运行：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20
```

正式运行会生成或更新：

- `.agents/skills/research_radar.db`（**网页主数据源**，相关论文全部 upsert）
- `.agents/skills/Ptychography_论文全量库.xlsx`（导出/备份）
- `.agents/skills/daily_reports/YYYY-MM-DD_profile.md`
- `.agents/skills/ptychography-paper-tracker/paper_history.json`（去重与 last_seen，不再阻止重复入库）

**入库与推荐分离**：通过相关性过滤的论文一律写入 SQLite；首页/微信 TOP3 仅按 `is_recommended` 与星级展示。

## 5. 年度综述

检索指定年份范围，并生成适合年度调研的候选论文列表：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode annual_summary --start_year 2024 --end_year 2026 --max_papers_per_year 50 --dry_run
```

确认结果后去掉 `--dry_run`：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode annual_summary --start_year 2024 --end_year 2026 --max_papers_per_year 50
```

## 6. Google Scholar / SerpApi

Google Scholar 模式需要 SerpApi Key。推荐通过环境变量传入，不要写入代码或提交到 GitHub。

```powershell
$env:SERPAPI_API_KEY="你的 SerpApi Key"
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode google_scholar --max_results 20 --dry_run
```

也可以显式传参：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode google_scholar --max_results 20 --serp_api_key "你的 SerpApi Key"
```

公开项目时不要把真实 Key 写进 README、脚本、配置文件或历史命令记录。

## 7. Server 酱推送

设置 Server 酱 Turbo Key：

```powershell
$env:SCT_KEY="你的 Server 酱 SendKey"
```

测试推送：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --test_notify --notify serverchan
```

每日运行并推送摘要：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20 --notify serverchan
```

## 8. 自定义研究方向

可以用交互方式创建新方向：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --create_profile
```

也可以直接编辑：

```text
.agents/skills/ptychography-paper-tracker/research_profiles.json
```

建议为每个方向配置：

- `include_keywords`：核心关键词。
- `exclude_keywords`：排除词，用于降低无关论文得分。
- `must_have_any`：至少命中一个才更可能被保留的关键词。
- `research_focus`：你的具体研究关注点。
- `arxiv_categories`：arXiv 分类范围。
- `score_rules.min_score`：最低相关性分数（正常 tier 入库门槛）。
- `ingest_min_score`：分层入库最低分；未命中 `must_have_any` 但分数 ≥ 此值时以低优先级 tier 入库（`is_recommended=0`）。
- `ingest_below_must_have`：是否启用上述低优先级入库（默认 `true`）。

调新方向时先用 `--dry_run`，避免把不相关论文写入历史和 Excel。

## 9. GitHub 公开发布清单

建议提交：

- `README.md`
- `docs/USAGE_ZH.md`
- `requirements.txt`
- `.gitignore`
- `.agents/skills/ptychography-paper-tracker/SKILL.md`
- `.agents/skills/ptychography-paper-tracker/research_profiles.json`
- `.agents/skills/ptychography-paper-tracker/scripts/ptychography_tracker.py`
- 其他确实需要公开的脚本。

建议不要提交：

- `.idea/`
- `.claude/settings.local.json`
- `__pycache__/`
- 真实运行生成的 Excel 文件。
- 真实日报和历史记录。
- 任何 API Key、SendKey、Token。

首次创建 Git 仓库可以运行：

```powershell
git init
git add README.md docs .gitignore requirements.txt .agents\skills\ptychography-paper-tracker
git commit -m "Publish ptychography paper tracker"
```

如果你要创建 GitHub 远程仓库：

```powershell
git branch -M main
git remote add origin https://github.com/<your-name>/<repo-name>.git
git push -u origin main
```

推送前建议运行：

```powershell
git status --short
```

确认没有把私密文件、缓存文件和运行结果放进暂存区。

## 10. 常见问题

### 运行时中文乱码

主脚本已尝试把标准输出切换为 UTF-8。如果终端仍乱码，可在 PowerShell 中运行：

```powershell
chcp 65001
```

### 搜不到足够论文

可以适当增加：

```powershell
--time_range_days 30 --max_results 50
```

或者放宽 `research_profiles.json` 中的 `must_have_any` 和 `score_rules.min_score`。

### 结果里混入无关论文

优先调整：

- 增加 `exclude_keywords`。
- 提高 `score_rules.min_score`。
- 收紧 `must_have_any`。
- 在 `research_focus` 中加入更具体的研究对象或方法。

### 不想写入 Excel

先使用：

```powershell
--dry_run
```

目前 `--dry_run` 是最安全的预览方式。

## 11. 推荐工作流

日常使用：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20 --dry_run
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20
```

新方向调试：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --create_profile
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile 新方向ID --mode daily --time_range_days 14 --max_results 30 --dry_run
```

年度调研：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode annual_summary --start_year 2024 --end_year 2026 --max_papers_per_year 50 --dry_run
```

## 12. Web 可视化界面

第一版 Web 使用：

```text
FastAPI + Jinja2 + SQLite
```

启动前先安装依赖：

```powershell
pip install -r requirements.txt
```

启动服务：

```powershell
python -m uvicorn web.web_app:app --reload --host 127.0.0.1 --port 8000
```

浏览器打开：

```text
http://127.0.0.1:8000
```

侧栏为 **六个主入口**（旧 URL 保留重定向）：

| 入口 | 路径 | 聚合能力 |
|------|------|----------|
| 今日工作台 | `/` | Dashboard TOP3、「查看全部今日」→ `/papers?date=today`；`/fetch`、`/runs/*`、`/daily`、`/wechat` |
| 文献库 | `/library` | `/papers` 全量列表（IF/引用/星级/tier 筛选）、`/annual/{year}`、`/reading`、`/roadmap` |
| **研究进展** | `/progress` | 文献/阅读/关键词/里程碑/写作准备度仪表盘；导出 `progress_reports/` |
| 灵感笔记 | `/ideas` | 想法 CRUD，可关联论文与导出 `idea_notes/` |
| 写作中心 | `/writing` | 日报/周报、综述、组会、引言草稿、`/citations` 引用篮 |
| 设置 | `/settings` | API Key、期刊指标、`/profiles` 研究方向（含分层入库配置）、`--doctor` 健康检查 |

常用子路径：

- `/fetch`：可控运行（篇数、写 DB/Excel/Markdown、微信推送等），完成后 `/runs/{id}` 含「下一步」按钮。
- `/papers/{id}`：原文/PDF 链接、上传 PDF 精读、`my_notes/`。
- `/progress`：个人课题推进 KPI（区别于 `/roadmap` 领域时间线）。
- `/generate` → 重定向 `/writing`。

### 数据流（Web 与 CLI 统一）

| 层级 | 职责 |
|------|------|
| SQLite | 论文主表、阅读笔记、runs、review_jobs |
| Excel / Markdown | 导出与报告产物 |
| paper_history.json | 去重键 + `last_seen`，不阻止 upsert |
| research_profiles.json | 研究方向配置唯一源 |

推荐日常流程：打开 Web → `/fetch` 运行 daily → `/papers` 浏览 → 点原文 → 详情页自动生成 AI 笔记 → 保存「我的笔记」。

如果数据库为空，请从 `/fetch` 运行 `daily` 模式（无需 SerpApi）。
