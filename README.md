# Ptychography Paper Tracker

一个可配置的文献追踪工具，用于检索、筛选和归档 ptychography、4D-STEM、电子显微、X-ray ptychography 以及医学影像 AI 等方向的论文。项目支持 arXiv、Semantic Scholar、Google Scholar/SerpApi，能够生成中文日报、追加 Excel 文献库，并按研究方向做相关性评分。

## 主要功能

- 多研究方向配置：在 `research_profiles.json` 中维护关键词、排除词、arXiv 分类、评分规则和解析重点。
- 每日文献雷达：按最近 N 天检索论文，筛选新增论文，生成中文 Markdown 报告。
- 年度综述模式：按年份检索并汇总论文，适合做阶段性调研。
- Excel 归档：把论文标题、链接、摘要翻译、创新点、实验结果、未来展望等字段追加到工作簿。
- 去重历史：基于 DOI、arXiv ID、Semantic Scholar ID 和标题哈希减少重复记录。
- 推送提醒：支持 Server 酱 Turbo，把每日摘要推送到微信。
- SQLite 数据层：正式运行结果会写入 `.agents/skills/research_radar.db`，供网页读取。
- Web 可视化：第一版提供 Dashboard、文献库、文献详情、研究方向配置和设置页。

## 快速开始

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --doctor
```

列出内置研究方向：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --list_profiles
```

预览最近 7 天电子 ptychography 方向的结果，不写入文件：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 7 --max_results 20 --dry_run
```

正式运行并写入 Excel、日报和历史记录：

```powershell
python .agents\skills\ptychography-paper-tracker\scripts\ptychography_tracker.py --profile electron_ptychography --mode daily --time_range_days 1 --max_results 20
```

启动 Web 界面：

```powershell
pip install -r requirements.txt
python -m uvicorn web.web_app:app --reload --host 127.0.0.1 --port 8000
```

然后打开：

```text
http://127.0.0.1:8000
```

更多命令、配置方式和 GitHub 公开发布建议见 [中文使用教程](docs/USAGE_ZH.md)。

## 内置研究方向

- `electron_ptychography`：电子显微、4D-STEM、电子 ptychography、相位恢复、WDD/SSB 等。
- `xray_ptychography`：X-ray ptychography、CDI、同步辐射、纳米成像。
- `medical_ai`：医学影像 AI、诊断、分割、临床验证。

## 公开发布前建议

发布到 GitHub 前，建议保留脚本、配置模板、技能说明和教程，忽略本地 IDE 配置、缓存、私密设置、运行生成的 Excel/日报/历史记录。仓库已经提供 `.gitignore` 作为默认公开清单。

推荐首次提交：

```powershell
git init
git add README.md docs .gitignore requirements.txt .agents\skills\ptychography-paper-tracker
git commit -m "Publish ptychography paper tracker"
```

如果你还想同时支持 Claude Code，可按需提交 `.claude\skills\ptychography-paper-tracker`，但不要提交 `.claude\settings.local.json`。

## License

公开前建议补充一个开源许可证。若希望别人自由使用和修改，常见选择是 MIT License。
