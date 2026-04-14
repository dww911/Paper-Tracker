# Paper-Tracker
领域论文自动追踪工具（谷歌学术+arXiv）
领域论文自动检索、解析、归档工具

## ✨ 功能
- 支持 **谷歌学术（SerpApi）** 精准检索领域论文
- 支持 arXiv 每日新论文追踪
- 自动去重、按年份归档到 Excel
- 自动提取：标题、作者、期刊、发表时间、链接、摘要
- 适配 Claude Skill 一键运行

## 🛠️ 安装依赖
```bash
pip install requests pandas openpyxl feedparser serpapi
