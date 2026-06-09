# 期末复习AI导师 — Claude Code Skill

一个帮助大学生从零掌握课程知识并通过期末考试的 AI 导师 Skill。

## 特点

- 🧑‍🏫 **私人家教模式** — 不是总结器，是耐心讲解每个概念的导师
- 📐 **理科友好** — 自动补全 PPT 提取中缺失的 LaTeX 公式，MathJax 渲染
- 🌐 **浏览器双屏** — 公式在浏览器完美渲染，终端只做简短交互
- 📝 **分阶段练习** — 基础巩固 → 综合应用 → 真题模拟
- 🔄 **热更新** — Q&A 实时追加到讲义 HTML，自动刷新浏览器

## 安装

1. 将 `final-exam-review.md` 放到项目的 `.claude/skills/` 目录下
2. 将 `extract_document.py` 放到同一目录
3. 安装依赖：`pip install unstructured`
4. 重启 Claude Code，skill 自动加载

## 使用

```
/期末复习
```

然后按提示对话即可。把课程讲义（PPT/PDF/Word/图片）放到项目目录下，skill 会自动识别和提取。

## 文件结构

```
你的项目/
├── .claude/skills/
│   ├── final-exam-review.md    # Skill 定义
│   └── extract_document.py     # 文档提取脚本
├── 高等数学/                    # 放你的PPT/PDF等
│   ├── A-第一章 xxx/
│   └── ...
└── _extracted/                 # 自动生成的缓存和讲义HTML
    └── 高等数学/
        └── ...
```

## 支持的文档格式

PDF、DOCX、PPTX、图片（自动OCR）、HTML、Markdown、TXT、CSV、EPUB

## 免责声明

本 Skill 仅供学习辅助，课程资料版权归原作者所有，请勿将讲义文件上传至公开仓库。
