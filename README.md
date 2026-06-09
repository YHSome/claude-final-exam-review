# 期末复习AI导师 — Claude Code Skill

一个帮你从零掌握课程知识的 AI 家教。把讲义扔进去，它会像私人家教一样，在浏览器里一步一步给你讲解知识点、出题练习、批改答疑。

## 📦 你需要准备什么

在开始之前，你只需要电脑上有两样东西：

1. **Claude Code**（AI 助手软件）
2. 你上课的**讲义文件**（PPT、PDF、Word、甚至手机拍的板书照片都行）

> 不会装 Claude Code？往下看👇

---

## 🚀 第一步：安装 Claude Code

1. 打开浏览器，访问 **https://claude.ai/code**
2. 下载对应你系统的版本（Windows/Mac）
3. 安装并登录你的 Anthropic 账号
4. 安装完成后，打开终端（Windows 按 `Win+R`，输入 `cmd` 回车），输入 `claude` 回车，看到欢迎界面就成功了

> ⚠️ Claude Code 目前需要付费 API 额度。如果你是学生，可以关注 Anthropic 的学生优惠。

---

## 📥 第二步：下载这个 Skill

> 如果你会 git：`git clone https://github.com/YHSome/claude-final-exam-review.git`

如果不会 git（或者上面的命令看不懂），跟着下面做：

1. 打开 https://github.com/YHSome/claude-final-exam-review
2. 点击绿色的 **Code** 按钮 → 选择 **Download ZIP**
3. 把下载的 ZIP 解压到你喜欢的文件夹（比如桌面上的 `期末复习` 文件夹）
4. 解压后你会看到这些文件：

```
期末复习/
├── README.md                      ← 你现在看的这个文件
├── stop-server.cmd                ← 双击可以关闭后台服务
├── .gitignore
└── .claude/
    └── skills/
        ├── final-exam-review.md   ← AI 家教的核心大脑
        └── extract_document.py    ← 用来读取讲义的脚本
```

---

## 📚 第三步：放入你的讲义

把你上课的讲义文件放到 `期末复习` 文件夹下面。建议按科目分文件夹，比如：

```
期末复习/
├── 高等数学/
│   ├── A-第一章 函数与极限/
│   │   ├── 1-1 映射与函数.ppt
│   │   └── 1-2 数列极限.ppt
│   └── B-第二章 导数与微分/
│       └── ...
├── 线性代数/
│   └── ...
└── 大学物理/
    └── ...
```

**支持的格式**：PPT (.ppt/.pptx)、PDF、Word (.doc/.docx)、图片（会自动 OCR 识别文字）、HTML、Markdown

> 💡 **提示**：文件名建议用"章节号-节名"格式，比如 `1-3 函数极限.ppt`，这样 AI 能更好地组织内容。

---

## 🎓 第四步：开始复习！

1. 打开终端（`Win+R` → `cmd` → 回车）
2. 进入你的复习文件夹：
   ```
   cd C:\Users\你的用户名\Desktop\期末复习
   ```
3. 启动 Claude Code：
   ```
   claude
   ```
4. 输入以下命令激活家教模式：
   ```
   /final-exam-review
   ```
5. 然后像聊天一样告诉它你想复习什么：
   ```
   我要复习高等数学，从第一章开始
   ```

**AI 家教会自动**：
- 扫描你的讲义文件
- 在浏览器里打开一个排版精美的讲义页面（公式用 LaTeX 渲染）
- 一步一步给你讲解每个知识点
- 每讲完一节停下来问你有没有疑问
- 出练习题给你做，做完帮你批改

> 🖥️ **浏览器里看公式**：讲解内容会在浏览器中显示，公式渲染效果跟教材一样清晰。终端只用来简单对话。

---

## 🔧 常见问题

### Q: 浏览器里公式显示不出来，全是乱码？

A: 这是 MathJax 没加载成功。试试：
1. 检查网络连接（MathJax 需要从 CDN 加载一次）
2. 刷新浏览器页面（`F5`）
3. 如果还不行，确认你是通过 `http://localhost:8888/...` 打开的，不是双击 `.html` 文件

### Q: 双击了 `stop-server.cmd` 后浏览器打不开了？

A: 这个文件会关闭后台服务。如果不小心双击了，在 Claude Code 里说"重启服务器"就行。

### Q: 提示 "pip install unstructured" 失败？

A: 这是讲义提取脚本的依赖。如果你不需要自动提取 PPT/PDF 文字（比如你用的是纯文本笔记），可以不装。否则，在终端输入：
```
pip install unstructured
```

### Q: Claude Code 的终端对话是纯文字的，怎么看公式？

A: **所有公式都在浏览器里**！终端只用来简短对话（"懂了""继续""出几道题"），真正的讲解内容、公式推导、例题都在浏览器页面里。这是设计好的双屏模式。

### Q: 我没有讲义文件，能直接用吗？

A: 可以！AI 家教会用自己的知识给你讲解。但建议至少放一份目录或考试大纲，这样它知道要讲哪些内容。

---

## 📁 项目文件说明

| 文件 | 用途 |
|:---|:---|
| `.claude/skills/final-exam-review.md` | Skill 的核心定义（AI 家教的"大脑"） |
| `.claude/skills/extract_document.py` | 文档提取脚本（读取你的 PPT/PDF） |
| `stop-server.cmd` | 一键关闭后台 HTTP 服务（复习完双击即可） |

---

## 🧑‍💻 给想自己修改的人

如果你会一点点编程，可以自己改 Skill：

- 编辑 `.claude/skills/final-exam-review.md` 来调整 AI 的行为
- 比如：让它更啰嗦/更简洁、调整题目难度、增加新的科目支持
- 改完后在 Claude Code 里说 `reload` 就能生效

---

## 👨‍💻 面向计算机专业学生（简洁版）

```bash
# 1. 克隆仓库
git clone https://github.com/YHSome/claude-final-exam-review.git
cd claude-final-exam-review

# 2. 放入讲义（支持 ppt/pptx/pdf/docx/html/md/txt/csv + 图片OCR）
mkdir 高等数学 && # 把你的课件扔进去

# 3. 可选：安装文档提取依赖
pip install unstructured

# 4. 启动 Claude Code
claude

# 在 Claude Code 中：
/final-exam-review
我要复习高等数学第八章第四节
```

**Skill 结构**：
- `final-exam-review.md` — Skill 定义（Markdown frontmatter + 完整教学流水线）
- `extract_document.py` — 基于 `unstructured` 库的文档提取，自动识别格式、保留结构、图片 OCR
- Skill 输出所有教学内容到 `_extracted/科目/章/xxx.讲义.html`（MathJax 渲染），通过本地 HTTP 服务器 `localhost:8888` 在浏览器展示
- 追加内容（Q&A、练习、答案）用 Edit 工具直接操作 HTML，避免 bash/Python 转义 LaTeX

**技术文档**：

| 文档 | 链接 |
|:---|:---|
| Claude Code 官方文档 | https://docs.anthropic.com/en/docs/claude-code |
| Claude Code Skill 开发指南 | https://docs.anthropic.com/en/docs/claude-code/skills |
| Skill 最佳实践 | https://docs.anthropic.com/en/docs/claude-code/skills#skill-best-practices |
| Claude API / Anthropic SDK | https://docs.anthropic.com/en/api |
| MathJax 3 文档 | https://docs.mathjax.org/en/latest/ |
| Unstructured 库文档 | https://docs.unstructured.io/ |

---

## 📄 免责声明

本 Skill 仅供学习辅助。课程讲义版权归原作者所有，请勿将老师的讲义文件上传到公开仓库。
