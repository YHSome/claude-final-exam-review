"""
通用文档文本提取脚本
使用 unstructured 库自动识别文件类型并提取文本内容。
支持 PDF、DOCX、PPTX、图片(PNG/JPG/TIFF)、HTML、MD、TXT、CSV、EPUB 等格式。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
核心设计：提取后存盘，不直接输出全文到聊天区
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- 默认模式：提取并缓存全文，只输出结构概览（章节标题、篇幅）
- 缓存目录镜像源文件结构：科目 → 章 → 节
  例如: 高等数学/A-第一章 函数与极限/1-1 映射与函数.ppt
    → _extracted/高等数学/A-第一章 函数与极限/1-1 映射与函数.ppt.txt
- 概览足够让 AI 判断"资料讲了什么、哪些章节需要重点讲解"
- 讲解时按章节逐段读取，避免长文本撑爆上下文窗口
- `--list --tree` 可查看科目→章→节的树状结构

依赖安装：
    pip install unstructured
    # 可选扩展：
    # pip install unstructured[pdf]     # PDF图片/扫描件OCR
    # pip install unstructured[image]   # 图片OCR
    # pip install unstructured[docx]    # Word文档
    # pip install unstructured[pptx]    # PowerPoint

用法：
    python extract_document.py <file1> [<file2> ...]           # 提取文件
    python extract_document.py --list                          # 列出已缓存文件(平铺)
    python extract_document.py --list --tree                   # 树状结构: 科目→章→节
    python extract_document.py --read <缓存键>                  # 读取全文
    python extract_document.py --read <缓存键> --from 10 --to 60  # 按行号读取
    python extract_document.py --subject 高等数学 --chapter 第一章  # 读取某章所有节
"""

import sys
import os
import re
import json
from pathlib import Path

# 确保 stdout/stderr 支持 Unicode（Windows GBK 终端兼容）
if sys.platform == "win32":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# 项目根目录 = 此脚本的上上级（.claude/skills/ → .claude/ → 项目根）
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
CACHE_DIR = PROJECT_ROOT / "_extracted"
CACHE_INDEX = CACHE_DIR / "index.json"


# ─── 缓存基础操作 ───────────────────────────────────────────

def ensure_cache_dir():
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if not CACHE_INDEX.exists():
        CACHE_INDEX.write_text("{}", encoding="utf-8")


def load_index():
    ensure_cache_dir()
    with open(CACHE_INDEX, "r", encoding="utf-8") as f:
        return json.load(f)


def save_index(idx):
    ensure_cache_dir()
    with open(CACHE_INDEX, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=2)


def relpath(path):
    """将绝对/相对路径转为相对于项目根目录的路径"""
    abs_path = os.path.abspath(path)
    try:
        return os.path.relpath(abs_path, PROJECT_ROOT)
    except ValueError:
        return abs_path


# ─── 文档提取 ───────────────────────────────────────────────

def _is_old_ppt(file_path):
    """检测是否为旧版 .ppt 格式 (OLE2/CFB)"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext != ".ppt":
        return False
    try:
        with open(file_path, "rb") as f:
            return f.read(4) == b"\xd0\xcf\x11\xe0"
    except Exception:
        return False


def _extract_old_ppt_text(file_path):
    """
    从旧版 .ppt (OLE2) 中提取文本。
    优先使用 PowerPoint COM 将 .ppt 转为 .pptx，再用 unstructured 原生解析，
    可完整捕获文本、公式和结构。COM 不可用时降级为二进制记录解析。
    """
    # 尝试 COM 转 .pptx + unstructured 原生解析（最佳质量）
    try:
        return _extract_via_com_convert(file_path)
    except Exception as e:
        print(f"   ⚠️  COM转换失败 ({e})，降级为记录解析...", file=sys.stderr)

    # 降级：二进制记录解析（纯文本）
    return _extract_via_record_parser(file_path)


def _extract_via_com_convert(file_path):
    """
    使用 PowerPoint COM 将旧 .ppt 另存为 .pptx，
    然后用 unstructured 原生解析 .pptx，精准提取文本、公式和结构。
    """
    import tempfile
    import win32com.client

    abs_path = os.path.abspath(file_path)
    tmp_dir = tempfile.mkdtemp(prefix="ppt_convert_")

    try:
        # Step 1: COM 打开 .ppt，另存为 .pptx
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            ppt_app.Visible = False
        except Exception:
            pass

        pptx_path = os.path.join(tmp_dir, "converted.pptx")
        presentation = ppt_app.Presentations.Open(abs_path, WithWindow=False)
        # 24 = ppSaveAsOpenXMLPresentation (.pptx), 32 = ppSaveAsDefault
        try:
            presentation.SaveAs(pptx_path, 24)  # .pptx
        except Exception:
            presentation.SaveAs(pptx_path, 32)  # fallback

        presentation.Close()
        ppt_app.Quit()

        # Step 2: 用 unstructured 原生解析 .pptx（精确提取文本、公式、表格）
        from unstructured.partition.pptx import partition_pptx

        elements = partition_pptx(filename=pptx_path)
        text_parts = []
        titles = []

        for element in elements:
            element_type = type(element).__name__
            text = str(element).strip()
            if not text:
                continue

            if element_type == "Title":
                text_parts.append(f"## {text}")
                titles.append({"text": text, "line": len(text_parts)})
            elif element_type == "Table":
                text_parts.append(f"[表格]\n{text}")
            elif element_type == "ListItem":
                text_parts.append(f"• {text}")
            else:
                text_parts.append(text)

        full_text = "\n\n".join(text_parts)

        if len(full_text.strip()) < 20:
            return "[旧版PPT: 转换后仍未提取到有效文本。]", []

        return full_text, titles

    finally:
        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)
        try:
            ppt_app.Quit()
        except Exception:
            pass


def _extract_via_record_parser(file_path):
    """
    降级方案：递归解析 PPT 二进制记录 (TextCharsAtom/TextBytesAtom)。
    只能提取纯文本，无法捕获公式和复杂排版。
    """
    import olefile
    import struct

    TEXT_CHARS_ATOM = 0x0FA0
    TEXT_BYTES_ATOM = 0x0FA8
    CONTAINER_MASK = 0x0F

    def parse_records(data, offset, length, texts):
        pos = offset
        end = offset + length
        while pos + 8 <= end:
            b0 = data[pos]
            rec_ver = b0 & CONTAINER_MASK
            rec_type = struct.unpack_from("<H", data, pos + 2)[0]
            rec_len = struct.unpack_from("<I", data, pos + 4)[0]
            body_start = pos + 8
            body_end = min(body_start + rec_len, end)

            if rec_type == TEXT_CHARS_ATOM:
                try:
                    text = data[body_start:body_end].decode("utf-16-le", errors="ignore")
                    text = text.strip().replace("\x00", "")
                    if text and len(text) >= 2:
                        texts.append(text)
                except Exception:
                    pass
            elif rec_type == TEXT_BYTES_ATOM:
                try:
                    text = data[body_start:body_end].decode("latin-1", errors="ignore")
                    text = text.strip()
                    if text and len(text) >= 2:
                        texts.append(text)
                except Exception:
                    pass
            elif rec_ver == CONTAINER_MASK and rec_len > 0:
                parse_records(data, body_start, rec_len, texts)

            pos += 8 + rec_len

    ole = olefile.OleFileIO(file_path)
    all_texts = []

    for stream_path_parts in ole.listdir():
        try:
            data = ole.openstream("/".join(stream_path_parts)).read()
        except Exception:
            continue
        if len(data) >= 8:
            parse_records(data, 0, len(data), all_texts)

    ole.close()

    if not all_texts:
        return (
            f"[旧版PPT: 未能提取到文本。建议将 .ppt 转为 .pptx 后重试。]\n"
            f"文件: {file_path}"
        ), []

    seen = set()
    unique = [t for t in all_texts if not (t in seen or seen.add(t))]

    def is_garbage(text):
        if len(text) < 2:
            return True
        if text.isdigit() or text.replace(",", "").replace(".", "").isdigit():
            return len(text) < 6
        weird = sum(1 for c in text if not c.isprintable() or ("一" > c > "\x7f"))
        return weird > len(text) * 0.5

    filtered = [t for t in unique if not is_garbage(t)]

    if not filtered:
        return f"[旧版PPT: 提取文本均为元数据，无有效内容。]", []

    titles = []
    final_parts = []
    for t in filtered:
        is_title = len(t) <= 40 and t[-1] not in "，。、；：？！,.;:?!）)"
        if is_title:
            final_parts.append(f"## {t}")
            titles.append({"text": t, "line": len(final_parts)})
        else:
            final_parts.append(t)

    return "\n\n".join(final_parts), titles


def extract_document_text(file_path):
    """使用 unstructured 提取文档。旧 .ppt 自动走 olefile 降级路径。"""
    # 旧 .ppt → 用 olefile 降级提取，避免 unstructured 崩溃
    if _is_old_ppt(file_path):
        return _extract_old_ppt_text(file_path)

    from unstructured.partition.auto import partition

    elements = partition(filename=file_path)
    text_parts = []
    titles = []

    for element in elements:
        element_type = type(element).__name__
        text = str(element).strip()
        if not text:
            continue

        if element_type == "Title":
            text_parts.append(f"## {text}")
            titles.append({"text": text, "line": len(text_parts)})
        elif element_type == "Table":
            text_parts.append(f"[表格]\n{text}")
        elif element_type == "ListItem":
            text_parts.append(f"• {text}")
        else:
            text_parts.append(text)

    return "\n\n".join(text_parts), titles


def parse_hierarchy(rel_path):
    """
    从相对路径中解析 科目 / 章 / 节。
    假设目录结构为: 科目/章目录/节文件
    例如: 高等数学/A-第一章 函数与极限/1-1 映射与函数.ppt
      → subject=高等数学, chapter=A-第一章 函数与极限, section=1-1 映射与函数
    """
    parts = Path(rel_path).parts
    if len(parts) >= 3:
        return {"subject": parts[0], "chapter": parts[1], "section": parts[2]}
    elif len(parts) == 2:
        return {"subject": parts[0], "chapter": parts[1], "section": None}
    elif len(parts) == 1:
        return {"subject": parts[0], "chapter": None, "section": None}
    return {"subject": rel_path, "chapter": None, "section": None}


def build_overview(text, titles, source_path, cache_key):
    """根据提取内容构建结构概览（这是输出到聊天区的部分）"""
    lines = text.split("\n")
    total_chars = len(text)
    total_lines = len(lines)
    hierarchy = parse_hierarchy(cache_key)

    overview = []
    overview.append(f"📄 {source_path}")

    # 显示层级归属
    loc = []
    if hierarchy["subject"]:
        loc.append(f"科目: {hierarchy['subject']}")
    if hierarchy["chapter"]:
        loc.append(f"章: {hierarchy['chapter']}")
    if hierarchy["section"]:
        loc.append(f"节: {hierarchy['section']}")
    if loc:
        overview.append(f"   {' | '.join(loc)}")

    overview.append(f"   缓存键: {cache_key}")
    overview.append(f"   总篇幅: {total_chars} 字符 / {total_lines} 行")

    if titles:
        overview.append(f"\n📑 检测到的标题 ({len(titles)} 个):")
        for i, t in enumerate(titles, 1):
            preview = t["text"][:80] + ("..." if len(t["text"]) > 80 else "")
            overview.append(f"   {i:3d}. [行{t['line']:>5d}] {preview}")
    else:
        preview_lines = [l for l in lines[:10] if l.strip()][:5]
        if preview_lines:
            overview.append(f"\n📑 未检测到明显标题结构，前几行预览:")
            for l in preview_lines:
                overview.append(f"   · {l[:100]}")

    est_tokens = total_chars // 3
    overview.append(f"\n⚠️  全文约 {est_tokens} tokens，已缓存到磁盘。")
    overview.append(f"   讲解时用 --read 按章节逐段读取。")

    return "\n".join(overview)


def do_extract(file_paths):
    """提取文件到缓存，镜像源目录结构，输出结构概览"""
    ensure_cache_dir()
    index = load_index()

    for path in file_paths:
        if not os.path.exists(path):
            print(f"❌ 文件不存在: {path}")
            continue

        # 缓存键 = 相对于项目根的路径
        cache_key = relpath(path)
        cache_path = CACHE_DIR / (cache_key + ".txt")

        try:
            print(f"⏳ 正在提取: {cache_key} ...", file=sys.stderr)
            text, titles = extract_document_text(path)

            # 确保子目录存在
            cache_path.parent.mkdir(parents=True, exist_ok=True)

            # 存盘
            cache_path.write_text(text, encoding="utf-8")

            hierarchy = parse_hierarchy(cache_key)

            # 更新索引
            index[cache_key] = {
                "source": os.path.abspath(path),
                "cache": str(cache_path),
                "chars": len(text),
                "lines": len(text.split("\n")),
                "title_count": len(titles),
                "titles": titles,
                "subject": hierarchy["subject"],
                "chapter": hierarchy["chapter"],
                "section": hierarchy["section"],
            }
            save_index(index)

            # 输出概览
            print(build_overview(text, titles, cache_key, cache_key))

        except ImportError:
            print("❌ 请先安装 unstructured: pip install unstructured")
            sys.exit(1)
        except Exception as e:
            print(f"❌ 提取失败 [{cache_key}]: {e}")


# ─── 列出缓存 ───────────────────────────────────────────────

def do_list(tree_mode=False):
    """列出已缓存的文件"""
    ensure_cache_dir()
    index = load_index()
    if not index:
        print("📭 暂无缓存文件")
        return

    print(f"📂 缓存目录: {CACHE_DIR}")
    print(f"   共 {len(index)} 个文件\n")

    if tree_mode:
        _do_list_tree(index)
    else:
        for key, info in sorted(index.items()):
            est_tokens = info["chars"] // 3
            print(f"   {key}")
            print(f"      篇幅: {info['chars']} 字符 / {info['lines']} 行 (~{est_tokens} tokens)")
            print(f"      标题: {info['title_count']} 个")
            print()


def _do_list_tree(index):
    """树状结构：科目 → 章 → 节"""
    # 构建树: subject → chapter → [(section_key, info)]
    tree = {}
    for key, info in sorted(index.items()):
        subj = info.get("subject") or "(无科目)"
        chap = info.get("chapter") or "(无章节)"
        tree.setdefault(subj, {}).setdefault(chap, []).append((key, info))

    for subject, chapters in sorted(tree.items()):
        # 统计该科目总 token 数
        total_tokens = sum(
            info["chars"] // 3
            for sections in chapters.values()
            for _, info in sections
        )
        print(f"📚 {subject}  (~{total_tokens} tokens)")

        for chapter, sections in sorted(chapters.items()):
            chap_tokens = sum(info["chars"] // 3 for _, info in sections)
            print(f"   📂 {chapter}  (~{chap_tokens} tokens)")

            for key, info in sections:
                est_tokens = info["chars"] // 3
                title_count = info["title_count"]
                print(f"      📄 {info.get('section') or key}  ({info['chars']}字, {title_count}个标题)")

            print()
        print()


# ─── 读取缓存 ───────────────────────────────────────────────

def do_read(cache_ref, from_line=None, to_line=None):
    """从缓存读取某文件的全文或指定行范围"""
    ensure_cache_dir()
    index = load_index()

    match_key = _resolve_cache_key(cache_ref, index)

    if not match_key:
        print(f"❌ 未找到缓存: {cache_ref}")
        print(f"   使用 --list 查看已缓存的文件")
        sys.exit(1)

    info = index[match_key]
    cache_path = Path(info["cache"])

    if not cache_path.exists():
        print(f"❌ 缓存文件丢失: {cache_path}")
        sys.exit(1)

    text = cache_path.read_text(encoding="utf-8")
    lines = text.split("\n")

    if from_line is not None or to_line is not None:
        start = max(0, (from_line or 1) - 1)
        end = min(len(lines), to_line or len(lines))
        selected = "\n".join(lines[start:end])
        print(f"📖 {match_key} [行 {start+1}-{end} / 共 {len(lines)} 行]")
        print(f"{'─'*60}")
        print(selected)
    else:
        print(f"📖 {match_key} [全文 / {len(lines)} 行]")
        print(f"{'─'*60}")
        print(text)


def do_read_by_subject_chapter(subject, chapter=None):
    """读取某科目/某章下所有节的全文"""
    ensure_cache_dir()
    index = load_index()

    matched = []
    for key, info in index.items():
        if info.get("subject") == subject:
            if chapter is None or info.get("chapter") == chapter:
                matched.append((key, info))

    if not matched:
        target = f"{subject}/{chapter}" if chapter else subject
        print(f"❌ 未找到匹配的缓存: {target}")
        sys.exit(1)

    matched.sort(key=lambda x: x[0])

    print(f"📚 {'科目' if not chapter else '章节'}: {subject + (f' / {chapter}' if chapter else '')}")
    print(f"   共 {len(matched)} 个文件\n")
    print(f"{'='*60}")

    for key, info in matched:
        cache_path = Path(info["cache"])
        if cache_path.exists():
            text = cache_path.read_text(encoding="utf-8")
            print(f"\n{'─'*60}")
            print(f"📄 {info.get('section') or key}")
            print(f"{'─'*60}\n")
            print(text)
        else:
            print(f"\n❌ 缓存丢失: {key}")


def _resolve_cache_key(cache_ref, index):
    """按多种方式匹配缓存键：精确匹配 / 包含匹配 / 按章节匹配"""
    # 精确匹配
    if cache_ref in index:
        return cache_ref

    # 包含匹配（支持只输入文件名或章节关键词）
    matches = [k for k in index if cache_ref in k]
    if len(matches) == 1:
        return matches[0]
    elif len(matches) > 1:
        print(f"⚠️  多个匹配: {cache_ref}")
        for m in matches:
            print(f"     {m}")
        print(f"   请使用更精确的关键词")
        return None

    return None


# ─── 入口 ────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    # --list [--tree]
    if sys.argv[1] == "--list":
        tree_mode = "--tree" in sys.argv
        do_list(tree_mode=tree_mode)
        sys.exit(0)

    # --subject X [--chapter Y]
    if sys.argv[1] == "--subject":
        if len(sys.argv) < 3:
            print("用法: python extract_document.py --subject <科目> [--chapter <章>]")
            sys.exit(1)
        subject = sys.argv[2]
        chapter = None
        for i, arg in enumerate(sys.argv):
            if arg == "--chapter" and i + 1 < len(sys.argv):
                chapter = sys.argv[i + 1]
        do_read_by_subject_chapter(subject, chapter)
        sys.exit(0)

    # --read <key> [--from N] [--to M]
    if sys.argv[1] == "--read":
        if len(sys.argv) < 3:
            print("用法: python extract_document.py --read <缓存键> [--from N] [--to M]")
            sys.exit(1)

        cache_ref = sys.argv[2]
        from_line = None
        to_line = None

        for i, arg in enumerate(sys.argv):
            if arg == "--from" and i + 1 < len(sys.argv):
                from_line = int(sys.argv[i + 1])
            if arg == "--to" and i + 1 < len(sys.argv):
                to_line = int(sys.argv[i + 1])

        do_read(cache_ref, from_line, to_line)
        sys.exit(0)

    # 默认模式：提取文件
    files = sys.argv[1:]
    do_extract(files)
