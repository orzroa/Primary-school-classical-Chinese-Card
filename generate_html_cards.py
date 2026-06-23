#!/usr/bin/env python3
import json
import re
import argparse
import sys

def get_poem_lines(content, max_chars_per_line=8):
    """Split poem content lines by newline and split long lines at punctuation to fit within max_chars_per_line."""
    raw_lines = content.strip().split('\n')
    lines = []
    for r_line in raw_lines:
        r_line = r_line.strip()
        if not r_line:
            continue
        if len(r_line) <= max_chars_per_line:
            lines.append(r_line)
        else:
            # Split by punctuation (commas, periods, semicolons, etc.)
            parts = re.split(r'([，。；？！、])', r_line)
            current = ""
            for i in range(0, len(parts), 2):
                part = parts[i]
                punctuation = parts[i+1] if i+1 < len(parts) else ""
                chunk = part + punctuation
                if not chunk:
                    continue
                if len(current) + len(chunk) <= max_chars_per_line:
                    current += chunk
                else:
                    if current:
                        lines.append(current)
                    current = chunk
            if current:
                lines.append(current)
    return lines

def format_poem(content):
    """Dynamically determine the best line-split limit and font size to fit the card dimensions."""
    # Test layout steps to see which font size and line split count fits best within 320pt height budget.
    layouts = [
        (8, "21pt"),
        (10, "18pt"),
        (12, "16pt"),
        (14, "14pt"),
        (16, "12pt"),
        (20, "10pt"),
        (24, "8pt")
    ]
    for limit, font_size in layouts:
        lines = get_poem_lines(content, limit)
        fs_pt = float(font_size.replace("pt", ""))
        line_h = fs_pt * 1.4 + 3  # 1.4 line-height + 3pt line gap
        total_h = len(lines) * line_h
        if total_h <= 320:  # Card height is 14.85cm ≈ 387pt, content space ≈ 350pt
            return lines, font_size
    return get_poem_lines(content, 24), "8pt"

def format_author(author_str):
    """Format dynasty and author string standardizing to '[朝代] 作者' format."""
    if not author_str:
        return ""
    if author_str.startswith('《') and author_str.endswith('》'):
        return author_str
    if author_str == "汉乐府":
        return "[汉] 乐府"
    if '·' in author_str:
        parts = author_str.rsplit('·', 1)
        return f"[{parts[0]}] {parts[1]}"
    return author_str

def format_title_size(title):
    """Adjust font size dynamically for long titles to look balanced."""
    length = len(title)
    if length <= 6:
        return "25pt"
    elif length <= 10:
        return "21pt"
    else:
        return "18pt"

def generate_html(poems, output_path):
    # CSS template matching perfectly (3x2 vertical grid layout)
    css_content = """  /* 全局重置与满幅打印优化 */
  @page {
    size: A4 portrait;
    margin: 0;
  }
  * {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
  }
  body {
    background-color: #f5f5f5;
    font-family: "SimSun", "Songti SC", "Noto Serif CJK SC", "Source Han Serif SC", serif;
    color: #1a1a1a;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  /* 非打印区域样式（仅在浏览器中显示） */
  .no-print {
    max-width: 900px;
    margin: 20px auto;
    padding: 20px;
    background: #ffffff;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    font-family: system-ui, -apple-system, sans-serif;
  }
  .no-print h1 {
    font-size: 20px;
    margin-bottom: 10px;
    color: #2c3e50;
  }
  .no-print ol {
    margin-left: 20px;
    line-height: 1.6;
    color: #555;
  }
  .no-print .btn {
    display: inline-block;
    margin-top: 15px;
    padding: 10px 20px;
    background: #555;
    color: #fff;
    text-decoration: none;
    border-radius: 4px;
    font-weight: bold;
    cursor: pointer;
    border: none;
  }
  .no-print .btn:hover {
    background: #333;
  }

  /* 打印页面容器 (完全对应 A4 纵向尺寸：21.0cm x 29.7cm) */
  .page {
    width: 21.0cm;
    height: 29.7cm;
    background: #ffffff;
    position: relative;
    page-break-after: always;
    overflow: hidden;
    display: grid;
    grid-template-columns: repeat(3, 7.0cm);
    grid-template-rows: repeat(2, 14.85cm);
    /* 在屏幕显示时加上阴影，方便预览 */
    margin: 30px auto;
    box-shadow: 0 0 15px rgba(0,0,0,0.15);
  }

  /* ── 卡片外框定义 (3x2 竖卡布局，7.0cm x 14.85cm) ── */
  .card {
    width: 7.0cm;
    height: 14.85cm;
    position: relative;
    border-right: 1px dashed #999;
    border-bottom: 1px dashed #999;
  }
  /* 最后一列去掉右边框 */
  .card:nth-child(3n) {
    border-right: none;
  }
  /* 第二行去掉下边框 */
  .card:nth-child(n+4) {
    border-bottom: none;
  }

  /* ── 卡片正反面样式 ── */
  .card-front {
    width: 100%;
    height: 100%;
    padding: 0.6cm;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    align-items: center;
    position: relative;
  }
  /* 正面装饰性灰色框 */
  .card-front::before {
    content: "";
    position: absolute;
    top: 0.3cm;
    bottom: 0.3cm;
    left: 0.3cm;
    right: 0.3cm;
    border: 1px solid #ccc;
    pointer-events: none;
  }

  .card-back {
    width: 100%;
    height: 100%;
    padding: 0.6cm;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    align-items: center;
    position: relative;
  }
  /* 背面经典双线内框 */
  .card-back::before {
    content: "";
    position: absolute;
    top: 0.3cm;
    bottom: 0.3cm;
    left: 0.3cm;
    right: 0.3cm;
    border: 3px double #555;
    pointer-events: none;
  }

  /* 手动双面打印：背面所有卡片文字方向保持正向，无需旋转 */

  /* ── 诗文排版 ── */
  .seq-no {
    font-size: 13pt;
    color: #555;
    font-style: italic;
    font-weight: bold;
    height: 1.0cm;
    display: flex;
    align-items: center;
  }
  .poem-body {
    flex-grow: 1;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    width: 100%;
  }
  .poem-line {
    font-weight: bold;
    color: #111;
    line-height: 1.4;
    text-align: center;
    margin: 3px 0;
    white-space: nowrap;
    letter-spacing: 0.05em;
  }

  /* ── 背面信息排版 ── */
  .back-seq {
    font-size: 11pt;
    font-weight: bold;
    color: #1a1a1a;
    font-family: Arial, sans-serif;
    letter-spacing: 0.05em;
  }
  .back-deco {
    font-size: 12pt;
    color: #555;
    margin-top: 2px;
  }
  .back-title {
    font-weight: bold;
    color: #000;
    text-align: center;
    line-height: 1.3;
    max-width: 90%;
    word-break: break-all;
  }
  .back-author {
    font-size: 15pt;
    color: #555;
    font-weight: normal;
  }
  .back-page {
    font-size: 11pt;
    color: #777;
    font-weight: normal;
  }

  /* 打印时隐藏预览辅助组件 */
  @media print {
    body {
      background-color: #fff;
    }
    .no-print {
      display: none !important;
    }
    .page {
      margin: 0 !important;
      box-shadow: none !important;
    }
  }
"""

    html = []
    html.append("<!DOCTYPE html>")
    html.append('<html lang="zh-CN">')
    html.append('<head>')
    html.append('  <meta charset="UTF-8">')
    html.append('  <title>小学古诗文卡片 - 批量打印排版</title>')
    html.append('  <style>')
    html.append(css_content)
    html.append('  </style>')
    html.append('</head>')
    html.append('<body>')

    # Add no-print instructions
    html.append('  <div class="no-print">')
    html.append('    <h1>小学古诗文卡片 (HTML 纵向 A4 手动双面对齐版)</h1>')
    html.append('    <ol>')
    html.append('      <li><b>排版已更新为 3×2 全竖卡布局</b>：每页包含 6 张等尺寸竖置卡片（宽 7.0cm × 高 14.85cm），无缝利用整张 A4 纸。</li>')
    html.append('      <li><b>背面已完成左右翻转对齐</b>：背面卡片顺序调整为 3、2、1、6、5、4，且文字方向保持正常，以完美契合左右翻面（长边翻转）手动双面打印。</li>')
    html.append('      <li><b>打印步骤</b>：点击下方的“直接打印PDF”按钮或按快捷键 <code>Ctrl + P</code>。</li>')
    html.append('      <li><b>重要打印设置</b>：')
    html.append('        <ul>')
    html.append('          <li>方向：<b>纵向 (Portrait)</b></li>')
    html.append('          <li>纸张大小：<b>A4</b></li>')
    html.append('          <li>页边距：必须设置为<b>“无” (None)</b> 以获得最佳比例。</li>')
    html.append('          <li>缩放：<b>实际大小 / 默认 (100% / Default)</b></li>')
    html.append('          <li>背景图形 (Background graphics)：<b>勾选</b></li>')
    html.append('        </ul>')
    html.append('      </li>')
    html.append('    </ol>')
    html.append('    <button class="btn" onclick="window.print()">直接打印 / 另存为 PDF</button>')
    html.append('  </div>')

    # Chunk into 6-card pages
    chunks = [poems[i:i+6] for i in range(0, len(poems), 6)]

    for page_idx, chunk in enumerate(chunks, 1):
        # ── FRONT PAGE ──
        html.append(f'  <!-- ── 页面 {page_idx}：正面（诗词内容） ── -->')
        html.append('  <div class="page page-front">')
        for i in range(6):
            if i < len(chunk) and chunk[i].get('seq', 0) > 0:
                p = chunk[i]
                lines, fs = format_poem(p['content'])
                lines_html = "".join([f'<div class="poem-line" style="font-size: {fs};">{line}</div>' for line in lines])
                html.append(f'    <div class="card">')
                html.append(f'      <div class="card-front">')
                html.append(f'        <div class="seq-no">· {p["seq"]} ·</div>')
                html.append(f'        <div class="poem-body">')
                html.append(f'          {lines_html}')
                html.append(f'        </div>')
                html.append(f'      </div>')
                html.append(f'    </div>')
            else:
                html.append(f'    <div class="card">')
                html.append(f'      <div class="card-front"></div>')
                html.append(f'    </div>')
        html.append('  </div>')

        # ── BACK PAGE ──
        html.append(f'  <!-- ── 页面 {page_idx}：背面（题目信息） ── -->')
        html.append('  <div class="page page-back">')
        # Back cards order is: 2, 1, 0, 5, 4, 3 (which corresponds to 3, 2, 1, 6, 5, 4 in 1-based indexing)
        for i in [2, 1, 0, 5, 4, 3]:
            if i < len(chunk) and chunk[i].get('seq', 0) > 0:
                p = chunk[i]
                title_fs = format_title_size(p['title'])
                author_formatted = format_author(p['author'])
                html.append(f'    <div class="card">')
                html.append(f'      <div class="card-back">')
                html.append(f'        <div class="back-seq">NO. {p["seq"]:03d}</div>')
                html.append(f'        <div class="back-deco">—— ✦ ——</div>')
                html.append(f'        <div class="back-title" style="font-size: {title_fs};">{p["title"]}</div>')
                html.append(f'        <div class="back-author">{author_formatted}</div>')
                html.append(f'        <div class="back-deco">— —</div>')
                html.append(f'        <div class="back-page">《课本》{p["page"]}</div>')
                html.append(f'      </div>')
                html.append(f'    </div>')
            else:
                html.append(f'    <div class="card">')
                html.append(f'      <div class="card-back"></div>')
                html.append(f'    </div>')
        html.append('  </div>')

    html.append('</body>')
    html.append('</html>')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(html))

def main():
    parser = argparse.ArgumentParser(description="批量生成诗词卡片 HTML 打印模板")
    parser.add_argument("--start", type=int, default=1, help="起始诗词序号 (1-200)")
    parser.add_argument("--end", type=int, default=200, help="结束诗词序号 (1-200)")
    parser.add_argument("--output", type=str, default="poem_cards_final.html", help="输出HTML路径")
    args = parser.parse_args()

    try:
        with open("poems_index.json", "r", encoding="utf-8") as f:
            all_poems = json.load(f)
    except Exception as e:
        print(f"读取 poems_index.json 失败: {e}", file=sys.stderr)
        sys.exit(1)

    # Filter poems by sequence range
    selected_poems = [p for p in all_poems if args.start <= p.get('seq', 0) <= args.end]
    
    # Pad selected poems with blanks to be a multiple of 6
    pad_count = (6 - (len(selected_poems) % 6)) % 6
    for i in range(pad_count):
        selected_poems.append({"seq": 0, "title": "", "author": "", "page": "", "content": ""})

    generate_html(selected_poems, args.output)
    print(f"✓ 成功生成: {args.output} (共 {len(selected_poems)} 首诗词，共 {len(selected_poems)//6} 张 A4 纸双面)")

if __name__ == "__main__":
    main()
