#!/usr/bin/env python3
import json
import argparse
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

def split_long_lines(content_str, max_len=9):
    if not content_str:
        return []
    raw_lines = content_str.split('\n')
    processed = []
    for line in raw_lines:
        line = line.strip()
        if not line:
            continue
        if len(line) <= max_len:
            processed.append(line)
        else:
            # 遇到超长行，根据常见标点符号拆分
            parts = re.split(r'(，|；|、|。)', line)
            current = ""
            for i in range(0, len(parts), 2):
                chunk = parts[i]
                if i + 1 < len(parts):
                    chunk += parts[i+1]
                
                if len(current) + len(chunk) <= max_len and current != "":
                    current += chunk
                else:
                    if current:
                        processed.append(current)
                    current = chunk
            if current:
                processed.append(current)
    return processed


# 通用列配置 - 无间隙
CARD_COLS = 4
CARD_ROWS = 2
COL_WIDTH = 38     # 完美平分 A4 宽度

# 正面行配置 (细网格)
FRONT_ROW_SPAN = 29
FRONT_ROW_HEIGHT = 10

# 背面行配置 (粗网格)
BACK_ROW_SPAN = 8
BACK_ROW_HEIGHT = 36.25

COLOR_TITLE = "1A1A1A"
COLOR_AUTHOR = "595959"
COLOR_DECO = "C0392B"
COLOR_PAGE = "8C8C8C"
COLOR_CONTENT = "1A1A1A"


def setup_sheet_base(ws):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    
    ws.page_margins = PageMargins(left=0.1, right=0.1, top=0.1, bottom=0.1)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True

    for i in range(1, CARD_COLS + 1):
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTH


def setup_sheet_front(ws):
    setup_sheet_base(ws)
    total_rows = CARD_ROWS * FRONT_ROW_SPAN
    for i in range(1, total_rows + 1):
        ws.row_dimensions[i].height = FRONT_ROW_HEIGHT


def setup_sheet_back(ws):
    setup_sheet_base(ws)
    total_rows = CARD_ROWS * BACK_ROW_SPAN
    for i in range(1, total_rows + 1):
        ws.row_dimensions[i].height = BACK_ROW_HEIGHT


def get_card_position_front(index):
    col_idx = index % CARD_COLS
    row_idx = index // CARD_COLS
    col_start = col_idx + 1
    row_start = row_idx * FRONT_ROW_SPAN + 1
    return col_start, row_start


def get_card_position_back(index):
    col_idx = index % CARD_COLS
    row_idx = index // CARD_COLS
    mirrored_col_idx = (CARD_COLS - 1) - col_idx
    col_start = mirrored_col_idx + 1
    row_start = row_idx * BACK_ROW_SPAN + 1
    return col_start, row_start


def add_card_border(ws, col_start, row_start, row_span, color="8C8C8C", style="dashed"):
    border_style = Side(style=style, color=color)
    ws.cell(row=row_start, column=col_start).border = Border(top=border_style, left=border_style, right=border_style)
    ws.cell(row=row_start + row_span - 1, column=col_start).border = Border(bottom=border_style, left=border_style, right=border_style)
    for r in range(row_start + 1, row_start + row_span - 1):
        ws.cell(row=r, column=col_start).border = Border(left=border_style, right=border_style)


def get_line_layout(num_lines):
    if num_lines <= 4:
        return 22, 6  # 60pt
    elif num_lines <= 6:
        return 20, 4  # 40pt
    elif num_lines <= 8:
        return 18, 3  # 30pt
    else:
        return 16, 2  # 20pt


def fill_card_front(ws, col_start, row_start, poem):
    add_card_border(ws, col_start, row_start, FRONT_ROW_SPAN, color="8C8C8C", style="dashed")
    if not poem.get('title'):
        return

    seq_row_start = row_start
    seq_row_end = row_start + 2
    ws.merge_cells(start_row=seq_row_start, start_column=col_start, end_row=seq_row_end, end_column=col_start)
    cell_seq = ws.cell(row=seq_row_start, column=col_start)
    cell_seq.value = f"·  {poem['seq']}  ·"
    cell_seq.font = Font(name='宋体', size=12, color=COLOR_DECO, italic=True)
    cell_seq.alignment = Alignment(horizontal='center', vertical='center')

    lines = poem["content"]
    num_lines = len(lines)
    if num_lines == 0:
        return
        
    font_size, rows_per_line = get_line_layout(num_lines)

    total_content_rows = num_lines * rows_per_line
    available_rows = 26
    start_offset = (available_rows - total_content_rows) // 2

    current_row = row_start + 3 + start_offset

    for line in lines:
        cell = ws.cell(row=current_row, column=col_start)
        if rows_per_line > 1:
            ws.merge_cells(start_row=current_row, start_column=col_start,
                           end_row=current_row + rows_per_line - 1, end_column=col_start)
        cell.value = line
        cell.font = Font(name='宋体', size=font_size, color=COLOR_CONTENT, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        current_row += rows_per_line


def fill_card_back(ws, col_start, row_start, poem):
    add_card_border(ws, col_start, row_start, BACK_ROW_SPAN, color="8C8C8C", style="dashed")
    if not poem.get('title'):
        return

    seq_row = row_start
    cell_seq = ws.cell(row=seq_row, column=col_start)
    cell_seq.value = f"NO. {poem['seq']:03d}"
    cell_seq.font = Font(name='Arial', size=12, color=COLOR_DECO, bold=True)
    cell_seq.alignment = Alignment(horizontal='center', vertical='center')

    deco_row = row_start + 1
    cell_deco = ws.cell(row=deco_row, column=col_start)
    cell_deco.value = "—— ✦ ——"
    cell_deco.font = Font(name='宋体', size=14, color=COLOR_DECO)
    cell_deco.alignment = Alignment(horizontal='center', vertical='center')

    title_row = row_start + 2
    cell_title = ws.cell(row=title_row, column=col_start)
    ws.merge_cells(start_row=title_row, start_column=col_start, end_row=title_row + 2, end_column=col_start)
    cell_title.value = poem["title"]
    cell_title.font = Font(name='宋体', size=24, color=COLOR_TITLE, bold=True)
    cell_title.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    author_row = row_start + 5
    cell_author = ws.cell(row=author_row, column=col_start)
    cell_author.value = poem["author"]
    cell_author.font = Font(name='宋体', size=16, color=COLOR_AUTHOR)
    cell_author.alignment = Alignment(horizontal='center', vertical='center')

    deco2_row = row_start + 6
    cell_deco2 = ws.cell(row=deco2_row, column=col_start)
    cell_deco2.value = "— —"
    cell_deco2.font = Font(name='宋体', size=12, color=COLOR_DECO)
    cell_deco2.alignment = Alignment(horizontal='center', vertical='center')

    page_row = row_start + 7
    cell_page = ws.cell(row=page_row, column=col_start)
    cell_page.value = f"《课本》{poem['page']}"
    cell_page.font = Font(name='宋体', size=12, color=COLOR_PAGE)
    cell_page.alignment = Alignment(horizontal='center', vertical='center')


def main():
    parser = argparse.ArgumentParser(description="生成诗词卡片 Excel 模板")
    parser.add_argument("start", type=int, help="起始序号")
    parser.add_argument("end", type=int, help="结束序号")
    args = parser.parse_args()

    if args.start < 1 or args.end < args.start:
        print("错误：范围无效")
        return
    
    count = args.end - args.start + 1
    if count % 8 != 0:
        print(f"错误：生成的卡片数量必须是 8 的倍数，当前范围包含 {count} 首。")
        return

    try:
        with open("poems_index.json", "r", encoding="utf-8") as f:
            all_poems = json.load(f)
    except Exception as e:
        print(f"读取 poems_index.json 失败: {e}")
        return
    
    poems_to_generate = [p for p in all_poems if args.start <= p["seq"] <= args.end]
    
    if len(poems_to_generate) < count:
        print(f"警告：JSON中找到的诗词数量 ({len(poems_to_generate)}) 少于请求数量 ({count})，将用空白卡片补齐。")

    # 补齐到 8 的倍数
    while len(poems_to_generate) % 8 != 0 or len(poems_to_generate) < count:
        poems_to_generate.append({"seq": 0, "title": "", "author": "", "page": "", "content": ""})

    wb = Workbook()
    wb.remove(wb.active) 

    chunks = [poems_to_generate[i:i+8] for i in range(0, len(poems_to_generate), 8)]
    
    for chunk in chunks:
        # 获取真实的起始和结束序号（忽略用来占位的空卡片）
        valid_seqs = [p['seq'] for p in chunk if p['seq'] != 0]
        if not valid_seqs:
            continue
        c_start = valid_seqs[0]
        c_end = valid_seqs[-1]
        
        ws_front = wb.create_sheet(f"正面({c_start}-{c_end})")
        setup_sheet_front(ws_front)
        for i, poem in enumerate(chunk):
            col, row = get_card_position_front(i)
            poem_copy = poem.copy()
            poem_copy['content'] = split_long_lines(poem.get('content', ''))
            fill_card_front(ws_front, col, row, poem_copy)

        ws_back = wb.create_sheet(f"背面({c_start}-{c_end})")
        setup_sheet_back(ws_back)
        for i, poem in enumerate(chunk):
            col, row = get_card_position_back(i)
            fill_card_back(ws_back, col, row, poem)

    out_file = f"poem_cards_{args.start}-{args.end}.xlsx"
    wb.save(out_file)
    print(f"✓ 成功生成: {out_file} (共 {len(chunks)} 页双面A4)")

if __name__ == "__main__":
    main()