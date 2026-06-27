import argparse
import json
import math
import os
from pathlib import Path

import fitz


FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\arialuni.ttf",
]


def find_font():
    for candidate in FONT_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return None


def clean_text(value):
    return " ".join(str(value or "").replace("\u00a0", " ").split())


def text_width(font, text, size):
    try:
        return font.text_length(text, fontsize=size)
    except Exception:
        return len(text) * size * 0.55


def wrap_text(font, text, size, max_width):
    text = clean_text(text)
    if not text:
        return []

    has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in text)
    tokens = list(text) if has_cjk else []
    if not has_cjk:
        for chunk in text.split(" "):
            if chunk:
                tokens.append(chunk)
                tokens.append(" ")

    lines = []
    line = ""
    for token in tokens:
        next_line = f"{line}{token}"
        if line and text_width(font, next_line.rstrip(), size) > max_width:
            lines.append(line.rstrip())
            line = token.lstrip()
        else:
            line = next_line
    if line.strip():
        lines.append(line.rstrip())
    return lines or [text]


def draw_wrapped_text(page, font_name, font_file, font, text, rect, size, color=(0.08, 0.09, 0.10), line_gap=1.35):
    lines = wrap_text(font, text, size, max(10, rect.width))
    line_height = size * line_gap
    y = rect.y0
    drawn = 0
    for line in lines:
        if y + line_height > rect.y1:
            break
        page.insert_text(
            fitz.Point(rect.x0, y + size),
            line,
            fontsize=size,
            fontname=font_name,
            fontfile=font_file,
            color=color,
        )
        y += line_height
        drawn += 1
    return drawn, len(lines), y


def segment_order_key(segment, page_width):
    bounds = segment.get("bounds") or {}
    x = float(bounds.get("x") or 0)
    y = float(bounds.get("y") or 0)
    column = 0 if x < page_width * 0.52 else 1
    return (column, -y, x)


def is_heading(segment, page_width):
    bounds = segment.get("bounds") or {}
    text = clean_text(segment.get("text"))
    width = float(bounds.get("width") or 0)
    x = float(bounds.get("x") or 0)
    font_size = float(segment.get("fontSize") or 0)
    char_count = len(text)
    if not text:
        return False
    if font_size >= 12.5 and char_count <= 80:
        return True
    if width >= page_width * 0.62 and char_count <= 42:
        return True
    if char_count <= 28 and x < page_width * 0.30:
        return True
    return False


def add_page(doc, width, height, title, page_label, font_name, font_file, font):
    page = doc.new_page(width=width, height=height)
    margin = max(36, min(width, height) * 0.055)
    page.insert_text(
        fitz.Point(margin, margin * 0.62),
        f"{title}  {page_label}",
        fontsize=9,
        fontname=font_name,
        fontfile=font_file,
        color=(0.38, 0.42, 0.44),
    )
    page.draw_line(
        fitz.Point(margin, margin * 0.85),
        fitz.Point(width - margin, margin * 0.85),
        color=(0.78, 0.82, 0.84),
        width=0.6,
    )
    return page, margin, margin * 1.18


def build_pdf(payload, output_path):
    font_file = find_font()
    font_name = "msyh" if font_file else "helv"
    font = fitz.Font(fontfile=font_file) if font_file else fitz.Font("helv")
    title = clean_text(payload.get("title")) or "PDF译文阅读版"
    pages = payload.get("pages") or []
    segments = payload.get("segments") or []
    by_page = {}
    for segment in segments:
        text = clean_text(segment.get("text"))
        if not text:
            continue
        page_number = int(segment.get("page") or 1)
        by_page.setdefault(page_number, []).append({**segment, "text": text})

    doc = fitz.open()
    default_width = 595.28
    default_height = 841.89

    for source_page_index, source_page in enumerate(pages or [{"width": default_width, "height": default_height}], start=1):
        width = float(source_page.get("width") or default_width)
        height = float(source_page.get("height") or default_height)
        width = max(420, min(width, 900))
        height = max(560, min(height, 1200))
        page_segments = by_page.get(source_page_index, [])
        page_segments.sort(key=lambda item: segment_order_key(item, width))

        output_page, margin, y = add_page(doc, width, height, title, f"原第 {source_page_index} 页", font_name, font_file, font)
        gap = max(18, width * 0.035)
        column_width = (width - margin * 2 - gap) / 2
        bottom = height - margin * 0.85
        left_x = margin
        right_x = margin + column_width + gap
        current_column = 0
        column_y = [y, y]

        for segment in page_segments:
            heading = is_heading(segment, width)
            size = 11.5 if heading else 9.8
            line_gap = 1.45 if heading else 1.38
            if heading:
                target_x = margin
                target_width = width - margin * 2
                target_y = max(column_y)
                estimated_lines = len(wrap_text(font, segment["text"], size, target_width))
                needed = estimated_lines * size * line_gap + 8
                if target_y + needed > bottom:
                    output_page, margin, target_y = add_page(doc, width, height, title, f"原第 {source_page_index} 页续", font_name, font_file, font)
                    gap = max(18, width * 0.035)
                    column_width = (width - margin * 2 - gap) / 2
                    bottom = height - margin * 0.85
                    left_x = margin
                    right_x = margin + column_width + gap
                    column_y = [target_y, target_y]
                rect = fitz.Rect(target_x, target_y, target_x + target_width, bottom)
                drawn, total, new_y = draw_wrapped_text(output_page, font_name, font_file, font, segment["text"], rect, size, line_gap=line_gap)
                column_y = [new_y + 8, new_y + 8]
                current_column = 0
                continue

            column = current_column
            estimated_lines = len(wrap_text(font, segment["text"], size, column_width))
            needed = estimated_lines * size * line_gap + 5
            if column_y[column] + needed > bottom:
                if column == 0:
                    column = 1
                    current_column = 1
                else:
                    output_page, margin, y = add_page(doc, width, height, title, f"原第 {source_page_index} 页续", font_name, font_file, font)
                    gap = max(18, width * 0.035)
                    column_width = (width - margin * 2 - gap) / 2
                    bottom = height - margin * 0.85
                    left_x = margin
                    right_x = margin + column_width + gap
                    column_y = [y, y]
                    column = 0
                    current_column = 0

            target_x = left_x if column == 0 else right_x
            rect = fitz.Rect(target_x, column_y[column], target_x + column_width, bottom)
            drawn, total, new_y = draw_wrapped_text(output_page, font_name, font_file, font, segment["text"], rect, size, line_gap=line_gap)
            column_y[column] = new_y + 4

    if doc.page_count == 0:
        page, _, _ = add_page(doc, default_width, default_height, title, "", font_name, font_file, font)
        page.insert_text(fitz.Point(72, 96), "没有可导出的译文。", fontsize=12, fontname=font_name, fontfile=font_file)

    doc.save(output_path, deflate=True, garbage=4)
    doc.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    with open(args.input, "r", encoding="utf-8") as handle:
        payload = json.load(handle)
    build_pdf(payload, args.output)
    print(json.dumps({"output": os.path.abspath(args.output), "bytes": os.path.getsize(args.output)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
