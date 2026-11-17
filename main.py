import os
import uuid
import zipfile
import io
import sys
import time
import shutil
import re
import html
import subprocess
from datetime import datetime
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, make_response
from PIL import Image
import img2pdf
import pdf2docx
import openpyxl
import pandas as pd
from pypdf import PdfReader, PdfWriter
import docx
from docx import Document
import fitz  # PyMuPDF

# --- Проверка опциональных зависимостей ---
try:
    import pdf2image
    POPPLER_INSTALLED = True
except ImportError:
    POPPLER_INSTALLED = False

try:
    import pytesseract
    pytesseract.get_tesseract_version()
    TESSERACT_INSTALLED = True
except Exception:
    TESSERACT_INSTALLED = False

try:
    import tabula
    TABULA_INSTALLED = True
except Exception:
    TABULA_INSTALLED = False

try:
    import cv2
    import numpy as np
    CV2_INSTALLED = True
except Exception:
    CV2_INSTALLED = False


app = Flask(__name__)
app.secret_key = 'office_master_super_secret_key_2026'

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB max upload

# --- Поиск шрифтов для кириллицы ---
CYRILLIC_FONT_PATH = None
for candidate in [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts', 'Arial.ttf'),
    '/System/Library/Fonts/Supplemental/Arial.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    '/Library/Fonts/Arial.ttf',
    'C:\\Windows\\Fonts\\arial.ttf'
]:
    if os.path.isfile(candidate):
        CYRILLIC_FONT_PATH = candidate
        break

print(f"Используемый кириллический шрифт: {CYRILLIC_FONT_PATH}")


# --- Вспомогательные функции ---

def safe_clean_filename(filename):
    """Очищает имя файла, сохраняя русские буквы и безопасные символы."""
    filename = os.path.basename(filename)
    cleaned = re.sub(r'[^\w\s\.\-\(\)\[\]]', '_', filename, flags=re.UNICODE)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned or f"document_{uuid.uuid4().hex[:6]}"


def cleanup_old_uploads():
    """Удаляет временные папки старше 2 часов для экономии места."""
    now = time.time()
    try:
        for item in os.listdir(UPLOAD_FOLDER):
            item_path = os.path.join(UPLOAD_FOLDER, item)
            # Если старше 2 часов (7200 сек)
            if os.path.isdir(item_path) and now - os.path.getmtime(item_path) > 7200:
                shutil.rmtree(item_path, ignore_errors=True)
            elif os.path.isfile(item_path) and now - os.path.getmtime(item_path) > 7200:
                os.remove(item_path)
    except Exception:
        pass


def parse_page_range(range_str, total_pages):
    """Парсит строку диапазонов страниц (например '1-3, 5, 8') в список 0-based индексов."""
    indices = set()
    if not range_str or not range_str.strip():
        return []
    parts = [p.strip() for p in range_str.split(',') if p.strip()]
    for part in parts:
        if '-' in part:
            sub = [s.strip() for s in part.split('-') if s.strip()]
            if len(sub) == 2 and sub[0].isdigit() and sub[1].isdigit():
                start, end = int(sub[0]), int(sub[1])
                if start > end:
                    start, end = end, start
                for p in range(start, end + 1):
                    if 1 <= p <= total_pages:
                        indices.add(p - 1)
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                indices.add(p - 1)
    return sorted(list(indices))


def read_csv_smart(filepath):
    """Умное чтение CSV с авто-определением кодировки (UTF-8, UTF-8-sig, Windows-1251 для 1С/банков)."""
    encodings = ['utf-8-sig', 'cp1251', 'utf-8', 'latin1']
    for enc in encodings:
        try:
            return pd.read_csv(filepath, encoding=enc, sep=None, engine='python')
        except Exception:
            continue
    # Fallback с запятой
    return pd.read_csv(filepath, encoding='utf-8', errors='ignore')


# --- Функции работы с PDF и документами ---

def insert_font_to_page(page, font_alias="CyrillicFont"):
    if CYRILLIC_FONT_PATH and os.path.exists(CYRILLIC_FONT_PATH):
        try:
            page.insert_font(fontname=font_alias, fontfile=CYRILLIC_FONT_PATH)
            return font_alias
        except Exception:
            pass
    return "helv"


def draw_stamp_on_pdf(doc, stamp_text='ОПЛАЧЕНО', add_date=True, color_name='red', position='top-right', page_target='all'):
    """Наносит официальный бухгалтерский штамп или водяной знак на PDF."""
    colors = {
        'red': (0.86, 0.15, 0.15),
        'blue': (0.15, 0.38, 0.78),
        'green': (0.13, 0.58, 0.28),
        'gray': (0.35, 0.38, 0.42)
    }
    color_rgb = colors.get(color_name, (0.86, 0.15, 0.15))
    date_str = datetime.now().strftime('%d.%m.%Y') if add_date else None

    total_pages = len(doc)
    if page_target == 'first':
        target_indices = [0] if total_pages > 0 else []
    elif page_target == 'last':
        target_indices = [total_pages - 1] if total_pages > 0 else []
    else:
        target_indices = list(range(total_pages))

    for idx in target_indices:
        page = doc[idx]
        fname = insert_font_to_page(page, "StampFont")
        w, h = page.rect.width, page.rect.height

        if position == 'center_watermark':
            center_x, center_y = w / 2, h / 2
            wm_color = tuple(min(1.0, c * 0.4 + 0.6) for c in color_rgb)
            text_lines = [stamp_text]
            if date_str:
                text_lines.append(date_str)
            full_txt = " • ".join(text_lines)

            point = fitz.Point(center_x - 140, center_y + 40)
            morph = (point, fitz.Matrix(38))
            page.insert_text(point, full_txt, fontname=fname, fontsize=32, color=wm_color, morph=morph)
        else:
            box_w, box_h = 185, 65
            margin = 35

            if position == 'top-right':
                rect = fitz.Rect(w - box_w - margin, margin, w - margin, margin + box_h)
            elif position == 'bottom-right':
                rect = fitz.Rect(w - box_w - margin, h - box_h - margin, w - margin, h - margin)
            elif position == 'bottom-left':
                rect = fitz.Rect(margin, h - box_h - margin, margin + box_w, h - margin)
            else:
                rect = fitz.Rect(margin, margin, margin + box_w, margin + box_h)

            page.draw_rect(rect, color=color_rgb, width=2.2)
            inner_rect = fitz.Rect(rect.x0 + 3.5, rect.y0 + 3.5, rect.x1 - 3.5, rect.y1 - 3.5)
            page.draw_rect(inner_rect, color=color_rgb, width=0.8)

            content = f"{stamp_text}\n{date_str}" if date_str else stamp_text
            text_rect = fitz.Rect(rect.x0 + 4, rect.y0 + 8, rect.x1 - 4, rect.y1 - 4)
            page.insert_textbox(text_rect, content, fontname=fname, fontsize=13, color=color_rgb, align=fitz.TEXT_ALIGN_CENTER)


def add_page_numbers_to_pdf(doc, num_format='page_of_total', position='bottom-center', start_num=1, start_page=1):
    """Проставляет номера страниц («Стр. X из Y») для суда, налоговой, архива."""
    total_pages = len(doc)
    start_num = int(start_num) if start_num else 1
    start_page = int(start_page) if start_page else 1

    for idx in range(start_page - 1, total_pages):
        page = doc[idx]
        fname = insert_font_to_page(page, "NumFont")
        current_num = start_num + (idx - (start_page - 1))

        if num_format == 'page_of_total':
            text = f"Стр. {current_num} из {total_pages}"
        elif num_format == 'simple_of_total':
            text = f"{current_num} / {total_pages}"
        elif num_format == 'dashes':
            text = f"- {current_num} -"
        else:
            text = str(current_num)

        w, h = page.rect.width, page.rect.height
        margin = 25
        box_h = 20

        if position == 'bottom-center':
            rect = fitz.Rect(0, h - margin - box_h, w, h - margin)
            align = fitz.TEXT_ALIGN_CENTER
        elif position == 'bottom-right':
            rect = fitz.Rect(w - 200 - margin, h - margin - box_h, w - margin, h - margin)
            align = fitz.TEXT_ALIGN_RIGHT
        elif position == 'top-center':
            rect = fitz.Rect(0, margin, w, margin + box_h)
            align = fitz.TEXT_ALIGN_CENTER
        elif position == 'top-right':
            rect = fitz.Rect(w - 200 - margin, margin, w - margin, margin + box_h)
            align = fitz.TEXT_ALIGN_RIGHT
        else:
            rect = fitz.Rect(0, h - margin - box_h, w, h - margin)
            align = fitz.TEXT_ALIGN_CENTER

        page.insert_textbox(rect, text, fontname=fname, fontsize=10, color=(0.35, 0.35, 0.35), align=align)


def rotate_pdf_pages(doc, angle=90, pages_filter='all', custom_range=''):
    """Поворачивает страницы PDF (90, 180, 270 градусов)."""
    angle = int(angle)
    total = len(doc)
    custom_indices = set(parse_page_range(custom_range, total)) if pages_filter == 'custom' else set()

    for idx, page in enumerate(doc):
        should_rotate = False
        if pages_filter == 'all':
            should_rotate = True
        elif pages_filter == 'odd' and (idx + 1) % 2 == 1:
            should_rotate = True
        elif pages_filter == 'even' and (idx + 1) % 2 == 0:
            should_rotate = True
        elif pages_filter == 'custom' and idx in custom_indices:
            should_rotate = True

        if should_rotate:
            page.set_rotation((page.rotation + angle) % 360)


def compress_pdf_file(input_pdf_path, output_pdf_path, level='optimal'):
    """Сжимает PDF файл (оптимальное для почты или экстремальное для лимитов Госуслуг)."""
    doc = fitz.open(input_pdf_path)
    if level == 'extreme':
        comp_doc = fitz.open()
        for page in doc:
            pix = page.get_pixmap(dpi=120)
            img_data = pix.pil_tobytes(format='JPEG', quality=60, optimize=True)
            img_pdf = fitz.open('pdf', fitz.open('jpeg', img_data).convert_to_pdf())
            comp_doc.insert_pdf(img_pdf)
        comp_doc.save(output_pdf_path, garbage=4, deflate=True)
        comp_doc.close()
    else:
        doc.save(output_pdf_path, garbage=4, deflate=True, clean=True, deflate_images=True, deflate_fonts=True)
    doc.close()


def images_to_pdf(image_paths, output_pdf_path):
    """Объединяет любые картинки (PNG, JPG, WEBP, BMP, TIFF) в один PDF файл."""
    doc = fitz.open()
    for path in image_paths:
        with Image.open(path) as img:
            if img.mode in ('RGBA', 'LA', 'P'):
                bg = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                bg.paste(img, mask=img.split()[-1] if 'A' in img.getbands() else None)
                buf = io.BytesIO()
                bg.save(buf, format='JPEG', quality=95)
                buf.seek(0)
                img_pdf = fitz.open('pdf', fitz.open('jpeg', buf.getvalue()).convert_to_pdf())
            else:
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=95)
                buf.seek(0)
                img_pdf = fitz.open('pdf', fitz.open('jpeg', buf.getvalue()).convert_to_pdf())
            doc.insert_pdf(img_pdf)
    doc.save(output_pdf_path)
    doc.close()


def extract_text_from_doc_binary(doc_path):
    """Извлекает структурированный текст из устаревших бинарных файлов Word 97-2003 (.doc)."""
    # 1. Попытка через soffice (LibreOffice headless), если установлен
    if shutil.which('soffice'):
        try:
            out_dir = os.path.dirname(doc_path)
            cmd = ['soffice', '--headless', '--convert-to', 'txt:Text', doc_path, '--outdir', out_dir]
            subprocess.run(cmd, timeout=25, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            base = os.path.splitext(os.path.basename(doc_path))[0]
            txt_path = os.path.join(out_dir, base + '.txt')
            if os.path.exists(txt_path):
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as tf:
                    t = tf.read().strip()
                try:
                    os.remove(txt_path)
                except Exception:
                    pass
                if t:
                    return t
        except Exception:
            pass

    # 2. Попытка через antiword CLI
    if shutil.which('antiword'):
        try:
            res = subprocess.run(['antiword', '-m', 'UTF-8.txt', doc_path], capture_output=True, timeout=15, text=True)
            if res.returncode == 0 and res.stdout.strip():
                return res.stdout.strip()
        except Exception:
            pass

    # 3. Встроенный парсер бинарного формата Microsoft Word (.doc)
    try:
        with open(doc_path, 'rb') as f:
            data = f.read()

        valid_chars = set('АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя'
                          'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 '
                          '.,;:№"\'()-_/«»!?+=%@\t\n')

        def parse_decoded(text):
            raw_paragraphs = re.split(r'[\r\n\x07\x0c]+', text)
            filtered = []
            for p in raw_paragraphs:
                p = p.strip()
                if not p:
                    continue
                valid_count = sum(1 for c in p if c in valid_chars)
                if valid_count / len(p) >= 0.85:
                    letters = [c for c in p if c.isalpha()]
                    if len(letters) >= 3:
                        if not any(p.startswith(ign) for ign in ['Root Entry', 'WordDocument', 'SummaryInformation', 'DocumentSummaryInformation', 'CompObj']):
                            if not any(ign in p for ign in ['Absatz-Standardschriftart', 'WW-Absatz']):
                                filtered.append(re.sub(r'\s+', ' ', p).strip())
            return filtered

        lines = parse_decoded(data.decode('utf-16le', errors='ignore'))
        if not lines:
            lines = parse_decoded(data.decode('cp1251', errors='ignore'))

        if lines:
            return '\n'.join(lines)
    except Exception:
        pass

    return "Не удалось извлечь содержимое документа .doc"


def doc_to_pdf(doc_path, output_pdf_path):
    """Конвертирует устаревший бинарный Word 97-2003 (.doc) в PDF."""
    # Попытка через soffice
    if shutil.which('soffice'):
        try:
            out_dir = os.path.dirname(output_pdf_path)
            cmd = ['soffice', '--headless', '--convert-to', 'pdf', doc_path, '--outdir', out_dir]
            subprocess.run(cmd, timeout=35, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            base = os.path.splitext(os.path.basename(doc_path))[0]
            gen_pdf = os.path.join(out_dir, base + '.pdf')
            if os.path.exists(gen_pdf):
                if gen_pdf != output_pdf_path:
                    shutil.move(gen_pdf, output_pdf_path)
                return
        except Exception:
            pass

    # Встроенный рендеринг через fitz.Story
    text = extract_text_from_doc_binary(doc_path)
    html_parts = ['<div style="font-family: sans-serif; font-size: 11pt; line-height: 1.5; color: #1e293b;">']
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            html_parts.append('<p style="margin: 4px 0;">&nbsp;</p>')
        else:
            html_parts.append(f'<p style="margin: 6px 0;">{html.escape(line)}</p>')
    html_parts.append('</div>')

    story = fitz.Story(''.join(html_parts))
    writer = fitz.DocumentWriter(output_pdf_path)
    more = 1
    while more:
        device = writer.begin_page(fitz.paper_rect('A4'))
        more, _ = story.place(fitz.Rect(40, 40, 555, 802))
        story.draw(device)
        writer.end_page()
    writer.close()


def docx_to_pdf(docx_path, output_pdf_path):
    """Преобразует Word DOCX в PDF с заголовками, стилями и таблицами."""
    # Если передан .doc или файл не является OpenXML архивом
    try:
        d = docx.Document(docx_path)
    except Exception:
        # Fallback на парсер .doc
        doc_to_pdf(docx_path, output_pdf_path)
        return

    html_parts = ['<div style="font-family: sans-serif; font-size: 11pt; line-height: 1.45; color: #1e293b;">']

    for element in d.element.body:
        tag = element.tag.split('}')[-1]
        if tag == 'p':
            p = docx.text.paragraph.Paragraph(element, d)
            text = p.text.strip()
            if not text:
                html_parts.append('<p style="margin: 4px 0;">&nbsp;</p>')
                continue

            p_style = p.style.name if p.style else ''
            if 'Heading 1' in p_style or 'Title' in p_style:
                html_parts.append(f'<h1 style="font-size: 18pt; color: #0f172a; margin: 14px 0 8px 0; font-weight: bold;">{html.escape(text)}</h1>')
            elif 'Heading 2' in p_style:
                html_parts.append(f'<h2 style="font-size: 15pt; color: #1e293b; margin: 12px 0 6px 0; font-weight: bold;">{html.escape(text)}</h2>')
            elif 'Heading 3' in p_style:
                html_parts.append(f'<h3 style="font-size: 13pt; color: #334155; margin: 10px 0 4px 0; font-weight: bold;">{html.escape(text)}</h3>')
            else:
                p_html = ''
                for run in p.runs:
                    txt = html.escape(run.text)
                    if run.bold: txt = f'<b>{txt}</b>'
                    if run.italic: txt = f'<i>{txt}</i>'
                    if run.underline: txt = f'<u>{txt}</u>'
                    p_html += txt

                align = 'left'
                if p.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
                    align = 'center'
                elif p.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT:
                    align = 'right'
                elif p.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY:
                    align = 'justify'

                if 'List' in p_style:
                    html_parts.append(f'<li style="margin-left: 20px; text-align: {align};">{p_html or html.escape(text)}</li>')
                else:
                    html_parts.append(f'<p style="text-align: {align}; margin: 5px 0;">{p_html or html.escape(text)}</p>')

        elif tag == 'tbl':
            tbl = docx.table.Table(element, d)
            html_parts.append('<table style="width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 10pt;">')
            for r_idx, row in enumerate(tbl.rows):
                bg = '#f1f5f9' if r_idx == 0 else ('#f8fafc' if r_idx % 2 == 1 else '#ffffff')
                weight = 'bold' if r_idx == 0 else 'normal'
                html_parts.append(f'<tr style="background: {bg}; font-weight: {weight};">')
                for cell in row.cells:
                    html_parts.append(f'<td style="border: 1px solid #cbd5e1; padding: 6px 10px;">{html.escape(cell.text.strip())}</td>')
                html_parts.append('</tr>')
            html_parts.append('</table>')

    html_parts.append('</div>')
    story = fitz.Story(''.join(html_parts))
    writer = fitz.DocumentWriter(output_pdf_path)
    more = 1
    while more:
        device = writer.begin_page(fitz.paper_rect('A4'))
        more, _ = story.place(fitz.Rect(40, 40, 555, 802))
        story.draw(device)
        writer.end_page()
    writer.close()


def excel_to_pdf(input_path, output_pdf_path):
    """Преобразует Excel (XLSX, XLS) или CSV в альбомный печатный PDF документ."""
    ext = input_path.lower().rsplit('.', 1)[-1]
    if ext == 'csv':
        df = read_csv_smart(input_path)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = os.path.basename(input_path).rsplit('.', 1)[0][:30]
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append(list(row))
    else:
        wb = openpyxl.load_workbook(input_path, data_only=True)

    html_parts = ['<div style="font-family: sans-serif; font-size: 9pt; color: #1e293b;">']
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        html_parts.append(f'<h3 style="font-size: 13pt; color: #1e3a8a; margin: 15px 0 8px 0; border-bottom: 2px solid #3b82f6; padding-bottom: 4px;">Лист: {html.escape(sheetname)}</h3>')
        html_parts.append('<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">')
        for r_idx, row in enumerate(rows):
            if all(c is None or str(c).strip() == '' for c in row):
                continue
            is_hdr = (r_idx == 0)
            bg = '#e2e8f0' if is_hdr else ('#f8fafc' if r_idx % 2 == 1 else '#ffffff')
            weight = 'bold' if is_hdr else 'normal'
            html_parts.append(f'<tr style="background: {bg}; font-weight: {weight};">')
            for cell in row:
                val = '' if cell is None else str(cell)
                align = 'right' if isinstance(cell, (int, float)) else 'left'
                html_parts.append(f'<td style="border: 1px solid #94a3b8; padding: 4px 6px; text-align: {align};">{html.escape(val)}</td>')
            html_parts.append('</tr>')
        html_parts.append('</table>')
    html_parts.append('</div>')

    story = fitz.Story(''.join(html_parts))
    writer = fitz.DocumentWriter(output_pdf_path)
    more = 1
    while more:
        device = writer.begin_page(fitz.paper_rect('A4-L'))
        more, _ = story.place(fitz.Rect(30, 30, 812, 565))
        story.draw(device)
        writer.end_page()
    writer.close()


def pdf_to_excel(pdf_path, output_xlsx_path):
    """Извлекает таблицы из PDF в книгу Excel."""
    doc = fitz.open(pdf_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    tables_found = 0

    for p_idx, page in enumerate(doc):
        try:
            tabs = page.find_tables()
            for t_idx, tab in enumerate(tabs):
                df = tab.to_pandas()
                if df is not None and not df.empty:
                    ws = wb.create_sheet(title=f"Стр_{p_idx+1}_Табл_{t_idx+1}"[:30])
                    ws.append(list(df.columns))
                    for row in df.itertuples(index=False):
                        ws.append(list(row))
                    tables_found += 1
        except Exception:
            pass

    if tables_found == 0 and TABULA_INSTALLED:
        try:
            dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
            if dfs:
                for t_idx, df in enumerate(dfs):
                    if not df.empty:
                        ws = wb.create_sheet(title=f"Таблица_{t_idx+1}"[:30])
                        ws.append(list(df.columns))
                        for row in df.itertuples(index=False):
                            ws.append(list(row))
                        tables_found += 1
        except Exception:
            pass

    if tables_found == 0:
        for p_idx, page in enumerate(doc):
            lines = page.get_text('text').split('\n')
            non_empty = [l for l in lines if l.strip()]
            if non_empty:
                ws = wb.create_sheet(title=f"Страница_{p_idx+1}"[:30])
                for line in non_empty:
                    cols = [c.strip() for c in line.split('\t') if c.strip()]
                    if len(cols) <= 1:
                        cols = [c.strip() for c in line.split('   ') if c.strip()]
                    ws.append(cols if cols else [line.strip()])
                tables_found += 1

    if tables_found == 0:
        ws = wb.create_sheet(title="Данные")
        ws.append(["Текст не обнаружен в документе"])

    wb.save(output_xlsx_path)


def merge_excel_csv_to_workbook(file_paths, output_xlsx_path):
    """Объединяет несколько файлов CSV или Excel в одну книгу Excel по отдельным листам."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for idx, path in enumerate(file_paths):
        base_name = os.path.basename(path).rsplit('.', 1)[0][:25]
        ext = path.lower().rsplit('.', 1)[-1]

        if ext == 'csv':
            df = read_csv_smart(path)
            sheet_title = f"{base_name}"[:30]
            ws = wb.create_sheet(title=sheet_title)
            ws.append(list(df.columns))
            for row in df.itertuples(index=False):
                ws.append(list(row))
        elif ext in ('xlsx', 'xls'):
            src_wb = openpyxl.load_workbook(path, data_only=True)
            for s_name in src_wb.sheetnames:
                src_ws = src_wb[s_name]
                title = f"{base_name}_{s_name}"[:30] if len(src_wb.sheetnames) > 1 else base_name[:30]
                ws = wb.create_sheet(title=title)
                for row in src_ws.iter_rows(values_only=True):
                    ws.append(list(row))

    wb.save(output_xlsx_path)


def txt_to_pdf(txt_path, output_pdf_path):
    """Конвертирует текстовый файл TXT в форматированный PDF."""
    with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
        text = f.read()
    escaped = html.escape(text).replace('\n', '<br>')
    html_content = f'<div style="font-family: monospace; font-size: 10pt; line-height: 1.45; color: #1e293b;">{escaped}</div>'
    story = fitz.Story(html_content)
    writer = fitz.DocumentWriter(output_pdf_path)
    more = 1
    while more:
        device = writer.begin_page(fitz.paper_rect('A4'))
        more, _ = story.place(fitz.Rect(40, 40, 555, 802))
        story.draw(device)
        writer.end_page()
    writer.close()


def extract_text_from_file(file_path):
    """Извлекает весь чистый текст из PDF, DOCX, DOC, XLSX или TXT."""
    ext = file_path.lower().rsplit('.', 1)[-1]
    if ext == 'pdf':
        doc = fitz.open(file_path)
        parts = []
        for i, page in enumerate(doc):
            parts.append(f"=== Страница {i+1} ===\n{page.get_text()}")
        return "\n\n".join(parts)
    elif ext == 'doc':
        return extract_text_from_doc_binary(file_path)
    elif ext == 'docx':
        try:
            d = docx.Document(file_path)
            return "\n".join([p.text for p in d.paragraphs if p.text.strip()])
        except Exception:
            return extract_text_from_doc_binary(file_path)
    elif ext in ('xlsx', 'xls', 'csv'):
        df = read_csv_smart(file_path) if ext == 'csv' else pd.read_excel(file_path)
        return df.to_string()
    elif ext == 'txt':
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    return "Формат не поддерживается для прямого извлечения текста."


# --- OCR функции ---

def ocr_text_from_image(image_path):
    if not TESSERACT_INSTALLED:
        raise EnvironmentError("Tesseract OCR не установлен.")
    custom_config = r'--oem 3 --psm 6 -l rus+eng'
    img = Image.open(image_path)
    return pytesseract.image_to_string(img, config=custom_config)


def ocr_table_data_from_image(image_path):
    if not TESSERACT_INSTALLED:
        raise EnvironmentError("Tesseract OCR не установлен.")
    img = Image.open(image_path)
    custom_config = r'--oem 3 --psm 6 -l rus+eng'
    data = pytesseract.image_to_data(img, config=custom_config, output_type=pytesseract.Output.DICT)

    rows = []
    current_row_words = []
    last_top = -1
    y_threshold = 20

    n_boxes = len(data['text'])
    for i in range(n_boxes):
        if int(data['conf'][i]) > 0:
            text = data['text'][i]
            top = data['top'][i]
            left = data['left'][i]

            if last_top == -1:
                last_top = top

            if abs(top - last_top) > y_threshold:
                if current_row_words:
                    current_row_words.sort(key=lambda x: x['left'])
                    rows.append([w['text'] for w in current_row_words])
                current_row_words = []
                last_top = top

            if text.strip():
                current_row_words.append({'text': text, 'left': left})

    if current_row_words:
        current_row_words.sort(key=lambda x: x['left'])
        rows.append([w['text'] for w in current_row_words])

    return rows


# --- ГЛАВНЫЙ МАРШРУТ FLASK ---

@app.route('/', methods=['GET', 'POST'])
def index():
    cleanup_old_uploads()

    if request.method == 'POST':
        uploaded_files = request.files.getlist('files') or request.files.getlist('file')
        valid_files = [f for f in uploaded_files if f and f.filename and f.filename.strip()]

        target_format = request.form.get('format', 'MERGE_PDF')
        download_token = request.form.get('download_token')

        if not valid_files:
            flash('Пожалуйста, выберите хотя бы один файл.')
            return redirect(url_for('index'))

        req_id = str(uuid.uuid4())
        req_dir = os.path.join(app.config['UPLOAD_FOLDER'], req_id)
        os.makedirs(req_dir, exist_ok=True)

        # Сохранение файлов
        saved_files = []
        for f in valid_files:
            orig_name = safe_clean_filename(f.filename)
            ext = orig_name.rsplit('.', 1)[-1].lower() if '.' in orig_name else ''
            disk_path = os.path.join(req_dir, f"{len(saved_files)}_{orig_name}")
            f.save(disk_path)
            saved_files.append((disk_path, orig_name, ext))

        try:
            # 1. ОБЪЕДИНЕНИЕ В ОДИН PDF (MERGE_PDF)
            if target_format == 'MERGE_PDF':
                merged_doc = fitz.open()
                for disk_path, orig_name, ext in saved_files:
                    if ext == 'pdf':
                        src = fitz.open(disk_path)
                        merged_doc.insert_pdf(src)
                        src.close()
                    elif ext in ('png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'tif'):
                        temp_pdf = os.path.join(req_dir, f"temp_{uuid.uuid4().hex[:6]}.pdf")
                        images_to_pdf([disk_path], temp_pdf)
                        src = fitz.open(temp_pdf)
                        merged_doc.insert_pdf(src)
                        src.close()
                    elif ext == 'doc':
                        temp_pdf = os.path.join(req_dir, f"temp_{uuid.uuid4().hex[:6]}.pdf")
                        doc_to_pdf(disk_path, temp_pdf)
                        src = fitz.open(temp_pdf)
                        merged_doc.insert_pdf(src)
                        src.close()
                    elif ext == 'docx':
                        temp_pdf = os.path.join(req_dir, f"temp_{uuid.uuid4().hex[:6]}.pdf")
                        try:
                            docx_to_pdf(disk_path, temp_pdf)
                        except Exception:
                            doc_to_pdf(disk_path, temp_pdf)
                        src = fitz.open(temp_pdf)
                        merged_doc.insert_pdf(src)
                        src.close()
                    else:
                        flash(f'Файл {orig_name} пропущен (не поддерживается для склейки).')

                if len(merged_doc) == 0:
                    flash('Не удалось собрать ни одной страницы в PDF.')
                    return redirect(url_for('index'))

                out_file = os.path.join(req_dir, 'merged_document.pdf')
                merged_doc.save(out_file)
                merged_doc.close()
                return make_file_response(out_file, 'merged_document.pdf', 'application/pdf', download_token)

            # 2. КАРТИНКИ В ОДИН PDF (PDF_FROM_IMAGES)
            elif target_format == 'PDF_FROM_IMAGES':
                img_paths = [p for p, name, ext in saved_files if ext in ('png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'tif')]
                if not img_paths:
                    flash('Не найдено изображений для объединения в PDF.')
                    return redirect(url_for('index'))

                out_file = os.path.join(req_dir, 'images_combined.pdf')
                images_to_pdf(img_paths, out_file)
                return make_file_response(out_file, 'images_combined.pdf', 'application/pdf', download_token)

            # 3. ОБЪЕДИНИТЬ НЕСКОЛЬКО EXCEL / CSV В ОДНУ КНИГУ (MERGE_EXCEL)
            elif target_format == 'MERGE_EXCEL':
                excel_csv_paths = [p for p, name, ext in saved_files if ext in ('xlsx', 'xls', 'csv')]
                if not excel_csv_paths:
                    flash('Для объединения выберите файлы Excel (.xlsx, .xls) или CSV.')
                    return redirect(url_for('index'))

                out_file = os.path.join(req_dir, 'combined_sheets.xlsx')
                merge_excel_csv_to_workbook(excel_csv_paths, out_file)
                return make_file_response(out_file, 'combined_worksheets.xlsx',
                                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', download_token)

            # 4. РАЗДЕЛИТЬ PDF ПОСТРАНИЧНО (SPLIT_PDF)
            elif target_format == 'SPLIT_PDF':
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for disk_path, orig_name, ext in saved_files:
                        if ext != 'pdf':
                            continue
                        src = fitz.open(disk_path)
                        base_orig = orig_name.rsplit('.', 1)[0]
                        for i, page in enumerate(src):
                            single_doc = fitz.open()
                            single_doc.insert_pdf(src, from_page=i, to_page=i)
                            page_bytes = single_doc.tobytes()
                            single_doc.close()
                            zf.writestr(f"{base_orig}_стр_{i+1}.pdf", page_bytes)
                        src.close()

                zip_buffer.seek(0)
                return make_stream_response(zip_buffer, 'split_pages.zip', 'application/zip', download_token)

            # 5. PDF В КАРТИНКИ (IMAGES_FROM_PDF)
            elif target_format == 'IMAGES_FROM_PDF':
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for disk_path, orig_name, ext in saved_files:
                        if ext != 'pdf':
                            continue
                        src = fitz.open(disk_path)
                        base_orig = orig_name.rsplit('.', 1)[0]
                        for i, page in enumerate(src):
                            pix = page.get_pixmap(dpi=200)
                            zf.writestr(f"{base_orig}_стр_{i+1}.png", pix.tobytes('png'))
                        src.close()

                zip_buffer.seek(0)
                return make_stream_response(zip_buffer, 'pdf_pages_images.zip', 'application/zip', download_token)

            # --- ОПЕРАЦИИ, ОБРАБАТЫВАЮЩИЕ КАЖДЫЙ ФАЙЛ (ПАКЕТНАЯ ОБРАБОТКА) ---
            processed_items = []

            for disk_path, orig_name, ext in saved_files:
                base_name = orig_name.rsplit('.', 1)[0] if '.' in orig_name else orig_name

                # А. ШТАМП НА ДОКУМЕНТ (STAMP_PDF)
                if target_format == 'STAMP_PDF':
                    if ext != 'pdf':
                        continue
                    doc = fitz.open(disk_path)
                    stamp_text = request.form.get('stamp_text', 'ОПЛАЧЕНО').strip() or 'ОПЛАЧЕНО'
                    add_date = (request.form.get('stamp_add_date') == 'true')
                    color = request.form.get('stamp_color', 'red')
                    position = request.form.get('stamp_position', 'top-right')
                    pages_target = request.form.get('stamp_pages', 'all')

                    draw_stamp_on_pdf(doc, stamp_text, add_date, color, position, pages_target)
                    out_path = os.path.join(req_dir, f"{base_name}_stamped.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_stamped.pdf", 'application/pdf'))

                # Б. НУМЕРАЦИЯ СТРАНИЦ (NUMBER_PAGES)
                elif target_format == 'NUMBER_PAGES':
                    if ext != 'pdf':
                        continue
                    doc = fitz.open(disk_path)
                    num_format = request.form.get('num_format', 'page_of_total')
                    position = request.form.get('num_position', 'bottom-center')
                    start_num = int(request.form.get('num_start', 1) or 1)
                    start_page = int(request.form.get('num_start_page', 1) or 1)

                    add_page_numbers_to_pdf(doc, num_format, position, start_num, start_page)
                    out_path = os.path.join(req_dir, f"{base_name}_numbered.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_numbered.pdf", 'application/pdf'))

                # В. ПОВОРОТ СТРАНИЦ (ROTATE_PDF)
                elif target_format == 'ROTATE_PDF':
                    if ext != 'pdf':
                        continue
                    doc = fitz.open(disk_path)
                    angle = int(request.form.get('rotate_angle', 90) or 90)
                    filter_mode = request.form.get('rotate_filter', 'all')
                    custom_range = request.form.get('rotate_custom_range', '')

                    rotate_pdf_pages(doc, angle, filter_mode, custom_range)
                    out_path = os.path.join(req_dir, f"{base_name}_rotated.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_rotated.pdf", 'application/pdf'))

                # Г. ИЗВЛЕЧЬ СТРАНИЦЫ (EXTRACT_PAGES)
                elif target_format == 'EXTRACT_PAGES':
                    if ext != 'pdf':
                        continue
                    doc = fitz.open(disk_path)
                    page_range = request.form.get('page_range', '')
                    indices = parse_page_range(page_range, len(doc))
                    if not indices:
                        flash(f'Укажите корректные номера страниц для извлечения из {orig_name}.')
                        return redirect(url_for('index'))
                    doc.select(indices)
                    out_path = os.path.join(req_dir, f"{base_name}_extracted.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_extracted.pdf", 'application/pdf'))

                # Д. УДАЛИТЬ СТРАНИЦЫ (DELETE_PAGES)
                elif target_format == 'DELETE_PAGES':
                    if ext != 'pdf':
                        continue
                    doc = fitz.open(disk_path)
                    page_range = request.form.get('page_range', '')
                    indices = parse_page_range(page_range, len(doc))
                    if not indices:
                        flash(f'Укажите корректные номера страниц для удаления из {orig_name}.')
                        return redirect(url_for('index'))
                    if len(indices) >= len(doc):
                        flash('Нельзя удалить все страницы документа.')
                        return redirect(url_for('index'))
                    doc.delete_pages(indices)
                    out_path = os.path.join(req_dir, f"{base_name}_trimmed.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_trimmed.pdf", 'application/pdf'))

                # Е. СЖАТИЕ PDF (COMPRESS_PDF)
                elif target_format == 'COMPRESS_PDF':
                    if ext != 'pdf':
                        continue
                    level = request.form.get('compress_level', 'optimal')
                    out_path = os.path.join(req_dir, f"{base_name}_compressed.pdf")
                    compress_pdf_file(disk_path, out_path, level=level)
                    processed_items.append((out_path, f"{base_name}_compressed.pdf", 'application/pdf'))

                # Ж. WORD -> PDF (WORD_TO_PDF)
                elif target_format == 'WORD_TO_PDF':
                    if ext not in ('docx', 'doc'):
                        continue
                    out_path = os.path.join(req_dir, f"{base_name}.pdf")
                    if ext == 'doc':
                        doc_to_pdf(disk_path, out_path)
                    else:
                        try:
                            docx_to_pdf(disk_path, out_path)
                        except Exception:
                            doc_to_pdf(disk_path, out_path)
                    processed_items.append((out_path, f"{base_name}.pdf", 'application/pdf'))

                # З. PDF -> WORD (WORD_FROM_PDF)
                elif target_format == 'WORD_FROM_PDF':
                    if ext != 'pdf':
                        continue
                    out_path = os.path.join(req_dir, f"{base_name}.docx")
                    try:
                        cv = pdf2docx.Converter(disk_path)
                        cv.convert(out_path, start=0, end=None)
                        cv.close()
                    except Exception:
                        d = Document()
                        src = fitz.open(disk_path)
                        for page in src:
                            lines = page.get_text('text').split('\n')
                            for l in lines:
                                if l.strip():
                                    d.add_paragraph(l)
                            d.add_page_break()
                        d.save(out_path)
                        src.close()
                    processed_items.append((out_path, f"{base_name}.docx",
                                            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))

                # И. EXCEL / CSV -> PDF (EXCEL_TO_PDF)
                elif target_format == 'EXCEL_TO_PDF':
                    if ext not in ('xlsx', 'xls', 'csv'):
                        continue
                    out_path = os.path.join(req_dir, f"{base_name}.pdf")
                    excel_to_pdf(disk_path, out_path)
                    processed_items.append((out_path, f"{base_name}.pdf", 'application/pdf'))

                # К. PDF -> EXCEL (PDF_TO_EXCEL или EXCEL_FROM_PDF)
                elif target_format in ('PDF_TO_EXCEL', 'EXCEL_FROM_PDF'):
                    if ext != 'pdf':
                        continue
                    out_path = os.path.join(req_dir, f"{base_name}_tables.xlsx")
                    pdf_to_excel(disk_path, out_path)
                    processed_items.append((out_path, f"{base_name}_tables.xlsx",
                                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))

                # Л. TXT -> PDF (TXT_TO_PDF)
                elif target_format == 'TXT_TO_PDF':
                    if ext != 'txt':
                        continue
                    out_path = os.path.join(req_dir, f"{base_name}.pdf")
                    txt_to_pdf(disk_path, out_path)
                    processed_items.append((out_path, f"{base_name}.pdf", 'application/pdf'))

                # М. ИЗВЛЕЧЬ ТЕКСТ В TXT (PDF_TO_TXT)
                elif target_format == 'PDF_TO_TXT':
                    out_path = os.path.join(req_dir, f"{base_name}.txt")
                    extracted = extract_text_from_file(disk_path)
                    with open(out_path, 'w', encoding='utf-8') as f_out:
                        f_out.write(extracted)
                    processed_items.append((out_path, f"{base_name}.txt", 'text/plain; charset=utf-8'))

                # Н. ЗАЩИТИТЬ ПАРОЛЕМ (PROTECT_PDF)
                elif target_format == 'PROTECT_PDF':
                    if ext != 'pdf':
                        continue
                    pw = request.form.get('password', '').strip()
                    if not pw:
                        flash('Пожалуйста, введите пароль для защиты документа.')
                        return redirect(url_for('index'))
                    doc = fitz.open(disk_path)
                    out_path = os.path.join(req_dir, f"{base_name}_protected.pdf")
                    doc.save(out_path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=pw, owner_pw=pw)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_protected.pdf", 'application/pdf'))

                # О. СНЯТЬ ПАРОЛЬ (UNLOCK_PDF)
                elif target_format == 'UNLOCK_PDF':
                    if ext != 'pdf':
                        continue
                    pw = request.form.get('password', '').strip()
                    doc = fitz.open(disk_path)
                    if doc.is_encrypted:
                        rc = doc.authenticate(pw)
                        if not rc:
                            flash(f'Неверный пароль для документа {orig_name}.')
                            return redirect(url_for('index'))
                    out_path = os.path.join(req_dir, f"{base_name}_unlocked.pdf")
                    doc.save(out_path)
                    doc.close()
                    processed_items.append((out_path, f"{base_name}_unlocked.pdf", 'application/pdf'))

                # П. СКАН В WORD (WORD_OCR)
                elif target_format == 'WORD_OCR':
                    out_path = os.path.join(req_dir, f"{base_name}_ocr.docx")
                    d = Document()
                    if ext == 'pdf':
                        src = fitz.open(disk_path)
                        for p_idx, page in enumerate(src):
                            if TESSERACT_INSTALLED:
                                pix = page.get_pixmap(dpi=250)
                                temp_img = os.path.join(req_dir, f"ocr_p_{p_idx}.png")
                                pix.save(temp_img)
                                txt = ocr_text_from_image(temp_img)
                            else:
                                txt = page.get_text('text')
                            d.add_paragraph(txt)
                            d.add_page_break()
                        src.close()
                    elif ext in ('png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'tif'):
                        if TESSERACT_INSTALLED:
                            txt = ocr_text_from_image(disk_path)
                        else:
                            txt = "[Tesseract OCR не найден на сервере. Распознавание скана недоступно.]"
                        d.add_paragraph(txt)
                    d.save(out_path)
                    processed_items.append((out_path, f"{base_name}_ocr.docx",
                                            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))

                # Р. СКАН ТАБЛИЦЫ В EXCEL (EXCEL_OCR)
                elif target_format == 'EXCEL_OCR':
                    out_path = os.path.join(req_dir, f"{base_name}_ocr.xlsx")
                    if ext == 'pdf':
                        pdf_to_excel(disk_path, out_path)
                    elif ext in ('png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'tif'):
                        if TESSERACT_INSTALLED:
                            rows = ocr_table_data_from_image(disk_path)
                            if rows:
                                max_cols = max(len(r) for r in rows)
                                padded = [r + [''] * (max_cols - len(r)) for r in rows]
                                df = pd.DataFrame(padded)
                                df.to_excel(out_path, index=False, header=False)
                            else:
                                wb = openpyxl.Workbook()
                                wb.active.append(['Табличные данные не найдены'])
                                wb.save(out_path)
                        else:
                            wb = openpyxl.Workbook()
                            wb.active.append(['Tesseract OCR не установлен'])
                            wb.save(out_path)
                    processed_items.append((out_path, f"{base_name}_ocr.xlsx",
                                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))

            # --- ФОРМИРОВАНИЕ ОТВЕТА (ОДИН ФАЙЛ ИЛИ ZIP АРХИВ) ---
            if not processed_items:
                flash('Ни один файл не подошел под выбранную операцию. Проверьте форматы файлов.')
                return redirect(url_for('index'))

            if len(processed_items) == 1:
                item_path, item_name, item_mime = processed_items[0]
                return make_file_response(item_path, item_name, item_mime, download_token)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                seen_names = {}
                for item_path, item_name, item_mime in processed_items:
                    base, ext_out = os.path.splitext(item_name)
                    cnt = seen_names.get(item_name, 0)
                    seen_names[item_name] = cnt + 1
                    entry_name = item_name if cnt == 0 else f"{base}_{cnt+1}{ext_out}"
                    with open(item_path, 'rb') as f_read:
                        zf.writestr(entry_name, f_read.read())

            zip_buffer.seek(0)
            return make_stream_response(zip_buffer, 'converted_documents.zip', 'application/zip', download_token)

        except Exception as e:
            import traceback
            traceback.print_exc()
            flash(f'Ошибка обработки: {str(e)}')
            return redirect(url_for('index'))

    now_date = datetime.now().strftime('%d.%m.%Y')
    return render_template('index.html',
                           poppler=POPPLER_INSTALLED,
                           tesseract=TESSERACT_INSTALLED,
                           now_date=now_date)


def make_file_response(file_path, download_name, mimetype, token=None):
    """Формирует HTTP-ответ с файлом и кукой для завершения спиннера в браузере."""
    resp = send_file(file_path, as_attachment=True, download_name=download_name, mimetype=mimetype)
    if token:
        resp.set_cookie('download_token', str(token), path='/', max_age=60, samesite='Lax')
    return resp


def make_stream_response(stream, download_name, mimetype, token=None):
    """Формирует HTTP-ответ из потока BytesIO и устанавливает куку загрузки."""
    resp = send_file(stream, as_attachment=True, download_name=download_name, mimetype=mimetype)
    if token:
        resp.set_cookie('download_token', str(token), path='/', max_age=60, samesite='Lax')
    return resp


if __name__ == '__main__':
    print("Запуск сервера ОфисМастер...")
    print(f"Статус OCR Tesseract: {'Установлен' if TESSERACT_INSTALLED else 'Не найден'}")
    print(f"Статус Poppler: {'Установлен' if POPPLER_INSTALLED else 'Векторный PyMuPDF используется'}")
    app.run(host='0.0.0.0', port=5000, debug=True)
