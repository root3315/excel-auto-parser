#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор PDF-листинга исходного кода проекта Excel Smart Parser.

Собирает исходники проекта в один большой PDF-документ:
  * обложка со статистикой (строки, размер, SHA-256);
  * содержание с номерами страниц;
  * подсветка синтаксиса (Pygments) + нумерация строк;
  * колонтитулы, закладки (outline) для навигации;
  * авто-подбор кегля так, чтобы объём попал в заданный диапазон страниц.

Зависимости:
    pip install reportlab pygments

Пример:
    python tools/generate_code_pdf.py --out excel_smart_parser_code.pdf
    python tools/generate_code_pdf.py --target-pages 30 50 --font-size 7.5
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import math
import os
import sys
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas as _canvas
except ImportError:  # pragma: no cover
    sys.exit("Не найден reportlab. Установите: pip install reportlab pygments")

try:
    from pygments import lex as _pygments_lex
    from pygments.lexers import get_lexer_for_filename, guess_lexer
    from pygments.styles import get_style_by_name
    from pygments.util import ClassNotFound
except ImportError:  # pragma: no cover
    sys.exit("Не найден pygments. Установите: pip install reportlab pygments")


# --------------------------------------------------------------------------- #
# Константы оформления
# --------------------------------------------------------------------------- #

MONO = "Courier"
MONO_BOLD = "Courier-Bold"
MONO_ITALIC = "Courier-Oblique"
MONO_BOLD_ITALIC = "Courier-BoldOblique"
SANS = "Helvetica"
SANS_BOLD = "Helvetica-Bold"

CHAR_WIDTH_RATIO = 0.6  # ширина глифа моноширинного шрифта в долях кегля

# Базовые шрифты PDF (Courier/Helvetica) не содержат кириллицы, поэтому при наличии
# в системе TTF-шрифтов с Unicode они регистрируются и используются вместо них.
MONO_TTF_FAMILIES: Tuple[Tuple[str, Tuple[str, str, str, str]], ...] = (
    ("DejaVuSansMono", (
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Oblique.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-BoldOblique.ttf",
    )),
    ("LiberationMono", (
        "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationMono-Italic.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationMono-BoldItalic.ttf",
    )),
)

SANS_TTF_FAMILIES: Tuple[Tuple[str, Tuple[str, str, str, str]], ...] = (
    ("DejaVuSans", (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
    )),
    ("LiberationSans", (
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf",
    )),
)


def _register_family(family: str, paths: Sequence[str]) -> Optional[Tuple[str, str, str, str]]:
    """Регистрирует TTF-семейство в reportlab; возвращает имена начертаний."""
    if not all(os.path.exists(path) for path in paths):
        return None
    names = (family, family + "-Bold", family + "-Italic", family + "-BoldItalic")
    try:
        for name, path in zip(names, paths):
            pdfmetrics.registerFont(TTFont(name, path))
        registerFontFamily(family, normal=names[0], bold=names[1],
                           italic=names[2], boldItalic=names[3])
    except Exception:  # pragma: no cover — битый или недоступный шрифт
        return None
    return names


def register_fonts() -> None:
    """Подключает Unicode-шрифты (кириллица в комментариях и подписях)."""
    global MONO, MONO_BOLD, MONO_ITALIC, MONO_BOLD_ITALIC, SANS, SANS_BOLD, CHAR_WIDTH_RATIO

    for family, paths in MONO_TTF_FAMILIES:
        names = _register_family(family, paths)
        if names:
            MONO, MONO_BOLD, MONO_ITALIC, MONO_BOLD_ITALIC = names
            CHAR_WIDTH_RATIO = pdfmetrics.stringWidth("0" * 10, MONO, 100) / 1000.0
            break
    else:
        print("[!] Моноширинный Unicode-шрифт не найден — кириллица в листинге "
              "может отображаться некорректно.")

    for family, paths in SANS_TTF_FAMILIES:
        names = _register_family(family, paths)
        if names:
            SANS, SANS_BOLD = names[0], names[1]
            break


INK = HexColor("#1b1f24")
MUTED = HexColor("#6b7480")
RULE = HexColor("#d5dae1")
GUTTER_BG = HexColor("#f2f4f7")
GUTTER_INK = HexColor("#9aa4b1")
ACCENT = HexColor("#1f6feb")
TITLE_BG = HexColor("#eef2f7")

PAGE_SIZES = {"a4": A4, "letter": letter}

# Файлы проекта, попадающие в листинг по умолчанию.
DEFAULT_SOURCES: Tuple[Tuple[str, str], ...] = (
    ("excel_smart_parser.py", "Ядро парсера — excel_smart_parser.py"),
    ("TEST/test_all_features.py", "Тесты — TEST/test_all_features.py"),
    ("excel_viewer/index.html", "Веб-просмотрщик — excel_viewer/index.html"),
)


# --------------------------------------------------------------------------- #
# Модель данных
# --------------------------------------------------------------------------- #

Span = Tuple[str, HexColor, str]  # (текст, цвет, шрифт)
Row = Tuple[Optional[int], List[Span]]  # (номер строки или None для переноса, спаны)


@dataclass
class Config:
    page_size: Tuple[float, float] = A4
    margin_left: float = 46.0
    margin_right: float = 34.0
    margin_top: float = 44.0
    margin_bottom: float = 40.0
    header_height: float = 24.0
    footer_height: float = 22.0
    font_size: float = 8.0
    leading_ratio: float = 1.20
    gutter_pad: float = 6.0
    title_block_rows: int = 4  # сколько строк кода «съедает» заголовок файла
    style_name: str = "default"
    project: str = "Excel Smart Parser"

    @property
    def leading(self) -> float:
        return round(self.font_size * self.leading_ratio, 3)

    @property
    def char_width(self) -> float:
        return self.font_size * CHAR_WIDTH_RATIO

    @property
    def page_width(self) -> float:
        return self.page_size[0]

    @property
    def page_height(self) -> float:
        return self.page_size[1]

    @property
    def body_top(self) -> float:
        return self.page_height - self.margin_top - self.header_height

    @property
    def body_bottom(self) -> float:
        return self.margin_bottom + self.footer_height

    @property
    def body_height(self) -> float:
        return self.body_top - self.body_bottom

    def gutter_width(self, max_lineno: int) -> float:
        digits = max(3, len(str(max_lineno)))
        return digits * self.char_width + 2 * self.gutter_pad

    def rows_per_page(self) -> int:
        return max(1, int(self.body_height // self.leading))

    def chars_per_row(self, gutter: float) -> int:
        avail = self.page_width - self.margin_left - self.margin_right - gutter - 4
        return max(20, int(avail // self.char_width))


@dataclass
class SourceFile:
    path: str
    title: str
    text: str = ""
    rows: List[Row] = field(default_factory=list)
    n_lines: int = 0
    n_bytes: int = 0
    sha256: str = ""
    language: str = ""
    start_page: int = 0
    page_count: int = 0


@dataclass
class Page:
    source_index: int
    rows: List[Row]
    is_first: bool
    max_lineno: int


# --------------------------------------------------------------------------- #
# Чтение и лексический разбор
# --------------------------------------------------------------------------- #

def load_source(root: str, rel_path: str, title: str) -> SourceFile:
    full = os.path.join(root, rel_path)
    with open(full, "rb") as fh:
        raw = fh.read()
    text = raw.decode("utf-8", errors="replace").replace("\r\n", "\n").replace("\r", "\n")
    src = SourceFile(path=rel_path, title=title, text=text)
    src.n_bytes = len(raw)
    src.n_lines = text.count("\n") + (0 if text.endswith("\n") else 1)
    src.sha256 = hashlib.sha256(raw).hexdigest()
    return src


def pick_font(bold: bool, italic: bool) -> str:
    if bold and italic:
        return MONO_BOLD_ITALIC
    if bold:
        return MONO_BOLD
    if italic:
        return MONO_ITALIC
    return MONO


def lex_to_lines(src: SourceFile, style) -> List[List[Span]]:
    """Разбирает файл лексером Pygments в список строк со спанами (текст, цвет, шрифт)."""
    try:
        lexer = get_lexer_for_filename(src.path, stripnl=False)
    except ClassNotFound:  # pragma: no cover
        lexer = guess_lexer(src.text)
    src.language = lexer.name

    cache: dict = {}
    lines: List[List[Span]] = [[]]
    for ttype, value in _pygments_lex(src.text, lexer):
        key = str(ttype)
        if key not in cache:
            st = style.style_for_token(ttype)
            color = HexColor("#" + (st.get("color") or "24292f"))
            cache[key] = (color, pick_font(bool(st.get("bold")), bool(st.get("italic"))))
        color, font = cache[key]
        chunks = value.split("\n")
        for i, chunk in enumerate(chunks):
            if chunk:
                lines[-1].append((chunk.expandtabs(4), color, font))
            if i < len(chunks) - 1:
                lines.append([])
    while lines and not lines[-1]:
        lines.pop()
    return lines


def wrap_line(spans: Sequence[Span], width: int) -> List[List[Span]]:
    """Режет длинную строку на несколько физических строк по ширине в символах."""
    rows: List[List[Span]] = []
    current: List[Span] = []
    used = 0
    for text, color, font in spans:
        while text:
            if used >= width:
                rows.append(current)
                current, used = [], 0
            room = width - used
            head, text = text[:room], text[room:]
            current.append((head, color, font))
            used += len(head)
    rows.append(current)
    return rows


def build_rows(src: SourceFile, cfg: Config, style) -> None:
    logical = lex_to_lines(src, style)
    gutter = cfg.gutter_width(max(1, len(logical)))
    width = cfg.chars_per_row(gutter)
    rows: List[Row] = []
    for number, spans in enumerate(logical, start=1):
        for i, piece in enumerate(wrap_line(spans, width)):
            rows.append((number if i == 0 else None, piece))
    src.rows = rows


# --------------------------------------------------------------------------- #
# Пагинация
# --------------------------------------------------------------------------- #

def paginate(sources: Sequence[SourceFile], cfg: Config) -> List[Page]:
    per_page = cfg.rows_per_page()
    pages: List[Page] = []
    for idx, src in enumerate(sources):
        max_lineno = max((n for n, _ in src.rows if n), default=1)
        cursor = 0
        first = True
        total = len(src.rows)
        while cursor < total or first:
            capacity = per_page - (cfg.title_block_rows if first else 0)
            chunk = src.rows[cursor:cursor + capacity]
            pages.append(Page(source_index=idx, rows=chunk, is_first=first, max_lineno=max_lineno))
            cursor += capacity
            first = False
    return pages


def front_matter_pages(n_sources: int) -> int:
    """Обложка + содержание (при большом числе файлов — несколько страниц содержания)."""
    return 1 + max(1, math.ceil(n_sources / 24))


def layout(sources: Sequence[SourceFile], cfg: Config, style) -> List[Page]:
    for src in sources:
        build_rows(src, cfg, style)
    pages = paginate(sources, cfg)
    offset = front_matter_pages(len(sources))
    for idx, src in enumerate(sources):
        own = [i for i, p in enumerate(pages) if p.source_index == idx]
        src.start_page = own[0] + 1 + offset
        src.page_count = len(own)
    return pages


def autofit(sources: Sequence[SourceFile], cfg: Config, style,
            target: Tuple[int, int], candidates: Sequence[float]) -> Tuple[Config, List[Page]]:
    """Подбирает наибольший кегль, при котором объём документа попадает в диапазон."""
    lo, hi = target
    attempts: List[Tuple[float, int, List[Page], Config]] = []
    for size in sorted(candidates, reverse=True):
        trial = Config(**{**cfg.__dict__, "font_size": size})
        pages = layout(sources, trial, style)
        total = len(pages) + front_matter_pages(len(sources))
        attempts.append((size, total, pages, trial))
        if lo <= total <= hi:
            return trial, pages
    # Ни один кегль не попал в диапазон — берём ближайший по числу страниц.
    size, total, pages, trial = min(attempts, key=lambda a: min(abs(a[1] - lo), abs(a[1] - hi)))
    print(f"[!] Диапазон {lo}-{hi} стр. недостижим; выбран кегль {size} pt → {total} стр.")
    layout(sources, trial, style)
    return trial, pages


# --------------------------------------------------------------------------- #
# Отрисовка
# --------------------------------------------------------------------------- #

class PdfRenderer:
    def __init__(self, path: str, cfg: Config, sources: Sequence[SourceFile], pages: Sequence[Page]):
        self.cfg = cfg
        self.sources = sources
        self.pages = pages
        self.total_pages = len(pages) + front_matter_pages(len(sources))
        self.canvas = _canvas.Canvas(path, pagesize=cfg.page_size)
        self.canvas.setTitle(f"{cfg.project} — полный листинг исходного кода")
        self.canvas.setAuthor(cfg.project)
        self.canvas.setSubject("Исходный код проекта в виде PDF-листинга")
        self.canvas.setCreator("tools/generate_code_pdf.py")

    # -- служебное --------------------------------------------------------- #

    def _rule(self, y: float, x0: Optional[float] = None, x1: Optional[float] = None,
              color=RULE, width: float = 0.5) -> None:
        cfg = self.cfg
        self.canvas.setStrokeColor(color)
        self.canvas.setLineWidth(width)
        self.canvas.line(x0 if x0 is not None else cfg.margin_left,
                         y,
                         x1 if x1 is not None else cfg.page_width - cfg.margin_right,
                         y)

    def _header(self, left: str, right: str) -> None:
        cfg = self.cfg
        c = self.canvas
        y = cfg.page_height - cfg.margin_top - 8
        c.setFont(SANS_BOLD, 8)
        c.setFillColor(MUTED)
        c.drawString(cfg.margin_left, y, left[:70])
        c.setFont(SANS, 8)
        c.drawRightString(cfg.page_width - cfg.margin_right, y, right[:50])
        self._rule(y - 5)

    def _footer(self, page_no: int) -> None:
        cfg = self.cfg
        c = self.canvas
        y = cfg.margin_bottom + 8
        self._rule(y + 10)
        c.setFont(SANS, 7.5)
        c.setFillColor(MUTED)
        c.drawString(cfg.margin_left, y, cfg.project)
        c.drawCentredString(cfg.page_width / 2, y, f"стр. {page_no} из {self.total_pages}")
        c.drawRightString(cfg.page_width - cfg.margin_right, y,
                          _dt.date.today().strftime("%d.%m.%Y"))

    # -- обложка ----------------------------------------------------------- #

    def cover(self) -> None:
        cfg = self.cfg
        c = self.canvas
        c.bookmarkPage("cover")
        c.addOutlineEntry("Обложка", "cover", level=0)

        top = cfg.page_height - 150
        c.setFillColor(ACCENT)
        c.rect(cfg.margin_left, top + 46, 96, 5, stroke=0, fill=1)

        c.setFillColor(INK)
        c.setFont(SANS_BOLD, 30)
        c.drawString(cfg.margin_left, top, cfg.project)
        c.setFont(SANS, 15)
        c.setFillColor(MUTED)
        c.drawString(cfg.margin_left, top - 26, "Полный листинг исходного кода")

        total_lines = sum(s.n_lines for s in self.sources)
        total_bytes = sum(s.n_bytes for s in self.sources)

        box_top = top - 66
        box_h = 34 + 22 * len(self.sources)
        c.setFillColor(TITLE_BG)
        c.rect(cfg.margin_left, box_top - box_h, cfg.page_width - cfg.margin_left - cfg.margin_right,
               box_h, stroke=0, fill=1)

        y = box_top - 22
        c.setFillColor(INK)
        c.setFont(SANS_BOLD, 10)
        c.drawString(cfg.margin_left + 14, y, "Состав документа")
        y -= 18
        for src in self.sources:
            c.setFont(MONO, 9)
            c.setFillColor(INK)
            c.drawString(cfg.margin_left + 14, y, src.path)
            c.setFont(SANS, 9)
            c.setFillColor(MUTED)
            c.drawRightString(cfg.page_width - cfg.margin_right - 14, y,
                              f"{src.n_lines} строк · {src.n_bytes / 1024:.1f} КБ · стр. {src.start_page}")
            y -= 22

        y = box_top - box_h - 40
        rows = [
            ("Всего строк кода", f"{total_lines}"),
            ("Общий объём исходников", f"{total_bytes / 1024:.1f} КБ"),
            ("Страниц в документе", f"{self.total_pages}"),
            ("Кегль листинга", f"{cfg.font_size} pt / интерлиньяж {cfg.leading} pt"),
            ("Дата сборки", _dt.datetime.now().strftime("%d.%m.%Y %H:%M")),
        ]
        for label, value in rows:
            c.setFont(SANS, 10)
            c.setFillColor(MUTED)
            c.drawString(cfg.margin_left, y, label)
            c.setFont(SANS_BOLD, 10)
            c.setFillColor(INK)
            c.drawString(cfg.margin_left + 190, y, value)
            y -= 17

        y -= 12
        c.setFont(SANS_BOLD, 9)
        c.setFillColor(INK)
        c.drawString(cfg.margin_left, y, "Контрольные суммы SHA-256")
        y -= 15
        for src in self.sources:
            c.setFont(MONO, 7)
            c.setFillColor(MUTED)
            c.drawString(cfg.margin_left, y, f"{src.sha256}  {src.path}")
            y -= 11

        c.setFont(SANS, 8)
        c.setFillColor(MUTED)
        c.drawString(cfg.margin_left, cfg.margin_bottom + 14,
                     "Документ сгенерирован автоматически: tools/generate_code_pdf.py")
        c.showPage()

    # -- содержание -------------------------------------------------------- #

    def toc(self) -> None:
        cfg = self.cfg
        c = self.canvas
        pages_needed = front_matter_pages(len(self.sources)) - 1
        chunk = math.ceil(len(self.sources) / pages_needed) if pages_needed else len(self.sources)
        page_no = 2
        for part in range(pages_needed):
            c.bookmarkPage(f"toc-{part}")
            if part == 0:
                c.addOutlineEntry("Содержание", "toc-0", level=0)
            self._header("Содержание", cfg.project)
            y = cfg.body_top - 24
            c.setFillColor(INK)
            c.setFont(SANS_BOLD, 18)
            c.drawString(cfg.margin_left, y, "Содержание")
            y -= 30
            for src in self.sources[part * chunk:(part + 1) * chunk]:
                c.setFont(SANS_BOLD, 11)
                c.setFillColor(INK)
                c.drawString(cfg.margin_left, y, src.title)
                c.setFont(SANS, 11)
                c.setFillColor(ACCENT)
                c.drawRightString(cfg.page_width - cfg.margin_right, y, str(src.start_page))
                y -= 14
                c.setFont(MONO, 8)
                c.setFillColor(MUTED)
                c.drawString(cfg.margin_left + 10, y,
                             f"{src.language} · {src.n_lines} строк · {src.page_count} стр. листинга")
                y -= 10
                self._rule(y, color=HexColor("#e8ebef"))
                y -= 16
            self._footer(page_no)
            page_no += 1
            c.showPage()

    # -- заголовок файла --------------------------------------------------- #

    def _title_block(self, src: SourceFile, y: float) -> float:
        cfg = self.cfg
        c = self.canvas
        height = cfg.title_block_rows * cfg.leading
        c.setFillColor(TITLE_BG)
        c.rect(cfg.margin_left, y - height + 6,
               cfg.page_width - cfg.margin_left - cfg.margin_right, height - 4, stroke=0, fill=1)
        c.setFillColor(ACCENT)
        c.rect(cfg.margin_left, y - height + 6, 3.2, height - 4, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont(SANS_BOLD, 11)
        c.drawString(cfg.margin_left + 12, y - 6, src.title)
        c.setFont(MONO, 7.5)
        c.setFillColor(MUTED)
        c.drawString(cfg.margin_left + 12, y - 18,
                     f"{src.path} · {src.language} · {src.n_lines} строк · "
                     f"{src.n_bytes / 1024:.1f} КБ · sha256:{src.sha256[:16]}")
        return y - height

    # -- страница листинга ------------------------------------------------- #

    def code_page(self, page: Page, page_no: int) -> None:
        cfg = self.cfg
        c = self.canvas
        src = self.sources[page.source_index]
        gutter = cfg.gutter_width(page.max_lineno)
        x_gutter = cfg.margin_left
        x_code = cfg.margin_left + gutter + 4

        if page.is_first:
            c.bookmarkPage(f"src-{page.source_index}")
            c.addOutlineEntry(src.title, f"src-{page.source_index}", level=0)

        self._header(src.path, f"{src.language} · {cfg.project}")

        y = cfg.body_top
        if page.is_first:
            y = self._title_block(src, y)

        rows_height = len(page.rows) * cfg.leading
        if page.rows:
            c.setFillColor(GUTTER_BG)
            c.rect(x_gutter - 2, y - rows_height, gutter, rows_height, stroke=0, fill=1)
            self._rule(y - rows_height, x0=x_gutter + gutter - 2, x1=x_gutter + gutter - 2,
                       color=RULE)
            c.setStrokeColor(RULE)
            c.setLineWidth(0.5)
            c.line(x_gutter + gutter - 2, y, x_gutter + gutter - 2, y - rows_height)

        baseline = y - cfg.font_size
        for lineno, spans in page.rows:
            if lineno is not None:
                c.setFont(MONO, cfg.font_size * 0.92)
                c.setFillColor(GUTTER_INK)
                c.drawRightString(x_gutter + gutter - cfg.gutter_pad - 2, baseline, str(lineno))
            else:
                c.setFont(MONO, cfg.font_size * 0.92)
                c.setFillColor(HexColor("#c3cad3"))
                c.drawRightString(x_gutter + gutter - cfg.gutter_pad - 2, baseline, "↳")
            offset = 0
            for text, color, font in spans:
                if text.strip():
                    c.setFont(font, cfg.font_size)
                    c.setFillColor(color)
                    c.drawString(x_code + offset * cfg.char_width, baseline, text)
                offset += len(text)
            baseline -= cfg.leading

        self._footer(page_no)
        c.showPage()

    # -- сборка ------------------------------------------------------------ #

    def render(self) -> None:
        self.cover()
        self.toc()
        page_no = front_matter_pages(len(self.sources)) + 1
        for page in self.pages:
            self.code_page(page, page_no)
            page_no += 1
        self.canvas.save()


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="PDF-листинг исходного кода проекта")
    ap.add_argument("--out", default="DOCS/excel_smart_parser_source_code.pdf",
                    help="путь к итоговому PDF")
    ap.add_argument("--root", default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    help="корень проекта")
    ap.add_argument("--file", action="append", default=None, metavar="PATH",
                    help="добавить файл в листинг (можно повторять); по умолчанию — файлы проекта")
    ap.add_argument("--target-pages", nargs=2, type=int, default=(30, 50), metavar=("MIN", "MAX"),
                    help="желаемый объём документа в страницах (по умолчанию 30 50)")
    ap.add_argument("--font-size", type=float, default=None,
                    help="фиксированный кегль листинга (отключает авто-подбор)")
    ap.add_argument("--page-size", choices=sorted(PAGE_SIZES), default="a4")
    ap.add_argument("--style", default="default", help="цветовая схема Pygments")
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    root = args.root
    if args.file:
        wanted: Iterable[Tuple[str, str]] = [(p, f"Листинг — {p}") for p in args.file]
    else:
        wanted = DEFAULT_SOURCES

    sources: List[SourceFile] = []
    for rel, title in wanted:
        if not os.path.exists(os.path.join(root, rel)):
            print(f"[!] Пропущен отсутствующий файл: {rel}")
            continue
        sources.append(load_source(root, rel, title))
    if not sources:
        sys.exit("Нет ни одного файла для листинга.")

    register_fonts()
    style = get_style_by_name(args.style)
    cfg = Config(page_size=PAGE_SIZES[args.page_size], style_name=args.style)

    if args.font_size:
        cfg = Config(**{**cfg.__dict__, "font_size": args.font_size})
        pages = layout(sources, cfg, style)
    else:
        cfg, pages = autofit(sources, cfg, style, tuple(args.target_pages),
                             candidates=(9.0, 8.5, 8.0, 7.5, 7.0, 6.5, 6.0, 5.5))

    out = args.out if os.path.isabs(args.out) else os.path.join(root, args.out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    renderer = PdfRenderer(out, cfg, sources, pages)
    renderer.render()

    total = len(pages) + front_matter_pages(len(sources))
    size_kb = os.path.getsize(out) / 1024
    print(f"PDF готов: {out}")
    print(f"  страниц: {total} (кегль {cfg.font_size} pt, {cfg.rows_per_page()} строк на странице)")
    print(f"  файлов: {len(sources)}, строк кода: {sum(s.n_lines for s in sources)}")
    print(f"  размер: {size_kb:.1f} КБ")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
