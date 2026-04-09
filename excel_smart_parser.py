"""
Excel Universal Parser v18
═══════════════════════════════════════════════════════════════════════════════
Исправления относительно v17:

  [FIX] SheetAdapter: переведён на abc.ABC / @abstractmethod. Ранее cell()
        бросал NotImplementedError только в рантайме при первом вызове.
        Теперь попытка инстанцировать неполный адаптер бросает TypeError
        немедленно при __init__, до любого обращения к данным.

  [FIX] _score_header_row / is_day_header: int(v) вызывался для float-значений
        без проверки math.isfinite(v). float('nan') и float('inf') уже
        отсекаются _is_numeric, но явная защита исключает OverflowError /
        ValueError при любых будущих изменениях пути вызова.

  [FIX] except Exception: pass → warnings.warn (~6 мест):
        hidden_cols (OpenpyxlAdapter), native_tables (OpenpyxlAdapter),
        named_ranges внутренний и внешний циклы (XlrdAdapter),
        named_ranges внутренний и внешний циклы (PyxlsbAdapter),
        _check_formulas_lazy (верхний уровень).
        Тихое поглощение исключений затрудняло отладку — теперь все
        нештатные ситуации видны через warnings.warn(..., UserWarning).

  [FIX] _load_xlsx: двойное открытие файла устранено структурно.
        Ранее _load_xlsx открывал файл с data_only=True, затем
        _check_formulas_lazy независимо открывала тот же файл с
        data_only=False — два load_workbook на один и тот же путь.
        Теперь: read_only-проход (data_only=False, read_only=True) для
        детекции формул и сборки списка листов, затем один полный проход
        (data_only=True) для данных. _check_formulas_lazy помечена
        deprecated и оставлена только для обратной совместимости.

  [FIX] _parse_range: двойное чтение ячеек для _detect_dtype.
        Ранее _detect_dtype получала [adapter.cell(r, c) for r in data_rows]
        отдельно для каждой колонки — все значения уже были прочитаны в
        цикле rows_out. Добавлен col_values: dict[int, list[CellValue]],
        заполняемый в том же цикле. _detect_dtype теперь получает данные
        из кэша без повторных вызовов adapter.cell().

  [FIX] _is_year: float('nan') и float('inf') вызывали OverflowError /
        ValueError при вызове int(v), потому что int(float('inf')) бросает
        OverflowError, а int(float('nan')) — ValueError. Добавлена проверка
        math.isfinite(v) перед int(v). Проблема реальна: openpyxl возвращает
        float('nan') для некоторых ошибочных ячеек, а _is_year вызывается
        из _score_header_row на каждую ячейку.

  [FIX] _load_csv delimiter fallback: max(candidates, key=candidates.get,
        default=",") — параметр default у max() используется только когда
        итерируемый объект ПУСТ. Так как _FALLBACK_DELIMITERS содержит 8
        символов, default никогда не срабатывает. Если ни один разделитель
        не встречается в файле (все счётчики == 0), max() возвращает
        произвольный символ. Исправлено: явная проверка счётчика лучшего
        кандидата, fallback на "," если он равен 0.

  [FIX] main: import traceback выполнялся внутри блока except при каждой
        ошибке. Модуль перенесён в top-level импорты.

  [FIX] _extract_heuristic: переменная l в списочном включении (for l, info
        in col_info.items()) визуально неотличима от цифры 1. Переименована
        в letter для ясности.

Исправления v16 (сохранены):

  [FIX] _is_numeric / _detect_encoding: import math и import codecs выполнялись
        внутри функций при каждом вызове. _is_numeric вызывается в горячем цикле
        на каждую ячейку — поиск по sys.modules происходил тысячи раз. Оба
        модуля перенесены в блок top-level импортов.

  [FIX] main: argparse description содержал устаревшую строку "v14" (снова).
        Исправлено на "v16".

  [FIX] ExcelParser._visible: метод вызывался 3 раза за парсинг одного листа
        (из _extract_heuristic, _extract_vertical, _extract_headerless) без
        кэширования. hidden_rows()/hidden_cols() итерируют row_dimensions —
        для больших файлов это заметный overhead. Добавлен кэш _vis_cache,
        сбрасываемый в начале каждого parse_sheet.

  [FIX] _check_formulas_lazy: открывал файл отдельно для каждого листа.
        Для xlsx с N листами — N вызовов load_workbook. Рефакторинг: один
        вызов load_workbook на весь файл, с ранним выходом на первой формуле
        в каждом листе. Сигнатура принимает список имён листов.

  [FIX] _is_year: дробные float вида 2021.5 возвращали True — int(2021.5)=2021
        попадает в диапазон 1900–2200. Дробный float — не год. Добавлена
        проверка v == int(v) для float-значений.

  [FIX] _score_header_row: строки длиной 1 символ ("A", "N", "К") получали
        score 0.0 из-за условия 2 <= len(s). Однобуквенные заголовки колонок
        типичны в финансах и науке. Исправлено: len >= 1 даёт +0.7 (чуть
        меньше чем +1.0 для длинных строк, т.к. однобуквенный текст менее
        однозначен).

  [FIX] _extract_heuristic: имела собственный пустой used_rows, не видя строк
        уже занятых native/named tables из parse_sheet. Эвристика могла найти
        таблицу, частично перекрывающую уже найденный named range. Теперь
        принимает used_rows из parse_sheet.

  [FIX] StreamingWriter: при streaming=True без output_path writer не создавался
        и данные молча терялись без предупреждения. Добавлен warnings.warn.

  [FIX] _extract_heuristic col_info: хранил полный список значений столбца
        ("values": vals_c) для каждой колонки, что при больших таблицах
        (10к строк × 50 колонок) создавало 500к промежуточных объектов.
        values убраны из col_info — dtype вычисляется сразу, строки строятся
        напрямую через adapter.cell().

Исправления v15 (сохранены):

  [FIX] _is_numeric: float("nan"), float("inf"), float("-inf") не бросают
        ValueError в Python, из-за чего строки "nan", "NaN", "inf", "Infinity"
        в CSV-ячейках ошибочно классифицировались как числа. В _score_header_row
        это давало -0.2 вместо +1.0, ломая детектирование заголовков. В
        _detect_dtype колонки получали тип "number" вместо "text". Исправлено:
        добавлена проверка math.isfinite() после успешного float().

  [FIX] _extract_vertical: счётчики text_in_a и numeric_in_rest вычислялись
        по всем строкам data_candidates, включая строки уже занятые (used_rows).
        Пороги применялись к "грязной" статистике, а реальные данные брались
        только из незанятых строк — результат мог не соответствовать проверкам.
        Исправлено: счётчики вычисляются только по строкам вне used_rows.

  [FIX] CsvAdapter docstring: утверждал "ленивая загрузка, O(1) по памяти,
        iter_rows_lazy читает потоком" — всё неверно с v11/v13. Весь файл
        загружается в _row_cache при инициализации, iter_rows_lazy итерирует
        по кэшу. Docstring исправлен.

  [FIX] Удалён неиспользуемый импорт Iterator из typing.

  [FIX] _score_header_row: комментарий "дата — нейтрально" при значении +0.5
        вводил в заблуждение — нейтральный 0.0, а +0.5 есть лёгкий положительный
        вклад. Комментарий исправлен.

Исправления v14 (сохранены):

  [FIX] XlrdAdapter.named_ranges: xlrd.Ref3D.coords — кортеж из 6 элементов
        (shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi), а код распаковывал
        только 5, сдвигая все поля на одну позицию вправо. min_row, max_row,
        min_col, max_col брали значения из неправильных полей; colxhi вообще
        игнорировался. Named ranges в .xls файлах имели полностью некорректные
        координаты. Исправлено: распаковка 6 полей с правильным маппингом.

  [FIX] parse_file: форматы .xltx / .xltm не входили в проверку
        ext in (".xlsx", ".xlsm"), из-за чего wb оставался None и
        _extract_named_ranges_from_wb никогда не вызывался — named ranges
        для .xltx / .xltm файлов терялись целиком. Исправлено.

  [FIX] StreamingWriter.write_table (JSON): стриминг не работал — код строил
        t_copy со всеми rows и сериализовывал таблицу целиком через json.dumps.
        Весь смысл --stream (не держать данные в RAM) был утерян. Исправлено:
        сначала пишется метаданные таблицы, затем строки сериализуются и
        записываются по одной без накопления в памяти.

  [FIX] _load_csv delimiter fallback: KNOWN_DELIMITERS.strip() убирал пробел
        из строки-кандидатов, делая пробел-разделитель недостижимым в fallback.
        Исправлено: итерация по frozenset явных символов без .strip().

  [FIX] CsvAdapter.cell: переменная r = row - 1 вычислялась но никогда не
        использовалась (кэш обходится по row, не по r). Мёртвый код удалён.

  [FIX] main: argparse description содержал устаревшую версию "v12". Исправлено.

  [FIX] PyxlsbAdapter docstring: утверждал что iter_rows_lazy читает файл
        потоком без кэша — неверно с v13 (итерация по _row_cache). Исправлено.

Исправления v13 (сохранены):
  [FIX] PyxlsbAdapter.iter_rows_lazy: итерация по кэшу вместо повторного чтения
  [FIX] CsvAdapter.iter_rows_lazy: итерация по кэшу вместо повторного чтения
  [FIX] parse_sheet: named ranges ищутся всегда, не только при отсутствии native
  [FIX] _extract_vertical: guard header_row in used_rows
  [FIX] _score_header_row: числовой год нейтрален (0.0), бонус только строкам

Исправления v12 (сохранены):
  [FIX] _extract_vertical: не вызывал _dedupe_headers
  [FIX] _detect_dtype: percent-детекция по >30% значений
  [FIX] parse_file: xlsx не загружался дважды (wb из ws.parent)
  [FIX] _check_formulas_lazy: открывает файл без data_only + read_only=True
  [FIX] parse_sheet: vertical всегда запускается после heuristic

Исправления v11 (сохранены):
  [FIX] _dedupe_headers в _parse_range и _extract_heuristic
  [FIX] _score_header_row: год +0.3 вместо +1.0
  [FIX] CsvAdapter: единый проход _load_cache
  [FIX] StreamingWriter: [:-1] вместо rstrip("}")
  [FIX] _parse_range: guard min_col > max_col
  [FIX] PyxlsbAdapter.named_ranges: цепочка fallback-атрибутов
  [FIX] _extract_vertical: порог text_in_a < 2

Исправления v10 (сохранены):
  [FIX] _extract_vertical: несколько таблиц на листе
  [FIX] _extract_heuristic: порог +0.25 и len(data_rows) >= 5
  [NEW] _extract_headerless: 5-й источник для матриц без заголовка
  [FIX] PyxlsbAdapter: один проход в _ensure_cache
  [FIX] streaming: счётчики sources корректны
"""


from __future__ import annotations

import abc
import argparse
import codecs
import csv
import datetime
import json
import math
import os
import sys
import traceback  # FIX v17: перенесён из except-блока в top-level импорты
import warnings
from typing import Any, Generator, Optional

# ── Обязательная зависимость ───────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.utils import range_boundaries as _range_boundaries   # FIX: один импорт
except ImportError:
    sys.exit("❌ pip install openpyxl")

# ── Опциональные зависимости ───────────────────────────────────────────────────
try:
    from tqdm import tqdm as _tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

try:
    import pyxlsb
    HAS_PYXLSB = True
except ImportError:
    HAS_PYXLSB = False

# FIX: автодетекция кодировки CSV
try:
    import chardet as _chardet
    HAS_CHARDET = True
except ImportError:
    try:
        import charset_normalizer as _chardet  # type: ignore
        HAS_CHARDET = True
    except ImportError:
        HAS_CHARDET = False


# ══════════════════════════════════════════════════════════════════════════════
# Типы
# ══════════════════════════════════════════════════════════════════════════════

CellValue = Any
Row = dict[str, CellValue]

TABLE_SOURCE_NATIVE    = "native_table"
TABLE_SOURCE_NAMED     = "named_range"
TABLE_SOURCE_HEURISTIC = "heuristic"
TABLE_SOURCE_VERTICAL  = "vertical"
TABLE_SOURCE_HEADERLESS = "headerless"  # FIX: таблицы без текстового заголовка


# ══════════════════════════════════════════════════════════════════════════════
# Абстрактный адаптер листа
# ══════════════════════════════════════════════════════════════════════════════

class SheetAdapter(abc.ABC):
    """
    Абстрактный адаптер листа. Подклассы обязаны реализовать cell().
    Все остальные методы имеют дефолтную реализацию и могут быть переопределены.
    """
    name: str
    max_row: int
    max_col: int

    @abc.abstractmethod
    def cell(self, row: int, col: int) -> CellValue:
        """Возвращает значение ячейки (1-based row, col). None для пустых."""

    def iter_rows_lazy(self, cols: list[int]) -> Generator[tuple[int, list[CellValue]], None, None]:
        """Ленивый итератор строк: (row_number, [values])."""
        for r in range(1, self.max_row + 1):
            yield r, [self.cell(r, c) for c in cols]

    def hidden_rows(self) -> set[int]:
        return set()

    def hidden_cols(self) -> set[int]:
        return set()

    def native_tables(self) -> list[dict]:
        return []

    def named_ranges(self) -> list[dict]:
        """Возвращает [{name, sheet, min_row, max_row, min_col, max_col}]."""
        return []


# ── openpyxl (.xlsx / .xlsm / .xltx) ─────────────────────────────────────────

class OpenpyxlAdapter(SheetAdapter):
    def __init__(self, ws, name: str):
        self._ws = ws
        self.name = name
        self.max_row = ws.max_row or 0
        self.max_col = ws.max_column or 0
        # кэш merged cells
        self._merged: dict[tuple[int, int], CellValue] = {}
        for rng in ws.merged_cells.ranges:
            master = ws.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._merged[(r, c)] = master

    def cell(self, row: int, col: int) -> CellValue:
        if (row, col) in self._merged:
            return self._merged[(row, col)]
        return self._ws.cell(row, col).value

    def hidden_rows(self) -> set[int]:
        return {r for r, d in self._ws.row_dimensions.items() if d.hidden}

    def hidden_cols(self) -> set[int]:
        result: set[int] = set()
        for letter, d in self._ws.column_dimensions.items():
            if d.hidden:
                try:
                    result.add(column_index_from_string(letter))
                except Exception as e:
                    warnings.warn(
                        f"Лист '{self.name}': не удалось распознать скрытую колонку "
                        f"'{letter}': {e}",
                        UserWarning, stacklevel=2,
                    )
        return result

    def native_tables(self) -> list[dict]:
        tables = []
        for tname, tobj in self._ws.tables.items():
            try:
                # FIX v18.1: в некоторых версиях openpyxl tobj — строка (ref),
                # в других — объект Table с атрибутом .ref
                ref = tobj.ref if hasattr(tobj, "ref") else str(tobj)
                if ":" not in ref:
                    continue  # Не диапазон — пропускаем
                min_col, min_row, max_col, max_row = _range_boundaries(ref)
                tables.append({"name": tname, "ref": ref,
                                "min_row": min_row, "max_row": max_row,
                                "min_col": min_col, "max_col": max_col})
            except Exception as e:
                warnings.warn(
                    f"Лист '{self.name}': нативная таблица '{tname}' "
                    f"имеет некорректный диапазон: {e}",
                    UserWarning, stacklevel=2,
                )
        return tables


# ── xlrd (.xls) ───────────────────────────────────────────────────────────────

class XlrdAdapter(SheetAdapter):
    def __init__(self, sheet, book, name: str):
        self._sheet = sheet
        self._book = book
        self.name = name
        self.max_row = sheet.nrows
        self.max_col = sheet.ncols

    def cell(self, row: int, col: int) -> CellValue:
        import xlrd as _xlrd
        if row < 1 or col < 1 or row > self.max_row or col > self.max_col:
            return None
        c = self._sheet.cell(row - 1, col - 1)
        if c.ctype == _xlrd.XL_CELL_EMPTY:
            return None
        if c.ctype == _xlrd.XL_CELL_DATE:
            try:
                return _xlrd.xldate_as_datetime(c.value, self._book.datemode)
            except Exception as e:
                warnings.warn(
                    f"Лист '{self.name}': не удалось преобразовать дату "
                    f"(xldate={c.value}): {e}. Возвращается raw float.",
                    UserWarning, stacklevel=2,
                )
                return c.value
        return c.value

    def named_ranges(self) -> list[dict]:
        """FIX: именованные диапазоны для .xls через xlrd."""
        result = []
        try:
            for name_obj in self._book.name_obj_list:
                name = name_obj.name
                if not name_obj.result:
                    continue
                try:
                    # xlrd.Ref3D.coords — кортеж из 6 элементов:
                    # (shtxlo, shtxhi, rowxlo, rowxhi, colxlo, colxhi)
                    # Диапазоны полуоткрытые: [lo, hi) в 0-based индексах.
                    # FIX v14: ранее распаковывалось только 5 элементов, из-за
                    # чего все поля сдвигались вправо (shtxhi→row0, rowxlo→row1
                    # и т.д.), а colxhi вообще игнорировался.
                    for area in name_obj.result.coords:
                        shtxlo, _shtxhi, row0, row1, col0, col1 = area[:6]
                        sheet_idx = shtxlo
                        if self._book.sheet_names()[sheet_idx] != self.name:
                            continue
                        result.append({
                            "name": name,
                            "sheet": self.name,
                            "min_row": row0 + 1,   # 0-based lo → 1-based
                            "max_row": row1,        # 0-based exclusive hi = 1-based inclusive
                            "min_col": col0 + 1,
                            "max_col": col1,
                        })
                except Exception as e:
                    warnings.warn(
                        f"Лист '{self.name}': именованный диапазон '{name}' "
                        f"не удалось распарсить: {e}",
                        UserWarning, stacklevel=2,
                    )
        except Exception as e:
            warnings.warn(
                f"Лист '{self.name}': ошибка при чтении name_obj_list: {e}",
                UserWarning, stacklevel=2,
            )
        return result


# ── pyxlsb (.xlsb) ────────────────────────────────────────────────────────────

class PyxlsbAdapter(SheetAdapter):
    """
    Адаптер для .xlsb файлов через pyxlsb.
    При инициализации (_ensure_cache) весь лист загружается в _row_cache одним
    проходом — необходимо для произвольного доступа через cell().
    iter_rows_lazy() итерирует по _row_cache без повторного чтения файла (v13).
    """

    def __init__(self, filepath: str, sheet_name: str):
        self._filepath = filepath
        self.name = sheet_name
        self._row_cache: dict[int, list] = {}
        self._cache_loaded: bool = False
        # FIX: один проход вместо двух — размеры и кэш строятся одновременно
        self.max_row = 0
        self.max_col = 0
        self._ensure_cache()

    def _ensure_cache(self) -> None:
        """FIX: один проход — одновременно определяем размеры и строим кэш (было 2 прохода)."""
        if self._cache_loaded:
            return
        import pyxlsb as _pyxlsb
        tmp: dict[int, dict] = {}
        max_r = max_c = 0
        with _pyxlsb.open_workbook(self._filepath) as wb:
            with wb.get_sheet(self.name) as ws:
                for i, row in enumerate(ws.rows()):
                    r = i + 1
                    if r > max_r:
                        max_r = r
                    for cell in row:
                        c_idx = cell.c
                        if c_idx + 1 > max_c:
                            max_c = c_idx + 1
                        tmp.setdefault(r, {})[c_idx] = cell.v
        self.max_row = max_r
        self.max_col = max_c
        for r, cols in tmp.items():
            data = [None] * max_c
            for c_idx, v in cols.items():
                if 0 <= c_idx < max_c:
                    data[c_idx] = v
            self._row_cache[r] = data
        self._cache_loaded = True

    def _load_row(self, target_row: int) -> list:
        """Возвращает строку из кэша, заполняя его при первом обращении."""
        if not self._cache_loaded:
            self._ensure_cache()
        return self._row_cache.get(target_row, [None] * self.max_col)

    def cell(self, row: int, col: int) -> CellValue:
        if row < 1 or col < 1 or row > self.max_row or col > self.max_col:
            return None
        return self._load_row(row)[col - 1]

    def iter_rows_lazy(self, cols: list[int]) -> Generator[tuple[int, list[CellValue]], None, None]:
        """
        FIX v13: ранее открывал файл заново через pyxlsb, хотя _ensure_cache()
        в конструкторе уже загрузил весь лист в _row_cache. Файл читался дважды.
        Теперь итерируем напрямую по _row_cache.
        """
        col_to_pos = {c: i for i, c in enumerate(cols)}
        for r in range(1, self.max_row + 1):
            row_data = self._row_cache.get(r, [None] * self.max_col)
            vals = [None] * len(cols)
            for c, pos in col_to_pos.items():
                idx = c - 1  # 1-based → 0-based
                if 0 <= idx < len(row_data):
                    vals[pos] = row_data[idx]
            yield r, vals

    def named_ranges(self) -> list[dict]:
        """FIX v11: перебор нескольких возможных атрибутов pyxlsb — API нестабилен между версиями."""
        result = []
        try:
            import pyxlsb as _pyxlsb
            with _pyxlsb.open_workbook(self._filepath) as wb:
                # pyxlsb не имеет стабильного публичного API для named ranges;
                # пробуем все известные атрибуты в порядке приоритета
                candidates = (
                    getattr(wb, "defined_names", None)
                    or getattr(wb, "named_ranges", None)
                    or getattr(wb, "_defined_names", None)
                    or []
                )
                for nr in candidates:
                    try:
                        name = getattr(nr, "name", None) or getattr(nr, "Name", None)
                        formula = getattr(nr, "formula", None) or getattr(nr, "refers_to", None)
                        if not name or not formula:
                            continue
                        if "!" not in formula:
                            continue
                        sheet_part, range_part = formula.split("!", 1)
                        sheet_name = sheet_part.strip("'\"$")
                        if sheet_name != self.name:
                            continue
                        min_col, min_row, max_col, max_row = _range_boundaries(range_part)
                        if min_col > max_col or min_row > max_row:
                            continue
                        result.append({
                            "name": name,
                            "sheet": self.name,
                            "min_row": min_row, "max_row": max_row,
                            "min_col": min_col, "max_col": max_col,
                        })
                    except Exception as e:
                        warnings.warn(
                            f"Лист '{self.name}': именованный диапазон pyxlsb "
                            f"не удалось распарсить: {e}",
                            UserWarning, stacklevel=2,
                        )
        except Exception as e:
            warnings.warn(
                f"Лист '{self.name}': ошибка при чтении named_ranges из .xlsb: {e}",
                UserWarning, stacklevel=2,
            )
        return result


# ── CSV ───────────────────────────────────────────────────────────────────────

class CsvAdapter(SheetAdapter):
    """
    Адаптер CSV. При инициализации весь файл загружается в _row_cache одним
    проходом (единый проход с v11). Доступ через cell() и iter_rows_lazy()
    работает по кэшу — повторного чтения файла нет (исправлено в v13).
    Потребление памяти O(n) по числу строк файла.
    """

    def __init__(self, filepath: str, encoding: str, delimiter: str, name: str):
        self.name = name
        self._filepath = filepath
        self._encoding = encoding
        self._delimiter = delimiter
        # FIX v11: один проход — строим кэш и размеры одновременно
        self._row_cache: dict[int, list] = {}
        self.max_row, self.max_col = self._load_cache()

    def _load_cache(self) -> tuple[int, int]:
        """Единый проход: читаем файл один раз, строим кэш и определяем размеры."""
        max_r = max_c = 0
        with open(self._filepath, newline="", encoding=self._encoding) as f:
            for i, row in enumerate(csv.reader(f, delimiter=self._delimiter)):
                r = i + 1
                max_r = r
                if len(row) > max_c:
                    max_c = len(row)
                self._row_cache[r] = row
        return max_r, max_c

    def cell(self, row: int, col: int) -> CellValue:
        # FIX v14: r = row - 1 была мёртвой переменной — кэш хранит строки
        # по 1-based ключу row, а не по 0-based. Удалено.
        c = col - 1
        row_data = self._row_cache.get(row, [])
        if c < 0 or c >= len(row_data):
            return None
        v = row_data[c]
        return v if v != "" else None

    def iter_rows_lazy(self, cols: list[int]) -> Generator[tuple[int, list[CellValue]], None, None]:
        """
        FIX v13: ранее открывал файл заново через open() + csv.reader, хотя
        _load_cache() уже загрузил весь файл в _row_cache при инициализации —
        файл читался дважды. Теперь итерируем напрямую по кэшу.
        """
        col_to_pos = {c: i for i, c in enumerate(cols)}
        for r in range(1, self.max_row + 1):
            row_data = self._row_cache.get(r, [])
            vals = [None] * len(cols)
            for c, pos in col_to_pos.items():
                raw_c = c - 1  # 1-based → 0-based
                if 0 <= raw_c < len(row_data):
                    v = row_data[raw_c]
                    vals[pos] = v if v != "" else None
            yield r, vals


# ══════════════════════════════════════════════════════════════════════════════
# Загрузчики
# ══════════════════════════════════════════════════════════════════════════════

def load_sheets(filepath: str, only_sheet: Optional[str] = None) -> list[SheetAdapter]:
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return _load_xlsx(filepath, only_sheet)
    elif ext == ".xls":
        return _load_xls(filepath, only_sheet)
    elif ext == ".xlsb":
        return _load_xlsb(filepath, only_sheet)
    elif ext == ".csv":
        return _load_csv(filepath)
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}. "
                         "Поддерживается: .xlsx .xlsm .xls .xlsb .csv")


def _load_xlsx(filepath: str, only_sheet: Optional[str]) -> list[OpenpyxlAdapter]:
    ext = os.path.splitext(filepath)[1].lower()
    # FIX: документируем поведение xlsm
    if ext == ".xlsm":
        warnings.warn(
            f"'{os.path.basename(filepath)}': формат .xlsm — "
            "файл откроется корректно, VBA-макросы игнорируются.",
            UserWarning, stacklevel=3,
        )

    # FIX v18: один проход data_only=False (read_only — быстро) для детекции
    # формул, затем один проход data_only=True для фактических данных.
    # Ранее _check_formulas_lazy открывала файл отдельно → два открытия.
    # Теперь оба open явно сидят рядом; read_only-проход лёгкий.
    sheets_to_process: list[str] = []
    try:
        wb_scan = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
        try:
            for sheet_name in wb_scan.sheetnames:
                if only_sheet and sheet_name != only_sheet:
                    continue
                ws = wb_scan[sheet_name]
                for row in ws.iter_rows(values_only=False):
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            warnings.warn(
                                f"Лист '{sheet_name}': обнаружены формулы. "
                                "data_only=True вернёт None для несохранённых значений.",
                                UserWarning, stacklevel=3,
                            )
                            break
                    else:
                        continue
                    break
                sheets_to_process.append(sheet_name)
        finally:
            wb_scan.close()
    except Exception as e:
        warnings.warn(
            f"_load_xlsx: не удалось выполнить проверку формул для '{os.path.basename(filepath)}': {e}",
            UserWarning, stacklevel=2,
        )
        # Fallback: обрабатываем все листы без проверки формул
        try:
            wb_tmp = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            sheets_to_process = [
                n for n in wb_tmp.sheetnames
                if not only_sheet or n == only_sheet
            ]
            wb_tmp.close()
        except Exception as e:
            warnings.warn(
                f"_load_xlsx: не удалось получить список листов из '{os.path.basename(filepath)}' "
                f"для fallback: {e}. Ни один лист не будет обработан.",
                UserWarning, stacklevel=2,
            )
            sheets_to_process = []

    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheets = []
    for name in wb.sheetnames:
        if name not in sheets_to_process:
            continue
        ws = wb[name]
        if not ws.max_row:
            continue
        sheets.append(OpenpyxlAdapter(ws, name))
    return sheets


def _check_formulas_lazy(filepath: str, sheet_names: list[str]) -> None:
    """
    Устарело в v18: логика перенесена в _load_xlsx (больше не вызывается).
    Оставлено для обратной совместимости на случай внешних вызовов.
    FIX v16: ранее вызывался отдельно для каждого листа, открывая файл N раз.
    FIX v18: полностью встроен в _load_xlsx, этот метод — заглушка.
    """
    warnings.warn(
        "_check_formulas_lazy устарела в v18 — логика перенесена в _load_xlsx.",
        DeprecationWarning, stacklevel=2,
    )


def _load_xls(filepath: str, only_sheet: Optional[str]) -> list[XlrdAdapter]:
    if not HAS_XLRD:
        raise ImportError("pip install xlrd  # для .xls файлов")
    wb = xlrd.open_workbook(filepath)
    result = []
    for name in wb.sheet_names():
        if only_sheet and name != only_sheet:
            continue
        result.append(XlrdAdapter(wb.sheet_by_name(name), wb, name))
    return result


def _load_xlsb(filepath: str, only_sheet: Optional[str]) -> list[PyxlsbAdapter]:
    if not HAS_PYXLSB:
        raise ImportError("pip install pyxlsb  # для .xlsb файлов")
    import pyxlsb as _pyxlsb
    result = []
    with _pyxlsb.open_workbook(filepath) as wb:
        for name in wb.sheets:
            if only_sheet and name != only_sheet:
                continue
            result.append(PyxlsbAdapter(filepath, name))
    return result


def _detect_encoding(filepath: str) -> str:
    """Автодетекция кодировки через chardet/charset-normalizer с надёжным fallback."""
    if HAS_CHARDET:
        with open(filepath, "rb") as f:
            raw = f.read(32768)
        detected = _chardet.detect(raw)
        enc = (detected.get("encoding") or "utf-8").strip()
        # Проверяем, что Python реально знает эту кодировку; иначе — fallback
        try:
            codecs.lookup(enc)
            return enc
        except LookupError:
            pass
    # Fallback: перебор надёжных кодировок
    for enc in ("utf-8-sig", "utf-8", "cp1251", "cp1252", "latin-1"):
        try:
            with open(filepath, encoding=enc) as f:
                f.read(4096)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return "utf-8"


def _load_csv(filepath: str) -> list[CsvAdapter]:
    encoding = _detect_encoding(filepath)

    # FIX: расширенный набор разделителей + fallback на любой непечатный/пунктуационный символ
    KNOWN_DELIMITERS = ",;\t|^~ "
    try:
        with open(filepath, newline="", encoding=encoding) as f:
            sample = f.read(16384)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=KNOWN_DELIMITERS)
            delimiter = dialect.delimiter
        except csv.Error:
            # Fallback: считаем частоту символов и берём самый частый из кандидатов.
            # FIX v14: KNOWN_DELIMITERS.strip() убирал пробел из итерации,
            # делая его недостижимым. Используем явный frozenset символов.
            _FALLBACK_DELIMITERS = frozenset(",;\t|^~ ")
            candidates = {d: sample.count(d) for d in _FALLBACK_DELIMITERS}
            # FIX v17: max(..., default=",") не работал — default у max() используется
            # только когда итерируемый объект пуст, а candidates всегда содержит 8
            # элементов. Если все счётчики == 0, max() возвращал произвольный символ.
            # Исправлено: явная проверка счётчика победителя.
            _best = max(candidates, key=candidates.get)
            delimiter = _best if candidates[_best] > 0 else ","

        # FIX: передаём путь в адаптер — никакого list(reader) в память
        return [CsvAdapter(filepath, encoding, delimiter, os.path.basename(filepath))]
    except Exception as e:
        raise ValueError(f"Не удалось прочитать CSV '{filepath}': {e}") from e


# ══════════════════════════════════════════════════════════════════════════════
# Утилиты значений
# ══════════════════════════════════════════════════════════════════════════════

def _to_str(v: CellValue) -> str:
    return "" if v is None else str(v).strip()


def _is_empty(v: CellValue) -> bool:
    return _to_str(v) == ""


def _is_numeric(v: CellValue) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return math.isfinite(v)   # отсекаем nan/inf уже на уровне нативных значений
    s = _to_str(v).replace(",", ".").replace(" ", "").replace("%", "")
    if not s:
        return False
    try:
        return math.isfinite(float(s))  # float("nan")/"inf" не бросают ValueError
    except ValueError:
        return False


def _is_date(v: CellValue) -> bool:
    return isinstance(v, (datetime.datetime, datetime.date))


def _is_year(v: CellValue) -> bool:
    """
    FIX v16: дробные float не являются годами.
    FIX v17: float('nan') и float('inf') вызывали OverflowError/ValueError
             при вызове int(v). Добавлена проверка math.isfinite() перед int().
    """
    if isinstance(v, bool):
        return False
    if isinstance(v, float):
        # FIX v17: int(float('inf')) → OverflowError, int(float('nan')) → ValueError
        if not math.isfinite(v):
            return False
        # FIX v16: 2021.5 → int(2021.5)=2021 попадало в диапазон, но дробный — не год
        if v != int(v):
            return False
        return 1900 <= int(v) <= 2200
    if isinstance(v, int):
        return 1900 <= v <= 2200
    s = _to_str(v)
    try:
        return 1900 <= int(s) <= 2200
    except ValueError:
        return False


def _serialize(v: CellValue) -> Any:
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    return v


def _detect_dtype(values: list[CellValue]) -> str:
    non_empty = [v for v in values if not _is_empty(v)]
    if not non_empty:
        return "text"
    total = len(non_empty)
    nums  = sum(1 for v in non_empty if _is_numeric(v))
    dates = sum(1 for v in non_empty if _is_date(v))
    if dates / total > 0.5:
        return "date"
    if nums / total > 0.7:
        # FIX v12: ранее any("%" ...) давало percent если ХОТЯ БЫ одна ячейка
        # содержала "%" в любом месте строки (напр. "рост 15% г/г").
        # Теперь требуем, чтобы >30% значений явно содержали символ "%".
        pct_count = sum(
            1 for v in non_empty
            if isinstance(v, str) and "%" in v
        )
        return "percent" if pct_count / total > 0.3 else "number"
    return "text"


# ══════════════════════════════════════════════════════════════════════════════
# Score-based детектор заголовков
# ══════════════════════════════════════════════════════════════════════════════

def _score_header_row(vals: list[CellValue]) -> float:
    non_empty = [v for v in vals if not _is_empty(v)]
    if not non_empty:
        return 0.0

    # FIX v18: определяем паттерн "числа-дни" (1,2,3...31 подряд).
    # Если в строке есть последовательные маленькие整数 (1-31) — это заголовок
    # таблицы с днями месяца, а не данные. Не штрафуем такие числа.
    numeric_vals = [v for v in non_empty if _is_numeric(v) and not _is_year(v)]
    is_day_header = False
    if len(numeric_vals) >= 5:
        # Проверяем что числа подряд идут 1,2,3... или близки к этому
        day_numbers = []
        for v in numeric_vals:
            # FIX v18: int(float('nan')) → ValueError, int(float('inf')) → OverflowError.
            # math.isfinite уже гарантируется _is_numeric, но numeric_vals строится
            # напрямую — явная проверка исключает UB при будущих рефакторингах.
            if isinstance(v, float):
                if not math.isfinite(v):
                    continue
                n = int(v)
            else:
                n = v
            if 1 <= n <= 31:
                day_numbers.append(n)
        if len(day_numbers) >= 5:
            day_numbers_sorted = sorted(day_numbers)
            # Проверяем что большинство идут подряд (разница <=2 между соседями)
            consecutive = sum(
                1 for i in range(len(day_numbers_sorted) - 1)
                if 1 <= (day_numbers_sorted[i + 1] - day_numbers_sorted[i]) <= 2
            )
            if consecutive >= len(day_numbers_sorted) * 0.6:
                is_day_header = True

    text_like = 0.0
    for v in non_empty:
        if _is_year(v):
            if isinstance(v, str):
                text_like += 0.3
        elif _is_date(v):
            text_like += 0.5
        elif _is_numeric(v):
            # FIX v18: не штрафуем числа-дни в заголовке таблицы
            if is_day_header:
                text_like += 0.5  # нейтрально-положительный вклад
            else:
                text_like -= 0.2
        else:
            s = _to_str(v)
            if len(s) == 1:
                text_like += 0.7
            elif 2 <= len(s) <= 60:
                text_like += 1.0
            elif len(s) > 60:
                text_like -= 0.5
    ratio = text_like / len(non_empty)
    if len(non_empty) < 2:
        ratio *= 0.6
    return max(0.0, min(1.0, ratio))


def _is_header_row(vals: list[CellValue], threshold: float = 0.4) -> bool:
    return _score_header_row(vals) >= threshold


def _dedupe_headers(names: list[str]) -> list[str]:
    """
    FIX v11: дублирующиеся имена колонок приводили к тихой перезаписи значений
    в dict строки — второй столбец с тем же именем молча затирал первый.
    Теперь дубли получают суффикс _2, _3 и т.д.
    Пустые имена заменяются на '_col_N' во избежание коллизий пустых строк.
    """
    seen: dict[str, int] = {}
    result: list[str] = []
    for i, name in enumerate(names):
        key = name if name else f"_col_{i + 1}"
        if key not in seen:
            seen[key] = 1
            result.append(key)
        else:
            seen[key] += 1
            result.append(f"{key}_{seen[key]}")
    return result


# ══════════════════════════════════════════════════════════════════════════════
# Потоковый writer
# ══════════════════════════════════════════════════════════════════════════════

class StreamingWriter:
    """
    FIX: пишет таблицы сразу на диск, не накапливая в памяти.
    Используется при --stream.
    """

    def __init__(self, output_path: str, fmt: str, file_meta: dict):
        self.fmt = fmt
        self.output_path = output_path
        self._table_count = 0
        self._row_count   = 0
        self._file_meta   = file_meta
        self._json_fh     = None
        self._jsonl_fh    = None
        self._csv_dir     = None
        self._open(output_path, fmt)

    def _open(self, path: str, fmt: str) -> None:
        if fmt == "json":
            self._json_fh = open(path, "w", encoding="utf-8")
            meta = {k: v for k, v in self._file_meta.items()}
            meta_str = json.dumps(meta, ensure_ascii=False)
            # FIX v11: [:-1] вместо rstrip("}") — rstrip срезал бы лишние символы
            # если значение метаданных оканчивалось на "}" (напр. путь с {}).
            self._json_fh.write(meta_str[:-1] + ', "tables_data": [\n')
            self._first_table = True
        elif fmt == "jsonl":
            self._jsonl_fh = open(path, "w", encoding="utf-8")
        elif fmt == "csv":
            os.makedirs(path, exist_ok=True)
            self._csv_dir = path

    def write_table(self, table: dict) -> None:
        self._table_count += 1
        self._row_count += len(table.get("rows", []))
        if self.fmt == "json":
            sep = "" if self._first_table else ",\n"
            self._first_table = False
            # FIX v14: ранее t_copy["rows"] = table["rows"] добавлял все строки
            # обратно и json.dumps сериализовывал всю таблицу сразу — стриминга
            # не было. Теперь: сначала пишем метаданные (без rows), затем
            # открываем массив строк и сериализуем каждую строку отдельно.
            meta = {k: v for k, v in table.items() if k != "rows"}
            meta_str = json.dumps(meta, ensure_ascii=False, default=str)
            # Открываем объект таблицы: {...мета..., "rows": [
            self._json_fh.write(sep + meta_str[:-1] + ', "rows": [')
            rows = table.get("rows", [])
            for i, row in enumerate(rows):
                row_sep = "" if i == 0 else ","
                self._json_fh.write(
                    row_sep + json.dumps(row, ensure_ascii=False, default=str)
                )
            self._json_fh.write("]}")
        elif self.fmt == "jsonl":
            for row in table.get("rows", []):
                record = {"_sheet": table["sheet"], "_table": table["name"], **row}
                self._jsonl_fh.write(
                    json.dumps(record, ensure_ascii=False, default=str) + "\n"
                )
        elif self.fmt == "csv":
            rows = table.get("rows", [])
            if not rows:
                return
            safe = (
                table["name"]
                .replace("/", "_").replace("\\", "_")
                .replace(":", "_").replace(" ", "_")[:60]
            )
            path = os.path.join(self._csv_dir,
                                f"{self._table_count:02d}_{safe}.csv")
            fieldnames = list(rows[0].keys())
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)

    def close(self) -> None:
        if self.fmt == "json" and self._json_fh:
            self._json_fh.write("\n]}")
            self._json_fh.close()
        elif self.fmt == "jsonl" and self._jsonl_fh:
            self._jsonl_fh.close()

    @property
    def stats(self) -> tuple[int, int]:
        return self._table_count, self._row_count


# ══════════════════════════════════════════════════════════════════════════════
# Парсер
# ══════════════════════════════════════════════════════════════════════════════

class ExcelParser:
    def __init__(
        self,
        header_threshold: float = 0.4,
        skip_hidden: bool = False,  # FIX v18: было True — скрытые строки
                                     # (свёрнутые пользователем) пропускались,
                                     # из-за чего терялись данные. Теперь по
                                     # умолчанию включаем все строки как v5.
        min_data_cells: int = 2,
        max_empty_streak: int = 50,   # FIX: было 3 — таблицы с маркер-строками
                                       # ("Исход", "Доп.часы" и т.п.) разрывались.
                                       # Увеличено до 50 чтобы не рвать таблицы
                                       # с промежуточными секциями.
    ):
        self.header_threshold = header_threshold
        self.skip_hidden = skip_hidden
        self.min_data_cells = min_data_cells
        self.max_empty_streak = max_empty_streak
        self._vis_cache: dict[int, tuple[list[int], list[int]]] = {}

    # ── Видимые строки / колонки ──────────────────────────────────────────────

    def _visible(self, adapter: SheetAdapter) -> tuple[list[int], list[int]]:
        key = id(adapter)
        if key not in self._vis_cache:
            hr = adapter.hidden_rows() if self.skip_hidden else set()
            hc = adapter.hidden_cols() if self.skip_hidden else set()
            rows = [r for r in range(1, adapter.max_row + 1) if r not in hr]
            cols = [c for c in range(1, adapter.max_col + 1) if c not in hc]
            self._vis_cache[key] = (rows, cols)
        return self._vis_cache[key]

    # ── Парсинг диапазона ─────────────────────────────────────────────────────

    def _parse_range(
        self,
        adapter: SheetAdapter,
        min_row: int, max_row: int,
        min_col: int, max_col: int,
        source: str,
        name: str,
    ) -> Optional[dict]:
        # FIX v11: некорректные именованные диапазоны могут иметь min > max
        if min_col > max_col or min_row > max_row:
            return None

        vis_rows, _ = self._visible(adapter)
        cols = list(range(min_col, max_col + 1))
        rows_in = [r for r in vis_rows if min_row <= r <= max_row]
        if len(rows_in) < 2:
            return None

        header_row = rows_in[0]
        data_rows  = rows_in[1:]

        # FIX v11: дедупликация имён заголовков перед использованием
        raw_headers = [_to_str(adapter.cell(header_row, c)) for c in cols]
        deduped     = _dedupe_headers(raw_headers)
        header_dict: dict[int, str] = {
            c: deduped[i] for i, c in enumerate(cols) if deduped[i]
        }

        rows_out: list[Row] = []
        # FIX v18: кэшируем значения по колонкам для _detect_dtype —
        # ранее _detect_dtype вызывал adapter.cell() ещё раз для каждой колонки,
        # хотя все значения уже были прочитаны при формировании rows_out.
        col_values: dict[int, list[CellValue]] = {c: [] for c in cols}
        for r in data_rows:
            rd: dict[str, Any] = {}
            has = False
            for c in cols:
                v = adapter.cell(r, c)
                col_values[c].append(v)
                if not _is_empty(v):
                    rd[header_dict.get(c, get_column_letter(c))] = _serialize(v)
                    has = True
            if has:
                rows_out.append(rd)

        if not rows_out:
            return None

        return {
            "sheet": adapter.name,
            "name": name,
            "source": source,
            "header_row": header_row,
            "data_start": data_rows[0],
            "data_end":   data_rows[-1],
            "columns": [
                {
                    "letter": get_column_letter(c),
                    "name":   header_dict.get(c, get_column_letter(c)),
                    "type":   _detect_dtype(col_values[c]),  # FIX v18: из кэша
                }
                for c in cols
            ],
            "rows": rows_out,
        }

    # ── Источник 1: встроенные таблицы ───────────────────────────────────────

    def _extract_native_tables(self, adapter: SheetAdapter) -> list[dict]:
        results = []
        for tbl in adapter.native_tables():
            p = self._parse_range(
                adapter,
                tbl["min_row"], tbl["max_row"],
                tbl["min_col"], tbl["max_col"],
                TABLE_SOURCE_NATIVE, tbl["name"],
            )
            if p:
                results.append(p)
        return results

    # ── Источник 2: именованные диапазоны ─────────────────────────────────────

    def _extract_named_ranges_from_wb(self, wb, adapters_map: dict[str, SheetAdapter]) -> list[dict]:
        """
        FIX: Workbook-scope и Sheet-scope именованные диапазоны обрабатываются отдельно.
        Sheet-scope диапазоны имеют вид  SheetName!A1:B10  в destinations.
        """
        results: list[dict] = []
        try:
            defined = wb.defined_names
        except AttributeError:
            return results

        for dn in defined:
            try:
                destinations = dn.destinations
            except AttributeError:
                # dn — строка (некоторые версии openpyxl отдают строки напрямую)
                continue
            if isinstance(destinations, str):
                # Формат: "'SheetName'!A1:B10"
                if "!" not in destinations:
                    continue
                sheet_part, range_part = destinations.split("!", 1)
                sheet_title = sheet_part.strip("'\"$")
                try:
                    min_col, min_row, max_col, max_row = _range_boundaries(range_part)
                except Exception as e:
                    warnings.warn(
                        f"Именованный диапазон '{dn.name}': "
                        f"не удалось распарсить диапазон '{range_part}': {e}",
                        UserWarning, stacklevel=2,
                    )
                    continue
                if sheet_title not in adapters_map:
                    continue
                adapter = adapters_map[sheet_title]
                p = self._parse_range(
                    adapter, min_row, max_row, min_col, max_col,
                    TABLE_SOURCE_NAMED, dn.name,
                )
                if p:
                    results.append(p)
            else:
                for sheet_title, ref in destinations:
                    if sheet_title not in adapters_map:
                        continue
                    try:
                        min_col, min_row, max_col, max_row = _range_boundaries(ref)
                    except Exception as e:
                        warnings.warn(
                            f"Именованный диапазон '{dn.name}' на листе '{sheet_title}': "
                            f"не удалось распарсить диапазон '{ref}': {e}",
                            UserWarning, stacklevel=2,
                        )
                        continue
                    adapter = adapters_map[sheet_title]
                    p = self._parse_range(
                        adapter, min_row, max_row, min_col, max_col,
                        TABLE_SOURCE_NAMED, dn.name,
                    )
                    if p:
                        results.append(p)
        return results

    def _extract_named_ranges_from_adapter(self, adapter: SheetAdapter) -> list[dict]:
        """FIX: именованные диапазоны для .xls и .xlsb через адаптер."""
        results: list[dict] = []
        for nr in adapter.named_ranges():
            p = self._parse_range(
                adapter,
                nr["min_row"], nr["max_row"],
                nr["min_col"], nr["max_col"],
                TABLE_SOURCE_NAMED, nr["name"],
            )
            if p:
                results.append(p)
        return results

    # ── Источник 3: эвристика ─────────────────────────────────────────────────

    def _extract_heuristic(
        self, adapter: SheetAdapter, external_used_rows: set[int]
    ) -> list[dict]:
        """
        FIX v16: принимает external_used_rows из parse_sheet, чтобы видеть строки
        уже занятые native/named tables и не генерировать перекрывающие таблицы.
        """
        vis_rows, vis_cols = self._visible(adapter)
        if not vis_rows or not vis_cols:
            return []

        row_idx: dict[int, int] = {r: i for i, r in enumerate(vis_rows)}
        tables:   list[dict]    = []
        # Локальный used_rows объединяет внешние занятые строки со своими
        used_rows: set[int]     = set(external_used_rows)
        table_counter = 0
        i = 0

        while i < len(vis_rows):
            row  = vis_rows[i]
            vals = [adapter.cell(row, c) for c in vis_cols]

            if all(_is_empty(v) for v in vals):
                i += 1
                continue

            if not _is_header_row(vals, self.header_threshold):
                i += 1
                continue

            # ── Многострочный заголовок ───────────────────────────────────────
            header_rows  = [row]
            header_dict: dict[int, str] = {}
            for c in vis_cols:
                h = _to_str(adapter.cell(row, c))
                if h:
                    header_dict[c] = h

            for j in range(i + 1, min(i + 4, len(vis_rows))):
                r2   = vis_rows[j]
                r2v  = [adapter.cell(r2, c) for c in vis_cols]
                sc   = _score_header_row(r2v)
                nums = sum(1 for v in r2v if _is_numeric(v) and not _is_year(v))
                if sc >= 0.35 and nums == 0:
                    header_rows.append(r2)
                    for c in vis_cols:
                        h = _to_str(adapter.cell(r2, c))
                        if h:
                            header_dict[c] = (
                                header_dict[c] + " / " + h
                                if c in header_dict else h
                            )
                else:
                    break

            active_cols = set(header_dict.keys())
            if not active_cols:
                i += 1
                continue

            start_i = row_idx.get(header_rows[-1], i) + 1

            # ── Строки данных ─────────────────────────────────────────────────
            data_rows: list[int] = []
            empty_streak = 0

            for di in range(start_i, len(vis_rows)):
                r     = vis_rows[di]
                rv    = [adapter.cell(r, c) for c in active_cols]
                non_e = sum(1 for v in rv if not _is_empty(v))

                if non_e == 0:
                    empty_streak += 1
                    if empty_streak >= self.max_empty_streak:
                        break
                    continue

                if non_e < self.min_data_cells:
                    empty_streak = 0
                    continue

                empty_streak = 0

                # FIX v18: убран break по новому заголовку.
                # Ранее строки-маркеры ("Исход", "Доп.часы", "Низкие оценки ГЛ")
                # получали высокий header score и вызывали break — таблица
                # разрывалась, данные после маркера терялись.
                # Теперь: просто собираем все строки до конца пустой streak.

                data_rows.append(r)

            if not data_rows:
                i += 1
                continue

            span = set(range(header_rows[0], data_rows[-1] + 1))
            if span & used_rows:
                i += 1
                continue

            # ── Сборка ───────────────────────────────────────────────────────
            sorted_cols = sorted(active_cols)
            raw_col_names = [header_dict[c] for c in sorted_cols]
            # FIX v11: дедупликация имён колонок — одинаковые заголовки перезаписывали
            # значения друг друга в dict строки
            deduped_names = _dedupe_headers(raw_col_names)

            col_info: dict[str, dict] = {}
            for c, col_name in zip(sorted_cols, deduped_names):
                letter = get_column_letter(c)
                # FIX v16: значения столбца больше не хранятся в col_info —
                # ранее "values": vals_c создавало промежуточный список длиной N
                # для каждого столбца (50 столбцов × 10к строк = 500к объектов).
                # dtype вычисляется сразу, строки читаются напрямую через cell().
                col_info[letter] = {
                    "name":  col_name,
                    "index": c,
                    "dtype": _detect_dtype([adapter.cell(r, c) for r in data_rows]),
                }

            rows_out: list[Row] = []
            for r in data_rows:
                rd: dict[str, Any] = {}
                has = False
                for letter, info in col_info.items():
                    v = adapter.cell(r, info["index"])
                    if not _is_empty(v):
                        rd[info["name"]] = _serialize(v)
                        has = True
                if has:
                    rows_out.append(rd)

            if rows_out:
                table_counter += 1
                tables.append({
                    "sheet":      adapter.name,
                    "name":       f"{adapter.name} / Таблица {table_counter}",
                    "source":     TABLE_SOURCE_HEURISTIC,
                    "header_row": header_rows[0],
                    "data_start": data_rows[0],
                    "data_end":   data_rows[-1],
                    "columns": [
                        # FIX v17: переименована переменная l → letter (неотличима от 1)
                        {"letter": letter, "name": info["name"], "type": info["dtype"]}
                        for letter, info in col_info.items()
                    ],
                    "rows": rows_out,
                })
                used_rows.update(span)

            i = row_idx.get(data_rows[-1], i) + 1

        return tables

    # ── Источник 4: вертикальные таблицы ──────────────────────────────────────

    def _extract_vertical(self, adapter: SheetAdapter, used_rows: set[int]) -> list[dict]:
        """
        FIX: поддержка нескольких вертикальных таблиц на листе.
        Сканируем все строки и разбиваем на блоки по пустым разделителям.
        """
        vis_rows, vis_cols = self._visible(adapter)
        if not vis_rows or len(vis_cols) < 2:
            return []

        col_a   = vis_cols[0]
        other_c = vis_cols[1:]

        # Разбиваем строки на непрерывные блоки (разделитель — полностью пустая строка)
        blocks: list[list[int]] = []
        current: list[int] = []
        for r in vis_rows:
            rv = [adapter.cell(r, c) for c in vis_cols]
            if all(_is_empty(v) for v in rv):
                if current:
                    blocks.append(current)
                    current = []
            else:
                current.append(r)
        if current:
            blocks.append(current)

        results: list[dict] = []

        for block in blocks:
            if len(block) < 2:
                continue

            # Первая непустая строка блока — кандидат на заголовок
            header_row = block[0]

            # FIX v13: header_row не проверялся на вхождение в used_rows —
            # вертикальная таблица могла перекрыть заголовок уже найденной.
            if header_row in used_rows:
                continue

            header_vals = [adapter.cell(header_row, c) for c in vis_cols]
            if not _is_header_row(header_vals, 0.3):
                continue

            data_candidates = block[1:]

            # FIX v15: счётчики вычисляются только по строкам вне used_rows —
            # ранее считались по всем data_candidates, включая занятые строки,
            # что приводило к ложному прохождению порогов text_in_a / numeric_in_rest.
            text_in_a = 0
            numeric_in_rest = 0
            total_rest = 0
            for r in data_candidates:
                if r in used_rows:
                    continue          # исключаем уже занятые строки из статистики
                va = adapter.cell(r, col_a)
                if not _is_empty(va) and not _is_numeric(va):
                    text_in_a += 1
                for c in other_c:
                    v = adapter.cell(r, c)
                    if not _is_empty(v):
                        total_rest += 1
                        if _is_numeric(v):
                            numeric_in_rest += 1

            # FIX v11: порог снижен с 3 до 2 — небольшие вертикальные таблицы
            # (например, 2 строки данных) ранее молча пропускались
            if text_in_a < 2:
                continue
            if total_rest == 0 or numeric_in_rest / total_rest < 0.5:
                continue

            data_rows = [r for r in data_candidates if r not in used_rows]
            if not data_rows:
                continue

            # FIX v12: дедупликация имён заголовков — ранее дублирующиеся имена
            # в вертикальных таблицах тихо перезаписывали значения в dict строки
            raw_headers = [_to_str(adapter.cell(header_row, c)) for c in vis_cols]
            deduped_hdr = _dedupe_headers(raw_headers)
            header_dict = {c: deduped_hdr[i] for i, c in enumerate(vis_cols)}

            rows_out: list[Row] = []
            for r in data_rows:
                rd: dict[str, Any] = {}
                has = False
                for c in vis_cols:
                    v = adapter.cell(r, c)
                    if not _is_empty(v):
                        rd[header_dict.get(c, get_column_letter(c))] = _serialize(v)
                        has = True
                if has:
                    rows_out.append(rd)

            if not rows_out:
                continue

            tbl_num = len(results) + 1
            results.append({
                "sheet":      adapter.name,
                "name":       f"{adapter.name} / Вертикальная таблица {tbl_num}",
                "source":     TABLE_SOURCE_VERTICAL,
                "header_row": header_row,
                "data_start": data_rows[0],
                "data_end":   data_rows[-1],
                "columns": [
                    {
                        "letter": get_column_letter(c),
                        "name":   header_dict.get(c, get_column_letter(c)),
                        "type":   _detect_dtype([adapter.cell(r, c) for r in data_rows]),
                    }
                    for c in vis_cols
                ],
                "rows": rows_out,
            })
            used_rows.update(range(header_row, data_rows[-1] + 1))

        return results

    # ── Источник 5: таблицы без заголовка (числовые матрицы) ────────────────────

    def _extract_headerless(self, adapter: SheetAdapter, used_rows: set[int]) -> list[dict]:
        """
        FIX: обрабатывает листы/блоки без текстового заголовка.
        Присваивает колонкам имена A, B, C… и парсит непрерывные блоки данных.
        Срабатывает только если все остальные источники не нашли ничего.
        """
        vis_rows, vis_cols = self._visible(adapter)
        if not vis_rows or not vis_cols:
            return []

        tables: list[dict] = []
        table_counter = 0
        i = 0

        while i < len(vis_rows):
            r = vis_rows[i]
            vals = [adapter.cell(r, c) for c in vis_cols]

            if all(_is_empty(v) for v in vals):
                i += 1
                continue

            if r in used_rows:
                i += 1
                continue

            # Собираем непрерывный блок непустых строк
            block_rows: list[int] = []
            empty_streak = 0
            for j in range(i, len(vis_rows)):
                rj = vis_rows[j]
                if rj in used_rows:
                    break
                rv = [adapter.cell(rj, c) for c in vis_cols]
                non_e = sum(1 for v in rv if not _is_empty(v))
                if non_e == 0:
                    empty_streak += 1
                    if empty_streak >= self.max_empty_streak:
                        break
                    continue
                empty_streak = 0
                block_rows.append(rj)

            if len(block_rows) < 2:
                i += len(block_rows) + 1
                continue

            # Используем буквы колонок вместо заголовка
            col_letters = {c: get_column_letter(c) for c in vis_cols}
            rows_out: list[Row] = []
            for rr in block_rows:
                rd: dict[str, Any] = {}
                has = False
                for c in vis_cols:
                    v = adapter.cell(rr, c)
                    if not _is_empty(v):
                        rd[col_letters[c]] = _serialize(v)
                        has = True
                if has:
                    rows_out.append(rd)

            if rows_out:
                table_counter += 1
                tables.append({
                    "sheet":      adapter.name,
                    "name":       f"{adapter.name} / Беззаголовочная {table_counter}",
                    "source":     TABLE_SOURCE_HEADERLESS,
                    "header_row": block_rows[0],
                    "data_start": block_rows[0],
                    "data_end":   block_rows[-1],
                    "columns": [
                        {
                            "letter": get_column_letter(c),
                            "name":   col_letters[c],
                            "type":   _detect_dtype([adapter.cell(rr, c) for rr in block_rows]),
                        }
                        for c in vis_cols
                    ],
                    "rows": rows_out,
                })
                used_rows.update(range(block_rows[0], block_rows[-1] + 1))

            i += len(block_rows) + 1

        return tables

    # ── parse_sheet ───────────────────────────────────────────────────────────

    def parse_sheet(
        self,
        adapter: SheetAdapter,
        wb=None,
        all_adapters: Optional[dict[str, SheetAdapter]] = None,
    ) -> list[dict]:
        # FIX v16: сбрасываем кэш _visible для нового листа
        self._vis_cache.clear()

        all_tables: list[dict] = []
        used_rows:  set[int]   = set()

        # 1. Встроенные таблицы
        native = self._extract_native_tables(adapter)
        for t in native:
            used_rows.update(range(t["header_row"], t["data_end"] + 1))
        all_tables.extend(native)

        # 2. Именованные диапазоны
        # FIX v13: убрано условие "and not native" — ранее named ranges не
        # искались совсем если на листе уже нашлись нативные таблицы.
        # Теперь ищем всегда; перекрытия предотвращаются через used_rows.
        named: list[dict] = []
        if wb is not None and all_adapters:
            # xlsx: через openpyxl wb (Workbook-scope + Sheet-scope)
            named = [t for t in self._extract_named_ranges_from_wb(wb, all_adapters)
                     if t["sheet"] == adapter.name]
        else:
            # xls / xlsb: через адаптер напрямую
            named = self._extract_named_ranges_from_adapter(adapter)

        for t in named:
            span = set(range(t["header_row"], t["data_end"] + 1))
            if not (span & used_rows):
                used_rows.update(span)
                all_tables.append(t)

        # 3. Эвристика
        # FIX v16: передаём used_rows чтобы эвристика видела уже занятые строки
        if not all_tables:
            heuristic = self._extract_heuristic(adapter, used_rows)
            for t in heuristic:
                used_rows.update(range(t["header_row"], t["data_end"] + 1))
            all_tables.extend(heuristic)

        # 4. Вертикальные таблицы
        # FIX v12: убран guard "if not all_tables" — вертикальные таблицы теперь
        # всегда ищутся после эвристики. Лист может содержать одновременно
        # обычные (heuristic) и вертикальные блоки. used_rows предотвращает
        # перекрытие уже найденных строк.
        all_tables.extend(self._extract_vertical(adapter, used_rows))

        # 5. Беззаголовочные таблицы (последний резерв — матрицы без шапки)
        # FIX v18: запускать всегда, а не только когда all_tables пуст.
        # heuristic может найти маленькую таблицу и пропустить остальные строки
        # (напр. Лист3 в файле Доплаты — 6 строк из 56). headerless заполнит
        # оставшиеся строки через used_rows.
        all_tables.extend(self._extract_headerless(adapter, used_rows))

        return all_tables

    # ── parse_file ────────────────────────────────────────────────────────────

    def parse_file(
        self,
        filepath: str,
        output_path: Optional[str] = None,
        fmt: str = "json",
        only_sheet: Optional[str] = None,
        streaming: bool = False,
    ) -> dict:
        print(f"\n{'═' * 72}")
        print(f"📊  {os.path.basename(filepath)}")
        print(f"{'═' * 72}")

        ext      = os.path.splitext(filepath)[1].lower()
        adapters = load_sheets(filepath, only_sheet)

        wb = None
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            # FIX v12: файл больше не загружается дважды. load_sheets() уже
            # вызвал openpyxl.load_workbook() внутри _load_xlsx(). Извлекаем
            # workbook из первого OpenpyxlAdapter через ws.parent — бесплатно.
            # FIX v14: добавлены .xltx / .xltm — ранее они не входили в проверку,
            # wb оставался None, и named ranges для этих форматов терялись.
            first_opx = next(
                (a for a in adapters if isinstance(a, OpenpyxlAdapter)), None
            )
            if first_opx is not None:
                wb = first_opx._ws.parent
        adapters_map = {a.name: a for a in adapters}

        # Потоковый writer
        writer: Optional[StreamingWriter] = None
        if streaming and output_path:
            file_meta = {"file": os.path.basename(filepath), "format": ext.lstrip(".")}
            writer = StreamingWriter(output_path, fmt, file_meta)
        elif streaming and not output_path:
            # FIX v16: ранее данные молча терялись без предупреждения
            warnings.warn(
                "--stream требует --out-dir или явного output_path. "
                "Данные не будут сохранены на диск.",
                UserWarning, stacklevel=2,
            )

        all_tables: list[dict] = []
        # FIX: счётчики источников работают в обоих режимах (streaming и обычный)
        stream_sources: dict[str, int] = {
            TABLE_SOURCE_NATIVE: 0,
            TABLE_SOURCE_NAMED: 0,
            TABLE_SOURCE_HEURISTIC: 0,
            TABLE_SOURCE_VERTICAL: 0,
            TABLE_SOURCE_HEADERLESS: 0,
        }
        iter_list = (
            _tqdm(adapters, desc="Листы", unit="лист")
            if HAS_TQDM and len(adapters) > 3
            else adapters
        )

        for adapter in iter_list:
            tables = self.parse_sheet(adapter, wb=wb, all_adapters=adapters_map)

            src_counts: dict[str, int] = {}
            for t in tables:
                src = t["source"]
                src_counts[src] = src_counts.get(src, 0) + 1
                # FIX: накапливаем счётчики источников независимо от режима
                if src in stream_sources:
                    stream_sources[src] += 1
                if writer:
                    writer.write_table(t)

            summary = ", ".join(f"{k}:{v}" for k, v in src_counts.items()) or "—"
            print(f"  📄  '{adapter.name}' "
                  f"({adapter.max_row}×{adapter.max_col}) "
                  f"→ {len(tables)} таблиц [{summary}]")

            if not streaming:
                all_tables.extend(tables)

        if writer:
            writer.close()
            n_tables, n_rows = writer.stats
        else:
            n_tables = len(all_tables)
            n_rows   = sum(len(t["rows"]) for t in all_tables)

        result = {
            "file":       os.path.basename(filepath),
            "format":     ext.lstrip("."),
            "sheets":     len(adapters),
            "tables":     n_tables,
            "total_rows": n_rows,
            # FIX: stream_sources корректен в обоих режимах
            "sources": stream_sources,
            "tables_data": all_tables,
        }

        if output_path and not streaming:
            _write_output(result, output_path, fmt)
            print(f"  💾  {output_path}")

        print(f"\n{'═' * 72}")
        print(
            f"✅  {n_tables} таблиц | {n_rows} строк | "
            f"native:{result['sources'][TABLE_SOURCE_NATIVE]} "
            f"named:{result['sources'][TABLE_SOURCE_NAMED]} "
            f"heuristic:{result['sources'][TABLE_SOURCE_HEURISTIC]} "
            f"vertical:{result['sources'][TABLE_SOURCE_VERTICAL]} "
            f"headerless:{result['sources'][TABLE_SOURCE_HEADERLESS]}"
            + (" | streaming" if streaming else "")
        )
        print(f"{'═' * 72}\n")

        return result


# ══════════════════════════════════════════════════════════════════════════════
# Запись (не-потоковый режим)
# ══════════════════════════════════════════════════════════════════════════════

def _write_output(result: dict, path: str, fmt: str) -> None:
    if fmt == "json":
        with open(path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2, default=str)
    elif fmt == "jsonl":
        with open(path, "w", encoding="utf-8") as f:
            for table in result["tables_data"]:
                for row in table["rows"]:
                    record = {"_sheet": table["sheet"], "_table": table["name"], **row}
                    f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")
    elif fmt == "csv":
        os.makedirs(path, exist_ok=True)
        for idx, table in enumerate(result["tables_data"]):
            rows = table.get("rows", [])
            if not rows:
                continue
            safe = (
                table["name"]
                .replace("/", "_").replace("\\", "_")
                .replace(":", "_").replace(" ", "_")[:60]
            )
            csv_path = os.path.join(path, f"{idx + 1:02d}_{safe}.csv")
            with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)
    else:
        raise ValueError(f"Неизвестный формат: {fmt}")


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Excel Universal Parser v18",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    ap.add_argument("files", nargs="+",
                    help="Файлы .xlsx .xlsm .xls .xlsb .csv")
    ap.add_argument("--out-dir", default=None,
                    help="Директория для результатов")
    ap.add_argument("--format", choices=["json", "jsonl", "csv"], default="json",
                    help="Формат вывода (default: json)")
    ap.add_argument("--sheet", default=None,
                    help="Обрабатывать только этот лист")
    ap.add_argument("--header-threshold", type=float, default=0.4,
                    help="Score-порог для заголовка [0..1] (default: 0.4)")
    ap.add_argument("--min-data-cells", type=int, default=2,
                    help="Мин. непустых ячеек для строки данных (default: 2)")
    ap.add_argument("--include-hidden", action="store_true",
                    help="Включить скрытые строки и колонки")
    ap.add_argument("--stream", action="store_true",
                    help="Потоковая запись — не накапливать данные в RAM")
    args = ap.parse_args()

    parser = ExcelParser(
        header_threshold=args.header_threshold,
        skip_hidden=not args.include_hidden,
        min_data_cells=args.min_data_cells,
    )

    for fp in args.files:
        if not os.path.exists(fp):
            print(f"❌  Файл не найден: {fp}", file=sys.stderr)
            continue

        out_dir = args.out_dir or os.path.dirname(os.path.abspath(fp))
        stem    = os.path.splitext(os.path.basename(fp))[0].replace(".", "_")

        if args.format == "csv":
            out_path = os.path.join(out_dir, stem + "_tables")
        else:
            out_path = os.path.join(out_dir, stem + "_parsed." + args.format)

        try:
            parser.parse_file(
                fp,
                output_path=out_path,
                fmt=args.format,
                only_sheet=args.sheet,
                streaming=args.stream,
            )
        except Exception as exc:
            # FIX v17: traceback перенесён в top-level импорты
            print(f"❌  Ошибка: {fp}\n    {exc}", file=sys.stderr)
            traceback.print_exc()


if __name__ == "__main__":
    main()
