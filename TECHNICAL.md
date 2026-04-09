# Техническая документация: Excel Smart Parser v18.1

> Полное описание архитектуры, алгоритмов и внутренних механизмов парсера.
> 1895 строк кода, 68 тестов, 5 форматов файлов.

---

## Оглавление

1. [Общая архитектура](#1-общая-архитектура)
2. [Адаптеры форматов](#2-адаптеры-форматов)
3. [Загрузчики (load_sheets)](#3-загрузчики-load_sheets)
4. [Score-based детекция заголовков](#4-score-based-детекция-заголовков)
5. [Пять источников таблиц](#5-пять-источников-таблиц)
6. [Координация: used_rows](#6-координация-usedrows)
7. [Утилиты значений](#7-утилиты-значений)
8. [Потоковый writer](#8-потоковый-writer)
9. [Дедупликация заголовков](#9-дедупликация-заголовков)
10. [Обработка скрытых строк/колонок](#10-обработка-скрытых-строкколонок)
11. [Кэширование и производительность](#11-кэширование-и-производительность)
12. [Диагностика формул](#12-диагностика-формул)
13. [CSV: определение кодировки и разделителя](#13-csv-определение-кодировки-и-разделителя)
14. [Обработка merged cells](#14-обработка-merged-cells)
15. [Определение типов данных (dtype)](#15-определение-типов-данных-dtype)
16. [Поток данных parse_file](#16-поток-данных-parse_file)
17. [История багов и фиксов](#17-история-багов-и-фиксов)

---

## 1. Общая архитектура

Парсер построен по паттерну **Adapter + Multi-source Extraction**.

```
                    ┌─────────────────────┐
                    │     parse_file()    │
                    │   (точка входа)     │
                    └─────────┬───────────┘
                              │
                    ┌─────────▼───────────┐
                    │    load_sheets()    │
                    │  (загрузчик листов) │
                    └─────────┬───────────┘
                              │
          ┌──────────┬────────┼────────┬──────────┐
          ▼          ▼        ▼        ▼          ▼
    ┌──────────┐ ┌────────┐ ┌────────┐ ┌────────┐ ┌────────┐
    │Openpyxl  │ │ Xlrd   │ │Pyxlsb  │ │  CSV   │ │ (будущие│
    │ Adapter  │ │Adapter │ │Adapter │ │Adapter │ │ форматы)│
    └────┬─────┘ └───┬────┘ └───┬────┘ └───┬────┘ └─────────┘
         │           │          │           │
         └───────────┴──────────┴───────────┘
                              │
                    ┌─────────▼───────────┐
                    │   parse_sheet()     │
                    │ (координация 5-ти   │
                    │   источников)       │
                    └─────────┬───────────┘
                              │
          ┌──────────┬────────┼────────┬──────────┐
          ▼          ▼        ▼        ▼          ▼
    ┌──────────┐ ┌────────┐ ┌────────┐ ┌────────┐ ┌──────────┐
    │ Native   │ │ Named  │ │Heurist │ │Vertic. │ │Headerless│
    │ Tables   │ │ Ranges │ │ ic     │ │ al     │ │          │
    └──────────┘ └────────┘ └────────┘ └────────┘ └──────────┘
```

### Базовый класс SheetAdapter (abc.ABC)

```python
class SheetAdapter(abc.ABC):
    name: str          # имя листа
    max_row: int       # макс. номер строки
    max_col: int       # макс. номер колонки

    @abc.abstractmethod
    def cell(self, row: int, col: int) -> CellValue:
        """Возвращает значение ячейки (1-based). None для пустых."""

    # Методы с дефолтной реализацией (переопределяются при необходимости):
    def iter_rows_lazy(self, cols: list[int]) -> Generator[...]: ...
    def hidden_rows(self) -> set[int]: ...
    def hidden_cols(self) -> set[int]: ...
    def native_tables(self) -> list[dict]: ...
    def named_ranges(self) -> list[dict]: ...
```

**Зачем ABC**: если кто-то создаст подкласс без `cell()`, ошибка (`TypeError`)
возникнет немедленно при `__init__`, а не в рантайме при первом вызове.

---

## 2. Адаптеры форматов

### 2.1 OpenpyxlAdapter (.xlsx, .xlsm, .xltx, .xltm)

**Инициализация:**
```python
class OpenpyxlAdapter(SheetAdapter):
    def __init__(self, ws, name: str):
        self._ws = ws
        self.name = name
        self.max_row = ws.max_row or 0
        self.max_col = ws.max_column or 0
        # Кэш merged cells — строится один раз при инициализации
        self._merged: dict[tuple[int, int], CellValue] = {}
        for rng in ws.merged_cells.ranges:
            master = ws.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._merged[(r, c)] = master
```

**Почему кэш merged cells при init**: без кэша каждый вызов `cell()` проверял бы
`merged_cells.ranges` — O(N) на каждую ячейку, где N = число merged диапазонов.
С кэшем — O(1) dict lookup.

**cell():**
```python
def cell(self, row: int, col: int) -> CellValue:
    if (row, col) in self._merged:
        return self._merged[(row, col)]
    return self._ws.cell(row, col).value
```

**hidden_cols()** использует `warnings.warn` вместо `pass`:
```python
def hidden_cols(self) -> set[int]:
    result: set[int] = set()
    for letter, d in self._ws.column_dimensions.items():
        if d.hidden:
            try:
                result.add(column_index_from_string(letter))
            except Exception as e:
                warnings.warn(f"Лист '{self.name}': скрытая колонка '{letter}': {e}")
    return result
```

### 2.2 XlrdAdapter (.xls)

**Ключевые отличия от openpyxl:**
- xlrd работает с **0-based** индексами, адаптер конвертирует в 1-based
- Даты в xlrd хранятся как float (Excel serial date) — конвертируется через `xldate_as_datetime`
- `named_ranges` парсит `name_obj_list` с `xlrd.Ref3D.coords` (6 элементов)

**Именованные диапазоны:**
```python
def named_ranges(self) -> list[dict]:
    for name_obj in self._book.name_obj_list:
        for area in name_obj.result.coords:
            shtxlo, _shtxhi, row0, row1, col0, col1 = area[:6]
            # row0, row1 — 0-based, [lo, hi) полуоткрытый
            # colxhi — 0-based exclusive → 1-based inclusive = row1
            result.append({
                "min_row": row0 + 1,  # 0-based lo → 1-based
                "max_row": row1,       # 0-based exclusive hi = 1-based inclusive
                "min_col": col0 + 1,
                "max_col": col1,
            })
```

### 2.3 PyxlsbAdapter (.xlsb)

**Особенности:**
- При `__init__` весь лист загружается в `_row_cache` **одним проходом**
- `max_row` и `max_col` определяются **в том же проходе** (раньше было 2 прохода)
- `iter_rows_lazy()` итерирует по `_row_cache` без повторного чтения файла

```python
def _ensure_cache(self) -> None:
    """Один проход: размеры + кэш одновременно."""
    tmp: dict[int, dict] = {}
    max_r = max_c = 0
    with pyxlsb.open_workbook(self._filepath) as wb:
        with wb.get_sheet(self.name) as ws:
            for i, row in enumerate(ws.rows()):
                r = i + 1
                if r > max_r: max_r = r
                for cell in row:
                    c_idx = cell.c
                    if c_idx + 1 > max_c: max_c = c_idx + 1
                    tmp.setdefault(r, {})[c_idx] = cell.v
    # ... заполнение _row_cache из tmp
```

### 2.4 CsvAdapter (.csv)

- Весь файл загружается в `_row_cache` **одним проходом** при `__init__`
- `cell()` и `iter_rows_lazy()` работают по кэшу — повторного чтения нет
- Потребление памяти: **O(n)** по числу строк файла

---

## 3. Загрузчики (load_sheets)

### 3.1 Главная функция

```python
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
```

### 3.2 _load_xlsx — устранение двойного открытия

**Было (v17):** `load_workbook(data_only=True)` → `_check_formulas_lazy` открывает
файл заново с `data_only=False`. **Два load_workbook на один файл.**

**Стало (v18):**
```python
def _load_xlsx(filepath: str, only_sheet: Optional[str]) -> list[OpenpyxlAdapter]:
    # 1. Лёгкий read_only-проход для проверки формул
    wb_scan = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
    for sheet_name in wb_scan.sheetnames:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value.startswith("="):
                    warnings.warn("Обнаружены формулы...")
                    break
            else: continue
            break
        sheets_to_process.append(sheet_name)
    wb_scan.close()

    # 2. Один полный проход для данных
    wb = openpyxl.load_workbook(filepath, data_only=True)
    for name in sheets_to_process:
        ws = wb[name]
        sheets.append(OpenpyxlAdapter(ws, name))
    return sheets
```

**read_only=True** потребляет минимум RAM — только структура листа, не данные.

---

## 4. Score-based детекция заголовков

### 4.1 Алгоритм _score_header_row

Функция оценивает строку по шкале **0.0–1.0** — насколько она похожа на заголовок.

```python
def _score_header_row(vals: list[CellValue]) -> float:
    non_empty = [v for v in vals if not _is_empty(v)]
    if not non_empty:
        return 0.0
```

### 4.2 Detection паттерна "числа-дни" (FIX v18)

Заголовки таблиц с днями месяца содержат числа 1–31:

| ФИО | op | 1 | 2 | 3 | ... | 31 | ИТОГО |
|-----|----|---|---|---|-----|----|---|

Без специальной обработки числа 1–31 штрафовались бы как `-0.2 × 31 = -6.2`,
давая score ≈ 0.0. **Фикс v18** распознаёт этот паттерн:

```python
    # Определяем паттерн "числа-дни"
    numeric_vals = [v for v in non_empty if _is_numeric(v) and not _is_year(v)]
    is_day_header = False
    if len(numeric_vals) >= 5:
        day_numbers = []
        for v in numeric_vals:
            if isinstance(v, float):
                if not math.isfinite(v): continue  # FIX v18: nan/inf защита
                n = int(v)
            else:
                n = v
            if 1 <= n <= 31:
                day_numbers.append(n)
        if len(day_numbers) >= 5:
            day_numbers_sorted = sorted(day_numbers)
            # Проверяем что большинство идут подряд (разница ≤2)
            consecutive = sum(
                1 for i in range(len(day_numbers_sorted) - 1)
                if 1 <= (day_numbers_sorted[i + 1] - day_numbers_sorted[i]) <= 2
            )
            if consecutive >= len(day_numbers_sorted) * 0.6:
                is_day_header = True
```

**Порог 60%**: допускает пропуски (например, нет выходных дней).

### 4.3 Таблица вкладов

| Тип значения | В обычном режиме | Если is_day_header |
|-------------|-----------------|-------------------|
| Числовой год (int 1900–2200) | 0.0 | 0.0 |
| Строковый год ("2024") | +0.3 | +0.3 |
| Дата (datetime) | +0.5 | +0.5 |
| Число | **−0.2** | **+0.5** |
| Текст 1 символ | +0.7 | +0.7 |
| Текст 2–60 символов | +1.0 | +1.0 |
| Текст >60 символов | −0.5 | −0.5 |

### 4.4 Нормализация

```python
    ratio = text_like / len(non_empty)
    if len(non_empty) < 2:
        ratio *= 0.6   # Штраф за малоколоночные строки
    return max(0.0, min(1.0, ratio))
```

### 4.5 Практические значения

| Строка | Score | Результат |
|--------|-------|-----------|
| `["ФИО", "op", 1, 2, 3, ..., 31, "ИТОГО"]` | ~0.7 | ✅ Заголовок |
| `["Имя", "Возраст", "Город"]` | 1.0 | ✅ Заголовок |
| `["Исход"]` | 0.6 | ⚠️ Маркер |
| `[100, 200, 300, 400]` | 0.0 | ❌ Данные |
| `[None, None, None]` | 0.0 | ❌ Пустая |

---

## 5. Пять источников таблиц

### 5.1 Порядок и логика

```
1. native_table  — точные (созданы пользователем через Ctrl+T)
2. named_range   — точные (именованные диапазоны)
3. heuristic     — эвристика (score-based поиск заголовков)
4. vertical      — вертикальные таблицы (заголовки в колонке A)
5. headerless    — матрицы без заголовка (колонки = A, B, C...)
```

Каждый источник записывает найденные строки в `used_rows`. Следующий источник
пропускает эти строки.

### 5.2 Источник 1: native_table

Excel-таблицы, созданные через **Вставка → Таблица** (Ctrl+T).

```python
def _extract_native_tables(self, adapter) -> list[dict]:
    for tbl in adapter.native_tables():
        p = self._parse_range(adapter,
            tbl["min_row"], tbl["max_row"],
            tbl["min_col"], tbl["max_col"],
            TABLE_SOURCE_NATIVE, tbl["name"])
        if p: results.append(p)
```

### 5.3 Источник 2: named_range

Именованные диапазони из `defined_names` (Workbook-scope) или
`sheet.name_obj_list` (.xls).

**Разделение scope:**
```python
if wb is not None and all_adapters:
    # xlsx: через openpyxl Workbook (Workbook-scope + Sheet-scope)
    named = self._extract_named_ranges_from_wb(wb, all_adapters)
else:
    # xls / xlsb: через адаптер
    named = self._extract_named_ranges_from_adapter(adapter)
```

### 5.4 Источник 3: heuristic

Главный источник для файлов без явных таблиц.

**Алгоритм:**
```
for each visible row:
    score = _score_header_row(vals)
    if score >= threshold:
        # Нашли потенциальный заголовок!

        # 1. Проверяем многострочный заголовок (до 4 строк)
        header_rows = [current_row]
        for next_rows in range(1, 4):
            if score >= 0.35 and numeric_count == 0:
                header_rows.append(next_row)
                # Объединяем заголовки через " / "

        # 2. Определяем active_cols (колонки с непустыми заголовками)
        active_cols = {c for c in header_dict if header_dict[c]}

        # 3. Собираем строки данных
        data_rows = []
        empty_streak = 0
        for next_row in remaining_rows:
            non_empty = count_non_empty(active_cols)
            if non_empty == 0:
                empty_streak += 1
                if empty_streak >= max_empty_streak (50): break
                continue
            if non_empty < min_data_cells (2):
                empty_streak = 0  # Сброс streak
                continue
            data_rows.append(row)

        # 4. Проверяем перекрытие с used_rows
        span = set(range(header_row, last_data_row + 1))
        if span & used_rows:
            continue  # Пропускаем — уже занято

        # 5. Собираем таблицу
        table = build_table(...)
        used_rows.update(span)
```

**Ключевые параметры:**
- `max_empty_streak = 50` — не разрывать таблицу при коротких разреженных секциях
- `min_data_cells = 2` — минимум непустых ячеек в строке данных
- `header_threshold = 0.4` — порог score для заголовка

### 5.5 Источник 4: vertical

Для таблиц где заголовки находятся в **первой колонке** (A), а данные — в остальных.

```
Блок строк (между пустыми разделителями):
┌──────────────┬───────┬───────┐
│ Заголовок A  │ Заг. B│ Заг. C│  ← header_row
├──────────────┼───────┼───────┤
│ Текст в A    │  100  │  200  │  ← data_row
│ Текст в A    │  300  │  400  │  ← data_row
│ Текст в A    │  500  │  600  │  ← data_row
└──────────────┴───────┴───────┘
```

**Критерии:**
1. Первая строка блока — заголовок (`_is_header_row >= 0.3`)
2. `text_in_a >= 2` — хотя бы 2 строки с текстом в колонке A
3. `numeric_in_rest / total_rest >= 0.5` — >50% остальных ячеек числовые

### 5.6 Источник 5: headerless (FIX v18)

Для матриц без текстового заголовка. Присваивает колонкам имена A, B, C...

**FIX v18**: запускается **всегда**, а не только когда `all_tables` пуст.
Это позволяет подхватить строки, пропущенные heuristic.

```python
# Было (v17):
if not all_tables:
    all_tables.extend(self._extract_headerless(adapter, used_rows))

# Стало (v18):
all_tables.extend(self._extract_headerless(adapter, used_rows))
```

---

## 6. Координация: used_rows

`used_rows` — `set[int]` который отслеживает строки уже найденных таблиц.

### 6.1 Как используется

```python
# В parse_sheet():
used_rows: set[int] = set()

# 1. Native таблицы
for t in native:
    used_rows.update(range(t["header_row"], t["data_end"] + 1))

# 2. Named ranges — только если не пересекаются
for t in named:
    span = set(range(t["header_row"], t["data_end"] + 1))
    if not (span & used_rows):  # Пересечение == пустое множество
        used_rows.update(span)
        all_tables.append(t)

# 3. Heuristic — видит used_rows
heuristic = self._extract_heuristic(adapter, used_rows)

# 4. Vertical — тоже
all_tables.extend(self._extract_vertical(adapter, used_rows))

# 5. Headerless — тоже
all_tables.extend(self._extract_headerless(adapter, used_rows))
```

### 6.2 Почему это важно

Без `used_rows`:
- Heuristic нашёл бы таблицу поверх named range
- Vertical нашёл бы таблицу поверх heuristic
- Один и тот же ряд данных попал бы в несколько таблиц

---

## 7. Утилиты значений

### 7.1 _is_numeric

```python
def _is_numeric(v: CellValue) -> bool:
    if isinstance(v, bool):
        return False                    # bool — подкласс int, но не число
    if isinstance(v, (int, float)):
        return math.isfinite(v)         # отсекаем nan/inf
    s = _to_str(v).replace(",", ".").replace(" ", "").replace("%", "")
    if not s:
        return False
    try:
        return math.isfinite(float(s))  # float("nan") не бросает ValueError
    except ValueError:
        return False
```

**Почему `math.isfinite`**: `float("nan")` и `float("inf")` не бросают
`ValueError` при `float()` — они успешно парсятся. Без `isfinite` строки
"nan", "Infinity" в CSV ошибочно классифицировались бы как числа.

### 7.2 _is_year

```python
def _is_year(v: CellValue) -> bool:
    if isinstance(v, bool): return False
    if isinstance(v, float):
        if not math.isfinite(v): return False   # FIX v17: nan/inf
        if v != int(v): return False            # FIX v16: 2021.5 — не год
        return 1900 <= int(v) <= 2200
    if isinstance(v, int):
        return 1900 <= v <= 2200
    s = _to_str(v)
    try:
        return 1900 <= int(s) <= 2200
    except ValueError:
        return False
```

### 7.3 _is_date

```python
def _is_date(v: CellValue) -> bool:
    return isinstance(v, (datetime.datetime, datetime.date))
```

### 7.4 _serialize

```python
def _serialize(v: CellValue) -> Any:
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()   # datetime → "2024-03-15T00:00:00"
    return v                    # остальное как есть
```

---

## 8. Потоковый writer

### 8.1 Зачем

Без `--stream` все таблицы накапливаются в `all_tables` → `json.dump()` пишет
всё сразу. Для файлов с миллионами строк — OutOfMemory.

`StreamingWriter` пишет каждую строку **сразу на диск**.

### 8.2 JSON режим

```python
def write_table(self, table: dict) -> None:
    if self.fmt == "json":
        # FIX v14: ранее строил t_copy со всеми rows и сериализовывал целиком
        # Теперь: мета без rows, затем rows по одному

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
```

**Почему `meta_str[:-1]`**: `json.dumps({"a": 1})` → `{"a": 1}`.
`[:-1]` убирает `}` → `{"a": 1` → добавляем `,"rows": [...]}`.

FIX v11: раньше использовался `rstrip("}")` — срезал бы символ `}` если
**значение** в метаданных заканчивалось на `}` (напр. путь `C:\{test}`).

### 8.3 JSONL режим

```python
for row in table.get("rows", []):
    record = {"_sheet": table["sheet"], "_table": table["name"], **row}
    self._jsonl_fh.write(json.dumps(record, ...) + "\n")
```

### 8.4 CSV режим

```python
safe = table["name"].replace("/", "_")...[:60]
path = os.path.join(self._csv_dir, f"{table_num:02d}_{safe}.csv")
with open(path, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
    w.writeheader()
    w.writerows(rows)
```

**utf-8-sig**: с BOM для корректного открытия в Excel.

---

## 9. Дедупликация заголовков

### 9.1 Проблема

Заголовки `["Сумма", "Сумма", "Сумма"]` → dict-строка:
```python
{"Сумма": 100, "Сумма": 200, "Сумма": 300}  →  {"Сумма": 300}
```
Вторая и третья колонки **тихо перезаписывают** первую.

### 9.2 Решение

```python
def _dedupe_headers(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for i, name in enumerate(names):
        key = name if name else f"_col_{i + 1}"  # Пустые → _col_N
        if key not in seen:
            seen[key] = 1
            result.append(key)
        else:
            seen[key] += 1
            result.append(f"{key}_{seen[key]}")
    return result
```

**Результат:** `["Сумма", "Сумма_2", "Сумма_3"]`

### 9.3 Многострочные заголовки

При объединении заголовков строк 1–4:
```
Строка 1: Способ отражения | Способ отражения | ...
Строка 2: Подразделение    | Подразделение    | ...
Строка 3: №                | Сотрудник        | ...
Строка 4: 20.01 Зарплата   | 20.01 Зарплата   | ...
```
→ `"Способ отражения / Подразделение / № / 20.01 Зарплата"`

Затем `_dedupe_headers` uniquifies дубликаты.

---

## 10. Обработка скрытых строк/колонок

### 10.1 _visible() с кэшем

```python
def _visible(self, adapter: SheetAdapter) -> tuple[list[int], list[int]]:
    key = id(adapter)
    if key not in self._vis_cache:
        hr = adapter.hidden_rows() if self.skip_hidden else set()
        hc = adapter.hidden_cols() if self.skip_hidden else set()
        rows = [r for r in range(1, adapter.max_row + 1) if r not in hr]
        cols = [c for c in range(1, adapter.max_col + 1) if c not in hc]
        self._vis_cache[key] = (rows, cols)
    return self._vis_cache[key]
```

**FIX v16**: без кэша `_visible` вызывался 3 раза за лист
(heuristic + vertical + headerless) — каждый раз итерация по `row_dimensions`.

### 10.2 skip_hidden = False (FIX v18)

```python
def __init__(self, ..., skip_hidden: bool = False, ...):
```

**Почему False по умолчанию**: файлы могут содержать свёрнутые строки
с данными. `skip_hidden=True` молча их пропускал — пользователь терял данные
без предупреждения.

---

## 11. Кэширование и производительность

### 11.1 Таблица кэшей

| Кэш | Что хранит | Где используется | Сброс |
|-----|-----------|-----------------|-------|
| `_vis_cache` | (видимые строки, видимые колонки) | `_visible()` | Каждый `parse_sheet` |
| `_merged` | {(row, col): value} для merged cells | `OpenpyxlAdapter.cell()` | При init адаптера |
| `_row_cache` (xlsb) | {row: [values]} для всех строк | `PyxlsbAdapter` | При init адаптера |
| `_row_cache` (csv) | {row: [values]} для всех строк | `CsvAdapter` | При init адаптера |
| `col_values` | {col_index: [values]} для колонки | `_parse_range` dtype | Локальный, на таблицу |

### 11.2 _parse_range: кэш col_values (FIX v18)

**Было (v17):**
```python
# Для каждой колонки отдельно:
"type": _detect_dtype([adapter.cell(r, c) for r in data_rows])
```
N колонок × M строк = **N вызовов adapter.cell()** для dtype, **плюс** ещё N×M
для формирования rows_out.

**Стало (v18):**
```python
col_values: dict[int, list[CellValue]] = {c: [] for c in cols}
for r in data_rows:
    for c in cols:
        v = adapter.cell(r, c)
        col_values[c].append(v)      # Кэшируем
        if not _is_empty(v):
            rd[name] = _serialize(v)
            has = True

# ...
"columns": [{
    "type": _detect_dtype(col_values[c])  # Из кэша, 0 вызовов cell()
}]
```

### 11.3 Hot path: _is_numeric

`_is_numeric` вызывается в `_score_header_row` для **каждой ячейки каждой строки**.
В файле 1000×50 = 50 000 вызовов. Поэтому:

- `math` импортирован на top-level (не внутри функции)
- Минимум аллокаций
- Ранний return для `bool`, `int`, `float`

---

## 12. Диагностика формул

### 12.1 Проблема

openpyxl с `data_only=True` возвращает **последнее сохранённое значение** формулы.
Если файл не был сохранён после редактирования формул — возвращается `None`.

### 12.2 Решение

```python
def _load_xlsx(filepath, only_sheet):
    # 1. Открываем read_only с data_only=False (сырые формулы)
    wb_scan = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
    for sheet_name in wb_scan.sheetnames:
        ws = wb_scan[sheet_name]
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value.startswith("="):
                    warnings.warn(
                        f"Лист '{sheet_name}': обнаружены формулы. "
                        "data_only=True вернёт None для несохранённых значений."
                    )
                    break
            else: continue
            break
    wb_scan.close()

    # 2. Открываем с data_only=True для данных
    wb = openpyxl.load_workbook(filepath, data_only=True)
```

---

## 13. CSV: определение кодировки и разделителя

### 13.1 _detect_encoding

```python
def _detect_encoding(filepath: str) -> str:
    if HAS_CHARDET:
        with open(filepath, "rb") as f:
            raw = f.read(32768)   # Первые 32KB
        detected = _chardet.detect(raw)
        enc = (detected.get("encoding") or "utf-8").strip()
        try:
            codecs.lookup(enc)    # Проверяем что Python знает кодировку
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
```

### 13.2 Определение разделителя

```python
def _load_csv(filepath: str) -> list[CsvAdapter]:
    with open(filepath, newline="", encoding=encoding) as f:
        sample = f.read(16384)

    # 1. Пробуем csv.Sniffer
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|^~ ")
        delimiter = dialect.delimiter
    except csv.Error:
        # 2. Fallback: считаем частоту символов
        _FALLBACK_DELIMITERS = frozenset(",;\t|^~ ")
        candidates = {d: sample.count(d) for d in _FALLBACK_DELIMITERS}
        _best = max(candidates, key=candidates.get)
        delimiter = _best if candidates[_best] > 0 else ","
```

**FIX v17**: `max(..., default=",")` не работал — `default` у `max()`
используется только когда итерируемый объект **пуст**. `candidates` всегда
содержит 8 элементов, поэтому `max()` всегда возвращал какой-то символ,
даже если его счётчик = 0.

---

## 14. Обработка merged cells

### 14.1 Как openpyxl хранит merged cells

```python
ws.merged_cells.ranges  # список Range объектов
# Range: min_row, max_row, min_col, max_col
```

### 14.2 Кэш при инициализации

```python
self._merged: dict[tuple[int, int], CellValue] = {}
for rng in ws.merged_cells.ranges:
    master = ws.cell(rng.min_row, rng.min_col).value  # Значение из master-ячейки
    for r in range(rng.min_row, rng.max_row + 1):
        for c in range(rng.min_col, rng.max_col + 1):
            self._merged[(r, c)] = master
```

### 14.3 cell() с кэшем

```python
def cell(self, row: int, col: int) -> CellValue:
    if (row, col) in self._merged:
        return self._merged[(row, col)]  # O(1) dict lookup
    return self._ws.cell(row, col).value
```

---

## 15. Определение типов данных (dtype)

### 15.1 Алгоритм

```python
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
        # Percent-детекция: >30% значений содержат "%"
        pct_count = sum(1 for v in non_empty if isinstance(v, str) and "%" in v)
        return "percent" if pct_count / total > 0.3 else "number"
    return "text"
```

### 15.2 Пороги

| Тип | Условие |
|-----|---------|
| `date` | >50% значений — даты |
| `number` | >70% значений — числа, ≤30% — "%" |
| `percent` | >70% значений — числа, >30% — "%" |
| `text` | всё остальное |

---

## 16. Поток данных parse_file

```
parse_file(filepath, output_path, fmt, only_sheet, streaming)
│
├─ load_sheets(filepath, only_sheet)
│   ├─ _load_xlsx → [OpenpyxlAdapter, ...]
│   ├─ _load_xls  → [XlrdAdapter, ...]
│   ├─ _load_xlsb → [PyxlsbAdapter, ...]
│   └─ _load_csv  → [CsvAdapter]
│
├─ for adapter in adapters:
│   │
│   ├─ parse_sheet(adapter, wb, all_adapters)
│   │   ├─ _vis_cache.clear()
│   │   ├─ _extract_native_tables()    → used_rows += range
│   │   ├─ _extract_named_ranges()     → if not overlap: used_rows += range
│   │   ├─ _extract_heuristic()        → used_rows += range
│   │   ├─ _extract_vertical()         → used_rows += range
│   │   └─ _extract_headerless()       → used_rows += range
│   │
│   ├─ if streaming:
│   │   └─ writer.write_table(table)  → сразу на диск
│   └─ else:
│       └─ all_tables.extend(tables)
│
├─ if not streaming:
│   └─ _write_output(result, output_path, fmt)
│
└─ return result
```

---

## 17. История багов и фиксов

### Критические баги (потеря данных)

| Баг | Версия | Симптом | Фикс |
|-----|--------|---------|------|
| `_is_header_row` score 0.0 для заголовков с днями 1–31 | v17 | Таблицы с днями не находились | +0.5 для чисел-дней |
| `break` по маркер-строкам | v17 | Таблицы разрывались на "Исход", "Доп.часы" | Убран break |
| `max_empty_streak = 3` | v17 | Короткие разреженные секции разрывали таблицу | Увеличено до 50 |
| `skip_hidden = True` | v17 | Скрытые строки терялись (86 из 218 в "Все") | По умолчанию False |
| Headerless только если `not all_tables` | v17 | Маленькая heuristic таблица блокировала headerless | Запускать всегда |
| Двойное открытие файла | v17 | 2× load_workbook для xlsx | Один scan + один full |
| `_detect_dtype` читал ячейки дважды | v17 | N×M лишних вызовов adapter.cell() | Кэш col_values |
| `int(float('nan'))` → ValueError | v17 | Краш на ячейках с NaN | `math.isfinite` проверка |
| `except: pass` в 6+ местах | v18 | Ошибки молча игнорировались | `warnings.warn` |

### Баги корректности

| Баг | Версия | Симптом | Фикс |
|-----|--------|---------|------|
| `_dedupe_headers` не было | v10 | Дубли заголовков перезаписывали данные | Добавлен dedupe |
| xlrd.Ref3D.coords распаковка 5 вместо 6 | v14 | Named ranges в .xls сдвинуты | Правильная распаковка |
| `_extract_vertical` считал по всем строкам | v15 | Грязная статистика из used_rows | Счётчики только по свободным |
| `percent` при одной "%" ячейке | v12 | "рост 15% г/г" → вся колонка percent | >30% значений с "%" |
| `max()` с default для delimiter | v17 | Fallback возвращал произвольный символ | Явная проверка счётчика |

### Баги производительности

| Баг | Версия | Симптом | Фикс |
|-----|--------|---------|------|
| `_check_formulas_lazy` N раз | v16 | N× load_workbook для N листов | Один вызов на файл |
| `_visible` без кэша | v16 | 3× итерация row_dimensions на лист | Кэш _vis_cache |
| Pyxlsb 2 прохода кэша | v13 | Файл читался дважды | Один проход |
| CsvAdapter iter_rows_lazy читал файл | v13 | Файл читался дважды | Итерация по кэшу |
| StreamingWriter накапливал rows | v14 | `--stream` не работал | Потоковая запись |

### Архитектурные улучшения

| Изменение | Версия | Описание |
|-----------|--------|----------|
| `SheetAdapter` → `abc.ABC` | v18.1 | TypeError при неполном подклассе |
| 68 unit-тестов | v18.1 | Покрытие всех фиксов v11–v18 |
| `_extract_vertical` всегда | v12 | Убран guard `if not all_tables` |
| `named_ranges` всегда | v13 | Убран guard `and not native` |
| `max()` → явная проверка | v17 | Fallback delimiter корректен |
| `import traceback` → top-level | v17 | Не искать sys.modules в except |
| `l` → `letter` в col_info | v17 | Переименована неоднозначная переменная |
