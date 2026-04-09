# Техническая документация: Excel Viewer

> Веб-просмотрщик результатов парсинга. Автономный HTML-файл без зависимостей.

---

## Оглавление

1. [Обзор](#1-обзор)
2. [Архитектура](#2-архитектура)
3. [Загрузка файлов](#3-загрузка-файлов)
4. [Структура данных](#4-структура-данных)
5. [Рендеринг таблиц](#5-рендеринг-таблиц)
6. [Поиск и фильтрация](#6-поиск-и-фильтрация)
7. [Сортировка](#7-сортировка)
8. [Производительность](#8-производительность)
9. [CSS-архитектура](#9-css-архитектура)
10. [Совместимость](#10-совместимость)

---

## 1. Обзор

Excel Viewer — одностраничное веб-приложение (SPA) для визуализации JSON-результатов парсинга Excel-файлов. Работает полностью в браузере без серверной части.

### Ключевые характеристики

| Параметр | Значение |
|----------|----------|
| **Файлов** | 1 (`index.html`) |
| **Зависимости** | Нет (чистый HTML + CSS + JS) |
| **Размер** | ~20 КБ |
| **Макс. строк** | 2000 одновременно (остальные через поиск) |
| **Тема** | Тёмная (GitHub Dark palette) |
| **Язык интерфейса** | Русский |

### Сценарий использования

```
1. Сотрудник парсит Excel → получает файл_parsed.json
2. Открывает excel_viewer/index.html в браузере
3. Перетаскивает JSON на страницу
4. Видит таблицы с поиском, сортировкой и фильтрацией
```

---

## 2. Архитектура

### 2.1 Структура приложения

```
┌─────────────────────────────────────────────┐
│                 TopBar                      │
│  [Logo] [Загрузить]        [Файл] [Статист.]│
├─────────────────────────────────────────────┤
│           Upload Screen (начальный)          │
│  ┌───────────────────────────────────────┐  │
│  │  [Иконка] Загрузите JSON-файл         │  │
│  │  [Drag & Drop Zone]                   │  │
│  │  [Features grid: 4 возможности]       │  │
│  └───────────────────────────────────────┘  │
├─────────────────────────────────────────────┤
│           Data View (после загрузки)         │
│  ┌───────────────────────────────────────┐  │
│  │  [Tabs] [Поиск____________] [Badge]   │  │
│  ├───────────────────────────────────────┤  │
│  │  ┌───┬───┬───┬───┬───┐               │  │
│  │  │ A │ B │ C │ D │ E │  ← Header     │  │
│  ├───┼───┼───┼───┼───┤               │  │
│  │  │   │   │   │   │   │  ← Body       │  │
│  │  │   │   │   │   │   │               │  │
│  └──┴───┴───┴───┴───┴───────────────┘  │  │
└─────────────────────────────────────────────┘
```

### 2.1 Компоненты

| Компонент | Описание |
|-----------|----------|
| **TopBar** | Sticky навигация с логотипом, кнопкой загрузки, статистикой |
| **UploadScreen** | Начальный экран с drag-and-drop зоной и описанием возможностей |
| **DataView** | Основной вид: вкладки таблиц, поиск, таблица данных |
| **TableRenderer** | Динамический рендерер заголовков и строк |
| **SearchEngine** | Фильтрация строк с подсветкой совпадений |
| **SortEngine** | Сортировка по числовым и текстовым колонкам |

### 2.2 Состояние приложения

```javascript
let data = null;           // Распарсенный JSON
let curTable = 0;          // Индекс текущей таблицы
let sort = { col: null, asc: true };  // Текущая сортировка
let query = '';            // Поисковый запрос
```

---

## 3. Загрузка файлов

### 3.1 Способы загрузки

| Способ | Триггер | Обработчик |
|--------|---------|------------|
| **Кнопка «Загрузить»** | `click` на `.upload-btn` | `fileInput.click()` |
| **Клик по drop-зоне** | `click` на `.drop-zone` | `dropInput.click()` |
| **Drag & Drop** | `drop` на `.drop-zone` | `read(e.dataTransfer.files[0])` |
| **Глобальный Drop** | `drop` на `document.body` | Проверка `.json` расширения |

### 3.2 Обработка кнопки «Загрузить»

```javascript
uploadBtn.addEventListener('click', (e) => {
    e.preventDefault();
    fileInput.click();
});
```

**Важно:** `e.preventDefault()` предотвращает стандартное поведение `<label>`,
а `fileInput.click()` программно открывает системный диалог выбора файла.

### 3.3 Drag & Drop

```javascript
dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');  // Визуальная обратная связь
});

dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) read(e.dataTransfer.files[0]);
});
```

### 3.4 Чтение файла

```javascript
function read(file) {
    const r = new FileReader();
    r.onload = e => {
        try {
            data = JSON.parse(e.target.result);
            boot();
        } catch (err) {
            alert('Bad JSON: ' + err.message);
        }
    };
    r.readAsText(file);
}
```

**Обработка ошибок:** Невалидный JSON показывается через `alert` с описанием ошибки.

---

## 4. Структура данных

### 4.1 Ожидаемый формат JSON

Viewer ожидает результат работы `excel_smart_parser.py`:

```json
{
  "file": "имя_файла.xlsx",
  "format": "xlsx",
  "sheets": 3,
  "tables": 3,
  "total_rows": 847,
  "sources": { "heuristic": 3 },
  "tables_data": [
    {
      "sheet": "ИмяЛиста",
      "name": "ИмяЛиста / Таблица 1",
      "source": "heuristic",
      "columns": [
        { "letter": "A", "name": "ФИО", "type": "text" },
        { "letter": "B", "name": "Сумма", "type": "number" }
      ],
      "rows": [
        { "ФИО": "Иванов Иван", "Сумма": 15000 }
      ]
    }
  ]
}
```

### 4.2 Ключевые поля

| Поле | Тип | Описание |
|------|-----|----------|
| `tables_data` | `array` | Массив всех найденных таблиц |
| `tables_data[].columns` | `array` | Метаданные колонок (имя, тип, буква) |
| `tables_data[].columns[].type` | `string` | `"text"` или `"number"` |
| `tables_data[].rows` | `array` | Массив строк данных (объекты) |
| `tables_data[].rows[]` | `object` | Ключи = имена колонок, значения = данные |

### 4.3 Типизация колонок

| `type` | Поведение |
|--------|-----------|
| `"number"` | Выравнивание вправо, форматирование с разделителями, числовая сортировка |
| `"text"` | Выравнивание влево, строковая сортировка, без форматирования |

---

## 5. Рендеринг таблиц

### 5.1 Инициализация (`boot()`)

```javascript
function boot() {
    uploadScreen.style.display = 'none';    // Скрыть экран загрузки
    dataView.classList.add('active');       // Показать вид данных
    statsBar.style.display = 'flex';        // Показать статистику

    $('statFile').textContent = data.file || '—';
    $('statTables').textContent = data.tables || 0;
    $('statRows').textContent = (data.total_rows || 0).toLocaleString('ru-RU');

    buildTabs();
    switchTable(0);
}
```

### 5.2 Построение заголовков

```javascript
thead.innerHTML = '<tr>' + cols.map((c, i) => {
    const sorted = sort.col === i ? ' sorted' : '';
    const arrow = sort.col === i ? (sort.asc ? '▲' : '▼') : '⇅';
    return `<th data-i="${i}" class="${sorted}">
        <span class="th-inner">${esc(c.name)}
            <span class="sort-arrow">${arrow}</span>
        </span>
    </th>`;
}).join('') + '</tr>';
```

**XSS-защита:** Имена колонок проходят через `esc()` — экранирование `<`, `>`, `&`, `"`.

### 5.3 Рендеринг строк

```javascript
tbody.innerHTML = show.map(r =>
    '<tr>' + cols.map((c, i) => {
        let v = r[c.name];
        if (v == null) v = '';
        const n = numSet.has(i) && typeof v === 'number';
        const cls = n ? ' class="num"' : '';
        const txt = n ? fmtNum(v) : esc(String(v));
        return `<td${cls}>${hl(txt)}</td>`;
    }).join('') + '</tr>'
).join('');
```

### 5.4 Ограничение рендеринга

```javascript
const MAX = 2000;
const show = rows.slice(0, MAX);
```

**Причина:** Рендеринг >5000 DOM-элементов вызывает лаги браузера.
При превышении показывается сообщение «… ещё N строк — уточните поиск».

### 5.5 Форматирование чисел

```javascript
function fmtNum(n) {
    return Number.isInteger(n)
        ? n.toLocaleString('ru-RU')
        : n.toLocaleString('ru-RU', { maximumFractionDigits: 4 });
}
```

**Примеры:**
- `15000` → `15 000`
- `3537.9` → `3 537,9`
- `438.4015` → `438,4015`

---

## 6. Поиск и фильтрация

### 6.1 Алгоритм фильтрации

```javascript
if (query) {
    const q = query.toLowerCase();
    rows = rows.filter(r =>
        cols.some(c => {
            const v = r[c.name];
            return v != null && String(v).toLowerCase().includes(q);
        })
    );
}
```

**Логика:** Строка остаётся если **хотя бы одно** поле содержит запрос (case-insensitive).

### 6.2 Debounce ввода

```javascript
let tid;
searchInput.addEventListener('input', e => {
    clearTimeout(tid);
    tid = setTimeout(() => {
        query = e.target.value.trim();
        renderRows();
    }, 180);
});
```

**Причина:** Без debounce рендеринг вызывался бы на каждый символ — лаги при быстром вводе.

### 6.3 Подсветка совпадений

```javascript
function hl(t) {
    if (!query) return t;
    return t.replace(
        new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'),
        '<mark>$&</mark>'
    );
}
```

**Экранирование regex:** `replace(/[.*+?^${}()|[\]\\]/g, '\\$&')` предотвращает
ошибки при поиске спецсимволов (`.`, `*`, `?`, `[`, `]`, и т.д.).

### 6.4 Badge с результатом

```javascript
if (query) {
    rowCount.className = 'badge badge-match';
    rowCount.textContent = `${rows.length} из ${total}`;
} else {
    rowCount.className = 'badge badge-default';
    rowCount.textContent = `${rows.length} строк`;
}
```

---

## 7. Сортировка

### 7.1 Инициализация

Клик по заголовку `<th>` переключает сортировку:

```javascript
thead.querySelectorAll('th').forEach(th => {
    th.onclick = () => {
        const i = parseInt(th.dataset.i);
        if (sort.col === i) sort.asc = !sort.asc;   // Переключение направления
        else { sort.col = i; sort.asc = true; }     // Новая колонка, по возрастанию
        renderRows();
    };
});
```

### 7.2 Алгоритм сортировки

```javascript
rows.sort((a, b) => {
    let va = a[key], vb = b[key];
    if (va == null) return 1;       // null в конец
    if (vb == null) return -1;

    // Числовая сортировка
    if (num && typeof va === 'number' && typeof vb === 'number')
        return sort.asc ? va - vb : vb - va;

    // Строковая сортировка
    va = String(va).toLowerCase();
    vb = String(vb).toLowerCase();
    if (va < vb) return sort.asc ? -1 : 1;
    if (va > vb) return sort.asc ? 1 : -1;
    return 0;
});
```

### 7.3 Типы сортировки

| Тип колонки | Метод | Пример |
|-------------|-------|--------|
| `number` + число | Числовое сравнение | `100 < 1000` |
| `text` или mixed | Строковое (lexicographic) | `"Алексей" < "Борис"` |
| `null`/`undefined` | Всегда в конец | — |

### 7.4 Визуальная индикация

| Состояние | Стрелка |
|-----------|---------|
| Не сортируется | `⇅` (полупрозрачная) |
| По возрастанию | `▲` (яркая) |
| По убыванию | `▼` (яркая) |

---

## 8. Производительность

### 8.1 Таблица оптимизаций

| Оптимизация | Описание | Эффект |
|-------------|----------|--------|
| **Debounce поиска** | 180мс задержка перед рендерингом | -80% перерисовок при вводе |
| **Лимит 2000 строк** | `slice(0, 2000)` перед рендерингом | Стабильный FPS при больших данных |
| **Set для числовых колонок** | `numSet = new Set()` | O(1) проверка типа vs O(n) каждый раз |
| **InnerHTML batch** | Весь tbody за одну операцию | Один reflow вместо N |
| **Esc на лету** | `esc()` при генерации HTML | Нет дополнительного прохода |

### 8.2 Сложность операций

| Операция | Сложность | Описание |
|----------|-----------|----------|
| **Фильтрация** | O(n × m) | n строк × m колонок |
| **Сортировка** | O(n log n × m) | Сравнение по m колонкам |
| **Рендеринг** | O(k × m) | k отображаемых строк (≤2000) |
| **Поиск + подсветка** | O(k × m × L) | L = длина строки для regex |

### 8.3 Память

| Компонент | Потребление |
|-----------|-------------|
| **Парсенный JSON** | ~100-500 КБ на 1000 строк |
| **DOM-дерево** | ~2-5 МБ на 2000 строк |
| **Общий расход** | ~5-10 МБ для типичного файла |

---

## 9. CSS-архитектура

### 9.1 Цветовая палитра (GitHub Dark)

| Переменная | Значение | Назначение |
|------------|----------|------------|
| `--bg` | `#0d1117` | Основной фон страницы |
| `--bg-secondary` | `#161b22` | Фон topbar-статистики |
| `--surface` | `#1c2128` | Фон карточек, таблиц, контролов |
| `--border` | `#30363d` | Границы элементов |
| `--border-light` | `#3d444d` | Скроллбар |
| `--text` | `#e6edf3` | Основной текст |
| `--text-secondary` | `#8b949e` | Вторичный текст |
| `--text-muted` | `#6e7681` | Placeholder, подсказки |
| `--accent` | `#58a6ff` | Основной акцент (кнопки, ссылки) |
| `--accent-hover` | `#79b8ff` | Ховер акцента |
| `--green` | `#3fb950` | Успех, галочки |
| `--orange` | `#d29922` | Подсветка поиска |
| `--purple` | `#bc8cff` | Декоративный |

### 9.2 Sticky-элементы

| Элемент | Позиция | Z-index |
|---------|---------|---------|
| **TopBar** | `top: 0` | 100 |
| **Table thead** | `top: 0` | 10 |

### 9.3 Backdrop-filter

```css
.topbar {
    background: rgba(13,17,23,.85);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
}
```

**Эффект:** Полупрозрачный topbar с размытием контента под ним (как в macOS).

### 9.4 Адаптивность

```css
@media (max-width: 640px) {
    .topbar { padding: 0 16px; }
    .ctrl-bar { padding: 10px 16px; }
    .topbar-stats { display: none; }
    .features { grid-template-columns: 1fr; }
    .upload-card { padding: 32px 24px; }
}
```

**Поведение на мобильных:**
- Статистика в topbar скрывается
- Features grid становится одноколоночным
- Уменьшены отступы

---

## 10. Совместимость

### 10.1 Браузеры

| Браузер | Мин. версия | Примечание |
|---------|-------------|------------|
| Chrome | 88+ | Полная поддержка |
| Firefox | 90+ | Полная поддержка |
| Safari | 14+ | `-webkit-backdrop-filter` fallback |
| Edge | 88+ | Chromium-based |

### 10.2 Используемые API

| API | Поддержка |
|-----|-----------|
| `FileReader` | Все современные браузеры |
| `backdrop-filter` | Chrome 76+, Firefox 103+, Safari 9+ |
| `CSS custom properties` | Все современные браузеры |
| `Array.prototype.flatMap` | Chrome 69+, Firefox 62+, Safari 12+ |
| `Set` | Все современные браузеры |
| `Intl.NumberFormat` | Все современные браузеры |

### 10.3 Ограничения

| Ограничение | Описание |
|-------------|----------|
| **2000 строк** | Больше — через поиск. Для полного просмотра нужен серверный рендеринг |
| **Только JSON** | Не работает с XLSX напрямую — нужен парсер |
| **Нет экспорта** | Просмотр только, без скачивания в другом формате |
| **Нет пагинации** | Бесконечный скролл внутри таблицы |

---

## 11. Структура файла

```
index.html
├── <head>
│   ├── <meta charset>, <meta viewport>
│   └── <style>           /* ~300 строк CSS */
├── <body>
│   ├── .topbar           /* Sticky навигация */
│   ├── .upload-screen    /* Начальный экран */
│   └── .data-view        /* Основной вид */
│       ├── .ctrl-bar     /* Вкладки + поиск */
│       ├── .table-wrap   /* Контейнер таблицы */
│       └── .empty-state  /* «Ничего не найдено» */
└── <script>              /* ~200 строк JS */
    ├── State variables
    ├── DOM refs
    ├── Upload handlers
    ├── boot()
    ├── buildTabs()
    ├── render()
    ├── renderRows()
    ├── Search handler
    └── Helpers (esc, hl, fmtNum)
```

---

## 12. Безопасность

### 12.1 XSS-защита

Все пользовательские данные экранируются перед вставкой в HTML:

```javascript
function esc(s) {
    return s.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
}
```

**Применяется к:**
- Именам колонок (из JSON)
- Значениям ячеек (из JSON)

**Не применяется к:**
- HTML-атрибутам (используются template literals с `data-*`)
- Поисковым совпадениям (оборачиваются в `<mark>`, но текст совпадения экранирован до regex)

### 12.2 Regex-безопасность

```javascript
query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
```

Предотвращает **ReDoS** и синтаксические ошибки при поиске спецсимволов.

### 12.3 Локальная обработка

- Данные **не покидают** компьютер пользователя
- Нет сетевых запросов (fetch/XHR)
- Нет cookie, localStorage, sessionStorage
- Файл работает через `file://` протокол

---

## 13. Расширение

### 13.1 Возможные улучшения

| Фича | Сложность | Описание |
|------|-----------|----------|
| **Экспорт в CSV** | Низкая | Кнопка на таблице, генерация Blob |
| **Пагинация** | Средняя | Разбиение на страницы по 100 строк |
| **Фильтры по колонкам** | Средняя | Dropdown с уникальными значениями |
| **Светлая тема** | Низкая | CSS переменные + toggle |
| **Прямая загрузка XLSX** | Высокая | Интеграция SheetJS (xlsx.js) |
| **Горячие клавиши** | Низкая | `Ctrl+F`, `Ctrl+O`, `←/→` для вкладок |

### 13.2 Добавление новой фичи

Пример: кнопка экспорта в CSV

```javascript
// В .ctrl-bar добавить кнопку
const exportBtn = document.createElement('button');
exportBtn.textContent = 'Экспорт CSV';
exportBtn.onclick = () => exportCSV();
ctrlBar.appendChild(exportBtn);

function exportCSV() {
    const tbl = data.tables_data[curTable];
    const cols = tbl.columns;
    const header = cols.map(c => c.name).join(',');
    const rows = tbl.rows.map(r =>
        cols.map(c => JSON.stringify(r[c.name] ?? '')).join(',')
    );
    const blob = new Blob([header + '\n' + rows.join('\n')], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${tbl.name}.csv`;
    a.click();
}
```
