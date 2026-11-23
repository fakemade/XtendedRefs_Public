# XtendedRefs — COM-надстройка для Microsoft Word

> Управление перекрёстными ссылками на рисунки, формулы, таблицы и литературу прямо в документе.

[![GitHub release](https://img.shields.io/badge/dynamic/json?label=Release&query=$.tag_name&url=https%3A//api.github.com/repos/SelivanovU/XtendedRefs_sharp/releases/latest)](../../releases)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()
[![Language](https://img.shields.io/badge/language-C%23-orange)]()

## 📋 Содержание

- [Что это](#что-это)
- [Возможности (статус)](#возможности-статус)
- [Архитектура проекта](#архитектура-проекта)
- [Установка](#установка)
  - [Для пользователя (готовый установщик)](#для-пользователя-готовый-установщик)
  - [Для разработчика (dev-build)](#для-разработчика-dev-build)
- [Использование](#использование)
  - [Вкладка «Оформление»](#вкладка-оформление)
  - [Настройки документа](#настройки-документа)
  - [Где хранятся данные](#где-хранятся-данные)
- [Скриншоты](#скриншоты)
- [Технические детали](#технические-детали)
- [FAQ / Известные ограничения](#faq--известные-ограничения)

---

## Что это

**XtendedRefs** — COM-надстройка (DLL, загружается как InprocServer32), которая встраивается в Microsoft Word и добавляет вкладку **«Оформление»** на ленту. Она помогает авторам технической документации:

- Создавать пронумерованные подписи к рисункам и таблицам с учётом глав.
- Работать с формулами MathType (создание, удаление, превью).
- Ведать список литературы: добавлять источники, вставлять ссылки на них в текст, перенумеровывать при удалении.
- Всё хранится **внутри DOCX** — закладки, поля REF, переменные документа и кеш медиафайлов (EMF) в `%TEMP%`.

**Технические параметры:**

| | |
|---|---|
| Язык / фреймворк | C# · .NET Framework 4.7.2 |
| Тип COM-объекта | InprocServer32, `HKCU` (без прав администратора) |
| CLSID | `{1A2D0B5D-3B55-4F1D-9C69-2D6A54FB2C44}` |
| ProgID | `XtendedRefs.ComAddIn` |
| Целевой Word | 2016+ (x86 / x64, определяется автоматически) |
| UI | Ribbon XML + WPF-диалоги (pickers) |

---

## Возможности (статус)

Статус соответствует текущему **разрабатываемому** состоянию. Новые команды пишутся в `Application/Commands/NewCommands/` — старая логика (`FigureCommands`, `TableCommands`) будет заменена по мере готовности новой доменной модели.

### ✅ Полностью рабочая

| Фича | Что умеет сейчас |
|---|---|
| **Литература** | Добавление источника, удаление (с авто-перенумеровкой), вставка REF-ссылки на источник из picker'а. Двойные закладки: `[N]` и полный текст. |

### 🟡 Почти готово (нужны доработки)

| Фича | Что работает | Чего не хватает |
|---|---|---|
| **Формулы** | Создание формулы из MathType (таблица 3 колонки, SEQ-номер), вставка REF-ссылки, обновление превью EMF. | ❌ Нет команды «Удалить формулу» · Баг: не все поля REF обновляются после изменения документа. |

### 🟢 Завершено (альфа)

| Фича | Что работает |
|---|---|
| **Главы** | Сканирование заголовков, построение оглавления на лету, изменение стилей ссылок, обновление при правках текста. |

### 🔴 Не реализовано / переписывается

| Фича | Текущее состояние | Планы |
|---|---|---|
| **Рисунки** | Старая версия (`FigureCommands`) — создание подписи, вставка REF. Будет переписана на новую доменную модель `NewModels/`. | Новая версия с unified picker'ом и поддержкой шаблонов с несколькими `$` (по главам). |
| **Таблицы** | Старая версия (`TableCommands`) — подпись над/под таблицей, REF. Будет переписана вместе с рисунками. | Та же новая модель: подписи `Таблица {STYLEREF "Heading 1" \n}.{SEQ Table_Seq \n} — ...` |
| **Шаблоны** | Настройки хранятся как переменные документа (`DocumentSettings`). | Формат шаблонов для переноса между документами/устройствами. |
| **Батч-вставка литературы** | По одной за раз. | Вставка источников по абзацам, пакетная вставка ссылок. |

---

## Архитектура проекта

```
┌─────────────────────────────────────┐
│  Ribbon XML (MainRibbon.xml)        │  Декларация UI-групп: Рисунки / Формулы / Литература / Таблицы
├─────────────────────────────────────┤
│  Connect.cs + AddInContext           │  COM entry point · IDTExtensibility2, IRibbonExtensibility
├──────────┬──────────────────────────┤
│ Features/│ UI/Windows (WPF)         │  Команды пользователя + диалоговые picker'ы
├──────────┴──────────────────────────┤
│ Infrastructure/                     │  Сервисы Word API: закладки, поля REF, переменные, EMF-буфер, DOCX-media map
├─────────────────────────────────────┤
│ Domain/                             │  Чистые модели: DocumentSettings, NewModels/, NumberingState
└─────────────────────────────────────┘
```

Зависимости идут **сверху вниз**: Commands → Infrastructure → Domain. Никогда наоборот.

**Потоки данных:**

1. Пользователь нажимает кнопку на ленте → `Connect` делегирует в нужную команду.
2. Команда использует `Infrastructure/` сервисы для манипуляций с Word (закладки, поля REF).
3. Состояние сохраняется в переменные документа (`DocumentVariableStore`) и в `RuntimeStore` (in-memory singleton).
4. Превью изображений (EMF) кешируются в `%TEMP%\XtendedRefs\{Figures|Formulas|Tables}\` и извлекаются при открытии документа через `DocxMediaPreviewMap` (парсинг DOCX как ZIP-архива).

---

## Установка

### Для пользователя (готовый установщик)

Скрипт сборки:

```powershell
# В корне проекта:
.\Build-Installer.ps1           # соберёт DLL и запустит ISCC setup.iss
```

Результат — `bin\XtendedRefsInstaller.exe`. При запуске он:

1. Определяет разрядность установленного Word (x86 / x64) по ключам реестра Office 14.0–16.0.
2. Копирует DLL в папку `HKCU\Software\Microsoft\Office\Word\Addins\XtendedRefs.ComAddIn`.
3. Регистрирует COM-объект и создаёт ярлык на ленте.

> ✅ Не требует прав администратора — всё пишется в HKCU.

### Для разработчика (dev-build)

```powershell
.\build-dev.ps1                 # компиляция Debug DLL + dev-registration
# или
Register-DevAddin.ps1           # только регистрация (если DLL уже есть)
Use-ReleaseAddin.ps1            # для release-версии
```

Для отписки: `dev-unregister.ps1` / `UnregisterAddin.ps1`.

---

## Использование

### Вкладка «Оформление»

После установки в Word появляется вкладка **«Оформление»** с четырьмя группами кнопок:

| Группа | Кнопки |
|---|---|
| **Рисунки** | Создать подпись · Удалить подпись · Вставить ссылку |
| **Формулы** | Создать формулу · Удалить формулу · Обновить превью · Вставить ссылку |
| **Литература** | Добавить источник · Удалить источник · Вставить ссылку |
| **Таблицы** | Создать подпись · Удалить подпись · Вставить ссылку |

### Настройки документа

Нажмите **«Настройки»** (иконка шестерёнки) — откроется WPF-окно `SettingsWindow`. Там можно задать:

- Шаблон подписи рисунка / таблицы (`$` = вставка номера главы).
- Стиль заголовка для извлечения номерa главы (`STYLEREF`).
- Разделитель между номером и текстом источника литературы (пробел, табуляция, пустое).
- Позицию подписи таблицы (над / под).

Настройки сохраняются как **переменные документа** — переносятся с файлом DOCX.

### Где хранятся данные

| Что | Где |
|---|---|
| Шаблоны, стили, разделители | Переменные Word (`DocumentVariableStore`) |
| Список литературы, формул, таблиц | `RuntimeStore` (in-memory) + закладки в документе |
| Превью изображений | `%TEMP%\XtendedRefs\{Figures|Formulas|Tables}\` (EMF-файлы) |
| Карта медиа DOCX | Переменная `MediaMap` в документе (сериализованный JSON) |

---

## Скриншоты

<!-- Screenshot: Ribbon tab «Оформление» -->
![Скриншот ленты](https://placehold.co/900x400?text=+Риббон+%2F+Ribbon+Tab+)

<!-- Screenshot: SettingsWindow — настройки шаблонов и стилей -->
![Настройки](https://placehold.co/750x600?text=Settings+Window)

<!-- Screenshot: LiteraturePicker — список источников с превью -->
![Literature picker](https://placehold.co/800x500?text=Literature+Picker)

<!-- Screenshot: Formula preview update -->
![Formula preview](https://placehold.co/900x350?text=Formula+Preview+Update)

<!-- Screenshot: Table caption above/below toggle -->
![Table caption position](https://placehold.co/700x400?text=Table+Caption+Position)

---

## Технические детали

### Команда на ленте — поток выполнения

```
User clicks button
  → Connect.OnCreateFigurePressed(IRibbonControl)
  → RunCommand(Action, "Create Figure")   // try-catch + MessageBox
  → new FigureCommands().CreateFigure()   // или LiteratureCommands/AddLiterature() и т.д.
  → Infrastructure/BookmarkNavigator    (найти существующие закладки)
  → Infrastructure/ReferenceFieldUpdater (манипуляции REF-полями)
  → Infrastructure/ClipboardEmfService    (EMF через P/Invoke Win32 API)
  → Domain/DocumentSettings               (читать шаблоны из переменных документа)
```

### Подпись рисунка / таблицы

Шаблон по умолчанию: `Рисунок {STYLEREF "Heading 1" \n}.{SEQ Figure_Seq \n} — ...`

- `{STYLEREF "Heading 1" \n}` → номер главы из стиля заголовка (2.3).
- `{SEQ Figure_Seq \n}` → автоматический счётчик Word (2).
- Результат: `Рисунок 2.3 — Схема архитектуры`.

Поддерживается несколько `$` для многоуровневой нумерации (`{STYLEREF "Heading 1" \n}.{STYLEREF "Heading 2" \n}....`).

### Закладки (bookmark prefixes)

| Префикс | Значение |
|---|---|
| `Figure_NNN` | Полный абзац подписи рисунка |
| `FigureNum_NNN` | Только номер |
| `Formula_NNN` | Средний столбец формулы (уравнение) |
| `Table_NNN` | Полный абзац подписи таблицы |
| `TableNum_NNN` | Только номер |
| `Literature_NNN` | Порядковый номер источника `[N]` |
| `LiteratureText_NNN` | Полный текст источника |

### События Word (автоматика)

- **DocumentOpen** → загрузка карты медиа, фоновое кеширование превью.
- **DocumentBeforeSave** → пересчёт карты медиа, сохранение в DOCX.
- **WindowSelectionChange после MathType** → обновить EMF формулы через 1 сек.

---

## FAQ / Известные ограничения

### Q: Работает ли надстройка на Mac?
A: Нет — COM InprocServer32 + P/Invoke Win32 API. Только Windows.

### Q: Что если Word установлен не в стандартной папке?
A: Установщик ищет по ключам реестра `HKLM\SOFTWARE\Microsoft\Office\*.0\Word\InstallRoot`. Работает для 14.0 (2010), 15.0 (2013) и 16.0 (2016+).

### Q: Ссылки REF не обновляются после удаления источника?
A: В текущей версии `DeleteLiterature` пересчитывает номера и вызывает `ReferenceFieldUpdater.RenameBookmarksInRefFields()`. Если что-то осталось — нажмите **Обновить все поля** (Ctrl+A → F9) в Word.

### Q: Превью формул не появляется?
A: Проверьте, что MathType открыт при нажатии «Создать формулу». EMF захватывается из буфера обмена через `CopyEnhMetaFile`. Если нет — нажмите кнопку «Обновить превью» повторно.

---

<!-- ENGLISH STUB -->

## English (coming soon)

> This extension requires localization and adaptation for international audiences.
> The README is currently in Russian only. We plan to add full English documentation,
> translated UI labels, and locale-aware default templates in a future release.

---

*Сборка: `Build-Installer.ps1` · Идентификатор пакета: `{A8C3F156-2E47-4B9D-B362-DC4F87213A5E}`*