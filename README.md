# Sprut Importer

Python-скрипт для сборки иерархической базы спецификаций `.xls/.xlsx` в Excel-шаблон импорта SprutTP.

## Что делает

- Рекурсивно обходит `input/specs` и индексирует спецификации по коду из имени файла.
- Берёт головные спецификации из `input/head` (или `input/roots.txt`, если head пуст).
- Строит дерево вложенности по позициям типа `Сборочные единицы` и `Комплекты`.
- Поддерживает исполнения по нескольким колонкам количеств.
- Формирует:
  - `output/sprut_import.xlsx` с колонками `Type | Code | Name | Kol | Vhod | Pos`.
  - `output/report.xlsx` с вкладками ошибок/статистики.
  - `logs/run.log`.

## Установка

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e .
```

Если нужна поддержка старых `.xls`, пакет `xlrd==2.0.1` уже указан в зависимостях.

## Структура рабочей зоны

Скрипт автоматически создаёт каталоги, если их нет:

```text
C:\CODEX\SPRUTTECHNOBASE
  input
    head
    specs
  template
    sprut_template.xlsx
  output
    sprut_import.xlsx
    report.xlsx
  logs
    run.log
```

## Запуск

> Внимание: скрипт **не запускается автоматически** при установке.

```bash
python -m sprut_importer \
  --base "C:\\CODEX\\SPRUTTECHNOBASE" \
  --head_dir "input\\head" \
  --specs_dir "input\\specs" \
  --template "template\\sprut_template.xlsx" \
  --out "output\\sprut_import.xlsx" \
  --report "output\\report.xlsx"
```

## Правила обработки

- Имя файла: `<CODE> <NAME>.xls/xlsx`.
- Код головы берётся из начала имени файла до первого пробела.
- Для исполнений:
  - `k=0` — база.
  - `k=1` — `-01`, `k=2` — `-02`, ...
  - если в базе уже есть `-NN`, используется `NN + k`.
- В выгрузку попадают строки с `Kol > 0` для выбранного исполнения.
- Неизвестный `TypeText` маппится в `DOCUMENT` и попадает в отчёт `unknown_types`.
- Отсутствующие дочерние спецификации пишутся в `missing_spec_files`.
- Циклы вложенности ловятся и пишутся в `cycles`.

## Отчёт `report.xlsx`

Вкладки:

1. `summary`
2. `missing_spec_files`
3. `unknown_types`
4. `rows_without_qty`
5. `parse_errors`
6. `duplicates_by_code`
7. `cycles`
