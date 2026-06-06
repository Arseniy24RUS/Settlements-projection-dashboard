# Good First Issues

These tasks are intentionally scoped for new contributors. Please open or claim an issue before starting work.

## Documentation

- Add `docs/data-dictionary.md` with table names, key fields, and short descriptions.
- Add a short glossary for projection terms used in the dashboard: TFR, CDR, net migration, age-sex profile, central settlement.
- Add source-provenance placeholders to `docs/methodology.md` for the next data refresh.
- Review Russian and English documentation for terminology consistency.

## Testing And QA

- Add a script that checks local image and Markdown links in `README.md` and `docs/`.
- Add a lightweight JSON validation check for `data/config.json`.
- Add a smoke test that opens the methodology dialog in both languages.
- Document manual QA steps for map rendering and XLSX/PNG exports.

## Data Readiness

- Create a draft data manifest schema for Parquet files, checksums, row counts, and source dates.
- Add a checklist for validating `settlement_id` uniqueness and `oktmo_syn` coverage.
- Draft a small sensitivity-analysis template comparing `withMIG` and `noMIG`.

---

# Простые Задачи Для Первого Вклада

Эти задачи специально ограничены по объему для новых участников. Перед началом работы откройте issue или напишите, что берете его в работу.

## Документация

- Добавить `docs/data-dictionary.md` с названиями таблиц, ключевыми полями и краткими описаниями.
- Добавить краткий глоссарий терминов: СКР, ОКС, сальдо миграции, половозрастной профиль, опорный населенный пункт.
- Добавить заготовки provenance в `docs/methodology.md` для следующего обновления данных.
- Проверить русскую и английскую документацию на единообразие терминов.

## Тестирование И QA

- Добавить script для проверки локальных изображений и Markdown-ссылок в `README.md` и `docs/`.
- Добавить легкую проверку JSON для `data/config.json`.
- Добавить smoke-test, который открывает диалог методики на двух языках.
- Описать ручные QA-шаги для карты и экспорта XLSX/PNG.

## Готовность Данных

- Создать draft schema для data manifest: Parquet-файлы, контрольные суммы, счетчики строк и даты источников.
- Добавить checklist проверки уникальности `settlement_id` и покрытия `oktmo_syn`.
- Подготовить шаблон анализа чувствительности для сравнения `withMIG` и `noMIG`.
