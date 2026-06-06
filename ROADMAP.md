# Roadmap

This roadmap is non-binding and intended to help contributors choose useful work.

## Near Term

- Maintain `docs/methodology.md` as the canonical research-readiness document.
- Add a compact data dictionary for all Parquet tables and key fields.
- Add automated README and docs link checks.
- Keep i18n tests stable across desktop and mobile viewports.
- Record data build identifiers, checksums, and row counts for each research release.

## Medium Term

- Add a machine-readable data manifest with source, extraction date, and preprocessing metadata.
- Add QA checks for duplicate settlement identifiers, missing coordinates, and municipal key coverage.
- Add scenario-comparison exports for `withMIG` and `noMIG`.
- Document sensitivity protocols for fertility, mortality, migration, and age-sex allocation assumptions.
- Improve accessibility of controls, tables, and chart export labels.

## Longer Term

- Support uncertainty or percentile bands if upstream projection data are available.
- Add optional downloadable research bundles with manifest, checksums, and citation metadata.
- Add more granular validation against regional or national control totals.
- Provide reproducible preprocessing notebooks or scripts when source-data licensing allows.

---

# Дорожная Карта

Дорожная карта не является обязательством по срокам. Она помогает выбирать полезные задачи.

## Ближайшие Задачи

- Поддерживать `docs/methodology.md` как основной документ исследовательской готовности.
- Добавить компактный словарь данных для всех Parquet-таблиц и ключевых полей.
- Добавить автоматическую проверку ссылок README и документации.
- Поддерживать i18n-тесты на desktop и mobile viewports.
- Фиксировать идентификаторы сборки данных, контрольные суммы и счетчики строк для каждого исследовательского релиза.

## Среднесрочные Задачи

- Добавить машинно-читаемый data manifest с источниками, датами выгрузки и метаданными предобработки.
- Добавить QA-проверки дублей идентификаторов, отсутствующих координат и покрытия муниципальных ключей.
- Добавить экспорт сравнения сценариев `withMIG` и `noMIG`.
- Описать протоколы чувствительности для рождаемости, смертности, миграции и аллокации половозрастной структуры.
- Улучшить доступность элементов управления, таблиц и подписей экспорта графиков.

## Дальнейшие Задачи

- Поддержать интервалы неопределенности или percentile-bands, если такие данные появятся upstream.
- Добавить скачиваемые исследовательские bundles с manifest, checksums и citation metadata.
- Расширить валидацию относительно региональных или национальных контрольных сумм.
- Опубликовать воспроизводимые scripts/notebooks предобработки, если это допускают лицензии исходных данных.
