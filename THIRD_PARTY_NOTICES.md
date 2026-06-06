# Third-Party Notices

This repository combines original project code and documentation with external data sources, open-source libraries, map services, and institutional names. The repository licenses cover only the materials for which the project maintainers hold the relevant rights.

## Data, Statistics, And Geodata

- Settlement names, administrative hierarchy, population counts, demographic components, geospatial coordinates, and central-place classifications may be derived from official statistics, public administrative registries, geodata providers, or author preprocessing of those sources.
- Official statistical and geodata materials remain subject to the terms, attribution requirements, and publication rules of their original providers.
- Reuse of derived datasets should preserve source attribution, extraction date, preprocessing notes, and any known restrictions.
- Territorial names and administrative labels are used as data attributes and do not imply endorsement, political position, or legal interpretation by the project.

## Map Services And Basemap Data

- The dashboard uses CARTO raster basemap tiles from `https://basemaps.cartocdn.com/`.
- CARTO basemaps use OpenStreetMap data and require attribution to OpenStreetMap contributors and CARTO.
- OpenStreetMap data is governed by the Open Database License and related OpenStreetMap Foundation terms.
- CARTO tiles and services remain subject to CARTO's own terms of service.

## Client-Side Libraries And Tooling

The dashboard loads or uses third-party software under the licenses published by each upstream project. Current runtime and development dependencies include:

| Component | Purpose |
| --- | --- |
| DuckDB-Wasm | In-browser analytical queries over Parquet files |
| MapLibre GL JS | Web map rendering |
| deck.gl | Point and text layers on the map |
| Plotly.js | Interactive charts and PNG export |
| Tom Select | Multi-select controls |
| noUiSlider | Time and threshold sliders |
| SheetJS `xlsx` | XLSX export |
| html2canvas | DOM-to-PNG export support |
| Playwright | Browser-based tests |
| http-server | Local static server for tests and development |

CDN delivery through jsDelivr and unpkg is subject to the service terms of those providers.

## Logos, Names, And Badges

- GitHub, Creative Commons, OpenStreetMap, CARTO, library names, institutional names, and any related marks remain the property of their respective owners.
- Shields.io badges and external service badges are used for informational repository metadata only.
- The presence of a name, badge, or service reference does not imply endorsement.

---

# Уведомления О Материалах Третьих Сторон

Этот репозиторий объединяет оригинальный код и документацию проекта со сторонними источниками данных, библиотеками, картографическими сервисами и институциональными наименованиями. Лицензии репозитория распространяются только на материалы, права на которые принадлежат авторам проекта.

## Данные, Статистика И Геоданные

- Названия населенных пунктов, административная иерархия, численность населения, демографические компоненты, координаты и классификации опорных пунктов могут быть получены из официальной статистики, публичных административных реестров, геоданных или авторской предобработки таких источников.
- Официальные статистические и географические материалы остаются под условиями их первоначальных поставщиков.
- При повторном использовании производных наборов данных следует сохранять указание источников, дату выгрузки, описание предобработки и известные ограничения.
- Территориальные названия и административные метки используются как атрибуты данных и не означают одобрения, политической позиции или юридической интерпретации со стороны проекта.

## Картографические Сервисы И Подложка

- Дашборд использует растровую подложку CARTO с `https://basemaps.cartocdn.com/`.
- Подложки CARTO используют данные OpenStreetMap и требуют указания OpenStreetMap contributors и CARTO.
- Данные OpenStreetMap регулируются Open Database License и связанными условиями OpenStreetMap Foundation.
- Тайлы и сервисы CARTO остаются под условиями CARTO.

## Клиентские Библиотеки И Инструменты

Дашборд загружает или использует стороннее программное обеспечение на условиях лицензий, опубликованных соответствующими проектами. Текущие зависимости включают:

| Компонент | Назначение |
| --- | --- |
| DuckDB-Wasm | Аналитические запросы к Parquet-файлам в браузере |
| MapLibre GL JS | Отрисовка веб-карты |
| deck.gl | Точечные и текстовые слои на карте |
| Plotly.js | Интерактивные графики и экспорт PNG |
| Tom Select | Элементы множественного выбора |
| noUiSlider | Слайдеры времени и порогов |
| SheetJS `xlsx` | Экспорт XLSX |
| html2canvas | Экспорт DOM-элементов в PNG |
| Playwright | Браузерные тесты |
| http-server | Локальный статический сервер для тестов и разработки |

Доставка через CDN jsDelivr и unpkg регулируется условиями этих сервисов.

## Логотипы, Названия И Бейджи

- GitHub, Creative Commons, OpenStreetMap, CARTO, названия библиотек, институций и связанные обозначения принадлежат соответствующим правообладателям.
- Бейджи Shields.io и бейджи внешних сервисов используются только как справочная метаинформация репозитория.
- Упоминание названия, бейджа или сервиса не означает одобрения проекта соответствующей организацией.
