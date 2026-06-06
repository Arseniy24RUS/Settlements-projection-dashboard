# Settlements Projection Dashboard · Spatial Population Analytics

[English](#english) · [Русский](#русский)

[![Live demo](https://img.shields.io/badge/demo-GitHub%20Pages-blue)](https://arseniy24rus.github.io/Settlements-projection-dashboard/)
[![Code: MIT](https://img.shields.io/badge/code-MIT-blue.svg)](LICENSE)
[![Data/docs: CC BY 4.0](https://img.shields.io/badge/data%20%26%20docs-CC%20BY%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)

---

## English

### Overview

`Settlements-projection-dashboard` is a static web dashboard for spatial analysis of settlement-level population projections. It is designed for GitHub Pages and other static hosting environments and uses compact data files, client-side visualization and map-based exploration.

The project is intended as an applied instrument for regional studies, urban and rural analysis, public administration courses and exploratory demographic visualization. It helps connect demographic projection results with the geography of settlements and makes spatial differences more visible to researchers, students and policy analysts.

### Visual overview

![Settlements Projection Dashboard English hero](assets/visuals/hero-en.png)

![English scenario walkthrough: filters, settlement click, charts and pyramid update](assets/visuals/scenario-en.gif)

![English architecture diagram](assets/visuals/architecture-en.svg)

![English methodology diagram](assets/visuals/methodology-en.svg)

### Live dashboard

GitHub Pages: <https://arseniy24rus.github.io/Settlements-projection-dashboard/>

### Analytical purpose

The dashboard answers a practical question: how can projected demographic change be inspected at the settlement level rather than only at the level of large administrative regions? The interface is designed to support comparison, filtering, mapping and visual inspection of population dynamics across localities.

Detailed research methodology: [docs/methodology.md](docs/methodology.md).

### Repository structure

```text
assets/       JavaScript, CSS, icons and auxiliary front-end assets
data/         Dashboard datasets, including compact analytical files
index.html    Main static dashboard page
README.md     Project documentation
```

### Deployment model

The dashboard is a static website. It can be published directly from the repository root through GitHub Pages. Because browsers restrict access to local files, the dashboard should be opened through HTTP/HTTPS rather than as a `file://` document.

Local launch:

```bash
python -m http.server 8000
```

Then open <http://localhost:8000/>.

### Data model

The project is organized around compact client-side data files. Depending on the build version, these may include Parquet, CSV, JSON or GeoJSON-like resources used for tables, maps and charts. The recommended documentation standard is to maintain a `data_dictionary.md` file describing settlement identifiers, territorial hierarchy, years, demographic variables, projection scenario labels and preprocessing steps.

### Suggested interpretation

The dashboard should be used as a visual and exploratory analytical tool. It is suitable for identifying broad spatial patterns, comparing settlements and preparing research questions. It should not be treated as a substitute for a full methodological report or an official demographic forecast.

### Quality checks

Recommended checks include: loading the page through a local HTTP server; verifying that all data files are reachable; testing map rendering; checking that filters and charts react to user input; checking that export functions, if present, return valid data; and verifying that the dashboard works on current desktop browsers.

### Citation

If you use the dashboard, data structure or visualization design, please cite:

> Sitkovskiy, A. M. (2026). Settlements Projection Dashboard: spatial population analytics. GitHub. https://github.com/Arseniy24RUS/Settlements-projection-dashboard

### License

| Material | License | Notes |
| --- | --- | --- |
| Source code | [MIT](LICENSE) | Application code, tests and build/runtime scripts. |
| Documentation, data and dashboard content | [CC BY 4.0](LICENSE-DOCS-AND-DATA.md) | Documentation, data files, figures, screenshots and text unless otherwise stated. |
| Third-party data, maps, libraries, logos and institutional names | Original provider terms | See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). |

---

## Русский

### Обзор

`Settlements-projection-dashboard` — статический веб-дашборд для пространственного анализа демографических прогнозов на уровне населённых пунктов. Проект рассчитан на публикацию через GitHub Pages и другие статические хостинги, использует компактные файлы данных, клиентскую визуализацию и картографическое представление.

Проект предназначен для региональных исследований, анализа городских и сельских территорий, учебных курсов по государственному управлению и разведочной демографической визуализации. Он помогает связать результаты демографического прогноза с географией населённых пунктов и делает пространственные различия более наглядными для исследователей, студентов и аналитиков политики.

### Визуальный обзор

![Главный экран дашборда на русском языке](assets/visuals/hero-ru.png)

![Сценарий на русском языке: фильтры, выбор населённого пункта, обновление графиков и пирамиды](assets/visuals/scenario-ru.gif)

![Диаграмма архитектуры на русском языке](assets/visuals/architecture-ru.svg)

![Методологическая диаграмма на русском языке](assets/visuals/methodology-ru.svg)

### Публичный дашборд

GitHub Pages: <https://arseniy24rus.github.io/Settlements-projection-dashboard/>

### Аналитическая задача

Дашборд отвечает на практический вопрос: как изучать прогнозируемые демографические изменения на уровне населённых пунктов, а не только крупных административных регионов? Интерфейс ориентирован на сравнение, фильтрацию, картографирование и визуальный анализ динамики численности населения по локальным территориям.

Подробная методология исследования: [docs/methodology.md](docs/methodology.md).

### Структура репозитория

```text
assets/       JavaScript, CSS, иконки и вспомогательные фронтенд-ресурсы
data/         Наборы данных дашборда, включая компактные аналитические файлы
index.html    Главная статическая страница дашборда
README.md     Документация проекта
```

### Модель публикации

Дашборд является статическим сайтом. Его можно публиковать непосредственно из корня репозитория через GitHub Pages. Поскольку браузеры ограничивают доступ к локальным файлам, дашборд следует открывать через HTTP/HTTPS, а не как документ `file://`.

Локальный запуск:

```bash
python -m http.server 8000
```

Затем откройте <http://localhost:8000/>.

### Модель данных

Проект построен вокруг компактных клиентских файлов данных. В зависимости от версии сборки это могут быть Parquet, CSV, JSON или GeoJSON-подобные ресурсы, используемые для таблиц, карт и графиков. Рекомендуемый стандарт документации — поддерживать файл `data_dictionary.md` с описанием идентификаторов населённых пунктов, территориальной иерархии, лет, демографических переменных, сценариев прогноза и этапов предобработки.

### Интерпретация результатов

Дашборд следует использовать как визуальный и разведочный аналитический инструмент. Он подходит для выявления общих пространственных закономерностей, сравнения населённых пунктов и постановки исследовательских вопросов. Его не следует рассматривать как замену полноценного методологического отчёта или официального демографического прогноза.

### Контроль качества

Рекомендуемые проверки включают: загрузку страницы через локальный HTTP-сервер; доступность всех файлов данных; корректность отрисовки карты; реакцию фильтров и графиков на действия пользователя; корректность экспортных функций, если они присутствуют; работоспособность в актуальных настольных браузерах.

### Как цитировать

При использовании дашборда, структуры данных или визуального решения, пожалуйста, цитируйте:

> Ситковский А. М. Settlements Projection Dashboard: spatial population analytics. GitHub, 2026. https://github.com/Arseniy24RUS/Settlements-projection-dashboard

### Лицензия

| Материал | Лицензия | Примечание |
| --- | --- | --- |
| Исходный код | [MIT](LICENSE) | Код приложения, тесты и scripts сборки/запуска. |
| Документация, данные и контент дашборда | [CC BY 4.0](LICENSE-DOCS-AND-DATA.md) | Документация, файлы данных, рисунки, скриншоты и тексты, если явно не указано иное. |
| Сторонние данные, карты, библиотеки, логотипы и институциональные названия | Условия первоначальных поставщиков | См. [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). |
