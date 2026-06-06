# Methodology

## English

### Scope And Research Use

`Settlements-projection-dashboard` is a static analytical dashboard for settlement-level population projection results. The unit of display is an individual settlement with coordinates, administrative attributes, observed census/control-point population values, and one or more projected population trajectories. The unit of demographic component modeling is the municipality, because age-sex structure, fertility, mortality, and migration components are more stable and more commonly available at municipal scale than at the settlement scale.

The dashboard is designed for exploratory research, teaching, publication figures, and policy analysis. It should be cited as an analytical visualization and data-structure project, not as an official demographic forecast. Any paper, report, or policy memo using the outputs should separately document the source statistics, extraction dates, territorial harmonization rules, and scenario assumptions used to create the underlying Parquet files.

### Versioned Analytical Frame

The current dashboard configuration reports 169,909 settlements, 91 regions, 2,367 municipalities, 2,732 central settlements, and 2,344 municipal age-sex profiles. The interface uses census/control years 1989, 2002, 2010, and 2021; settlement projections for 2021-2100; fertility and mortality actual component years 2008-2023; migration actual component years 2010-2023; and migration age-sex projection years 2021-2100.

The runtime build is identified in the application code as `2026-03-02-migage-data-v7`, with migration age data version `20260302-migage-fix1`. A research release should keep these identifiers aligned with `data/config.json`, the published repository commit, and any external data provenance notes.

### Settlement And Municipal Data Model

The dashboard consumes compact Parquet files in `data/` through DuckDB-Wasm in the browser.

| File family | Analytical role |
| --- | --- |
| `settlement_index.parquet` | Core settlement table: `settlement_id`, region, municipality, settlement names, type, coordinates, observed population values, central-place flag text, and municipal join keys such as `oktmo_syn`. |
| `settlement_details.parquet` | Extended attributes shown in the selected-settlement card. |
| `settlement_forecast_wide_withMIG.parquet` and `settlement_forecast_wide_noMIG.parquet` | Settlement-level projected population columns `pop_YEAR` for each scenario. |
| `municipal_age/scenario=SCENARIO/year=YEAR.parquet` | Municipal age-sex population profiles used to construct settlement pyramids. |
| `municipal_components_actual.parquet` | Actual municipal births, deaths, fertility, mortality, and migration components. |
| `municipal_components_wide_withMIG.parquet` and `municipal_components_wide_noMIG.parquet` | Scenario-specific municipal component trajectories in wide format. |
| `municipal_migration_age/scenario=SCENARIO/year=YEAR.parquet` | Age-sex structure of municipal inflow and outflow by five-year age group. |
| `data/config.json` | Dashboard metadata, years, scenario labels, defaults, and headline counts. |

`settlement_id` is the primary settlement identifier used for map selection and row lookup. `Region` and `Municipality` are user-facing spatial filters. `oktmo_syn` links settlements to municipal component and age-sex profile tables. Because administrative names and boundaries change over time, stable research use should treat `settlement_id` and harmonized municipal keys as analytical identifiers, while names remain display labels.

### Population Projection Quantities

The dashboard displays precomputed projection outputs. The demographic identity behind each settlement trajectory can be written as:

```text
P_s,t+1,k = P_s,t,k + B_s,t,k - D_s,t,k + I_s,t,k - O_s,t,k
```

where `P` is population, `B` births, `D` deaths, `I` in-migration, `O` out-migration, `s` settlement, `t` year, and `k` scenario. For the `noMIG` scenario, migration is excluded from the counterfactual trajectory:

```text
P_s,t+1,noMIG = P_s,t,noMIG + B_s,t,noMIG - D_s,t,noMIG
```

Observed/control values are shown for 1989, 2002, 2010, and 2021. For 2021, the dashboard uses the observed settlement value when available and falls back to `pop_2021` from the scenario table. For projection years after 2021, values come from the scenario-specific forecast table.

Aggregate population for a current spatial selection `A` is:

```text
P_A,t,k = sum(P_s,t,k for s in A)
```

Absolute and relative change against a selected base year `b` are:

```text
Delta_abs_s,t,b,k = P_s,t,k - P_s,b,k
Delta_pct_s,t,b,k = 100 * Delta_abs_s,t,b,k / P_s,b,k
```

Relative change is undefined when the base population is missing or zero.

### Scenario Logic

The dashboard supports two scenarios.

| Scenario | Interpretation |
| --- | --- |
| `withMIG` | Projection with fertility, mortality, and migration components. Migration inflow, outflow, net migration, and migration age-sex structure are available where supplied by the municipal component tables. |
| `noMIG` | Counterfactual projection without migration. Population change is driven by natural increase or decrease. Migration age-sex flows are zero by definition and should be interpreted as a contrast to `withMIG`, not as observed migration behavior. |

For research reporting, scenario assumptions should be described outside the dashboard as a structured model protocol: base population, fertility schedule or total fertility rate path, mortality schedule or crude death rate path, migration rates/flows, territorial harmonization, and any constraints or smoothing applied to small settlements.

### Age-Sex Structure

Settlement age-sex pyramids are derived by proportional transfer from the corresponding municipality. For municipality `m`, age `a`, sex `x`, year `t`, and scenario `k`:

```text
share_m,a,x,t,k = N_m,a,x,t,k / sum(N_m,a,x,t,k over all a,x)
N_hat_s,a,x,t,k = round(P_s,t,k * share_m,a,x,t,k)
```

This means the settlement pyramid is not an independently observed settlement-level age-sex distribution. It assumes that the selected settlement has the same age-sex proportions as its municipality in the selected year and scenario. This is usually preferable to suppressing the age-sex view entirely, but it should be treated as an allocation model. The assumption is strongest for small settlements, mono-functional settlements, institutional settlements, and municipalities with strong center-periphery differences.

Migration age-sex pyramids are aggregated from municipal inflow and outflow tables over the current spatial selection:

```text
In_A,a,x,t,k = sum(In_m,a,x,t,k for m linked to selected settlements)
Out_A,a,x,t,k = sum(Out_m,a,x,t,k for m linked to selected settlements)
Net_A,a,x,t,k = In_A,a,x,t,k - Out_A,a,x,t,k
```

### Fertility, Mortality, And Migration Components

Fertility is represented by births and total fertility rate (TFR). For aggregated selections, the dashboard recomputes a weighted TFR using an effective exposure denominator. When direct exposure `expo_eff` is unavailable, it is approximated from births and TFR:

```text
Exposure_fert_m,t = 35 * Births_m,t / TFR_m,t
TFR_A,t = 35 * sum(Births_m,t) / sum(Exposure_fert_m,t)
```

The factor `35` approximates the reproductive age interval. It makes aggregated TFR more interpretable than a simple arithmetic mean of municipal TFR values.

Mortality is represented by deaths and crude death rate (CDR). Aggregated CDR is recomputed from deaths and mid-year population. When `pop_mid` is unavailable, it is approximated from deaths and CDR:

```text
Pop_mid_m,t = 1000 * Deaths_m,t / CDR_m,t
CDR_A,t = 1000 * sum(Deaths_m,t) / sum(Pop_mid_m,t)
```

Migration is represented by inflow, outflow, and net migration:

```text
Mig_net_A,t,k = Mig_in_A,t,k - Mig_out_A,t,k
```

For `withMIG`, these series describe the scenario-specific migration component. For `noMIG`, they should be zero or omitted depending on the source table design.

### Spatial Filters And Map Semantics

The dashboard filters by Russian region and municipality using multiple selection. Municipality options are activated after one or more regions are selected. Internally, municipality filters use a `(Region, Municipality)` pair so that same-name municipalities in different regions remain distinct.

The map displays settlements with finite coordinates and positive population in the selected map year. A zero or missing population in the selected year removes the settlement from the map, which makes potential settlement disappearance visible over time. Symbol size is based on fixed population classes. Symbol color is based on relative change from the selected base year to the selected map year. The map view can be reset to the national extent or fitted to the current filtered selection.

Central settlements are identified through the `Central_places` field and a representative-settlement selection rule that chooses one representative per region/name group, prioritizing larger recent population values. The blue outline and letter marker are visual aids; the original central-place text remains available in the popup and selected-settlement card.

### Data Versioning And Provenance Requirements

Every research-grade refresh should preserve:

- source provider names, source URLs or archive references, extraction dates, and access dates;
- raw file checksums where redistribution is allowed;
- preprocessing scripts or notebooks and their commit hashes;
- territorial harmonization tables for settlement and municipal identifiers;
- scenario labels and assumptions in machine-readable form;
- generated Parquet checksums, row counts, and key uniqueness checks;
- dashboard build identifier, `data/config.json`, and repository commit.

At minimum, a release note or data manifest should state whether rows, identifiers, coordinates, component values, or scenario assumptions changed relative to the previous version.

### Sensitivity And Validation

Recommended sensitivity checks:

- compare `withMIG` and `noMIG` trajectories for each region, municipality, settlement-size class, and central-place group;
- stress-test fertility, mortality, and migration assumptions separately, especially after 2050;
- compare aggregate results against official regional or national control totals when those controls exist;
- examine small settlements separately because rounding and zero-population thresholds have large relative effects;
- test alternative age-sex allocation rules for settlements where municipal proportional transfer is unlikely to be valid;
- verify coordinate outliers, duplicate settlement names, and administrative boundary changes before publication maps are exported.

The dashboard currently presents deterministic trajectories. It does not display probabilistic intervals. If uncertainty intervals are needed, they should be computed upstream and added as separate scenario or percentile datasets.

### Limitations

- The dashboard is an exploratory research instrument, not an official forecast.
- Settlement-level projections are only as strong as the upstream source data, identifier harmonization, and scenario assumptions.
- Settlement age-sex pyramids are municipal proportional allocations, not observed settlement-level age-sex structures.
- Small settlements can show large relative changes from small absolute differences.
- Missing or zero base-year population makes relative change undefined.
- Coordinates, names, and municipal membership can contain source or harmonization errors.
- Administrative boundaries and territorial labels may change over time and should be documented in research outputs.
- External map tiles, libraries, and official data remain subject to third-party terms.

### Reproducible Commands

From the repository root:

```bash
npm install
npm test
```

Local static launch:

```bash
npx http-server . -a 127.0.0.1 -p 4175 -c-1
```

Open <http://127.0.0.1:4175/>. The dashboard should be served over HTTP or HTTPS; opening `index.html` through `file://` will not reliably load local data.

Basic metadata checks on Windows PowerShell:

```powershell
Get-Content -Raw -Encoding UTF8 .\data\config.json | ConvertFrom-Json | Select-Object title, defaultScenario, defaultMapYear
Get-FileHash .\data\config.json
git rev-parse --short HEAD
git status --short
```

---

## Русский

### Область Применения И Исследовательское Использование

`Settlements-projection-dashboard` - статический аналитический дашборд для результатов демографического прогноза на уровне населенных пунктов. Единицей отображения является отдельный населенный пункт с координатами, административными атрибутами, наблюдаемыми контрольными значениями численности и одной или несколькими прогнозными траекториями. Единицей компонентного демографического моделирования является муниципальное образование, поскольку половозрастная структура, рождаемость, смертность и миграция обычно надежнее и доступнее именно на муниципальном уровне.

Дашборд предназначен для разведочного анализа, преподавания, подготовки иллюстраций для публикаций и прикладной аналитики. Его следует цитировать как проект аналитической визуализации и структуры данных, а не как официальный демографический прогноз. Любая статья, записка или доклад, использующие результаты, должны отдельно фиксировать источники статистики, даты выгрузки, правила территориальной гармонизации и сценарные предпосылки, на основе которых были созданы Parquet-файлы.

### Версионированная Аналитическая Рамка

Текущая конфигурация дашборда содержит 169 909 населенных пунктов, 91 регион, 2 367 муниципальных образований, 2 732 опорных населенных пункта и 2 344 муниципальных половозрастных профиля. Интерфейс использует переписные или контрольные годы 1989, 2002, 2010 и 2021; прогнозы по населенным пунктам на 2021-2100 годы; фактические компонентные ряды рождаемости и смертности за 2008-2023 годы; фактические ряды миграции за 2010-2023 годы; половозрастные миграционные профили на 2021-2100 годы.

Сборка приложения обозначена в коде как `2026-03-02-migage-data-v7`, версия миграционных половозрастных данных - `20260302-migage-fix1`. В исследовательском релизе эти идентификаторы должны быть согласованы с `data/config.json`, опубликованным коммитом репозитория и внешними заметками о происхождении данных.

### Модель Данных Населенных Пунктов И Муниципалитетов

Дашборд читает компактные Parquet-файлы из `data/` через DuckDB-Wasm непосредственно в браузере.

| Семейство файлов | Аналитическая роль |
| --- | --- |
| `settlement_index.parquet` | Основная таблица населенных пунктов: `settlement_id`, регион, муниципалитет, названия, тип, координаты, наблюдаемая численность, текст статуса опорного пункта и муниципальные ключи вроде `oktmo_syn`. |
| `settlement_details.parquet` | Расширенные атрибуты, показываемые в карточке выбранного населенного пункта. |
| `settlement_forecast_wide_withMIG.parquet` и `settlement_forecast_wide_noMIG.parquet` | Прогнозные столбцы `pop_YEAR` по населенным пунктам для каждого сценария. |
| `municipal_age/scenario=SCENARIO/year=YEAR.parquet` | Муниципальные половозрастные профили для построения пирамид населенных пунктов. |
| `municipal_components_actual.parquet` | Фактические муниципальные компоненты рождаемости, смертности и миграции. |
| `municipal_components_wide_withMIG.parquet` и `municipal_components_wide_noMIG.parquet` | Сценарные муниципальные компонентные траектории в широком формате. |
| `municipal_migration_age/scenario=SCENARIO/year=YEAR.parquet` | Половозрастная структура муниципального притока и оттока по пятилетним возрастным группам. |
| `data/config.json` | Метаданные дашборда, годы, подписи сценариев, значения по умолчанию и основные счетчики. |

`settlement_id` является основным идентификатором населенного пункта для выбора на карте и поиска строки. `Region` и `Municipality` используются как пользовательские пространственные фильтры. `oktmo_syn` связывает населенные пункты с муниципальными компонентами и половозрастными профилями. Из-за изменений административных названий и границ устойчивый исследовательский анализ должен считать аналитическими идентификаторами `settlement_id` и гармонизированные муниципальные ключи, а названия - отображаемыми подписями.

### Величины Прогноза Численности

Дашборд показывает заранее рассчитанные прогнозные результаты. Демографическое тождество, лежащее за траекторией населенного пункта, можно записать так:

```text
P_s,t+1,k = P_s,t,k + B_s,t,k - D_s,t,k + I_s,t,k - O_s,t,k
```

где `P` - численность населения, `B` - рождения, `D` - смерти, `I` - прибытия, `O` - выбытия, `s` - населенный пункт, `t` - год, `k` - сценарий. Для сценария `noMIG` миграция исключается из контрфактической траектории:

```text
P_s,t+1,noMIG = P_s,t,noMIG + B_s,t,noMIG - D_s,t,noMIG
```

Наблюдаемые или контрольные значения показываются для 1989, 2002, 2010 и 2021 годов. Для 2021 года дашборд использует наблюдаемое значение по населенному пункту, если оно доступно, и значение `pop_2021` из сценарной таблицы в качестве резерва. Для прогнозных лет после 2021 года значения берутся из сценарной прогнозной таблицы.

Суммарная численность по текущей пространственной выборке `A`:

```text
P_A,t,k = sum(P_s,t,k for s in A)
```

Абсолютное и относительное изменение к выбранному базовому году `b`:

```text
Delta_abs_s,t,b,k = P_s,t,k - P_s,b,k
Delta_pct_s,t,b,k = 100 * Delta_abs_s,t,b,k / P_s,b,k
```

Относительное изменение не определяется, если базовая численность отсутствует или равна нулю.

### Логика Сценариев

Дашборд поддерживает два сценария.

| Сценарий | Интерпретация |
| --- | --- |
| `withMIG` | Прогноз с компонентами рождаемости, смертности и миграции. При наличии муниципальных таблиц доступны приток, отток, сальдо и половозрастная структура миграции. |
| `noMIG` | Контрфактический прогноз без миграции. Изменение численности определяется естественным приростом или убылью. Половозрастные миграционные потоки равны нулю по определению и должны использоваться как контраст к `withMIG`, а не как описание наблюдаемой миграции. |

Для исследовательского отчета сценарные предпосылки следует описывать вне дашборда как структурированный протокол модели: базовая численность, траектория возрастных коэффициентов рождаемости или СКР, траектория смертности или общего коэффициента смертности, миграционные rates/flows, территориальная гармонизация, ограничения и сглаживание малых населенных пунктов.

### Половозрастная Структура

Половозрастные пирамиды населенных пунктов строятся через пропорциональный перенос профиля соответствующего муниципального образования. Для муниципалитета `m`, возраста `a`, пола `x`, года `t` и сценария `k`:

```text
share_m,a,x,t,k = N_m,a,x,t,k / sum(N_m,a,x,t,k over all a,x)
N_hat_s,a,x,t,k = round(P_s,t,k * share_m,a,x,t,k)
```

Следовательно, пирамида населенного пункта не является независимо наблюдаемой половозрастной структурой. Она предполагает, что выбранный населенный пункт имеет те же половозрастные доли, что и муниципалитет в выбранном году и сценарии. Такой подход лучше, чем полное отсутствие половозрастного представления, но его нужно трактовать как аллокационную модель. Предположение особенно чувствительно для малых населенных пунктов, монофункциональных поселений, институциональных населенных пунктов и муниципалитетов с сильным различием между центром и периферией.

Миграционные половозрастные пирамиды агрегируются из муниципальных таблиц притока и оттока по текущей пространственной выборке:

```text
In_A,a,x,t,k = sum(In_m,a,x,t,k for m linked to selected settlements)
Out_A,a,x,t,k = sum(Out_m,a,x,t,k for m linked to selected settlements)
Net_A,a,x,t,k = In_A,a,x,t,k - Out_A,a,x,t,k
```

### Рождаемость, Смертность И Миграционные Компоненты

Рождаемость представлена числом рождений и суммарным коэффициентом рождаемости (СКР). Для агрегированных выборок дашборд пересчитывает взвешенный СКР через эффективный знаменатель экспозиции. Если прямое поле `expo_eff` отсутствует, знаменатель приближается по числу рождений и СКР:

```text
Exposure_fert_m,t = 35 * Births_m,t / TFR_m,t
TFR_A,t = 35 * sum(Births_m,t) / sum(Exposure_fert_m,t)
```

Коэффициент `35` приближает длину репродуктивного возрастного интервала. Это делает агрегированный СКР содержательнее простой арифметической средней муниципальных значений.

Смертность представлена числом умерших и общим коэффициентом смертности (ОКС). Агрегированный ОКС пересчитывается через число умерших и среднегодовую численность. Если `pop_mid` отсутствует, она приближается по числу умерших и ОКС:

```text
Pop_mid_m,t = 1000 * Deaths_m,t / CDR_m,t
CDR_A,t = 1000 * sum(Deaths_m,t) / sum(Pop_mid_m,t)
```

Миграция представлена притоком, оттоком и сальдо:

```text
Mig_net_A,t,k = Mig_in_A,t,k - Mig_out_A,t,k
```

Для `withMIG` эти ряды описывают сценарную миграционную компоненту. Для `noMIG` они должны быть нулевыми или отсутствовать в зависимости от устройства исходной таблицы.

### Пространственные Фильтры И Семантика Карты

Дашборд фильтрует данные по субъектам РФ и муниципальным образованиям с поддержкой множественного выбора. Список муниципалитетов активируется после выбора одного или нескольких регионов. Внутри приложения муниципальный фильтр использует пару `(Region, Municipality)`, чтобы одноименные муниципалитеты в разных регионах оставались различимыми.

На карте показываются населенные пункты с корректными координатами и положительной численностью в выбранном году карты. Нулевая или отсутствующая численность в выбранном году исключает пункт с карты, что делает возможное исчезновение поселений видимым в динамике. Размер символа основан на фиксированных классах численности. Цвет символа основан на относительном изменении от выбранного базового года к году карты. Вид карты можно сбросить к общероссийскому охвату или масштабировать к текущей выборке.

Опорные населенные пункты определяются через поле `Central_places` и правило выбора представителя: один представитель на группу регион/название, с приоритетом большей недавней численности. Синяя окантовка и буквенный маркер являются визуальными подсказками; исходный текст статуса сохраняется во всплывающей подсказке и карточке выбранного пункта.

### Версионирование Данных И Требования К Provenance

Каждое исследовательское обновление должно сохранять:

- названия поставщиков источников, ссылки или архивные указания, даты выгрузки и даты доступа;
- контрольные суммы сырых файлов, если их можно распространять;
- скрипты или notebooks предобработки и их commit hash;
- таблицы территориальной гармонизации идентификаторов населенных пунктов и муниципалитетов;
- сценарные подписи и предпосылки в машинно-читаемом виде;
- контрольные суммы Parquet-файлов, счетчики строк и проверки уникальности ключей;
- идентификатор сборки дашборда, `data/config.json` и коммит репозитория.

Минимально релизная заметка или data manifest должны указывать, изменились ли строки, идентификаторы, координаты, компонентные значения или сценарные предпосылки по сравнению с предыдущей версией.

### Чувствительность И Валидация

Рекомендуемые проверки чувствительности:

- сравнивать траектории `withMIG` и `noMIG` по регионам, муниципалитетам, классам размера населенных пунктов и группам опорных пунктов;
- отдельно стресс-тестировать предпосылки рождаемости, смертности и миграции, особенно после 2050 года;
- сравнивать агрегированные результаты с официальными региональными или национальными контрольными суммами, если они доступны;
- анализировать малые населенные пункты отдельно, поскольку округление и пороги нулевой численности дают большие относительные эффекты;
- проверять альтернативные правила аллокации половозрастной структуры там, где муниципальный пропорциональный перенос маловероятен;
- проверять координатные выбросы, дубли названий и административные изменения перед экспортом карт для публикации.

В текущем виде дашборд показывает детерминированные траектории. Вероятностные интервалы не отображаются. Если они нужны, их следует рассчитывать upstream и добавлять как отдельные сценарные или percentile-наборы данных.

### Ограничения

- Дашборд является разведочным исследовательским инструментом, а не официальным прогнозом.
- Прогнозы на уровне населенных пунктов зависят от качества исходных данных, гармонизации идентификаторов и сценарных предпосылок.
- Половозрастные пирамиды населенных пунктов являются муниципальными пропорциональными аллокациями, а не наблюдаемыми структурами.
- В малых населенных пунктах небольшие абсолютные изменения дают большие относительные изменения.
- Если базовая численность отсутствует или равна нулю, относительное изменение не определяется.
- Координаты, названия и муниципальная принадлежность могут содержать ошибки источника или гармонизации.
- Административные границы и территориальные названия могут меняться во времени и должны описываться в исследовательских материалах.
- Внешние картографические тайлы, библиотеки и официальные данные остаются под условиями третьих сторон.

### Воспроизводимые Команды

Из корня репозитория:

```bash
npm install
npm test
```

Локальный статический запуск:

```bash
npx http-server . -a 127.0.0.1 -p 4175 -c-1
```

Откройте <http://127.0.0.1:4175/>. Дашборд следует запускать через HTTP или HTTPS; открытие `index.html` через `file://` не обеспечивает надежную загрузку локальных данных.

Базовые проверки метаданных в Windows PowerShell:

```powershell
Get-Content -Raw -Encoding UTF8 .\data\config.json | ConvertFrom-Json | Select-Object title, defaultScenario, defaultMapYear
Get-FileHash .\data\config.json
git rev-parse --short HEAD
git status --short
```
