import * as duckdb from 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@1.33.1-dev18.0/+esm';

const DASHBOARD_BUILD = '2026-03-01-centralplaces-v5';
window.__RUSSIA_SETTLEMENTS_DASHBOARD_BUILD__ = DASHBOARD_BUILD;

const DEFAULT_MAP_VIEW = { center: [95, 63], zoom: 2.35 };

const POPULATION_CLASSES = [
  { min: 1, max: 100, label: '1–100', shortLabel: '1–100', radius: 4 },
  { min: 101, max: 500, label: '101–500', shortLabel: '101–500', radius: 5 },
  { min: 501, max: 1000, label: '501–1 000', shortLabel: '501–1 тыс.', radius: 6 },
  { min: 1001, max: 5000, label: '1 001–5 000', shortLabel: '1–5 тыс.', radius: 8 },
  { min: 5001, max: 10000, label: '5 001–10 000', shortLabel: '5–10 тыс.', radius: 10 },
  { min: 10001, max: 50000, label: '10 001–50 000', shortLabel: '10–50 тыс.', radius: 12 },
  { min: 50001, max: 100000, label: '50 001–100 000', shortLabel: '50–100 тыс.', radius: 14 },
  { min: 100001, max: 500000, label: '100 001–500 000', shortLabel: '100–500 тыс.', radius: 16 },
  { min: 500001, max: 1000000, label: '500 001–1 000 000', shortLabel: '500 тыс.–1 млн', radius: 18 },
  { min: 1000001, max: 5000000, label: '1 000 001–5 000 000', shortLabel: '1–5 млн', radius: 22 },
  { min: 5000001, max: Infinity, label: 'Более 5 000 001', shortLabel: '> 5 млн', radius: 26 },
];

const SCENARIO_TRACE_STYLES = {
  withMIG: { color: '#15803d', dash: 'solid' },
  noMIG: { color: '#b45309', dash: 'dash' },
};

const CENTRAL_PLACE_GROUPS = [
  { id: 'services', symbol: 'С', label: 'центр социальных услуг' },
  { id: 'investment', symbol: 'И', label: 'инвестиционный центр' },
  { id: 'aggl_core', symbol: 'Я', label: 'ядро агломерации' },
  { id: 'agglomeration', symbol: 'А', label: 'агломерационный пункт' },
  { id: 'security', symbol: 'Б', label: 'безопасность / граница' },
  { id: 'infrastructure', symbol: 'К', label: 'критическая инфраструктура' },
  { id: 'zato', symbol: 'З', label: 'ЗАТО' },
  { id: 'science', symbol: 'Н', label: 'наукоград' },
  { id: 'regional_capital', symbol: 'Р', label: 'центр субъекта РФ' },
  { id: 'other', symbol: 'О', label: 'иной тип опорного пункта' },
];

const CENTRAL_GROUP_BY_ID = new Map(CENTRAL_PLACE_GROUPS.map((item) => [item.id, item]));

const els = {
  scenarioButtons: document.getElementById('scenarioButtons'),
  chartRangeSlider: document.getElementById('chartRangeSlider'),
  baseYearDisplay: document.getElementById('baseYearDisplay'),
  chartSpanDisplay: document.getElementById('chartSpanDisplay'),
  chartEndDisplay: document.getElementById('chartEndDisplay'),
  mapYearSlider: document.getElementById('mapYearSlider'),
  mapYearDisplay: document.getElementById('mapYearDisplay'),
  pyramidYearSlider: document.getElementById('pyramidYearSlider'),
  pyramidYearDisplay: document.getElementById('pyramidYearDisplay'),
  migrationYearSlider: document.getElementById('migrationYearSlider'),
  migrationYearDisplay: document.getElementById('migrationYearDisplay'),
  migrationScenarioNote: document.getElementById('migrationScenarioNote'),
  regionSelect: document.getElementById('regionSelect'),
  municipalitySelect: document.getElementById('municipalitySelect'),
  municipalityHint: document.getElementById('municipalityHint'),
  centralPlacesToggle: document.getElementById('centralPlacesToggle'),
  autoPyramidToggle: document.getElementById('autoPyramidToggle'),
  deltaNegativeInput: document.getElementById('deltaNegativeInput'),
  deltaPositiveInput: document.getElementById('deltaPositiveInput'),
  applyFiltersBtn: document.getElementById('applyFiltersBtn'),
  clearFiltersBtn: document.getElementById('clearFiltersBtn'),
  exportSelectionXlsxBtn: document.getElementById('exportSelectionXlsxBtn'),
  exportMapPngBtn: document.getElementById('exportMapPngBtn'),
  exportMapXlsxBtn: document.getElementById('exportMapXlsxBtn'),
  exportAggregatePngBtn: document.getElementById('exportAggregatePngBtn'),
  exportAggregateXlsxBtn: document.getElementById('exportAggregateXlsxBtn'),
  exportSettlementPngBtn: document.getElementById('exportSettlementPngBtn'),
  exportSettlementXlsxBtn: document.getElementById('exportSettlementXlsxBtn'),
  exportPyramidPngBtn: document.getElementById('exportPyramidPngBtn'),
  exportPyramidXlsxBtn: document.getElementById('exportPyramidXlsxBtn'),
  exportSelectedXlsxBtn: document.getElementById('exportSelectedXlsxBtn'),
  exportFertilityPngBtn: document.getElementById('exportFertilityPngBtn'),
  exportFertilityXlsxBtn: document.getElementById('exportFertilityXlsxBtn'),
  exportMortalityPngBtn: document.getElementById('exportMortalityPngBtn'),
  exportMortalityXlsxBtn: document.getElementById('exportMortalityXlsxBtn'),
  exportMigrationPngBtn: document.getElementById('exportMigrationPngBtn'),
  exportMigrationXlsxBtn: document.getElementById('exportMigrationXlsxBtn'),
  aggregateScenarioWithMIG: document.getElementById('aggregateScenarioWithMIG'),
  aggregateScenarioNoMIG: document.getElementById('aggregateScenarioNoMIG'),
  settlementScenarioWithMIG: document.getElementById('settlementScenarioWithMIG'),
  settlementScenarioNoMIG: document.getElementById('settlementScenarioNoMIG'),
  fertilityScenarioWithMIG: document.getElementById('fertilityScenarioWithMIG'),
  fertilityScenarioNoMIG: document.getElementById('fertilityScenarioNoMIG'),
  mortalityScenarioWithMIG: document.getElementById('mortalityScenarioWithMIG'),
  mortalityScenarioNoMIG: document.getElementById('mortalityScenarioNoMIG'),
  migrationScenarioWithMIG: document.getElementById('migrationScenarioWithMIG'),
  migrationScenarioNoMIG: document.getElementById('migrationScenarioNoMIG'),
  openMethodologyBtn: document.getElementById('openMethodologyBtn'),
  resetMapViewBtn: document.getElementById('resetMapViewBtn'),
  loadingOverlay: document.getElementById('loadingOverlay'),
  loadingTitle: document.getElementById('loadingTitle'),
  loadingSubtitle: document.getElementById('loadingSubtitle'),
  errorBanner: document.getElementById('errorBanner'),
  mapCaption: document.getElementById('mapCaption'),
  mapColorLegendTitle: document.getElementById('mapColorLegendTitle'),
  mapColorLegendLabels: document.getElementById('mapColorLegendLabels'),
  sizeLegend: document.getElementById('sizeLegend'),
  centralTypeLegend: document.getElementById('centralTypeLegend'),
  map: document.getElementById('map'),
  kpiSettlements: document.getElementById('kpiSettlements'),
  kpiFilterNote: document.getElementById('kpiFilterNote'),
  kpiPopulation: document.getElementById('kpiPopulation'),
  kpiPopulationNote: document.getElementById('kpiPopulationNote'),
  kpiDelta: document.getElementById('kpiDelta'),
  kpiDeltaNote: document.getElementById('kpiDeltaNote'),
  kpiCentral: document.getElementById('kpiCentral'),
  kpiCentralNote: document.getElementById('kpiCentralNote'),
  selectedSettlementSummary: document.getElementById('selectedSettlementSummary'),
  selectedSettlementMeta: document.getElementById('selectedSettlementMeta'),
  selectionHint: document.getElementById('selectionHint'),
  topSettlements: document.getElementById('topSettlements'),
  fertilityBlock: document.getElementById('fertilityBlock'),
  mortalityBlock: document.getElementById('mortalityBlock'),
  migrationBlock: document.getElementById('migrationBlock'),
  methodologyDialog: document.getElementById('methodologyDialog'),
};

const state = {
  config: null,
  db: null,
  conn: null,
  map: null,
  deckOverlay: null,
  selectedSettlementId: null,
  currentMapData: [],
  currentAggregateSeries: null,
  currentSelectionBundle: null,
  currentPyramidBundle: null,
  currentComponentSeries: null,
  currentMigrationAgeBundle: null,
  registeredFiles: new Set(),
  filterOptions: {
    regions: [],
    municipalities: [],
    municipalitiesByRegion: new Map(),
  },
  centralPlaceRepresentatives: [],
  centralPlaceRepresentativeIds: new Set(),
  centralPlaceTextById: new Map(),
  ui: {
    activeScenario: null,
    regionTom: null,
    municipalityTom: null,
    hoverPopup: null,
    refreshTimer: null,
    refreshToken: 0,
    mapReady: null,
    pendingFitBounds: false,
  },
};



document.addEventListener('DOMContentLoaded', () => {
  init().catch((error) => {
    console.error(error);
    showError(`Не удалось инициализировать дашборд: ${error?.message || error}`);
    setLoading(false);
  });
});

async function init() {
  setLoading(true, 'Инициализация дашборда…', 'Загружается конфигурация и подключается движок DuckDB-Wasm.');
  await loadConfig();
  initControls();
  bindEvents();
  await initDuckDB();
  await prepareTables();
  await loadCentralPlaceRepresentatives();
  await loadFilterOptions();
  initMap();
  await state.ui.mapReady;
  await refreshAll();
  setLoading(false);
}

async function loadConfig() {
  const response = await fetch('./data/config.json');
  if (!response.ok) {
    throw new Error(`Не удалось загрузить конфигурацию (HTTP ${response.status}).`);
  }
  state.config = await response.json();
  state.config.allYears = [...new Set([...state.config.censusYears, ...state.config.forecastYears])].sort((a, b) => a - b);
  state.config.defaultChartEndYear = Math.max(...state.config.forecastYears);
  state.config.allYearIndex = Object.fromEntries(state.config.allYears.map((year, index) => [year, index]));
  state.config.mapYearIndex = Object.fromEntries(state.config.mapYears.map((year, index) => [year, index]));
  state.config.pyramidYearIndex = Object.fromEntries(state.config.pyramidYears.map((year, index) => [year, index]));
  state.config.migrationYearIndex = Object.fromEntries((state.config.componentYears?.migrationPyramid || []).map((year, index) => [year, index]));
}



function initControls() {
  state.ui.activeScenario = state.config.defaultScenario;
  renderScenarioButtons();

  createRangeYearSlider(
    els.chartRangeSlider,
    state.config.allYears,
    [state.config.allYearIndex[state.config.defaultBaseYear], state.config.allYearIndex[state.config.defaultChartEndYear]],
    () => updateTimelineReadout(),
    () => scheduleRefresh(),
  );
  createSingleYearSlider(
    els.mapYearSlider,
    state.config.mapYears,
    state.config.mapYearIndex[state.config.defaultMapYear],
    () => updateMapYearReadout(),
    () => {
      syncPyramidYearIfNeeded();
      scheduleRefresh();
    },
  );
  createSingleYearSlider(
    els.pyramidYearSlider,
    state.config.pyramidYears,
    state.config.pyramidYearIndex[Math.max(2021, state.config.defaultBaseYear)],
    () => updatePyramidYearReadout(),
    () => renderPyramid().catch(handleError),
  );
  createSingleYearSlider(
    els.migrationYearSlider,
    state.config.componentYears.migrationPyramid,
    state.config.migrationYearIndex[state.config.defaultMigrationPyramidYear] ?? 0,
    () => updateMigrationYearReadout(),
    () => renderMigrationAgePyramids().catch(handleError),
  );

  updateTimelineReadout();
  updateMapYearReadout();
  updatePyramidYearReadout();
  updateMigrationYearReadout();

  [
    [els.aggregateScenarioWithMIG, els.aggregateScenarioNoMIG],
    [els.settlementScenarioWithMIG, els.settlementScenarioNoMIG],
    [els.fertilityScenarioWithMIG, els.fertilityScenarioNoMIG],
    [els.mortalityScenarioWithMIG, els.mortalityScenarioNoMIG],
    [els.migrationScenarioWithMIG, els.migrationScenarioNoMIG],
  ].forEach(([withMigEl, noMigEl]) => {
    withMigEl.checked = state.ui.activeScenario === 'withMIG';
    noMigEl.checked = state.ui.activeScenario === 'noMIG';
  });

  state.ui.regionTom = new TomSelect(els.regionSelect, {
    plugins: ['remove_button'],
    maxItems: null,
    maxOptions: 5000,
    create: false,
    persist: false,
    closeAfterSelect: false,
    placeholder: 'Все субъекты РФ',
  });

  state.ui.municipalityTom = new TomSelect(els.municipalitySelect, {
    plugins: ['remove_button'],
    maxItems: null,
    maxOptions: 50000,
    create: false,
    persist: false,
    closeAfterSelect: false,
    placeholder: 'Сначала выберите субъект РФ',
  });
  state.ui.municipalityTom.disable();
  state.ui.municipalityTom.inputState();

  renderSizeLegend();
  renderCentralTypeLegend();
  renderMapLegend();
  updateMigrationScenarioNote();
}



function bindEvents() {
  els.scenarioButtons.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-scenario]');
    if (!button) return;
    setActiveScenario(button.dataset.scenario, { syncCharts: true, refresh: true });
  });

  state.ui.regionTom.on('change', () => updateMunicipalityOptions());
  state.ui.municipalityTom.on('change', () => {});

  els.applyFiltersBtn.addEventListener('click', () => {
    state.ui.pendingFitBounds = true;
    scheduleRefresh();
  });
  els.clearFiltersBtn.addEventListener('click', () => clearFilters(true));
  els.resetMapViewBtn.addEventListener('click', () => resetMapView());
  els.openMethodologyBtn.addEventListener('click', () => els.methodologyDialog.showModal());

  els.centralPlacesToggle.addEventListener('change', () => updateMapLayers());
  els.autoPyramidToggle.addEventListener('change', () => {
    if (els.autoPyramidToggle.checked) {
      syncPyramidYearIfNeeded(true);
    }
  });

  const rerenderMapAppearance = () => {
    renderMapCaption();
    updateMapLayers();
  };
  els.deltaNegativeInput.addEventListener('change', rerenderMapAppearance);
  els.deltaPositiveInput.addEventListener('change', rerenderMapAppearance);

  bindScenarioCheckboxGroup(els.aggregateScenarioWithMIG, els.aggregateScenarioNoMIG, () => renderAggregateChart());
  bindScenarioCheckboxGroup(els.settlementScenarioWithMIG, els.settlementScenarioNoMIG, () => renderSelectedSettlementChart());
  bindScenarioCheckboxGroup(els.fertilityScenarioWithMIG, els.fertilityScenarioNoMIG, () => renderFertilityBlock());
  bindScenarioCheckboxGroup(els.mortalityScenarioWithMIG, els.mortalityScenarioNoMIG, () => renderMortalityBlock());
  bindScenarioCheckboxGroup(els.migrationScenarioWithMIG, els.migrationScenarioNoMIG, () => renderMigrationBlock());

  els.exportSelectionXlsxBtn.addEventListener('click', () => exportCurrentSelectionXlsx());
  els.exportMapXlsxBtn.addEventListener('click', () => exportCurrentSelectionXlsx());
  els.exportMapPngBtn.addEventListener('click', () => exportMapPng());
  els.exportAggregateXlsxBtn.addEventListener('click', () => exportAggregateXlsx());
  els.exportAggregatePngBtn.addEventListener('click', () => exportPlotPng('aggregateChart', buildFilename('aggregate')));
  els.exportSettlementXlsxBtn.addEventListener('click', () => exportSettlementSeriesXlsx());
  els.exportSettlementPngBtn.addEventListener('click', () => exportPlotPng('settlementChart', buildFilename('settlement_series')));
  els.exportPyramidXlsxBtn.addEventListener('click', () => exportPyramidXlsx());
  els.exportPyramidPngBtn.addEventListener('click', () => exportPlotPng('pyramidChart', buildFilename('pyramid')));
  els.exportSelectedXlsxBtn.addEventListener('click', () => exportSelectedDetailsXlsx());
  els.exportFertilityPngBtn.addEventListener('click', () => exportElementPng('fertilityBlock', buildFilename('fertility')));
  els.exportFertilityXlsxBtn.addEventListener('click', () => exportFertilityXlsx());
  els.exportMortalityPngBtn.addEventListener('click', () => exportElementPng('mortalityBlock', buildFilename('mortality')));
  els.exportMortalityXlsxBtn.addEventListener('click', () => exportMortalityXlsx());
  els.exportMigrationPngBtn.addEventListener('click', () => exportElementPng('migrationBlock', buildFilename('migration')));
  els.exportMigrationXlsxBtn.addEventListener('click', () => exportMigrationXlsx());
}



function bindScenarioCheckboxGroup(firstEl, secondEl, onChange) {
  const handler = () => {
    if (!firstEl.checked && !secondEl.checked) {
      const fallback = getScenario();
      if (firstEl.dataset.scenario === fallback) {
        firstEl.checked = true;
      } else if (secondEl.dataset.scenario === fallback) {
        secondEl.checked = true;
      } else {
        firstEl.checked = true;
      }
    }
    onChange();
  };
  firstEl.addEventListener('change', handler);
  secondEl.addEventListener('change', handler);
}

function createRangeYearSlider(element, years, startIndices, onUpdate, onChange) {
  noUiSlider.create(element, {
    start: startIndices,
    connect: true,
    step: 1,
    range: { min: 0, max: years.length - 1 },
    behaviour: 'tap-drag',
    tooltips: [makeSliderTooltip(years), makeSliderTooltip(years)],
  });
  element.noUiSlider.on('update', onUpdate);
  element.noUiSlider.on('change', onChange);
}

function createSingleYearSlider(element, years, startIndex, onUpdate, onChange) {
  noUiSlider.create(element, {
    start: [startIndex],
    connect: [true, false],
    step: 1,
    range: { min: 0, max: years.length - 1 },
    behaviour: 'tap',
    tooltips: [makeSliderTooltip(years)],
  });
  element.noUiSlider.on('update', onUpdate);
  element.noUiSlider.on('change', onChange);
}

function makeSliderTooltip(years) {
  return {
    to(value) {
      return String(years[Math.round(Number(value))]);
    },
    from(value) {
      return Number(value);
    },
  };
}

function setActiveScenario(scenarioId, { syncCharts = true, refresh = true } = {}) {
  const isValid = state.config.scenarios.some((scenario) => scenario.id === scenarioId);
  if (!isValid) return;
  state.ui.activeScenario = scenarioId;
  renderScenarioButtons();
  updateMigrationScenarioNote();

  if (syncCharts) {
    [
      [els.aggregateScenarioWithMIG, els.aggregateScenarioNoMIG],
      [els.settlementScenarioWithMIG, els.settlementScenarioNoMIG],
      [els.fertilityScenarioWithMIG, els.fertilityScenarioNoMIG],
      [els.mortalityScenarioWithMIG, els.mortalityScenarioNoMIG],
      [els.migrationScenarioWithMIG, els.migrationScenarioNoMIG],
    ].forEach(([firstEl, secondEl]) => syncScenarioCheckboxGroupToActive(firstEl, secondEl, scenarioId));
  }

  if (refresh) {
    scheduleRefresh();
  }
}



function syncScenarioCheckboxGroupToActive(firstEl, secondEl, activeScenarioId) {
  const checkedCount = [firstEl, secondEl].filter((input) => input.checked).length;
  if (checkedCount <= 1) {
    firstEl.checked = firstEl.dataset.scenario === activeScenarioId;
    secondEl.checked = secondEl.dataset.scenario === activeScenarioId;
  } else {
    if (firstEl.dataset.scenario === activeScenarioId) firstEl.checked = true;
    if (secondEl.dataset.scenario === activeScenarioId) secondEl.checked = true;
  }
}

function renderScenarioButtons() {
  els.scenarioButtons.innerHTML = state.config.scenarios.map((scenario) => `
    <button type="button" data-scenario="${escapeHtml(scenario.id)}" class="${scenario.id === state.ui.activeScenario ? 'active' : ''}">${escapeHtml(scenario.label)}</button>
  `).join('');
}

function updateTimelineReadout() {
  const [baseYear, chartEndYear] = getChartRangeYears();
  els.baseYearDisplay.textContent = formatAny(baseYear);
  els.chartEndDisplay.textContent = formatAny(chartEndYear);
  const span = Math.max(0, chartEndYear - baseYear);
  els.chartSpanDisplay.textContent = `${span.toLocaleString('ru-RU')} ${pluralYears(span)}`;
}

function updateMapYearReadout() {
  els.mapYearDisplay.textContent = formatAny(getMapYear());
}

function updatePyramidYearReadout() {
  els.pyramidYearDisplay.textContent = formatAny(getPyramidYear());
}

function updateMigrationYearReadout() {
  els.migrationYearDisplay.textContent = formatAny(getMigrationPyramidYear());
}

function updateMigrationScenarioNote() {
  if (!els.migrationScenarioNote) return;
  els.migrationScenarioNote.textContent = `Активный сценарий: ${getScenarioLabel(getScenario())}`;
}

async function initDuckDB() {
  setLoading(true, 'Подключение вычислительного движка…', 'Подготавливается DuckDB-Wasm для выполнения запросов непосредственно в браузере.');
  const bundles = duckdb.getJsDelivrBundles();
  const bundle = await duckdb.selectBundle(bundles);
  const workerUrl = URL.createObjectURL(
    new Blob([`importScripts("${bundle.mainWorker}");`], { type: 'text/javascript' }),
  );
  const worker = new Worker(workerUrl);
  const logger = new duckdb.ConsoleLogger();
  state.db = new duckdb.AsyncDuckDB(logger, worker);
  await state.db.instantiate(bundle.mainModule, bundle.pthreadWorker);
  URL.revokeObjectURL(workerUrl);
  state.conn = await state.db.connect();
}

async function prepareTables() {
  setLoading(true, 'Регистрация файлов данных…', 'Подключаются локальные Parquet-файлы, предназначенные для работы на GitHub Pages.');
  await ensureRegisteredFile('settlement_index.parquet', './data/settlement_index.parquet');
  await ensureRegisteredFile('settlement_details.parquet', './data/settlement_details.parquet');
  await ensureRegisteredFile('settlement_forecast_wide_withMIG.parquet', './data/settlement_forecast_wide_withMIG.parquet');
  await ensureRegisteredFile('settlement_forecast_wide_noMIG.parquet', './data/settlement_forecast_wide_noMIG.parquet');
  await ensureRegisteredFile('municipal_components_actual.parquet', './data/municipal_components_actual.parquet');
  await ensureRegisteredFile('municipal_components_wide_withMIG.parquet', './data/municipal_components_wide_withMIG.parquet');
  await ensureRegisteredFile('municipal_components_wide_noMIG.parquet', './data/municipal_components_wide_noMIG.parquet');
  await queryExec(`CREATE OR REPLACE TABLE settlement_index AS SELECT * FROM read_parquet('settlement_index.parquet')`);
  await queryExec(`
    CREATE OR REPLACE TABLE central_place_representatives AS
    WITH flagged AS (
      SELECT
        CAST(settlement_id AS BIGINT) AS settlement_id,
        Region,
        Municipality,
        COALESCE(
          NULLIF(TRIM(CAST(Settlement_name AS VARCHAR)), ''),
          NULLIF(TRIM(CAST(Settlement_full_name AS VARCHAR)), '')
        ) AS settlement_name_norm,
        CAST(Central_places AS VARCHAR) AS Central_places,
        CAST(Population_2021 AS DOUBLE) AS pop_2021,
        CAST(Population_2010 AS DOUBLE) AS pop_2010,
        CAST(Population_2002 AS DOUBLE) AS pop_2002,
        CAST(Population_1989 AS DOUBLE) AS pop_1989
      FROM settlement_index
      WHERE Central_places IS NOT NULL
        AND TRIM(CAST(Central_places AS VARCHAR)) <> ''
        AND LOWER(TRIM(CAST(Central_places AS VARCHAR))) <> 'nan'
    ),
    ranked AS (
      SELECT *,
        ROW_NUMBER() OVER (
          PARTITION BY Region, settlement_name_norm
          ORDER BY pop_2021 DESC NULLS LAST, pop_2010 DESC NULLS LAST, pop_2002 DESC NULLS LAST, pop_1989 DESC NULLS LAST, Municipality ASC, settlement_id ASC
        ) AS rn
      FROM flagged
      WHERE settlement_name_norm IS NOT NULL
    )
    SELECT
      settlement_id,
      Region,
      Municipality,
      settlement_name_norm,
      Central_places
    FROM ranked
    WHERE rn = 1
  `);
}



async function loadCentralPlaceRepresentatives() {
  const rows = await queryRows(`
    SELECT
      CAST(settlement_id AS BIGINT) AS settlement_id,
      Region,
      Municipality,
      settlement_name_norm,
      Central_places
    FROM central_place_representatives
    ORDER BY Region, Municipality, settlement_name_norm
  `);

  state.centralPlaceRepresentatives = rows.map((row) => ({
    settlement_id: toNumber(row.settlement_id),
    Region: row.Region,
    Municipality: row.Municipality,
    settlement_name_norm: row.settlement_name_norm,
    Central_places: row.Central_places,
  }));
  state.centralPlaceRepresentativeIds = new Set(state.centralPlaceRepresentatives.map((row) => row.settlement_id));
  state.centralPlaceTextById = new Map(state.centralPlaceRepresentatives.map((row) => [row.settlement_id, row.Central_places]));
}



async function loadFilterOptions() {
  setLoading(true, 'Формирование фильтров…', 'Из таблицы населённых пунктов извлекаются перечни субъектов РФ и муниципальных образований.');
  const regionRows = await queryRows(`
    SELECT DISTINCT Region
    FROM settlement_index
    WHERE Region IS NOT NULL AND TRIM(CAST(Region AS VARCHAR)) <> ''
    ORDER BY Region
  `);
  const municipalityRows = await queryRows(`
    SELECT DISTINCT Region, Municipality
    FROM settlement_index
    WHERE Region IS NOT NULL AND Municipality IS NOT NULL
      AND TRIM(CAST(Region AS VARCHAR)) <> ''
      AND TRIM(CAST(Municipality AS VARCHAR)) <> ''
    ORDER BY Region, Municipality
  `);

  state.filterOptions.regions = regionRows.map((row) => row.Region);
  state.filterOptions.municipalities = municipalityRows.map((row) => ({
    Region: row.Region,
    Municipality: row.Municipality,
  }));
  state.filterOptions.municipalitiesByRegion = new Map();

  for (const region of state.filterOptions.regions) {
    const municipalities = municipalityRows
      .filter((row) => row.Region === region)
      .map((row) => row.Municipality);
    state.filterOptions.municipalitiesByRegion.set(region, municipalities);
  }

  state.ui.regionTom.clearOptions();
  state.ui.regionTom.addOptions(state.filterOptions.regions.map((region) => ({ value: region, text: region })));
  state.ui.regionTom.settings.maxOptions = Math.max(5000, state.filterOptions.regions.length + 10);
  state.ui.regionTom.refreshOptions(false);
  updateMunicipalityOptions();
}



function updateMunicipalityOptions() {
  const selectedRegions = normalizeMultiValue(state.ui.regionTom.getValue());
  const selectedMunicipalities = new Set(normalizeMultiValue(state.ui.municipalityTom.getValue()));

  if (!selectedRegions.length) {
    state.ui.municipalityTom.clear(true);
    state.ui.municipalityTom.clearOptions();
    state.ui.municipalityTom.settings.placeholder = 'Сначала выберите субъект РФ';
    state.ui.municipalityTom.disable();
    state.ui.municipalityTom.inputState();
    if (els.municipalityHint) {
      els.municipalityHint.textContent = 'Список муниципалитетов раскрывается после выбора одного или нескольких субъектов РФ.';
    }
    return;
  }

  const allowedMunicipalities = new Set();
  for (const region of selectedRegions) {
    for (const municipality of state.filterOptions.municipalitiesByRegion.get(region) || []) {
      allowedMunicipalities.add(`${region}|||${municipality}`);
    }
  }
  const options = state.filterOptions.municipalities.filter((row) => allowedMunicipalities.has(`${row.Region}|||${row.Municipality}`));

  state.ui.municipalityTom.enable();
  state.ui.municipalityTom.clear(true);
  state.ui.municipalityTom.clearOptions();
  state.ui.municipalityTom.addOptions(options.map((row) => ({
    value: `${row.Region}|||${row.Municipality}`,
    text: `${row.Municipality} — ${row.Region}`,
  })));
  state.ui.municipalityTom.settings.maxOptions = Math.max(50000, options.length + 10);
  state.ui.municipalityTom.settings.placeholder = 'Все муниципальные образования выбранных субъектов РФ';

  const stillValid = [...selectedMunicipalities].filter((municipalityKey) =>
    options.some((row) => `${row.Region}|||${row.Municipality}` === municipalityKey));
  state.ui.municipalityTom.setValue(stillValid, true);
  state.ui.municipalityTom.inputState();
  state.ui.municipalityTom.refreshOptions(false);

  if (els.municipalityHint) {
    els.municipalityHint.textContent = `Доступно муниципалитетов: ${formatInteger(options.length)}.`;
  }
}



function initMap() {
  const style = {
    version: 8,
    sources: {
      carto: {
        type: 'raster',
        tiles: ['https://basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png'],
        tileSize: 256,
        attribution: '© OpenStreetMap contributors © CARTO',
      },
    },
    layers: [{ id: 'carto', type: 'raster', source: 'carto' }],
  };

  state.map = new maplibregl.Map({
    container: 'map',
    style,
    center: DEFAULT_MAP_VIEW.center,
    zoom: DEFAULT_MAP_VIEW.zoom,
    maxZoom: 13,
    renderWorldCopies: false,
    attributionControl: true,
    preserveDrawingBuffer: true,
  });
  state.map.addControl(new maplibregl.NavigationControl({ showCompass: false }), 'top-right');
  state.ui.hoverPopup = new maplibregl.Popup({ closeButton: false, closeOnClick: false, maxWidth: '360px', offset: 14 });

  state.ui.mapReady = new Promise((resolve) => {
    state.map.on('load', () => {
      state.deckOverlay = new deck.MapboxOverlay({ layers: [] });
      state.map.addControl(state.deckOverlay);
      state.map.on('zoomend', () => updateMapLayers());
      resolve();
    });
  });
}

function resetMapView() {
  if (!state.map) return;
  state.map.easeTo({ center: DEFAULT_MAP_VIEW.center, zoom: DEFAULT_MAP_VIEW.zoom, duration: 700 });
}

function fitMapToCurrentData() {
  if (!state.map) return;
  state.ui.pendingFitBounds = false;

  if (!state.currentMapData.length) {
    resetMapView();
    return;
  }

  if (state.currentMapData.length === 1) {
    const point = state.currentMapData[0];
    state.map.flyTo({ center: [point.Longitude, point.Latitude], zoom: 8.2, duration: 700 });
    return;
  }

  const bounds = new maplibregl.LngLatBounds();
  for (const row of state.currentMapData) {
    bounds.extend([row.Longitude, row.Latitude]);
  }
  state.map.fitBounds(bounds, { padding: 48, duration: 700, maxZoom: 9.5 });
}

function scheduleRefresh() {
  window.clearTimeout(state.ui.refreshTimer);
  state.ui.refreshTimer = window.setTimeout(() => {
    refreshAll().catch(handleError);
  }, 220);
}

function clearFilters(applyFitBounds = true) {
  state.ui.regionTom.clear(true);
  updateMunicipalityOptions();
  if (state.ui.municipalityTom) state.ui.municipalityTom.clear(true);
  if (applyFitBounds) {
    state.ui.pendingFitBounds = true;
  }
  scheduleRefresh();
}

async function refreshAll() {
  const token = ++state.ui.refreshToken;
  clearError();
  setLoading(true, 'Обновление дашборда…', 'Пересчитываются картографическая выборка, временные ряды и сопутствующие показатели.');

  const mapData = await loadCurrentMapData();
  if (token !== state.ui.refreshToken) return;
  state.currentMapData = mapData;
  renderMapCaption();
  renderKpis();
  renderTopList();
  updateMapLayers();
  if (state.ui.pendingFitBounds) {
    fitMapToCurrentData();
  }

  state.currentAggregateSeries = await loadAggregateSeries();
  if (token !== state.ui.refreshToken) return;
  renderAggregateChart();

  state.currentComponentSeries = await loadComponentSeries();
  if (token !== state.ui.refreshToken) return;
  renderFertilityBlock();
  renderMortalityBlock();
  renderMigrationBlock();
  await renderMigrationAgePyramids();
  if (token !== state.ui.refreshToken) return;

  if (state.selectedSettlementId != null) {
    await refreshSelectedSettlement();
    if (token !== state.ui.refreshToken) return;
  } else {
    state.currentSelectionBundle = null;
    renderSelectedSettlementCard();
    renderSelectedSettlementChart();
    await renderPyramid();
  }

  setLoading(false);
}



async function loadCurrentMapData() {
  const scenario = getScenario();
  const forecastAlias = getForecastAlias(scenario);
  const mapYear = getMapYear();
  const baseYear = getBaseYear();
  const whereClause = buildWhereClause('i');
  const popExpr = populationExpression(mapYear, 'i', 'w');
  const baseExpr = populationExpression(baseYear, 'i', 'w');

  const needsForecastJoin = mapYear >= 2021 || baseYear >= 2021;
  const joinSql = needsForecastJoin
    ? `LEFT JOIN read_parquet('${forecastAlias}') w ON CAST(i.settlement_id AS BIGINT) = CAST(w.settlement_id AS BIGINT)`
    : '';

  const sql = `
    SELECT
      CAST(i.settlement_id AS BIGINT) AS settlement_id,
      i.Region,
      i.Municipality,
      i.Settlement_full_name,
      i.Settlement_name,
      i.Type_settlement,
      CAST(i.Latitude AS DOUBLE) AS Latitude,
      CAST(i.Longitude AS DOUBLE) AS Longitude,
      CAST(c.settlement_id IS NOT NULL AS BOOLEAN) AS is_central_place,
      c.Central_places AS Central_places,
      ${popExpr} AS pop,
      ${baseExpr} AS base_pop
    FROM settlement_index i
    LEFT JOIN central_place_representatives c ON CAST(i.settlement_id AS BIGINT) = c.settlement_id
    ${joinSql}
    ${whereClause}
    ORDER BY pop DESC NULLS LAST
  `;

  const rows = await queryRows(sql);
  return rows
    .map((row) => {
      const pop = toNumber(row.pop);
      const basePop = toNumber(row.base_pop);
      const deltaAbs = Number.isFinite(pop) && Number.isFinite(basePop) ? pop - basePop : null;
      const deltaPct = Number.isFinite(basePop) && basePop > 0 && Number.isFinite(deltaAbs) ? (deltaAbs / basePop) * 100 : null;
      const centralMeta = row.is_central_place ? buildCentralPlaceMeta(row.Central_places) : null;
      return {
        settlement_id: toNumber(row.settlement_id),
        Region: row.Region,
        Municipality: row.Municipality,
        Settlement_full_name: row.Settlement_full_name,
        Settlement_name: row.Settlement_name,
        Type_settlement: row.Type_settlement,
        Latitude: toNumber(row.Latitude),
        Longitude: toNumber(row.Longitude),
        is_central_place: Boolean(row.is_central_place),
        Central_places: row.Central_places,
        central_group_id: centralMeta?.id || null,
        central_symbol: centralMeta?.symbol || null,
        central_group_label: centralMeta?.label || null,
        pop,
        base_pop: basePop,
        delta_abs: deltaAbs,
        delta_pct: deltaPct,
      };
    })
    .filter((row) => Number.isFinite(row.Latitude) && Number.isFinite(row.Longitude) && Number.isFinite(row.pop) && row.pop > 0);
}

async function loadAggregateSeries() {
  const whereClause = buildWhereClause('i');

  const actualSql = `
    SELECT
      SUM(CAST(i.Population_1989 AS DOUBLE)) AS pop_1989,
      SUM(CAST(i.Population_2002 AS DOUBLE)) AS pop_2002,
      SUM(CAST(i.Population_2010 AS DOUBLE)) AS pop_2010,
      SUM(CAST(i.Population_2021 AS DOUBLE)) AS pop_2021
    FROM settlement_index i
    ${whereClause}
  `;

  const forecastSelect = state.config.forecastYears.map((year) => `SUM(CAST(w.pop_${year} AS DOUBLE)) AS pop_${year}`).join(',\n      ');
  const actualRows = await queryRows(actualSql);
  const actual = actualRows[0] || {};
  const forecastsByScenario = {};

  await Promise.all(state.config.scenarios.map(async (scenario) => {
    const forecastSql = `
      SELECT
        ${forecastSelect}
      FROM read_parquet('${getForecastAlias(scenario.id)}') w
      JOIN settlement_index i USING (settlement_id)
      ${whereClause}
    `;
    const rows = await queryRows(forecastSql);
    const forecast = rows[0] || {};
    forecastsByScenario[scenario.id] = {
      years: state.config.forecastYears,
      values: state.config.forecastYears.map((year) => toNumber(forecast[`pop_${year}`])),
    };
  }));

  return {
    actualYears: state.config.censusYears,
    actualValues: state.config.censusYears.map((year) => toNumber(actual[`pop_${year}`])),
    forecastsByScenario,
  };
}

function buildSelectedOktmoSubquery() {
  const whereClause = buildWhereClause('i');
  if (whereClause) {
    return `SELECT DISTINCT i.oktmo_syn FROM settlement_index i ${whereClause} AND i.oktmo_syn IS NOT NULL AND TRIM(CAST(i.oktmo_syn AS VARCHAR)) <> ''`;
  }
  return `SELECT DISTINCT i.oktmo_syn FROM settlement_index i WHERE i.oktmo_syn IS NOT NULL AND TRIM(CAST(i.oktmo_syn AS VARCHAR)) <> ''`;
}

function buildForecastComponentSelectParts(years) {
  const parts = [];
  years.forEach((year) => {
    const birthDen = `CASE WHEN w.tfr_${year} IS NOT NULL AND w.tfr_${year} > 0 AND w.births_total_${year} IS NOT NULL THEN (w.births_total_${year} * 35.0) / w.tfr_${year} END`;
    const deathDen = `CASE WHEN w.cdr_${year} IS NOT NULL AND w.cdr_${year} > 0 AND w.deaths_total_${year} IS NOT NULL THEN (w.deaths_total_${year} * 1000.0) / w.cdr_${year} END`;
    parts.push(`SUM(COALESCE(w.births_total_${year}, 0)) AS births_total_${year}`);
    parts.push(`CASE WHEN SUM(${birthDen}) > 0 THEN 35.0 * SUM(COALESCE(w.births_total_${year}, 0)) / SUM(${birthDen}) END AS tfr_${year}`);
    parts.push(`SUM(COALESCE(w.deaths_total_${year}, 0)) AS deaths_total_${year}`);
    parts.push(`CASE WHEN SUM(${deathDen}) > 0 THEN 1000.0 * SUM(COALESCE(w.deaths_total_${year}, 0)) / SUM(${deathDen}) END AS cdr_${year}`);
    parts.push(`SUM(COALESCE(w.mig_in_total_${year}, 0)) AS mig_in_total_${year}`);
    parts.push(`SUM(COALESCE(w.mig_out_total_${year}, 0)) AS mig_out_total_${year}`);
    parts.push(`SUM(COALESCE(w.mig_net_total_${year}, 0)) AS mig_net_total_${year}`);
  });
  return parts.join(',\n        ');
}

async function loadForecastComponentSeries(scenarioId) {
  const subquery = buildSelectedOktmoSubquery();
  const selectParts = buildForecastComponentSelectParts(state.config.forecastYears);
  const rows = await queryRows(`
    WITH selected_oktmo AS (
      ${subquery}
    )
    SELECT
      ${selectParts}
    FROM read_parquet('municipal_components_wide_${scenarioId}.parquet') w
    JOIN selected_oktmo s USING (oktmo_syn)
  `);
  const row = rows[0] || {};
  return {
    years: state.config.forecastYears,
    births: state.config.forecastYears.map((year) => toNumber(row[`births_total_${year}`])),
    tfr: state.config.forecastYears.map((year) => toNumber(row[`tfr_${year}`])),
    deaths: state.config.forecastYears.map((year) => toNumber(row[`deaths_total_${year}`])),
    cdr: state.config.forecastYears.map((year) => toNumber(row[`cdr_${year}`])),
    migIn: state.config.forecastYears.map((year) => toNumber(row[`mig_in_total_${year}`])),
    migOut: state.config.forecastYears.map((year) => toNumber(row[`mig_out_total_${year}`])),
    migNet: state.config.forecastYears.map((year) => toNumber(row[`mig_net_total_${year}`])),
  };
}

async function loadComponentSeries() {
  const subquery = buildSelectedOktmoSubquery();
  const actualRows = await queryRows(`
    WITH selected_oktmo AS (
      ${subquery}
    )
    SELECT
      CAST(a.year AS BIGINT) AS year,
      SUM(COALESCE(a.births_total, 0)) AS births_total,
      CASE
        WHEN SUM(COALESCE(a.expo_eff, CASE WHEN a.tfr IS NOT NULL AND a.tfr > 0 AND a.births_total IS NOT NULL THEN (a.births_total * 35.0) / a.tfr END)) > 0
        THEN 35.0 * SUM(COALESCE(a.births_total, 0)) / SUM(COALESCE(a.expo_eff, CASE WHEN a.tfr IS NOT NULL AND a.tfr > 0 AND a.births_total IS NOT NULL THEN (a.births_total * 35.0) / a.tfr END))
      END AS tfr,
      SUM(COALESCE(a.deaths_total, 0)) AS deaths_total,
      CASE
        WHEN SUM(COALESCE(a.pop_mid, CASE WHEN a.cdr IS NOT NULL AND a.cdr > 0 AND a.deaths_total IS NOT NULL THEN (a.deaths_total * 1000.0) / a.cdr END)) > 0
        THEN 1000.0 * SUM(COALESCE(a.deaths_total, 0)) / SUM(COALESCE(a.pop_mid, CASE WHEN a.cdr IS NOT NULL AND a.cdr > 0 AND a.deaths_total IS NOT NULL THEN (a.deaths_total * 1000.0) / a.cdr END))
      END AS cdr,
      SUM(COALESCE(a.mig_in_total, 0)) AS mig_in_total,
      SUM(COALESCE(a.mig_out_total, 0)) AS mig_out_total,
      SUM(COALESCE(a.mig_net_total, 0)) AS mig_net_total
    FROM read_parquet('municipal_components_actual.parquet') a
    JOIN selected_oktmo s USING (oktmo_syn)
    GROUP BY a.year
    ORDER BY a.year
  `);

  const actualYears = actualRows.map((row) => toNumber(row.year)).filter((year) => Number.isFinite(year));
  const actualBirths = actualRows.map((row) => toNumber(row.births_total));
  const actualTfr = actualRows.map((row) => toNumber(row.tfr));
  const actualDeaths = actualRows.map((row) => toNumber(row.deaths_total));
  const actualCdr = actualRows.map((row) => toNumber(row.cdr));
  const actualMigIn = actualRows.map((row) => toNumber(row.mig_in_total));
  const actualMigOut = actualRows.map((row) => toNumber(row.mig_out_total));
  const actualMigNet = actualRows.map((row) => toNumber(row.mig_net_total));

  const forecastsByScenario = {};
  await Promise.all(state.config.scenarios.map(async (scenario) => {
    forecastsByScenario[scenario.id] = await loadForecastComponentSeries(scenario.id);
  }));

  return {
    fertility: {
      actualYears,
      birthsActual: actualBirths,
      tfrActual: actualTfr,
      forecastsByScenario,
    },
    mortality: {
      actualYears,
      deathsActual: actualDeaths,
      cdrActual: actualCdr,
      forecastsByScenario,
    },
    migration: {
      actualYears,
      migInActual: actualMigIn,
      migOutActual: actualMigOut,
      migNetActual: actualMigNet,
      forecastsByScenario,
    },
  };
}


function renderAggregateChart() {
  if (!state.currentAggregateSeries) {
    renderEmptyPlot('aggregateChart', 'Нет данных для текущего набора фильтров.');
    return;
  }

  const traces = [];
  const actualSeries = filterSeriesToWindow(state.currentAggregateSeries.actualYears, state.currentAggregateSeries.actualValues);
  if (actualSeries.x.length) {
    traces.push({
      x: actualSeries.x,
      y: actualSeries.y,
      type: 'scatter',
      mode: 'lines+markers',
      name: 'Фактические данные',
      line: { color: '#1d4ed8', width: 3 },
      marker: { color: '#1d4ed8', size: 9 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    });
  }

  for (const scenarioId of getSelectedChartScenarios('aggregate')) {
    const style = SCENARIO_TRACE_STYLES[scenarioId] || { color: '#334155', dash: 'solid' };
    const scenarioSeries = state.currentAggregateSeries.forecastsByScenario[scenarioId];
    const filtered = filterSeriesToWindow(scenarioSeries?.years || [], scenarioSeries?.values || []);
    if (!filtered.x.length) continue;
    traces.push({
      x: filtered.x,
      y: filtered.y,
      type: 'scatter',
      mode: 'lines',
      name: `Прогноз (${getScenarioLabel(scenarioId)})`,
      line: { color: style.color, width: 3, dash: style.dash },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    });
  }

  if (!traces.length) {
    renderEmptyPlot('aggregateChart', 'В пределах выбранного временного интервала данные отсутствуют.');
    return;
  }

  const layout = buildLineLayout('Численность населения', describeCurrentFilter());
  Plotly.newPlot('aggregateChart', traces, layout, buildPlotConfig(buildFilename('aggregate')));
}

function buildActualForecastTraces({ actualYears, actualValues, forecastSeriesByScenario, selectedScenarios, actualName, yLabel, rateMode = false }) {
  const traces = [];
  const actualSeries = filterSeriesToWindow(actualYears || [], actualValues || []);
  if (actualSeries.x.length) {
    traces.push({
      x: actualSeries.x,
      y: actualSeries.y,
      type: 'scatter',
      mode: 'lines+markers',
      name: actualName,
      line: { color: '#1d4ed8', width: 3 },
      marker: { color: '#1d4ed8', size: 8 },
      hovertemplate: rateMode
        ? `<b>%{x}</b><br>${yLabel}: %{y:.2f}<extra></extra>`
        : `<b>%{x}</b><br>${yLabel}: %{y:,}<extra></extra>`,
    });
  }

  for (const scenarioId of selectedScenarios) {
    const style = SCENARIO_TRACE_STYLES[scenarioId] || { color: '#334155', dash: 'solid' };
    const series = forecastSeriesByScenario?.[scenarioId];
    const filtered = filterSeriesToWindow(series?.years || [], series?.values || []);
    if (!filtered.x.length) continue;
    traces.push({
      x: filtered.x,
      y: filtered.y,
      type: 'scatter',
      mode: 'lines',
      name: `Прогноз (${getScenarioLabel(scenarioId)})`,
      line: { color: style.color, width: 3, dash: style.dash },
      hovertemplate: rateMode
        ? `<b>%{x}</b><br>${yLabel}: %{y:.2f}<extra></extra>`
        : `<b>%{x}</b><br>${yLabel}: %{y:,}<extra></extra>`,
    });
  }
  return traces;
}

function renderFertilityBlock() {
  const bundle = state.currentComponentSeries?.fertility;
  if (!bundle) {
    renderEmptyPlot('fertilityBirthsChart', 'Нет данных по рождаемости для текущей выборки.');
    renderEmptyPlot('fertilityTfrChart', 'Нет данных по СКР для текущей выборки.');
    return;
  }

  const selectedScenarios = getSelectedChartScenarios('fertility');
  const birthsForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.births }]));
  const tfrForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.tfr }]));

  const birthsTraces = buildActualForecastTraces({
    actualYears: bundle.actualYears,
    actualValues: bundle.birthsActual,
    forecastSeriesByScenario: birthsForecast,
    selectedScenarios,
    actualName: 'Фактические данные',
    yLabel: 'Число рождений',
    rateMode: false,
  });
  if (birthsTraces.length) {
    const layout = buildLineLayout('Число рождений', describeCurrentFilter());
    layout.yaxis.tickformat = ',d';
    Plotly.newPlot('fertilityBirthsChart', birthsTraces, layout, buildPlotConfig(buildFilename('fertility_births')));
  } else {
    renderEmptyPlot('fertilityBirthsChart', 'В пределах выбранного временного интервала данные отсутствуют.');
  }

  const tfrTraces = buildActualForecastTraces({
    actualYears: bundle.actualYears,
    actualValues: bundle.tfrActual,
    forecastSeriesByScenario: tfrForecast,
    selectedScenarios,
    actualName: 'Фактические данные',
    yLabel: 'СКР',
    rateMode: true,
  });
  if (tfrTraces.length) {
    const layout = buildLineLayout('Суммарный коэффициент рождаемости', describeCurrentFilter());
    delete layout.yaxis.tickformat;
    Plotly.newPlot('fertilityTfrChart', tfrTraces, layout, buildPlotConfig(buildFilename('fertility_tfr')));
  } else {
    renderEmptyPlot('fertilityTfrChart', 'В пределах выбранного временного интервала данные отсутствуют.');
  }
}

function renderMortalityBlock() {
  const bundle = state.currentComponentSeries?.mortality;
  if (!bundle) {
    renderEmptyPlot('mortalityDeathsChart', 'Нет данных по смертности для текущей выборки.');
    renderEmptyPlot('mortalityCdrChart', 'Нет данных по ОКС для текущей выборки.');
    return;
  }

  const selectedScenarios = getSelectedChartScenarios('mortality');
  const deathsForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.deaths }]));
  const cdrForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.cdr }]));

  const deathsTraces = buildActualForecastTraces({
    actualYears: bundle.actualYears,
    actualValues: bundle.deathsActual,
    forecastSeriesByScenario: deathsForecast,
    selectedScenarios,
    actualName: 'Фактические данные',
    yLabel: 'Число умерших',
    rateMode: false,
  });
  if (deathsTraces.length) {
    const layout = buildLineLayout('Число умерших', describeCurrentFilter());
    layout.yaxis.tickformat = ',d';
    Plotly.newPlot('mortalityDeathsChart', deathsTraces, layout, buildPlotConfig(buildFilename('mortality_deaths')));
  } else {
    renderEmptyPlot('mortalityDeathsChart', 'В пределах выбранного временного интервала данные отсутствуют.');
  }

  const cdrTraces = buildActualForecastTraces({
    actualYears: bundle.actualYears,
    actualValues: bundle.cdrActual,
    forecastSeriesByScenario: cdrForecast,
    selectedScenarios,
    actualName: 'Фактические данные',
    yLabel: 'ОКС',
    rateMode: true,
  });
  if (cdrTraces.length) {
    const layout = buildLineLayout('Общий коэффициент смертности', describeCurrentFilter());
    delete layout.yaxis.tickformat;
    Plotly.newPlot('mortalityCdrChart', cdrTraces, layout, buildPlotConfig(buildFilename('mortality_cdr')));
  } else {
    renderEmptyPlot('mortalityCdrChart', 'В пределах выбранного временного интервала данные отсутствуют.');
  }
}

function renderMigrationBlock() {
  const bundle = state.currentComponentSeries?.migration;
  if (!bundle) {
    renderEmptyPlot('migrationTotalsChart', 'Нет данных по миграции для текущей выборки.');
    return;
  }

  const selectedScenarios = getSelectedChartScenarios('migration');
  const measureMeta = [
    { key: 'migIn', actualValues: bundle.migInActual, forecastKey: 'migIn', actualName: 'Фактический приток', colorActual: '#1d4ed8', colorForecast: '#60a5fa' },
    { key: 'migOut', actualValues: bundle.migOutActual, forecastKey: 'migOut', actualName: 'Фактический отток', colorActual: '#7c3aed', colorForecast: '#c084fc' },
    { key: 'migNet', actualValues: bundle.migNetActual, forecastKey: 'migNet', actualName: 'Фактическое сальдо', colorActual: '#0f766e', colorForecast: '#34d399' },
  ];

  const traces = [];
  measureMeta.forEach((meta) => {
    const actualSeries = filterSeriesToWindow(bundle.actualYears, meta.actualValues);
    if (actualSeries.x.length) {
      traces.push({
        x: actualSeries.x,
        y: actualSeries.y,
        type: 'scatter',
        mode: 'lines+markers',
        name: meta.actualName,
        line: { color: meta.colorActual, width: 3 },
        marker: { color: meta.colorActual, size: 7 },
        hovertemplate: `<b>%{x}</b><br>${meta.actualName.replace('Фактический ', '').replace('Фактическое ', '')}: %{y:,}<extra></extra>`,
      });
    }
  });

  selectedScenarios.forEach((scenarioId) => {
    const style = SCENARIO_TRACE_STYLES[scenarioId] || { dash: 'solid' };
    const forecastData = bundle.forecastsByScenario?.[scenarioId];
    measureMeta.forEach((meta) => {
      const filtered = filterSeriesToWindow(forecastData?.years || [], forecastData?.[meta.forecastKey] || []);
      if (!filtered.x.length) return;
      const baseName = meta.actualName.replace('Фактический ', '').replace('Фактическое ', '');
      traces.push({
        x: filtered.x,
        y: filtered.y,
        type: 'scatter',
        mode: 'lines',
        name: `Прогноз: ${baseName} (${getScenarioLabel(scenarioId)})`,
        line: { color: meta.colorForecast, width: 2.8, dash: style.dash },
        hovertemplate: `<b>%{x}</b><br>${baseName}: %{y:,}<extra></extra>`,
      });
    });
  });

  if (!traces.length) {
    renderEmptyPlot('migrationTotalsChart', 'В пределах выбранного временного интервала данные отсутствуют.');
    return;
  }

  const layout = buildLineLayout('Миграционные потоки', describeCurrentFilter());
  layout.yaxis.tickformat = ',d';
  Plotly.newPlot('migrationTotalsChart', traces, layout, buildPlotConfig(buildFilename('migration_totals')));
}

function buildAgeBandSortKey(label) {
  const text = String(label || '').trim();
  const match = text.match(/^(\d+)/);
  if (match) return Number(match[1]);
  if (/100\+|100\s*\+|100 и более/i.test(text)) return 100;
  return 999;
}

function buildAgeBandCategoryArray(rows) {
  return [...new Set(rows.map((row) => row.age_band))].sort((a, b) => buildAgeBandSortKey(a) - buildAgeBandSortKey(b));
}

function sexRank(value) {
  const text = String(value || '').trim().toLowerCase();
  if (text.startsWith('м')) return 0;
  if (text.startsWith('f') || text.startsWith('ж')) return 1;
  return 2;
}

function renderMigrationAgePyramids() {
  updateMigrationScenarioNote();
  return (async () => {
    const scenario = getScenario();
    const year = getMigrationPyramidYear();
    const alias = getMunicipalMigrationAgeAlias(scenario, year);
    await ensureRegisteredFile(alias, `./data/municipal_migration_age/scenario=${scenario}/year=${year}.parquet`);

    const subquery = buildSelectedOktmoSubquery();
    const rows = await queryRows(`
      WITH selected_oktmo AS (
        ${subquery}
      )
      SELECT sex, age_band, SUM(COALESCE(inflow, 0)) AS inflow, SUM(COALESCE(outflow, 0)) AS outflow, SUM(COALESCE(net, 0)) AS net
      FROM read_parquet('${alias}') m
      JOIN selected_oktmo s USING (oktmo_syn)
      GROUP BY sex, age_band
      ORDER BY age_band, sex
    `);

    const cleanedRows = rows
      .map((row) => ({
        sex: String(row.sex || '').trim(),
        age_band: String(row.age_band || '').trim(),
        inflow: Math.max(0, toNumber(row.inflow) || 0),
        outflow: Math.max(0, toNumber(row.outflow) || 0),
        net: toNumber(row.net) || 0,
      }))
      .filter((row) => row.age_band)
      .sort((a, b) => buildAgeBandSortKey(a.age_band) - buildAgeBandSortKey(b.age_band) || sexRank(a.sex) - sexRank(b.sex));

    state.currentMigrationAgeBundle = { year, scenario, rows: cleanedRows };

    renderMigrationAgePyramidChart('migrationInflowPyramidChart', cleanedRows, 'inflow', 'Половозрастная структура притока');
    renderMigrationAgePyramidChart('migrationOutflowPyramidChart', cleanedRows, 'outflow', 'Половозрастная структура оттока');
  })();
}

function renderMigrationAgePyramidChart(containerId, rows, metric, title) {
  if (!rows.length) {
    renderEmptyPlot(containerId, 'Для выбранной пространственной выборки возрастно-половые данные миграции отсутствуют.');
    return;
  }
  const totalMetric = rows.reduce((sum, row) => sum + (toNumber(row[metric]) || 0), 0);
  if (!Number.isFinite(totalMetric) || totalMetric <= 0) {
    renderEmptyPlot(containerId, 'В выбранном году миграционные потоки по возрастным группам отсутствуют.');
    return;
  }

  const ages = buildAgeBandCategoryArray(rows);
  const male = ages.map((age) => -1 * (rows.find((row) => row.age_band === age && String(row.sex).toLowerCase().startsWith('м'))?.[metric] || 0));
  const female = ages.map((age) => rows.find((row) => row.age_band === age && !String(row.sex).toLowerCase().startsWith('м'))?.[metric] || 0);
  const metricLabel = metric === 'inflow' ? 'Приток' : 'Отток';

  const traces = [
    {
      x: male,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Мужчины',
      marker: { color: '#2563eb' },
      hovertemplate: `<b>%{y}</b><br>Мужчины: %{customdata:,}<extra></extra>`,
      customdata: male.map((value) => Math.abs(value)),
    },
    {
      x: female,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Женщины',
      marker: { color: '#ec4899' },
      hovertemplate: `<b>%{y}</b><br>Женщины: %{x:,}<extra></extra>`,
    },
  ];

  const layout = {
    title: {
      text: `${title} · ${getMigrationPyramidYear()}`,
      font: { size: 16 },
      x: 0.02,
      xanchor: 'left',
    },
    barmode: 'relative',
    bargap: 0.05,
    margin: { l: 84, r: 32, t: 58, b: 48 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    legend: { orientation: 'h', x: 0, y: 1.08 },
    xaxis: {
      title: metricLabel,
      zeroline: true,
      zerolinewidth: 1,
      zerolinecolor: '#94a3b8',
      gridcolor: '#e2e8f0',
      tickformat: ',d',
    },
    yaxis: {
      title: 'Возрастная группа',
      categoryorder: 'array',
      categoryarray: ages,
      gridcolor: '#f1f5f9',
    },
    annotations: [
      {
        text: `Суммарно: ${formatInteger(totalMetric)} · сценарий: ${getScenarioLabel(getScenario())}`,
        xref: 'paper',
        yref: 'paper',
        x: 0,
        y: 1.16,
        xanchor: 'left',
        yanchor: 'bottom',
        showarrow: false,
        font: { size: 12, color: '#475569' },
      },
    ],
  };

  Plotly.newPlot(containerId, traces, layout, buildPlotConfig(buildFilename(metric === 'inflow' ? 'migration_inflow_pyramid' : 'migration_outflow_pyramid')));
}


function renderTopList() {
  const rows = [...state.currentMapData]
    .sort((a, b) => (b.pop || 0) - (a.pop || 0))
    .slice(0, 25);

  if (!rows.length) {
    els.topSettlements.innerHTML = '<div class="selection-summary empty-state">По текущему набору фильтров данных нет.</div>';
    return;
  }

  const tableRows = rows.map((row, index) => `
    <tr data-settlement-id="${row.settlement_id}">
      <td class="rank">${index + 1}</td>
      <td>${escapeHtml(row.Settlement_full_name || row.Settlement_name || 'Без названия')}</td>
      <td>${escapeHtml(row.Municipality || '—')}</td>
      <td>${formatInteger(row.pop)}</td>
      <td>${formatPct(row.delta_pct)}</td>
    </tr>
  `).join('');

  els.topSettlements.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>№</th>
          <th>Населённый пункт</th>
          <th>Муниципалитет</th>
          <th>Численность</th>
          <th>Δ, %</th>
        </tr>
      </thead>
      <tbody>${tableRows}</tbody>
    </table>
  `;

  els.topSettlements.querySelectorAll('tbody tr').forEach((rowEl) => {
    rowEl.addEventListener('click', () => {
      const settlementId = Number(rowEl.dataset.settlementId);
      selectSettlement(settlementId, true).catch(handleError);
    });
  });
}

function renderKpis() {
  const settlements = state.currentMapData.length;
  const totalPopulation = state.currentMapData.reduce((sum, row) => sum + (Number.isFinite(row.pop) ? row.pop : 0), 0);
  const validDeltaRows = state.currentMapData.filter((row) => Number.isFinite(row.delta_pct));
  const avgDelta = validDeltaRows.length
    ? validDeltaRows.reduce((sum, row) => sum + row.delta_pct, 0) / validDeltaRows.length
    : null;
  const centralVisibleCount = state.currentMapData.filter((row) => row.is_central_place).length;
  const centralTotalCount = countCentralRepresentativesForCurrentFilter();

  els.kpiSettlements.textContent = formatInteger(settlements);
  els.kpiFilterNote.textContent = describeCurrentFilter();
  els.kpiPopulation.textContent = formatInteger(totalPopulation);
  els.kpiPopulationNote.textContent = `Год на карте: ${getMapYear()} · сценарий: ${getScenarioLabel(getScenario())}`;
  els.kpiDelta.textContent = formatPct(avgDelta);
  els.kpiDeltaNote.textContent = `Относительно ${getBaseYear()} года`;
  els.kpiCentral.textContent = formatInteger(centralTotalCount);
  els.kpiCentralNote.textContent = `Официальный перечень в выборке; на карте в ${getMapYear()} г. видны ${formatInteger(centralVisibleCount)}${els.centralPlacesToggle.checked ? '' : ' · слой отключён'}`;
}

function renderMapCaption() {
  const mapYear = getMapYear();
  const baseYear = getBaseYear();
  const min = getColorNegativeThreshold();
  const max = getColorPositiveThreshold();
  els.mapCaption.textContent = `Год на карте: ${mapYear}. Цвет: прирост численности населения ${mapYear} г. к ${baseYear} г.; шкала от ${formatSignedPctThreshold(min)} до ${formatSignedPctThreshold(max)}. Размер окружности отражает фиксированные классы численности населения. Фильтр: ${describeCurrentFilter()}.`;
  renderMapLegend();
}

function renderMapLegend() {
  els.mapColorLegendTitle.textContent = `Цвет: прирост численности населения ${getMapYear()} г. к ${getBaseYear()} г.`;
  els.mapColorLegendLabels.innerHTML = `
    <span>${formatSignedPctThreshold(getColorNegativeThreshold())}</span>
    <span>0%</span>
    <span>${formatSignedPctThreshold(getColorPositiveThreshold())}</span>
  `;
}

function renderSizeLegend() {
  els.sizeLegend.innerHTML = POPULATION_CLASSES.map((item) => {
    const size = Math.max(8, item.radius * 2);
    return `
      <div class="size-legend-item">
        <span class="size-legend-bubble" style="width:${size}px;height:${size}px;"></span>
        <span>${escapeHtml(item.shortLabel)}</span>
      </div>
    `;
  }).join('');
}

function renderCentralTypeLegend() {
  const visibleGroups = CENTRAL_PLACE_GROUPS.filter((item) => item.id !== 'other');
  els.centralTypeLegend.innerHTML = visibleGroups.map((item) => `
    <div class="central-legend-item">
      <span class="central-outline-demo">${escapeHtml(item.symbol)}</span>
      <span><span class="legend-emphasis">${escapeHtml(item.symbol)}</span> — ${escapeHtml(item.label)}</span>
    </div>
  `).join('');
}

function updateMapLayers() {
  if (!state.deckOverlay) return;

  const selectedId = state.selectedSettlementId;
  const centralVisible = els.centralPlacesToggle.checked;
  const currentZoom = state.map?.getZoom() ?? DEFAULT_MAP_VIEW.zoom;
  const centralData = centralVisible ? state.currentMapData.filter((d) => d.is_central_place) : [];
  const centralSymbolData = centralData.filter((d) => d.central_symbol && (radiusForPopulation(d.pop) >= 8 || currentZoom >= 5.2));

  const layers = [
    new deck.ScatterplotLayer({
      id: 'settlements-main',
      data: state.currentMapData,
      pickable: true,
      opacity: 0.82,
      stroked: true,
      filled: true,
      radiusUnits: 'pixels',
      radiusScale: 1,
      lineWidthMinPixels: 0.6,
      getPosition: (d) => [d.Longitude, d.Latitude],
      getRadius: (d) => radiusForPopulation(d.pop),
      getFillColor: (d) => colorForDelta(d.delta_pct),
      getLineColor: (d) => (d.settlement_id === selectedId ? [255, 255, 255, 245] : [255, 255, 255, 128]),
      getLineWidth: (d) => (d.settlement_id === selectedId ? 2.2 : 0.6),
      autoHighlight: true,
      highlightColor: [17, 24, 39, 220],
      onHover: handleMapHover,
      onClick: (info) => {
        if (info.object) {
          selectSettlement(info.object.settlement_id, false).catch(handleError);
        }
      },
      updateTriggers: {
        getLineColor: [selectedId, getMapYear(), getBaseYear(), getColorNegativeThreshold(), getColorPositiveThreshold()],
        getLineWidth: [selectedId],
        getFillColor: [getMapYear(), getBaseYear(), getColorNegativeThreshold(), getColorPositiveThreshold()],
      },
    }),
    new deck.ScatterplotLayer({
      id: 'central-outline',
      data: centralData,
      pickable: false,
      stroked: true,
      filled: false,
      radiusUnits: 'pixels',
      radiusScale: 1,
      lineWidthMinPixels: 2.2,
      getPosition: (d) => [d.Longitude, d.Latitude],
      getRadius: (d) => Math.min(radiusForPopulation(d.pop) + 2.6, 30),
      getLineColor: [37, 99, 235, 240],
    }),
    new deck.TextLayer({
      id: 'central-symbols',
      data: centralSymbolData,
      pickable: false,
      characterSet: 'auto',
      sizeUnits: 'pixels',
      fontFamily: 'Arial, sans-serif',
      fontWeight: 700,
      getPosition: (d) => [d.Longitude, d.Latitude],
      getText: (d) => d.central_symbol,
      getSize: (d) => Math.max(10, radiusForPopulation(d.pop) * 0.95),
      getColor: [37, 99, 235, 255],
      getTextAnchor: 'middle',
      getAlignmentBaseline: 'center',
    }),
  ];

  state.deckOverlay.setProps({ layers });
}

function handleMapHover(info) {
  if (!state.ui.hoverPopup) return;
  const object = info.object;
  if (!object) {
    state.ui.hoverPopup.remove();
    return;
  }

  const centralLine = object.is_central_place
    ? `<br><span class="popup-emph">Опорный пункт:</span> ${escapeHtml(object.central_group_label || 'да')}`
    : '';

  const html = `
    <div class="popup-card">
      <strong>${escapeHtml(object.Settlement_full_name || object.Settlement_name || 'Населённый пункт')}</strong><br>
      ${escapeHtml(object.Type_settlement || '')}${object.Type_settlement ? ' · ' : ''}${escapeHtml(object.Municipality || '')}<br>
      ${escapeHtml(object.Region || '')}<br>
      <span>Численность (${getMapYear()}): ${formatInteger(object.pop)}</span><br>
      <span>Изменение к ${getBaseYear()} году: ${formatPct(object.delta_pct)}</span>${centralLine}
    </div>
  `;
  state.ui.hoverPopup
    .setLngLat(info.coordinate)
    .setHTML(html)
    .addTo(state.map);
}

async function selectSettlement(settlementId, flyTo = false) {
  state.selectedSettlementId = Number(settlementId);
  await refreshSelectedSettlement();
  const selected = state.currentMapData.find((row) => row.settlement_id === state.selectedSettlementId);
  if (selected && flyTo && state.map) {
    state.map.flyTo({ center: [selected.Longitude, selected.Latitude], zoom: Math.max(state.map.getZoom(), 6.2), duration: 700 });
  }
  updateMapLayers();
}

async function refreshSelectedSettlement() {
  const settlementId = state.selectedSettlementId;
  if (settlementId == null) {
    renderSelectedSettlementCard();
    renderSelectedSettlementChart();
    await renderPyramid();
    return;
  }

  const details = await loadSettlementDetails(settlementId);
  const forecastEntries = await Promise.all(
    state.config.scenarios.map(async (scenario) => [scenario.id, await loadSettlementForecastRow(settlementId, scenario.id)]),
  );
  const forecastRowsByScenario = Object.fromEntries(forecastEntries);
  const summary = buildSelectedSettlementSummary(details, forecastRowsByScenario);
  state.currentSelectionBundle = summary;
  renderSelectedSettlementCard();
  renderSelectedSettlementChart();
  await renderPyramid();
}

async function loadSettlementDetails(settlementId) {
  const rows = await queryRows(`
    SELECT *
    FROM read_parquet('settlement_details.parquet')
    WHERE CAST(settlement_id AS BIGINT) = ${Number(settlementId)}
    LIMIT 1
  `);
  return rows[0] || null;
}

async function loadSettlementForecastRow(settlementId, scenario) {
  const alias = getForecastAlias(scenario);
  const rows = await queryRows(`
    SELECT *
    FROM read_parquet('${alias}')
    WHERE CAST(settlement_id AS BIGINT) = ${Number(settlementId)}
    LIMIT 1
  `);
  return rows[0] || null;
}

function buildSelectedSettlementSummary(details, forecastRowsByScenario) {
  if (!details) {
    return null;
  }

  const actualYears = [];
  const actualValues = [];
  for (const year of state.config.censusYears) {
    const value = toNumber(details[`Population_${year}`]);
    if (Number.isFinite(value)) {
      actualYears.push(year);
      actualValues.push(value);
    }
  }

  const seriesByScenario = {};
  const byYearByScenario = {};

  for (const scenario of state.config.scenarios) {
    const forecastRow = forecastRowsByScenario[scenario.id] || null;
    const forecastYears = [];
    const forecastValues = [];
    if (forecastRow) {
      for (const year of state.config.forecastYears) {
        const value = toNumber(forecastRow[`pop_${year}`]);
        if (Number.isFinite(value)) {
          forecastYears.push(year);
          forecastValues.push(value);
        }
      }
    }

    const byYear = new Map();
    actualYears.forEach((year, index) => byYear.set(year, actualValues[index]));
    forecastYears.forEach((year, index) => byYear.set(year, forecastValues[index]));

    seriesByScenario[scenario.id] = {
      years: forecastYears,
      values: forecastValues,
    };
    byYearByScenario[scenario.id] = byYear;
  }

  const activeScenario = getScenario();
  const activeMap = byYearByScenario[activeScenario] || new Map();
  const selectedYear = getMapYear();
  const baseYear = getBaseYear();
  const selectedPop = activeMap.get(selectedYear) ?? null;
  const basePop = activeMap.get(baseYear) ?? null;
  const deltaAbs = Number.isFinite(selectedPop) && Number.isFinite(basePop) ? selectedPop - basePop : null;
  const deltaPct = Number.isFinite(basePop) && basePop > 0 && Number.isFinite(deltaAbs) ? (deltaAbs / basePop) * 100 : null;

  return {
    details,
    forecastRowsByScenario,
    actualYears,
    actualValues,
    seriesByScenario,
    byYearByScenario,
    selectedPop,
    basePop,
    deltaAbs,
    deltaPct,
  };
}

function renderSelectedSettlementCard() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    els.selectedSettlementSummary.className = 'selection-summary empty-state';
    els.selectedSettlementSummary.textContent = 'Населённый пункт ещё не выбран.';
    els.selectedSettlementMeta.innerHTML = '';
    els.selectionHint.textContent = 'Щёлкните по окружности на карте или выберите пункт из таблицы-лидера.';
    return;
  }

  const details = summary.details;
  const activeScenarioLabel = getScenarioLabel(getScenario());
  const effectiveCentralText = getEffectiveCentralPlaceText(details.settlement_id);
  els.selectedSettlementSummary.className = 'selection-summary';
  els.selectionHint.textContent = 'Карточка показывает атрибуты выбранного населённого пункта из файла Settlement_names.csv.';

  els.selectedSettlementSummary.innerHTML = `
    <div>
      <div style="color:#475569;font-size:0.92rem;">${escapeHtml(details.Type_settlement || 'Тип не указан')} · ${escapeHtml(details.Municipality || 'Муниципалитет не указан')}</div>
      <h3 style="margin:0.35rem 0 0;font-size:1.25rem;">${escapeHtml(details.Settlement_full_name || details.Settlement_name || 'Без названия')}</h3>
      <div style="margin-top:0.35rem;color:#475569;">${escapeHtml(details.Region || 'Регион не указан')}</div>
      <div style="margin-top:0.55rem;color:#0f172a;">${effectiveCentralText ? `Статус опорного пункта: <strong>${escapeHtml(String(effectiveCentralText))}</strong>.` : 'Статус опорного пункта не отмечен.'}</div>
      <div style="margin-top:0.35rem;color:#475569;">Активный сценарий карты и пирамиды: ${escapeHtml(activeScenarioLabel)}</div>
    </div>
    <div class="selection-grid">
      <div class="summary-kpi"><div class="label">Численность (${getMapYear()})</div><div class="value">${formatInteger(summary.selectedPop)}</div></div>
      <div class="summary-kpi"><div class="label">Базовый год (${getBaseYear()})</div><div class="value">${formatInteger(summary.basePop)}</div></div>
      <div class="summary-kpi"><div class="label">Абсолютное изменение</div><div class="value">${formatSigned(summary.deltaAbs)}</div></div>
      <div class="summary-kpi"><div class="label">Относительное изменение</div><div class="value">${formatPct(summary.deltaPct)}</div></div>
    </div>
  `;

  const excludedKeys = new Set(['latitude', 'longitude', 'settlement_id', 'oktmo_stable', 'oktmo_syn', 'is_central_place', 'central_places']);
  const rows = [
    `
      <tr>
        <th>Central_places</th>
        <td>${escapeHtml(effectiveCentralText || 'не отмечен')}</td>
      </tr>
    `,
    ...Object.entries(details)
      .filter(([, value]) => !isBlank(value))
      .filter(([key]) => !excludedKeys.has(String(key).toLowerCase()))
      .map(([key, value]) => `
        <tr>
          <th>${escapeHtml(key)}</th>
          <td>${escapeHtml(formatAny(value))}</td>
        </tr>
      `),
  ].join('');

  els.selectedSettlementMeta.innerHTML = `
    <table>
      <tbody>${rows}</tbody>
    </table>
  `;
}



function renderSelectedSettlementChart() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    renderEmptyPlot('settlementChart', 'Выберите населённый пункт на карте, чтобы построить его временной ряд.');
    return;
  }

  const traces = [];
  const actualSeries = filterSeriesToWindow(summary.actualYears, summary.actualValues);
  if (actualSeries.x.length) {
    traces.push({
      x: actualSeries.x,
      y: actualSeries.y,
      type: 'scatter',
      mode: 'lines+markers',
      name: 'Фактические данные',
      line: { color: '#1d4ed8', width: 3 },
      marker: { color: '#1d4ed8', size: 9 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    });
  }

  for (const scenarioId of getSelectedChartScenarios('settlement')) {
    const style = SCENARIO_TRACE_STYLES[scenarioId] || { color: '#334155', dash: 'solid' };
    const forecastSeries = summary.seriesByScenario[scenarioId];
    const filtered = filterSeriesToWindow(forecastSeries?.years || [], forecastSeries?.values || []);
    if (!filtered.x.length) continue;
    traces.push({
      x: filtered.x,
      y: filtered.y,
      type: 'scatter',
      mode: 'lines',
      name: `Прогноз (${getScenarioLabel(scenarioId)})`,
      line: { color: style.color, width: 3, dash: style.dash },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    });
  }

  if (!traces.length) {
    renderEmptyPlot('settlementChart', 'В пределах выбранного временного интервала данные отсутствуют.');
    return;
  }

  const title = summary.details.Settlement_full_name || summary.details.Settlement_name || 'Населённый пункт';
  const layout = buildLineLayout('Численность населения', title);
  Plotly.newPlot('settlementChart', traces, layout, buildPlotConfig(buildFilename('settlement_series')));
}

async function renderPyramid() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    state.currentPyramidBundle = null;
    renderEmptyPlot('pyramidChart', 'Выберите населённый пункт, чтобы рассчитать половозрастную пирамиду.');
    return;
  }

  const details = summary.details;
  const year = getPyramidYear();
  const scenario = getScenario();
  const oktmo = details.oktmo_syn;
  const settlementPop = toNumber(summary.byYearByScenario?.[scenario]?.get(year));

  if (!oktmo || !Number.isFinite(settlementPop) || settlementPop <= 0) {
    state.currentPyramidBundle = null;
    renderEmptyPlot('pyramidChart', 'Для выбранного пункта отсутствуют данные, необходимые для построения пирамиды.');
    return;
  }

  const alias = getMunicipalAgeAlias(scenario, year);
  await ensureRegisteredFile(alias, `./data/municipal_age/scenario=${scenario}/year=${year}.parquet`);
  const rows = await queryRows(`
    SELECT sex, age, pop
    FROM read_parquet('${alias}')
    WHERE oktmo_syn = ${sqlLiteral(oktmo)}
    ORDER BY CAST(age AS INTEGER), sex
  `);

  if (!rows.length) {
    state.currentPyramidBundle = null;
    renderEmptyPlot('pyramidChart', 'Для выбранного сочетания «населённый пункт — год — сценарий» муниципальный возрастно-половой профиль отсутствует.');
    return;
  }

  const scaledRows = scaleAgeProfile(rows, Math.round(settlementPop));
  state.currentPyramidBundle = {
    year,
    scenario,
    settlementPop: Math.round(settlementPop),
    oktmo_syn: oktmo,
    rows: scaledRows,
  };

  const ages = [...new Set(scaledRows.map((row) => row.age))].sort((a, b) => a - b);
  const male = ages.map((age) => -1 * (scaledRows.find((row) => row.age === age && row.sex === 'M')?.scaled_pop || 0));
  const female = ages.map((age) => scaledRows.find((row) => row.age === age && row.sex === 'F')?.scaled_pop || 0);
  const tickVals = ages.filter((age) => age % 5 === 0 || age >= 100);
  const tickText = tickVals.map((age) => (age >= 100 ? '100+' : String(age)));

  const traces = [
    {
      x: male,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Мужчины',
      marker: { color: '#2563eb' },
      hovertemplate: '<b>Возраст %{customdata}</b><br>Мужчины: %{text:,}<extra></extra>',
      customdata: ages.map((age) => (age >= 100 ? '100+' : age)),
      text: male.map((v) => Math.abs(v)),
    },
    {
      x: female,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Женщины',
      marker: { color: '#ec4899' },
      hovertemplate: '<b>Возраст %{customdata}</b><br>Женщины: %{x:,}<extra></extra>',
      customdata: ages.map((age) => (age >= 100 ? '100+' : age)),
    },
  ];

  const layout = {
    title: {
      text: `${details.Settlement_full_name || details.Settlement_name || 'Населённый пункт'} · ${year}`,
      font: { size: 16 },
      x: 0.02,
      xanchor: 'left',
    },
    barmode: 'relative',
    bargap: 0.05,
    margin: { l: 64, r: 36, t: 58, b: 48 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    legend: { orientation: 'h', x: 0, y: 1.08 },
    xaxis: {
      title: 'Численность',
      zeroline: true,
      zerolinewidth: 1,
      zerolinecolor: '#94a3b8',
      gridcolor: '#e2e8f0',
      tickformat: ',d',
      tickmode: 'auto',
    },
    yaxis: {
      title: 'Возраст',
      dtick: 5,
      gridcolor: '#f1f5f9',
      tickmode: 'array',
      tickvals: tickVals,
      ticktext: tickText,
    },
    annotations: [
      {
        text: `Общая численность: ${formatInteger(settlementPop)} · сценарий: ${getScenarioLabel(scenario)}`,
        xref: 'paper',
        yref: 'paper',
        x: 0,
        y: 1.16,
        xanchor: 'left',
        yanchor: 'bottom',
        showarrow: false,
        font: { size: 12, color: '#475569' },
      },
    ],
  };

  Plotly.newPlot('pyramidChart', traces, layout, buildPlotConfig(buildFilename('pyramid')));
}



function scaleAgeProfile(rows, targetTotal) {
  const cleaned = rows.map((row) => ({
    sex: String(row.sex || '').trim(),
    age: Number(row.age),
    pop: Math.max(0, toNumber(row.pop) || 0),
  })).filter((row) => Number.isFinite(row.age));

  const total = cleaned.reduce((sum, row) => sum + row.pop, 0);
  if (!Number.isFinite(total) || total <= 0 || !Number.isFinite(targetTotal) || targetTotal <= 0) {
    return cleaned.map((row) => ({ ...row, scaled_pop: 0 }));
  }

  const raw = cleaned.map((row) => (row.pop / total) * targetTotal);
  const floors = raw.map((value) => Math.floor(value));
  let diff = targetTotal - floors.reduce((sum, value) => sum + value, 0);
  const ranked = raw.map((value, index) => ({ index, frac: value - floors[index] })).sort((a, b) => b.frac - a.frac);
  for (let i = 0; i < ranked.length && diff > 0; i += 1, diff -= 1) {
    floors[ranked[i].index] += 1;
  }

  return cleaned.map((row, index) => ({ ...row, scaled_pop: floors[index] }));
}

function renderEmptyPlot(containerId, message) {
  Plotly.newPlot(containerId, [], {
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    margin: { l: 30, r: 20, t: 40, b: 30 },
    xaxis: { visible: false },
    yaxis: { visible: false },
    annotations: [{
      text: message,
      showarrow: false,
      xref: 'paper',
      yref: 'paper',
      x: 0.5,
      y: 0.5,
      font: { size: 14, color: '#475569' },
      align: 'center',
    }],
  }, { displayModeBar: false, responsive: true });
}

function buildLineLayout(yTitle, subtitle) {
  return {
    title: {
      text: subtitle,
      font: { size: 16 },
      x: 0.02,
      xanchor: 'left',
    },
    margin: { l: 64, r: 24, t: 60, b: 52 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    legend: { orientation: 'h', x: 0, y: 1.12 },
    hovermode: 'x unified',
    annotations: [
      {
        text: `Диапазон графика: ${getBaseYear()}–${getChartEndYear()}`,
        xref: 'paper',
        yref: 'paper',
        x: 0,
        y: 1.17,
        xanchor: 'left',
        yanchor: 'bottom',
        showarrow: false,
        font: { size: 12, color: '#475569' },
      },
    ],
    xaxis: {
      title: 'Год',
      tickmode: 'auto',
      gridcolor: '#e2e8f0',
      zeroline: false,
    },
    yaxis: {
      title: yTitle,
      tickformat: ',d',
      gridcolor: '#e2e8f0',
      zeroline: false,
    },
  };
}

function buildPlotConfig(filename) {
  return {
    displaylogo: false,
    responsive: true,
    toImageButtonOptions: {
      format: 'png',
      filename,
      height: 800,
      width: 1400,
      scale: 2,
    },
  };
}

async function exportMapPng() {
  try {
    const canvas = await html2canvas(document.querySelector('.map-panel'), {
      useCORS: true,
      backgroundColor: '#ffffff',
      scale: 2,
      logging: false,
    });
    canvas.toBlob((blob) => {
      if (!blob) {
        showError('Не удалось подготовить PNG карты.');
        return;
      }
      downloadBlob(blob, `${buildFilename('map')}.png`);
    });
  } catch (error) {
    handleError(new Error(`Не удалось сохранить карту в PNG: ${error?.message || error}`));
  }
}

async function exportElementPng(elementId, filename) {
  const element = document.getElementById(elementId);
  if (!element) {
    showError('Не удалось найти элемент для экспорта PNG.');
    return;
  }
  try {
    const canvas = await html2canvas(element, {
      useCORS: true,
      backgroundColor: '#ffffff',
      scale: 2,
      logging: false,
    });
    canvas.toBlob((blob) => {
      if (!blob) {
        showError('Не удалось подготовить PNG.');
        return;
      }
      downloadBlob(blob, `${filename}.png`);
    });
  } catch (error) {
    handleError(new Error(`Не удалось сохранить PNG: ${error?.message || error}`));
  }
}

function exportPlotPng(containerId, filename) {
  const node = document.getElementById(containerId);
  if (!node || !node.data || !node.data.length) {
    showError('Для экспорта PNG сначала нужно построить соответствующий график.');
    return;
  }
  Plotly.downloadImage(node, { format: 'png', filename, width: 1400, height: 800, scale: 2 });
}

function exportCurrentSelectionXlsx() {
  if (!state.currentMapData.length) {
    showError('Нет данных для экспорта: текущая выборка пуста.');
    return;
  }
  const rows = state.currentMapData.map((row) => ({
    settlement_id: row.settlement_id,
    Region: row.Region,
    Municipality: row.Municipality,
    Settlement_full_name: row.Settlement_full_name,
    Settlement_name: row.Settlement_name,
    Type_settlement: row.Type_settlement,
    Latitude: row.Latitude,
    Longitude: row.Longitude,
    pop: row.pop,
    base_pop: row.base_pop,
    delta_abs: row.delta_abs,
    delta_pct: row.delta_pct,
    is_central_place: row.is_central_place,
    Central_places: row.Central_places,
    central_group_label: row.central_group_label,
  }));
  exportWorkbook([{ name: 'selection', rows }], `${buildFilename('selection')}.xlsx`);
}

function exportAggregateXlsx() {
  if (!state.currentAggregateSeries) {
    showError('Сначала нужно построить агрегированный временной ряд.');
    return;
  }
  const rows = [];
  const actualSeries = filterSeriesToWindow(state.currentAggregateSeries.actualYears, state.currentAggregateSeries.actualValues);
  actualSeries.x.forEach((year, index) => {
    rows.push({ year, population: actualSeries.y[index], source: 'Фактические данные' });
  });
  for (const scenarioId of getSelectedChartScenarios('aggregate')) {
    const scenarioSeries = state.currentAggregateSeries.forecastsByScenario[scenarioId];
    const filtered = filterSeriesToWindow(scenarioSeries?.years || [], scenarioSeries?.values || []);
    filtered.x.forEach((year, index) => {
      rows.push({ year, population: filtered.y[index], source: `Прогноз (${getScenarioLabel(scenarioId)})` });
    });
  }
  exportWorkbook([{ name: 'aggregate', rows }], `${buildFilename('aggregate')}.xlsx`);
}

function exportSettlementSeriesXlsx() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    showError('Сначала выберите населённый пункт.');
    return;
  }
  const rows = [];
  const actualSeries = filterSeriesToWindow(summary.actualYears, summary.actualValues);
  actualSeries.x.forEach((year, index) => {
    rows.push({ year, population: actualSeries.y[index], source: 'Фактические данные' });
  });
  for (const scenarioId of getSelectedChartScenarios('settlement')) {
    const filtered = filterSeriesToWindow(summary.seriesByScenario[scenarioId]?.years || [], summary.seriesByScenario[scenarioId]?.values || []);
    filtered.x.forEach((year, index) => {
      rows.push({ year, population: filtered.y[index], source: `Прогноз (${getScenarioLabel(scenarioId)})` });
    });
  }
  exportWorkbook([{ name: 'settlement_series', rows }], `${buildFilename('settlement_series')}.xlsx`);
}

function exportPyramidXlsx() {
  if (!state.currentPyramidBundle?.rows?.length) {
    showError('Сначала постройте половозрастную пирамиду.');
    return;
  }
  const rows = state.currentPyramidBundle.rows.map((row) => ({
    age: row.age >= 100 ? '100+' : row.age,
    sex: row.sex,
    scaled_population: row.scaled_pop,
    municipal_profile_population: row.pop,
    scenario: getScenarioLabel(state.currentPyramidBundle.scenario),
    year: state.currentPyramidBundle.year,
  }));
  exportWorkbook([{ name: 'pyramid', rows }], `${buildFilename('pyramid')}.xlsx`);
}

function buildChartExportRows(actualYears, actualValues, forecastMap, selectedScenarios, valueField) {
  const rows = [];
  const actualSeries = filterSeriesToWindow(actualYears || [], actualValues || []);
  actualSeries.x.forEach((year, index) => {
    rows.push({ year, [valueField]: actualSeries.y[index], source: 'Фактические данные' });
  });
  selectedScenarios.forEach((scenarioId) => {
    const filtered = filterSeriesToWindow(forecastMap?.[scenarioId]?.years || [], forecastMap?.[scenarioId]?.values || []);
    filtered.x.forEach((year, index) => {
      rows.push({ year, [valueField]: filtered.y[index], source: `Прогноз (${getScenarioLabel(scenarioId)})` });
    });
  });
  return rows;
}

function exportFertilityXlsx() {
  const bundle = state.currentComponentSeries?.fertility;
  if (!bundle) {
    showError('Сначала постройте блок рождаемости.');
    return;
  }
  const selectedScenarios = getSelectedChartScenarios('fertility');
  const birthsForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.births }]));
  const tfrForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.tfr }]));
  exportWorkbook([
    { name: 'births', rows: buildChartExportRows(bundle.actualYears, bundle.birthsActual, birthsForecast, selectedScenarios, 'births_total') },
    { name: 'tfr', rows: buildChartExportRows(bundle.actualYears, bundle.tfrActual, tfrForecast, selectedScenarios, 'tfr') },
  ], `${buildFilename('fertility')}.xlsx`);
}

function exportMortalityXlsx() {
  const bundle = state.currentComponentSeries?.mortality;
  if (!bundle) {
    showError('Сначала постройте блок смертности.');
    return;
  }
  const selectedScenarios = getSelectedChartScenarios('mortality');
  const deathsForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.deaths }]));
  const cdrForecast = Object.fromEntries(Object.entries(bundle.forecastsByScenario || {}).map(([scenarioId, data]) => [scenarioId, { years: data.years, values: data.cdr }]));
  exportWorkbook([
    { name: 'deaths', rows: buildChartExportRows(bundle.actualYears, bundle.deathsActual, deathsForecast, selectedScenarios, 'deaths_total') },
    { name: 'cdr', rows: buildChartExportRows(bundle.actualYears, bundle.cdrActual, cdrForecast, selectedScenarios, 'cdr') },
  ], `${buildFilename('mortality')}.xlsx`);
}

function exportMigrationXlsx() {
  const bundle = state.currentComponentSeries?.migration;
  if (!bundle) {
    showError('Сначала постройте блок миграции.');
    return;
  }
  const rows = [];
  const actualMeta = [
    { key: 'migInActual', label: 'Фактический приток' },
    { key: 'migOutActual', label: 'Фактический отток' },
    { key: 'migNetActual', label: 'Фактическое сальдо' },
  ];
  actualMeta.forEach((meta) => {
    const filtered = filterSeriesToWindow(bundle.actualYears, bundle[meta.key]);
    filtered.x.forEach((year, index) => {
      rows.push({ year, value: filtered.y[index], source: meta.label });
    });
  });

  const selectedScenarios = getSelectedChartScenarios('migration');
  selectedScenarios.forEach((scenarioId) => {
    const forecast = bundle.forecastsByScenario?.[scenarioId];
    [
      ['migIn', 'Прогноз притока'],
      ['migOut', 'Прогноз оттока'],
      ['migNet', 'Прогноз сальдо'],
    ].forEach(([field, label]) => {
      const filtered = filterSeriesToWindow(forecast?.years || [], forecast?.[field] || []);
      filtered.x.forEach((year, index) => {
        rows.push({ year, value: filtered.y[index], source: `${label} (${getScenarioLabel(scenarioId)})` });
      });
    });
  });

  const inflowRows = state.currentMigrationAgeBundle?.rows?.map((row) => ({ age_band: row.age_band, sex: row.sex, inflow: row.inflow, year: state.currentMigrationAgeBundle.year, scenario: getScenarioLabel(state.currentMigrationAgeBundle.scenario) })) || [];
  const outflowRows = state.currentMigrationAgeBundle?.rows?.map((row) => ({ age_band: row.age_band, sex: row.sex, outflow: row.outflow, year: state.currentMigrationAgeBundle.year, scenario: getScenarioLabel(state.currentMigrationAgeBundle.scenario) })) || [];

  exportWorkbook([
    { name: 'migration_totals', rows },
    { name: 'migration_inflow_age', rows: inflowRows },
    { name: 'migration_outflow_age', rows: outflowRows },
  ], `${buildFilename('migration')}.xlsx`);
}

function exportSelectedDetailsXlsx() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    showError('Сначала выберите населённый пункт.');
    return;
  }
  const effectiveCentralText = getEffectiveCentralPlaceText(summary.details.settlement_id);
  const excludedKeys = new Set(['latitude', 'longitude', 'settlement_id', 'oktmo_stable', 'oktmo_syn', 'is_central_place', 'central_places']);
  const rows = [
    { field: 'Central_places', value: effectiveCentralText || 'не отмечен' },
    ...Object.entries(summary.details)
      .filter(([, value]) => !isBlank(value))
      .filter(([field]) => !excludedKeys.has(String(field).toLowerCase()))
      .map(([field, value]) => ({ field, value: formatAny(value) })),
  ];
  exportWorkbook([{ name: 'settlement_details', rows }], `${buildFilename('settlement_details')}.xlsx`);
}

function exportWorkbook(sheets, filename) {
  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name.slice(0, 31));
  });
  XLSX.writeFile(workbook, filename);
}



function filterSeriesToWindow(years, values) {
  const [startYear, endYear] = getChartRangeYears();
  const x = [];
  const y = [];
  years.forEach((year, index) => {
    const value = values[index];
    if (year < startYear || year > endYear) return;
    if (!Number.isFinite(toNumber(value))) return;
    x.push(year);
    y.push(toNumber(value));
  });
  return { x, y };
}

function buildFilename(stem) {
  return `${stem}_${getScenario()}_${getMapYear()}_base${getBaseYear()}_end${getChartEndYear()}`.replace(/[^a-zA-Z0-9_\-]+/g, '_');
}

function getScenario() {
  return state.ui.activeScenario || state.config.defaultScenario;
}

function getChartRangeYears() {
  const raw = els.chartRangeSlider.noUiSlider.get();
  const indices = (Array.isArray(raw) ? raw : [raw]).map((value) => Math.round(Number(value)));
  return [state.config.allYears[indices[0]], state.config.allYears[indices[1]]];
}

function getBaseYear() {
  return getChartRangeYears()[0];
}

function getChartEndYear() {
  return getChartRangeYears()[1];
}

function getMapYear() {
  const index = Math.round(Number(els.mapYearSlider.noUiSlider.get()));
  return state.config.mapYears[index];
}

function getPyramidYear() {
  const index = Math.round(Number(els.pyramidYearSlider.noUiSlider.get()));
  return state.config.pyramidYears[index];
}

function getMigrationPyramidYear() {
  const index = Math.round(Number(els.migrationYearSlider.noUiSlider.get()));
  return state.config.componentYears.migrationPyramid[index];
}

function getSelectedChartScenarios(kind) {
  const groups = {
    aggregate: [els.aggregateScenarioWithMIG, els.aggregateScenarioNoMIG],
    settlement: [els.settlementScenarioWithMIG, els.settlementScenarioNoMIG],
    fertility: [els.fertilityScenarioWithMIG, els.fertilityScenarioNoMIG],
    mortality: [els.mortalityScenarioWithMIG, els.mortalityScenarioNoMIG],
    migration: [els.migrationScenarioWithMIG, els.migrationScenarioNoMIG],
  };
  const group = groups[kind] || groups.settlement;
  const ids = [];
  group.forEach((input) => {
    if (input?.checked) ids.push(input.dataset.scenario);
  });
  return ids.length ? ids : [getScenario()];
}

function getForecastAlias(scenario) {
  return `settlement_forecast_wide_${scenario}.parquet`;
}

function getMunicipalAgeAlias(scenario, year) {
  return `municipal_age_${scenario}_${year}.parquet`;
}

function getMunicipalMigrationAgeAlias(scenario, year) {
  return `municipal_migration_age_${scenario}_${year}.parquet`;
}



function populationExpression(year, indexAlias = 'i', forecastAlias = 'w') {
  if (year === 1989) return `CAST(${indexAlias}.Population_1989 AS DOUBLE)`;
  if (year === 2002) return `CAST(${indexAlias}.Population_2002 AS DOUBLE)`;
  if (year === 2010) return `CAST(${indexAlias}.Population_2010 AS DOUBLE)`;
  if (year === 2021) return `COALESCE(CAST(${indexAlias}.Population_2021 AS DOUBLE), CAST(${forecastAlias}.pop_2021 AS DOUBLE))`;
  return `CAST(${forecastAlias}.pop_${year} AS DOUBLE)`;
}

function buildWhereClause(alias = 'i') {
  const regions = normalizeMultiValue(state.ui.regionTom.getValue());
  const municipalityKeys = normalizeMultiValue(state.ui.municipalityTom.getValue());
  const clauses = [];
  if (regions.length) {
    clauses.push(`${alias}.Region IN (${regions.map(sqlLiteral).join(', ')})`);
  }
  if (municipalityKeys.length) {
    const tuples = municipalityKeys.map((key) => {
      const [region, municipality] = String(key).split('|||');
      return `(${sqlLiteral(region)}, ${sqlLiteral(municipality)})`;
    });
    clauses.push(`(${alias}.Region, ${alias}.Municipality) IN (${tuples.join(', ')})`);
  }
  return clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
}

function normalizeMultiValue(value) {
  if (Array.isArray(value)) return value.filter(Boolean);
  if (value == null || value === '') return [];
  return [value];
}

function setSliderToYear(element, yearArray, year) {
  const index = yearArray.indexOf(year);
  if (index >= 0) {
    element.noUiSlider.set(index);
  }
}

function syncPyramidYearIfNeeded(forceRender = false) {
  if (!els.autoPyramidToggle.checked) return;
  const target = Math.max(2021, getMapYear());
  if (getPyramidYear() !== target) {
    setSliderToYear(els.pyramidYearSlider, state.config.pyramidYears, target);
  } else if (forceRender) {
    renderPyramid().catch(handleError);
  }
}

function getScenarioLabel(scenarioId) {
  return state.config.scenarios.find((scenario) => scenario.id === scenarioId)?.label || scenarioId;
}

function describeCurrentFilter() {
  const selectedRegions = normalizeMultiValue(state.ui.regionTom.getValue());
  const selectedMunicipalities = normalizeMultiValue(state.ui.municipalityTom.getValue());
  if (!selectedRegions.length && !selectedMunicipalities.length) {
    return 'Вся Россия';
  }
  const parts = [];
  if (selectedRegions.length) parts.push(`субъекты РФ: ${selectedRegions.length}`);
  if (selectedMunicipalities.length) parts.push(`муниципалитеты: ${selectedMunicipalities.length}`);
  return parts.join(' · ');
}

function countCentralRepresentativesForCurrentFilter() {
  if (!state.centralPlaceRepresentatives.length) {
    return state.currentMapData.filter((row) => row.is_central_place).length;
  }
  const selectedRegions = new Set(normalizeMultiValue(state.ui.regionTom.getValue()));
  const selectedMunicipalityKeys = new Set(normalizeMultiValue(state.ui.municipalityTom.getValue()));

  return state.centralPlaceRepresentatives.filter((row) => {
    if (selectedRegions.size && !selectedRegions.has(row.Region)) return false;
    if (selectedMunicipalityKeys.size && !selectedMunicipalityKeys.has(`${row.Region}|||${row.Municipality}`)) return false;
    return true;
  }).length;
}

function getEffectiveCentralPlaceText(settlementId) {
  return state.centralPlaceTextById.get(Number(settlementId)) || null;
}

function isEffectiveCentralPlace(settlementId) {
  return state.centralPlaceRepresentativeIds.has(Number(settlementId));
}

function getColorNegativeThreshold() {
  const parsed = Number(els.deltaNegativeInput.value);
  if (!Number.isFinite(parsed) || parsed >= 0) {
    els.deltaNegativeInput.value = '-10';
    return -10;
  }
  return parsed;
}

function getColorPositiveThreshold() {
  const parsed = Number(els.deltaPositiveInput.value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    els.deltaPositiveInput.value = '10';
    return 10;
  }
  return parsed;
}

function populationClassFor(pop) {
  if (!Number.isFinite(pop) || pop <= 0) return POPULATION_CLASSES[0];
  return POPULATION_CLASSES.find((item) => pop >= item.min && pop <= item.max) || POPULATION_CLASSES[POPULATION_CLASSES.length - 1];
}

function radiusForPopulation(pop) {
  if (!Number.isFinite(pop) || pop <= 0) return 0;
  return populationClassFor(pop).radius;
}

function colorForDelta(deltaPct) {
  if (!Number.isFinite(deltaPct)) return [148, 163, 184, 170];
  const min = getColorNegativeThreshold();
  const max = getColorPositiveThreshold();

  if (deltaPct <= 0) {
    const t = clamp((deltaPct - min) / (0 - min), 0, 1);
    return lerpColor([220, 38, 38, 218], [250, 204, 21, 210], t);
  }
  const t = clamp(deltaPct / max, 0, 1);
  return lerpColor([250, 204, 21, 210], [22, 163, 74, 222], t);
}

function buildCentralPlaceMeta(value) {
  if (isBlank(value)) return null;
  const text = String(value).replace(/\s+/g, ' ').trim().toLowerCase();
  let groupId = 'other';
  if (text.includes('ядром городской агломерации')) {
    groupId = 'aggl_core';
  } else if (text.includes('входит в состав городской агломерации')) {
    groupId = 'agglomeration';
  } else if (text.includes('основным центром предоставления социальных услуг')) {
    groupId = 'services';
  } else if (text.includes('инвестиционные проекты')) {
    groupId = 'investment';
  } else if (text.includes('национальную безопасность') || text.includes('государственной границы')) {
    groupId = 'security';
  } else if (text.includes('критически важной инфраструктуры')) {
    groupId = 'infrastructure';
  } else if (text.includes('зато')) {
    groupId = 'zato';
  } else if (text.includes('наукограда')) {
    groupId = 'science';
  } else if (text.includes('административного центра субъекта')) {
    groupId = 'regional_capital';
  }
  const group = CENTRAL_GROUP_BY_ID.get(groupId) || CENTRAL_GROUP_BY_ID.get('other');
  return {
    ...group,
    sourceLabel: String(value).trim(),
  };
}

function lerpColor(a, b, t) {
  return a.map((value, index) => Math.round(value + (b[index] - value) * t));
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function formatInteger(value) {
  if (!Number.isFinite(value)) return '—';
  return Math.round(value).toLocaleString('ru-RU');
}

function formatSigned(value) {
  if (!Number.isFinite(value)) return '—';
  return `${value > 0 ? '+' : ''}${Math.round(value).toLocaleString('ru-RU')}`;
}

function formatPct(value) {
  if (!Number.isFinite(value)) return '—';
  return `${value > 0 ? '+' : ''}${value.toLocaleString('ru-RU', { maximumFractionDigits: 2, minimumFractionDigits: 0 })}%`;
}

function formatSignedPctThreshold(value) {
  if (!Number.isFinite(value)) return '—';
  return `${value > 0 ? '+' : ''}${Math.round(value).toLocaleString('ru-RU')}%`;
}

function pluralYears(value) {
  const abs = Math.abs(Number(value)) % 100;
  const last = abs % 10;
  if (abs > 10 && abs < 20) return 'лет';
  if (last === 1) return 'год';
  if (last >= 2 && last <= 4) return 'года';
  return 'лет';
}

function isLikelyYear(value) {
  return Number.isInteger(value) && value >= 1800 && value <= 2200;
}

function formatAny(value) {
  if (value == null || value === '') return '—';
  if (typeof value === 'number') {
    if (isLikelyYear(value)) return String(value);
    if (Number.isInteger(value)) return value.toLocaleString('ru-RU');
    return value.toLocaleString('ru-RU', { maximumFractionDigits: 6 });
  }
  if (typeof value === 'boolean') return value ? 'Да' : 'Нет';
  return String(value);
}

function isBlank(value) {
  return value == null || value === '' || (typeof value === 'string' && value.trim() === '') || String(value) === 'NaN';
}

function sqlLiteral(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function toNumber(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value === 'bigint') return Number(value);
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeArrowValue(value) {
  if (typeof value === 'bigint') return Number(value);
  if (Array.isArray(value)) return value.map((item) => normalizeArrowValue(item));
  if (value && typeof value === 'object') {
    const out = {};
    for (const [key, innerValue] of Object.entries(value)) {
      out[key] = normalizeArrowValue(innerValue);
    }
    return out;
  }
  return value;
}

async function ensureRegisteredFile(alias, relativePath) {
  if (state.registeredFiles.has(alias)) return;
  const absoluteUrl = new URL(relativePath, window.location.href).toString();
  await state.db.registerFileURL(alias, absoluteUrl, duckdb.DuckDBDataProtocol.HTTP, false);
  state.registeredFiles.add(alias);
}

async function queryRows(sql) {
  const result = await state.conn.query(sql);
  return result.toArray().map((row) => normalizeArrowValue(row.toJSON()));
}

async function queryExec(sql) {
  await state.conn.query(sql);
}

function setLoading(isVisible, title = 'Загрузка…', subtitle = '') {
  els.loadingTitle.textContent = title;
  els.loadingSubtitle.textContent = subtitle;
  els.loadingOverlay.classList.toggle('visible', Boolean(isVisible));
}

function showError(message) {
  els.errorBanner.textContent = message;
  els.errorBanner.classList.remove('hidden');
}

function clearError() {
  els.errorBanner.classList.add('hidden');
  els.errorBanner.textContent = '';
}

function handleError(error) {
  console.error(error);
  showError(error?.message || String(error));
  setLoading(false);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}