import * as duckdb from 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm/+esm';

const els = {
  scenarioSelect: document.getElementById('scenarioSelect'),
  mapYearSelect: document.getElementById('mapYearSelect'),
  baseYearSelect: document.getElementById('baseYearSelect'),
  pyramidYearSelect: document.getElementById('pyramidYearSelect'),
  regionSelect: document.getElementById('regionSelect'),
  municipalitySelect: document.getElementById('municipalitySelect'),
  centralPlacesToggle: document.getElementById('centralPlacesToggle'),
  autoPyramidToggle: document.getElementById('autoPyramidToggle'),
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
  openMethodologyBtn: document.getElementById('openMethodologyBtn'),
  resetMapViewBtn: document.getElementById('resetMapViewBtn'),
  loadingOverlay: document.getElementById('loadingOverlay'),
  loadingTitle: document.getElementById('loadingTitle'),
  loadingSubtitle: document.getElementById('loadingSubtitle'),
  errorBanner: document.getElementById('errorBanner'),
  mapCaption: document.getElementById('mapCaption'),
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
  registeredFiles: new Set(),
  filterOptions: {
    regions: [],
    municipalities: [],
    municipalitiesByRegion: new Map(),
  },
  ui: {
    regionTom: null,
    municipalityTom: null,
    hoverPopup: null,
    refreshTimer: null,
    refreshToken: 0,
    mapReady: null,
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
}

function initControls() {
  fillSelect(
    els.scenarioSelect,
    state.config.scenarios.map((scenario) => ({ value: scenario.id, text: scenario.label })),
    state.config.defaultScenario,
  );
  fillSelect(
    els.mapYearSelect,
    state.config.mapYears.map((year) => ({ value: String(year), text: String(year) })),
    String(state.config.defaultMapYear),
  );
  fillSelect(
    els.baseYearSelect,
    state.config.mapYears.map((year) => ({ value: String(year), text: String(year) })),
    String(state.config.defaultBaseYear),
  );
  fillSelect(
    els.pyramidYearSelect,
    state.config.pyramidYears.map((year) => ({ value: String(year), text: String(year) })),
    String(Math.max(state.config.defaultBaseYear, 2021)),
  );

  state.ui.regionTom = new TomSelect(els.regionSelect, {
    plugins: ['remove_button'],
    maxItems: null,
    create: false,
    persist: false,
    closeAfterSelect: false,
    placeholder: 'Все субъекты РФ',
  });

  state.ui.municipalityTom = new TomSelect(els.municipalitySelect, {
    plugins: ['remove_button'],
    maxItems: null,
    create: false,
    persist: false,
    closeAfterSelect: false,
    placeholder: 'Все муниципальные образования',
  });
}

function bindEvents() {
  els.scenarioSelect.addEventListener('change', () => scheduleRefresh());
  els.mapYearSelect.addEventListener('change', () => {
    syncPyramidYearIfNeeded();
    scheduleRefresh();
  });
  els.baseYearSelect.addEventListener('change', () => scheduleRefresh());
  els.pyramidYearSelect.addEventListener('change', () => renderPyramid().catch(handleError));
  els.centralPlacesToggle.addEventListener('change', () => updateMapLayers());
  els.autoPyramidToggle.addEventListener('change', () => syncPyramidYearIfNeeded());

  state.ui.regionTom.on('change', () => {
    updateMunicipalityOptions();
  });
  state.ui.municipalityTom.on('change', () => {
    // explicit apply button is supported, but refresh is lightweight enough for interactive work
  });

  els.applyFiltersBtn.addEventListener('click', () => scheduleRefresh());
  els.clearFiltersBtn.addEventListener('click', () => clearFilters());
  els.resetMapViewBtn.addEventListener('click', () => resetMapView());
  els.openMethodologyBtn.addEventListener('click', () => els.methodologyDialog.showModal());

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
  await queryExec(`CREATE OR REPLACE TABLE settlement_index AS SELECT * FROM read_parquet('settlement_index.parquet')`);
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
    const municipalities = municipalityRows.filter((row) => row.Region === region).map((row) => row.Municipality);
    state.filterOptions.municipalitiesByRegion.set(region, municipalities);
  }

  state.ui.regionTom.clearOptions();
  state.ui.regionTom.addOptions(state.filterOptions.regions.map((region) => ({ value: region, text: region })));
  state.ui.regionTom.refreshOptions(false);
  updateMunicipalityOptions();
}

function updateMunicipalityOptions() {
  const selectedRegions = normalizeMultiValue(state.ui.regionTom.getValue());
  const selectedMunicipalities = new Set(normalizeMultiValue(state.ui.municipalityTom.getValue()));

  let options = state.filterOptions.municipalities;
  if (selectedRegions.length) {
    const allowedMunicipalities = new Set();
    for (const region of selectedRegions) {
      for (const municipality of state.filterOptions.municipalitiesByRegion.get(region) || []) {
        allowedMunicipalities.add(`${region}|||${municipality}`);
      }
    }
    options = state.filterOptions.municipalities.filter((row) => allowedMunicipalities.has(`${row.Region}|||${row.Municipality}`));
  }

  state.ui.municipalityTom.clear(true);
  state.ui.municipalityTom.clearOptions();
  state.ui.municipalityTom.addOptions(options.map((row) => ({
    value: `${row.Region}|||${row.Municipality}`,
    text: `${row.Municipality} — ${row.Region}`,
  })));
  const stillValid = [...selectedMunicipalities].filter((municipalityKey) => options.some((row) => `${row.Region}|||${row.Municipality}` === municipalityKey));
  state.ui.municipalityTom.setValue(stillValid, true);
  state.ui.municipalityTom.refreshOptions(false);
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
    center: [95, 63],
    zoom: 2.35,
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
      resolve();
    });
  });
}

function resetMapView() {
  if (!state.map) return;
  state.map.easeTo({ center: [95, 63], zoom: 2.35, duration: 700 });
}

function scheduleRefresh() {
  window.clearTimeout(state.ui.refreshTimer);
  state.ui.refreshTimer = window.setTimeout(() => {
    refreshAll().catch(handleError);
  }, 220);
}

function clearFilters() {
  state.ui.regionTom.clear(true);
  updateMunicipalityOptions();
  state.ui.municipalityTom.clear(true);
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

  state.currentAggregateSeries = await loadAggregateSeries();
  if (token !== state.ui.refreshToken) return;
  renderAggregateChart();

  if (state.selectedSettlementId != null) {
    await refreshSelectedSettlement();
    if (token !== state.ui.refreshToken) return;
  } else {
    renderSelectedSettlementCard();
    renderSelectedSettlementChart();
    renderPyramid();
  }

  setLoading(false);
}

async function loadCurrentMapData() {
  const scenario = getScenario();
  const forecastAlias = getForecastAlias(scenario);
  const mapYear = getMapYear();
  const whereClause = buildWhereClause('i');
  const popExpr = populationExpression(mapYear, 'i', 'w');
  const baseExpr = populationExpression(getBaseYear(), 'i', 'w');

  const needsForecastJoin = mapYear >= 2021 || getBaseYear() >= 2021;
  const joinSql = needsForecastJoin ? `LEFT JOIN read_parquet('${forecastAlias}') w USING (settlement_id)` : '';

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
      CAST(i.is_central_place AS BOOLEAN) AS is_central_place,
      i.Central_places,
      ${popExpr} AS pop,
      ${baseExpr} AS base_pop,
      CAST(i.Population_1989 AS DOUBLE) AS Population_1989,
      CAST(i.Population_2002 AS DOUBLE) AS Population_2002,
      CAST(i.Population_2010 AS DOUBLE) AS Population_2010,
      CAST(i.Population_2021 AS DOUBLE) AS Population_2021
    FROM settlement_index i
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
        pop,
        base_pop: basePop,
        delta_abs: deltaAbs,
        delta_pct: deltaPct,
      };
    })
    .filter((row) => Number.isFinite(row.Latitude) && Number.isFinite(row.Longitude) && Number.isFinite(row.pop) && row.pop > 0);
}

async function loadAggregateSeries() {
  const scenario = getScenario();
  const forecastAlias = getForecastAlias(scenario);
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
  const forecastSql = `
    SELECT
      ${forecastSelect}
    FROM read_parquet('${forecastAlias}') w
    JOIN settlement_index i USING (settlement_id)
    ${whereClause}
  `;

  const actualRows = await queryRows(actualSql);
  const forecastRows = await queryRows(forecastSql);
  const actual = actualRows[0] || {};
  const forecast = forecastRows[0] || {};

  return {
    actualYears: state.config.censusYears,
    actualValues: state.config.censusYears.map((year) => toNumber(actual[`pop_${year}`])),
    forecastYears: state.config.forecastYears,
    forecastValues: state.config.forecastYears.map((year) => toNumber(forecast[`pop_${year}`])),
  };
}

function renderAggregateChart() {
  if (!state.currentAggregateSeries) {
    renderEmptyPlot('aggregateChart', 'Нет данных для текущего набора фильтров.');
    return;
  }

  const traces = [
    {
      x: state.currentAggregateSeries.actualYears,
      y: state.currentAggregateSeries.actualValues,
      type: 'scatter',
      mode: 'lines+markers',
      name: 'Фактические данные',
      line: { color: '#1d4ed8', width: 3 },
      marker: { color: '#1d4ed8', size: 9 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    },
    {
      x: state.currentAggregateSeries.forecastYears,
      y: state.currentAggregateSeries.forecastValues,
      type: 'scatter',
      mode: 'lines',
      name: `Прогноз (${getScenarioLabel(getScenario())})`,
      line: { color: '#d97706', width: 3 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    },
  ];

  const layout = buildLineLayout('Численность населения', describeCurrentFilter());
  Plotly.newPlot('aggregateChart', traces, layout, buildPlotConfig(buildFilename('aggregate')));
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
  const avgDelta = validDeltaRows.length ? validDeltaRows.reduce((sum, row) => sum + row.delta_pct, 0) / validDeltaRows.length : null;
  const centralCount = state.currentMapData.filter((row) => row.is_central_place).length;

  els.kpiSettlements.textContent = formatInteger(settlements);
  els.kpiFilterNote.textContent = describeCurrentFilter();
  els.kpiPopulation.textContent = formatInteger(totalPopulation);
  els.kpiPopulationNote.textContent = `Год: ${getMapYear()} · сценарий: ${getScenarioLabel(getScenario())}`;
  els.kpiDelta.textContent = formatPct(avgDelta);
  els.kpiDeltaNote.textContent = `Относительно ${getBaseYear()} года`;
  els.kpiCentral.textContent = formatInteger(centralCount);
  els.kpiCentralNote.textContent = `Видимы на карте: ${els.centralPlacesToggle.checked ? 'да' : 'слой отключён'}`;
}

function renderMapCaption() {
  els.mapCaption.textContent = `Год на карте: ${getMapYear()}. Базовый год для динамики: ${getBaseYear()}. Размер окружности пропорционален численности населения, цвет отражает изменение численности. Фильтр: ${describeCurrentFilter()}.`;
}

function updateMapLayers() {
  if (!state.deckOverlay) return;

  const selectedId = state.selectedSettlementId;
  const centralVisible = els.centralPlacesToggle.checked;

  const layers = [
    new deck.ScatterplotLayer({
      id: 'settlements-main',
      data: state.currentMapData,
      pickable: true,
      opacity: 0.78,
      stroked: true,
      filled: true,
      radiusUnits: 'pixels',
      radiusScale: 1,
      lineWidthMinPixels: 0.4,
      getPosition: (d) => [d.Longitude, d.Latitude],
      getRadius: (d) => radiusForPopulation(d.pop),
      getFillColor: (d) => colorForDelta(d.delta_pct),
      getLineColor: (d) => (d.settlement_id === selectedId ? [255, 255, 255, 240] : [255, 255, 255, 120]),
      getLineWidth: (d) => (d.settlement_id === selectedId ? 2 : 0.5),
      autoHighlight: true,
      highlightColor: [17, 24, 39, 220],
      onHover: handleMapHover,
      onClick: (info) => {
        if (info.object) {
          selectSettlement(info.object.settlement_id, false).catch(handleError);
        }
      },
      updateTriggers: {
        getLineColor: [selectedId],
        getLineWidth: [selectedId],
      },
    }),
    new deck.ScatterplotLayer({
      id: 'central-layer',
      data: centralVisible ? state.currentMapData.filter((d) => d.is_central_place) : [],
      pickable: false,
      stroked: true,
      filled: false,
      radiusUnits: 'pixels',
      radiusScale: 1,
      lineWidthMinPixels: 2,
      getPosition: (d) => [d.Longitude, d.Latitude],
      getRadius: (d) => Math.min(radiusForPopulation(d.pop) + 2.2, 28),
      getLineColor: [250, 204, 21, 235],
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

  const html = `
    <div class="popup-card">
      <strong>${escapeHtml(object.Settlement_full_name || object.Settlement_name || 'Населённый пункт')}</strong><br>
      ${escapeHtml(object.Type_settlement || '')}${object.Type_settlement ? ' · ' : ''}${escapeHtml(object.Municipality || '')}<br>
      ${escapeHtml(object.Region || '')}<br>
      <span>Численность (${getMapYear()}): ${formatInteger(object.pop)}</span><br>
      <span>Изменение к ${getBaseYear()} году: ${formatPct(object.delta_pct)}</span>
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
  if (selected) {
    if (flyTo && state.map) {
      state.map.flyTo({ center: [selected.Longitude, selected.Latitude], zoom: Math.max(state.map.getZoom(), 6.2), duration: 700 });
    }
  }
  updateMapLayers();
}

async function refreshSelectedSettlement() {
  const settlementId = state.selectedSettlementId;
  if (settlementId == null) {
    renderSelectedSettlementCard();
    renderSelectedSettlementChart();
    renderPyramid();
    return;
  }

  const details = await loadSettlementDetails(settlementId);
  const forecastRow = await loadSettlementForecastRow(settlementId, getScenario());
  const summary = buildSelectedSettlementSummary(details, forecastRow);
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

function buildSelectedSettlementSummary(details, forecastRow) {
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

  const selectedYear = getMapYear();
  const baseYear = getBaseYear();
  const selectedPop = byYear.get(selectedYear) ?? null;
  const basePop = byYear.get(baseYear) ?? null;
  const deltaAbs = Number.isFinite(selectedPop) && Number.isFinite(basePop) ? selectedPop - basePop : null;
  const deltaPct = Number.isFinite(basePop) && basePop > 0 && Number.isFinite(deltaAbs) ? (deltaAbs / basePop) * 100 : null;

  return {
    details,
    forecastRow,
    actualYears,
    actualValues,
    forecastYears,
    forecastValues,
    byYear,
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
    return;
  }

  const details = summary.details;
  els.selectedSettlementSummary.className = 'selection-summary';
  els.selectionHint.textContent = `Выбранный населённый пункт связан с муниципальным образованием через oktmo_syn = ${details.oktmo_syn || '—'}.`;

  els.selectedSettlementSummary.innerHTML = `
    <div>
      <div style="color:#475569;font-size:0.92rem;">${escapeHtml(details.Type_settlement || 'Тип не указан')} · ${escapeHtml(details.Municipality || 'Муниципалитет не указан')}</div>
      <h3 style="margin:0.35rem 0 0;font-size:1.25rem;">${escapeHtml(details.Settlement_full_name || details.Settlement_name || 'Без названия')}</h3>
      <div style="margin-top:0.35rem;color:#475569;">${escapeHtml(details.Region || 'Регион не указан')}</div>
      <div style="margin-top:0.6rem;color:#0f172a;">${details.Central_places ? `Статус центрального / опорного пункта: <strong>${escapeHtml(String(details.Central_places))}</strong>.` : 'Статус центрального / опорного пункта не отмечен.'}</div>
    </div>
    <div class="selection-grid">
      <div class="summary-kpi"><div class="label">Численность (${getMapYear()})</div><div class="value">${formatInteger(summary.selectedPop)}</div></div>
      <div class="summary-kpi"><div class="label">Базовый год (${getBaseYear()})</div><div class="value">${formatInteger(summary.basePop)}</div></div>
      <div class="summary-kpi"><div class="label">Абсолютное изменение</div><div class="value">${formatSigned(summary.deltaAbs)}</div></div>
      <div class="summary-kpi"><div class="label">Относительное изменение</div><div class="value">${formatPct(summary.deltaPct)}</div></div>
    </div>
  `;

  const excludedKeys = new Set(['latitude', 'longitude']);
  const rows = Object.entries(details)
    .filter(([, value]) => !isBlank(value))
    .filter(([key]) => !excludedKeys.has(key))
    .map(([key, value]) => `
      <tr>
        <th>${escapeHtml(key)}</th>
        <td>${escapeHtml(formatAny(value))}</td>
      </tr>
    `)
    .join('');

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

  const traces = [
    {
      x: summary.actualYears,
      y: summary.actualValues,
      type: 'scatter',
      mode: 'lines+markers',
      name: 'Фактические данные',
      line: { color: '#1d4ed8', width: 3 },
      marker: { color: '#1d4ed8', size: 9 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    },
    {
      x: summary.forecastYears,
      y: summary.forecastValues,
      type: 'scatter',
      mode: 'lines',
      name: `Прогноз (${getScenarioLabel(getScenario())})`,
      line: { color: '#d97706', width: 3 },
      hovertemplate: '<b>%{x}</b><br>Население: %{y:,}<extra></extra>',
    },
  ];

  const layout = buildLineLayout('Численность населения', summary.details.Settlement_full_name || summary.details.Settlement_name || 'Населённый пункт');
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
  const settlementPop = toNumber(summary.byYear.get(year));

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

  const traces = [
    {
      x: male,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Мужчины',
      marker: { color: '#2563eb' },
      hovertemplate: '<b>Возраст %{y}</b><br>Мужчины: %{customdata:,}<extra></extra>',
      customdata: male.map((v) => Math.abs(v)),
    },
    {
      x: female,
      y: ages,
      type: 'bar',
      orientation: 'h',
      name: 'Женщины',
      marker: { color: '#ec4899' },
      hovertemplate: '<b>Возраст %{y}</b><br>Женщины: %{x:,}<extra></extra>',
    },
  ];

  const layout = {
    title: {
      text: `${escapeHtml(details.Settlement_full_name || details.Settlement_name || 'Населённый пункт')} · ${year}`,
      font: { size: 16 },
      x: 0.02,
      xanchor: 'left',
    },
    barmode: 'relative',
    bargap: 0.05,
    margin: { l: 64, r: 36, t: 48, b: 48 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    legend: { orientation: 'h', x: 0, y: 1.1 },
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
      autorange: 'reversed',
      dtick: 5,
      gridcolor: '#f1f5f9',
    },
    annotations: [
      {
        text: `Общая численность: ${formatInteger(settlementPop)} · сценарий: ${getScenarioLabel(scenario)}`,
        xref: 'paper',
        yref: 'paper',
        x: 0,
        y: 1.18,
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
    margin: { l: 64, r: 24, t: 52, b: 52 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    legend: { orientation: 'h', x: 0, y: 1.12 },
    hovermode: 'x unified',
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
  }));
  exportWorkbook([{ name: 'selection', rows }], `${buildFilename('selection')}.xlsx`);
}

function exportAggregateXlsx() {
  if (!state.currentAggregateSeries) {
    showError('Сначала нужно построить агрегированный временной ряд.');
    return;
  }
  const rows = [];
  state.currentAggregateSeries.actualYears.forEach((year, index) => {
    rows.push({ year, population: state.currentAggregateSeries.actualValues[index], source: 'Фактические данные' });
  });
  state.currentAggregateSeries.forecastYears.forEach((year, index) => {
    rows.push({ year, population: state.currentAggregateSeries.forecastValues[index], source: `Прогноз (${getScenarioLabel(getScenario())})` });
  });
  exportWorkbook([{ name: 'aggregate', rows }], `${buildFilename('aggregate')}.xlsx`);
}

function exportSettlementSeriesXlsx() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    showError('Сначала выберите населённый пункт.');
    return;
  }
  const rows = [];
  summary.actualYears.forEach((year, index) => {
    rows.push({ year, population: summary.actualValues[index], source: 'Фактические данные' });
  });
  summary.forecastYears.forEach((year, index) => {
    rows.push({ year, population: summary.forecastValues[index], source: `Прогноз (${getScenarioLabel(getScenario())})` });
  });
  exportWorkbook([{ name: 'settlement_series', rows }], `${buildFilename('settlement_series')}.xlsx`);
}

function exportPyramidXlsx() {
  if (!state.currentPyramidBundle?.rows?.length) {
    showError('Сначала постройте половозрастную пирамиду.');
    return;
  }
  const rows = state.currentPyramidBundle.rows.map((row) => ({
    age: row.age,
    sex: row.sex,
    scaled_population: row.scaled_pop,
    municipal_profile_population: row.pop,
    scenario: getScenarioLabel(state.currentPyramidBundle.scenario),
    year: state.currentPyramidBundle.year,
    oktmo_syn: state.currentPyramidBundle.oktmo_syn,
  }));
  exportWorkbook([{ name: 'pyramid', rows }], `${buildFilename('pyramid')}.xlsx`);
}

function exportSelectedDetailsXlsx() {
  const summary = state.currentSelectionBundle;
  if (!summary?.details) {
    showError('Сначала выберите населённый пункт.');
    return;
  }
  const rows = Object.entries(summary.details).map(([field, value]) => ({ field, value: formatAny(value) }));
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

function buildFilename(stem) {
  const scenario = getScenario();
  const year = getMapYear();
  return `${stem}_${scenario}_${year}`.replace(/[^a-zA-Z0-9_\-]+/g, '_');
}

function fillSelect(select, options, selectedValue) {
  select.innerHTML = options.map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.text)}</option>`).join('');
  select.value = selectedValue;
}

function getScenario() {
  return els.scenarioSelect.value;
}

function getMapYear() {
  return Number(els.mapYearSelect.value);
}

function getBaseYear() {
  return Number(els.baseYearSelect.value);
}

function getPyramidYear() {
  return Number(els.pyramidYearSelect.value);
}

function getForecastAlias(scenario) {
  return `settlement_forecast_wide_${scenario}.parquet`;
}

function getMunicipalAgeAlias(scenario, year) {
  return `municipal_age_${scenario}_${year}.parquet`;
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

function syncPyramidYearIfNeeded() {
  if (!els.autoPyramidToggle.checked) return;
  const mapYear = getMapYear();
  els.pyramidYearSelect.value = String(Math.max(2021, mapYear));
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

function radiusForPopulation(pop) {
  if (!Number.isFinite(pop) || pop <= 0) return 1.5;
  return Math.max(1.5, Math.min(26, Math.sqrt(pop) / 12));
}

function colorForDelta(deltaPct) {
  if (!Number.isFinite(deltaPct)) return [148, 163, 184, 170];
  const clamped = Math.max(-100, Math.min(100, deltaPct));
  if (clamped < 0) {
    const t = (clamped + 100) / 100;
    return lerpColor([153, 27, 27, 215], [226, 232, 240, 185], t);
  }
  const t = clamped / 100;
  return lerpColor([226, 232, 240, 185], [29, 78, 216, 220], t);
}

function lerpColor(a, b, t) {
  return a.map((value, index) => Math.round(value + (b[index] - value) * t));
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

function formatAny(value) {
  if (value == null || value === '') return '—';
  if (typeof value === 'number') {
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
