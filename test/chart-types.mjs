// FF-22: every classic chart type the core catalog keeps exports the exact
// native construct for its Aspose.Slides ChartType and imports back to the
// same id; deprecated ids export like their replacement; other ids keep the
// legacy construct.
import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {CHART_TYPES, DEPRECATED_CHART_TYPES, resolveChartType} from '../dist/chart-types.js';

const decoder = new TextDecoder();
const categoryData = {columns: ['Quarter', 'North', 'South', 'West'], rows: [['Q1', 4, 3, 2], ['Q2', 5, 2, 3], ['Q3', 6, 4, 1]]};
const scatterData = {columns: ['Point', 'Spend', 'Revenue', 'Margin'], rows: [['a', 1, 10, 3], ['b', 2, 14, 4], ['c', 4, 19, 6]]};
const dataFor = (id) => (CHART_TYPES[id].family === 'xy' ? scatterData : categoryData);

async function exportChart(type, data) {
  const bytes = await toPptx({slides: [{title: type, chart: {type, data}}]});
  const entries = unzipSync(bytes);
  const xml = decoder.decode(entries['ppt/charts/chart1.xml']);
  const imported = (await fromPptx(bytes)).slides[0];
  const chart = imported.chart ?? imported.blocks?.find((block) => block.chart)?.chart;
  return {xml, chart};
}

function construct(xml) {
  const plot = xml.match(/<c:plotArea>([\s\S]*)<\/c:plotArea>/)[1];
  const element = plot.match(/<c:(\w+Chart)>/)[1];
  const body = plot.match(new RegExp(`<c:${element}>([\\s\\S]*?)</c:${element}>`))[1];
  const val = (name) => body.match(new RegExp(`<c:${name} val="([^"]*)"/>`))?.[1];
  const series = [...body.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map((m) => m[1]);
  return {
    element,
    body,
    barDir: val('barDir'),
    grouping: val('grouping'),
    overlap: val('overlap'),
    radarStyle: val('radarStyle'),
    series,
    symbols: series.map((ser) => ser.match(/<c:marker>\s*<c:symbol val="([^"]*)"/)?.[1]),
    valFormat: plot.match(/<c:valAx>[\s\S]*?<c:numFmt formatCode="([^"]*)"/)?.[1],
  };
}

const expected = {
  column: {element: 'barChart', barDir: 'col', grouping: 'clustered'},
  'stacked-column-3x': {element: 'barChart', barDir: 'col', grouping: 'stacked', overlap: '100'},
  '100pct-stacked-column-3x': {element: 'barChart', barDir: 'col', grouping: 'percentStacked', overlap: '100', valFormat: '0%'},
  bar: {element: 'barChart', barDir: 'bar', grouping: 'clustered'},
  'stacked-bar-3x': {element: 'barChart', barDir: 'bar', grouping: 'stacked', overlap: '100'},
  '100pct-stacked-bar-3x': {element: 'barChart', barDir: 'bar', grouping: 'percentStacked', overlap: '100', valFormat: '0%'},
  line: {element: 'lineChart', grouping: 'standard', symbol: 'none'},
  'line-with-markers': {element: 'lineChart', grouping: 'standard', symbol: 'circle'},
  'stacked-line-3x': {element: 'lineChart', grouping: 'stacked', symbol: 'none'},
  'stacked-line-with-markers-3x': {element: 'lineChart', grouping: 'stacked', symbol: 'circle'},
  area: {element: 'areaChart', grouping: 'standard'},
  'stacked-area-3x': {element: 'areaChart', grouping: 'stacked'},
  '100pct-stacked-area-3x': {element: 'areaChart', grouping: 'percentStacked', valFormat: '0%'},
  pie: {element: 'pieChart', seriesCount: 1},
  doughnut: {element: 'doughnutChart', seriesCount: 1},
  scatter: {element: 'scatterChart', seriesCount: 2, symbol: 'circle'},
  radar: {element: 'radarChart', radarStyle: 'standard', symbol: 'none'},
  'radar-with-markers': {element: 'radarChart', radarStyle: 'marker', symbol: 'circle'},
  'filled-radar': {element: 'radarChart', radarStyle: 'filled'},
};

const classic = Object.keys(CHART_TYPES).filter((id) => CHART_TYPES[id].family !== 'chartex');
assert.deepEqual(classic.sort(), Object.keys(expected).sort(), 'every classic kept chart type has an expectation');

let checked = 0;
for (const [id, want] of Object.entries(expected)) {
  const data = dataFor(id);
  const {xml, chart} = await exportChart(id, data);
  const got = construct(xml);
  for (const key of ['element', 'barDir', 'grouping', 'overlap', 'radarStyle', 'valFormat']) {
    if (want[key] !== undefined) assert.equal(got[key], want[key], `${id}: ${key}`);
  }
  if (want.seriesCount !== undefined) assert.equal(got.series.length, want.seriesCount, `${id}: series count`);
  if (want.symbol !== undefined) assert.ok(got.symbols.every((symbol) => symbol === want.symbol), `${id}: marker symbols ${got.symbols}`);
  if (got.element === 'lineChart' || got.element === 'areaChart') {
    assert.ok(got.body.startsWith(`<c:grouping val="${want.grouping}"/>`), `${id}: grouping is the first child (CT_${got.element})`);
    assert.doesNotMatch(got.body, /<c:invertIfNegative/, `${id}: invertIfNegative belongs to bar series only`);
  }
  assert.equal(chart.type, id, `${id}: imports back to the same chart type id`);
  if (id === 'pie' || id === 'doughnut') assert.deepEqual(chart.data, {columns: data.columns.slice(0, 2), rows: data.rows.map((row) => row.slice(0, 2))});
  else if (id === 'scatter') assert.deepEqual(chart.data.rows.map((row) => row.slice(1)), data.rows.map((row) => row.slice(1)), 'scatter X/Y values round-trip');
  else assert.deepEqual(chart.data, data, `${id}: data round-trips`);
  checked++;
}

// Deprecated ids export exactly like their replacement and import as it.
for (const [deprecated, replacement] of Object.entries(DEPRECATED_CHART_TYPES)) {
  assert.equal(resolveChartType(deprecated).id, replacement);
  if (CHART_TYPES[replacement].family === 'chartex') continue;
  const data = dataFor(replacement);
  const [a, b] = await Promise.all([exportChart(deprecated, data), exportChart(replacement, data)]);
  const strip = (xml) => construct(xml).body;
  assert.equal(strip(a.xml), strip(b.xml), `${deprecated} exports like ${replacement}`);
  assert.equal(a.chart.type, replacement, `${deprecated} imports as ${replacement}`);
  checked++;
}

// Ids outside the core catalog keep the legacy heuristic.
for (const [id, element, barDir] of [['custom-kpi', 'barChart', 'col'], ['my-bar', 'barChart', 'bar'], ['trend-line', 'lineChart', undefined], ['donut', 'doughnutChart', undefined]]) {
  const got = construct((await exportChart(id, categoryData)).xml);
  assert.equal(got.element, element, id);
  if (barDir) assert.equal(got.barDir, barDir, id);
  checked++;
}

console.log(`Chart types passed: ${checked} exports (${Object.keys(expected).length} classic Aspose chart types, ${Object.keys(DEPRECATED_CHART_TYPES).length} deprecated aliases, legacy ids) with exact constructs and same-id reimport.`);
