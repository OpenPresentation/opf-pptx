// A chart whose data is a single column (histogram, dot plot: values without a category column) was replaced by
// a text placeholder with no chart part, no graphic frame and no diagnostic. It now exports a native chart and
// says how the data was adapted; data that cannot be plotted keeps its placeholder and reports why.
import assert from 'node:assert/strict';
import {strFromU8, unzipSync} from 'fflate';
import {catalogs} from '@openpresentation/opf';
import {fromPptx, toPptx} from '../dist/index.js';

const deck = (type, data) => ({name: 'single column chart', language: 'english', design: {fontScheme: 'calibri'}, slides: [{id: 'a', layout: 'chart-1x', title: 'Chart', chart: {type, data}, text: 'Body'}]});
async function exported(type, data) {
  const diagnostics = [];
  const bytes = await toPptx(deck(type, data), {onDiagnostic: (diagnostic) => diagnostics.push(diagnostic)});
  const parts = unzipSync(bytes);
  const slide = strFromU8(parts['ppt/slides/slide1.xml']);
  const charts = Object.keys(parts).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name));
  return {bytes, parts, slide, charts, diagnostics, chart: charts.length ? strFromU8(parts[charts[0]]) : '', frame: slide.includes('graphicFrame'), chartDiagnostics: diagnostics.filter((diagnostic) => diagnostic.code.startsWith('chart-'))};
}
const numbers = (xml, tag) => [...xml.matchAll(new RegExp(`<c:${tag}>([\\s\\S]*?)</c:${tag}>`, 'g'))].map(([, body]) => [...body.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map(([, value]) => value));

// The FF-09 case: a one-column histogram.
const values = [3, 5, 8, 13];
const histogram = await exported('histogram', {columns: ['Value'], rows: values.map((value) => [value])});
assert.equal(histogram.charts.length, 1, 'a chart part is exported');
assert.ok(histogram.frame, 'the slide has a graphic frame');
assert.doesNotMatch(histogram.slide, /&quot;type&quot;/, 'no placeholder dump replaces the chart');
assert.equal(histogram.chartDiagnostics.length, 1);
assert.deepEqual({code: histogram.chartDiagnostics[0].code, adaptation: histogram.chartDiagnostics[0].adaptation, path: histogram.chartDiagnostics[0].path}, {code: 'chart-data-adapted', adaptation: 'histogram-binned', path: 'slides.0.chart'});
assert.match(histogram.chartDiagnostics[0].message, /4 values.*3 equal-width bins/);
// Sturges: ceil(log2(4)) + 1 = 3 bins over [3, 13]; the maximum falls in the last bin.
assert.match(histogram.chart, /<c:barDir val="col"\/>/);
assert.deepEqual(numbers(histogram.chart, 'cat'), [['3–6.33333', '6.33333–9.66667', '9.66667–13']]);
assert.deepEqual(numbers(histogram.chart, 'val'), [['2', '1', '1']]);
assert.equal(numbers(histogram.chart, 'val')[0].reduce((total, count) => total + Number(count), 0), values.length, 'every value is counted once');
assert.deepEqual(await toPptx(deck('histogram', {columns: ['Value'], rows: values.map((value) => [value])})), histogram.bytes, 'deterministic');
// The export re-imports as a valid column chart of the binned counts.
const imported = await fromPptx(histogram.bytes);
const importedChart = (imported.slides[0].chart ?? imported.slides[0].blocks?.find((block) => block.chart)?.chart);
assert.equal(importedChart.type, 'column');
assert.deepEqual(importedChart.data.columns, ['Bin', 'Frequency']);
assert.deepEqual(importedChart.data.rows.map((row) => row[1]), [2, 1, 1]);

// Numeric strings count, other cells are ignored; a constant column is one bin; one value is one bin.
const strings = await exported('histogram', {columns: ['Value'], rows: [['1'], ['2'], ['x'], [null], [3], ['4.5']]});
assert.equal(numbers(strings.chart, 'val')[0].reduce((total, count) => total + Number(count), 0), 4, 'non-numeric cells are not counted as zeros');
const constant = await exported('histogram', {columns: ['Value'], rows: [[7], [7], [7]]});
assert.deepEqual(numbers(constant.chart, 'cat'), [['7']]);
assert.deepEqual(numbers(constant.chart, 'val'), [['3']]);
const one = await exported('histogram', {columns: ['Value'], rows: [[7]]});
assert.deepEqual(numbers(one.chart, 'val'), [['1']]);
const wide = await exported('histogram', {columns: ['Value'], rows: Array.from({length: 1000}, (_, index) => [index])});
assert.equal(numbers(wide.chart, 'val')[0].length, 11, 'Sturges bins for 1000 values');
assert.equal(numbers(wide.chart, 'val')[0].reduce((total, count) => total + Number(count), 0), 1000);

// The same silent drop applied to every chart type with one data column (dot plot is the other catalog one).
const singleColumn = catalogs.chartTypes.filter((record) => record.columns?.length === 1).map((record) => record.id);
assert.deepEqual(singleColumn.sort(), ['dot-plot', 'histogram'], 'the catalog chart types with one data column');
const types = new Set([...catalogs.chartTypes.map((record) => record.id), 'column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar']);
for (const type of types) {
  const result = await exported(type, {columns: ['Value'], rows: [[3], [5], [8]]});
  assert.equal(result.charts.length, 1, `${type}: a chart part is exported`);
  assert.ok(result.frame, `${type}: the slide has a graphic frame`);
  assert.deepEqual(result.chartDiagnostics.map((diagnostic) => diagnostic.code), ['chart-data-adapted'], `${type}: the adaptation is reported`);
  assert.equal(result.chartDiagnostics[0].adaptation, type === 'histogram' ? 'histogram-binned' : 'row-numbers', type);
}
const dots = await exported('dot-plot', {columns: ['2023'], rows: [[3], [5], [8]]});
assert.deepEqual(numbers(dots.chart, 'cat'), [['1', '2', '3']], 'rows are numbered');
assert.deepEqual(numbers(dots.chart, 'val'), [['3', '5', '8']]);
assert.match(dots.chartDiagnostics[0].message, /'2023'.*row numbers/);
const scatter = await exported('scatter', {columns: ['Y'], rows: [[3], [5], [8]]});
assert.deepEqual(numbers(scatter.chart, 'yVal'), [['3', '5', '8']]);
assert.deepEqual(numbers(scatter.chart, 'xVal'), [['1', '2', '3']]);

// Data that cannot be plotted keeps the placeholder frame and says why, instead of dropping the chart quietly.
for (const [label, data, reason] of [
  ['text only', {columns: ['Category'], rows: [['a'], ['b']]}, 'single-column-not-numeric'],
  ['external data', {src: 'https://example.invalid/data.csv'}, 'data-not-inline']
]) {
  const result = await exported('histogram', data);
  assert.equal(result.charts.length, 0, label);
  assert.match(result.slide, /Chart/, `${label}: the placeholder frame stays`);
  assert.deepEqual(result.chartDiagnostics.map((diagnostic) => [diagnostic.code, diagnostic.reason, diagnostic.path]), [['chart-data-unplottable', reason, 'slides.0.chart']], label);
  assert.ok(result.chartDiagnostics[0].message.length > 20, label);
}

// Charts with a category column and one or more series are unchanged: same shape, no chart diagnostics.
for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'radar', 'histogram', 'waterfall']) {
  const result = await exported(type, {columns: ['Category', 'Value'], rows: [['a', 3], ['b', 4]]});
  assert.equal(result.charts.length, 1, type);
  assert.deepEqual(result.chartDiagnostics, [], `${type}: no diagnostic for ordinary data`);
  assert.deepEqual(numbers(result.chart, 'cat'), [['a', 'b']], type);
}
console.log(`Single-column charts passed: histogram binning (${values.length} values to 3 bins, constant, single and 1000 values), ${types.size} chart types exported with a reported adaptation, unplottable data reported, ordinary charts unchanged.`);
