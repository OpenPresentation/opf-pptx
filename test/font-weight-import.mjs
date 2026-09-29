import assert from 'node:assert/strict';
import {validatePresentation} from '@openpresentation/opf';
import {prepareNodeFonts} from '@openpresentation/opf-render/fonts-node';
import PptxGenJS from '../vendor/pptxgenjs/pptxgen.es.js';
import {toPptx, fromPptx} from '../dist/index.js';
import {FACE_STYLE_WORDS, FONT_WEIGHT_WORDS, splitWeightFace} from '../src/font-weights.js';

// The exporter names a chosen family's weight faces by their native style-link
// names ("Roboto SemiBold"). Import maps them back to family + OPF bold and never
// alters a genuine family that merely ends in a weight word ("Arial Black").
const runs = deck => deck.slides[0].blocks.filter(block => block.type === 'text').flatMap(block => [block.text].flat());
async function nativeRuns(faces, theme) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  if (theme) pptx.theme = theme;
  const slide = pptx.addSlide();
  faces.forEach((face, index) => slide.addText(`Sample ${index}`, {x: 1, y: 0.5 + index * 0.6, w: 6, h: 0.5, fontSize: 18, fontFace: face}));
  const diagnostics = [];
  const deck = await fromPptx(await pptx.write({outputType: 'uint8array'}), {onDiagnostic: item => diagnostics.push(item)});
  assert.equal(validatePresentation(deck).valid, true);
  return {runs: runs(deck), diagnostics};
}
const family = run => run.fontFamily;

// 1. Exporter round trip through the real weight faces: heading-weight quote body is Roboto SemiBold, attribution Roboto Medium.
const {options} = await prepareNodeFonts();
const source = {design: {fontScheme: 'roboto'}, slides: [{title: 'Weights', quote: {text: 'Retain the selected source.', attribution: 'Reviewer', source: 'Recorded interview'}}]};
const exported = await toPptx(structuredClone(source), {...options, strictAssets: true});
const diagnostics = [];
const round = await fromPptx(exported, {onDiagnostic: item => diagnostics.push(item)});
assert.equal(validatePresentation(round).valid, true);
assert.ok(!JSON.stringify(round).includes('Roboto SemiBold') && !JSON.stringify(round).includes('Roboto Medium'), 'native weight-face names must not reach OPF');
const [body, footer] = runs(round);
assert.deepEqual([family(body), body.bold], ['Roboto', true], 'SemiBold maps to Roboto + bold');
assert.deepEqual([family(footer), footer.bold], ['Roboto', undefined], 'Medium maps to plain Roboto');
assert.ok(diagnostics.some(item => item.code === 'approximate-body-font-weight'), 'the lost numeric weight is reported');

// 2. Native weight faces of a known family, whatever the theme.
{
  const {runs: imported, diagnostics: notes} = await nativeRuns(['Roboto Medium', 'Roboto SemiBold', 'Roboto ExtraBold', 'Roboto Light', 'Roboto Bold', 'Roboto']);
  assert.deepEqual(imported.map(family), Array(6).fill('Roboto'));
  assert.deepEqual(imported.map(run => run.bold), [undefined, true, true, undefined, true, undefined]);
  assert.deepEqual(notes.filter(item => item.code === 'approximate-body-font-weight').length, 4, 'Medium, SemiBold, ExtraBold and Light are approximated; Bold and regular are exact');
}

// 3. The deck's chosen family owns its weight faces; other families do not.
{
  const chosen = await nativeRuns(['Inter SemiBold', 'Inter Medium'], {headFontFace: 'Inter', bodyFontFace: 'Inter'});
  assert.deepEqual(chosen.runs.map(family), ['Inter', 'Inter']);
  assert.deepEqual(chosen.runs.map(run => run.bold), [true, undefined]);
  const foreign = await nativeRuns(['Inter SemiBold'], {headFontFace: 'Calibri', bodyFontFace: 'Calibri'});
  assert.equal(foreign.runs[0].fontFamily, 'Inter SemiBold', 'a face of an unrelated family is left as authored');
  assert.equal(foreign.runs[0].bold, undefined);
}

// 4. Genuine families ending in a weight word stay exactly as authored, even when their base family is chosen.
for (const [face, theme] of [['Arial Black', undefined], ['Arial Black', {headFontFace: 'Arial', bodyFontFace: 'Arial'}], ['Segoe UI Semibold', {headFontFace: 'Segoe UI', bodyFontFace: 'Segoe UI'}], ['Calibri Light', {headFontFace: 'Calibri', bodyFontFace: 'Calibri'}], ['Roboto Mono', undefined]]) {
  const {runs: imported, diagnostics: notes} = await nativeRuns([face], theme);
  assert.equal(imported[0].fontFamily, face, `${face} keeps its genuine family name`);
  assert.equal(imported[0].bold, undefined, `${face} is not turned bold`);
  assert.ok(!notes.some(item => item.code === 'approximate-body-font-weight'), `${face} reports no weight approximation`);
}

// 5. Table cell runs use the same mapping.
{
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.addSlide().addTable([[{text: 'Cell', options: {fontFace: 'Roboto SemiBold', fontSize: 14}}]], {x: 1, y: 1, w: 4, colW: [4]});
  const cell = (await fromPptx(await pptx.write({outputType: 'uint8array'}))).slides[0].blocks.find(block => block.table).table.rows[0][0];
  const cellRuns = [cell.value].flat().filter(run => typeof run === 'object');
  assert.deepEqual(cellRuns.map(run => [run.fontFamily, run.bold]), [['Roboto', true]]);
}

// 6. One table: every recognized weight word both keeps a chosen family's face (exporter) and splits it (importer).
for (const [word, weight] of FONT_WEIGHT_WORDS) {
  assert.ok(FACE_STYLE_WORDS.test(word), word);
  assert.equal(splitWeightFace(`Roboto ${word}`, ['Roboto'])?.weight, weight, word);
}
assert.equal(splitWeightFace('Roboto Fancy', ['Roboto']), undefined);
assert.equal(splitWeightFace('Arial Black', ['Arial']), undefined);
assert.deepEqual(splitWeightFace('Roboto Medium Italic', ['Roboto']), {family: 'Roboto', weight: 500, italic: true});

console.log(JSON.stringify({test: 'font-weight-import', passed: true, words: FONT_WEIGHT_WORDS.length}));
