import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {composeSlide, layoutTable} from '@openpresentation/opf/composition';
import {renderSvgDeck} from '@openpresentation/opf-render';
import {toPptx, fromPptx} from '../dist/index.js';

// FF-39: the native table frame equals the composed box (as the preview and the
// parity harness measure it), and header/footer furniture is added after every
// content shape, the order opf-render paints it in (furniture last).
const dec = new TextDecoder(), EMU = 9525, TOLERANCE_EMU = 0.02 / 72 * 914400;
const slideXml = async (deck, options) => dec.decode(unzipSync(await toPptx(deck, options))['ppt/slides/slide1.xml']);
const frameOf = xml => {
  const frame = [...xml.matchAll(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g)].map(match => match[0]).find(text => text.includes('<a:tbl>'));
  const [, x, y] = frame.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/), [, cx, cy] = frame.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
  const rows = [...frame.matchAll(/<a:tr h="(\d+)"/g)].map(match => Number(match[1]));
  return {x: +x, y: +y, cx: +cx, cy: +cy, rowSum: rows.reduce((a, b) => a + b, 0), rows: rows.length};
};

// 1. Table frame ext equals the composed box whether the rows fill it or not.
const short = {columns: ['Region', 'Value'], rows: [['North', '1'], ['South', '2']]};
const tall = {columns: ['Region', 'Value'], rows: Array.from({length: 40}, (_, index) => [`Row ${index} with a longer descriptive label`, String(index)])};
let checked = 0, slack = 0, packed = 0;
for (const [label, table] of [['short', short], ['tall', tall]]) for (const scale of [1, 0.5]) for (const withText of [false, true]) {
  const slide = withText
    ? {title: 'Frame', composition: {mode: 'row', weights: [1, 3]}, blocks: [{text: 'Context'}, {table}]}
    : {title: 'Frame', table};
  const deck = {design: {dimensions: {widthInches: 1280 * scale / 96, heightInches: 720 * scale / 96}}, slides: [slide]};
  const geometry = composeSlide(deck.slides[0], {width: 1280 * scale, height: 720 * scale});
  const item = geometry.items.find(entry => entry.field === 'table');
  const frame = frameOf(await slideXml(deck));
  const note = `${label} scale ${scale} text ${withText}`;
  for (const [actual, expected, name] of [[frame.x, item.box.x, 'x'], [frame.y, item.box.y, 'y'], [frame.cx, item.box.width, 'width'], [frame.cy, item.box.height, 'height']])
    assert.ok(Math.abs(actual - expected * EMU) <= TOLERANCE_EMU, `${note}: frame ${name} ${actual / EMU} equals composed box ${expected}`);
  // Rows keep the measured preview heights, which never exceed the composed box.
  const rows = layoutTable(table, item.box, {scale}).rows.reduce((sum, row) => sum + row.box.height, 0);
  assert.ok(frame.rowSum <= frame.cy + frame.rows, `${note}: rows never exceed the declared frame`);
  if (frame.cy - frame.rowSum > 100 * EMU) slack++; else packed++;
  checked++;
}
assert.ok(slack > 0, 'The suite covers tables whose rows do not fill the composed box.');
assert.ok(packed > 0, 'The suite covers tables whose rows fill the composed box.');

// 2. Furniture follows all content in spTree, matching preview paint order.
const names = xml => [...xml.matchAll(/<p:cNvPr id="\d+" name="([^"]*)"/g)].map(match => match[1]).filter(Boolean);
const furniture = name => name.startsWith('OPF furniture');
const decks = [
  {design: {header: {left: {text: 'Confidential'}, right: {date: true, dateFormat: 'yyyy-MM-dd'}}, footer: {left: {text: 'Acme'}, right: {slideNumber: true, slideNumberFormat: '{current} / {total}'}}},
    slides: [{title: 'First', text: 'Body', table: short}, {title: 'Second', bullets: ['One', 'Two']}]},
  {design: {footer: {center: {text: 'Only a footer'}}}, slides: [{title: 'Solo', text: 'Body', table: short}]},
];
let orders = 0;
for (const deck of decks) {
  const entries = unzipSync(await toPptx(deck, {date: '2026-01-02'}));
  const svgs = renderSvgDeck(deck, {trace: true, date: '2026-01-02'});
  for (const [index, slide] of deck.slides.entries()) {
    const xml = dec.decode(entries[`ppt/slides/slide${index + 1}.xml`]), order = names(xml);
    assert.ok(order.some(furniture), `slide ${index + 1} exports furniture`);
    assert.ok(order.some(name => !furniture(name)), `slide ${index + 1} exports content`);
    const firstFurniture = order.findIndex(furniture), lastContent = order.findLastIndex(name => !furniture(name));
    assert.ok(firstFurniture > lastContent, `slide ${index + 1}: furniture ${JSON.stringify(order)} follows every content shape`);
    assert.ok(order.slice(firstFurniture).every(furniture), 'Nothing but furniture follows the first furniture shape.');
    // Same order in the preview SVG: the first furniture element follows the last content path.
    const slideSvg = svgs[index];
    const paint = slideSvg.indexOf('data-opf-furniture-kind');
    const contentPaint = Math.max(...['title', 'text', 'table', 'bullets'].map(field => slideSvg.lastIndexOf(`data-opf-path="slides.${index}.${field}`)));
    assert.ok(paint > contentPaint && contentPaint > 0, `slide ${index + 1}: preview paints furniture after content`);
    orders++;
  }
  // Native fields and re-import do not depend on furniture position.
  const first = dec.decode(entries['ppt/slides/slide1.xml']);
  if (deck.design.footer.right) assert.match(first, /<a:fld [^>]*type="slidenum"/);
  const imported = await fromPptx(await toPptx(deck, {date: '2026-01-02'}));
  assert.deepEqual(imported.design.footer, deck.design.footer);
  assert.deepEqual(imported.slides.map(slide => slide.title), deck.slides.map(slide => slide.title));
}
console.log(`Table frame and furniture order passed: ${checked} table frames equal the composed box (${slack} with slack rows, ${packed} packed); ${orders} slides paint furniture after content in spTree and SVG.`);
