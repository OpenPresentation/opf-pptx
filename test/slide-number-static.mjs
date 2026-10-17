import assert from 'node:assert/strict';
import {unzipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {isBareSlideNumber} from '../dist/furniture-fields.js';

// RR-91: Keynote replaces the whole paragraph that holds a slide-number field (placeholder or plain text box) with the number, so
// `{{slide.number}} / {{deck.slideCount}}` shows only "2". `slideNumberWithText: 'static'` writes the number of such a part as fixed
// text in an ordinary tagged shape; the default keeps the live field in the Slide Number placeholder.
const dec = new TextDecoder();
const options = {seed: 1, timestamp: '2026-01-01T00:00:00Z', zipDate: '2026-01-01T00:00:00Z'};
const deck = right => ({design: {footer: {left: {text: 'Acme Corp'}, right: {text: right}}}, slides: [{title: 'One'}, {title: 'Two'}, {title: 'Three'}]});
const slide = (entries, index) => dec.decode(entries[`ppt/slides/slide${index}.xml`]);
const shapes = xml => [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map(match => match[0]).filter(shape => /name="OPF furniture/.test(shape));
const phTypes = xml => shapes(xml).map(shape => shape.match(/<p:ph type="(\w+)"/)?.[1]).filter(Boolean);
const texts = shape => [...shape.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(match => match[1]).join('');
const read = async bytes => { const issues = [], imported = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)}); return {imported, issues}; };

assert.equal(isBareSlideNumber({text: '2', fields: [{type: 'slideNumber', start: 0, end: 1}]}), true);
assert.equal(isBareSlideNumber({text: '2 / 6', fields: [{type: 'slideNumber', start: 0, end: 1}]}), false);
assert.equal(isBareSlideNumber({text: 'Page 2', fields: [{type: 'slideNumber', start: 5, end: 6}]}), false);
assert.equal(isBareSlideNumber({text: 'Acme'}), false);

// Default: the placeholder holds the live field and the fixed words, as before.
{
  const entries = unzipSync(await toPptx(deck('{{slide.number}} / {{deck.slideCount}}'), options));
  for (const index of [1, 2, 3]) {
    assert.deepEqual(phTypes(slide(entries, index)), ['ftr', 'sldNum']);
    assert.match(slide(entries, index), new RegExp(`type="slidenum">(?:(?!</a:fld>)[\\s\\S])*<a:t>${index}</a:t></a:fld><a:r>(?:(?!</a:r>)[\\s\\S])*<a:t> / 3</a:t>`));
  }
}

// Static: "2 / 3" is one fixed run in an ordinary shape; no field and no Slide Number placeholder. The authored text still returns.
for (const right of ['{{slide.number}} / {{deck.slideCount}}', 'Page {{slide.number}}', '{{slide.number}}{{slide.number}}']) {
  const bytes = await toPptx(deck(right), {...options, slideNumberWithText: 'static'});
  const entries = unzipSync(bytes);
  for (const index of [1, 2, 3]) {
    const xml = slide(entries, index), parts = shapes(xml);
    assert.deepEqual(phTypes(xml), ['ftr'], `${right}: only the footer text is a placeholder`);
    assert.equal(/type="slidenum"/.test(xml), false, `${right}: no live field`);
    assert.equal(texts(parts[1]), right.replace(/\{\{slide\.number\}\}/g, String(index)).replace('{{deck.slideCount}}', '3'));
    assert.equal(parts[1].match(/<a:r>/g).length, 1, `${right}: one fixed run (what Keynote shows as written)`);
  }
  const master = dec.decode(entries['ppt/slideMasters/slideMaster1.xml']);
  assert.match(master, /<p:hf sldNum="0" hdr="0" ftr="1" dt="0"\/>/, `${right}: the dialog shows no slide number`);
  const {imported, issues} = await read(bytes);
  assert.deepEqual(imported.design.footer, deck(right).design.footer, `${right}: the authored footer text returns`);
  assert.deepEqual(issues.filter(issue => !/-import-reflow$/.test(issue.code)), [], `${right}: no provenance diagnostic`);
}

// A bare number keeps its field and placeholder in static mode (Keynote shows a lone number correctly).
{
  const entries = unzipSync(await toPptx(deck('{{slide.number}}'), {...options, slideNumberWithText: 'static'}));
  for (const index of [1, 2, 3]) {
    assert.deepEqual(phTypes(slide(entries, index)), ['ftr', 'sldNum']);
    assert.match(slide(entries, index), /type="slidenum"/);
  }
}

// A header number (never a placeholder) also becomes fixed text when it has words around it.
{
  const source = {design: {header: {right: {text: 'Page {{slide.number}}'}}}, slides: [{title: 'One'}, {title: 'Two'}]};
  assert.match(slide(unzipSync(await toPptx(source, options)), 2), /type="slidenum"/);
  const entries = unzipSync(await toPptx(source, {...options, slideNumberWithText: 'static'}));
  assert.equal(/type="slidenum"/.test(slide(entries, 2)), false);
  const xml = [slide(entries, 2), dec.decode(entries['ppt/slideMasters/slideMaster1.xml'])].join('');
  assert.ok(/<a:t>Page 2<\/a:t>/.test(xml));
}

await assert.rejects(toPptx(deck('{{slide.number}}'), {...options, slideNumberWithText: 'live'}), error => error.code === 'invalid-slide-number-option' && error.path === 'options.slideNumberWithText');
console.log('Static slide numbers passed: default placeholder unchanged, three number-with-words forms fixed text, bare number and header, round trip and option validation.');
