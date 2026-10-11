import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {unzipSync, zipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {regionIndexes, foreignLayoutReference} from '../dist/template-layouts.js';

const layouts = JSON.parse(await readFile(new URL('./fixtures/layout-templates-0.19.json', import.meta.url))).layouts;
const catalogs = [{source: 'pkg:@openpresentation/gallery@2', layouts}];
const options = {catalogs, seed: 1, timestamp: '2026-01-01T00:00:00Z', strict: false, provenance: 'full'};
const dec = new TextDecoder(), enc = new TextEncoder();
const many = {name: 'Maximum regions', areas: [], regions: {}};
for (let y = 0; y < 12; y++) {
  const row = [];
  for (let x = 0; x < 12; x++) {
    const name = `cell-${y}-${x}`;
    row.push(name);
    many.regions[name] = {accepts: x === 11 && y === 11 ? ['text'] : ['text', 'image']};
  }
  many.areas.push(row.join(' '));
}
const indexes = Object.values(regionIndexes(many));
assert.equal(new Set(indexes).size, 144, 'Maximum custom grid has no placeholder-index collisions across kinds');

const source = {slides: [{layout: 'text', title: 'Short title', text: 'Body'}, {layout: 'text', title: 'A very long title '.repeat(40), text: 'Body'}, {layout: 'text', title: [{text: 'Rich title', bold: true}], text: 'Body'}]};
const parts = unzipSync(await toPptx(source, options));
const titleShape = n => [...dec.decode(parts[`ppt/slides/slide${n}.xml`]).matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map(([shape]) => shape).filter(shape => /name="OPF heading slides\.\d+\.title/.test(shape));
assert.equal(titleShape(1).length, 1);
assert.match(titleShape(1)[0], /<a:spLocks noGrp="1"\/>/, 'Native title placeholders cannot be grouped');
assert.match(titleShape(1)[0], /<p:ph type="title"\/>/, 'One accepted plain title line is a native title placeholder');
for (const n of [2, 3]) assert.ok(titleShape(n).every(shape => !/<p:ph\b/.test(shape)), 'Long and rich titles remain tagged line shapes');
const templateParts = Object.entries(parts).filter(([path]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)).map(([, bytes]) => dec.decode(bytes));
assert.ok(templateParts.some(xml => /name="Title"[\s\S]*?<a:lstStyle><a:lvl1pPr[^>]*><a:defRPr sz="\d+"/.test(xml)), 'Layout prompts inherit an explicit OPF font size');
assert.ok(templateParts.some(xml => /<a:latin typeface="\+mj-lt"/.test(xml)), 'Layout title prompts inherit the selected heading family from their own master');
assert.deepEqual((await fromPptx(zipSync(parts), options)).slides.slice(0, 2), source.slides.slice(0, 2), 'Native title identity preserves authored titles through import');

const automaticSource = {slides: [{title: 'Automatic', text: 'Body'}, {layout: 'auto', title: 'Explicit automatic', text: 'Body'}]};
const automaticBytes = await toPptx(automaticSource, options);
const automaticParts = unzipSync(automaticBytes);
const automaticLayouts = Object.entries(automaticParts).filter(([path]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)).map(([, bytes]) => dec.decode(bytes));
assert.equal(automaticLayouts.filter(xml => /name="OPF auto"/.test(xml) && /type="titleOnly"/.test(xml)).length, 1, 'Automatic slides use one OPF auto titleOnly layout per master');
const automaticReports = [];
assert.deepEqual((await fromPptx(automaticBytes, {...options, onDiagnostic: item => automaticReports.push(item)})).slides, automaticSource.slides, 'Automatic layout preserves implicit versus explicit authored layout');
const mixed = {slides: [{layout: 'text', title: 'Template', text: 'Body'}, {title: 'Automatic', text: 'Body'}]};
const mixedParts = unzipSync(await toPptx(mixed, options));
const rel = dec.decode(mixedParts['ppt/slides/_rels/slide2.xml.rels']).match(/<Relationship[^>]*Type="[^"]*\/slideLayout"[^>]*Target="([^"]+)"/)[1];
let firstRelationships = dec.decode(mixedParts['ppt/slides/_rels/slide1.xml.rels']);
firstRelationships = firstRelationships.replace(/(<Relationship[^>]*Type="[^"]*\/slideLayout"[^>]*Target=")[^"]+/, `$1${rel}`);
mixedParts['ppt/slides/_rels/slide1.xml.rels'] = enc.encode(firstRelationships);
const changedReports = [];
const changed = await fromPptx(zipSync(mixedParts), {...options, onDiagnostic: item => changedReports.push(item)});
assert.equal(changed.slides[0].layout, 'auto', 'Change Layout from a template to OPF auto wins over stored layout');
assert.ok(changedReports.some(item => item.code === 'layout-changed'));
assert.ok(!automaticReports.some(item => item.code === 'unresolved-reference' && /auto/.test(item.message)), 'auto is a built-in sentinel, not a catalog reference');
const bareAuto = unzipSync(await toPptx({slides: [{title: 'Bare automatic', text: 'Body'}]}, {...options, catalogs: []}));
assert.ok(Object.entries(bareAuto).some(([path, bytes]) => /^ppt\/slideLayouts\//.test(path) && /name="OPF auto"/.test(dec.decode(bytes))), 'Bare automatic decks use OPF auto without registered catalogs');
const foreign = unzipSync(await toPptx({slides: [{title: 'Foreign', text: 'Body'}]}, {...options, catalogs: [], provenance: false}));
const foreignRelationship = dec.decode(foreign['ppt/slides/_rels/slide1.xml.rels']).match(/<Relationship[^>]*Type="[^"]*\/slideLayout"[^>]*Target="([^"]+)"/)[1];
const layoutPath = `ppt/${foreignRelationship.replace(/^\.\.\//, '')}`;
foreign[layoutPath] = enc.encode(dec.decode(foreign[layoutPath]).replace(/<p:custDataLst>[\s\S]*?<\/p:custDataLst>/g, ''));
for (const [type, name, expected] of [['title', 'Foreign', 'cover'], ['secHead', 'Foreign', 'section'], ['obj', 'Foreign', 'text'], ['tx', 'Foreign', 'text'], ['twoObj', 'Foreign', 'two-column'], ['twoTxTwoObj', 'Foreign', 'two-column'], ['picTx', 'Foreign', 'image-beside'], ['blank', 'Foreign', 'auto'], ['titleOnly', 'Foreign', 'auto'], ['cust', 'cHaRt BeSiDe', 'chart-beside'], ['cust', 'Unknown', 'auto']]) {
  const copy = {...foreign};
  let xml = dec.decode(copy[layoutPath]).replace(/<p:sldLayout\b[^>]*>/, open => open.replace(/\s+type="[^"]*"/, '') .replace(/>$/, ` type="${type}">`));
  xml = xml.replace(/<p:cSld\b[^>]*>/, open => open.replace(/\s+name="[^"]*"/, '').replace(/>$/, ` name="${name}">`));
  copy[layoutPath] = enc.encode(xml);
  const imported = await fromPptx(zipSync(copy), options);
  assert.equal(imported.slides[0].layout, expected === 'auto' ? expected : `default:${expected}`, `${type}/${name}: foreign layout maps by type or built-in name`);
}
console.log('Template layout review: 144 region identities, native plain titles, prompt typography and foreign layout mapping passed. Native Office acceptance remains pending.');

assert.equal(foreignLayoutReference({type: 'titleOnly'}), 'auto');
assert.equal(foreignLayoutReference({type: 'blank'}, []), 'auto');
assert.equal(foreignLayoutReference({type: 'title'}, []), undefined);

const competing = [{source: 'first', layouts: {cover: {...layouts.cover, name: 'First cover'}}}, {source: 'second', layouts: {cover: {...layouts.cover, name: 'Second cover'}}}];
assert.equal(foreignLayoutReference({type: 'cust', 'p:cSld': {name: 'Second cover'}}, competing), 'auto', 'a shadowed host record cannot supply a misleading match');
assert.equal(foreignLayoutReference({type: 'cust', 'p:cSld': {name: 'First cover'}}, competing), 'default:cover');
assert.equal((await fromPptx(zipSync(foreign))).slides[0].layout, 'auto', 'normal import maps a foreign titleOnly without options');
assert.equal((await fromPptx(zipSync(foreign), {catalogs: []})).slides[0].layout, 'auto', 'explicitly empty catalogs retain automatic mapping');
