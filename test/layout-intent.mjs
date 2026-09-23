import assert from 'node:assert/strict';
import {unzipSync, zipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {validatePresentation, catalogs} from '@openpresentation/opf';

// FF-29: slide layout intent (layout id, type, composition, composition hints
// and the inline catalogs.layouts record) is part of each OPF_SLIDE_V1 record.
// Import restores it with the document record (FF-32) and without it, for
// example after a slide is pasted into another deck.
const enc = new TextEncoder(), dec = new TextDecoder();
const base = catalogs.layouts.find(record => record.id === 'title-subtitle');
const heroA = {...structuredClone(base), id: 'gallery-hero', name: 'Gallery Hero'};
const heroB = {...structuredClone(heroA), name: 'Gallery Hero (other deck)'};
const sideB = {...structuredClone(base), id: 'gallery-side', name: 'Gallery Side'};
const deckA = {
  $schema: 'https://openpresentation.org/schema/opf/v1', name: 'Deck A', narrative: 'problem-solution',
  design: {contentAlignment: 'center'}, catalogs: {layouts: {records: [heroA]}},
  slides: [
    {layout: 'gallery-hero', title: 'Hello', subtitle: 'World', composition: {mode: 'column'}, design: {titleAlignment: 'center', contentBox: false}},
    {layout: 'title-subtitle', title: 'Two', subtitle: 'Bundled'}
  ]
};
const deckB = {
  $schema: 'https://openpresentation.org/schema/opf/v1', name: 'Deck B', catalogs: {layouts: {records: [heroB, sideB]}},
  slides: [{layout: 'gallery-hero', title: 'From B', subtitle: 'Different record'}, {layout: 'gallery-side', title: 'Side', subtitle: 'Only in B'}]
};
for (const deck of [deckA, deckB]) assert.equal(validatePresentation(deck).valid, true, JSON.stringify(validatePresentation(deck).errors));
const exportedA = await toPptx(structuredClone(deckA)), exportedB = await toPptx(structuredClone(deckB));

const read = async bytes => {
  const issues = [];
  const deck = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)});
  assert.equal(validatePresentation(deck).valid, true);
  return {deck, provenance: issues.filter(issue => /provenance|reference|slide-id/.test(issue.code)).map(issue => [issue.code, issue.path])};
};
const modify = (bytes, mutate) => { const entries = unzipSync(bytes); mutate(entries); return zipSync(entries); };
const text = (entries, path, mutate) => { entries[path] = enc.encode(mutate(dec.decode(entries[path]))); };
const tagValue = bytes => JSON.parse(Buffer.from(dec.decode(bytes).match(/\bval="([^"]+)"/)[1], 'hex').toString('utf8'));
const stripDocument = entries => text(entries, 'ppt/presentation.xml', xml => xml.replace(/<p:custDataLst>[\s\S]*?<\/p:custDataLst>/, ''));
const intent = slide => ({layout: slide.layout, type: slide.type, composition: slide.composition, titleAlignment: slide.design?.titleAlignment, contentBox: slide.design?.contentBox});
const expectedA = [
  {layout: 'gallery-hero', type: undefined, composition: {mode: 'column'}, titleAlignment: 'center', contentBox: false},
  {layout: 'title-subtitle', type: undefined, composition: undefined, titleAlignment: undefined, contentBox: undefined}
];
// Paste slide `index` of another package after the last slide, as PowerPoint
// does: the slide keeps its own OPF_SLIDE_V1 tag, the host keeps its document tag.
const paste = (target, source, index) => modify(target, entries => {
  const from = unzipSync(source);
  const count = Object.keys(entries).filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path)).length, n = count + 1;
  entries[`ppt/slides/slide${n}.xml`] = from[`ppt/slides/slide${index}.xml`];
  entries[`ppt/tags/opfPasted${n}.xml`] = from[`ppt/tags/opfSlide${index}.xml`];
  entries[`ppt/slides/_rels/slide${n}.xml.rels`] = enc.encode(dec.decode(from[`ppt/slides/_rels/slide${index}.xml.rels`])
    .replace(`../tags/opfSlide${index}.xml`, `../tags/opfPasted${n}.xml`).replace(/<Relationship [^>]*notesSlide[^>]*\/>/, ''));
  text(entries, '[Content_Types].xml', xml => xml.replace('</Types>', `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/tags/opfPasted${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/></Types>`));
  text(entries, 'ppt/_rels/presentation.xml.rels', xml => xml.replace('</Relationships>', `<Relationship Id="rIdPasted${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`));
  text(entries, 'ppt/presentation.xml', xml => xml.replace('</p:sldIdLst>', `<p:sldId id="${500 + n}" r:id="rIdPasted${n}"/></p:sldIdLst>`));
});
let cases = 0;

// Package shape: the inline record travels with its slide; bundled ids carry none.
{
  const entries = unzipSync(exportedA);
  const [one, two] = [1, 2].map(index => tagValue(entries[`ppt/tags/opfSlide${index}.xml`]));
  assert.deepEqual({layout: one.layout, type: one.type, composition: one.composition, layoutRecord: one.layoutRecord, design: one.design},
    {layout: 'gallery-hero', type: undefined, composition: {mode: 'column'}, layoutRecord: heroA, design: {titleAlignment: 'center', contentBox: false}});
  assert.equal(two.layout, 'title-subtitle');
  assert.equal(two.layoutRecord, undefined, 'Bundled layouts need no stored record.');
  assert.deepEqual(tagValue(entries['ppt/tags/opfDocument.xml']).catalogs, {layouts: {records: [heroA]}});
  assert.equal(Object.keys(entries).some(path => /opfLayout/i.test(path)), false, 'No separate layout tag part.');
  const refs = unzipSync(await toPptx(structuredClone(deckA), {provenance: 'references-only'}));
  assert.deepEqual(tagValue(refs['ppt/tags/opfSlide1.xml']).layoutRecord, heroA, 'A layout record without sources is a reference.');
  const none = unzipSync(await toPptx(structuredClone(deckA), {provenance: false}));
  assert.equal(Object.keys(none).some(path => /^ppt\/tags\/opf(Document|Slide)/.test(path)), false);
  cases++;
}

// With the document record (FF-32): every layout intent returns, and re-export writes the same records.
{
  const {deck, provenance} = await read(exportedA);
  assert.deepEqual(provenance, []);
  assert.deepEqual(deck.slides.map(intent), expectedA);
  assert.deepEqual(deck.catalogs, {layouts: {records: [heroA]}});
  assert.equal(deck.narrative, 'problem-solution');
  const again = unzipSync(await toPptx(deck)), before = unzipSync(exportedA);
  for (const index of [1, 2]) {
    const {native: _a, ...x} = tagValue(before[`ppt/tags/opfSlide${index}.xml`]), {native: _b, ...y} = tagValue(again[`ppt/tags/opfSlide${index}.xml`]);
    assert.deepEqual(y, x, `slide ${index} provenance is idempotent`);
  }
  cases++;
}

// Without document provenance: slide tags alone restore layout intent and records.
{
  const stripped = await read(modify(exportedA, stripDocument));
  assert.deepEqual(stripped.provenance, []);
  assert.deepEqual(stripped.deck.slides.map(intent), expectedA);
  assert.deepEqual(stripped.deck.catalogs, {layouts: {records: [heroA]}});
  assert.equal(stripped.deck.narrative, undefined, 'Deck metadata needs the document record.');
  assert.equal(stripped.deck.design?.contentAlignment, undefined, 'Deck defaults need the document record.');
  const damaged = await read(modify(exportedA, entries => text(entries, 'ppt/tags/opfDocument.xml', xml => xml.replace(/val="[0-9A-F]{8}/, 'val="ZZZZZZZZ'))));
  assert.deepEqual(damaged.provenance, [['invalid-document-provenance', '']]);
  assert.deepEqual(damaged.deck.slides.map(intent), expectedA);
  assert.deepEqual(damaged.deck.catalogs, {layouts: {records: [heroA]}});
  cases++;
}

// A pasted slide brings the record the host document lacks.
{
  const {deck, provenance} = await read(paste(exportedA, exportedB, 2));
  assert.deepEqual(deck.slides.map(slide => slide.layout), ['gallery-hero', 'title-subtitle', 'gallery-side']);
  assert.deepEqual(deck.catalogs, {layouts: {records: [heroA, sideB]}});
  // The deck default is not restored once slides were added (FF-32).
  assert.deepEqual(provenance, [['design-reference-changed', 'design.contentAlignment']]);
  const stripped = await read(modify(paste(exportedA, exportedB, 2), stripDocument));
  assert.deepEqual(stripped.provenance, []);
  assert.deepEqual(stripped.deck.slides.map(slide => slide.layout), ['gallery-hero', 'title-subtitle', 'gallery-side']);
  assert.deepEqual(stripped.deck.catalogs, {layouts: {records: [heroA, sideB]}});
  cases++;
}

// A pasted slide whose record disagrees for the same id keeps its content without the layout id.
{
  const {deck, provenance} = await read(paste(exportedA, exportedB, 1));
  assert.deepEqual(deck.slides.map(slide => slide.layout), ['gallery-hero', 'title-subtitle', undefined]);
  assert.equal(deck.slides[2].title, 'From B');
  assert.deepEqual(deck.catalogs, {layouts: {records: [heroA]}});
  assert.deepEqual(provenance, [['layout-reference-changed', 'slides.2.layout'], ['design-reference-changed', 'design.contentAlignment']]);
  const stripped = await read(modify(paste(exportedA, exportedB, 1), stripDocument));
  assert.deepEqual(stripped.deck.slides.map(slide => slide.layout), ['gallery-hero', 'title-subtitle', undefined]);
  assert.deepEqual(stripped.deck.catalogs, {layouts: {records: [heroA]}});
  assert.deepEqual(stripped.provenance, [['layout-reference-changed', 'slides.2.layout']]);
  cases++;
}

// Tampering and edits report the specific field; the rest of the slide intent is unaffected.
{
  const retag = (entries, mutate) => text(entries, 'ppt/tags/opfSlide1.xml', xml => {
    const value = JSON.parse(Buffer.from(xml.match(/\bval="([^"]+)"/)[1], 'hex').toString('utf8'));
    mutate(value);
    return xml.replace(/val="[0-9A-F]+"/, `val="${Buffer.from(JSON.stringify(value)).toString('hex').toUpperCase()}"`);
  });
  const mismatched = await read(modify(exportedA, entries => { stripDocument(entries); retag(entries, value => { value.layoutRecord.id = 'other'; }); }));
  assert.deepEqual(mismatched.provenance, [['invalid-document-provenance', 'slides.0.layoutRecord']]);
  assert.equal(mismatched.deck.slides[0].layout, 'gallery-hero');
  assert.equal(mismatched.deck.catalogs, undefined);
  // The id itself still returns (an unknown layout id only warns in validation), as does the rest of the intent.
  assert.equal(mismatched.deck.slides[0].composition?.mode, 'column');
  const malformed = await read(modify(exportedA, entries => { stripDocument(entries); retag(entries, value => { value.layoutRecord = 'gallery-hero'; }); }));
  assert.deepEqual(malformed.provenance, [['invalid-document-provenance', 'slides.0']]);
  assert.equal(malformed.deck.slides[0].layout, undefined);
  assert.equal(malformed.deck.slides[1].layout, 'title-subtitle');
  const moved = await read(modify(exportedA, entries => { stripDocument(entries); text(entries, 'ppt/slides/slide1.xml', xml => xml.replace(/(<p:sp>[\s\S]*?<a:off x=")(\d+)"/, (_, before, x) => `${before}${Number(x) + 12700}"`)); }));
  assert.deepEqual(moved.provenance, [['layout-reference-changed', 'slides.0.layout'], ['slide-reference-changed', 'slides.0.composition'],
    ['design-reference-changed', 'slides.0.design.titleAlignment'], ['design-reference-changed', 'slides.0.design.contentBox']]);
  assert.deepEqual(moved.deck.slides.map(slide => slide.layout), [undefined, 'title-subtitle']);
  assert.equal(moved.deck.catalogs, undefined);
  cases++;
}

console.log(`Layout intent passed: ${cases} groups; OPF_SLIDE_V1 carries the layout record, and import restores layout, type, composition, hints and records with and without document provenance, for pasted slides, conflicting records, tampering and edits.`);
