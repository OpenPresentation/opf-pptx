import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {createRequire} from 'node:module';
import path from 'node:path';
import {hash, inspectArchive, compareNoopText, nativeFindings, importedFurnitureFindings} from '../scripts/keynote-qualification/pack.mjs';
import {sourceDeck} from '../scripts/keynote-qualification/fixtures.mjs';

const resolver = createRequire(path.join(process.env.OPF_KEYNOTE_TEST_CONSUMER ?? process.cwd(), 'package.json'));
const dependencies = createRequire(resolver.resolve('@openpresentation/opf-pptx/package.json'));
const engine = {fflate: dependencies('fflate'), xml: dependencies('fast-xml-parser')};
const directory = new URL('./fixtures/keynote-noop/', import.meta.url), manifest = JSON.parse(await readFile(new URL('manifest.json', directory)));
const controls = [];
for (const file of manifest.fixtures) {
  const bytes = await readFile(new URL(file.file, directory));
  assert.equal(bytes.length, file.bytes); assert.equal(hash(bytes), file.sha256, 'Retained native bytes must stay exact');
  controls.push(inspectArchive(bytes, engine));
}
const [baseline, native] = controls;
assert.equal(baseline.slides.length, 16); assert.equal(native.slides.length, 16);
const findings = nativeFindings(native, 'no-op', baseline);
assert.equal(findings.length, 16);
assert.ok(findings.every(value => /K\d{2} notes: no-op dynamic field types\/counts changed/.test(value)), 'Only notes-page field removal is an archive no-op finding');
for (let index = 0; index < 16; index++) {
  assert.deepEqual(compareNoopText(native.slides[index].textSemantics, baseline.slides[index].textSemantics, 'slide'), []);
  assert.deepEqual(native.slides[index].textSemantics.fields, {slidenum: 1});
  assert.deepEqual(baseline.slides[index].notesSemantics.fields, {slidenum: 1});
  assert.deepEqual(native.slides[index].notesSemantics.fields, {});
}
const before = baseline.slides[1].textSemantics, after = native.slides[1].textSemantics;
assert.equal(after.objects[0].name, 'OPF furniture 1 part 0 line 0');
assert.equal(after.objects[0].placeholderRole, null);
assert.equal(before.objects[0].placeholderRole, 'title');
const mutate = change => {const value = structuredClone(after); change(value); return compareNoopText(value, before, 'K02');};
assert.ok(mutate(value => {value.objects[2].paragraphs[0] = 'KN_ORIGINAL_PARAGRAPH. Keep sentence.';}).some(value => value.includes('ordered authored text changed')), 'Real word loss fails');
assert.ok(mutate(value => {value.objects[2].paragraphs[0] = 'KN_ORIGINAL_PARAGRAPH. sentence second this Keep.';}).some(value => value.includes('ordered authored text changed')), 'Same words in the wrong order fail');
assert.ok(mutate(value => {[value.objects[2], value.objects[3]] = [value.objects[3], value.objects[2]];}).some(value => value.includes('ordered authored text changed')), 'Paragraph reordering fails');
assert.ok(mutate(value => {value.objects.push(structuredClone(value.objects[0]));}).some(value => value.includes('literal furniture changed')), 'Duplicate footer fails');
assert.ok(mutate(value => {value.objects[0].paragraphs[0] = 'Changed footer';}).some(value => value.includes('literal furniture changed')), 'Changed literal footer fails');
assert.ok(mutate(value => {value.objects[0].name = 'unrelated body';}).some(value => value.includes('ordered authored text changed')), 'An unknown name without a footer role is not silently discarded');
assert.ok(mutate(value => {value.fields = {datetime: 1};}).some(value => value.includes('dynamic field types/counts changed')), 'Wrong field type fails even with the same count');
assert.ok(mutate(value => {value.fields = {};}).some(value => value.includes('dynamic field types/counts changed')), 'A missing slide-number field fails');
const badNotes = structuredClone(native.slides[1].notesSemantics); badNotes.objects[1].paragraphs[0] = 'K02 wrong notes';
assert.ok(compareNoopText(badNotes, baseline.slides[1].notesSemantics, 'K02 notes').some(value => value.includes('ordered authored text changed')), 'Wrong notes fail independently of generated notes-page fields');

// Exercise actual import as well: current-content authority exposes, rather than hides, the footer adaptation.
const {fromPptx} = await import(resolver.resolve('@openpresentation/opf-pptx'));
const imported = await fromPptx(await readFile(new URL('keynote-15.1.1-noop.pptx', directory)));
const furniture = importedFurnitureFindings(imported, sourceDeck('data:image/png;base64,AA=='), baseline);
assert.equal(furniture.length, 16, 'Current importer gap remains an explicit finding; no runtime fix or footer-word filtering is part of this change');
assert.ok(furniture.every(value => value.includes('literal furniture appears in current imported slide content')));
console.log('Keynote retained no-op: ordered words/notes and literal footer pass; slide fields survive without cached text; 16 notes-field removals and 16 current-import furniture gaps remain explicit. Word loss/reorder, wrong notes, duplicate/changed/unidentified footer and wrong/missing fields fail.');
