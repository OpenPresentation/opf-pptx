import assert from 'node:assert/strict';
import {unzipSync, zipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {validate} from '@openpresentation/opf';

// FA-04: built-in variables resolve before export. FA-31: in a header or footer `text` they are ordinary words
// (`{{organization.name}}`, `{{speaker.name}}, {{speaker.title}}`), and the stored document record returns the metadata.
const dec = new TextDecoder(), enc = new TextEncoder();
const slideXml = (entries, index = 1) => dec.decode(entries[`ppt/slides/slide${index}.xml`]);
const modify = (bytes, mutate) => { const entries = unzipSync(bytes); mutate(entries); return zipSync(entries); };
const read = async (bytes) => { const issues = []; const deck = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)}); assert.equal(validate(deck, {only: ['format']}).valid, true, JSON.stringify(validate(deck, {only: ['format']}).findings)); return {deck, issues}; };

const source = {
  name: 'Q4 Review',
  speaker: [{id: 'ada', name: 'Ada Lovelace', title: 'CTO'}, {id: 'grace', name: 'Grace Hopper'}],
  organization: {id: 'acme', name: 'Acme Corp', tagline: 'Build the future'},
  design: {footer: {left: {text: '{{organization.name}}\n{{speaker.name}}, {{speaker.title}}'}, right: {text: '{{organization.tagline}}'}}},
  slides: [{id: 'cover', title: '{{deck.name}}', subtitle: '{{speaker.name}}, {{speaker.title}} · {{organization.name}}'}, {id: 'next', title: 'By {{speakers}}', text: 'Body'}]
};
const before = structuredClone(source);
const bytes = await toPptx(source);
assert.deepEqual(source, before, 'Export leaves the source unchanged.');
assert.deepEqual(await toPptx(source), bytes, 'Export is deterministic.');
const entries = unzipSync(bytes), xml = slideXml(entries), second = slideXml(entries, 2);
for (const text of ['Q4 Review', 'Ada Lovelace, CTO · Acme Corp', 'Build the future', 'Acme Corp', 'Ada Lovelace, CTO']) assert.ok(xml.includes(`<a:t>${text}</a:t>`), text);
assert.ok(second.includes('<a:t>By Ada Lovelace, Grace Hopper</a:t>'), 'the speakers list joins');
assert.ok(!xml.includes('{{'), 'no unresolved token reaches the file');
// FA-31: the footer text is one string of the author's own words. Both values resolve before export, so the file holds fixed
// words in the order written (organization line, then speaker line); nothing is a field.
assert.ok(xml.indexOf('<a:t>Acme Corp</a:t>') < xml.indexOf('<a:t>Ada Lovelace, CTO</a:t>', xml.indexOf('<a:t>Acme Corp</a:t>')));
assert.ok(!xml.includes('type="slidenum"'));

// Import returns the words as the zone text; the stored document record (not the footer) returns the speakers and organization.
const {deck} = await read(bytes);
assert.equal(deck.design.footer.left.text, 'Acme Corp\nAda Lovelace, CTO');
assert.deepEqual(Object.keys(deck.design.footer.left), ['text'], 'no organization or speaker flag exists any more');
assert.deepEqual(deck.speaker, source.speaker, 'Stored speakers return from the document record.');
assert.deepEqual(deck.organization, source.organization, 'The stored organization returns.');
assert.ok(!(await read(bytes)).issues.some(issue => issue.code === 'invalid-furniture-provenance'));

// Without document provenance no speaker or organization metadata is rebuilt from the footer words.
const plain = await read(await toPptx(source, {provenance: false}));
assert.equal(plain.deck.speaker, undefined);
assert.equal(plain.deck.organization, undefined);
assert.equal(plain.deck.design.footer.left.text, 'Acme Corp\nAda Lovelace, CTO');
// A speaker with no title is the name alone.
const noTitle = structuredClone(source);
noTitle.speaker = {id: 'solo', name: 'Solo, Jr.'};
noTitle.design.footer.left.text = '{{organization.name}}\n{{speaker.name}}';
assert.equal((await read(await toPptx(noTitle, {provenance: false}))).deck.design.footer.left.text, 'Acme Corp\nSolo, Jr.');

// Current native words win: an edited line is the footer's new text, and the stored speakers stay as exported.
const retitled = await read(modify(bytes, entries => { for (const index of [1, 2]) entries[`ppt/slides/slide${index}.xml`] = enc.encode(slideXml(entries, index).replace('<a:t>Ada Lovelace, CTO</a:t>', '<a:t>Ada Lovelace, CEO</a:t>')); }));
assert.equal(retitled.deck.design.footer.left.text, 'Acme Corp\nAda Lovelace, CEO');
assert.deepEqual(retitled.deck.speaker, source.speaker, 'The footer no longer names the speaker, so the stored speakers are unchanged.');

// A missing source: the field is unresolved content, a built-in resolves to nothing and is reported.
const bare = structuredClone(source);
delete bare.speaker;
bare.slides = [{id: 'cover', title: 'By {{speaker.name}}'}];
const seen = [];
await toPptx(bare, {onDiagnostic: entry => seen.push(entry.code)});
assert.ok(seen.includes('variable-builtin-missing'));
console.log('PPTX built-in variables passed: resolution before export, footer variable text, provenance round trip, edits and missing sources.');
