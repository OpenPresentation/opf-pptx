import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {unzipSync} from 'fflate';
import * as opf from '@openpresentation/opf';
import {examples} from '@openpresentation/opf/examples';
import {resolvePresentation} from '@openpresentation/opf-render';
import {fromPptx, toPptx, inventoryPptxTypefaces} from '../dist/index.js';

// FF-49 / FF-50 (font-fidelity-everywhere). The presentation theme's major and
// minor `a:ea` and `a:cs` typefaces are never empty, for every font scheme and
// every language (Model C, core docs script-font-model.md):
// - a slot that the scheme or the language supplies names that script font;
// - every other slot repeats the theme's own latin face, the family the author
//   selected (never an open replacement and never a font nobody chose);
// - the language never changes the theme's latin fonts.
// The preview resolves the same slots (opf-render's per-slide script profile),
// the export is deterministic, and re-import keeps the scheme and language.

const {catalogs, resolveScriptFonts} = opf;
const slide = text => ({title: text, text});
const deck = (fontScheme, language, extra = {}) => ({name: 'Theme slots', ...(language === undefined ? {} : {language}), ...(fontScheme === undefined ? {} : {design: {fontScheme}}), ...extra, slides: [slide('Heading'), {title: 'Second', items: ['One', 'Two']}]});

const exported = async (presentation, options = {}) => {
  const diagnostics = [];
  const bytes = await toPptx(structuredClone(presentation), {...options, onDiagnostic: diagnostic => diagnostics.push(diagnostic)});
  return {bytes, diagnostics, inventory: inventoryPptxTypefaces(bytes)};
};
// The presentation package's own theme (prefix ''), slots as {latin, ea, cs}.
const theme = inventory => ({major: inventory.themes[''].major, minor: inventory.themes[''].minor});
const previewProfile = presentation => resolvePresentation(structuredClone(presentation)).slides[0].scriptFonts.profile;
const families = new Set(catalogs.fontSchemes.flatMap(record => [record.major, record.minor]));

const fontSchemeIds = catalogs.fontSchemes.map(record => record.id);
assert.ok(fontSchemeIds.length >= 89, 'the bundled catalog has the 89 upstream font schemes');

// 1. Every font scheme, in the default language: no empty slot, all slots the scheme's own families.
for (const record of catalogs.fontSchemes) {
  const {inventory, diagnostics} = await exported(deck(record.id));
  const {major, minor} = theme(inventory);
  const want = {major: {latin: record.major, ea: record.major, cs: record.major}, minor: {latin: record.minor, ea: record.minor, cs: record.minor}};
  assert.deepEqual({major, minor}, want, `${record.id}: theme slots repeat the scheme families`);
  assert.deepEqual(diagnostics.filter(diagnostic => /script|language/.test(diagnostic.code)), [], `${record.id}: no script diagnostics`);
  // Preview and export agree slot by slot.
  const profile = previewProfile(deck(record.id));
  assert.deepEqual([profile.heading.latin, profile.heading.eastAsian, profile.heading.complexScript], [major.latin, major.ea, major.cs], `${record.id}: preview heading slots`);
  assert.deepEqual([profile.body.latin, profile.body.eastAsian, profile.body.complexScript], [minor.latin, minor.ea, minor.cs], `${record.id}: preview body slots`);
}

// 2. Schemes x languages: Latin-only, CJK, Arabic, Hebrew, Indic, Thai and the Latin-slot scripts written with their own font.
const languages = [
  ['english-us', 'latin'], ['french', 'latin'], ['russian', 'latin'], ['greek', 'latin'],
  ['armenian', 'latin'], ['georgian', 'latin'], ['amharic', 'latin'],
  ['japanese', 'eastAsian'], ['chinese-simplified', 'eastAsian'], ['chinese-traditional', 'eastAsian'], ['korean', 'eastAsian'],
  ['arabic', 'complexScript'], ['hebrew', 'complexScript'], ['persian', 'complexScript'], ['urdu', 'complexScript'],
  ['hindi', 'complexScript'], ['bengali', 'complexScript'], ['tamil', 'complexScript'], ['thai', 'complexScript'], ['khmer', 'complexScript'],
];
const schemes = ['aptos', 'calibri', 'georgia', 'meiryo', 'arabic-typesetting', 'mangal'];
for (const scheme of schemes) {
  const record = catalogs.fontSchemes.find(entry => entry.id === scheme);
  for (const [language, role] of languages) {
    const presentation = deck(scheme, language);
    const label = `${scheme} + ${language}`;
    const resolved = resolveScriptFonts(presentation);
    assert.equal(resolved.scriptRole, role, `${label}: script role`);
    const {inventory, diagnostics, bytes} = await exported(presentation);
    const {major, minor} = theme(inventory);
    // Model C: the language never replaces the Latin scheme.
    assert.deepEqual([major.latin, minor.latin], [record.major, record.minor], `${label}: latin stays the scheme`);
    // Never empty; every slot names a scheme family or the language's own script font.
    const language_ = catalogs.languages.find(entry => entry.id === language);
    const languageFont = catalogs.fontSchemes.find(entry => entry.id === language_.fontScheme);
    const allowed = new Set([record.major, record.minor, languageFont?.major, languageFont?.minor]);
    for (const [name, slots] of [['major', major], ['minor', minor]]) for (const slot of ['ea', 'cs']) {
      assert.ok(slots[slot], `${label}: ${name} ${slot} is not empty`);
      assert.ok(allowed.has(slots[slot]), `${label}: ${name} ${slot} "${slots[slot]}" is a selected family`);
    }
    // The slot the language's script uses names the language's script font when the language supplies one.
    if (role !== 'latin' && resolved.sources[role] !== 'latin') {
      const key = role === 'eastAsian' ? 'ea' : 'cs';
      assert.deepEqual([major[key], minor[key]], [resolved.heading[role], resolved.body[role]], `${label}: ${key} names the script font`);
    }
    // Preview and export agree slot by slot.
    const profile = previewProfile(presentation);
    assert.deepEqual([profile.heading.eastAsian, profile.heading.complexScript, profile.body.eastAsian, profile.body.complexScript], [major.ea, major.cs, minor.ea, minor.cs], `${label}: preview slots equal the theme`);
    assert.equal(profile.lang, resolved.lang, `${label}: preview lang`);
    assert.equal(profile.rtl, resolved.rtl, `${label}: preview direction`);
    // Deterministic, and the scheme and language survive re-import with no script-font diagnostic.
    assert.deepEqual((await exported(presentation)).bytes, bytes, `${label}: deterministic`);
    const imported = [];
    const restored = await fromPptx(bytes, {onDiagnostic: diagnostic => imported.push(diagnostic.code)});
    assert.equal(restored.language, language, `${label}: language re-imports`);
    assert.equal(restored.design?.fontScheme, scheme, `${label}: font scheme re-imports`);
    // language-ambiguous is the pre-existing note for records that share one OOXML tag (bengali and chittagonian, bn-BD); the stored id still wins.
    assert.deepEqual(imported.filter(code => /script-font|language|rtl/.test(code) && code !== 'language-ambiguous'), [], `${label}: import diagnostics`);
    assert.deepEqual(diagnostics.filter(diagnostic => /script|language/.test(diagnostic.code)), [], `${label}: export diagnostics`);
  }
}

// 3. Re-import then export again reproduces the same theme slots (FF-32 provenance restores the scheme and language).
for (const [scheme, language] of [['meiryo', 'japanese'], ['calibri', 'english-us'], ['aptos', 'arabic'], ['mangal', 'hindi'], ['georgia', 'hebrew']]) {
  const {bytes, inventory} = await exported(deck(scheme, language));
  const restored = await fromPptx(bytes);
  const again = await exported(restored);
  assert.deepEqual(theme(again.inventory), theme(inventory), `${scheme} + ${language}: re-import exports the same theme slots`);
}

// 4. Inline (gallery legacy) font scheme records fill the slots too, and an explicit script slot still wins.
{
  const record = {id: 'legacy-face', name: 'Legacy face', app: 'PowerPoint', languageFamily: 'latin', languages: [], major: 'Legacy Head', minor: 'Legacy Body', type: 'sans-serif'};
  const {inventory} = await exported(deck('legacy-face', undefined, {catalogs: {fontSchemes: {records: [record]}}}));
  assert.deepEqual(theme(inventory), {major: {latin: 'Legacy Head', ea: 'Legacy Head', cs: 'Legacy Head'}, minor: {latin: 'Legacy Body', ea: 'Legacy Body', cs: 'Legacy Body'}});
  const explicit = {id: 'aptos', eastAsian: {major: 'Noto Sans JP', minor: 'Noto Sans JP'}, complexScript: {major: 'Noto Naskh Arabic', minor: 'Noto Naskh Arabic'}};
  const slots = theme((await exported(deck(explicit, 'english-us'))).inventory);
  assert.deepEqual([slots.major.ea, slots.major.cs, slots.minor.ea, slots.minor.cs], ['Noto Sans JP', 'Noto Naskh Arabic', 'Noto Sans JP', 'Noto Naskh Arabic']);
}

// 5. Run-level slots are unchanged: runs still repeat their own latin face in ea/cs for a Latin deck.
{
  const {bytes} = await exported(deck('calibri'));
  const slideXml = new TextDecoder().decode(unzipSync(bytes)['ppt/slides/slide1.xml']);
  const triples = [...slideXml.matchAll(/<a:latin typeface="([^"+][^"]*)"[^>]*\/>\s*<a:ea typeface="([^"]*)"[^>]*\/>\s*<a:cs\s+typeface="([^"]*)"/g)];
  assert.ok(triples.length > 0);
  for (const [, latin, ea, cs] of triples) assert.deepEqual([ea, cs], [latin, latin]);
}

// 6. No theme slot names an open replacement: the slots of a Calibri + Japanese deck are catalog families only.
{
  const {inventory} = await exported(deck('calibri', 'japanese'));
  for (const slots of Object.values(theme(inventory))) for (const face of Object.values(slots)) assert.ok(families.has(face), `${face} is a catalog family`);
}

// 7. The example corpus: every deck's presentation theme names all six slots.
// Structural gate: assets are substituted, as in test/export-corpus.mjs.
const image = new Uint8Array(await readFile(new URL('./fixtures/images/wide.png', import.meta.url)));
let corpus = 0;
for (const {file, deck} of examples) {
  const {major, minor} = theme((await exported(deck, {imageResolver: async () => image})).inventory);
  for (const slots of [major, minor]) for (const slot of ['latin', 'ea', 'cs']) assert.ok(slots[slot], `${file}: theme ${slot} is not empty`);
  corpus += 1;
}

console.log(JSON.stringify({test: 'theme-script-slots', passed: true, fontSchemes: fontSchemeIds.length, combinations: schemes.length * languages.length, corpusDecks: corpus}));
