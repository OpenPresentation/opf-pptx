import assert from 'node:assert/strict';
import * as opf from '@openpresentation/opf';
import {toPptx} from '../dist/index.js';

// opf-pptx#168. A slide may select its own script fonts (slides[].design.fontScheme, a slide theme, or an inline
// fontScheme with eastAsian/complexScript), and core resolves them per slide. Runs carry no explicit ea/cs (FF-05), so
// a slide's script fonts can only reach PowerPoint through the theme. A slide whose selected script fonts the
// presentation theme (slide 1's) does not carry is reported, never silently drawn in slide 1's fonts.

if (typeof opf.resolveScriptFonts !== 'function') {
  console.log('Per-slide script fonts: core has no resolveScriptFonts; nothing is resolved per slide.');
  process.exit(0);
}

const exported = async presentation => {
  const diagnostics = [];
  const bytes = await toPptx(structuredClone(presentation), {onDiagnostic: diagnostic => diagnostics.push(diagnostic)});
  return {bytes, diagnostics, perSlide: diagnostics.filter(diagnostic => diagnostic.code === 'script-font-per-slide-not-exported')};
};
const thai = 'ภาษาไทยเป็นภาษาที่มีวรรณยุกต์';
const inline = (slot, family) => ({major: 'Arial', minor: 'Arial', [slot]: {major: family, minor: family}});

// The opf-pptx#168 repro: Thai, slide 1 Angsana New, slide 2 DilleniaUPC.
const repro = {name: 'Per-slide script font', language: 'th', design: {theme: 'classic', dimensions: 'widescreen'}, slides: [
  {id: 'one', title: 'รายงานสรุปผล', design: {fontScheme: 'angsana-new'}, blocks: [{text: thai}]},
  {id: 'two', title: 'รายงานสรุปผล', design: {fontScheme: 'dilleniaupc'}, blocks: [{text: thai}]},
]};
// Japanese, slide 1 Meiryo, slide 2 MS Mincho.
const japanese = {name: 'Per-slide East Asian font', language: 'japanese', slides: [
  {title: '日本語の見出し', design: {fontScheme: 'meiryo'}, text: 'ひらがなとカタカナ'},
  {title: '日本語の見出し', design: {fontScheme: 'ms-mincho'}, text: 'ひらがなとカタカナ'},
]};
// Three inline script profiles: East Asian Meiryo, complex-script Traditional Arabic, complex-script Nirmala UI.
const threeProfiles = {name: 'Three script profiles', slides: [
  {title: '日本語', design: {fontScheme: inline('eastAsian', 'Meiryo')}, text: 'ひらがな'},
  {title: 'مرحبا', design: {fontScheme: inline('complexScript', 'Traditional Arabic')}, text: 'مرحبا بالعالم'},
  {title: 'नमस्ते', design: {fontScheme: inline('complexScript', 'Nirmala UI')}, text: 'नमस्ते दुनिया'},
]};

{
  const {perSlide} = await exported(repro);
  assert.deepEqual(perSlide.map(diagnostic => diagnostic.path), ['slides.1.design.fontScheme'], 'repro: slide 2 is reported');
  assert.match(perSlide[0].message, /DilleniaUPC/);
}
{
  const {perSlide} = await exported(japanese);
  assert.deepEqual(perSlide.map(diagnostic => diagnostic.path), ['slides.1.design.fontScheme'], 'Japanese: slide 2 is reported');
  assert.match(perSlide[0].message, /East Asian font "MS Mincho"/);
}
{
  const {perSlide} = await exported(threeProfiles);
  assert.deepEqual(perSlide.map(diagnostic => diagnostic.path), ['slides.1.design.fontScheme', 'slides.2.design.fontScheme'], 'three profiles: slides 2 and 3 are reported');
}

// Not reported: slides whose script fonts the theme carries, slides that select no script font, Latin-only overrides.
for (const [label, presentation] of [
  ['same profile through the language', {name: 'Same', language: 'th', design: {theme: 'classic'}, slides: [{title: thai, design: {fontScheme: 'angsana-new'}}, {title: thai}]}],
  ['Latin-only per-slide scheme', {name: 'Latin', slides: [{title: 'One', design: {fontScheme: 'georgia'}}, {title: 'Two', design: {fontScheme: 'arial'}}]}],
  ['later slide selects nothing', {name: 'Nothing', slides: [{title: '日本語', design: {fontScheme: inline('eastAsian', 'Meiryo')}}, {title: 'Two'}]}],
  ['one slide', {name: 'One', language: 'japanese', slides: [{title: '日本語', design: {fontScheme: 'ms-mincho'}}]}],
]) {
  const {perSlide} = await exported(presentation);
  assert.deepEqual(perSlide, [], `${label}: not reported`);
}

console.log('Per-slide script fonts: slides whose selected East Asian/complex-script fonts the theme does not carry are reported (opf-pptx#168).');
