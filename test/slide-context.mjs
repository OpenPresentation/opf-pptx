import assert from 'node:assert/strict';
import {strFromU8, unzipSync} from 'fflate';
import {resolveSlideContext} from '@openpresentation/opf';
import {toPptx} from '../dist/index.js';

// RR-55: the exporter composes every slide with the options of core's resolveSlideContext, so the deck's resolved font
// families (slide design, then deck design, then theme, then the engine default) are the ones the layout is measured in.
// A composeSlide option the engine does not know is ignored without an error, so the failure this test guards is silent:
// composition falling back to the default families while the package names the chosen ones.

const textSlide = {title: 'Heading text', subtitle: 'Subtitle text', text: 'Body copy of the slide.', items: ['One', 'Two']};
const deck = {
  name: 'Slide context',
  design: {fontScheme: 'roboto'},
  slides: [
    {id: 'inherits', ...textSlide},
    {id: 'own-scheme', ...textSlide, design: {fontScheme: 'georgia'}},
    {id: 'code', layout: 'code-1x', title: 'Code', code: {source: 'const x = 1;', language: 'ts'}},
  ],
};

// A measurement provider that records the family of every style the layout asks about.
const measuring = () => {
  const families = new Set();
  let calls = 0;
  return {families, calls: () => calls, textMeasurement: {measure: (text, size, style) => { calls++; families.add(style.fontFamily); return text.length * size * 0.5; }}};
};

for (const [index, slide] of deck.slides.entries()) {
  const expected = resolveSlideContext(deck, index).options.fontFamilies;
  const probe = measuring();
  const bytes = await toPptx({...deck, slides: [slide]}, {fonts: {textMeasurement: probe.textMeasurement}, strictAssets: true});
  assert.ok(probe.calls() > 0, `${slide.id}: the exporter measures with the fonts handle`);
  const allowed = new Set([expected.heading, expected.body, expected.code, expected.accent].filter(Boolean));
  assert.deepEqual([...probe.families].filter(family => !allowed.has(family)), [], `${slide.id}: layout is measured only in the deck's resolved families`);
  assert.ok(probe.families.has(expected.body), `${slide.id}: the body family ${expected.body} is measured`);
  // The package names the same families.
  const entries = unzipSync(bytes);
  const faces = new Set(Object.entries(entries).filter(([name]) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .flatMap(([, value]) => [...strFromU8(value).matchAll(/<a:latin typeface="([^"]*)"/g)].map(match => match[1])));
  assert.deepEqual([...faces].filter(family => !allowed.has(family)), [], `${slide.id}: the package names only the resolved families`);
}

// The families are the scheme's, not the default (Aptos) the engine uses when it is not told: roboto for the deck,
// georgia for the slide that sets its own scheme.
assert.equal(resolveSlideContext(deck, 0).options.fontFamilies.body, 'Roboto');
assert.notEqual(resolveSlideContext(deck, 1).options.fontFamilies.body, 'Roboto');
const georgia = measuring();
await toPptx({...deck, slides: [deck.slides[1]]}, {fonts: {textMeasurement: georgia.textMeasurement}});
assert.ok(!georgia.families.has('Aptos') && !georgia.families.has('Roboto'), 'a slide that sets its own scheme is not measured in the deck or default families');

// Without a fonts handle the layout uses core's portable estimate and the same families are named.
const estimated = unzipSync(await toPptx({...deck, slides: [deck.slides[1]]}));
assert.ok(strFromU8(estimated['ppt/slides/slide1.xml']).includes(`typeface="${resolveSlideContext(deck, 1).options.fontFamilies.body}"`));

// A reference that matches no record is reported once per place it is written, from core's context.
{
  const diagnostics = [];
  await toPptx({name: 'Unknown theme', design: {theme: 'no-such-theme'}, slides: [{title: 'One', text: 'a'}, {title: 'Two', text: 'b'}]}, {onDiagnostic: diagnostic => diagnostics.push(diagnostic)});
  const unresolved = diagnostics.filter(diagnostic => diagnostic.code === 'unresolved-theme');
  assert.deepEqual(unresolved.map(({path, id}) => ({path, id})), [{path: 'design.theme', id: 'no-such-theme'}]);
}

// An unresolved layout is refused, as before.
await assert.rejects(() => toPptx({name: 'Unknown layout', slides: [{layout: 'no-such-layout', title: 'x'}]}), {code: 'catalog-resolution-failed', path: 'slides.0.layout'});

console.log(JSON.stringify({test: 'slide-context', passed: true}));
