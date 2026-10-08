// FA-23: the 15 pptx.gallery image treatments (core docs/image-treatments.md) as OPF 0.15 image blocks and image
// backgrounds keep the native frames and paint order 0.14 exported for them as design.slideImage
// (test/fixtures/image-treatments-0.14.json, recorded from opf-pptx 0.14.0). Five treatments need a 0.14 frame that 0.15
// has no form for (a slide-wide frame inside the padding, a recolored or aspect-ratio background); they are listed in
// CHANGED with what 0.15 draws instead.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {unzipSync} from 'fflate';
import {XMLParser} from 'fast-xml-parser';
import {defaultCatalog} from '@openpresentation/opf/catalog';
import {toPptx} from '../dist/index.js';

const before = JSON.parse(await readFile(new URL('fixtures/image-treatments-0.14.json', import.meta.url), 'utf8'));
const hero = {src: `data:image/png;base64,${(await readFile(new URL('fixtures/images/treatment-hero.png', import.meta.url))).toString('base64')}`, alt: 'Project-authored posterized landscape with a sun', mediaType: 'image/png'};
const block = fields => ({type: 'image', image: 'asset:hero', ...fields});
const background = fields => ({type: 'image', src: 'asset:hero', ...fields});
const heading = id => ({title: id.split('-').map(word => word[0].toUpperCase() + word.slice(1)).join('-'), subtitle: `pptx.gallery image-treatments/${id}`});
// The 0.15 form of each reference slide (core docs/image-treatments.md, 0.14 design.slideImage on the right of each line).
const SLIDES = {
  'full-bleed': {design: {background: background({overlay: {color: 'dark1', opacity: 0.2}})}},                      // background, crop, overlay
  'text-overlay': {design: {background: background({overlay: {color: 'dark1', opacity: 0.55}})}},
  'side-by-side': {blocks: [block({placement: {edge: 'left', size: 0.46}})]},                                         // left 0.46 crop
  'masked-shape': {blocks: [block({placement: {edge: 'right', inset: true}, shape: 'hexagon', aspectRatio: 1.1547})]},
  'circular-crop': {blocks: [block({placement: {edge: 'left', size: 0.4, inset: true}, shape: 'circle'})]},
  'background-blur': {design: {contentBox: true, background: background({overlay: {color: 'light1', opacity: 0.4}})}},
  'image-strip': {blocks: [block({placement: {edge: 'bottom', size: 0.3}})]},
  'collage-grid': {design: {imageFit: 'cover'}, composition: {mode: 'grid', columns: 2}, blocks: [block({}), block({}), block({}), block({})]},
  'device-frame': {blocks: [block({placement: {edge: 'right', inset: true}, aspectRatio: 0.5, shape: 'rounded', cornerRadius: 0.12, border: {color: 'dark1', width: 12}})]},
  'cutout-subject': {blocks: [block({placement: {edge: 'right'}, fit: 'contain'})]},
  // CHANGED: 0.15 has no background inset, background mask or line, background recolor or background aspect ratio.
  'caption-overlay': {design: {background: background({overlay: {color: 'dark1', opacity: 0.75, edge: 'bottom', size: 0.25}})}},
  'rounded-card': {design: {background: background({})}},
  duotone: {design: {background: background({})}},
  watermark: {design: {background: background({fit: 'contain', opacity: 0.1})}},
  'cinematic-crop': {design: {background: background({})}}
};
const CHANGED = {
  'caption-overlay': 'the picture fills the slide instead of the slide inside its padding; the band covers the bottom quarter of the slide',
  'rounded-card': 'the picture fills the slide without the rounded mask and border',
  duotone: 'the picture is not recolored',
  watermark: 'the picture fills the slide (contain) without the grayscale recolor',
  'cinematic-crop': 'the picture covers the slide instead of a centered 2.39:1 band'
};
const deck = {assets: {hero}, design: {theme: 'classic'}, slides: Object.entries(SLIDES).map(([id, slide]) => ({id, ...(id === 'collage-grid' ? {title: 'Collage Grid'} : heading(id)), ...slide}))};
const bytes = await toPptx(deck, {imageFormat: 'preserve', strictAssets: true, provenance: false, catalogs: [defaultCatalog]});
const entries = unzipSync(bytes);
const flat = new XMLParser({ignoreAttributes: false, attributeNamePrefix: '', parseTagValue: false});
// 0.14's names for the same objects, so the paint order compares as is.
const legacy = (name, index) => name.replace(/^OPF background slides\.\d+ overlay$/, `OPF slide image overlay slides.${index}`).replace(/^OPF background overlay slides\.\d+$/, `OPF slide image overlay slides.${index}`)
  .replace(/^OPF background slides\.\d+$/, `OPF slide image slides.${index}`);
const near = (a, b) => Math.abs(Number(a) - Number(b)) <= 1;
const sameNumbers = (actual, expected, label) => {
  const flatten = (value, prefix = '') => value && typeof value === 'object' ? Object.entries(value).flatMap(([key, item]) => flatten(item, `${prefix}.${key}`)) : [[prefix, value]];
  const a = Object.fromEntries(flatten(actual)), e = Object.fromEntries(flatten(expected));
  assert.deepEqual(Object.keys(a).sort(), Object.keys(e).sort(), label);
  for (const key of Object.keys(e)) assert.ok(/^-?\d+$/.test(e[key]) ? near(a[key], e[key]) : a[key] === e[key], `${label}${key}: ${a[key]} vs ${e[key]}`);
};
let same = 0, changed = 0;
deck.slides.forEach((slide, index) => {
  const xml = new TextDecoder().decode(entries[`ppt/slides/slide${index + 1}.xml`]);
  const expected = before[slide.id];
  const pictures = [...xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)].map(([p]) => flat.parse(p)['p:pic']);
  if (CHANGED[slide.id]) { assert.ok(pictures.length <= 1, slide.id); changed++; return; }
  // Paint order: the picture, then its overlay directly above it, as in 0.14. A background stays first; a placed block now
  // follows the headings, in core's item order (headings, placed blocks, body) as the preview draws it: the band and the
  // headings never overlap, so the drawing is unchanged. Image blocks are numbered across the deck in 0.15 ("OPF image N"),
  // so names compare by kind.
  const order = [...xml.matchAll(/<p:(sp|pic|graphicFrame)>[\s\S]*?<p:cNvPr\b[^>]*\bname="([^"]*)"/g)].map(m => `${m[1]}:${legacy(m[2], index)}`)
    .map(entry => entry.replace(/^pic:OPF image \d+$/, slide.id === 'collage-grid' ? 'pic:OPF image' : `pic:OPF slide image slides.${index}`).replace(/^sp:OPF image \d+ overlay$/, `sp:OPF slide image overlay slides.${index}`));
  const placed = slide.blocks?.some(entry => entry.placement);
  const headingsFirst = list => placed ? [...list.filter(entry => /OPF heading/.test(entry)), ...list.filter(entry => !/OPF heading/.test(entry))] : list;
  assert.deepEqual(order, headingsFirst(expected.order.map(entry => entry.replace(/^pic:OPF image \d+$/, 'pic:OPF image'))), `${slide.id}: paint order`);
  assert.equal(pictures.length, expected.pictures.length, slide.id);
  pictures.forEach((picture, at) => {
    const was = expected.pictures[at];
    sameNumbers(picture['p:spPr']['a:xfrm'], was.xfrm, `${slide.id} frame`);
    sameNumbers(picture['p:spPr']['a:prstGeom'], was.geom, `${slide.id} mask`);
    sameNumbers(picture['p:spPr']['a:ln'] ?? null, was.ln, `${slide.id} line`);
    sameNumbers(picture['p:blipFill']['a:srcRect'] ?? null, was.srcRect, `${slide.id} crop`);
    sameNumbers(Object.fromEntries(Object.entries(picture['p:blipFill']['a:blip']).filter(([key]) => !key.startsWith('r:') && key !== 'a:extLst')), was.blip, `${slide.id} effects`);
  });
  const overlay = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map(([s]) => flat.parse(s)['p:sp']).find(s => /overlay/.test(s['p:nvSpPr']['p:cNvPr'].name));
  if (expected.overlay) {
    sameNumbers(overlay['p:spPr']['a:xfrm'], expected.overlay.xfrm, `${slide.id} overlay frame`);
    sameNumbers(overlay['p:spPr']['a:prstGeom'], expected.overlay.geom, `${slide.id} overlay shape`);
    sameNumbers(overlay['p:spPr']['a:solidFill'], expected.overlay.fill, `${slide.id} overlay fill`);
  } else assert.equal(overlay, undefined, slide.id);
  same++;
});
assert.equal(same + changed, 15);
console.log(`image treatments vs 0.14: ${same} treatments keep 0.14's frames and paint order; ${changed} changed by design (${Object.keys(CHANGED).join(', ')})`);
