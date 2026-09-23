import assert from 'node:assert/strict';
import {unzipSync, zipSync} from 'fflate';
import {validatePresentation} from '@openpresentation/opf';
import {toPptx, fromPptx} from '../dist/index.js';

// FF-29: video payloads export the preview's placeholder. The frame links to a
// web source and every placeholder shape carries OPF_MEDIA_V1, so import
// rebuilds the video block (and the asset entries it references) instead of
// leaving shape-name fallback blocks. Edited groups keep their caption text
// and report invalid-media-provenance.
const enc = new TextEncoder(), dec = new TextDecoder();
const deck = {
  $schema: 'https://openpresentation.org/schema/opf/v1', name: 'Media deck', design: {fontScheme: 'roboto'},
  assets: {'demo-video': {src: 'https://cdn.example.com/private/demo.mp4', title: 'Demo'}},
  slides: [
    {title: 'Titled', blocks: [{type: 'video', video: {src: 'https://example.com/walkthrough.mp4', title: 'Walkthrough', description: 'Two minute tour'}}]},
    {title: 'Asset', video: {src: 'asset:demo-video', title: 'Asset clip', description: 'From the registry'}},
    {title: 'Plain', video: 'https://example.com/plain.mp4'},
  ]
};
assert.equal(validatePresentation(deck).valid, true);
const exported = await toPptx(structuredClone(deck), {seed: 1});
const read = async bytes => {
  const issues = [];
  const doc = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)});
  assert.equal(validatePresentation(doc).valid, true);
  return {doc, media: issues.filter(issue => /media/.test(issue.code)).map(issue => [issue.code, issue.path])};
};
const modify = (bytes, mutate) => { const entries = unzipSync(bytes); mutate(entries); return zipSync(entries); };
const text = (entries, path, mutate) => { entries[path] = enc.encode(mutate(dec.decode(entries[path]))); };
const videos = doc => doc.slides.map(slide => (slide.blocks ?? []).filter(block => block.video !== undefined).map(block => block.video));
const junk = doc => doc.slides.flatMap(slide => slide.blocks ?? []).filter(block => /PowerPoint shape|OPF media/.test(block.text ?? ''));
const blockTexts = doc => doc.slides.flatMap(slide => slide.blocks ?? []).map(block => block.text).filter(Boolean);
const tagParts = entries => Object.keys(entries).filter(path => /^ppt\/tags\/opfMedia\d+\.xml$/.test(path));
const tagValue = xml => JSON.parse(Buffer.from(xml.match(/\bval="([^"]+)"/)[1], 'hex').toString('utf8'));
let checks = 0;

// Package shape: frame hyperlink to the web source; one OPF_MEDIA_V1 tag per placeholder shape.
{
  const entries = unzipSync(exported);
  const slide = n => dec.decode(entries[`ppt/slides/slide${n}.xml`]), rels = n => dec.decode(entries[`ppt/slides/_rels/slide${n}.xml.rels`]);
  assert.match(slide(1), /<p:cNvPr id="\d+" name="OPF media slides\.0\.blocks\.0\.video frame"><a:hlinkClick r:id="(rId\d+)"/);
  assert.match(rels(1), /Type="[^"]*\/hyperlink" Target="https:\/\/example\.com\/walkthrough\.mp4" TargetMode="External"/);
  assert.match(rels(2), /Target="https:\/\/cdn\.example\.com\/private\/demo\.mp4"/, 'Asset references link to the dereferenced source.');
  const records = tagParts(entries).map(path => tagValue(dec.decode(entries[path])));
  assert.deepEqual(records.filter(record => record.role === 'frame').map(record => record.video), deck.slides.map(slide => slide.video ?? slide.blocks[0].video));
  assert.deepEqual(records.find(record => record.path === 'slides.1.video' && record.role === 'frame').assets, deck.assets);
  for (const path of ['slides.0.blocks.0.video', 'slides.1.video', 'slides.2.video']) assert.deepEqual(records.filter(record => record.path === path).map(record => record.role).sort(), ['badge', 'caption', 'frame', 'play']);
  checks++;
}

// Unchanged package: every video returns exactly, asset entries included, with no fallback blocks or diagnostics.
{
  const {doc, media} = await read(exported);
  assert.deepEqual(media, []);
  assert.deepEqual(videos(doc), [[deck.slides[0].blocks[0].video], [deck.slides[1].video], [deck.slides[2].video]]);
  assert.deepEqual(doc.assets, deck.assets);
  assert.deepEqual(junk(doc), []);
  assert.deepEqual(blockTexts(doc), [], 'Captions are part of the restored video, not extra text.');
  // Re-export writes the same OPF values.
  const again = unzipSync(await toPptx(doc, {seed: 1}));
  assert.deepEqual(tagParts(again).map(path => tagValue(dec.decode(again[path]))).filter(record => record.role === 'frame').map(record => record.video), videos(doc).flat());
  checks++;
}

// references-only and provenance:false never copy asset registry entries into media tags.
for (const provenance of ['references-only', false]) {
  const entries = unzipSync(await toPptx(structuredClone(deck), {seed: 1, provenance}));
  const records = tagParts(entries).map(path => tagValue(dec.decode(entries[path])));
  assert.ok(records.every(record => record.assets === undefined));
  assert.ok(!records.some(record => JSON.stringify(record).includes('cdn.example.com')), 'No registry URL in media tags.');
  const {doc, media} = await read(zipSync(entries));
  assert.deepEqual(media, []);
  assert.deepEqual(videos(doc)[1], [deck.slides[1].video], 'The authored asset reference still returns.');
  assert.equal(doc.assets, undefined);
  checks++;
}

// Edits keep native content and name the video that was not restored; decoration never becomes a text block.
{
  const edited = await read(modify(exported, entries => text(entries, 'ppt/slides/slide1.xml', xml => xml.replace('<a:t>Walkthrough</a:t>', '<a:t>Renamed tour</a:t>'))));
  assert.deepEqual(edited.media, [['invalid-media-provenance', 'slides.0.blocks.0.video']]);
  assert.deepEqual(videos(edited.doc)[0], []);
  assert.deepEqual(edited.doc.slides[0].blocks.map(block => block.text), ['Renamed tour']);
  assert.deepEqual(videos(edited.doc).slice(1), [[deck.slides[1].video], [deck.slides[2].video]], 'Other slides are unaffected.');
  assert.deepEqual(junk(edited.doc), []);

  const removed = await read(modify(exported, entries => text(entries, 'ppt/slides/slide3.xml', xml => xml.replace(/<p:sp><p:nvSpPr><p:cNvPr id="\d+" name="OPF media slides\.2\.video badge">[\s\S]*?<\/p:sp>/, ''))));
  assert.deepEqual(removed.media, [['invalid-media-provenance', 'slides.2.video']]);
  assert.deepEqual(removed.doc.slides[2].blocks.map(block => block.text), ['https://example.com/plain.mp4'], 'The caption keeps the source as text.');
  assert.deepEqual(junk(removed.doc), []);

  const tampered = await read(modify(exported, entries => {
    const path = tagParts(entries).find(part => { const record = tagValue(dec.decode(entries[part])); return record.role === 'frame' && record.path === 'slides.0.blocks.0.video'; });
    text(entries, path, xml => xml.replace(/val="[0-9A-F]+"/, `val="${Buffer.from(JSON.stringify({v: 1, role: 'frame', path: 'slides.0.blocks.0.video', video: {src: 42}})).toString('hex').toUpperCase()}"`));
  }));
  assert.deepEqual(tampered.media, [['invalid-media-provenance', 'slides.0.blocks.0.video']]);
  assert.deepEqual(tampered.doc.slides[0].blocks.map(block => block.text), ['Walkthrough']);
  assert.deepEqual(junk(tampered.doc), []);
  checks++;
}

// An unreadable or ambiguous media tag is not trusted, so its shape is never
// removed by name: it goes through ordinary import. For the badge (a shape
// without text) that is one fallback block, {type: 'text', text: 'PowerPoint
// shape: OPF media <path> badge'}. The rest of the group is then incomplete:
// its other decoration is dropped, the caption stays as text, and
// invalid-media-provenance names the video. An ambiguous tag also reports the
// unattributable shape at the slide path.
{
  const badgeTag = entries => tagParts(entries).find(path => { const record = tagValue(dec.decode(entries[path])); return record.role === 'badge' && record.path === 'slides.0.blocks.0.video'; });
  const fallback = [{type: 'text', text: 'PowerPoint shape: OPF media slides.0.blocks.0.video badge'}, {type: 'text', text: 'Walkthrough'}];
  const cases = {
    // A second OPF_ identity on the same shape.
    ambiguous: [entries => text(entries, badgeTag(entries), xml => xml.replace('</p:tagLst>', '<p:tag name="OPF_CARD_V1" val="7B7D"/></p:tagLst>')),
      [['invalid-media-provenance', 'slides.0.blocks.0.video'], ['invalid-media-provenance', 'slides.0']]],
    // A tag part that is not XML, and a tag relationship whose part is missing.
    unreadable: [entries => { entries[badgeTag(entries)] = enc.encode('not xml <<<'); }, [['invalid-media-provenance', 'slides.0.blocks.0.video']]],
    missing: [entries => { delete entries[badgeTag(entries)]; }, [['invalid-media-provenance', 'slides.0.blocks.0.video']]],
  };
  for (const [label, [mutate, diagnostics]] of Object.entries(cases)) {
    const {doc, media} = await read(modify(exported, mutate));
    assert.deepEqual(media, diagnostics, `${label}: diagnostics`);
    assert.deepEqual(doc.slides[0].blocks, fallback, `${label}: one badge fallback block and the caption text`);
    assert.deepEqual(videos(doc).slice(1), [[deck.slides[1].video], [deck.slides[2].video]], `${label}: other slides are unaffected`);
  }
  checks++;
}

console.log(`Media placeholder passed: ${checks} groups; web-source hyperlinks, OPF_MEDIA_V1 tags, exact video and asset round trip, no registry copies outside 'full', edited groups keep captions with invalid-media-provenance and no fallback blocks, and an untrusted (unreadable or ambiguous) tag leaves one documented fallback block with the diagnostic.`);
