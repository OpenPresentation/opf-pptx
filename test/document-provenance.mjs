import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {unzipSync, zipSync} from 'fflate';
import {XMLValidator} from 'fast-xml-parser';
import {toPptx, fromPptx} from '../dist/index.js';
import {validatePresentation, catalogs} from '@openpresentation/opf';

// FF-32: catalog references, slide layout ids and authoring metadata survive
// export -> import while the native evidence they produced is unchanged, and a
// specific diagnostic replaces a silent loss after an edit.
const enc = new TextEncoder(), dec = new TextDecoder();
const logo = `data:image/png;base64,${(await readFile(new URL('fixtures/images/wide.png', import.meta.url))).toString('base64')}`;
const layoutRecord = {...structuredClone(catalogs.layouts.find(record => record.id === 'title-subtitle')), id: 'hero-title', name: 'Hero Title'};
const unused = {...structuredClone(layoutRecord), id: 'never-used'};
const source = {
  $schema: 'https://openpresentation.org/schema/opf/v1', name: 'Provenance deck',
  organization: [{id: 'acme', name: 'Acme', role: 'primary', logo: 'asset:acme-logo'}, {id: 'beta', name: 'Beta', role: 'partner'}],
  speaker: {id: 'alice', name: 'Alice Chen', title: 'VP', organizationId: 'acme'},
  audience: ['executives', {id: 'board', name: 'Board of directors'}], purpose: 'decide', language: 'japanese', tone: 'formal',
  narrative: 'problem-solution', takeaway: ['Ship it'], duration: 20, tags: ['gallery:test'],
  variables: {risk: '#C0392B'},
  assets: {'acme-logo': {src: logo, alt: 'Acme logo'}, unrelated: 'https://example.com/unused.png'},
  design: {theme: 'bold', colorScheme: 'boost', fontScheme: 'arial', dimensions: {preset: '16:9'}, background: {type: 'solid', color: '#FAFAFA'},
    contentAlignment: 'center'},
  catalogs: {layouts: {records: [layoutRecord, unused]}},
  slides: [
    {id: 'intro', beat: 'problem', layout: 'hero-title', title: 'Hello', subtitle: 'World', design: {titleAlignment: 'center', contentBox: false}},
    {id: 'body', layout: 'title-subtitle', title: 'Second', subtitle: 'Slide', design: {background: {type: 'solid', color: '#102030'}}},
    {layout: 'title-subtitle', title: 'Third', subtitle: 'Inherits'}
  ]
};
assert.equal(validatePresentation(source).valid, true, JSON.stringify(validatePresentation(source).errors));

const exported = await toPptx(structuredClone(source));
assert.deepEqual(await toPptx(structuredClone(source)), exported, 'Export stays deterministic.');
const read = async bytes => {
  const issues = [];
  const deck = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)});
  assert.equal(validatePresentation(deck).valid, true);
  return {deck, issues, provenance: issues.filter(issue => /provenance|reference|slide-id/.test(issue.code))};
};
const modify = (bytes, mutate) => { const entries = unzipSync(bytes); mutate(entries); return zipSync(entries); };
const text = (entries, path, mutate) => { entries[path] = enc.encode(mutate(dec.decode(entries[path]))); };
const tagValue = xml => JSON.parse(Buffer.from(xml.match(/\bval="([^"]+)"/)[1], 'hex').toString('utf8'));

// Package shape: standard customer-data tags, schema-ordered, typed, related.
{
  const entries = unzipSync(exported);
  const presentation = dec.decode(entries['ppt/presentation.xml']);
  assert.equal(XMLValidator.validate(presentation), true);
  assert.match(presentation, /<p:notesSz [^>]*\/><p:custDataLst><p:tags r:id="rIdOpfDocument"\/><\/p:custDataLst><p:defaultTextStyle>/);
  assert.match(dec.decode(entries['ppt/_rels/presentation.xml.rels']), /Id="rIdOpfDocument" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/tags" Target="tags\/opfDocument\.xml"/);
  const types = dec.decode(entries['[Content_Types].xml']);
  for (const part of ['opfDocument', 'opfSlide1', 'opfSlide2', 'opfSlide3']) {
    assert.ok(types.includes(`<Override PartName="/ppt/tags/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>`), part);
    assert.equal(XMLValidator.validate(dec.decode(entries[`ppt/tags/${part}.xml`])), true);
  }
  const documentXml = dec.decode(entries['ppt/tags/opfDocument.xml']);
  assert.match(documentXml, /<p:tag name="OPF_DOCUMENT_V1" val="[0-9A-F]+"\/>/);
  const document = tagValue(documentXml);
  assert.deepEqual(document.design, source.design);
  assert.deepEqual(Object.keys(document.metadata).sort(), ['audience', 'duration', 'language', 'narrative', 'organization', 'purpose', 'speaker', 'tags', 'takeaway', 'tone', 'variables']);
  assert.deepEqual(Object.keys(document.assets), ['acme-logo'], 'Only assets referenced by metadata are stored.');
  assert.deepEqual(document.catalogs.layouts.records.map(record => record.id), ['hero-title'], 'Only referenced inline records are stored.');
  assert.equal(document.native.size.cx, '12192000');
  assert.deepEqual(Object.keys(document.native.colors), ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']);
  assert.equal(document.native.fonts.major.latin, 'Arial Black');
  const slide = dec.decode(entries['ppt/slides/slide1.xml']);
  assert.match(slide, /<\/p:spTree><p:custDataLst><p:tags r:id="rIdOpfSlide"\/><\/p:custDataLst><\/p:cSld>/);
  const record = tagValue(dec.decode(entries['ppt/tags/opfSlide1.xml']));
  assert.deepEqual({id: record.id, beat: record.beat, layout: record.layout, design: record.design}, {id: 'intro', beat: 'problem', layout: 'hero-title', design: {titleAlignment: 'center', contentBox: false}});
  for (const key of ['structure', 'background', 'style']) assert.match(record.native[key], /^[0-9a-f]{14}$/);
  // Tags carry no typeface attribute, so package font inventories are unchanged.
  for (const part of Object.keys(entries).filter(path => path.startsWith('ppt/tags/'))) assert.doesNotMatch(dec.decode(entries[part]), /typeface=/);
}

// Unchanged package: every reference and metadata value returns.
{
  const {deck, provenance} = await read(exported);
  assert.deepEqual(provenance, []);
  assert.deepEqual(deck.design, source.design);
  for (const key of ['organization', 'speaker', 'audience', 'purpose', 'language', 'tone', 'narrative', 'takeaway', 'duration', 'tags', 'variables']) assert.deepEqual(deck[key], source[key], key);
  assert.deepEqual(deck.assets, {'acme-logo': source.assets['acme-logo']});
  assert.deepEqual(deck.catalogs, {layouts: {records: [layoutRecord]}});
  assert.deepEqual(deck.slides.map(slide => [slide.id, slide.beat, slide.layout]), [['intro', 'problem', 'hero-title'], ['body', undefined, 'title-subtitle'], [undefined, undefined, 'title-subtitle']]);
  assert.deepEqual(deck.slides.map(slide => slide.design), [{titleAlignment: 'center', contentBox: false}, {background: {type: 'solid', color: '#102030'}}, undefined],
    'Inherited backgrounds are not repeated as slide overrides; local ones return as authored.');
  // The reimported document exports again and keeps its references.
  const again = await read(await toPptx(deck));
  assert.deepEqual(again.provenance, []);
  assert.deepEqual(again.deck.design, source.design);
  assert.deepEqual(again.deck.slides.map(slide => slide.layout), ['hero-title', 'title-subtitle', 'title-subtitle']);
}

// Benign rewrites an editor performs on save do not count as edits.
{
  const rewritten = modify(exported, entries => {
    text(entries, 'ppt/theme/theme1.xml', xml => xml.replace(/lastClr="[0-9A-F]{6}"/g, 'lastClr="123456"'));
    for (const index of [1, 2, 3]) text(entries, `ppt/slides/slide${index}.xml`, xml => xml.replace(/<a:srgbClr val="([0-9A-F]{6})"><\/a:srgbClr>/g, '<a:srgbClr val="$1"/>')
      .replace(/(<p:sp>[\s\S]*?<\/p:sp>)(<p:sp>[\s\S]*?<\/p:sp>)/, '$2$1'));
  });
  const {deck, provenance} = await read(rewritten);
  assert.deepEqual(provenance, [], 'System color refresh, empty-element form and stacking order are not structural edits.');
  assert.deepEqual(deck.design, source.design);
}

// Theme colors changed: the color scheme and the theme bundle stay observed.
{
  const {deck, provenance} = await read(modify(exported, entries => text(entries, 'ppt/theme/theme1.xml', xml => xml.replace(/<a:accent1>[\s\S]*?<\/a:accent1>/, '<a:accent1><a:srgbClr val="00FF00"/></a:accent1>'))));
  // Stored references win only while the theme matches; FF-24's recovered scheme is the observed fallback.
  assert.deepEqual(deck.design.colorScheme, {id: 'boost', accent1: '#00FF00'});
  assert.notEqual(deck.design.theme, 'bold');
  assert.equal(deck.design.fontScheme, 'arial');
  assert.deepEqual(provenance.map(issue => [issue.code, issue.path]), [['design-reference-changed', 'design.theme'], ['design-reference-changed', 'design.colorScheme']]);
  assert.match(provenance[1].message, /accent1/);
  assert.match(provenance[1].message, /'boost'/);
}

// Theme fonts changed: the font scheme stays observed.
{
  const {deck, provenance} = await read(modify(exported, entries => text(entries, 'ppt/theme/theme1.xml', xml => xml.replace(/(<a:majorFont><a:latin typeface=")[^"]*"/, '$1Georgia"'))));
  assert.notEqual(deck.design.fontScheme, 'arial');
  assert.equal(deck.design.colorScheme, 'boost');
  assert.deepEqual(provenance.map(issue => [issue.code, issue.path]), [['design-reference-changed', 'design.theme'], ['design-reference-changed', 'design.fontScheme']]);
  assert.match(provenance[1].message, /major font/);
}

// Slide size changed: observed dimensions stay.
{
  const {deck, provenance} = await read(modify(exported, entries => text(entries, 'ppt/presentation.xml', xml => xml.replace('<p:sldSz cx="12192000" cy="6858000"/>', '<p:sldSz cx="9144000" cy="6858000"/>'))));
  assert.deepEqual(deck.design.dimensions, {widthInches: 10, heightInches: 7.5});
  assert.deepEqual(provenance.map(issue => [issue.code, issue.path]), [['design-reference-changed', 'design.dimensions']]);
}

// A moved object changes that slide's layout structure only.
{
  const {deck, provenance} = await read(modify(exported, entries => text(entries, 'ppt/slides/slide1.xml', xml => xml.replace(/(<p:sp>[\s\S]*?<a:off x=")(\d+)"/, (_, before, x) => `${before}${Number(x) + 12700}"`))));
  assert.deepEqual(deck.slides.map(slide => slide.layout), [undefined, 'title-subtitle', 'title-subtitle']);
  assert.equal(deck.slides[0].id, 'intro', 'Identity and beat do not depend on arrangement.');
  assert.equal(deck.slides[0].beat, 'problem');
  assert.equal(deck.design.contentAlignment, undefined, 'Deck composition defaults need every slide unchanged.');
  assert.deepEqual(provenance.map(issue => [issue.code, issue.path]), [
    ['layout-reference-changed', 'slides.0.layout'], ['design-reference-changed', 'slides.0.design.titleAlignment'],
    ['design-reference-changed', 'slides.0.design.contentBox'], ['design-reference-changed', 'design.contentAlignment']]);
  assert.deepEqual(deck.catalogs, undefined, 'An unrestored layout id does not keep its inline record.');
}

// Background edits: an edited inheriting slide keeps a local override; editing
// every inheriting slide leaves the deck background unrestored.
{
  const recolor = (entries, index) => text(entries, `ppt/slides/slide${index}.xml`, xml => xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'));
  const one = await read(modify(exported, entries => recolor(entries, 1)));
  assert.deepEqual(one.deck.design.background, source.design.background);
  assert.deepEqual(one.deck.slides[0].design.background, {type: 'solid', color: '#FF0000'});
  assert.equal(one.deck.slides[2].design, undefined);
  assert.deepEqual(one.provenance.map(issue => [issue.code, issue.path]), [['design-reference-changed', 'slides.0.design.background']]);
  const all = await read(modify(exported, entries => { recolor(entries, 1); recolor(entries, 3); recolor(entries, 2); }));
  assert.equal(all.deck.design.background, undefined);
  assert.deepEqual(all.deck.slides.map(slide => slide.design?.background?.color), ['#FF0000', '#FF0000', '#FF0000']);
  assert.deepEqual(all.provenance.map(issue => [issue.code, issue.path]), [['design-reference-changed', 'slides.1.design.background'], ['design-reference-changed', 'design.background']]);
}

// Duplicated slide (PowerPoint copies slide tags): layout returns, the id stays unique.
{
  const duplicated = modify(exported, entries => {
    entries['ppt/slides/slide4.xml'] = entries['ppt/slides/slide1.xml'];
    entries['ppt/slides/_rels/slide4.xml.rels'] = enc.encode(dec.decode(entries['ppt/slides/_rels/slide1.xml.rels']).replace(/opfSlide1\.xml/, 'opfSlide4.xml').replace(/notesSlide1\.xml/, 'notesSlide1.xml'));
    entries['ppt/tags/opfSlide4.xml'] = entries['ppt/tags/opfSlide1.xml'];
    text(entries, '[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/tags/opfSlide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/></Types>'));
    text(entries, 'ppt/_rels/presentation.xml.rels', xml => xml.replace('</Relationships>', '<Relationship Id="rIdCopy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/></Relationships>'));
    text(entries, 'ppt/presentation.xml', xml => xml.replace('</p:sldIdLst>', '<p:sldId id="400" r:id="rIdCopy"/></p:sldIdLst>'));
  });
  const {deck, provenance} = await read(duplicated);
  assert.deepEqual(deck.slides.map(slide => [slide.id, slide.layout]), [['intro', 'hero-title'], ['body', 'title-subtitle'], [undefined, 'title-subtitle'], [undefined, 'hero-title']]);
  assert.deepEqual(provenance.map(issue => [issue.code, issue.path]), [['duplicate-slide-id', 'slides.3.id'], ['design-reference-changed', 'design.contentAlignment']]);
}

// Stripped tags are an ordinary foreign deck; damaged or invalid ones fall back with a diagnostic.
{
  const stripped = await read(modify(exported, entries => text(entries, 'ppt/presentation.xml', xml => xml.replace(/<p:custDataLst>[\s\S]*?<\/p:custDataLst>/, ''))));
  assert.deepEqual(stripped.provenance, []);
  assert.equal(stripped.deck.narrative, undefined);
  assert.equal(stripped.deck.slides[0].layout, undefined);

  const damaged = await read(modify(exported, entries => text(entries, 'ppt/tags/opfDocument.xml', xml => xml.replace(/val="[0-9A-F]{8}/, 'val="ZZZZZZZZ'))));
  assert.deepEqual(damaged.provenance.map(issue => [issue.code, issue.path]), [['invalid-document-provenance', '']]);
  assert.equal(damaged.deck.narrative, undefined);

  const invalid = await read(modify(exported, entries => text(entries, 'ppt/tags/opfDocument.xml', xml => {
    const value = tagValue(xml);
    value.metadata.duration = -5;
    return xml.replace(/val="[0-9A-F]+"/, `val="${Buffer.from(JSON.stringify(value)).toString('hex').toUpperCase()}"`);
  })));
  assert.deepEqual(invalid.provenance.map(issue => issue.code), ['invalid-document-provenance']);
  assert.equal(invalid.deck.duration, undefined);
  assert.equal(invalid.deck.slides[0].layout, undefined, 'An invalid record restores nothing.');

  const unknown = await read(modify(exported, entries => text(entries, 'ppt/tags/opfDocument.xml', xml => {
    const value = tagValue(xml);
    value.metadata.script = 'x';
    const json = JSON.stringify(value).replace('"script":"x"', '"__proto__":{"polluted":true}');
    return xml.replace(/val="[0-9A-F]+"/, `val="${Buffer.from(json).toString('hex').toUpperCase()}"`);
  })));
  assert.deepEqual(unknown.provenance.map(issue => issue.code), ['invalid-document-provenance']);
  assert.equal({}.polluted, undefined);
}

// Furniture already owns a slide tag list; both records share it.
{
  const deck = {name: 'Footer deck', organization: {id: 'acme', name: 'Acme'}, design: {fontScheme: 'arial', footer: {left: {organization: true}, right: {slideNumber: true}}},
    slides: [{layout: 'title-subtitle', title: 'One', subtitle: 'Two'}]};
  const bytes = await toPptx(deck);
  const entries = unzipSync(bytes);
  const tags = dec.decode(entries['ppt/tags/opfFurnitureSlide0.xml']);
  assert.match(tags, /<p:tag name="OPF_FURNITURE_V1" val="[0-9A-F]+"\/><p:tag name="OPF_SLIDE_V1" val="[0-9A-F]+"\/><\/p:tagLst>/);
  assert.equal(entries['ppt/tags/opfSlide1.xml'], undefined);
  assert.match(dec.decode(entries['ppt/slides/slide1.xml']), /<\/p:spTree><p:custDataLst><p:tags r:id="rIdOpfFurnitureSlide"\/><\/p:custDataLst><\/p:cSld>/, 'One slide-level tag list.');
  const {deck: imported, issues} = await read(bytes);
  assert.equal(issues.some(issue => issue.code === 'invalid-furniture-provenance'), false);
  assert.deepEqual(imported.design.footer, deck.design.footer);
  assert.deepEqual(imported.organization, deck.organization);
  assert.equal(imported.slides[0].layout, 'title-subtitle');
  // The current organization name shown in the footer wins over the stored one.
  const renamed = await read(modify(bytes, entries => text(entries, 'ppt/slides/slide1.xml', xml => xml.replace('>Acme<', '>Acme Corp<'))));
  assert.deepEqual(renamed.deck.organization, {id: 'acme', name: 'Acme Corp'});
}

// Documents that state nothing keep their package unchanged; oversized metadata is reported, not stored.
{
  const plain = unzipSync(await toPptx({name: 'Plain', slides: [{title: 'Only text'}]}));
  assert.equal(Object.keys(plain).some(path => /opfDocument|opfSlide/.test(path)), false);
  assert.doesNotMatch(dec.decode(plain['ppt/presentation.xml']), /custDataLst/);
  const issues = [];
  const bytes = await toPptx({name: 'Large', speaker: {id: 'bio', name: 'Long', bio: 'x'.repeat(300 * 1024)}, tone: 'formal', slides: [{title: 'One'}]}, {onDiagnostic: issue => issues.push(issue)});
  assert.deepEqual(issues.filter(issue => issue.code === 'document-provenance-omitted').map(issue => issue.path), ['speaker']);
  const {deck} = await read(bytes);
  assert.equal(deck.speaker, undefined);
  assert.equal(deck.tone, 'formal');
}

console.log('Document provenance passed: package shape, full reference/metadata/layout round trip and re-export, benign rewrites, theme color/font, size, arrangement and background edits with specific diagnostics, duplicated slides, stripped/damaged/invalid tags, shared furniture tag lists and oversized metadata.');
