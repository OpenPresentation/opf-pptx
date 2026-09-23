import assert from 'node:assert/strict';
import {unzipSync, zipSync} from 'fflate';
import {toPptx, fromPptx} from '../dist/index.js';
import {validatePresentation, catalogs, schemas} from '@openpresentation/opf';
import {resolvePresentation} from '@openpresentation/opf-render';

// Generated socials furniture: every line is native text linked to its profile
// URL, and re-import rebuilds organization.socials from the current lines.
const enc = new TextEncoder(), dec = new TextDecoder();
const organization = {id: 'acme', name: 'Acme', socials: {linkedin: 'acme', x: '@acme', bluesky: 'https://bsky.app/profile/acme.bsky.social', custom: 'Visit us', legacy: 'http://acme.example/profile'}};
const source = {organization, design: {footer: {left: {organization: true}, right: {socials: true}}}, slides: [{title: 'One', text: 'Body'}, {title: 'Two', text: 'Body'}]};
const read = async bytes => { const issues = []; const deck = await fromPptx(bytes, {onDiagnostic: issue => issues.push(issue)}); assert.equal(validatePresentation(deck).valid, true); return {deck, issues}; };
const modify = (bytes, mutate) => { const entries = unzipSync(bytes); mutate(entries); return zipSync(entries); };
const slideXml = (entries, index = 1) => dec.decode(entries[`ppt/slides/slide${index}.xml`]);
const shapeText = (entries, text, replacement) => { for (const index of [1, 2]) entries[`ppt/slides/slide${index}.xml`] = enc.encode(slideXml(entries, index).replace(`<a:t>${text}</a:t>`, `<a:t>${replacement}</a:t>`)); };

const before = structuredClone(source), bytes = await toPptx(source);
assert.deepEqual(source, before, 'Export leaves the source unchanged.');
assert.deepEqual(await toPptx(source), bytes, 'Socials export is deterministic.');
const entries = unzipSync(bytes), xml = slideXml(entries), rels = dec.decode(entries['ppt/slides/_rels/slide1.xml.rels']);
const shown = ['linkedin.com/company/acme', 'x.com/acme', 'bsky.app/profile/acme.bsky.social', 'Visit us', 'http://acme.example/profile'];
for (const text of shown) assert.ok(xml.includes(`<a:t>${text}</a:t>`), text);
for (const url of ['https://linkedin.com/company/acme', 'https://x.com/acme', 'https://bsky.app/profile/acme.bsky.social', 'http://acme.example/profile'])
  assert.ok(rels.includes(`Target="${url}" TargetMode="External"`), url);
assert.equal((xml.match(/<a:hlinkClick /g) ?? []).length, 4, 'Only formatted or URL values link; raw text does not.');
assert.equal((xml.match(/hlinkClr[^>]*val="tx"/g) ?? []).length, 4, 'Links keep the furniture text colour.');
assert.equal((xml.match(/<a:rPr [^>]*u="none"[^>]*>(?:(?!<\/a:rPr>).)*<a:hlinkClick /g) ?? []).length, 4, 'Links are not underlined, matching the preview.');

// Without document provenance (FF-32), furniture alone rebuilds the socials as canonical profile URLs.
const canonicalSocials = {linkedin: 'https://linkedin.com/company/acme', x: 'https://x.com/acme',
  bluesky: 'https://bsky.app/profile/acme.bsky.social', custom: 'Visit us', legacy: 'http://acme.example/profile'};
const plain = await read(await toPptx(source, {provenance: false}));
assert.deepEqual(plain.deck.organization, {id: 'acme', name: 'Acme', socials: canonicalSocials}, 'Handles re-import as their canonical profile URLs.');
assert.deepEqual(plain.deck.design.footer, source.design.footer);
// With provenance, unedited lines keep the authored form (a handle stays a handle).
const {deck, issues} = await read(bytes);
assert.deepEqual(deck.design.footer, source.design.footer);
assert.deepEqual(deck.organization, organization, 'Unedited profile lines keep the authored socials values.');
assert.ok(issues.some(issue => issue.code === 'furniture-import-reflow'));
assert.ok(!issues.some(issue => issue.code === 'invalid-furniture-provenance'));
for (const slide of deck.slides) assert.ok(!JSON.stringify(slide.blocks ?? []).includes('x.com/acme'), 'Social lines are not duplicated as body text.');
// The re-imported deck exports the same visible profiles.
const again = unzipSync(await toPptx(deck));
for (const text of shown) assert.ok(slideXml(again).includes(`<a:t>${text}</a:t>`), text);

// Current native words win: an edited profile line becomes the new value.
const edited = await read(modify(bytes, entries => shapeText(entries, 'x.com/acme', 'x.com/acme_news')));
assert.deepEqual(edited.deck.organization.socials, {...organization.socials, x: 'https://x.com/acme_news'}, 'Only the edited line changes.');
assert.equal((await read(modify(await toPptx(source, {provenance: false}), entries => shapeText(entries, 'x.com/acme', 'x.com/acme_news')))).deck.organization.socials.x, 'https://x.com/acme_news');
// Disagreeing slides, cleared lines and unowned lines (no repeated organization
// name) fall back to ordinary current text with a specific diagnostic. Without
// provenance no socials are rebuilt; with it the stored authored socials return.
const disagreeing = entries => { entries['ppt/slides/slide2.xml'] = enc.encode(slideXml(entries, 2).replace('<a:t>x.com/acme</a:t>', '<a:t>x.com/other</a:t>')); };
const clearing = entries => shapeText(entries, 'Visit us', '');
const unnamed = {...source, design: {footer: {right: {socials: true}}}};
for (const [deckSource, mutate, pattern] of [[source, disagreeing, /social profile metadata disagrees/], [source, clearing, /social profile lines/], [unnamed, () => {}, /repeated organization name/]]) {
  const bare = await read(modify(await toPptx(deckSource, {provenance: false}), mutate));
  assert.equal(bare.deck.organization?.socials, undefined, String(pattern));
  assert.ok(bare.issues.some(issue => issue.code === 'invalid-furniture-provenance' && pattern.test(issue.message)), String(pattern));
  const stored = await read(modify(await toPptx(deckSource), mutate));
  assert.deepEqual(stored.deck.organization, organization, `stored ${pattern}`);
  assert.ok(stored.issues.some(issue => issue.code === 'invalid-furniture-provenance' && pattern.test(issue.message)), String(pattern));
  if (deckSource === unnamed) assert.ok(JSON.stringify(bare.deck.slides).includes('x.com/acme'), 'Unowned profile lines stay as current text.');
}

// Missing socials diagnose the controlling path; strict composition rejects them.
const missing = {...source, organization: {id: 'acme', name: 'Acme'}}, diagnostics = [];
await toPptx(missing, {onDiagnostic: issue => diagnostics.push(issue)});
assert.ok(diagnostics.some(issue => issue.code === 'unresolved-content' && issue.path === 'design.footer.right.socials'));

// Every bundled platform exports its own example handle as a linked profile URL.
for (const platform of catalogs.socialPlatforms) {
  const platformDeck = {...source, organization: {id: 'acme', name: 'Acme', socials: {[platform.id]: platform.handleExample}}, slides: [{text: 'Body'}]};
  assert.equal((await read(await toPptx(platformDeck))).deck.organization.socials[platform.id], platform.handleExample, platform.id);
  const one = await toPptx(platformDeck, {provenance: false});
  const {deck: back} = await read(one);
  assert.match(back.organization.socials[platform.id], /^https:\/\//, platform.id);
  assert.ok(dec.decode(unzipSync(one)['ppt/slides/_rels/slide1.xml.rels']).includes(`Target="${back.organization.socials[platform.id].replace(/&/g, '&amp;')}"`), platform.id);
}
// Export resolves socialPlatforms records exactly like the opf-render preview:
// inline records, then the document source (catalogSources), injected catalogs, bundled.
const record = (id, url) => ({$schema: 'https://openpresentation.org/schema/opf-social-platform/v1', id, name: id, profileUrlPattern: url, handlePrefix: '@'});
const sourced = {...source, catalogs: {socialPlatforms: {source: 'https://example.test/socials.json'}}, slides: [{text: 'Body'}]};
const hosts = [
  [{}, 'x.com/acme'],
  [{catalogs: {socialPlatforms: [record('x', 'https://injected.test/{handle}')]}}, 'injected.test/acme'],
  [{catalogs: {socialPlatforms: [record('x', 'https://injected.test/{handle}')]}, catalogSources: {'https://example.test/socials.json': {records: [record('x', 'https://sourced.test/{handle}')]}}}, 'sourced.test/acme'],
];
// CI links the pinned opf-render. The preview half of this check runs only when that
// renderer passes socials records (opf-render#34 or later); the export half always runs.
const previewText = (deck, host) => resolvePresentation(deck, host).slides[0].geometry.furniture.parts.find(part => part.field === 'socials').text.split('\n')[1];
const previewFormats = previewText(sourced, {}) === 'x.com/acme';
if (!previewFormats) console.log('Installed opf-render predates FF-34 socials records: preview parity skipped, export order still checked.');
for (const [host, expected] of hosts) for (const deck of [sourced, {...sourced, catalogs: {socialPlatforms: {source: sourced.catalogs.socialPlatforms.source, records: [record('x', 'https://inline.test/{handle}')]}}}]) {
  const want = deck.catalogs.socialPlatforms.records ? 'inline.test/acme' : expected;
  if (previewFormats) assert.equal(previewText(deck, host), want, 'preview');
  assert.ok(slideXml(unzipSync(await toPptx(deck, host))).includes(`<a:t>${want}</a:t>`), `export ${want}`);
}

// Every platform key the Socials schema accepts re-imports; the provenance check uses the schema pattern itself.
const keyPattern = new RegExp(schemas.presentation.$defs.Socials.propertyNames.pattern, 'u');
assert.equal(keyPattern.test('my_site'), false, 'Underscored keys are not valid Socials keys.');
const keyed = {...source, organization: {id: 'acme', name: 'Acme', socials: {'my-site2': 'Visit us', x9: '@acme'}}, slides: [{text: 'Body'}]};
assert.deepEqual((await read(await toPptx(keyed, {provenance: false}))).deck.organization.socials, {'my-site2': 'Visit us', x9: '@acme'});

console.log(`Socials furniture passed: linked profile lines, canonical re-import, edits, disagreement, orphan and ${catalogs.socialPlatforms.length} bundled platforms.`);
