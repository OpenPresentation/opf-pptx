import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {createRequire} from 'node:module';
import {lstat, mkdir, readFile, readdir, realpath, writeFile, copyFile, constants} from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {isDeepStrictEqual} from 'node:util';
import {sourceDeck, routes, outcomeTemplate, fixtureFontPolicy} from './fixtures.mjs';

const repository = fileURLToPath(new URL('../../', import.meta.url));
const json = value => Buffer.from(JSON.stringify(value, null, 2) + '\n');
export const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const array = value => value === undefined ? [] : Array.isArray(value) ? value : [value];
const inside = (root, file) => {const relative = path.relative(root, file); return relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative);};
const safe = value => typeof value === 'string' && !path.isAbsolute(value) && value.split(/[/\\]/).every(part => part && part !== '.' && part !== '..');

async function fresh(output, excluded = []) {
  const target = path.resolve(output), parent = await realpath(path.dirname(target));
  assert.equal(parent, path.dirname(target), 'Output parent must be canonical');
  for (const root of [await realpath(repository), ...excluded]) assert.ok(!inside(root, target), 'Output must be outside source and installed consumer');
  await mkdir(target); // Existing attempts, including failures, are immutable.
  return target;
}

async function tree(directory, relative = '') {
  const result = [];
  for (const entry of await readdir(path.join(directory, relative), {withFileTypes: true})) {
    const name = path.posix.join(relative, entry.name);
    assert.ok(!entry.isSymbolicLink(), `Sealed files cannot be symlinks: ${name}`);
    if (entry.isDirectory()) result.push(...await tree(directory, name));
    else if (entry.isFile()) {const bytes = await readFile(path.join(directory, name)); result.push({path: name, bytes: bytes.length, sha256: hash(bytes)});}
  }
  return result.sort((a, b) => a.path.localeCompare(b.path, 'en'));
}

export async function verifySeal(suiteDirectory) {
  const root = await realpath(suiteDirectory), suite = JSON.parse(await readFile(path.join(root, 'suite.json')));
  assert.equal(suite.schemaVersion, 1); assert.equal(suite.kind, 'keynote-qualification');
  for (const file of suite.files) {
    assert.ok(safe(file.path), 'Unsafe sealed file path');
    const absolute = path.join(root, file.path);
    assert.ok((await lstat(absolute)).isFile() && !(await lstat(absolute)).isSymbolicLink(), 'Sealed input must be a regular file');
    assert.ok(inside(root, await realpath(absolute)), 'Sealed input escapes suite');
    const bytes = await readFile(absolute);
    assert.equal(bytes.length, file.bytes, `${file.path}: bytes`); assert.equal(hash(bytes), file.sha256, `${file.path}: sha256`);
  }
  assert.equal(new Set(suite.files.map(file => file.path)).size, suite.files.length, 'Duplicate sealed file');
  const actual = (await tree(root)).filter(file => file.path !== 'suite.json');
  assert.deepEqual(actual, suite.files, 'Suite has missing, extra or changed files');
  return {root, suite};
}

export async function installed(consumerDirectory) {
  const consumer = await realpath(consumerDirectory), require = createRequire(path.join(consumer, 'package.json'));
  const packages = [];
  for (const name of ['opf', 'opf-render', 'opf-pptx', 'gallery'].map(value => '@openpresentation/' + value)) {
    const manifest = await realpath(require.resolve(name + '/package.json')), directory = path.dirname(manifest);
    assert.ok(inside(path.join(consumer, 'node_modules'), directory), `${name} must be inside the selected installed consumer`);
    const pkg = JSON.parse(await readFile(manifest)); assert.equal(pkg.name, name);
    const files = name.endsWith('/gallery') ? [] : [{path: 'package.json', bytes: (await readFile(manifest)).length, sha256: hash(await readFile(manifest))}];
    // Gallery ships catalog source rather than dist. Bind its complete package payload too.
    files.push(...await tree(directory, name.endsWith('/gallery') ? '' : 'dist'));
    if (name.endsWith('/opf-pptx')) files.push(...await tree(directory, 'vendor'));
    packages.push({name, version: pkg.version, directory, files, sha256: hash(json(files))});
  }
  for (const pkg of packages.filter(pkg => !pkg.name.endsWith('/gallery'))) {
    const local = createRequire(path.join(pkg.directory, 'package.json'));
    for (const name of ['@openpresentation/opf', '@openpresentation/gallery']) {
      if (name === pkg.name) continue;
      assert.equal(await realpath(path.dirname(local.resolve(name + '/package.json'))), packages.find(pkg => pkg.name === name).directory, `${pkg.name} resolves another ${name}`);
    }
  }
  const load = name => import(pathToFileURL(require.resolve(name)).href);
  const [core, pptx, render, fonts, gallery] = await Promise.all(['@openpresentation/opf', '@openpresentation/opf-pptx', '@openpresentation/opf-render', '@openpresentation/opf-render/fonts-node', '@openpresentation/gallery'].map(load));
  const pptxRequire = createRequire(require.resolve('@openpresentation/opf-pptx/package.json'));
  return {consumer, packages, core, pptx, render, fonts, gallery: gallery.gallery, fflate: pptxRequire('fflate'), xml: pptxRequire('fast-xml-parser')};
}

const deep = (node, key) => {
  if (!node || typeof node !== 'object') return [];
  return Object.entries(node).flatMap(([name, value]) => name === key ? array(value) : array(value).flatMap(child => deep(child, key)));
};
const texts = node => deep(node, 'a:t').map(value => typeof value === 'object' ? value['#text'] ?? '' : String(value));

export function inspectArchive(bytes, engine) {
  const parts = engine.fflate.unzipSync(bytes), {XMLParser, XMLValidator} = engine.xml;
  const parser = new XMLParser({ignoreAttributes: false, attributeNamePrefix: '', parseTagValue: false, trimValues: false});
  const xml = name => {assert.ok(parts[name], `Missing part: ${name}`); const value = engine.fflate.strFromU8(parts[name]); assert.equal(XMLValidator.validate(value), true, `Invalid XML: ${name}`); return parser.parse(value);};
  const relations = source => {
    const file = path.posix.join(path.posix.dirname(source), '_rels', path.posix.basename(source) + '.rels');
    if (!parts[file]) return [];
    return array(xml(file).Relationships?.Relationship).map(rel => {
      const target = decodeURIComponent(rel.Target.split('#')[0]);
      return {...rel, part: rel.TargetMode === 'External' ? null : path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(source), target))};
    });
  };
  for (const name of Object.keys(parts).filter(name => /\.(xml|rels)$/.test(name))) xml(name);
  for (const name of Object.keys(parts).filter(name => name.endsWith('.rels'))) {
    const source = name === '_rels/.rels' ? '' : path.posix.join(path.posix.dirname(path.posix.dirname(name)), path.posix.basename(name, '.rels'));
    const ids = new Set();
    for (const rel of relations(source)) {assert.ok(!ids.has(rel.Id), `Duplicate relationship: ${name}/${rel.Id}`); ids.add(rel.Id); if (rel.part) assert.ok(parts[rel.part], `Missing relationship target: ${rel.part}`);}
  }
  const presentation = xml('ppt/presentation.xml')['p:presentation'], rels = relations('ppt/presentation.xml');
  const slides = array(presentation['p:sldIdLst']?.['p:sldId']).map((entry, index) => {
    const file = rels.find(rel => rel.Id === entry['r:id'])?.part; assert.ok(file);
    const content = xml(file), links = relations(file), text = texts(content);
    const shapes = deep(content, 'p:sp').map(shape => ({name: shape['p:nvSpPr']?.['p:cNvPr']?.name, placeholder: shape['p:nvSpPr']?.['p:nvPr']?.['p:ph'] ?? null, text: texts(shape), paragraphs: array(shape['p:txBody']?.['a:p']).length, paragraphProperties: array(shape['p:txBody']?.['a:p']).map(paragraph => paragraph['a:pPr'] ?? null), bodyProperties: shape['p:txBody']?.['a:bodyPr'] ?? null, geometryEmu: shape['p:spPr']?.['a:xfrm'] ?? null, explicitLatinTypefaces: [...new Set(deep(shape, 'a:latin').map(value => value.typeface))], explicitSizesHundredthsPt: [...new Set([...deep(shape, 'a:rPr'), ...deep(shape, 'a:defRPr')].map(value => value.sz).filter(value => value !== undefined).map(Number))]}));
    const points = cache => array(cache?.['c:pt']).toSorted((a, b) => Number(a.idx) - Number(b.idx)).map(point => String(point['c:v']));
    const charts = deep(content, 'c:chart').map(chart => {
      const part = links.find(rel => rel.Id === chart['r:id'])?.part; assert.ok(part); const cache = xml(part);
      const series = deep(cache, 'c:ser').map(item => {
        const tx = item['c:tx'], cat = item['c:cat'], val = item['c:val'];
        const categories = points(cat?.['c:strRef']?.['c:strCache'] ?? cat?.['c:strLit'] ?? cat?.['c:numRef']?.['c:numCache'] ?? array(cat?.['c:multiLvlStrRef']?.['c:multiLvlStrCache']?.['c:lvl'])[0]);
        return {name: String(tx?.['c:v'] ?? points(tx?.['c:strRef']?.['c:strCache'])[0] ?? ''), categories, values: points(val?.['c:numRef']?.['c:numCache'] ?? val?.['c:numLit']).map(Number)};
      });
      return {part, series, values: deep(cache, 'c:numCache').flatMap(node => points(node).map(Number)), explicitColors: deep(cache, 'a:srgbClr').map(value => value.val)};
    });
    const notes = links.find(rel => rel.Type.endsWith('/notesSlide'))?.part;
    const pictures = deep(content, 'p:pic').map(picture => {const blip = deep(picture, 'a:blip')[0], part = links.find(rel => rel.Id === blip?.['r:embed'])?.part; return {part, sha256: part ? hash(parts[part]) : null, crop: deep(picture, 'a:srcRect')};});
    return {number: index + 1, part: file, text, marker: text.join(' ').match(/\bK\d{2}\b/)?.[0] ?? null, shapes, pictures, charts, tables: deep(content, 'a:tbl').map(table => ({rows: array(table['a:tr']).map(row => array(row['a:tc']).map(texts))})), notes: notes ? texts(xml(notes)) : []};
  });
  const layouts = Object.keys(parts).filter(name => /^ppt\/(?:slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)).sort().map(part => ({part, placeholders: deep(xml(part), 'p:ph')}));
  return {slides, layouts, nativeChartCount: slides.reduce((count, slide) => count + slide.charts.length, 0), nativeTableCount: slides.reduce((count, slide) => count + slide.tables.length, 0), archiveParts: Object.keys(parts).length};
}

export async function generate({consumer, output}) {
  assert.equal(process.versions.node.split('.')[0], '24', 'Qualification requires Node 24');
  const engine = await installed(consumer), root = await fresh(output, [engine.consumer]);
  const put = async (name, bytes) => {assert.ok(safe(name)); await mkdir(path.dirname(path.join(root, name)), {recursive: true}); await writeFile(path.join(root, name), bytes, {flag: 'wx'});};
  const image = await readFile(new URL('../../test/fixtures/images/wide.png', import.meta.url)), replacement = await readFile(new URL('../../test/fixtures/images/square.png', import.meta.url));
  const source = sourceDeck('data:image/png;base64,' + image.toString('base64')), before = structuredClone(source);
  const validation = engine.core.validate(source, {catalogs: [engine.gallery], only: ['format', 'references']}); assert.ok(validation.valid, JSON.stringify(validation.findings));
  const diagnostics = [], fonts = await engine.fonts.loadFonts({pack: 'office'});
  await fonts.ensure(source, {renderOptions: {catalogs: [engine.gallery]}, onDiagnostic: value => diagnostics.push(value)});
  const previewFaces = [400, 700].flatMap(fontWeight => [false, true].map(italic => fonts.registry.resolveFont({fontFamily: fixtureFontPolicy.selected, fontWeight, italic})));
  assert.ok(previewFaces.every(face => face.requestedFamily === 'Arial' && face.resolvedFamily === 'Arimo'), 'Selected Arial must use the normal licensed Arimo preview alias for all four styles');
  const options = {catalogs: [engine.gallery], fonts, seed: 86, timestamp: '2026-10-15T00:00:00Z', zipDate: '2026-10-15T00:00:00Z', date: '2026-10-15', provenance: 'full', strictAssets: true};
  const bytes = await engine.pptx.toPptx(source, {...options, onDiagnostic: value => diagnostics.push(value)});
  assert.deepEqual(source, before); assert.deepEqual(bytes, await engine.pptx.toPptx(source, options), 'Nondeterministic export');
  const inspection = inspectArchive(bytes, engine); assert.equal(inspection.slides.length, 16); assert.equal(inspection.nativeChartCount, 2); assert.equal(inspection.nativeTableCount, 2);
  const nativeLatinTypefaces = [...new Set(inspection.slides.flatMap(slide => slide.shapes.flatMap(shape => shape.explicitLatinTypefaces)))].sort();
  assert.ok(nativeLatinTypefaces.includes('Arial'), 'Native XML must retain selected Arial');
  assert.ok(!nativeLatinTypefaces.includes('Arimo'), 'Preview alias must not replace selected Arial in native XML');
  const importDiagnostics = [], imported = await engine.pptx.fromPptx(bytes, {catalogs: [engine.gallery], onDiagnostic: value => importDiagnostics.push(value)});
  await put('sources/deck.opf.json', json(source)); await put('inputs/deck.pptx', bytes);
  await put('assets/replacement.png', replacement); await put('assets/original.png', image);
  await put('inspection/source-archive.json', json(inspection));
  await put('inspection/source-import.json', json({document: imported, diagnostics: importDiagnostics, roundTripEqual: isDeepStrictEqual(source, imported), nativeAcceptance: false}));
  await put('inspection/validation.json', json({validation, diagnostics})); await put('actions/routes.json', json(routes));
  for (const route of Object.keys(routes)) await put(`actions/${route}.result-template.json`, json(outcomeTemplate(route)));
  for (let index = 0; index < source.slides.length; index++) await put(`previews/slide-${String(index + 1).padStart(2, '0')}.svg`, engine.render.toSvg(source, index + 1, options));
  const fontFiles = await Promise.all(fonts.fontFiles.map(async file => {const data = await readFile(file); return {file, bytes: data.length, sha256: hash(data)};}));
  const git = args => execFileSync('git', args, {cwd: repository, encoding: 'utf8', stdio: 'pipe'}).trim();
  const sourceFiles = ['scripts/keynote-qualification/fixtures.mjs', 'scripts/keynote-qualification/pack.mjs', 'scripts/keynote-qualification/cli.mjs', 'docs/keynote-qualification.md', 'test/keynote-qualification.mjs', 'test/fixtures/images/README.md', 'test/fixtures/images/wide.png', 'test/fixtures/images/square.png'];
  const suite = {schemaVersion: 1, kind: 'keynote-qualification', nativeAcceptance: false, registryAcceptance: false, runtime: {node: process.version, consumer: engine.consumer, packages: engine.packages}, source: {revision: git(['rev-parse', 'HEAD']), dirty: git(['status', '--porcelain']) !== '', files: await Promise.all(sourceFiles.map(async file => ({path: file, sha256: hash(await readFile(path.join(repository, file)))})))}, exportOptions: {...options, fonts: 'one prepared office handle', catalogs: ['installed @openpresentation/gallery']}, fonts: {manifestSha256: hash(json(fonts.manifest)), files: fontFiles, policy: fixtureFontPolicy, previewFaces, nativeLatinTypefaces, nativePhysicalFace: 'unobserved; supervisor verifies in Keynote', installedByGenerator: false}, directions: Object.keys(outcomeTemplate('no-op').directions), files: await tree(root)};
  await put('suite.json', json(suite)); await verifySeal(root); return suite;
}

export async function prepareCopy({suite: directory, route, output}) {
  const {root} = await verifySeal(directory); assert.ok(Object.hasOwn(routes, route), 'Unknown route');
  const target = await fresh(output, [root]);
  await copyFile(path.join(root, 'inputs/deck.pptx'), path.join(target, 'editable.pptx'), constants.COPYFILE_EXCL);
  await writeFile(path.join(target, 'result.json'), json(outcomeTemplate(route)), {flag: 'wx'});
  await writeFile(path.join(target, 'attempt.json'), json({route, suite: root, suiteSha256: hash(await readFile(path.join(root, 'suite.json'))), originalSha256: hash(await readFile(path.join(root, 'inputs/deck.pptx'))), inputSha256: hash(await readFile(path.join(target, 'editable.pptx'))), instructions: routes[route], nativeAcceptance: false}), {flag: 'wx'});
  return target;
}

export function nativeFindings(inspection, route, baseline) {
  assert.ok(Object.hasOwn(routes, route), 'Unknown route');
  const failures = [], check = (condition, message) => {if (!condition) failures.push(message);};
  check(inspection.slides.length === routes[route].expectedSlides, 'Unexpected exported slide count');
  check(inspection.nativeChartCount === 2, 'Two native chart objects must remain');
  check(inspection.nativeTableCount === 2, 'Two native table objects must remain');
  for (let index = 1; index <= 16; index++) {
    const marker = `K${String(index).padStart(2, '0')}`;
    check(inspection.slides.filter(slide => slide.marker === marker).length === 1, `${marker}: original slide marker missing or duplicated`);
  }
  for (const action of routes[route].actions) {
    const slide = action.id === 'new-slide' ? inspection.slides.find(slide => slide.text.join(' ').includes(action.marker)) : inspection.slides.find(slide => slide.marker === action.slide), text = slide?.text.join(' ') ?? '';
    if (action.marker) check(text.includes(action.marker), `${action.id}: expected edited text missing`);
    if (action.secondMarker) check(text.includes(action.secondMarker), `${action.id}: expected new body missing`);
    if (action.oldMarker) check(!text.includes(action.oldMarker), `${action.id}: replaced old text remains`);
    if (action.notesMarker) check(slide?.notes.join(' ').includes(action.notesMarker), `${action.id}: edited notes missing`);
    if (action.chartSeries) check(slide?.charts.some(chart => chart.series?.some(series => isDeepStrictEqual(series, action.chartSeries))), `${action.id}: edited native chart series/category/value cache missing or unrelated points changed`);
  }
  if (route === 'layout-structure') check(inspection.slides.findIndex(slide => slide.marker === 'K14') === inspection.slides.findIndex(slide => slide.marker === 'K15') + 1, 'K14 must immediately follow K15');
  if (route === 'no-op' && baseline) {
    const normalized = lines => lines.join(' ').replace(/\s+/g, ' ').trim();
    for (const source of baseline.slides) {
      const actual = inspection.slides.find(slide => slide.marker === source.marker);
      check(actual && normalized(actual.text) === normalized(source.text), `${source.marker}: no-op visible text changed (whitespace-normalized comparison)`);
      check(actual && normalized(actual.notes) === normalized(source.notes), `${source.marker}: no-op notes changed (whitespace-normalized comparison)`);
    }
  }
  return failures;
}

const inlineText = value => typeof value === 'string' || typeof value === 'number' ? String(value) : Array.isArray(value) ? value.map(inlineText).join('') : value && typeof value === 'object' ? inlineText(value.text) : '';
// Only current authored slide fields count. Never search serialized provenance, catalogs, extensions or other slides.
export function currentSlideText(slide) {
  if (!slide || typeof slide !== 'object') return '';
  const values = ['title', 'subtitle', 'tag', 'text'].map(field => inlineText(slide[field]));
  values.push(...array(slide.items).map(item => typeof item === 'object' && !Array.isArray(item) ? `${inlineText(item.text)} ${inlineText(item.description)}` : inlineText(item)));
  if (slide.metric !== undefined) values.push(typeof slide.metric === 'object' ? ['value', 'unit', 'label', 'description', 'delta', 'trend'].map(field => inlineText(slide.metric[field])).join(' ') : inlineText(slide.metric));
  if (slide.table) values.push(...[slide.table.columns, ...array(slide.table.rows)].flatMap(row => array(row).map(inlineText)));
  values.push(...array(slide.blocks).map(currentSlideText));
  return values.filter(Boolean).join(' ');
}

export function importedFindings(imported, route) {
  assert.ok(Object.hasOwn(routes, route), 'Unknown route');
  const failures = [];
  for (const action of routes[route].actions) {
    const slide = imported.slides?.find(slide => action.id === 'new-slide' ? currentSlideText(slide).includes(action.marker) : currentSlideText(slide).match(/\bK\d{2}\b/)?.[0] === action.slide);
    const text = currentSlideText(slide);
    if (action.marker && !text.includes(action.marker)) failures.push(`${action.id}: edited marker missing from current imported slide payload`);
    if (action.secondMarker && !text.includes(action.secondMarker)) failures.push(`${action.id}: new body missing from the same current imported slide`);
    if (action.oldMarker && text.includes(action.oldMarker)) failures.push(`${action.id}: replaced old text resurrected in current imported slide`);
    if (action.notesMarker && !inlineText(slide?.notes).includes(action.notesMarker)) failures.push(`${action.id}: edited marker missing from current imported notes`);
    if (action.chartSeries) {
      const charts = node => node ? [...(node.chart ? [node.chart] : []), ...array(node.blocks).flatMap(charts)] : [];
      const expected = action.chartSeries;
      const found = charts(slide).some(chart => {
        const column = chart.data?.columns?.findIndex(value => inlineText(value) === expected.name);
        return column > 0 && isDeepStrictEqual(chart.data.rows.map(row => inlineText(row[0])), expected.categories) && isDeepStrictEqual(chart.data.rows.map(row => row[column]), expected.values);
      });
      if (!found) failures.push(`${action.id}: current imported chart series/category/values differ`);
    }
  }
  return failures;
}

export async function inspectNative({suite: directory, consumer, input, route, output}) {
  const {root, suite} = await verifySeal(directory); assert.ok(Object.hasOwn(routes, route), 'Unknown route');
  const engine = await installed(consumer), target = await fresh(output, [root, engine.consumer]);
  assert.deepEqual(engine.packages.map(({name, version, sha256}) => ({name, version, sha256})), suite.runtime.packages.map(({name, version, sha256}) => ({name, version, sha256})), 'Inspection consumer differs from the sealed installed runtime');
  const bytes = await readFile(input), inspection = inspectArchive(bytes, engine), baseline = JSON.parse(await readFile(path.join(root, 'inspection/source-archive.json')));
  const failures = nativeFindings(inspection, route, baseline);
  const diagnostics = [], imported = await engine.pptx.fromPptx(bytes, {catalogs: [engine.gallery], onDiagnostic: value => diagnostics.push(value)});
  failures.push(...importedFindings(imported, route));
  if (route === 'content-edits') {
    const replacement = hash(await readFile(path.join(root, 'assets/replacement.png'))), actual = inspection.slides.find(slide => slide.marker === 'K08');
    if (!actual?.pictures.some(picture => picture.sha256 === replacement)) failures.push('replace-image: replacement media bytes missing from current slide');
  }
  const report = {scope: 'Offline inspection of the actual native-exported PPTX; no UI, font, raster or native acceptance claim.', route, input: {file: path.resolve(input), bytes: bytes.length, sha256: hash(bytes)}, packages: engine.packages, inspection, failures, staticAssertionsPassed: failures.length === 0, nativeAcceptance: false};
  await writeFile(path.join(target, 'inspection.json'), json(report), {flag: 'wx'});
  await writeFile(path.join(target, 'imported.opf.json'), json(imported), {flag: 'wx'});
  await writeFile(path.join(target, 'import-diagnostics.json'), json(diagnostics), {flag: 'wx'});
  await verifySeal(root); return report;
}
