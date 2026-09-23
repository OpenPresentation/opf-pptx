import {XMLParser} from 'fast-xml-parser';
import {decodeTextTag, encodeTextTag} from './code-provenance.js';

// FF-32: document and slide references survive a PPTX round trip.
//
// PPTX has no native field for an OPF catalog reference (design.colorScheme,
// fontScheme, theme, dimensions and background, a slide layout id) or for
// authoring metadata (narrative, tone, audience, purpose, language,
// organization, speaker ...). They are stored as standard PresentationML
// customer-data tags, the mechanism the other OPF provenance records already
// use: OPF_DOCUMENT_V1 on p:presentation/p:custDataLst and OPF_SLIDE_V1 on each
// slide's p:cSld/p:custDataLst. PowerPoint keeps tags through edits and saves,
// copies slide tags with a slide, and never shows them. Values are UTF-8 JSON as
// uppercase hex, because PowerPoint's Tags API is case-insensitive.
//
// Each design reference records the native evidence it produced (theme
// colors, theme fonts, slide size, slide background, slide object geometry).
// Import restores a reference only while that evidence is unchanged. After an
// edit the observed native values stay and a specific diagnostic names the
// reference that was not restored. Tags are untrusted input: they are parsed,
// size-limited, validated with the whole imported document, and never executed.

export const DOCUMENT_TAG = 'OPF_DOCUMENT_V1';
export const SLIDE_TAG = 'OPF_SLIDE_V1';
const REL_TAGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';
const NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const TAGS_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml';
const MAX_FIELD_BYTES = 256 * 1024, MAX_CATALOG_BYTES = 1024 * 1024, MAX_TAG_CHARS = 16 * 1024 * 1024;

const enc = new TextEncoder(), dec = new TextDecoder('utf-8', {fatal: true});
const parser = new XMLParser({ignoreAttributes: false, attributeNamePrefix: '', textNodeName: '#text', parseAttributeValue: false, parseTagValue: false, trimValues: false});
const array = value => value === undefined ? [] : Array.isArray(value) ? value : [value];
const object = value => value !== null && typeof value === 'object' && !Array.isArray(value);
const canonical = value => JSON.stringify(value, (_key, item) => object(item) ? Object.fromEntries(Object.keys(item).sort().map(key => [key, item[key]])) : item);
const same = (a, b) => canonical(a) === canonical(b);
const clone = value => value === undefined ? undefined : JSON.parse(JSON.stringify(value));
const own = (value, key) => object(value) && Object.hasOwn(value, key) && value[key] !== undefined;

// Deck-level fields. References are gated by native evidence; metadata has no
// native PowerPoint counterpart and round-trips from the stored value.
export const DESIGN_REFERENCES = Object.freeze(['theme', 'colorScheme', 'fontScheme', 'dimensions', 'background']);
export const COMPOSITION_HINTS = Object.freeze(['titleAlignment', 'contentAlignment', 'contentBox', 'contentDirection', 'chartPrimary', 'imageFill', 'listBullet']);
export const METADATA = Object.freeze(['narrative', 'tone', 'audience', 'purpose', 'language', 'organization', 'speaker', 'takeaway', 'duration', 'tags', 'variables']);
const STYLE_REFERENCES = ['theme', 'colorScheme', 'fontScheme'];
const SLIDE_STRUCTURE = ['layout', 'type', 'composition'];

// 53-bit non-cryptographic hash (cyrb53). Change detection only; tags are
// writable by anyone who can edit the file, so no authenticity is claimed.
const hash = text => hashUnits(text.length, index => text.charCodeAt(index));
const hashBytes = bytes => hashUnits(bytes.byteLength, index => bytes[index]);
function hashUnits(length, unit) {
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let index = 0; index < length; index += 1) {
    const code = unit(index);
    h1 = Math.imul(h1 ^ code, 2654435761);
    h2 = Math.imul(h2 ^ code, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, '0');
}

// ---------------------------------------------------------------------------
// Package helpers shared by export (final normalized parts) and import.

function relsPath(part) {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

function resolveTarget(source, target) {
  if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  const raw = target.startsWith('/') ? target.slice(1) : source.slice(0, source.lastIndexOf('/') + 1) + target;
  const parts = [];
  for (const part of raw.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop(); else parts.push(part);
  }
  return parts.join('/');
}

function relationships(entries, part) {
  const map = new Map();
  const bytes = entries[relsPath(part)];
  if (!bytes) return map;
  for (const rel of array(parser.parse(dec.decode(bytes))?.Relationships?.Relationship)) {
    if (!rel?.Id) continue;
    map.set(rel.Id, {type: rel.Type ?? '', external: rel.TargetMode === 'External', path: rel.TargetMode === 'External' ? null : resolveTarget(part, rel.Target ?? '')});
  }
  return map;
}

function slidePaths(entries, presentationRoot, rels) {
  return array(presentationRoot?.['p:sldIdLst']?.['p:sldId']).map(node => rels.get(node?.['r:id']))
    .filter(rel => rel?.type.endsWith('/slide') && rel.path && entries[rel.path]).map(rel => rel.path);
}

function themePath(entries, presentationRoot, rels) {
  const masters = array(presentationRoot?.['p:sldMasterIdLst']?.['p:sldMasterId']).map(node => rels.get(node?.['r:id']))
    .filter(rel => rel?.type.endsWith('/slideMaster') && rel.path && entries[rel.path]);
  for (const master of masters) {
    const theme = [...relationships(entries, master.path).values()].find(rel => rel.type.endsWith('/theme') && rel.path && entries[rel.path]);
    if (theme) return theme.path;
  }
  return [...rels.values()].find(rel => rel.type.endsWith('/theme') && rel.path && entries[rel.path])?.path ?? null;
}

const THEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

function colorValue(node) {
  if (!object(node)) return null;
  if (object(node['a:srgbClr']) || typeof node['a:srgbClr'] === 'string') {
    const color = node['a:srgbClr'], {val, ...rest} = object(color) ? color : {};
    return `srgb:${String(val ?? '').toUpperCase()}${Object.keys(rest).length ? canonical(rest) : ''}`;
  }
  // PowerPoint rewrites lastClr with the current system color; the reference is what persists.
  if (object(node['a:sysClr'])) return `sys:${node['a:sysClr'].val ?? ''}`;
  return canonical(withoutExtensions(node));
}

// Theme colors, theme fonts and slide size are the native counterparts of
// design.colorScheme, design.fontScheme (both via design.theme) and design.dimensions.
function nativeDocument(entries, presentationRoot, rels) {
  const path = themePath(entries, presentationRoot, rels);
  let elements;
  try { elements = path ? parser.parse(dec.decode(entries[path]))?.['a:theme']?.['a:themeElements'] : undefined; } catch { elements = undefined; }
  const scheme = elements?.['a:clrScheme'], fontScheme = elements?.['a:fontScheme'];
  const colors = object(scheme) ? Object.fromEntries(THEME_SLOTS.map(slot => [slot, colorValue(scheme[`a:${slot}`])])) : null;
  const face = (font, script) => { const node = fontScheme?.[font]?.[`a:${script}`]; return object(node) ? String(node.typeface ?? '') : null; };
  const fonts = object(fontScheme) ? Object.fromEntries(['major', 'minor'].map(kind => [kind, Object.fromEntries(['latin', 'ea', 'cs'].map(script => [script, face(`a:${kind}Font`, script)]))])) : null;
  const size = presentationRoot?.['p:sldSz'];
  return {colors, fonts, size: object(size) ? {cx: String(size.cx ?? ''), cy: String(size.cy ?? '')} : null};
}

function withoutExtensions(node) {
  if (Array.isArray(node)) return node.map(withoutExtensions);
  if (!object(node)) return node;
  return Object.fromEntries(Object.entries(node).filter(([key]) => key !== 'a:extLst' && key !== 'p:extLst').map(([key, value]) => [key, withoutExtensions(value)]));
}

// Relationship ids are renumbered by editors; the target bytes identify an image.
function withResolvedRelationships(node, rels, entries) {
  if (Array.isArray(node)) return node.map(item => withResolvedRelationships(item, rels, entries));
  if (!object(node)) return node;
  return Object.fromEntries(Object.entries(node).map(([key, value]) => {
    if (/^r:(embed|link|id)$/.test(key) && typeof value === 'string') {
      const rel = rels.get(value);
      const bytes = rel?.path ? entries[rel.path] : null;
      return [key, bytes ? `bytes:${bytes.byteLength}:${hashBytes(bytes)}` : `rel:${rel?.type ?? ''}`];
    }
    return [key, withResolvedRelationships(value, rels, entries)];
  }));
}

function box(xfrm, frame) {
  if (!object(xfrm)) return 'inherit';
  const off = xfrm['a:off'], ext = xfrm['a:ext'];
  // Table rows grow when a viewer reflows cell text, so a table frame's height
  // is not layout structure. Position and width are.
  return [off?.x, off?.y, ext?.cx, frame ? '' : ext?.cy, xfrm.rot ?? '0', xfrm.flipH ?? '0', xfrm.flipV ?? '0'].map(value => String(value ?? '')).join(',');
}

function structure(tree) {
  const kinds = {};
  const add = (kind, value) => (kinds[kind] ??= []).push(value);
  for (const node of array(tree?.['p:sp'])) add('sp', box(node?.['p:spPr']?.['a:xfrm']));
  for (const node of array(tree?.['p:pic'])) add('pic', box(node?.['p:spPr']?.['a:xfrm']));
  for (const node of array(tree?.['p:cxnSp'])) add('cxnSp', box(node?.['p:spPr']?.['a:xfrm']));
  for (const node of array(tree?.['p:graphicFrame'])) add('graphicFrame', box(node?.['p:xfrm'], true));
  for (const node of array(tree?.['p:grpSp'])) add('grpSp', `${box(node?.['p:grpSpPr']?.['a:xfrm'])}[${structure(node)}]`);
  for (const kind of ['p:contentPart', 'mc:AlternateContent']) if (tree?.[kind] !== undefined) add(kind, String(array(tree[kind]).length));
  // Stacking order is not structure; object kinds, positions and sizes are.
  return canonical(Object.fromEntries(Object.entries(kinds).map(([kind, values]) => [kind, values.sort()])));
}

function styleSignature(tree) {
  const faces = new Set(), colors = new Set();
  const visit = node => {
    if (Array.isArray(node)) { node.forEach(visit); return; }
    if (!object(node)) return;
    for (const [key, value] of Object.entries(node)) {
      if (/^a:(latin|ea|cs|sym)$/.test(key)) for (const font of array(value)) if (object(font) && font.typeface !== undefined) faces.add(String(font.typeface));
      if (/^a:(srgbClr|schemeClr|sysClr|prstClr)$/.test(key)) for (const color of array(value)) colors.add(`${key}:${object(color) ? String(color.val ?? '').toUpperCase() : ''}`);
      visit(value);
    }
  };
  visit(tree);
  return canonical({faces: [...faces].sort(), colors: [...colors].sort()});
}

function nativeSlide(entries, path) {
  const root = parser.parse(dec.decode(entries[path]))?.['p:sld'];
  const cSld = root?.['p:cSld'], tree = cSld?.['p:spTree'];
  const rels = relationships(entries, path);
  return {
    structure: hash(structure(tree)),
    background: hash(canonical(withResolvedRelationships(withoutExtensions(cSld?.['p:bg'] ?? null), rels, entries))),
    style: hash(styleSignature(withoutExtensions(tree ?? null)))
  };
}

// ---------------------------------------------------------------------------
// Export

function sizeOf(value) {
  return enc.encode(JSON.stringify(value)).byteLength;
}

function collectStrings(value, into = new Set()) {
  if (typeof value === 'string') into.add(value);
  else if (Array.isArray(value)) value.forEach(item => collectStrings(item, into));
  else if (object(value)) Object.values(value).forEach(item => collectStrings(item, into));
  return into;
}

// Inline catalog records are kept only when the document references their id,
// so restored ids resolve without copying unrelated catalog content.
function pruneCatalogs(catalogs, referenced) {
  if (!object(catalogs)) return undefined;
  const result = {};
  for (const [kind, entry] of Object.entries(catalogs)) {
    if (Array.isArray(entry)) {
      const records = entry.filter(record => typeof record?.id === 'string' && referenced.has(record.id));
      if (records.length) result[kind] = records;
    } else if (object(entry)) {
      const records = array(entry.records).filter(record => typeof record?.id === 'string' && referenced.has(record.id));
      const kept = {...(entry.source !== undefined ? {source: entry.source} : {}), ...(records.length ? {records} : {})};
      if (Object.keys(kept).length) result[kind] = kept;
    }
  }
  return Object.keys(result).length ? result : undefined;
}

function catalogReferences(catalogs) {
  const ids = new Set();
  for (const entry of Object.values(catalogs ?? {})) for (const record of array(Array.isArray(entry) ? entry : entry?.records)) collectStrings(record, ids);
  return ids;
}

function assetReferences(value) {
  return [...collectStrings(value)].filter(item => item.startsWith('asset:')).map(item => item.slice(6));
}

/**
 * The references and metadata to persist. Only values the document states are
 * stored; engine defaults are not. Returns null when there is nothing to keep.
 */
export function documentProvenance(presentation, report = () => {}) {
  const design = {}, metadata = {};
  for (const key of [...DESIGN_REFERENCES, ...COMPOSITION_HINTS]) if (own(presentation.design, key)) design[key] = clone(presentation.design[key]);
  for (const key of METADATA) {
    if (!own(presentation, key)) continue;
    const value = clone(presentation[key]);
    if (sizeOf(value) > MAX_FIELD_BYTES) report({code: 'document-provenance-omitted', path: key, message: `${key} is larger than ${MAX_FIELD_BYTES / 1024} KiB and is not stored in the PPTX; reimport cannot restore it.`});
    else metadata[key] = value;
  }
  const slides = presentation.slides.map((slide, index) => {
    const record = {v: 1, slide: index};
    for (const key of ['id', 'beat', ...SLIDE_STRUCTURE]) if (own(slide, key)) record[key] = clone(slide[key]);
    const slideDesign = {};
    for (const key of [...STYLE_REFERENCES, 'background', ...COMPOSITION_HINTS]) if (own(slide.design, key)) slideDesign[key] = clone(slide.design[key]);
    if (Object.keys(slideDesign).length) record.design = slideDesign;
    return record;
  });
  const stated = Object.keys(design).length || Object.keys(metadata).length || slides.some(record => Object.keys(record).length > 2);
  if (!stated) return null;
  const document = {v: 1, slides: slides.length};
  if (Object.keys(design).length) document.design = design;
  if (Object.keys(metadata).length) document.metadata = metadata;
  const assetIds = assetReferences(metadata).filter(id => own(presentation.assets, id));
  if (assetIds.length) {
    const assets = Object.fromEntries(assetIds.map(id => [id, clone(presentation.assets[id])]));
    if (sizeOf(assets) > MAX_FIELD_BYTES) report({code: 'document-provenance-omitted', path: 'assets', message: 'Assets referenced by document metadata are too large to store in the PPTX; reimport keeps their asset: references without the registry entries.'});
    else document.assets = assets;
  }
  const {catalogs, ...rest} = presentation;
  const referenced = collectStrings(rest);
  // A referenced record can itself refer to another inline record (a theme to its color scheme).
  let pruned = pruneCatalogs(catalogs, referenced);
  for (let pass = 0; pass < 3 && pruned; pass += 1) {
    const next = pruneCatalogs(catalogs, new Set([...referenced, ...catalogReferences(pruned)]));
    if (same(next, pruned)) break;
    pruned = next;
  }
  if (pruned) {
    if (sizeOf(pruned) > MAX_CATALOG_BYTES) report({code: 'document-provenance-omitted', path: 'catalogs', message: 'Referenced inline catalog records are larger than 1 MiB and are not stored in the PPTX; restored ids may need their catalogs supplied again.'});
    else document.catalogs = pruned;
  }
  return {document, slides};
}

function tagPart(name, value) {
  return enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="${NS}"><p:tag name="${name}" val="${encodeTextTag(value)}"/></p:tagLst>`);
}

function freeId(relsXml, base) {
  const ids = new Set([...relsXml.matchAll(/\bId="([^"]+)"/g)].map(match => match[1]));
  let id = base;
  while (ids.has(id)) id += '_';
  return id;
}

/**
 * Record native evidence from the final normalized parts and write the tags.
 * `entries` maps part paths to bytes and is updated in place.
 */
export function attachDocumentProvenance(entries, provenance) {
  if (!provenance) return;
  const presentationXml = dec.decode(entries['ppt/presentation.xml']);
  const presentationRoot = parser.parse(presentationXml)['p:presentation'];
  const rels = relationships(entries, 'ppt/presentation.xml');
  const paths = slidePaths(entries, presentationRoot, rels);
  if (paths.length !== provenance.slides.length) throw Error('Generated slide count differs from the document.');
  const types = [];
  const document = {...provenance.document, native: nativeDocument(entries, presentationRoot, rels)};

  // CT_Presentation: custDataLst follows photoAlbum and precedes kinsoku/defaultTextStyle/modifyVerifier/extLst.
  const documentPart = 'ppt/tags/opfDocument.xml';
  if (entries[documentPart] || /<p:custDataLst\b/.test(presentationXml)) throw Error('Colliding presentation customer data.');
  let presentationRels = dec.decode(entries['ppt/_rels/presentation.xml.rels']);
  const documentId = freeId(presentationRels, 'rIdOpfDocument');
  const anchor = presentationXml.search(/<p:(kinsoku|defaultTextStyle|modifyVerifier|extLst)\b|<\/p:presentation>/);
  if (anchor < 0) throw Error('Generated presentation part has no closing element.');
  entries['ppt/presentation.xml'] = enc.encode(`${presentationXml.slice(0, anchor)}<p:custDataLst><p:tags r:id="${documentId}"/></p:custDataLst>${presentationXml.slice(anchor)}`);
  entries['ppt/_rels/presentation.xml.rels'] = enc.encode(presentationRels.replace('</Relationships>', `<Relationship Id="${documentId}" Type="${REL_TAGS}" Target="tags/opfDocument.xml"/></Relationships>`));
  entries[documentPart] = tagPart(DOCUMENT_TAG, document);
  types.push(documentPart);

  for (const [index, path] of paths.entries()) {
    const record = {...provenance.slides[index], native: nativeSlide(entries, path)};
    const tag = `<p:tag name="${SLIDE_TAG}" val="${encodeTextTag(record)}"/>`;
    let xml = dec.decode(entries[path]);
    const slideRels = relsPath(path);
    let relsXml = dec.decode(entries[slideRels]);
    // One p:tags per custDataLst: join the slide's existing tag list (furniture) when present.
    const existing = xml.match(/<\/p:spTree>\s*<p:custDataLst>\s*<p:tags r:id="([^"]+)"\s*\/>/);
    if (existing) {
      const target = relationships(entries, path).get(existing[1]);
      if (target?.type !== REL_TAGS || !target.path || !entries[target.path]) throw Error('Generated slide tag relationship is missing.');
      entries[target.path] = enc.encode(dec.decode(entries[target.path]).replace('</p:tagLst>', `${tag}</p:tagLst>`));
      continue;
    }
    if (!xml.includes('</p:spTree>') || /<\/p:spTree>\s*<p:custDataLst\b/.test(xml)) throw Error('Generated slide has no shape tree or unexpected customer data.');
    const part = `ppt/tags/opfSlide${index + 1}.xml`;
    if (entries[part]) throw Error('Colliding slide tag part.');
    const id = freeId(relsXml, 'rIdOpfSlide');
    entries[path] = enc.encode(xml.replace('</p:spTree>', `</p:spTree><p:custDataLst><p:tags r:id="${id}"/></p:custDataLst>`));
    entries[slideRels] = enc.encode(relsXml.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL_TAGS}" Target="../tags/opfSlide${index + 1}.xml"/></Relationships>`));
    entries[part] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="${NS}">${tag}</p:tagLst>`);
    types.push(part);
  }
  const contentTypes = dec.decode(entries['[Content_Types].xml']);
  entries['[Content_Types].xml'] = enc.encode(contentTypes.replace('</Types>', types.map(part => `<Override PartName="/${part}" ContentType="${TAGS_TYPE}"/>`).join('') + '</Types>'));
}

// ---------------------------------------------------------------------------
// Import

function readTag(entries, container, rels, name) {
  const found = [];
  let unreadable = false;
  for (const link of array(container?.['p:tags'])) {
    const rel = rels.get(link?.['r:id']);
    if (rel?.type !== REL_TAGS || rel.external || !rel.path || !entries[rel.path]) { unreadable = true; continue; }
    try { found.push(...array(parser.parse(dec.decode(entries[rel.path]))?.['p:tagLst']?.['p:tag']).filter(tag => String(tag?.name ?? '').toUpperCase() === name)); }
    catch { unreadable = true; }
  }
  if (!found.length) return {missing: true, unreadable};
  if (found.length > 1) throw Error(`Multiple ${name} tags.`);
  const value = found[0].val;
  if (typeof value !== 'string' || value.length > MAX_TAG_CHARS) throw Error(`Oversized or empty ${name} tag.`);
  return {value: decodeTextTag(value)};
}

const validNative = native => object(native);

function validateDocument(value) {
  if (!object(value) || value.v !== 1 || !Number.isSafeInteger(value.slides) || value.slides < 1) throw Error('Unsupported document provenance version.');
  for (const key of ['design', 'metadata', 'catalogs', 'assets']) if (value[key] !== undefined && !object(value[key])) throw Error(`Invalid ${key} record.`);
  if (!validNative(value.native)) throw Error('Missing native evidence.');
  for (const key of Object.keys(value.design ?? {})) if (![...DESIGN_REFERENCES, ...COMPOSITION_HINTS].includes(key)) throw Error(`Unknown design field ${key}.`);
  for (const key of Object.keys(value.metadata ?? {})) if (!METADATA.includes(key)) throw Error(`Unknown metadata field ${key}.`);
  return value;
}

function validateSlide(value) {
  if (!object(value) || value.v !== 1 || !Number.isSafeInteger(value.slide) || value.slide < 0) throw Error('Unsupported slide provenance version.');
  if (!validNative(value.native) || typeof value.native.structure !== 'string' || typeof value.native.background !== 'string' || typeof value.native.style !== 'string') throw Error('Missing slide native evidence.');
  if (value.design !== undefined && !object(value.design)) throw Error('Invalid slide design record.');
  for (const key of Object.keys(value.design ?? {})) if (![...STYLE_REFERENCES, 'background', ...COMPOSITION_HINTS].includes(key)) throw Error(`Unknown slide design field ${key}.`);
  return value;
}

const label = value => typeof value === 'string' ? `'${value}'` : object(value) && typeof value.id === 'string' ? `'${value.id}' (with inline overrides)` : 'inline value';

function changedSlots(stored, observed) {
  if (!object(stored) || !object(observed)) return 'theme missing';
  const keys = [...new Set([...Object.keys(stored), ...Object.keys(observed)])].filter(key => !same(stored[key], observed[key]));
  return keys.join(', ') || 'unreadable';
}

/**
 * Read OPF_DOCUMENT_V1 / OPF_SLIDE_V1 and restore what still matches the
 * package. `imported` is modified in place. Returns per-slide layout evidence
 * ({layout, structure: 'match' | 'changed' | 'untagged'}) for layout recovery.
 */
export function restoreDocumentProvenance(imported, {entries, presentationRoot, presentationRels, slides, organizationConflict = false}, report) {
  const invalid = message => report({code: 'invalid-document-provenance', path: '', message: `${message} Ordinary import keeps the values observed in the PPTX.`});
  let document;
  try {
    const found = readTag(entries, presentationRoot?.['p:custDataLst'], presentationRels, DOCUMENT_TAG);
    if (found.missing) {
      if (found.unreadable) invalid('Presentation customer data could not be read.');
      return {slides: slides.map(() => ({structure: 'untagged'}))};
    }
    document = validateDocument(found.value);
  } catch (error) { invalid(`${error.message}`); return {slides: slides.map(() => ({structure: 'untagged'}))}; }

  const slideRecords = slides.map(({root, relationships: rels, path}, index) => {
    try {
      const found = readTag(entries, root?.['p:cSld']?.['p:custDataLst'], rels, SLIDE_TAG);
      if (found.missing) return null;
      return {record: validateSlide(found.value), native: nativeSlide(entries, path)};
    } catch (error) {
      report({code: 'invalid-document-provenance', path: `slides.${index}`, message: `${error.message} This slide keeps the values observed in the PPTX.`});
      return null;
    }
  });

  const native = nativeDocument(entries, presentationRoot, presentationRels);
  const stored = document.native;
  const match = {colors: same(stored.colors, native.colors) && native.colors !== null, fonts: same(stored.fonts, native.fonts) && native.fonts !== null, size: same(stored.size, native.size)};
  const design = {...(imported.design ?? {})};
  const changed = (path, message) => report({code: 'design-reference-changed', path, message});
  const storedDesign = document.design ?? {};

  const gates = {
    theme: [match.colors && match.fonts, () => `The PPTX theme colors or fonts changed since export (${[match.colors ? '' : changedSlots(stored.colors, native.colors), match.fonts ? '' : `${changedSlots(stored.fonts, native.fonts)} font`].filter(Boolean).join('; ')}).`],
    colorScheme: [match.colors, () => `The PPTX theme colors changed since export (${changedSlots(stored.colors, native.colors)}).`],
    fontScheme: [match.fonts, () => `The PPTX theme fonts changed since export (${changedSlots(stored.fonts, native.fonts)} font).`],
    dimensions: [match.size, () => `The PPTX slide size changed since export.`]
  };
  const restoredStyle = {};
  for (const key of ['theme', 'colorScheme', 'fontScheme', 'dimensions']) {
    if (storedDesign[key] === undefined) continue;
    const [ok, why] = gates[key];
    if (ok) { design[key] = clone(storedDesign[key]); restoredStyle[key] = true; }
    else { restoredStyle[key] = false; changed(`design.${key}`, `${why()} design.${key} ${label(storedDesign[key])} was not restored; the imported document keeps the values observed in the PPTX.`); }
  }
  const styleIntact = STYLE_REFERENCES.every(key => restoredStyle[key] !== false);

  // Slide-level references.
  const layouts = [];
  const seenIds = new Set();
  let allStructure = slideRecords.length === document.slides && slideRecords.every(Boolean);
  const backgroundInheritors = [];
  for (const [index, slide] of imported.slides.entries()) {
    const entry = slideRecords[index];
    if (!entry) { layouts.push({structure: 'untagged'}); allStructure = false; continue; }
    const {record, native: observed} = entry;
    const structureMatch = record.native.structure === observed.structure;
    if (!structureMatch) allStructure = false;
    layouts.push({layout: typeof record.layout === 'string' ? record.layout : undefined, structure: structureMatch ? 'match' : 'changed'});
    if (record.id !== undefined) {
      if (seenIds.has(record.id)) report({code: 'duplicate-slide-id', path: `slides.${index}.id`, message: `Slide id '${record.id}' appears on more than one slide (a duplicated slide); the first keeps it.`});
      else { seenIds.add(record.id); slide.id = clone(record.id); }
    }
    if (record.beat !== undefined) slide.beat = clone(record.beat);
    for (const key of SLIDE_STRUCTURE) {
      if (record[key] === undefined) continue;
      if (structureMatch) slide[key] = clone(record[key]);
      else report({code: key === 'layout' ? 'layout-reference-changed' : 'slide-reference-changed', path: `slides.${index}.${key}`,
        message: `The slide's objects were moved, resized, added or removed since export, so slides.${index}.${key} ${label(record[key])} was not restored; the imported slide keeps its observed arrangement.`});
    }
    const slideDesign = {...(slide.design ?? {})};
    const recordDesign = record.design ?? {};
    for (const key of COMPOSITION_HINTS) {
      if (recordDesign[key] === undefined) continue;
      if (structureMatch) slideDesign[key] = clone(recordDesign[key]);
      else changed(`slides.${index}.design.${key}`, `The slide's objects changed since export, so slides.${index}.design.${key} was not restored.`);
    }
    const styleMatch = record.native.style === observed.style;
    for (const key of STYLE_REFERENCES) {
      if (recordDesign[key] === undefined) continue;
      if (styleMatch) slideDesign[key] = clone(recordDesign[key]);
      else changed(`slides.${index}.design.${key}`, `The slide's fonts or colors changed since export, so slides.${index}.design.${key} ${label(recordDesign[key])} was not restored; the slide keeps its observed formatting.`);
    }
    const backgroundMatch = record.native.background === observed.background;
    if (recordDesign.background !== undefined) {
      if (backgroundMatch) slideDesign.background = clone(recordDesign.background);
      else changed(`slides.${index}.design.background`, `The slide background changed since export, so slides.${index}.design.background ${label(recordDesign.background)} was not restored; the slide keeps its observed background.`);
    } else backgroundInheritors.push({index, backgroundMatch, slideDesign});
    slide.design = slideDesign;
  }

  // A deck background is evidenced by the slides that inherit it. It is kept
  // while at least one of them still shows it (edited slides keep a local
  // override) or when every slide overrides it.
  let deckBackground = storedDesign.background === undefined;
  if (storedDesign.background !== undefined) {
    if (!backgroundInheritors.length || backgroundInheritors.some(item => item.backgroundMatch)) { design.background = clone(storedDesign.background); deckBackground = true; }
    else changed('design.background', `Every slide that inherited the deck background now shows a different background, so design.background ${label(storedDesign.background)} was not restored; slides keep their observed backgrounds.`);
  }
  if (deckBackground && styleIntact) {
    for (const {index, backgroundMatch, slideDesign} of backgroundInheritors) {
      if (backgroundMatch) delete slideDesign.background;
      else if (storedDesign.background !== undefined) changed(`slides.${index}.design.background`, `This slide's background changed since export; it is kept as a slide background override instead of inheriting design.background.`);
    }
  }
  for (const slide of imported.slides) if (object(slide.design) && !Object.keys(slide.design).length) delete slide.design;

  for (const key of COMPOSITION_HINTS) {
    if (storedDesign[key] === undefined) continue;
    if (allStructure) design[key] = clone(storedDesign[key]);
    else changed(`design.${key}`, `Slides were added, removed or rearranged since export, so the deck default design.${key} was not restored.`);
  }
  if (Object.keys(design).length) imported.design = design;

  // Authoring metadata has no native counterpart; current native text still
  // wins for the organization name shown in furniture.
  const metadata = document.metadata ?? {};
  for (const key of METADATA) {
    if (metadata[key] === undefined) continue;
    let value = clone(metadata[key]);
    if (key === 'organization' && organizationConflict) {
      report({code: 'metadata-reference-changed', path: 'organization', message: 'Slides now show different organization names in their furniture, so the stored organization was not restored; each slide keeps its visible text.'});
      continue;
    }
    if (key === 'organization' && object(imported.organization)) {
      // Fields read from native content (furniture name, linked socials) win
      // over their stored values; stored-only fields (logo, role ...) return.
      const observed = imported.organization;
      const list = array(value);
      const index = list.findIndex(item => object(item) && item.id === observed.id);
      if (index < 0) value = clone(observed);
      else if (Array.isArray(value)) value[index] = {...list[index], ...clone(observed)};
      else value = {...value, ...clone(observed)};
    }
    imported[key] = value;
  }
  if (object(document.assets)) {
    const assets = {...(object(imported.assets) ? imported.assets : {})};
    for (const [id, asset] of Object.entries(document.assets)) if (!Object.hasOwn(assets, id)) assets[id] = clone(asset);
    if (Object.keys(assets).length) imported.assets = assets;
  }
  if (object(document.catalogs)) {
    const {catalogs: _ignored, ...rest} = imported;
    const referenced = collectStrings(rest);
    let pruned = pruneCatalogs(document.catalogs, referenced);
    for (let pass = 0; pass < 3 && pruned; pass += 1) {
      const next = pruneCatalogs(document.catalogs, new Set([...referenced, ...catalogReferences(pruned)]));
      if (same(next, pruned)) break;
      pruned = next;
    }
    if (pruned) imported.catalogs = clone(pruned);
  }
  return {slides: layouts};
}
