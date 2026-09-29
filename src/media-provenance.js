import {XMLParser} from 'fast-xml-parser';
import {validatePresentation} from '@openpresentation/opf';
import {attachTextTags, decodeTextTag} from './code-provenance.js';

// A video payload exports as the preview's placeholder: a frame, a play badge,
// a play triangle and caption lines (FF-29). PPTX has no native field for the
// OPF video value, so every shape of the group carries one OPF_MEDIA_V1 shape
// tag and the frame's tag holds the authored value. Import recognizes an
// unchanged, fully tagged group and rebuilds the video block; any other tagged
// group drops only its tagged decoration shapes, keeps the caption text and
// reports invalid-media-provenance. A shape name alone never removes a shape.
const TAG = 'OPF_MEDIA_V1', REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';
const ROLES = ['frame', 'badge', 'play', 'caption'];
const MAX_VALUE_BYTES = 256 * 1024;
const parser = new XMLParser({ignoreAttributes: false, attributeNamePrefix: '', parseTagValue: false, trimValues: false});
const decoder = new TextDecoder('utf-8', {fatal: true}), encoder = new TextEncoder();
const array = value => value === undefined ? [] : Array.isArray(value) ? value : [value];
const object = value => value !== null && typeof value === 'object' && !Array.isArray(value);
const sizeOf = value => encoder.encode(JSON.stringify(value)).byteLength;

// A change detector, not authentication (the surrounding tags are writable).
// Do not copy a dereferenced asset URL into restricted-mode provenance tags.
export function mediaHrefFingerprint(value) {
  if (value === null) return null;
  if (typeof value !== 'string') return undefined;
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < value.length; i++) {
    h1 = Math.imul(h1 ^ value.charCodeAt(i), 2654435761);
    h2 = Math.imul(h2 ^ value.charCodeAt(i), 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, '0');
}

/** The caption the placeholder shows: title, else source, else "Media". */
export function mediaCaption(value) {
  const asset = typeof value === 'string' ? {src: value} : object(value) ? value : {src: ''};
  return asset.title || asset.src || 'Media';
}

/**
 * The frame record: the authored value and, in 'full' provenance mode, the
 * asset registry entries it references (without data: sources), so an
 * `asset:` reference still resolves after import. Oversized values are not
 * stored; import then reports the video as not restored.
 */
export function mediaFrameRecord(value, path, presentation, mode) {
  const record = {v: 1, role: 'frame', path};
  if (sizeOf(value) > MAX_VALUE_BYTES) return {...record, omitted: true};
  record.video = value;
  if (mode === 'full') {
    const assets = {};
    let source = typeof value === 'string' ? value : value?.src;
    for (let depth = 0; typeof source === 'string' && source.startsWith('asset:') && depth < 8; depth++) {
      const id = source.slice(6), asset = presentation.assets?.[id];
      if (asset === undefined || Object.hasOwn(assets, id) || JSON.stringify(asset).includes('data:') || sizeOf(asset) > MAX_VALUE_BYTES) break;
      assets[id] = asset;
      source = typeof asset === 'string' ? asset : asset?.src;
    }
    if (Object.keys(assets).length) record.assets = assets;
  }
  return record;
}

export function attachMediaTags(entries, records) {
  attachTextTags(entries, records, TAG, 'opfMedia', 'media');
}

function readTags(shape, relationships, entries) {
  const tags = [];
  let unreadable = false;
  for (const link of array(shape['p:nvSpPr']?.['p:nvPr']?.['p:custDataLst']?.['p:tags'])) {
    const rel = relationships.get(link['r:id']);
    if (rel?.type !== REL || rel.targetMode === 'External' || !entries[rel.path]) { unreadable = true; continue; }
    try { tags.push(...array(parser.parse(decoder.decode(entries[rel.path]))['p:tagLst']?.['p:tag'])); } catch { unreadable = true; }
  }
  return {tags, unreadable};
}

const validAsset = (id, asset) => /^[a-zA-Z0-9][a-zA-Z0-9._-]*$/.test(id) && validatePresentation({$schema: 'https://openpresentation.org/schema/opf/v1', name: 'Media', assets: {[id]: asset}, slides: [{title: 'Media'}]}).valid;
const validVideo = value => validatePresentation({$schema: 'https://openpresentation.org/schema/opf/v1', name: 'Media', slides: [{video: value}]}).valid;

// null means no native link; undefined means an ambiguous or unreadable link.
// Compare the resolved target, never the relationship id (Office may renumber it).
function frameHyperlink(shape, relationships) {
  const links = array(shape?.['p:nvSpPr']?.['p:cNvPr']?.['a:hlinkClick']);
  if (!links.length) return null;
  if (links.length !== 1) return undefined;
  const rel = relationships.get(links[0]?.['r:id']);
  return rel?.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
    && rel.targetMode === 'External' && typeof rel.target === 'string' ? rel.target : undefined;
}

/**
 * Returns {items, consumed, assets, captionShapes}. items are {shape, payload} for restored
 * video blocks (the frame supplies the position); consumed holds every shape
 * import must skip; assets are registry entries to add when absent.
 */
export function importMediaGroups(shapes, paragraphs, relationships, entries, slideIndex, report) {
  const groups = new Map(), items = [], consumed = new Set(), assets = {}, captionShapes = new Set();
  for (const [index, shape] of shapes.entries()) {
    const {tags, unreadable} = readTags(shape, relationships, entries);
    const own = tags.filter(tag => String(tag.name ?? '').toUpperCase() === TAG);
    if (!own.length) continue;
    let record = null;
    try {
      if (unreadable || own.length !== 1 || tags.some(tag => /^OPF_/i.test(tag.name ?? '') && String(tag.name).toUpperCase() !== TAG)) throw Error('ambiguous');
      record = decodeTextTag(own[0].val);
      if (!object(record) || record.v !== 1 || !ROLES.includes(record.role) || typeof record.path !== 'string' || !record.path.startsWith(`slides.${slideIndex}.`)) throw Error('invalid');
    } catch { record = null; }
    const key = record?.path ?? `slides.${slideIndex}`;
    if (!groups.has(key)) groups.set(key, {path: key, members: [], invalid: false});
    const group = groups.get(key);
    if (!record) { group.invalid = true; group.members.push({index, shape, role: null}); continue; }
    if (record.role === 'caption') captionShapes.add(shape);
    group.members.push({index, shape, record, role: record.role});
  }
  for (const group of groups.values()) {
    const name = (member, suffix) => member.shape['p:nvSpPr']?.['p:cNvPr']?.name === `OPF media ${group.path} ${suffix}`;
    const text = member => (paragraphs[member.index] ?? []).map(paragraph => paragraph.text).join('\n');
    const byRole = role => group.members.filter(member => member.role === role);
    const [frame] = byRole('frame'), captions = byRole('caption').sort((a, b) => a.record.line - b.record.line);
    const decorations = group.members.filter(member => ['frame', 'badge', 'play'].includes(member.role));
    const currentHref = frameHyperlink(frame?.shape, relationships);
    let reason = null;
    if (group.invalid) reason = 'a media tag is ambiguous or unreadable';
    else if (byRole('frame').length !== 1 || byRole('badge').length !== 1 || byRole('play').length !== 1 || !captions.length) reason = 'placeholder shapes were removed or duplicated';
    else if (!decorations.every(member => name(member, member.role) && !text(member)) || !captions.every(member => name(member, `caption line ${member.record.line}`))) reason = 'placeholder shapes were renamed or given text';
    else if (frame.record.omitted || !Object.hasOwn(frame.record, 'video')) reason = 'the video value was too large to store at export';
    else if (!validVideo(frame.record.video)) reason = 'the stored video value is not valid OPF';
    else if (!captions.every((member, index) => member.record.line === index && member.record.count === captions.length
      && typeof member.record.nativeText === 'string' && text(member) === member.record.nativeText)) reason = 'its caption text or line sequence changed, or exact caption evidence is missing';
    else if (!(frame.record.nativeHref === null || typeof frame.record.nativeHref === 'string')
      || mediaHrefFingerprint(currentHref) !== frame.record.nativeHref) reason = 'its source hyperlink changed or its hyperlink evidence is missing';
    if (!reason) {
      for (const member of group.members) consumed.add(member.shape);
      items.push({shape: frame.shape, payload: {type: 'video', video: structuredClone(frame.record.video)}});
      if (object(frame.record.assets)) for (const [id, asset] of Object.entries(frame.record.assets)) if (validAsset(id, asset)) assets[id] ??= structuredClone(asset);
      continue;
    }
    // Tagged decoration carries no content; caption lines stay ordinary text.
    for (const member of group.members) if (member.role && member.role !== 'caption' && !text(member)) consumed.add(member.shape);
    // Every current frame hyperlink is authored content, even when only its
    // caption changed or old tags lack evidence. Keep web targets as editable
    // linked text; other schemes remain inert plain text.
    for (const member of byRole('frame')) {
      const href = frameHyperlink(member.shape, relationships);
      if (typeof href === 'string') items.push({shape: member.shape, payload: {type: 'text', text: /^https?:\/\//i.test(href)
        ? [{text: href, link: href}] : href}});
    }
    report({code: 'invalid-media-provenance', path: group.path,
      message: `The video placeholder at ${group.path} was not restored as a video because ${reason}. Its caption stays as text; the video source is available as the placeholder's hyperlink when it was a web URL.`});
  }
  return {items, consumed, assets, captionShapes};
}
