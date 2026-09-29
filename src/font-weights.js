// One table of the weight and posture words that native legacy font families use
// as style-link names ("Roboto SemiBold"). The exporter recognizes them to keep
// a chosen family's own faces (sameTypeface) and the importer uses the same
// table to map such a face back to family plus OPF weight (splitWeightFace).
// Weights are CSS numeric weights. Longer spellings come first for matching.
export const FONT_WEIGHT_WORDS = Object.freeze([
  ['extra light', 200], ['extralight', 200], ['ultra light', 200], ['ultralight', 200],
  ['semi light', 350], ['semilight', 350], ['demi light', 350], ['demilight', 350],
  ['semi bold', 600], ['semibold', 600], ['demi bold', 600], ['demibold', 600],
  ['extra bold', 800], ['extrabold', 800], ['ultra bold', 800], ['ultrabold', 800],
  ['extra black', 950], ['extrablack', 950], ['ultra black', 950], ['ultrablack', 950],
  ['thin', 100], ['hairline', 100], ['light', 300], ['book', 400], ['regular', 400], ['normal', 400],
  ['medium', 500], ['bold', 700], ['black', 900], ['heavy', 900],
]);
export const FONT_POSTURE_WORDS = Object.freeze(['italic', 'oblique']);

const wordPattern = [...FONT_WEIGHT_WORDS.map(([word]) => word), ...FONT_POSTURE_WORDS]
  .map(word => word.replace(' ', ' ?')).join('|');
// Whole suffix made only of style words, e.g. "SemiBold", "Extra Bold Italic".
export const FACE_STYLE_WORDS = new RegExp(`^(?:(?:${wordPattern})[ \\t]*)+$`, 'i');

const weights = new Map(FONT_WEIGHT_WORDS.map(([word, weight]) => [word.replace(' ', ''), weight]));
const tokens = new RegExp(`(${wordPattern})`, 'gi');

// Faces the OPF bundles under legacy weight names. A provider may also supply
// other families; those are recognized through the deck's own chosen family.
export const KNOWN_WEIGHT_FACE_FAMILIES = Object.freeze(['Roboto']);
// Genuine standalone families whose names end in a weight word. They are never
// split, even when their base family is the chosen one.
export const STANDALONE_WEIGHT_NAMED_FAMILIES = Object.freeze([
  'Arial Black', 'Calibri Light', 'Segoe UI Black', 'Segoe UI Light', 'Segoe UI Semibold', 'Segoe UI Semilight',
  'Franklin Gothic Medium', 'Franklin Gothic Heavy', 'Franklin Gothic Demi', 'Gill Sans Ultra Bold',
  'Helvetica Neue Light', 'Helvetica Neue Medium', 'Bahnschrift SemiBold', 'Rockwell Extra Bold',
].map(name => name.toLowerCase()));

// Split "<family> <weight words>" into the family and its CSS weight.
// `families` are the chosen or known base families that may own weight faces.
// Returns undefined for any other name, so genuine families are never altered.
export function splitWeightFace(typeface, families) {
  if (typeof typeface !== 'string') return undefined;
  const name = typeface.trim().replace(/\s+/g, ' ');
  if (STANDALONE_WEIGHT_NAMED_FAMILIES.includes(name.toLowerCase())) return undefined;
  for (const family of families) {
    if (typeof family !== 'string' || !family.trim()) continue;
    const base = family.trim().replace(/\s+/g, ' ');
    if (name.length <= base.length + 1 || !name.toLowerCase().startsWith(`${base.toLowerCase()} `)) continue;
    const suffix = name.slice(base.length + 1);
    if (!FACE_STYLE_WORDS.test(suffix)) continue;
    let weight, italic = false;
    for (const [word] of suffix.matchAll(tokens)) {
      const key = word.toLowerCase().replace(' ', '');
      if (FONT_POSTURE_WORDS.includes(key)) italic = true;
      else weight = weights.get(key);
    }
    return {family: base, weight: weight ?? 400, italic};
  }
  return undefined;
}
