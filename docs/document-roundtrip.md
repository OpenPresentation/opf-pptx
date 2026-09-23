# Document references and metadata round trips (FF-32)

PPTX has no native field for an OPF catalog reference such as `design.colorScheme: "boost"` or a slide `layout: "title-slide"`. It also has none for authoring metadata such as `narrative`, `tone`, `audience`, `purpose`, `language`, `organization` or `speaker`. Before FF-32, `fromPptx` rebuilt only what the native parts showed, so every reference and metadata value was dropped without a diagnostic. The pptx.gallery parity harness recorded that loss for all 900 gallery values.

## Storage: PresentationML customer-data tags

Export writes two kinds of standard [customer-data tags](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.customerdatatags?view=openxml-3.0.1). They use the same mechanism as the existing `OPF_CODE_V1`, `OPF_HEADING_V1` and `OPF_FURNITURE_V1` provenance tags.

| tag | location | content |
|---|---|---|
| `OPF_DOCUMENT_V1` | `p:presentation/p:custDataLst` → `ppt/tags/opfDocument.xml` | deck `design` references (`theme`, `colorScheme`, `fontScheme`, `dimensions`, `background`) and composition defaults (`titleAlignment`, `contentAlignment`, `contentBox`, `contentDirection`, `chartPrimary`, `imageFill`, `listBullet`); metadata (`narrative`, `tone`, `audience`, `purpose`, `language`, `organization`, `speaker`, `takeaway`, `duration`, `tags`, `variables`); `assets` referenced by that metadata; referenced inline `catalogs` records; native evidence |
| `OPF_SLIDE_V1` | `p:cSld/p:custDataLst` of each slide → `ppt/tags/opfSlideN.xml`, or the slide's existing furniture tag list | slide `id`, `beat`, `layout`, `type`, `composition`, and slide `design` references and composition hints; native evidence |

Values are UTF-8 JSON encoded as uppercase hex, as for the other OPF tags, because PowerPoint's [Tags API](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags) treats tags case-insensitively. `custDataLst` is placed in schema order: after `notesSz` and before `defaultTextStyle` in `presentation.xml`, and after `spTree` in a slide. PresentationML allows one `p:tags` per `custDataLst`. A slide that already has an `OPF_FURNITURE_V1` tag list therefore gets its `OPF_SLIDE_V1` tag in that same list, and furniture import ignores it when it checks for ambiguous identities.

Alternatives considered:

- **Custom document properties (`docProps/custom.xml`)** are visible and editable in File > Info > Properties. Office limits text property values to 255 characters, which is too small for inline catalog records. They are also deck-wide only, so they cannot follow a slide that is copied or reordered.
- **A `customXml` data-store part** needs its own item-properties part, a GUID and a schema namespace, and it is also deck-wide only. It would be a second provenance mechanism next to the tags the package already uses.
- **Tags** are invisible to users, and PowerPoint preserves them on save. It copies a slide's tags with the slide. Every other OPF provenance record already uses them, so import has one reader and one threat model.

Only values the document states are stored; engine defaults are not. A document that states no reference, metadata or slide identity gets no tags and its bytes are unchanged. A metadata field larger than 256 KiB, or referenced inline catalogs larger than 1 MiB, are not stored and `toPptx` reports `document-provenance-omitted` with the field path.

## Native evidence and restore rules

Each stored reference records the native values it produced. Those values are read from the final normalized package, so later theme writers (for example FF-24 theme colors or FF-07 script fonts) are included automatically. Import restores a reference only while that evidence is unchanged:

| reference | restored when unchanged | on change |
|---|---|---|
| `design.colorScheme` | theme `clrScheme` slot colors (`sysClr` by system name; `lastClr` is ignored) | `design-reference-changed` at `design.colorScheme`, naming the changed slots |
| `design.fontScheme` | theme major/minor `latin`/`ea`/`cs` typefaces | `design-reference-changed` at `design.fontScheme` |
| `design.theme` | theme colors and theme fonts | `design-reference-changed` at `design.theme` |
| `design.dimensions` | `p:sldSz` | `design-reference-changed`; the observed inches stay |
| `design.background` | the background of the slides that inherited it: kept while at least one of them is unchanged, or when every slide overrides it | slides whose background changed keep a local override and report at `slides.N.design.background`; if all inheriting slides changed, the deck value is reported at `design.background` |
| slide `layout`, `type`, `composition`, slide composition hints | the slide's object kinds, positions, sizes, rotation and flips. Stacking order and table frame height (which viewers grow to fit text) are ignored | `layout-reference-changed` at `slides.N.layout`; `slide-reference-changed` / `design-reference-changed` for the others |
| slide `design.background` | that slide's `p:bg` (image backgrounds by target bytes, not relationship id) | `design-reference-changed` at `slides.N.design.background` |
| slide `design.theme`/`colorScheme`/`fontScheme` | the set of typefaces and colors used on the slide | `design-reference-changed` |
| deck composition defaults | every slide present and structurally unchanged | `design-reference-changed` at `design.<key>` |

Authoring metadata, slide `id` and slide `beat` have no native counterpart and are restored from the stored value. The organization name shown in current furniture text wins over the stored name. When a duplicated slide carries a copied tag, its layout is restored and the repeated slide id is reported as `duplicate-slide-id`. Inline catalog records return only for ids that the restored document still references. Referenced records therefore resolve and the reimported document exports again, while a layout id that is dropped does not bring its record back.

When a reference is not restored, the imported document keeps the values observed in the PPTX, for example the slide size in inches, the slide's own background or FF-24's recovered theme colors. A specific diagnostic always names the reference.

## Untrusted input

Tags can be written by anyone who can edit the file. The hash is a change detector, not a signature. Tag values are size-limited and decoded as JSON. Only known fields are accepted, and the whole candidate document must pass `validatePresentation`. Otherwise `invalid-document-provenance` is reported and the ordinary import is returned unchanged. Nothing in a tag is executed or fetched. A package whose tags were stripped imports as an ordinary foreign deck without a diagnostic.

## For layout recovery (FF-29)

`restoreDocumentProvenance()` in `src/document-provenance.js` returns, for each imported slide, `{layout, structure}`. `structure` is `'match'`, `'changed'` or `'untagged'`. Layout-structure recovery can use the stored id as a candidate when the arrangement changed.

## Scope

`npm test` runs `test/document-provenance.mjs`, which covers package shape, full round trip and re-export, benign editor rewrites, theme color, font, size, arrangement and background edits, duplicated slides, stripped, damaged and invalid tags, shared furniture tag lists and oversized metadata. These are XML-level checks. A native PowerPoint save of a tagged deck is a separate Office gate. `design.logo`, `design.watermark`, `design.slideImage` and `extensions` are not stored.
