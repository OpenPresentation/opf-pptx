# Maintained Keynote qualification (RR-86)

This is a versioned, project-authored baseline pack for the existing PPTX exporter and importer. It generates 16 OPF slides, a PPTX, SVG previews, edit actions and a sealed identity manifest. It does not change export defaults, implement native paragraph/list-body export, produce `.key` files, automate Keynote or establish native acceptance. The opt-in body prototype follows review of this baseline. Direct `.key` generation and advanced effects remain deferred.

Use Node 24 and a genuine installed consumer containing matching `@openpresentation/opf`, `opf-render`, `opf-pptx` and `gallery` packages plus the renderer's pinned font peers. The generator records actual installed versions and runtime hashes rather than assuming a release number; generate the accepted baseline again against the final coordinated registry graph. Source, installed-package, registry provenance and native claims remain separate.

```sh
node scripts/keynote-qualification/cli.mjs generate --consumer /absolute/installed-consumer --output /absolute/new-keynote-suite
node scripts/keynote-qualification/cli.mjs verify --suite /absolute/new-keynote-suite
```

The output directory must be new, outside the repository and installed consumer, under an existing canonical parent. Existing attempts are never overwritten. `suite.json` seals every generated source, archive, preview, action and static inspection file. It binds the generator revision/file hashes, installed runtime files, prepared font-file hashes and renderer manifest hash. The one prepared font handle is shared by preview and export. It neither installs native fonts nor asserts which physical face Keynote uses. The renderer verifies its bundled font/license pins normally.

The project-authored deck selects **Arial for both heading and body**, using the normal licensed **Arimo preview alias** in regular, bold, italic and bold italic styles. Native XML must retain Arial, without writing the preview alias as the selected native family. Proprietary Arial files are never bundled. Root's 2026-10-15 preflight reports four Arial style files available on the Mac and the Windows preflight verifies four styles. The reported regular-file hashes differ: macOS `525979822591a3447cfc49d943d6f7683508e25543407871c0ed8fed05fd2bd9`, Windows `b3658eadae55e682b5f69eb64c439c1ecc8f196c0bb8d4756d145d13bc86476a`. These attributed file observations do not prove identical font versions, physical per-glyph selection or native parity. Record actual native font/substitution observations in the UI results. The earlier Roboto suite is retained unchanged as dated smoke evidence and is not the canonical Arial baseline.

| Slides | Source case | Native observation or edit |
| --- | --- | --- |
| K01 | Cover/title/subtitle | Title, Body and Slide Number controls; imported layout chooser; create one native slide |
| K02–K03 | Scalar and rich paragraphs | Replace a marker, Return/new paragraph, bold/italic and object boundaries |
| K04 | List, description and nested child | Replace item, Return, indent, descriptions and list controls |
| K05–K06 | Long prose and wrapped list | Append paragraph/item, wrapping, clipping and editable-object boundaries |
| K07 | Two-column text and metric | Change layout, reapply original, preserve content and inspect geometry |
| K08–K09 | Synthetic PNG, contain and cropped cover | Replace image and use native crop/mask with caption preserved |
| K10–K11 | Native column and line charts | Change Q2 to 900 in native data editor; change line color |
| K12–K13 | Plain and rich native tables | Edit text cell and bold run; preserve table structure |
| K14–K15 | Slide-number fields | Reorder K14 after K15 and inspect actual displayed numbers |
| K16 | Notes and untouched final control | Append a notes marker and preserve the slide text |

Pictures reuse the repository's MIT synthetic quadrants/circle fixtures. `assets/replacement.png` is the replacement witness. The original retained long-text deck remains an independent review case in `test/fixtures/native-long-text/`; this pack does not replace or modify it.

## Native supervisor procedure

Only the root/supervisor operates Keynote. Each route starts from a separate fresh copy of the sealed baseline:

```sh
node scripts/keynote-qualification/cli.mjs copy --suite /absolute/new-keynote-suite --route no-op --output /absolute/new-keynote-noop
node scripts/keynote-qualification/cli.mjs copy --suite /absolute/new-keynote-suite --route content-edits --output /absolute/new-keynote-edits
node scripts/keynote-qualification/cli.mjs copy --suite /absolute/new-keynote-suite --route layout-structure --output /absolute/new-keynote-layouts
```

Open only that attempt's `editable.pptx`. Record Keynote/macOS versions, import warnings, actual font names and screenshots in its `result.json`. Follow the route actions in `attempt.json`; unavailable controls are `unsupported`, failed observations are `failed`, and unexecuted actions stay `not-run`. Use `passed` only with identified evidence. For each route, save a **new** `working.key`, export a **new** `native-export.pptx`, reopen both and retain any failure. Never save over sealed inputs, unrelated user files or an earlier attempt. Close only the test copies. Native slide-number display, crop controls, chart/table editability, text fit and actual fonts require real UI observations.

Keep four directional outcomes distinct in `result.json`: `opf-to-pptx-to-keynote`, `keynote-save-and-reopen`, `keynote-to-pptx`, and `keynote-pptx-to-opf`. A valid exported ZIP, recovered source provenance or successful import alone cannot pass a native edit action. In particular, line shapes are not proof of native paragraph or list semantics, and cached field text is not proof that slide numbers update.

Inspect the actual native-exported file without opening any app:

```sh
node scripts/keynote-qualification/cli.mjs inspect --suite /absolute/new-keynote-suite --consumer /absolute/installed-consumer --input /absolute/new-keynote-edits/native-export.pptx --route content-edits --output /absolute/new-keynote-edits-inspection
```

Repeat with the corresponding route for each output. This writes new `inspection.json`, `imported.opf.json` and import diagnostics. It validates XML, package relationship targets, native chart/table presence, slide ordering, marker text, notes and edited chart cache. It retains actual shape/placeholder/paragraph inventories and picture/crop hashes for independent review. Assertions are static and do not establish native geometry, glyph/font, baseline or pixel parity. The output always keeps `nativeAcceptance: false`; the supervisor reviews this inspection together with actual UI evidence and the directional result file.

## Portable checks

```sh
node test/keynote-qualification.mjs
node --check scripts/keynote-qualification/cli.mjs
node --check scripts/keynote-qualification/pack.mjs
```

The tests cover the action/source contract, placeholder and relationship inspection, presentation order independent of part numbering, fresh-copy protection and sealed-file tamper rejection. They use no native application, network or machine fonts. Normal test discovery includes the test automatically. Generated decks and native outputs stay outside tracked source; the fixture generator and actions are maintained here, with no transfer ZIP as a distribution requirement.
