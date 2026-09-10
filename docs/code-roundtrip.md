# Code source round trips (unreleased)

This feature needs the coordinated shared-code branches of core, renderer and converter. It is not available in the already published converter 0.6.0. No hosted service, font download or proprietary application is required at runtime.

Export writes one editable native text shape per accepted display line, including blank lines, and consumes the same filename/language/body fits as SVG. Literal tabs use accepted explicit native tab stops. Source text is never trimmed, uppercased or independently re-fitted during export.

Standard [PresentationML customer-data tags](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.customerdatatags?view=openxml-3.0.1) associate the code panel and line shapes. `OPF_CODE_V1` contains UTF-8 JSON encoded as uppercase hex because PowerPoint's [Tags API](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags) treats tags case-insensitively. The panel records the original code value and validated contiguous UTF-16 line ranges; line tags identify the group, part and display-line index. Shape names are only human-readable labels. This is source mapping, not a signature or authenticity claim; imported metadata is untrusted data and is never executed.

Import reconstructs only complete, uniquely identified groups. Current native text wins over old source text. Original hard line separators survive; visual soft wraps add no newline. New native paragraph/soft-break edits use that source part's first original newline convention, or LF when none exists. Original empty filename/language fields remain empty. Code source, filename and language remain editable OPF fields.

Missing/duplicate shapes, damaged tags/ranges, multiple identities, an edited panel or an edited generated label prevent reconstruction. `invalid-code-provenance` reports fallback to visible native shapes without resurrecting the old source. Entirely stripped tags are ordinary native content and cannot restore code semantics. Grouped code text can be recovered, but native group transforms and other group objects remain unsupported and produce `grouped-text-reflow`.

Successful source recovery reports `code-import-reflow`: native positioning, font theme, formatting and OPF readability policy are not reconstructed. Keep the original PPTX and review the newly composed OPF. This is not lossless arbitrary PowerPoint round-tripping or a pixel-equivalence guarantee.

Run `npm run test:code` for exact source/metadata, whitespace, Unicode serialization, native edits, renamed/reordered/grouped shapes and damaged-group checks. Astral Unicode serialization is separate from font coverage; the bundled monospace pack lacks the emoji used by that serialization fixture. The editor's `test:code-browser` runs actual offline editing/export/reimport.

On Windows with PowerPoint and local Courier New installed:

```powershell
node test/native-code.mjs generate artifacts/native-code
& test/native-code.ps1 -EvidenceDirectory artifacts/native-code
node test/native-code.mjs compare artifacts/native-code
```

The helper fingerprints code, dependencies, reference fonts and input/output files; it opens only generated fixtures and leaves unrelated presentations and PowerPoint open. Eight wide/portrait slides cover empty/whitespace-only source, case-sensitive metadata, tabs, blank/final lines, mixed line endings and long wrapped content. Original/native-saved/native-edited imports are checked separately. Native glyph containment is measured against accepted outer cells; separate rasters remain observations. Reference font bytes are neither copied nor embedded.
