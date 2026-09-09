# Prepared PPTX 0.6.0

Branch `codex/shared-quote-release-20260909` prepares PPTX 0.6.0 with core `^0.8.0` and optional renderer peer `^0.6.0` (development fixture version 0.6.0). It is unpublished. The original tested integration remains at `829103b67fdc96ae977bbb05f3e94457620645cd` on `codex/shared-quote-integration-20260909`. Core/CLI release PR [#49](https://github.com/OpenPresentation/opf/pull/49) and renderer 0.6.0 publication are predecessor gates.

The lockfile intentionally remains on the previous published set. Refresh it only after both predecessors are available from npm, then run clean Node 20/24 installations, full package/corpus/packed/browser checks, and renewed Windows PowerPoint fixtures. The prior twelve-case native report remains bound to its original source/font hashes; do not relabel it as verification of a different packaged runtime. Confirm editable lines, glyph containment/separation, save/reopen, reimport and raster differences separately.

Linux CI/publication use the exact Playwright 1.63.0 image already verified by coordinated CI `34399051732`, pin its digest and check the installed test version. The Windows matrix remains on native hosted Windows and installs Chromium separately. Open the final PR only after registry-dependent local gates can run; require clean CI/review before merge and a fresh `opf-pptx-v0.6.0` publication tag.

Native quote import retains editable text but loses OPF quote structure, typography and readability policy. Native chart geometry and general scalar wrapping remain separate fidelity gaps. There is no claim of arbitrary round-trip or pixel equivalence.
