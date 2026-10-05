---
type: fixed
---
FF-05 (opf-pptx#168): a slide whose own font scheme selects East Asian / complex-script fonts other than the first slide's is no longer exported silently with the first slide's script fonts; the export reports `script-font-per-slide-not-exported` (path `slides.N.design.fontScheme`).
