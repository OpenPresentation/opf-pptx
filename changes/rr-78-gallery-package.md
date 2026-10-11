---
type: changed
---
RR-78 (OPF 0.19, docs and tests only): the pptx.gallery catalog is now the package `@openpresentation/gallery`, because core 0.19 removed `@openpresentation/opf/catalog` with no alias. Register it with `import { gallery } from "@openpresentation/gallery"` and `catalogs: [gallery]`. Its display metadata is `catalogDisplay` from the same package. This package still bundles and fetches no catalog: the README, the declarations and the tests now name the gallery package, and it is a devDependency for the tests only.
