import type { LayoutDiagnostic, TextMeasurement } from "@openpresentation/opf/composition";
export declare const packageName = "@openpresentation/opf-pptx";

export declare const releaseLane: Readonly<{
  githubRepository: "OpenPresentation/opf-pptx";
  npmPackage: "@openpresentation/opf-pptx";
  compatibilityPackage: "@openpresentation/opf";
  rendererPackage: "@openpresentation/opf-render";
}>;

export declare const runtimePolicy: Readonly<{
  hostedServiceInCriticalPath: false;
  telemetry: false;
  commercialSdkInCriticalPath: false;
  requiredAiDependency: false;
  requiredLibreOfficeDependency: false;
  requiredNetworkCalls: false;
  deterministicLocalExecution: true;
}>;

export type ImageResolverResult =
  | string
  | Uint8Array
  | {
      data?: string | Uint8Array;
      path?: string;
      mediaType?: string;
    };

export interface ImageResolverContext {
  asset: unknown;
  presentation: unknown;
  path: string;
}

export interface FontSchemeDiagnostic { code: "unresolved-font-scheme"; path: string; message: string; id: string; fallback: string }
export interface MediaProvenanceDiagnostic { code: "media-provenance-omitted"; path: string; message: string }
/** The chart data cannot be plotted, so a placeholder frame stands in for the chart. */
export interface ChartDataUnplottableDiagnostic { code: "chart-data-unplottable"; path: string; message: string; reason: "data-not-inline" | "no-rows" | "no-columns" | "single-column-not-numeric" }
/** Content with no PowerPoint export (an empty table, an unsupported payload) is replaced by a plain-language placeholder frame. */
export interface ContentPlaceholderDiagnostic { code: "content-placeholder"; path: string; message: string; reason: "table-has-no-rows" | "unsupported-payload" }
/**
 * An image that cannot be exported as a picture: the preview's placeholder was exported instead (or no watermark, or the background colour).
 * `reason: "unsupported-format"`: the embedded bytes are no readable PNG, JPEG, GIF, WebP or SVG. For an SVG (a native SVG picture over a PNG fallback):
 * `"svg-malformed"` (not well-formed XML, or no `xmlns` SVG root), `"svg-no-size"` (no width and height and no viewBox), `"svg-too-large"` (over 8 MiB),
 * `"svg-unsafe"` (a DOCTYPE with external or markup entities), `"svg-rasterizer-unavailable"` (no `options.svgRasterizer` and `@openpresentation/opf-render`
 * is not installed; always the case in a browser build without `svgRasterizer`) `"svg-render-failed"` (the rasterizer threw or returned no PNG) or `"svg-unreadable"` (a local `.svg` path that could not be read).
 * `strictAssets` throws `unsupported-image-dimensions` for an unreadable raster, `invalid-svg-image` for the SVG content reasons above (and `svg-unreadable`), `svg-rasterizer-unavailable` for a
 * missing rasterizer and `svg-render-failed` for a failed one instead.
 */
export interface UnresolvedAssetDiagnostic { code: "unresolved-asset"; path: string; message: string; reason?: "unsupported-format" | "svg-malformed" | "svg-no-size" | "svg-too-large" | "svg-unsafe" | "svg-rasterizer-unavailable" | "svg-render-failed" | "svg-unreadable" }
/** An SVG picture had scripts, `foreignObject`, event handlers, references outside the file, `@import` rules or a DOCTYPE; they were removed from the embedded SVG (on import too). Nothing in an SVG is run or fetched. */
export interface SvgSanitizedDiagnostic { code: "svg-sanitized"; path: string; message: string }
/** An SVG slide image with a duotone recolor or a non-rectangular shape exports as its PNG raster, not as a native SVG picture, so the effect applies as in the preview (PowerPoint applies opacity, grayscale and a border to an SVG picture, and those stay native). */
export interface SvgImageRasterizedDiagnostic { code: "svg-image-rasterized"; path: string; message: string }
export interface WatermarkNotExportedDiagnostic { code: "watermark-not-exported"; path: string; message: string }
/**
 * The chart data was reshaped to export a native chart: a single value column was plotted against row numbers; a one-series construct
 * (pie, doughnut, and with `chartex: "native"` treemap, histogram, pareto, waterfall, funnel, map) kept only its first series; by default
 * a chartex chart type was written as a clustered column chart (`chartex-fallback`) and a one-column histogram was binned into counts
 * (`histogram-binned`).
 */
export interface ChartDataAdaptedDiagnostic { code: "chart-data-adapted"; path: string; message: string; adaptation: "histogram-binned" | "row-numbers" | "series-dropped" | "chartex-fallback" }
/** With `chartex: "native"`, a map chart (`world`) is exported as a chartex regionMap (`Requires="cx4"`) without cached geography: PowerPoint must fetch the region shapes from its online map service and shows "There was a problem getting the information for your map chart" with an empty chart until it does. */
export interface ChartMapGeodataDiagnostic { code: "chart-map-geodata"; path: string; message: string }

export interface ToPptxOptions {
  /** Default compatible converts WebP to a static PNG. Preserve embeds original WebP bytes. */
  imageFormat?: "compatible" | "preserve";
  /**
   * How the chartex chart types (treemap, histogram, pareto, box-and-whisker, waterfall, funnel, world) are exported.
   * "auto" (default) writes Office 2016 chartex parts (`cx:chartSpace`, style parts, an `mc:AlternateContent` frame with the
   * clustered column chart as fallback) for the constructs desktop PowerPoint confirmed natively on 2026-09-30 (treemap,
   * histogram, pareto, box-and-whisker, waterfall, funnel) and keeps `world` on the clustered column chart with
   * `chart-data-adapted` (`chartex-fallback`) because the regionMap part, though accepted with `Requires="cx4"`, draws
   * nothing until PowerPoint fetches map data online. "native" writes every chartex part, the map included (reports
   * `chart-map-geodata`). "fallback" writes clustered columns only
   * (`chartex-fallback`, `histogram-binned`), as releases before FF-22b did. Import of chartex parts is always on.
   */
  chartex?: "auto" | "native" | "fallback";
  /**
   * OPF_DOCUMENT_V1 / OPF_SLIDE_V1 customer-data tags that let fromPptx restore
   * catalog references, layout ids, authoring metadata (`filename` and `extensions`
   * included), `design.logo`, the asset registry, slide `section`/`extensions` and
   * the slide content structure: nested groups, promoted regions, the root payload
   * form, block ids and extensions, group composition (docs/document-roundtrip.md).
   * Tags are not shown in PowerPoint's UI. Default "full"; "references-only"
   * stores catalog references without organization, speaker, free text, slide ids,
   * assets or content structure; false writes no tags. Slide `section` labels are
   * also written as PowerPoint's native section list whatever this option says.
   */
  provenance?: "full" | "references-only" | false;
  textMeasurement?: TextMeasurement;
  /** Match preview/pagination clearance around supplied vector text outlines; default 1. */
  textRasterPadding?: number;
  /** Layout diagnostics, `media-provenance-omitted` when video data cannot be stored, plus `unresolved-font-scheme` (once per reference path) when a font-scheme id matches no record and the default `aptos` scheme is used as the base. */
  onDiagnostic?: (diagnostic: LayoutDiagnostic | FontSchemeDiagnostic | MediaProvenanceDiagnostic | ChartDataUnplottableDiagnostic | ChartDataAdaptedDiagnostic | ChartMapGeodataDiagnostic | ContentPlaceholderDiagnostic | UnresolvedAssetDiagnostic | WatermarkNotExportedDiagnostic | SvgSanitizedDiagnostic | SvgImageRasterizedDiagnostic) => void;
  baseDir?: string;
  compressionLevel?: number;
  imageResolver?: (src: string, context: ImageResolverContext) => ImageResolverResult | Promise<ImageResolverResult | null | undefined> | null | undefined;
  /**
   * Draws the PNG fallback of an SVG picture (the raster older viewers show; PowerPoint 2016 and Microsoft 365 draw the SVG itself). It receives the
   * sanitized SVG text and the pixel size to draw (the SVG's own aspect, about 192 dpi of the displayed size) and returns PNG bytes with a transparent
   * background. It must be deterministic for the export to be. Default in Node: opf-render's `svgToPng` (resvg with the bundled fonts), loaded only when an
   * SVG is exported (optional peer `@openpresentation/opf-render`). Without it, and without opf-render, an SVG picture exports as the "Image unavailable"
   * placeholder with an `unresolved-asset` diagnostic (`reason: "svg-rasterizer-unavailable"`); a browser build has no default.
   */
  svgRasterizer?: (svg: string, size: { width: number; height: number; scale: number; text: boolean }) => Uint8Array | Promise<Uint8Array>;
  seed?: number;
  strictAssets?: boolean;
  timestamp?: string;
  /**
   * ZIP calendar timestamps, including embedded workbooks, use UTC fields at
   * two-second resolution. Omitted/undefined retains the fixed 1980 default.
   * Accepts a valid Date, finite epoch milliseconds, YYYY-MM-DD (UTC midnight),
   * or YYYY-MM-DDTHH:mm[:ss[.fraction]] with Z or ±HH:mm. UTC years 1980–2099 only.
   * Ambiguous/legacy strings and invalid/out-of-range values throw OPFPptxError
   * at options.zipDate; null, empty string and zero no longer choose the default.
   */
  zipDate?: string | number | Date;
  /**
   * Today's calendar date (ISO YYYY-MM-DD) for `date: true` header/footer fields. The exporter never
   * reads a clock: it lays out this date as the cached text of a native PowerPoint date field, which
   * PowerPoint updates on open. Without it a current date is reported as unresolved content.
   */
  date?: string;
  /** Host-supplied catalog records, as in opf-render. Currently consulted for socialPlatforms (generated socials furniture). */
  catalogs?: Record<string, { records?: unknown[] } | unknown[]>;
  /** Records for document `catalogs.<kind>.source` URLs (a single source or each entry of an ordered search path), as in opf-render. Currently consulted for socialPlatforms. */
  catalogSources?: Record<string, { records?: unknown[] } | unknown[]>;
}

export interface FromPptxOptions {
  /** Reports native details that import cannot preserve, including code provenance fallback/reflow and grouped text transforms. Table paths identify native frame and row/cell indexes (including headers). Stored catalog references that no longer match the package report `design-reference-changed` / `layout-reference-changed` at the reference path; a slide layout id that resolves to no inline or bundled record reports `unresolved-layout-reference`; a slide whose imported blocks no longer fit the stored content structure reports `content-structure-changed` at `slides.N`, a block id repeated by a duplicated slide `duplicate-block-id`, and a footer section text that disagrees with PowerPoint's section list `section-reference-changed` (docs/document-roundtrip.md). */
  onDiagnostic?: (diagnostic: {code: string; path: string; message: string}) => void;
  fallbackName?: string;
  schema?: string;
  /** The export's host catalog records (as in ToPptxOptions). Currently used to recognize unedited socials lines, so authored handles return. */
  catalogs?: Record<string, { records?: unknown[] } | unknown[]>;
  /** The export's records for document `catalogs.<kind>.source` URLs (as in ToPptxOptions). Currently used for socialPlatforms. */
  catalogSources?: Record<string, { records?: unknown[] } | unknown[]>;
}

export declare class OPFPptxError extends Error {
  readonly code: string;
  readonly details: Record<string, unknown>;
  readonly issues?: unknown[];
  readonly path?: string;
  constructor(code: string, message: string, details?: Record<string, unknown>);
}

export declare function toPptx(input: unknown, options?: ToPptxOptions): Promise<Uint8Array>;

export declare function fromPptx(input: Uint8Array | ArrayBuffer, options?: FromPptxOptions): Promise<Record<string, unknown>>;

export interface TypefaceEntry {
  /** Part path; parts of nested packages use "outer.xlsx!/inner/part.xml". */
  part: string;
  kind: "drawingml" | "spreadsheetml";
  /** Local element name, for example latin, ea, cs, sym, buFont, font, name or rFont. */
  element: string;
  typeface: string;
  /** Theme font collection for theme parts. */
  theme?: "major" | "minor";
  /** Script tag of a theme script supplement (`<a:font script="…">`). */
  script?: string;
  pitchFamily?: number;
  /** Theme reference (+mj-lt, +mn-ea, …) resolved against the package theme; null when unresolved. */
  resolved?: string | null;
}

export interface TypefaceInventory {
  typefaces: TypefaceEntry[];
  themes: Record<string, {major: Partial<Record<"latin" | "ea" | "cs", string>>; minor: Partial<Record<"latin" | "ea" | "cs", string>>}>;
  /** docProps/app.xml "Fonts Used", or null when the package has no readable list. */
  fontsUsed: string[] | null;
}

export interface CheckPptxTypefacesOptions {
  /** Family names the document chose (heading, body, code, run fonts). Required. */
  fonts: string[];
  /** Chosen families that are monospace; their pitchFamily must be fixed pitch, and only theirs. */
  monospace?: string[];
  /** Allow empty theme ea/cs slots and references to them (FF-05). Default true. */
  allowEmptyThemeScripts?: boolean;
  /** Allowed theme script supplements; defaults to THEME_SCRIPT_SUPPLEMENTS. */
  scriptSupplements?: Readonly<Record<"major" | "minor", Readonly<Record<string, string>>>>;
}

export type TypefaceViolationReason =
  | "foreign-typeface"
  | "foreign-theme-reference"
  | "unresolved-theme-reference"
  | "empty-theme-reference"
  | "empty-typeface"
  | "foreign-script-supplement"
  | "monospace-not-fixed-pitch"
  | "fixed-pitch-not-monospace"
  | "inconsistent-pitch-family"
  | "missing-fonts-used"
  | "foreign-fonts-used"
  | "fonts-used-mismatch";

export interface TypefaceViolation {
  reason: TypefaceViolationReason;
  part: string;
  element?: string;
  typeface?: string;
  [detail: string]: unknown;
}

export declare const THEME_SCRIPT_SUPPLEMENTS: Readonly<Record<"major" | "minor", Readonly<Record<string, string>>>>;

export declare function inventoryPptxTypefaces(input: Uint8Array | ArrayBuffer | Record<string, Uint8Array>, options?: {nested?: boolean}): TypefaceInventory;

export declare function packageFontsUsed(inventory: TypefaceInventory): string[];

export declare function checkPptxTypefaces(input: Uint8Array | ArrayBuffer | Record<string, Uint8Array>, options: CheckPptxTypefacesOptions): {
  ok: boolean;
  violations: TypefaceViolation[];
  inventory: TypefaceInventory;
  /** The "Fonts Used" list the package's own fonts imply. */
  fontsUsed: string[];
};
