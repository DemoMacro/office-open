/**
 * Placeholder inheritance resolution.
 *
 * A slide placeholder omits geometry (xfrm) and other shape facets when it
 * inherits them from its layout, which in turn inherits from the slide master.
 * {@link resolvePlaceholder} walks that chain to recover the rendered position,
 * shape facets, text body and style.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { shapePropertiesDesc, textBodyDesc } from "@office-open/core/drawing";
import type { ShapePropertiesOptions, TextBodyOptions } from "@office-open/core/drawing";
import { attr, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { LayoutDefinition, MasterDefinition } from "./file";
import { readShapeStyle, type ShapeStyleOptions } from "./shape/shape";

/**
 * Maps a `p:ph/@type` token to the matching key on
 * {@link LayoutDefinition.placeholders} / {@link MasterDefinition.placeholders}.
 * `ctrTitle` (centered title) is normalized to `title`.
 */
export const PLACEHOLDER_TYPE_TO_KEY: Readonly<Record<string, string>> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  sub: "subtitle",
  dt: "date",
  ftr: "footer",
  sldNum: "slideNumber",
};

/** Position rect (a:xfrm off/ext). `number` is EMU on round-trip. */
export interface PlaceholderPosition {
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  width: number | UniversalMeasure;
  height: number | UniversalMeasure;
}

/**
 * Inheritable facets of a placeholder beyond position: shape properties
 * (geometry/fill/outline/effects/3D), text body, and shape style. Each field is
 * optional — present only when the layout/master placeholder defined it. A
 * slide placeholder inherits whichever facets it omits.
 */
export interface PlaceholderFacets {
  geometry?: ShapePropertiesOptions["geometry"];
  customGeometry?: ShapePropertiesOptions["customGeometry"];
  fill?: ShapePropertiesOptions["fill"];
  outline?: ShapePropertiesOptions["outline"];
  effects?: ShapePropertiesOptions["effects"];
  scene3d?: ShapePropertiesOptions["scene3d"];
  shape3d?: ShapePropertiesOptions["shape3d"];
  textBody?: TextBodyOptions;
  style?: ShapeStyleOptions;
}

/**
 * A complete placeholder template on a master/layout: position plus any
 * inheritable facets the source defined. Backward-compatible with the old
 * position-only input shape (flat x/y/width/height).
 */
export interface PlaceholderDefinition extends PlaceholderPosition, PlaceholderFacets {}

/** Result of resolving a placeholder against the layout/master chain. */
export interface ResolvedPlaceholder {
  /**
   * First defined position (layout takes precedence over master). Undefined
   * when neither carries a transform for this placeholder type.
   */
  position?: PlaceholderPosition;
  /**
   * Inherited shape/text/style facets. Master is the base; layout overrides
   * per-facet. Undefined when neither layer defines any facet.
   */
  facets?: PlaceholderFacets;
  /** True when layout or master carries `sz="0"` (hidden). */
  hidden?: boolean;
}

function asDefinition(ph: unknown): PlaceholderDefinition | undefined {
  return ph !== undefined && ph !== false && ph !== true && typeof ph === "object"
    ? (ph as PlaceholderDefinition)
    : undefined;
}

function pickPosition(def: PlaceholderDefinition | undefined): PlaceholderPosition | undefined {
  if (!def) return undefined;
  const { x, y, width, height } = def;
  if (x === undefined || y === undefined || width === undefined || height === undefined) {
    return undefined;
  }
  return { x, y, width, height };
}

function pickFacets(def: PlaceholderDefinition | undefined): PlaceholderFacets | undefined {
  if (!def) return undefined;
  const facets: PlaceholderFacets = {};
  if (def.geometry !== undefined) facets.geometry = def.geometry;
  if (def.customGeometry !== undefined) facets.customGeometry = def.customGeometry;
  if (def.fill !== undefined) facets.fill = def.fill;
  if (def.outline !== undefined) facets.outline = def.outline;
  if (def.effects !== undefined) facets.effects = def.effects;
  if (def.scene3d !== undefined) facets.scene3d = def.scene3d;
  if (def.shape3d !== undefined) facets.shape3d = def.shape3d;
  if (def.textBody !== undefined) facets.textBody = def.textBody;
  if (def.style !== undefined) facets.style = def.style;
  return Object.keys(facets).length > 0 ? facets : undefined;
}

function mergeFacets(
  base: PlaceholderFacets | undefined,
  override: PlaceholderFacets | undefined,
): PlaceholderFacets | undefined {
  if (!base && !override) return undefined;
  if (!base) return override;
  if (!override) return base;
  return { ...base, ...override };
}

/**
 * Resolve a slide placeholder's inherited position and facets from the
 * layout → master chain.
 *
 * @param placeholderType The `p:ph/@type` token from the slide placeholder
 *   (e.g. `"title"`, `"body"`, `"dt"`).
 * @param layout The slide's layout definition (carries placeholder templates).
 * @param master The layout's slide master definition.
 * @returns The resolved position and facets (or `{ hidden: true }` when either
 *   layer suppresses the placeholder with `sz="0"`).
 */
export function resolvePlaceholder(
  placeholderType: string,
  layout: LayoutDefinition | undefined,
  master: MasterDefinition | undefined,
): ResolvedPlaceholder {
  const key = PLACEHOLDER_TYPE_TO_KEY[placeholderType];
  if (!key) return {};

  const layoutPh = layout?.placeholders?.[key as keyof typeof layout.placeholders];
  const masterPh = master?.placeholders?.[key as keyof typeof master.placeholders];

  // sz="0" on either layer hides the placeholder downstream.
  if (layoutPh === false || masterPh === false) return { hidden: true };

  const layoutDef = asDefinition(layoutPh);
  const masterDef = asDefinition(masterPh);

  const result: ResolvedPlaceholder = {};

  const position = pickPosition(layoutDef) ?? pickPosition(masterDef);
  if (position) result.position = position;

  // Master is the base; layout overrides each facet it defines.
  const facets = mergeFacets(pickFacets(masterDef), pickFacets(layoutDef));
  if (facets) result.facets = facets;

  return result;
}

// ── Placeholder extraction (master/layout parse path) ──

/**
 * Read a placeholder's logical key and full definition (position + inheritable
 * facets), or `false` when hidden (`sz="0"`). Returns undefined for non-placeholder
 * shapes or ones without a transform. Shared by the slide-master and slide-layout
 * parse paths; the caller supplies its own `p:ph/@type` → key map.
 *
 * Facets beyond position are carried only when the source defined them: non-rect
 * geometry, fill/outline/effects/3D, custom text body (a run/break beyond the
 * default endParaRPr/field body), and shape style. This keeps default placeholders
 * position-only so the fresh emit path stays byte-equivalent with MS Office.
 */
export function extractPlaceholderDefinition(
  spEl: XmlElement,
  ctx: ReadContext,
  phTypeToKey: Readonly<Record<string, string>>,
): { key: string; def: PlaceholderDefinition | false } | undefined {
  const nvSpPr = findChild(spEl, "p:nvSpPr");
  const nvPr = nvSpPr ? findChild(nvSpPr, "p:nvPr") : undefined;
  const ph = nvPr ? findChild(nvPr, "p:ph") : undefined;
  if (!ph) return undefined;

  const phType = attr(ph, "type");
  const key = phType ? phTypeToKey[phType] : undefined;
  if (!key) return undefined;

  if (attr(ph, "sz") === "0") return { key, def: false };

  const def: Partial<PlaceholderDefinition> = {};

  // Position + spPr facets from shapePropertiesDesc.parse.
  const spPr = findChild(spEl, "p:spPr");
  if (spPr) {
    const spPrOpts = shapePropertiesDesc.parse(spPr, ctx);
    if (spPrOpts) {
      if (spPrOpts.x !== undefined) def.x = spPrOpts.x;
      if (spPrOpts.y !== undefined) def.y = spPrOpts.y;
      if (spPrOpts.width !== undefined) def.width = spPrOpts.width;
      if (spPrOpts.height !== undefined) def.height = spPrOpts.height;
      // rect is the placeholder geometry default — carry only non-rect.
      if (spPrOpts.customGeometry) {
        def.customGeometry = spPrOpts.customGeometry;
      } else if (spPrOpts.geometry !== undefined) {
        const preset =
          typeof spPrOpts.geometry === "string" ? spPrOpts.geometry : spPrOpts.geometry.preset;
        if (preset !== "rect") def.geometry = spPrOpts.geometry;
      }
      if (spPrOpts.fill !== undefined) def.fill = spPrOpts.fill;
      if (spPrOpts.outline !== undefined) def.outline = spPrOpts.outline;
      if (spPrOpts.effects !== undefined) def.effects = spPrOpts.effects;
      if (spPrOpts.scene3d !== undefined) def.scene3d = spPrOpts.scene3d;
      if (spPrOpts.shape3d !== undefined) def.shape3d = spPrOpts.shape3d;
    }
  }

  // textBody: carry only when it holds custom text (a run/break beyond the
  // default endParaRPr/field body), so default placeholders stay byte-equivalent.
  const txBody = findChild(spEl, "p:txBody");
  if (txBody && hasTextContent(txBody)) {
    def.textBody = textBodyDesc.parse(txBody, ctx);
  }

  // p:style — placeholders default to none; carry whenever present.
  const styleEl = findChild(spEl, "p:style");
  if (styleEl) def.style = readShapeStyle(styleEl);

  // A placeholder definition requires a concrete position.
  if (
    def.x === undefined ||
    def.y === undefined ||
    def.width === undefined ||
    def.height === undefined
  ) {
    return undefined;
  }
  return { key, def: def as PlaceholderDefinition };
}

/** True when a txBody holds an a:r/a:br run (custom content, not the empty default). */
function hasTextContent(txBody: XmlElement): boolean {
  for (const p of txBody.elements ?? []) {
    if (p.name !== "a:p") continue;
    for (const child of p.elements ?? []) {
      if (child.name === "a:r" || child.name === "a:br") return true;
    }
  }
  return false;
}
