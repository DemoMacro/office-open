/**
 * Placeholder inheritance resolution.
 *
 * A slide placeholder omits geometry (xfrm) when it inherits from its layout,
 * which in turn inherits from the slide master. {@link resolvePlaceholder}
 * walks that chain to recover the rendered position and visibility.
 *
 * @module
 */

import type { MasterPlaceholderPosition } from "@parts/slide-master";

import type { LayoutDefinition, MasterDefinition } from "./file";

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

/** Result of resolving a placeholder against the layout/master chain. */
export interface ResolvedPlaceholder {
  /**
   * First defined position (layout takes precedence over master). Undefined
   * when neither carries a position for this placeholder type.
   */
  position?: MasterPlaceholderPosition;
  /** True when layout or master carries `sz="0"` (hidden). */
  hidden?: boolean;
}

function asPosition(ph: unknown): MasterPlaceholderPosition | undefined {
  return ph !== undefined && ph !== false && ph !== true && typeof ph === "object"
    ? (ph as MasterPlaceholderPosition)
    : undefined;
}

/**
 * Resolve a slide placeholder's inherited position and visibility from the
 * layout → master chain.
 *
 * @param placeholderType The `p:ph/@type` token from the slide placeholder
 *   (e.g. `"title"`, `"body"`, `"dt"`).
 * @param layout The slide's layout definition (carries placeholder positions).
 * @param master The layout's slide master definition.
 * @returns The resolved position (or `{ hidden: true }` when either layer
 *   suppresses the placeholder with `sz="0"`).
 *
 * spPr (fill/outline/effects), textBody, and shape-style inheritance require
 * the layout/master placeholder to carry those properties beyond position;
 * extending the placeholder data model to a full `PlaceholderInheritance` is
 * the follow-up that unlocks complete inheritance resolution.
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

  const position = asPosition(layoutPh) ?? asPosition(masterPh);
  return position ? { position } : {};
}
