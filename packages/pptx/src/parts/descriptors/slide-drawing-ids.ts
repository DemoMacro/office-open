/**
 * Slide drawing-object id allocation.
 *
 * @module
 */

import type { WriteContext } from "@office-open/core/descriptor";

/**
 * Next drawing-object id for the slide being stringified, honoring a
 * reproducible scope when one is active: the scope counter is offset by one
 * because the slide's group shape (p:grpSpTree) always owns id 1. Returns
 * undefined without a scope so callers fall back to their module counters.
 */
export const nextSlideDrawingId = (ctx: WriteContext): number | undefined =>
  ctx.reproducible ? ctx.reproducible.nextDrawingId() + 1 : undefined;
