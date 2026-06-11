/**
 * Notes Slide (p:notes) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { buildNotesSlideXml } from "@parts/notes-slide";

// ── Types ──

export interface NotesSlideDescriptorOptions {
  /** Notes text content. */
  text?: string;
}

// ── Descriptor ──

export const notesSlideDesc: CustomDescriptor<NotesSlideDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildNotesSlideXml({ text: opts.text });
  },

  parse(_el, _ctx) {
    return {} as Partial<NotesSlideDescriptorOptions>;
  },
};
