/**
 * Notes Slide (p:notes) descriptor for PPTX.
 *
 * @module
 */

import { buildNotesSlideXml } from "@file/notes-slide";
import type { CustomDescriptor } from "@office-open/core/descriptor";

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
