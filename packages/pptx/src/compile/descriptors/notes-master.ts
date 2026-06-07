/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import { buildNotesMasterXml } from "@file/notes-master";
import type { NotesMasterOptions } from "@file/notes-master";
import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface NotesMasterDescriptorOptions {
  options?: NotesMasterOptions;
}

// ── Descriptor ──

export const notesMasterDesc: CustomDescriptor<NotesMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildNotesMasterXml(opts.options);
  },

  parse(_el, _ctx) {
    return {} as Partial<NotesMasterDescriptorOptions>;
  },
};
