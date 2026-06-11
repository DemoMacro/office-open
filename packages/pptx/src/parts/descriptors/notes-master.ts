/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { buildNotesMasterXml } from "@parts/notes-master";
import type { NotesMasterOptions } from "@parts/notes-master";

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
