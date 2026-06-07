/**
 * Notes Master (p:notesMaster) descriptor for PPTX.
 *
 * @module
 */

import { DefaultNotesMaster } from "@file/notes-master";
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
    const master = new DefaultNotesMaster(opts.options);
    return master.serialize();
  },

  parse(_el, _ctx) {
    return {} as Partial<NotesMasterDescriptorOptions>;
  },
};
