/**
 * Notes Slide (p:notes) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, textOf } from "@office-open/xml";
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

  parse(el, _ctx) {
    // Find notes body placeholder and extract text
    const cSld = findChild(el, "p:cSld");
    if (!cSld) return {} as NotesSlideDescriptorOptions;
    const spTree = findChild(cSld, "p:spTree");
    if (!spTree) return {} as NotesSlideDescriptorOptions;

    for (const sp of spTree.elements ?? []) {
      if (sp.name !== "p:sp") continue;
      // Check for notes body placeholder: p:nvSpPr > p:nvPr > p:ph with type="body"
      const nvPr = findChild(findChild(sp, "p:nvSpPr") ?? sp, "p:nvPr");
      const ph = nvPr ? findChild(nvPr, "p:ph") : undefined;
      if (!ph) continue;
      const phType = ph.attributes?.["type"] as string | undefined;
      if (phType !== "body") continue;

      // Extract text from p:txBody
      const txBody = findChild(sp, "p:txBody");
      if (!txBody) break;
      const parts: string[] = [];
      for (const p of txBody.elements ?? []) {
        if (p.name !== "a:p") continue;
        for (const r of p.elements ?? []) {
          if (r.name !== "a:r") continue;
          const t = findChild(r, "a:t");
          if (t) parts.push(textOf(t) ?? "");
        }
      }
      if (parts.length > 0) {
        return { text: parts.join("\n") } as NotesSlideDescriptorOptions;
      }
      break;
    }

    return {} as NotesSlideDescriptorOptions;
  },
};
