/**
 * Content Types ([Content_Types].xml) descriptor for PPTX.
 *
 * @module
 */

import type { ContentTypes } from "@file/content-types";
import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface ContentTypesDescriptorOptions {
  /** The mutable ContentTypes builder instance. */
  builder: ContentTypes;
}

// ── Descriptor ──

export const contentTypesDesc: CustomDescriptor<ContentTypesDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.builder.serialize();
  },

  parse(el, _ctx) {
    const overrides: { partName: string; contentType: string }[] = [];
    const defaults: { extension: string; contentType: string }[] = [];

    for (const child of el.elements ?? []) {
      if (child.name === "Override" && child.attributes) {
        overrides.push({
          partName: String(child.attributes["PartName"] ?? ""),
          contentType: String(child.attributes["ContentType"] ?? ""),
        });
      } else if (child.name === "Default" && child.attributes) {
        defaults.push({
          extension: String(child.attributes["Extension"] ?? ""),
          contentType: String(child.attributes["ContentType"] ?? ""),
        });
      }
    }

    return { overrides, defaults } as Record<
      string,
      unknown
    > as Partial<ContentTypesDescriptorOptions>;
  },
};
