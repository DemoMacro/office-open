/**
 * Table Styles (a:tblStyleLst) descriptor for PPTX.
 *
 * @module
 */

import { createTableStyleList, type TableStyleListOptions } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr } from "@office-open/xml";

// ── Types ──

export interface TableStylesDescriptorOptions {
  opts?: TableStyleListOptions;
}

// ── Descriptor ──

const DEFAULT_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

export const tableStylesDesc: CustomDescriptor<TableStylesDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (!opts.opts) {
      return `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${DEFAULT_STYLE_ID}"/>`;
    }
    return createTableStyleList(opts.opts).serialize();
  },

  parse(el, _ctx) {
    const defStyle = attr(el, "def") ?? DEFAULT_STYLE_ID;
    return { defStyle } as Record<string, unknown> as Partial<TableStylesDescriptorOptions>;
  },
};
