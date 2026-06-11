/**
 * Slide Sync Properties (p:sldSyncPr) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr } from "@office-open/xml";
import type { SlideSyncOptions } from "@parts/slide/slide-sync-properties";

// ── Descriptor ──

export const slideSyncDesc: CustomDescriptor<SlideSyncOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return (
      `<p:sldSyncPr serverSldId="${opts.serverSldId}" ` +
      `serverSldModifiedTime="${opts.serverSldModifiedTime}" ` +
      `clientInsertedTime="${opts.clientInsertedTime}"/>`
    );
  },

  parse(el, _ctx) {
    return {
      serverSldId: attr(el, "serverSldId") ?? "",
      serverSldModifiedTime: attr(el, "serverSldModifiedTime") ?? "",
      clientInsertedTime: attr(el, "clientInsertedTime") ?? "",
    };
  },
};
