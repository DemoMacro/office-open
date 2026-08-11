/**
 * Revision headers — xl/revisionHeaders.xml (CT_RevisionHeaders) descriptor.
 *
 * Index of revision logs in a shared workbook.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, children, escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { readBool, readNum } from "./parse";
import { R_NS, S_NS } from "./stringify";
import type { RevisionHeaderEntry, RevisionHeadersOptions } from "./types";

function stringifyHeader(h: RevisionHeaderEntry): string {
  const sheetIds = h.sheetIds.map((id) => `<sheetId val="${id}"/>`).join("");
  const reviewedXml =
    h.reviewed && h.reviewed.length > 0
      ? `<reviewedList count="${h.reviewed.length}">${h.reviewed
          .map((r) => `<reviewed rId="${r}"/>`)
          .join("")}</reviewedList>`
      : "";
  let attrs =
    ` guid="${escapeXml(h.guid)}" dateTime="${escapeXml(h.dateTime)}"` +
    ` maxSheetId="${h.maxSheetId}" userName="${escapeXml(h.userName)}"` +
    ` r:id="${escapeXml(h.rId)}"`;
  if (h.minRId !== undefined) attrs += ` minRId="${h.minRId}"`;
  if (h.maxRId !== undefined) attrs += ` maxRId="${h.maxRId}"`;
  return `<header${attrs}><sheetIdMap count="${h.sheetIds.length}">${sheetIds}</sheetIdMap>${reviewedXml}</header>`;
}

function parseHeader(el: XmlElement): RevisionHeaderEntry {
  const guid = attr(el, "guid") ?? "";
  const dateTime = attr(el, "dateTime") ?? "";
  const userName = attr(el, "userName") ?? "";
  // r:id lives in the relationships namespace
  const rId = String(el.attributes?.["r:id"] ?? el.attributes?.["id"] ?? "");
  const maxSheetId = Number(attr(el, "maxSheetId") ?? "0");
  const sheetIdMap = findChild(el, "sheetIdMap");
  const sheetIds = children(sheetIdMap, "sheetId").map((s) => Number(attr(s, "val") ?? "0"));
  const result: RevisionHeaderEntry = { guid, dateTime, userName, rId, maxSheetId, sheetIds };
  const reviewedList = findChild(el, "reviewedList");
  if (reviewedList) {
    const reviewed = children(reviewedList, "reviewed").map((r) => Number(attr(r, "rId") ?? "0"));
    if (reviewed.length > 0) result.reviewed = reviewed;
  }
  const minRId = attr(el, "minRId");
  if (minRId !== undefined) result.minRId = Number(minRId);
  const maxRId = attr(el, "maxRId");
  if (maxRId !== undefined) result.maxRId = Number(maxRId);
  return result;
}

export const revisionHeadersDesc: CustomDescriptor<RevisionHeadersOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.headers.length === 0) return undefined;
    let attrs = ` xmlns="${S_NS}" xmlns:r="${R_NS}" guid="${escapeXml(opts.guid)}"`;
    if (opts.lastGuid) attrs += ` lastGuid="${escapeXml(opts.lastGuid)}"`;
    if (opts.shared !== undefined) attrs += ` shared="${opts.shared ? 1 : 0}"`;
    if (opts.diskRevisions !== undefined) attrs += ` diskRevisions="${opts.diskRevisions ? 1 : 0}"`;
    if (opts.history !== undefined) attrs += ` history="${opts.history ? 1 : 0}"`;
    if (opts.trackRevisions !== undefined)
      attrs += ` trackRevisions="${opts.trackRevisions ? 1 : 0}"`;
    if (opts.exclusive !== undefined) attrs += ` exclusive="${opts.exclusive ? 1 : 0}"`;
    if (opts.revisionId !== undefined) attrs += ` revisionId="${opts.revisionId}"`;
    if (opts.version !== undefined) attrs += ` version="${opts.version}"`;
    if (opts.keepChangeHistory !== undefined)
      attrs += ` keepChangeHistory="${opts.keepChangeHistory ? 1 : 0}"`;
    if (opts.protected !== undefined) attrs += ` protected="${opts.protected ? 1 : 0}"`;
    if (opts.preserveHistory !== undefined) attrs += ` preserveHistory="${opts.preserveHistory}"`;
    return `<headers${attrs}>${opts.headers.map(stringifyHeader).join("")}</headers>`;
  },

  parse(el, _ctx) {
    const guid = attr(el, "guid") ?? "";
    const headers = children(el, "header").map(parseHeader);
    const result: Partial<RevisionHeadersOptions> = { guid, headers };
    const lastGuid = attr(el, "lastGuid");
    if (lastGuid !== undefined) result.lastGuid = lastGuid;
    readBool(el, "shared", (v) => (result.shared = v));
    readBool(el, "diskRevisions", (v) => (result.diskRevisions = v));
    readBool(el, "history", (v) => (result.history = v));
    readBool(el, "trackRevisions", (v) => (result.trackRevisions = v));
    readBool(el, "exclusive", (v) => (result.exclusive = v));
    readNum(el, "revisionId", (v) => (result.revisionId = v));
    readNum(el, "version", (v) => (result.version = v));
    readBool(el, "keepChangeHistory", (v) => (result.keepChangeHistory = v));
    readBool(el, "protected", (v) => (result.protected = v));
    readNum(el, "preserveHistory", (v) => (result.preserveHistory = v));
    return result as RevisionHeadersOptions;
  },
};
