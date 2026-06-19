/**
 * Revision log descriptors for shared-workbook change tracking.
 *
 * Three parts:
 *   - xl/revisionHeaders.xml        (CT_RevisionHeaders) — index of revision logs
 *   - xl/revisions/revisionN.xml    (CT_Revisions)        — one log per header entry
 *   - xl/users.xml                  (CT_Users)            — shared-workbook users
 *
 * Reference: OOXML transitional, sml.xsd — CT_RevisionHeaders (:1860),
 * CT_RevisionHeader (:1898), CT_Revisions (:1877, 12 revision elements),
 * CT_Users (:2100), AG_RevData (:1893).
 *
 * Nested CT_Cell (oc/nc) and CT_Dxf (odxf/ndxf/dxf) carry through verbatim as
 * `*Xml` strings — their full content model is large and revision round-trip only
 * needs lossless preservation, not a re-authored abstraction.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild, children, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ── Revision row/column action (ST_rwColActionType, sml.xsd:2076) ──

export type RowColumnAction = "insertRow" | "deleteRow" | "insertCol" | "deleteCol";

/** Revision action for comments / custom views (ST_RevisionAction, sml.xsd:2084). */
export type RevisionAction = "add" | "delete";

// ════════════════════════════════════════════════════════════════════════════
// Revision Headers — xl/revisionHeaders.xml (CT_RevisionHeaders)
// ════════════════════════════════════════════════════════════════════════════

/** One header entry (CT_RevisionHeader, sml.xsd:1898). */
export interface RevisionHeaderEntry {
  /** Unique GUID for this revision (required, ST_Guid). */
  guid: string;
  /** Date/time of the revision (required, xsd:dateTime). */
  dateTime: string;
  /** User who made the revision (required, ST_Xstring). */
  userName: string;
  /** Relationship ID pointing to the revision log file (required, r:id). */
  rId: string;
  /** Max sheet ID at the time of revision (required). */
  maxSheetId: number;
  /** Sheet ID map entries (required, ≥1). */
  sheetIds: number[];
  /** Reviewed revision IDs (reviewedList.reviewed[].rId, optional). */
  reviewed?: number[];
  /** Minimum revision ID (optional). */
  minRId?: number;
  /** Maximum revision ID (optional). */
  maxRId?: number;
}

/** Options for xl/revisionHeaders.xml (CT_RevisionHeaders, sml.xsd:1860). */
export interface RevisionHeadersOptions {
  /** Unique GUID for the headers collection (required). */
  guid: string;
  /** Revision header entries (required, ≥1). */
  headers: RevisionHeaderEntry[];
  /** Last GUID (optional). */
  lastGuid?: string;
  /** Shared workbook (default true). */
  shared?: boolean;
  /** Disk revisions (default false). */
  diskRevisions?: boolean;
  /** History kept (default true). */
  history?: boolean;
  /** Track revisions (default true). */
  trackRevisions?: boolean;
  /** Exclusive (default false). */
  exclusive?: boolean;
  /** Revision ID counter (default 0). */
  revisionId?: number;
  /** Version (default 1). */
  version?: number;
  /** Keep change history (default true). */
  keepChangeHistory?: boolean;
  /** Protected (default false). */
  protected?: boolean;
  /** Preserve history days (default 30). */
  preserveHistory?: number;
}

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

// ════════════════════════════════════════════════════════════════════════════
// Users — xl/users.xml (CT_Users, sml.xsd:2100)
// ════════════════════════════════════════════════════════════════════════════

/** One shared-workbook user (CT_SharedUser, sml.xsd:2106). */
export interface SharedUserOptions {
  /** GUID (required, ST_Guid). */
  guid: string;
  /** User name (required, ST_Xstring). */
  name: string;
  /** User ID (required). */
  id: number;
  /** Date/time of last edit (required, xsd:dateTime). */
  dateTime: string;
}

/** Options for xl/users.xml (CT_Users, sml.xsd:2100). */
export interface UsersOptions {
  /** Shared-user entries (optional, ≤256). */
  users?: SharedUserOptions[];
}

export const usersDesc: CustomDescriptor<UsersOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (!opts.users || opts.users.length === 0) return undefined;
    const users = opts.users
      .map(
        (u) =>
          `<userInfo guid="${escapeXml(u.guid)}" name="${escapeXml(u.name)}" id="${u.id}"` +
          ` dateTime="${escapeXml(u.dateTime)}"/>`,
      )
      .join("");
    return `<users xmlns="${S_NS}" count="${opts.users.length}">${users}</users>`;
  },

  parse(el, _ctx) {
    const users = children(el, "userInfo").map((u) => ({
      guid: attr(u, "guid") ?? "",
      name: attr(u, "name") ?? "",
      id: Number(attr(u, "id") ?? "0"),
      dateTime: attr(u, "dateTime") ?? "",
    }));
    const result: Partial<UsersOptions> = {};
    if (users.length > 0) result.users = users;
    return result as UsersOptions;
  },
};

// ════════════════════════════════════════════════════════════════════════════
// Revision Log — xl/revisions/revisionN.xml (CT_Revisions, sml.xsd:1877)
// ════════════════════════════════════════════════════════════════════════════

// AG_RevData (sml.xsd:1893): rId required, ua (undo), ra (rejected).
// Present on: rrc, rm, rsnm, ris, rcc, rdn, rcft. Absent on: rcv, rfmt, raf, rcmt, rqt.

/** Row/column insert/delete (CT_RevisionRowColumn, sml.xsd:1943). */
export interface RevisionRowColumnOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (sId, required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** Insertion/deletion action (required). */
  action: RowColumnAction;
  /** End of list (eol, default false). */
  endOfList?: boolean;
  /** Edge tracking (default false). */
  edge?: boolean;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Nested undo/rcc/rfmt children verbatim (optional). */
  childrenXml?: string;
}

/** Cell move (CT_RevisionMove, sml.xsd:1956). */
export interface RevisionMoveOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Source range (required, ST_Ref). */
  source: string;
  /** Destination range (required, ST_Ref). */
  destination: string;
  /** Source sheet ID (default 0). */
  sourceSheetId?: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Nested undo/rcc/rfmt children verbatim (optional). */
  childrenXml?: string;
}

/** Custom view add/delete (CT_RevisionCustomView, sml.xsd:1968 — no AG_RevData). */
export interface RevisionCustomViewOptions {
  /** GUID (required, ST_Guid). */
  guid: string;
  /** Action (required). */
  action: RevisionAction;
}

/** Sheet rename (CT_RevisionSheetRename, sml.xsd:1972). */
export interface RevisionSheetRenameOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Old sheet name (required, ST_Xstring). */
  oldName: string;
  /** New sheet name (required, ST_Xstring). */
  newName: string;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Sheet insertion (CT_RevisionInsertSheet, sml.xsd:1981). */
export interface RevisionInsertSheetOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (required). */
  sheetId: number;
  /** Sheet name (required, ST_Xstring). */
  name: string;
  /** Sheet position (required). */
  sheetPosition: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Cell value change (CT_RevisionCellChange, sml.xsd:1987). */
export interface RevisionCellChangeOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Sheet ID (sId, required). */
  sheetId: number;
  /** Has old differential format (odxf attr, default false). */
  hasOldDxf?: boolean;
  /** XF differential format (xfDxf, default false). */
  xfDxf?: boolean;
  /** Has style (s, default false). */
  style?: boolean;
  /** Has differential format (dxf attr, default false). */
  hasDxf?: boolean;
  /** Number format ID (optional). */
  numFmtId?: number;
  /** Quote prefix (default false). */
  quotePrefix?: boolean;
  /** Old quote prefix (default false). */
  oldQuotePrefix?: boolean;
  /** Phonetic (ph, default false). */
  phonetic?: boolean;
  /** Old phonetic (default false). */
  oldPhonetic?: boolean;
  /** End of list formula update (default false). */
  endOfListFormulaUpdate?: boolean;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Old cell (oc) verbatim CT_Cell XML (optional). */
  oldCellXml?: string;
  /** New cell (nc) verbatim CT_Cell XML (required). */
  newCellXml: string;
  /** Old differential format (odxf) verbatim CT_Dxf XML (optional). */
  oldDxfXml?: string;
  /** New differential format (ndxf) verbatim CT_Dxf XML (optional). */
  newDxfXml?: string;
}

/** Formatting change (CT_RevisionFormatting, sml.xsd:2008 — no AG_RevData). */
export interface RevisionFormattingOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Sequence of cell refs (sqref, required). */
  sqref: string;
  /** XF differential format (default false). */
  xfDxf?: boolean;
  /** Has style (s, default false). */
  style?: boolean;
  /** Start index (optional). */
  start?: number;
  /** Length (optional). */
  length?: number;
  /** Differential format (dxf) verbatim CT_Dxf XML (optional). */
  dxfXml?: string;
}

/** Auto formatting (CT_RevisionAutoFormatting, sml.xsd:2020 — no AG_RevData). */
export interface RevisionAutoFormattingOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** AG_AutoFormat attributes verbatim (optional, e.g. format/align/...). */
  autoFormatXml?: string;
}

/** Defined name change (CT_RevisionDefinedName, sml.xsd:2038). */
export interface RevisionDefinedNameOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Defined name (required, ST_Xstring). */
  name: string;
  /** Local sheet ID (optional; -1 or omitted = workbook-global). */
  localSheetId?: number;
  /** Custom view (default false). */
  customView?: boolean;
  /** New formula expression (optional). */
  formula?: string;
  /** Old formula expression (optional). */
  oldFormula?: string;
  /** Is function (default false). */
  function?: boolean;
  /** Old function flag (default false). */
  oldFunction?: boolean;
  /** Function group ID (optional). */
  functionGroupId?: number;
  /** Old function group ID (optional). */
  oldFunctionGroupId?: number;
  /** Shortcut key (optional, unsignedByte). */
  shortcutKey?: number;
  /** Old shortcut key (optional, unsignedByte). */
  oldShortcutKey?: number;
  /** Hidden (default false). */
  hidden?: boolean;
  /** Old hidden flag (default false). */
  oldHidden?: boolean;
  /** Custom menu (optional). */
  customMenu?: string;
  /** Old custom menu (optional). */
  oldCustomMenu?: string;
  /** Description (optional). */
  description?: string;
  /** Old description (optional). */
  oldDescription?: string;
  /** Help text (optional). */
  help?: string;
  /** Old help text (optional). */
  oldHelp?: string;
  /** Status bar text (optional). */
  statusBar?: string;
  /** Old status bar text (optional). */
  oldStatusBar?: string;
  /** Comment (optional). */
  comment?: string;
  /** Old comment (optional). */
  oldComment?: string;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
}

/** Comment revision (CT_RevisionComment, sml.xsd:2025 — no AG_RevData). */
export interface RevisionCommentOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Cell reference (required, ST_CellRef). */
  cell: string;
  /** GUID (required, ST_Guid). */
  guid: string;
  /** Action (default add). */
  action?: RevisionAction;
  /** Always show (default false). */
  alwaysShow?: boolean;
  /** Old comment flag (default false). */
  old?: boolean;
  /** Hidden row (default false). */
  hiddenRow?: boolean;
  /** Hidden column (default false). */
  hiddenColumn?: boolean;
  /** Author (required, ST_Xstring). */
  author: string;
  /** Old text length (default 0). */
  oldLength?: number;
  /** New text length (default 0). */
  newLength?: number;
}

/** Query table field change (CT_RevisionQueryTableField, sml.xsd:2071 — no AG_RevData). */
export interface RevisionQueryTableFieldOptions {
  /** Sheet ID (required). */
  sheetId: number;
  /** Affected range (required, ST_Ref). */
  ref: string;
  /** Field ID (required). */
  fieldId: number;
}

/** Conflict (CT_RevisionConflict, sml.xsd:2067). */
export interface RevisionConflictOptions {
  /** Revision ID (AG_RevData rId, required). */
  rId: number;
  /** Revision was performed by an undo (ua). */
  undo?: boolean;
  /** Revision was rejected (ra). */
  rejected?: boolean;
  /** Sheet ID (optional). */
  sheetId?: number;
}

/** Discriminated union of all CT_Revisions choice elements (sml.xsd:1877-1891). */
export type RevisionEntry =
  | { type: "rowColumn"; data: RevisionRowColumnOptions }
  | { type: "move"; data: RevisionMoveOptions }
  | { type: "customView"; data: RevisionCustomViewOptions }
  | { type: "sheetRename"; data: RevisionSheetRenameOptions }
  | { type: "insertSheet"; data: RevisionInsertSheetOptions }
  | { type: "cellChange"; data: RevisionCellChangeOptions }
  | { type: "formatting"; data: RevisionFormattingOptions }
  | { type: "autoFormatting"; data: RevisionAutoFormattingOptions }
  | { type: "definedName"; data: RevisionDefinedNameOptions }
  | { type: "comment"; data: RevisionCommentOptions }
  | { type: "queryTableField"; data: RevisionQueryTableFieldOptions }
  | { type: "conflict"; data: RevisionConflictOptions };

/** Options for xl/revisions/revisionN.xml (CT_Revisions, sml.xsd:1877). */
export interface RevisionLogOptions {
  /** Revision entries (the CT_Revisions choice sequence). */
  revisions: RevisionEntry[];
}

// ── Per-revision stringify ──

function agRevData(data: { rId: number; undo?: boolean; rejected?: boolean }): string {
  let s = ` rId="${data.rId}"`;
  if (data.undo) s += ` ua="1"`;
  if (data.rejected) s += ` ra="1"`;
  return s;
}

function stringifyEntry(entry: RevisionEntry): string {
  switch (entry.type) {
    case "rowColumn": {
      const d = entry.data;
      let a = agRevData(d) + ` sId="${d.sheetId}" ref="${escapeXml(d.ref)}" action="${d.action}"`;
      if (d.endOfList) a += ` eol="1"`;
      if (d.edge) a += ` edge="1"`;
      return `<rrc${a}>${d.childrenXml ?? ""}</rrc>`;
    }
    case "move": {
      const d = entry.data;
      let a =
        agRevData(d) +
        ` sheetId="${d.sheetId}" source="${escapeXml(d.source)}" destination="${escapeXml(d.destination)}"`;
      if (d.sourceSheetId !== undefined) a += ` sourceSheetId="${d.sourceSheetId}"`;
      return `<rm${a}>${d.childrenXml ?? ""}</rm>`;
    }
    case "customView": {
      const d = entry.data;
      return `<rcv guid="${escapeXml(d.guid)}" action="${d.action}"/>`;
    }
    case "sheetRename": {
      const d = entry.data;
      return `<rsnm${agRevData(d)} sheetId="${d.sheetId}" oldName="${escapeXml(
        d.oldName,
      )}" newName="${escapeXml(d.newName)}"/>`;
    }
    case "insertSheet": {
      const d = entry.data;
      return `<ris${agRevData(d)} sheetId="${d.sheetId}" name="${escapeXml(
        d.name,
      )}" sheetPosition="${d.sheetPosition}"/>`;
    }
    case "cellChange": {
      const d = entry.data;
      let a = agRevData(d) + ` sId="${d.sheetId}"`;
      if (d.hasOldDxf) a += ` odxf="1"`;
      if (d.xfDxf) a += ` xfDxf="1"`;
      if (d.style) a += ` s="1"`;
      if (d.hasDxf) a += ` dxf="1"`;
      if (d.numFmtId !== undefined) a += ` numFmtId="${d.numFmtId}"`;
      if (d.quotePrefix) a += ` quotePrefix="1"`;
      if (d.oldQuotePrefix) a += ` oldQuotePrefix="1"`;
      if (d.phonetic) a += ` ph="1"`;
      if (d.oldPhonetic) a += ` oldPh="1"`;
      if (d.endOfListFormulaUpdate) a += ` endOfListFormulaUpdate="1"`;
      const children = [d.oldCellXml ?? "", d.newCellXml, d.oldDxfXml ?? "", d.newDxfXml ?? ""]
        .filter(Boolean)
        .join("");
      return `<rcc${a}>${children}</rcc>`;
    }
    case "formatting": {
      const d = entry.data;
      let a = ` sheetId="${d.sheetId}" sqref="${escapeXml(d.sqref)}"`;
      if (d.xfDxf) a += ` xfDxf="1"`;
      if (d.style) a += ` s="1"`;
      if (d.start !== undefined) a += ` start="${d.start}"`;
      if (d.length !== undefined) a += ` length="${d.length}"`;
      return `<rfmt${a}>${d.dxfXml ?? ""}</rfmt>`;
    }
    case "autoFormatting": {
      const d = entry.data;
      return `<raf sheetId="${d.sheetId}" ref="${escapeXml(d.ref)}"${d.autoFormatXml ?? ""}/>`;
    }
    case "definedName": {
      const d = entry.data;
      let a = agRevData(d) + ` name="${escapeXml(d.name)}"`;
      if (d.localSheetId !== undefined) a += ` localSheetId="${d.localSheetId}"`;
      if (d.customView) a += ` customView="1"`;
      if (d.function) a += ` function="1"`;
      if (d.oldFunction) a += ` oldFunction="1"`;
      if (d.functionGroupId !== undefined) a += ` functionGroupId="${d.functionGroupId}"`;
      if (d.oldFunctionGroupId !== undefined) a += ` oldFunctionGroupId="${d.oldFunctionGroupId}"`;
      if (d.shortcutKey !== undefined) a += ` shortcutKey="${d.shortcutKey}"`;
      if (d.oldShortcutKey !== undefined) a += ` oldShortcutKey="${d.oldShortcutKey}"`;
      if (d.hidden) a += ` hidden="1"`;
      if (d.oldHidden) a += ` oldHidden="1"`;
      const xstring = (v: string | undefined, attr: string) =>
        v !== undefined ? ` ${attr}="${escapeXml(v)}"` : "";
      a += xstring(d.customMenu, "customMenu") + xstring(d.oldCustomMenu, "oldCustomMenu");
      a += xstring(d.description, "description") + xstring(d.oldDescription, "oldDescription");
      a += xstring(d.help, "help") + xstring(d.oldHelp, "oldHelp");
      a += xstring(d.statusBar, "statusBar") + xstring(d.oldStatusBar, "oldStatusBar");
      a += xstring(d.comment, "comment") + xstring(d.oldComment, "oldComment");
      const children = [
        d.formula !== undefined ? `<formula>${escapeXml(d.formula)}</formula>` : "",
        d.oldFormula !== undefined ? `<oldFormula>${escapeXml(d.oldFormula)}</oldFormula>` : "",
      ]
        .filter(Boolean)
        .join("");
      return `<rdn${a}>${children}</rdn>`;
    }
    case "comment": {
      const d = entry.data;
      let a = ` sheetId="${d.sheetId}" cell="${escapeXml(d.cell)}" guid="${escapeXml(
        d.guid,
      )}" action="${d.action ?? "add"}" author="${escapeXml(d.author)}"`;
      if (d.alwaysShow) a += ` alwaysShow="1"`;
      if (d.old) a += ` old="1"`;
      if (d.hiddenRow) a += ` hiddenRow="1"`;
      if (d.hiddenColumn) a += ` hiddenColumn="1"`;
      if (d.oldLength !== undefined) a += ` oldLength="${d.oldLength}"`;
      if (d.newLength !== undefined) a += ` newLength="${d.newLength}"`;
      return `<rcmt${a}/>`;
    }
    case "queryTableField": {
      const d = entry.data;
      return `<rqt sheetId="${d.sheetId}" ref="${escapeXml(d.ref)}" fieldId="${d.fieldId}"/>`;
    }
    case "conflict": {
      const d = entry.data;
      let a = agRevData(d);
      if (d.sheetId !== undefined) a += ` sheetId="${d.sheetId}"`;
      return `<rcft${a}/>`;
    }
  }
}

// ── Per-revision parse ──

/** Serializes an element's children back to a raw XML string (for rawXml passthrough). */
function childrenToXml(el: XmlElement | null): string {
  if (!el || !el.elements) return "";
  // Re-emit element children verbatim via their name + attributes + text + nested.
  return el.elements
    .filter((c) => c.type === "element")
    .map((c) => elementToXml(c as XmlElement))
    .join("");
}

function elementToXml(el: XmlElement): string {
  const attrStr = Object.entries(el.attributes ?? {})
    .map(([k, v]) => ` ${k}="${escapeXml(String(v))}"`)
    .join("");
  const inner = el.elements
    ? el.elements
        .map((c) => {
          if (c.type === "text") return escapeXml(textOf({ elements: [c] } as XmlElement) ?? "");
          if (c.type === "element") return elementToXml(c as XmlElement);
          return "";
        })
        .join("")
    : "";
  return `<${el.name}${attrStr}>${inner}</${el.name}>`;
}

/** Returns the first element child of a node as raw XML string. */
function firstChildXml(el: XmlElement | null, name: string): string | undefined {
  const child = findChild(el ?? undefined, name);
  return child ? elementToXml(child) : undefined;
}

function parseBool(el: XmlElement, name: string): boolean | undefined {
  const v = attr(el, name);
  if (v === undefined) return undefined;
  return v === "1" || v === "true";
}

function parseEntry(el: XmlElement): RevisionEntry | undefined {
  switch (el.name) {
    case "rrc": {
      const d: Partial<RevisionRowColumnOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        sheetId: Number(attr(el, "sId") ?? "0"),
        ref: attr(el, "ref") ?? "",
        action: (attr(el, "action") ?? "insertRow") as RowColumnAction,
      };
      const endOfList = parseBool(el, "eol");
      if (endOfList) d.endOfList = endOfList;
      const edge = parseBool(el, "edge");
      if (edge) d.edge = edge;
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      const childrenXml = childrenToXml(el);
      if (childrenXml) d.childrenXml = childrenXml;
      return { type: "rowColumn", data: d as RevisionRowColumnOptions };
    }
    case "rm": {
      const d: Partial<RevisionMoveOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        source: attr(el, "source") ?? "",
        destination: attr(el, "destination") ?? "",
      };
      const sourceSheetId = attr(el, "sourceSheetId");
      if (sourceSheetId !== undefined) d.sourceSheetId = Number(sourceSheetId);
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      const childrenXml = childrenToXml(el);
      if (childrenXml) d.childrenXml = childrenXml;
      return { type: "move", data: d as RevisionMoveOptions };
    }
    case "rcv":
      return {
        type: "customView",
        data: {
          guid: attr(el, "guid") ?? "",
          action: (attr(el, "action") ?? "add") as RevisionAction,
        },
      };
    case "rsnm": {
      const d: Partial<RevisionSheetRenameOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        oldName: attr(el, "oldName") ?? "",
        newName: attr(el, "newName") ?? "",
      };
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      return { type: "sheetRename", data: d as RevisionSheetRenameOptions };
    }
    case "ris": {
      const d: Partial<RevisionInsertSheetOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        name: attr(el, "name") ?? "",
        sheetPosition: Number(attr(el, "sheetPosition") ?? "0"),
      };
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      return { type: "insertSheet", data: d as RevisionInsertSheetOptions };
    }
    case "rcc": {
      const d: Partial<RevisionCellChangeOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        sheetId: Number(attr(el, "sId") ?? "0"),
        newCellXml: firstChildXml(el, "nc") ?? "",
      };
      const hasOldDxf = parseBool(el, "odxf");
      if (hasOldDxf) d.hasOldDxf = hasOldDxf;
      const xfDxf = parseBool(el, "xfDxf");
      if (xfDxf) d.xfDxf = xfDxf;
      const style = parseBool(el, "s");
      if (style) d.style = style;
      const hasDxf = parseBool(el, "dxf");
      if (hasDxf) d.hasDxf = hasDxf;
      const numFmtId = attr(el, "numFmtId");
      if (numFmtId !== undefined) d.numFmtId = Number(numFmtId);
      const quotePrefix = parseBool(el, "quotePrefix");
      if (quotePrefix) d.quotePrefix = quotePrefix;
      const oldQuotePrefix = parseBool(el, "oldQuotePrefix");
      if (oldQuotePrefix) d.oldQuotePrefix = oldQuotePrefix;
      const phonetic = parseBool(el, "ph");
      if (phonetic) d.phonetic = phonetic;
      const oldPhonetic = parseBool(el, "oldPh");
      if (oldPhonetic) d.oldPhonetic = oldPhonetic;
      const endOfList = parseBool(el, "endOfListFormulaUpdate");
      if (endOfList) d.endOfListFormulaUpdate = endOfList;
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      const oc = firstChildXml(el, "oc");
      if (oc) d.oldCellXml = oc;
      const odxf = firstChildXml(el, "odxf");
      if (odxf) d.oldDxfXml = odxf;
      const ndxf = firstChildXml(el, "ndxf");
      if (ndxf) d.newDxfXml = ndxf;
      return { type: "cellChange", data: d as RevisionCellChangeOptions };
    }
    case "rfmt": {
      const d: Partial<RevisionFormattingOptions> = {
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        sqref: attr(el, "sqref") ?? "",
      };
      const xfDxf = parseBool(el, "xfDxf");
      if (xfDxf) d.xfDxf = xfDxf;
      const style = parseBool(el, "s");
      if (style) d.style = style;
      const start = attr(el, "start");
      if (start !== undefined) d.start = Number(start);
      const length = attr(el, "length");
      if (length !== undefined) d.length = Number(length);
      const dxfXml = firstChildXml(el, "dxf");
      if (dxfXml) d.dxfXml = dxfXml;
      return { type: "formatting", data: d as RevisionFormattingOptions };
    }
    case "raf": {
      const d: Partial<RevisionAutoFormattingOptions> = {
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        ref: attr(el, "ref") ?? "",
      };
      // AG_AutoFormat is a fixed set of optional attributes; capture verbatim.
      const autoAttrs = Object.entries(el.attributes ?? {})
        .filter(([k]) => k !== "sheetId" && k !== "ref")
        .map(([k, v]) => ` ${k}="${escapeXml(String(v))}"`)
        .join("");
      if (autoAttrs) d.autoFormatXml = autoAttrs;
      return { type: "autoFormatting", data: d as RevisionAutoFormattingOptions };
    }
    case "rdn": {
      const d: Partial<RevisionDefinedNameOptions> = {
        rId: Number(attr(el, "rId") ?? "0"),
        name: attr(el, "name") ?? "",
      };
      const localSheetId = attr(el, "localSheetId");
      if (localSheetId !== undefined) d.localSheetId = Number(localSheetId);
      readBool(el, "customView", (v) => (d.customView = v));
      readBool(el, "function", (v) => (d.function = v));
      readBool(el, "oldFunction", (v) => (d.oldFunction = v));
      readNum(el, "functionGroupId", (v) => (d.functionGroupId = v));
      readNum(el, "oldFunctionGroupId", (v) => (d.oldFunctionGroupId = v));
      readNum(el, "shortcutKey", (v) => (d.shortcutKey = v));
      readNum(el, "oldShortcutKey", (v) => (d.oldShortcutKey = v));
      readBool(el, "hidden", (v) => (d.hidden = v));
      readBool(el, "oldHidden", (v) => (d.oldHidden = v));
      readStr(el, "customMenu", (v) => (d.customMenu = v));
      readStr(el, "oldCustomMenu", (v) => (d.oldCustomMenu = v));
      readStr(el, "description", (v) => (d.description = v));
      readStr(el, "oldDescription", (v) => (d.oldDescription = v));
      readStr(el, "help", (v) => (d.help = v));
      readStr(el, "oldHelp", (v) => (d.oldHelp = v));
      readStr(el, "statusBar", (v) => (d.statusBar = v));
      readStr(el, "oldStatusBar", (v) => (d.oldStatusBar = v));
      readStr(el, "comment", (v) => (d.comment = v));
      readStr(el, "oldComment", (v) => (d.oldComment = v));
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      const formula = findChild(el, "formula");
      if (formula) d.formula = textOf(formula) ?? "";
      const oldFormula = findChild(el, "oldFormula");
      if (oldFormula) d.oldFormula = textOf(oldFormula) ?? "";
      return { type: "definedName", data: d as RevisionDefinedNameOptions };
    }
    case "rcmt": {
      const d: Partial<RevisionCommentOptions> = {
        sheetId: Number(attr(el, "sheetId") ?? "0"),
        cell: attr(el, "cell") ?? "",
        guid: attr(el, "guid") ?? "",
        author: attr(el, "author") ?? "",
      };
      const action = attr(el, "action");
      if (action) d.action = action as RevisionAction;
      readBool(el, "alwaysShow", (v) => (d.alwaysShow = v));
      readBool(el, "old", (v) => (d.old = v));
      readBool(el, "hiddenRow", (v) => (d.hiddenRow = v));
      readBool(el, "hiddenColumn", (v) => (d.hiddenColumn = v));
      readNum(el, "oldLength", (v) => (d.oldLength = v));
      readNum(el, "newLength", (v) => (d.newLength = v));
      return { type: "comment", data: d as RevisionCommentOptions };
    }
    case "rqt":
      return {
        type: "queryTableField",
        data: {
          sheetId: Number(attr(el, "sheetId") ?? "0"),
          ref: attr(el, "ref") ?? "",
          fieldId: Number(attr(el, "fieldId") ?? "0"),
        },
      };
    case "rcft": {
      const d: Partial<RevisionConflictOptions> = { rId: Number(attr(el, "rId") ?? "0") };
      const undo = parseBool(el, "ua");
      if (undo) d.undo = undo;
      const rejected = parseBool(el, "ra");
      if (rejected) d.rejected = rejected;
      const sheetId = attr(el, "sheetId");
      if (sheetId !== undefined) d.sheetId = Number(sheetId);
      return { type: "conflict", data: d as RevisionConflictOptions };
    }
    default:
      return undefined;
  }
}

export const revisionLogDesc: CustomDescriptor<RevisionLogOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.revisions.length === 0) return undefined;
    return `<revisions xmlns="${S_NS}">${opts.revisions.map(stringifyEntry).join("")}</revisions>`;
  },

  parse(el, _ctx) {
    const revisions: RevisionEntry[] = [];
    for (const child of el.elements ?? []) {
      if (child.type !== "element") continue;
      const entry = parseEntry(child as XmlElement);
      if (entry) revisions.push(entry);
    }
    return { revisions } as RevisionLogOptions;
  },
};

// ── Attribute readers (set-callback style keeps parse paths on concrete types) ──

function readBool(el: XmlElement, name: string, set: (v: boolean) => void): void {
  const raw = attr(el, name);
  if (raw === "1" || raw === "true") set(true);
}

function readNum(el: XmlElement, name: string, set: (v: number) => void): void {
  const raw = attr(el, name);
  if (raw !== undefined) set(Number(raw));
}

function readStr(el: XmlElement, name: string, set: (v: string) => void): void {
  const raw = attr(el, name);
  if (raw !== undefined) set(raw);
}
