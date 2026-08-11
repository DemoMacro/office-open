/**
 * Revision log — per-revision parse helpers and shared attribute readers.
 *
 * @module
 */

import { attr, escapeXml, findChild, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  RevisionAction,
  RevisionAutoFormattingOptions,
  RevisionCellChangeOptions,
  RevisionCommentOptions,
  RevisionConflictOptions,
  RevisionDefinedNameOptions,
  RevisionEntry,
  RevisionFormattingOptions,
  RevisionInsertSheetOptions,
  RevisionMoveOptions,
  RevisionRowColumnOptions,
  RevisionSheetRenameOptions,
  RowColumnAction,
} from "./types";

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
  // nativeTypeAttributes (xlsx parse path) coerces "1" to number 1.
  return String(v) === "1" || v === "true";
}

export function parseEntry(el: XmlElement): RevisionEntry | undefined {
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

// ── Attribute readers (set-callback style keeps parse paths on concrete types) ──

export function readBool(el: XmlElement, name: string, set: (v: boolean) => void): void {
  const raw = attr(el, name);
  // nativeTypeAttributes (xlsx parse path) coerces "1" to number 1.
  if (String(raw) === "1" || raw === "true") set(true);
}

export function readNum(el: XmlElement, name: string, set: (v: number) => void): void {
  const raw = attr(el, name);
  if (raw !== undefined) set(Number(raw));
}

export function readStr(el: XmlElement, name: string, set: (v: string) => void): void {
  const raw = attr(el, name);
  if (raw !== undefined) set(raw);
}
