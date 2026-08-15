/**
 * Revision log — per-revision stringify helpers and namespace constants.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";

import type { RevisionEntry, RevisionNestedChild, RevisionUndoOptions } from "./types";

// Namespace constants shared across the headers/users/log descriptors.
const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ── Per-revision stringify ──

function agRevData(data: { rId: number; undo?: boolean; rejected?: boolean }): string {
  let s = ` rId="${data.rId}"`;
  if (data.undo) s += ` ua="1"`;
  if (data.rejected) s += ` ra="1"`;
  return s;
}

function stringifyUndo(d: RevisionUndoOptions): string {
  let a = ` index="${d.index}" exp="${d.expression}" dr="${escapeXml(d.dr)}"`;
  if (d.ref3D) a += ` ref3D="1"`;
  if (d.array) a += ` array="1"`;
  if (d.v) a += ` v="1"`;
  if (d.nf) a += ` nf="1"`;
  if (d.cs) a += ` cs="1"`;
  if (d.dn !== undefined) a += ` dn="${escapeXml(d.dn)}"`;
  if (d.r !== undefined) a += ` r="${escapeXml(d.r)}"`;
  if (d.sId !== undefined) a += ` sId="${d.sId}"`;
  return `<undo${a}/>`;
}

function stringifyNestedChildren(children: RevisionNestedChild[] | undefined): string {
  if (!children) return "";
  return children
    .map((child) => {
      switch (child.kind) {
        case "undo":
          return stringifyUndo(child.data);
        case "cellChange":
          return stringifyEntry({ type: "cellChange", data: child.data });
        case "formatting":
          return stringifyEntry({ type: "formatting", data: child.data });
      }
    })
    .join("");
}

export function stringifyEntry(entry: RevisionEntry): string {
  switch (entry.type) {
    case "rowColumn": {
      const d = entry.data;
      let a = agRevData(d) + ` sId="${d.sheetId}" ref="${escapeXml(d.ref)}" action="${d.action}"`;
      if (d.endOfList) a += ` eol="1"`;
      if (d.edge) a += ` edge="1"`;
      return `<rrc${a}>${stringifyNestedChildren(d.children)}</rrc>`;
    }
    case "move": {
      const d = entry.data;
      let a =
        agRevData(d) +
        ` sheetId="${d.sheetId}" source="${escapeXml(d.source)}" destination="${escapeXml(d.destination)}"`;
      if (d.sourceSheetId !== undefined) a += ` sourceSheetId="${d.sourceSheetId}"`;
      return `<rm${a}>${stringifyNestedChildren(d.children)}</rm>`;
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

export { R_NS, S_NS };
