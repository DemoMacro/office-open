/**
 * QueryTable types and descriptor for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_QueryTable
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml, findChild } from "@office-open/xml";

// ── Options ──

/** Deleted query-table field (CT_DeletedField). */
export interface QueryTableDeletedFieldOptions {
  /** Field name that was deleted (required) */
  name: string;
}

/** Query-table field (CT_QueryTableField). */
export interface QueryTableFieldOptions {
  /** Field ID (required, unique within the query table) */
  id: number;
  /** Field name */
  name?: string;
  /** Bound to a table column (default true) */
  dataBound?: boolean;
  /** Field shows row numbers (default false) */
  rowNumbers?: boolean;
  /** Fill formulas (default false) */
  fillFormulas?: boolean;
  /** Clipped (default false) */
  clipped?: boolean;
  /** Owning table column id (default 0) */
  tableColumnId?: number;
}

/** Query-table refresh info (CT_QueryTableRefresh). */
export interface QueryTableRefreshOptions {
  /** Preserve sort/filter layout (default true) */
  preserveSortFilterLayout?: boolean;
  /** Field ID wrapped (default false) */
  fieldIdWrapped?: boolean;
  /** Headers in last refresh (default true) */
  headersInLastRefresh?: boolean;
  /** Minimum refresh version (default 0) */
  minimumVersion?: number;
  /** Next unique ID for new rows (default 1) */
  nextId?: number;
  /** Unbound columns inserted left (default 0) */
  unboundColumnsLeft?: number;
  /** Unbound columns appended right (default 0) */
  unboundColumnsRight?: number;
  /** Field layout after refresh (CT_QueryTableRefresh → queryTableFields) */
  queryTableFields?: QueryTableFieldOptions[];
  /** Deleted fields (CT_QueryTableDeletedFields) */
  deletedFields?: QueryTableDeletedFieldOptions[];
}

/** Options for xl/queryTables/queryTable{n}.xml (CT_QueryTable). */
export interface QueryTableOptions {
  /** Query table name (required by XSD) */
  name?: string;
  /** Headers row shown (default true) */
  headers?: boolean;
  /** Row numbers shown (default false) */
  rowNumbers?: boolean;
  /** Refresh disabled (default false) */
  disableRefresh?: boolean;
  /** Background refresh (default true) */
  backgroundRefresh?: boolean;
  /** First background refresh (default false) */
  firstBackgroundRefresh?: boolean;
  /** Refresh on load (default false) */
  refreshOnLoad?: boolean;
  /** Grow/shrink behavior: "insertDelete" | "insertWhole" | "overwriteWhole" (default "insertDelete") */
  growShrinkType?: "insertDelete" | "insertWhole" | "overwriteWhole";
  /** Fill formulas on refresh (default false) */
  fillFormulas?: boolean;
  /** Remove data on save (default false) */
  removeDataOnSave?: boolean;
  /** Edit disabled (default false) */
  disableEdit?: boolean;
  /** Preserve formatting (default true) */
  preserveFormatting?: boolean;
  /** Adjust column width on refresh (default false) */
  adjustColumnWidth?: boolean;
  /** Connection ID referencing the workbook connection */
  connectionId: number;
  /** Auto format applied */
  autoFormat?: boolean;
  /** Intermediate rows preserved (CT_QueryTable `@intermediate`) */
  intermediate?: boolean;
  /** Refresh info (CT_QueryTableRefresh) */
  queryTableRefresh?: QueryTableRefreshOptions;
}

// ── Descriptor ──

export const queryTableDesc: CustomDescriptor<QueryTableOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const attrs: string[] = [];
    if (opts.name !== undefined) attrs.push(`name="${escapeXml(opts.name)}"`);
    if (opts.headers === false) attrs.push('headers="0"');
    if (opts.rowNumbers) attrs.push('rowNumbers="1"');
    if (opts.disableRefresh) attrs.push('disableRefresh="1"');
    if (opts.backgroundRefresh === false) attrs.push('backgroundRefresh="0"');
    if (opts.firstBackgroundRefresh) attrs.push('firstBackgroundRefresh="1"');
    if (opts.refreshOnLoad) attrs.push('refreshOnLoad="1"');
    if (opts.growShrinkType !== undefined) attrs.push(`growShrinkType="${opts.growShrinkType}"`);
    if (opts.fillFormulas) attrs.push('fillFormulas="1"');
    if (opts.removeDataOnSave) attrs.push('removeDataOnSave="1"');
    if (opts.disableEdit) attrs.push('disableEdit="1"');
    if (opts.preserveFormatting === false) attrs.push('preserveFormatting="0"');
    if (opts.adjustColumnWidth) attrs.push('adjustColumnWidth="1"');
    if (opts.autoFormat) attrs.push('autoFormat="1"');
    if (opts.intermediate) attrs.push('intermediate="1"');
    attrs.push(`connectionId="${opts.connectionId}"`);

    let inner = "";
    const r = opts.queryTableRefresh;
    if (r) {
      const rAttrs: string[] = [];
      if (r.preserveSortFilterLayout === false) rAttrs.push('preserveSortFilterLayout="0"');
      if (r.fieldIdWrapped) rAttrs.push('fieldIdWrapped="1"');
      if (r.headersInLastRefresh === false) rAttrs.push('headersInLastRefresh="0"');
      if (r.minimumVersion !== undefined) rAttrs.push(`minimumVersion="${r.minimumVersion}"`);
      if (r.nextId !== undefined) rAttrs.push(`nextId="${r.nextId}"`);
      if (r.unboundColumnsLeft !== undefined)
        rAttrs.push(`unboundColumnsLeft="${r.unboundColumnsLeft}"`);
      if (r.unboundColumnsRight !== undefined)
        rAttrs.push(`unboundColumnsRight="${r.unboundColumnsRight}"`);
      const rInner: string[] = [];
      if (r.queryTableFields && r.queryTableFields.length > 0) {
        const fParts: string[] = [`<queryTableFields count="${r.queryTableFields.length}">`];
        for (const f of r.queryTableFields) {
          const fAttrs: string[] = [`id="${f.id}"`];
          if (f.name !== undefined) fAttrs.push(`name="${escapeXml(f.name)}"`);
          if (f.dataBound === false) fAttrs.push('dataBound="0"');
          if (f.rowNumbers) fAttrs.push('rowNumbers="1"');
          if (f.fillFormulas) fAttrs.push('fillFormulas="1"');
          if (f.clipped) fAttrs.push('clipped="1"');
          if (f.tableColumnId !== undefined) fAttrs.push(`tableColumnId="${f.tableColumnId}"`);
          fParts.push(`<queryTableField ${fAttrs.join(" ")}/>`);
        }
        fParts.push("</queryTableFields>");
        rInner.push(fParts.join(""));
      }
      if (r.deletedFields && r.deletedFields.length > 0) {
        const dParts: string[] = [`<queryTableDeletedFields count="${r.deletedFields.length}">`];
        for (const d of r.deletedFields) {
          dParts.push(`<deletedField name="${escapeXml(d.name)}"/>`);
        }
        dParts.push("</queryTableDeletedFields>");
        rInner.push(dParts.join(""));
      }
      if (rInner.length > 0)
        inner = `<queryTableRefresh ${rAttrs.join(" ")}>${rInner.join("")}</queryTableRefresh>`;
      else inner = `<queryTableRefresh ${rAttrs.join(" ")}/>`;
    }

    return (
      `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${attrs.join(" ")}>` +
      `${inner}</queryTable>`
    );
  },

  parse(el, _ctx) {
    const result: Partial<QueryTableOptions> = {};
    if (attr(el, "name") !== undefined) result.name = attr(el, "name");
    if (String(attr(el, "headers")) === "0") result.headers = false;
    if (parseOnOff(attr(el, "rowNumbers"))) result.rowNumbers = true;
    if (parseOnOff(attr(el, "disableRefresh"))) result.disableRefresh = true;
    if (String(attr(el, "backgroundRefresh")) === "0") result.backgroundRefresh = false;
    if (parseOnOff(attr(el, "firstBackgroundRefresh"))) result.firstBackgroundRefresh = true;
    if (parseOnOff(attr(el, "refreshOnLoad"))) result.refreshOnLoad = true;
    const gst = attr(el, "growShrinkType");
    if (gst !== undefined) result.growShrinkType = gst as QueryTableOptions["growShrinkType"];
    if (parseOnOff(attr(el, "fillFormulas"))) result.fillFormulas = true;
    if (parseOnOff(attr(el, "removeDataOnSave"))) result.removeDataOnSave = true;
    if (parseOnOff(attr(el, "disableEdit"))) result.disableEdit = true;
    if (String(attr(el, "preserveFormatting")) === "0") result.preserveFormatting = false;
    if (parseOnOff(attr(el, "adjustColumnWidth"))) result.adjustColumnWidth = true;
    if (parseOnOff(attr(el, "autoFormat"))) result.autoFormat = true;
    if (parseOnOff(attr(el, "intermediate"))) result.intermediate = true;
    const cid = attrNum(el, "connectionId");
    if (cid !== undefined) result.connectionId = cid;

    const rEl = findChild(el, "queryTableRefresh");
    if (rEl) {
      const r: Partial<QueryTableRefreshOptions> = {};
      if (String(attr(rEl, "preserveSortFilterLayout")) === "0") r.preserveSortFilterLayout = false;
      if (parseOnOff(attr(rEl, "fieldIdWrapped"))) r.fieldIdWrapped = true;
      if (String(attr(rEl, "headersInLastRefresh")) === "0") r.headersInLastRefresh = false;
      const mv = attrNum(rEl, "minimumVersion");
      if (mv !== undefined) r.minimumVersion = mv;
      const nid = attrNum(rEl, "nextId");
      if (nid !== undefined) r.nextId = nid;
      const ucl = attrNum(rEl, "unboundColumnsLeft");
      if (ucl !== undefined) r.unboundColumnsLeft = ucl;
      const ucr = attrNum(rEl, "unboundColumnsRight");
      if (ucr !== undefined) r.unboundColumnsRight = ucr;
      const fieldsEl = findChild(rEl, "queryTableFields");
      if (fieldsEl) {
        const fields: QueryTableFieldOptions[] = [];
        for (const fEl of fieldsEl.elements ?? []) {
          if (fEl.name !== "queryTableField") continue;
          const f: QueryTableFieldOptions = { id: attrNum(fEl, "id") ?? 0 };
          if (attr(fEl, "name") !== undefined) f.name = attr(fEl, "name");
          if (String(attr(fEl, "dataBound")) === "0") f.dataBound = false;
          if (parseOnOff(attr(fEl, "rowNumbers"))) f.rowNumbers = true;
          if (parseOnOff(attr(fEl, "fillFormulas"))) f.fillFormulas = true;
          if (parseOnOff(attr(fEl, "clipped"))) f.clipped = true;
          const tci = attrNum(fEl, "tableColumnId");
          if (tci !== undefined) f.tableColumnId = tci;
          fields.push(f);
        }
        if (fields.length > 0) r.queryTableFields = fields;
      }
      const delEl = findChild(rEl, "queryTableDeletedFields");
      if (delEl) {
        const deleted: QueryTableDeletedFieldOptions[] = [];
        for (const dEl of delEl.elements ?? []) {
          if (dEl.name !== "deletedField") continue;
          deleted.push({ name: attr(dEl, "name") ?? "" });
        }
        if (deleted.length > 0) r.deletedFields = deleted;
      }
      result.queryTableRefresh = r as QueryTableRefreshOptions;
    }
    return result as QueryTableOptions;
  },
};
