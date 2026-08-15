/**
 * Connection types and descriptor for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Connections
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml, findChild } from "@office-open/xml";

// ── Options ──

/** Database (OLE DB) connection properties (CT_DbPr). */
export interface DatabasePropertiesOptions {
  /** OLE DB connection string (required) */
  connection: string;
  /** Command text */
  command?: string;
  /** Server command text (CT_DbPr @serverCommand) */
  serverCommand?: string;
  /** Command type: 1=cube, 2=SQL, 3=table (default 2) */
  commandType?: number;
}

/** One web-query table selection (CT_Tables choice: m/s/x). */
export type WebTableSelection = string | number | null;

/** Web query properties (CT_WebPr). */
export interface WebPropertiesOptions {
  /** Source URL (CT_WebPr @url) */
  url?: string;
  /** XML source query (CT_WebPr @xml) */
  xml?: boolean;
  /** Source data included (CT_WebPr @sourceData) */
  sourceData?: boolean;
  /** Parse PRE tags into columns (CT_WebPr @parsePre) */
  parsePre?: boolean;
  /** Consecutive delimiters as one (CT_WebPr @consecutive) */
  consecutive?: boolean;
  /** First row contains headers (CT_WebPr @firstRow) */
  firstRow?: boolean;
  /** XL97 compatible (CT_WebPr @xl97) */
  xl97?: boolean;
  /** Dates as text (CT_WebPr @textDates) */
  textDates?: boolean;
  /** XL2000 compatible (CT_WebPr @xl2000) */
  xl2000?: boolean;
  /** POST request body (CT_WebPr @post) */
  post?: string;
  /** HTML tables only (CT_WebPr @htmlTables) */
  htmlTables?: boolean;
  /** HTML formatting: "all" | "rtf" | "none" (default "none") */
  htmlFormat?: string;
  /** Edit page URL (CT_WebPr @editPage) */
  editPage?: string;
  /** Selected tables (CT_Tables: string name, number index, null = all/missing) */
  tables?: WebTableSelection[];
}

/** Text import field (CT_TextField). */
export interface TextFieldOptions {
  /** Field data type (ST_ExternalConnectionType, default "general") */
  type?: string;
  /** Field position (default 0) */
  position?: number;
}

/** Text import properties (CT_TextPr). */
export interface TextPropertiesOptions {
  /** Prompt for file name (default true) */
  prompt?: boolean;
  /** File type: "mac" | "win" | "dos" (default "win") */
  fileType?: "mac" | "win" | "dos";
  /** Code page (default 1252) */
  codePage?: number;
  /** Character set */
  characterSet?: string;
  /** First row to import (default 1) */
  firstRow?: number;
  /** Source file path */
  sourceFile?: string;
  /** Tab-delimited (default true) */
  delimited?: boolean;
  /** Decimal separator (default ".") */
  decimal?: string;
  /** Thousands separator (default ",") */
  thousands?: string;
  /** Tab delimiter (default true) */
  tab?: boolean;
  /** Space delimiter (default false) */
  space?: boolean;
  /** Comma delimiter (default false) */
  comma?: boolean;
  /** Semicolon delimiter (default false) */
  semicolon?: boolean;
  /** Consecutive delimiters as one (default false) */
  consecutive?: boolean;
  /** Text qualifier: "doubleQuote" | "singleQuote" | "none" */
  qualifier?: "doubleQuote" | "singleQuote" | "none";
  /** Field layout (CT_TextFields) */
  textFields?: TextFieldOptions[];
}

/** Query parameter (CT_Parameter). */
export interface ParameterOptions {
  /** Parameter name */
  name?: string;
  /** SQL data type (default 0) */
  sqlType?: number;
  /** Parameter type: "prompt" | "value" | "cell" (default "prompt") */
  parameterType?: "prompt" | "value" | "cell";
  /** Refresh when the parameter value changes (default false) */
  refreshOnChange?: boolean;
  /** Prompt text */
  prompt?: string;
  /** Boolean value */
  boolean?: boolean;
  /** Numeric value */
  double?: number;
  /** Integer value */
  integer?: number;
  /** String value */
  string?: string;
  /** Cell reference for the value */
  cell?: string;
}

/** Workbook connection (CT_Connection). */
export interface ConnectionOptions {
  /** Unique connection ID (required) */
  id: number;
  /** Connection name */
  name?: string;
  /** Connection type: 1=ODBC, 2=DAO, 3=OLE DB, 4=web, 5=text, 6=ADO, 7=DSP */
  type?: number;
  /** Refreshed version (required by XSD, default 6 at stringify) */
  refreshedVersion?: number;
  /** Minimum refreshable version (default 0) */
  minRefreshableVersion?: number;
  /** Refresh on load (default false) */
  refreshOnLoad?: boolean;
  /** Background refresh (default false) */
  backgroundRefresh?: boolean;
  /** Save data with the workbook (default false) */
  saveData?: boolean;
  /** Save password (default false) */
  savePassword?: boolean;
  /** Connection description */
  description?: string;
  /** Credentials method: "integrated" | "none" | "stored" | "prompt" (default "integrated") */
  credentials?: string;
  /** Refresh interval in minutes (default 0) */
  interval?: number;
  /** Keep connection alive (default false) */
  keepAlive?: boolean;
  /** Newly added connection (default false) */
  new?: boolean;
  /** Connection deleted (default false) */
  deleted?: boolean;
  /** Source file path (CT_Connection @sourceFile) */
  sourceFile?: string;
  /** ODC file path (CT_Connection @odcFile) */
  odcFile?: string;
  /** Only use the connection file (default false) */
  onlyUseConnectionFile?: boolean;
  /** Reconnection method (default 1) */
  reconnectionMethod?: number;
  /** Single sign-on ID (CT_Connection @singleSignOnId) */
  singleSignOnId?: string;
  /** Database properties (CT_DbPr) */
  dbPr?: DatabasePropertiesOptions;
  /** Web query properties (CT_WebPr) */
  webPr?: WebPropertiesOptions;
  /** Text import properties (CT_TextPr) */
  textPr?: TextPropertiesOptions;
  /** Parameters (CT_Parameters) */
  parameters?: ParameterOptions[];
}

/** Options for xl/connections.xml (CT_Connections). */
export interface ConnectionsOptions {
  /** Workbook connections */
  connections: ConnectionOptions[];
}

// ── Descriptor ──

export const connectionsDesc: CustomDescriptor<ConnectionsOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];
    for (const c of opts.connections) {
      const cAttrs: string[] = [`id="${c.id}"`];
      if (c.sourceFile !== undefined) cAttrs.push(`sourceFile="${escapeXml(c.sourceFile)}"`);
      if (c.odcFile !== undefined) cAttrs.push(`odcFile="${escapeXml(c.odcFile)}"`);
      if (c.keepAlive) cAttrs.push('keepAlive="1"');
      if (c.interval !== undefined) cAttrs.push(`interval="${c.interval}"`);
      if (c.name !== undefined) cAttrs.push(`name="${escapeXml(c.name)}"`);
      if (c.description !== undefined) cAttrs.push(`description="${escapeXml(c.description)}"`);
      if (c.type !== undefined) cAttrs.push(`type="${c.type}"`);
      if (c.reconnectionMethod !== undefined)
        cAttrs.push(`reconnectionMethod="${c.reconnectionMethod}"`);
      cAttrs.push(`refreshedVersion="${c.refreshedVersion ?? 6}"`);
      if (c.minRefreshableVersion !== undefined)
        cAttrs.push(`minRefreshableVersion="${c.minRefreshableVersion}"`);
      if (c.savePassword) cAttrs.push('savePassword="1"');
      if (c.new) cAttrs.push('new="1"');
      if (c.deleted) cAttrs.push('deleted="1"');
      if (c.onlyUseConnectionFile) cAttrs.push('onlyUseConnectionFile="1"');
      if (c.backgroundRefresh) cAttrs.push('background="1"');
      if (c.refreshOnLoad) cAttrs.push('refreshOnLoad="1"');
      if (c.saveData) cAttrs.push('saveData="1"');
      if (c.credentials !== undefined) cAttrs.push(`credentials="${escapeXml(c.credentials)}"`);
      if (c.singleSignOnId !== undefined)
        cAttrs.push(`singleSignOnId="${escapeXml(c.singleSignOnId)}"`);

      const inner: string[] = [];
      if (c.dbPr) {
        const dAttrs: string[] = [`connection="${escapeXml(c.dbPr.connection)}"`];
        if (c.dbPr.command !== undefined) dAttrs.push(`command="${escapeXml(c.dbPr.command)}"`);
        if (c.dbPr.serverCommand !== undefined)
          dAttrs.push(`serverCommand="${escapeXml(c.dbPr.serverCommand)}"`);
        if (c.dbPr.commandType !== undefined) dAttrs.push(`commandType="${c.dbPr.commandType}"`);
        inner.push(`<dbPr ${dAttrs.join(" ")}/>`);
      }
      if (c.webPr) {
        const w = c.webPr;
        const wAttrs: string[] = [];
        if (w.xml) wAttrs.push('xml="1"');
        if (w.sourceData) wAttrs.push('sourceData="1"');
        if (w.parsePre) wAttrs.push('parsePre="1"');
        if (w.consecutive) wAttrs.push('consecutive="1"');
        if (w.firstRow) wAttrs.push('firstRow="1"');
        if (w.xl97) wAttrs.push('xl97="1"');
        if (w.textDates) wAttrs.push('textDates="1"');
        if (w.xl2000) wAttrs.push('xl2000="1"');
        if (w.url !== undefined) wAttrs.push(`url="${escapeXml(w.url)}"`);
        if (w.post !== undefined) wAttrs.push(`post="${escapeXml(w.post)}"`);
        if (w.htmlTables) wAttrs.push('htmlTables="1"');
        if (w.htmlFormat !== undefined) wAttrs.push(`htmlFormat="${escapeXml(w.htmlFormat)}"`);
        if (w.editPage !== undefined) wAttrs.push(`editPage="${escapeXml(w.editPage)}"`);
        const tablesXml =
          w.tables && w.tables.length > 0
            ? `<tables count="${w.tables.length}">${w.tables
                .map((t) =>
                  t === null
                    ? "<m/>"
                    : typeof t === "number"
                      ? `<x v="${t}"/>`
                      : `<s v="${escapeXml(t)}"/>`,
                )
                .join("")}</tables>`
            : "";
        if (tablesXml) inner.push(`<webPr ${wAttrs.join(" ")}>${tablesXml}</webPr>`);
        else inner.push(`<webPr ${wAttrs.join(" ")}/>`);
      }
      if (c.textPr) {
        const t = c.textPr;
        const tAttrs: string[] = [];
        if (t.prompt === false) tAttrs.push('prompt="0"');
        if (t.fileType !== undefined) tAttrs.push(`fileType="${t.fileType}"`);
        if (t.codePage !== undefined) tAttrs.push(`codePage="${t.codePage}"`);
        if (t.characterSet !== undefined)
          tAttrs.push(`characterSet="${escapeXml(t.characterSet)}"`);
        if (t.firstRow !== undefined) tAttrs.push(`firstRow="${t.firstRow}"`);
        if (t.sourceFile !== undefined) tAttrs.push(`sourceFile="${escapeXml(t.sourceFile)}"`);
        if (t.delimited === false) tAttrs.push('delimited="0"');
        if (t.decimal !== undefined) tAttrs.push(`decimal="${escapeXml(t.decimal)}"`);
        if (t.thousands !== undefined) tAttrs.push(`thousands="${escapeXml(t.thousands)}"`);
        if (t.tab === false) tAttrs.push('tab="0"');
        if (t.space) tAttrs.push('space="1"');
        if (t.comma) tAttrs.push('comma="1"');
        if (t.semicolon) tAttrs.push('semicolon="1"');
        if (t.consecutive) tAttrs.push('consecutive="1"');
        if (t.qualifier !== undefined) tAttrs.push(`qualifier="${t.qualifier}"`);
        const fieldsXml =
          t.textFields && t.textFields.length > 0
            ? `<textFields count="${t.textFields.length}">${t.textFields
                .map((f) => {
                  const fAttrs: string[] = [];
                  if (f.type !== undefined) fAttrs.push(`type="${escapeXml(f.type)}"`);
                  if (f.position !== undefined) fAttrs.push(`position="${f.position}"`);
                  return `<textField ${fAttrs.join(" ")}/>`;
                })
                .join("")}</textFields>`
            : "";
        if (fieldsXml) inner.push(`<textPr ${tAttrs.join(" ")}>${fieldsXml}</textPr>`);
        else inner.push(`<textPr ${tAttrs.join(" ")}/>`);
      }
      if (c.parameters && c.parameters.length > 0) {
        const pmParts: string[] = [`<parameters count="${c.parameters.length}">`];
        for (const pm of c.parameters) {
          const pmAttrs: string[] = [];
          if (pm.name !== undefined) pmAttrs.push(`name="${escapeXml(pm.name)}"`);
          if (pm.sqlType !== undefined) pmAttrs.push(`sqlType="${pm.sqlType}"`);
          if (pm.parameterType !== undefined) pmAttrs.push(`parameterType="${pm.parameterType}"`);
          if (pm.refreshOnChange) pmAttrs.push('refreshOnChange="1"');
          if (pm.prompt !== undefined) pmAttrs.push(`prompt="${escapeXml(pm.prompt)}"`);
          if (pm.boolean !== undefined) pmAttrs.push(`boolean="${pm.boolean ? 1 : 0}"`);
          if (pm.double !== undefined) pmAttrs.push(`double="${pm.double}"`);
          if (pm.integer !== undefined) pmAttrs.push(`integer="${pm.integer}"`);
          if (pm.string !== undefined) pmAttrs.push(`string="${escapeXml(pm.string)}"`);
          if (pm.cell !== undefined) pmAttrs.push(`cell="${escapeXml(pm.cell)}"`);
          pmParts.push(`<parameter ${pmAttrs.join(" ")}/>`);
        }
        pmParts.push("</parameters>");
        inner.push(pmParts.join(""));
      }

      if (inner.length > 0)
        p.push(`<connection ${cAttrs.join(" ")}>${inner.join("")}</connection>`);
      else p.push(`<connection ${cAttrs.join(" ")}/>`);
    }
    p.push("</connections>");
    return p.join("");
  },

  parse(el, _ctx) {
    const connections: ConnectionOptions[] = [];
    for (const cEl of el.elements ?? []) {
      if (cEl.name !== "connection") continue;
      const c: Partial<ConnectionOptions> = { id: attrNum(cEl, "id") ?? 0 };
      if (attr(cEl, "sourceFile") !== undefined) c.sourceFile = attr(cEl, "sourceFile");
      if (attr(cEl, "odcFile") !== undefined) c.odcFile = attr(cEl, "odcFile");
      if (parseOnOff(attr(cEl, "keepAlive"))) c.keepAlive = true;
      const interval = attrNum(cEl, "interval");
      if (interval !== undefined) c.interval = interval;
      if (attr(cEl, "name") !== undefined) c.name = attr(cEl, "name");
      if (attr(cEl, "description") !== undefined) c.description = attr(cEl, "description");
      const type = attrNum(cEl, "type");
      if (type !== undefined) c.type = type;
      const rm = attrNum(cEl, "reconnectionMethod");
      if (rm !== undefined) c.reconnectionMethod = rm;
      const rv = attrNum(cEl, "refreshedVersion");
      if (rv !== undefined) c.refreshedVersion = rv;
      const mrv = attrNum(cEl, "minRefreshableVersion");
      if (mrv !== undefined) c.minRefreshableVersion = mrv;
      if (parseOnOff(attr(cEl, "savePassword"))) c.savePassword = true;
      if (parseOnOff(attr(cEl, "new"))) c.new = true;
      if (parseOnOff(attr(cEl, "deleted"))) c.deleted = true;
      if (parseOnOff(attr(cEl, "onlyUseConnectionFile"))) c.onlyUseConnectionFile = true;
      if (parseOnOff(attr(cEl, "background"))) c.backgroundRefresh = true;
      if (parseOnOff(attr(cEl, "refreshOnLoad"))) c.refreshOnLoad = true;
      if (parseOnOff(attr(cEl, "saveData"))) c.saveData = true;
      if (attr(cEl, "credentials") !== undefined) c.credentials = attr(cEl, "credentials");
      if (attr(cEl, "singleSignOnId") !== undefined) c.singleSignOnId = attr(cEl, "singleSignOnId");

      const dbEl = findChild(cEl, "dbPr");
      if (dbEl) {
        c.dbPr = {
          connection: attr(dbEl, "connection") ?? "",
          command: attr(dbEl, "command"),
          serverCommand: attr(dbEl, "serverCommand"),
          commandType: attrNum(dbEl, "commandType"),
        };
      }
      const webEl = findChild(cEl, "webPr");
      if (webEl) {
        const w: Partial<WebPropertiesOptions> = {};
        if (parseOnOff(attr(webEl, "xml"))) w.xml = true;
        if (parseOnOff(attr(webEl, "sourceData"))) w.sourceData = true;
        if (parseOnOff(attr(webEl, "parsePre"))) w.parsePre = true;
        if (parseOnOff(attr(webEl, "consecutive"))) w.consecutive = true;
        if (parseOnOff(attr(webEl, "firstRow"))) w.firstRow = true;
        if (parseOnOff(attr(webEl, "xl97"))) w.xl97 = true;
        if (parseOnOff(attr(webEl, "textDates"))) w.textDates = true;
        if (parseOnOff(attr(webEl, "xl2000"))) w.xl2000 = true;
        if (attr(webEl, "url") !== undefined) w.url = attr(webEl, "url");
        if (attr(webEl, "post") !== undefined) w.post = attr(webEl, "post");
        if (parseOnOff(attr(webEl, "htmlTables"))) w.htmlTables = true;
        if (attr(webEl, "htmlFormat") !== undefined) w.htmlFormat = attr(webEl, "htmlFormat");
        if (attr(webEl, "editPage") !== undefined) w.editPage = attr(webEl, "editPage");
        const tablesEl = findChild(webEl, "tables");
        if (tablesEl) {
          const tables: WebTableSelection[] = [];
          for (const tEl of tablesEl.elements ?? []) {
            if (tEl.name === "m") tables.push(null);
            else if (tEl.name === "s") tables.push(attr(tEl, "v") ?? "");
            else if (tEl.name === "x") tables.push(attrNum(tEl, "v") ?? 0);
          }
          if (tables.length > 0) w.tables = tables;
        }
        c.webPr = w as WebPropertiesOptions;
      }
      const textEl = findChild(cEl, "textPr");
      if (textEl) {
        const t: Partial<TextPropertiesOptions> = {};
        if (String(attr(textEl, "prompt")) === "0") t.prompt = false;
        const ft = attr(textEl, "fileType");
        if (ft !== undefined) t.fileType = ft as TextPropertiesOptions["fileType"];
        const cp = attrNum(textEl, "codePage");
        if (cp !== undefined) t.codePage = cp;
        if (attr(textEl, "characterSet") !== undefined)
          t.characterSet = attr(textEl, "characterSet");
        const fr = attrNum(textEl, "firstRow");
        if (fr !== undefined) t.firstRow = fr;
        if (attr(textEl, "sourceFile") !== undefined) t.sourceFile = attr(textEl, "sourceFile");
        if (String(attr(textEl, "delimited")) === "0") t.delimited = false;
        if (attr(textEl, "decimal") !== undefined) t.decimal = attr(textEl, "decimal");
        if (attr(textEl, "thousands") !== undefined) t.thousands = attr(textEl, "thousands");
        if (String(attr(textEl, "tab")) === "0") t.tab = false;
        if (parseOnOff(attr(textEl, "space"))) t.space = true;
        if (parseOnOff(attr(textEl, "comma"))) t.comma = true;
        if (parseOnOff(attr(textEl, "semicolon"))) t.semicolon = true;
        if (parseOnOff(attr(textEl, "consecutive"))) t.consecutive = true;
        const q = attr(textEl, "qualifier");
        if (q !== undefined) t.qualifier = q as TextPropertiesOptions["qualifier"];
        const fieldsEl = findChild(textEl, "textFields");
        if (fieldsEl) {
          const fields: TextFieldOptions[] = [];
          for (const fEl of fieldsEl.elements ?? []) {
            if (fEl.name !== "textField") continue;
            const f: TextFieldOptions = {};
            if (attr(fEl, "type") !== undefined) f.type = attr(fEl, "type");
            const pos = attrNum(fEl, "position");
            if (pos !== undefined) f.position = pos;
            fields.push(f);
          }
          if (fields.length > 0) t.textFields = fields;
        }
        c.textPr = t as TextPropertiesOptions;
      }
      const paramsEl = findChild(cEl, "parameters");
      if (paramsEl) {
        const params: ParameterOptions[] = [];
        for (const pmEl of paramsEl.elements ?? []) {
          if (pmEl.name !== "parameter") continue;
          const pm: ParameterOptions = {};
          if (attr(pmEl, "name") !== undefined) pm.name = attr(pmEl, "name");
          const st = attrNum(pmEl, "sqlType");
          if (st !== undefined) pm.sqlType = st;
          const pt = attr(pmEl, "parameterType");
          if (pt !== undefined) pm.parameterType = pt as ParameterOptions["parameterType"];
          if (parseOnOff(attr(pmEl, "refreshOnChange"))) pm.refreshOnChange = true;
          if (attr(pmEl, "prompt") !== undefined) pm.prompt = attr(pmEl, "prompt");
          if (attr(pmEl, "boolean") !== undefined) pm.boolean = parseOnOff(attr(pmEl, "boolean"));
          const dbl = attrNum(pmEl, "double");
          if (dbl !== undefined) pm.double = dbl;
          const int = attrNum(pmEl, "integer");
          if (int !== undefined) pm.integer = int;
          if (attr(pmEl, "string") !== undefined) pm.string = attr(pmEl, "string");
          if (attr(pmEl, "cell") !== undefined) pm.cell = attr(pmEl, "cell");
          params.push(pm);
        }
        if (params.length > 0) c.parameters = params;
      }
      connections.push(c as ConnectionOptions);
    }
    return { connections };
  },
};
