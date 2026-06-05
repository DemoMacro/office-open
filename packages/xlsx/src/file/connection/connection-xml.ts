/**
 * Connection XML generator — produces xl/connections.xml.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Connections
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

// ── Options ──

export interface DbPrOptions {
  /** OLE DB connection string (required) */
  readonly connection: string;
  /** Command text */
  readonly command?: string;
  /** Command type: "1"=cube, "2"=SQL, "3"=table, "4"=default */
  readonly commandType?: number;
  /** Server formatting */
  readonly serverCommand?: string;
}

export interface WebPrOptions {
  /** URL (required) */
  readonly url: string;
  /** Source data from: "csv" | "html" | ... */
  readonly sourceData?: boolean;
  /** HTML formatting: "all" | "rtf" | "none" */
  readonly htmlFormat?: string;
  /** HTML tables (1-based indices) */
  readonly htmlTables?: readonly string[];
  /** Consecutive property */
  readonly consecutive?: boolean;
  /** First row header */
  readonly firstRowHeader?: boolean;
  /** Parse PRE tags */
  readonly parsePre?: boolean;
  /** Same row dates */
  readonly xl2000?: boolean;
  /** Edit page URL (CT_WebPr @editPage) */
  readonly editPage?: string;
  /** First row headers (CT_WebPr @firstRow) */
  readonly firstRow?: boolean;
  /** POST request (CT_WebPr @post) */
  readonly post?: boolean;
  /** Text dates (CT_WebPr @textDates) */
  readonly textDates?: boolean;
  /** XL97 format (CT_WebPr @xl97) */
  readonly xl97?: boolean;
  /** Text fields configuration */
  readonly textFields?: readonly TextFieldOptions[];
}

export interface TextPrOptions {
  /** Text fields */
  readonly textFields?: readonly TextFieldOptions[];
  /** Code page */
  readonly codePage?: number;
  /** Character set */
  readonly characterSet?: string;
  /** Source file */
  readonly sourceFile?: string;
  /** Delimited */
  readonly delimited?: boolean;
  /** Tab delimiter */
  readonly tab?: boolean;
  /** Space delimiter */
  readonly space?: boolean;
  /** Comma delimiter */
  readonly comma?: boolean;
  /** Semicolon delimiter */
  readonly semicolon?: boolean;
  /** Custom delimiter character */
  readonly custom?: string;
  /** Decimal character */
  readonly decimal?: string;
  /** Thousands separator */
  readonly thousands?: string;
  /** Trailing minus */
  readonly trailingMinus?: boolean;
  /** Delimiter character (CT_TextPr @delimiter) */
  readonly delimiter?: string;
  /** File type (CT_TextPr @fileType) */
  readonly fileType?: "mac" | "win" | "dos";
  /** First row headers (CT_TextPr @firstRow) */
  readonly firstRow?: boolean;
  /** Qualifier (CT_TextPr @qualifier) */
  readonly qualifier?: "doubleQuote" | "singleQuote" | "none";
}

export interface TextFieldOptions {
  /** Column index (1-based) */
  readonly type: number;
  /** Field data type: "auto" | "text" | "MDY" | "DMY" | "YMD" | "MYD" | "DYM" | "YDM" | "skip" | "ELAPSED" | "SYS_DATE" | ... */
  readonly dataType?: string;
}

export interface ParameterOptions {
  /** Parameter name (required) */
  readonly name: string;
  /** SQL data type */
  readonly sqlType?: number;
  /** Character set */
  readonly characterSet?: string;
  /** String value */
  readonly stringValue?: string;
  /** Integer value */
  readonly integerValue?: number;
  /** Boolean value */
  readonly booleanValue?: boolean;
  /** Refresh on load */
  readonly refreshOnChange?: boolean;
  /** Prompt user */
  readonly prompt?: boolean;
  /** Integer parameter (CT_Parameter @integer) */
  readonly integer?: boolean;
  /** Cell reference for parameter value */
  readonly reference?: string;
  /** Parameter type: "prompt" | "value" | "cell" */
  readonly parameterType?: string;
}

export interface ConnectionOptions {
  /** Unique connection ID (required) */
  readonly id: number;
  /** Connection name */
  readonly name?: string;
  /** Connection type: 1=ODBC, 2=DAO, 3=OLE DB, 4=web, 5=text, 6=ADO, 7=DSP */
  readonly type?: number;
  /** Refresh on load */
  readonly refreshOnLoad?: boolean;
  /** Refreshed version */
  readonly refreshedVersion?: number;
  /** Background refresh */
  readonly backgroundRefresh?: boolean;
  /** Save data */
  readonly saveData?: boolean;
  /** Save password */
  readonly savePassword?: boolean;
  /** Connection description */
  readonly description?: string;
  /** Credentials method: "integrated" | "none" | "stored" | "prompt" */
  readonly credentials?: string;
  /** Refresh interval (CT_Connection @interval) */
  readonly interval?: number;
  /** Keep alive (CT_Connection @keepAlive) */
  readonly keepAlive?: boolean;
  /** New connection (CT_Connection @new) */
  readonly new?: boolean;
  /** ODC file path (CT_Connection @odcFile) */
  readonly odcFile?: string;
  /** Only use connection file (CT_Connection @onlyUseConnectionFile) */
  readonly onlyUseConnectionFile?: boolean;
  /** Reconnection method (CT_Connection @reconnectionMethod) */
  readonly reconnectionMethod?: number;
  /** Single sign-on ID (CT_Connection @singleSignOnId) */
  readonly singleSignOnId?: string;
  /** OLE DB properties */
  readonly dbPr?: DbPrOptions;
  /** Web query properties */
  readonly webPr?: WebPrOptions;
  /** Text import properties */
  readonly textPr?: TextPrOptions;
  /** Parameters */
  readonly parameters?: readonly ParameterOptions[];
}

// ── Component ──

export class ConnectionsXml extends BaseXmlComponent {
  private readonly connections: readonly ConnectionOptions[];

  public constructor(connections: readonly ConnectionOptions[]) {
    super("connections");
    this.connections = connections;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];

    for (const c of this.connections) {
      const cAttrs: Record<string, string | number | boolean | undefined> = {
        id: c.id,
      };
      if (c.name !== undefined) cAttrs.name = c.name;
      if (c.type !== undefined) cAttrs.type = c.type;
      if (c.refreshOnLoad) cAttrs.refreshOnLoad = 1;
      if (c.refreshedVersion !== undefined) cAttrs.refreshedVersion = c.refreshedVersion;
      else cAttrs.refreshedVersion = 6;
      if (c.backgroundRefresh) cAttrs.backgroundRefresh = 1;
      if (c.saveData === false) cAttrs.saveData = 0;
      if (c.savePassword) cAttrs.savePassword = 1;
      if (c.description !== undefined) cAttrs.description = c.description;
      if (c.credentials !== undefined) cAttrs.credentials = c.credentials;
      if (c.interval !== undefined) cAttrs.interval = c.interval;
      if (c.keepAlive) cAttrs.keepAlive = 1;
      if (c.new) cAttrs.new = 1;
      if (c.odcFile !== undefined) cAttrs.odcFile = c.odcFile;
      if (c.onlyUseConnectionFile) cAttrs.onlyUseConnectionFile = 1;
      if (c.reconnectionMethod !== undefined) cAttrs.reconnectionMethod = c.reconnectionMethod;
      if (c.singleSignOnId !== undefined) cAttrs.singleSignOnId = c.singleSignOnId;

      const children: string[] = [];

      // dbPr
      if (c.dbPr) {
        const dbAttrs: Record<string, string | number | undefined> = {
          connection: c.dbPr.connection,
        };
        if (c.dbPr.command !== undefined) dbAttrs.command = c.dbPr.command;
        if (c.dbPr.commandType !== undefined) dbAttrs.commandType = c.dbPr.commandType;
        if (c.dbPr.serverCommand !== undefined) dbAttrs.serverCommand = c.dbPr.serverCommand;
        children.push(`<dbPr${attrs(dbAttrs)}/>`);
      }

      // webPr
      if (c.webPr) {
        const wpAttrs: Record<string, string | number | boolean | undefined> = {
          url: c.webPr.url,
        };
        if (c.webPr.sourceData) wpAttrs.sourceData = 1;
        if (c.webPr.htmlFormat !== undefined) wpAttrs.htmlFormat = c.webPr.htmlFormat;
        if (c.webPr.consecutive) wpAttrs.consecutive = 1;
        if (c.webPr.firstRowHeader) wpAttrs.firstRowHeader = 1;
        if (c.webPr.parsePre) wpAttrs.parsePre = 1;
        if (c.webPr.xl2000) wpAttrs.xl2000 = 1;
        if (c.webPr.editPage !== undefined) wpAttrs.editPage = c.webPr.editPage;
        if (c.webPr.firstRow) wpAttrs.firstRow = 1;
        if (c.webPr.post) wpAttrs.post = 1;
        if (c.webPr.textDates) wpAttrs.textDates = 1;
        if (c.webPr.xl97) wpAttrs.xl97 = 1;

        const wpChildren: string[] = [];
        if (c.webPr.htmlTables && c.webPr.htmlTables.length > 0) {
          wpChildren.push(
            `<tables>${c.webPr.htmlTables.map((t) => `<x v="${escapeXml(t)}"/>`).join("")}</tables>`,
          );
        }
        if (c.webPr.textFields && c.webPr.textFields.length > 0) {
          wpChildren.push(buildTextFields(c.webPr.textFields));
        }
        if (wpChildren.length > 0) {
          children.push(`<webPr${attrs(wpAttrs)}>${wpChildren.join("")}</webPr>`);
        } else {
          children.push(`<webPr${attrs(wpAttrs)}/>`);
        }
      }

      // textPr
      if (c.textPr) {
        const tpAttrs: Record<string, string | number | boolean | undefined> = {};
        if (c.textPr.codePage !== undefined) tpAttrs.codePage = c.textPr.codePage;
        if (c.textPr.characterSet !== undefined) tpAttrs.characterSet = c.textPr.characterSet;
        if (c.textPr.sourceFile !== undefined) tpAttrs.sourceFile = c.textPr.sourceFile;
        if (c.textPr.delimited !== undefined) tpAttrs.delimited = c.textPr.delimited ? 1 : 0;
        if (c.textPr.tab !== undefined) tpAttrs.tab = c.textPr.tab ? 1 : 0;
        if (c.textPr.space !== undefined) tpAttrs.space = c.textPr.space ? 1 : 0;
        if (c.textPr.comma !== undefined) tpAttrs.comma = c.textPr.comma ? 1 : 0;
        if (c.textPr.semicolon !== undefined) tpAttrs.semicolon = c.textPr.semicolon ? 1 : 0;
        if (c.textPr.custom !== undefined) tpAttrs.custom = c.textPr.custom;
        if (c.textPr.decimal !== undefined) tpAttrs.decimal = c.textPr.decimal;
        if (c.textPr.thousands !== undefined) tpAttrs.thousands = c.textPr.thousands;
        if (c.textPr.trailingMinus !== undefined)
          tpAttrs.trailingMinus = c.textPr.trailingMinus ? 1 : 0;
        if (c.textPr.delimiter !== undefined) tpAttrs.delimiter = c.textPr.delimiter;
        if (c.textPr.fileType !== undefined) tpAttrs.fileType = c.textPr.fileType;
        if (c.textPr.firstRow) tpAttrs.firstRow = 1;
        if (c.textPr.qualifier !== undefined) tpAttrs.qualifier = c.textPr.qualifier;

        const tpChildren: string[] = [];
        if (c.textPr.textFields && c.textPr.textFields.length > 0) {
          tpChildren.push(buildTextFields(c.textPr.textFields));
        }
        if (tpChildren.length > 0) {
          children.push(`<textPr${attrs(tpAttrs)}>${tpChildren.join("")}</textPr>`);
        } else {
          children.push(`<textPr${attrs(tpAttrs)}/>`);
        }
      }

      // parameters
      if (c.parameters && c.parameters.length > 0) {
        const paramParts: string[] = [`<parameters count="${c.parameters.length}">`];
        for (const param of c.parameters) {
          const paramAttrs: Record<string, string | number | boolean | undefined> = {
            name: param.name,
          };
          if (param.sqlType !== undefined) paramAttrs.sqlType = param.sqlType;
          if (param.characterSet !== undefined) paramAttrs.characterSet = param.characterSet;
          if (param.stringValue !== undefined) paramAttrs.stringValue = param.stringValue;
          if (param.integerValue !== undefined) paramAttrs.integerValue = param.integerValue;
          if (param.booleanValue !== undefined)
            paramAttrs.booleanValue = param.booleanValue ? 1 : 0;
          if (param.refreshOnChange) paramAttrs.refreshOnChange = 1;
          if (param.prompt) paramAttrs.prompt = 1;
          if (param.integer) paramAttrs.integer = 1;
          if (param.reference !== undefined) paramAttrs.reference = param.reference;
          if (param.parameterType !== undefined) paramAttrs.parameterType = param.parameterType;
          paramParts.push(`<parameter${attrs(paramAttrs)}/>`);
        }
        paramParts.push("</parameters>");
        children.push(paramParts.join(""));
      }

      if (children.length > 0) {
        p.push(`<connection${attrs(cAttrs)}>${children.join("")}</connection>`);
      } else {
        p.push(`<connection${attrs(cAttrs)}/>`);
      }
    }

    p.push("</connections>");
    return p.join("");
  }
}

function buildTextFields(fields: readonly TextFieldOptions[]): string {
  const parts: string[] = [`<textFields count="${fields.length}">`];
  for (const f of fields) {
    const fAttrs: Record<string, string | number | undefined> = { type: f.type };
    if (f.dataType !== undefined) fAttrs.dataType = f.dataType;
    parts.push(`<textField${attrs(fAttrs)}/>`);
  }
  parts.push("</textFields>");
  return parts.join("");
}
