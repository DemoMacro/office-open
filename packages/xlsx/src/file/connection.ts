/**
 * Connection options for SpreadsheetML documents.
 *
 * Reference: OOXML transitional, sml.xsd, CT_Connections
 *
 * @module
 */

// ── Options ──

export interface DbPrOptions {
  /** OLE DB connection string (required) */
  connection: string;
  /** Command text */
  command?: string;
  /** Command type: "1"=cube, "2"=SQL, "3"=table, "4"=default */
  commandType?: number;
  /** Server formatting */
  serverCommand?: string;
}

export interface WebPrOptions {
  /** URL (required) */
  url: string;
  /** Source data from: "csv" | "html" | ... */
  sourceData?: boolean;
  /** HTML formatting: "all" | "rtf" | "none" */
  htmlFormat?: string;
  /** HTML tables (1-based indices) */
  htmlTables?: string[];
  /** Consecutive property */
  consecutive?: boolean;
  /** First row header */
  firstRowHeader?: boolean;
  /** Parse PRE tags */
  parsePre?: boolean;
  /** Same row dates */
  xl2000?: boolean;
  /** Edit page URL (CT_WebPr @editPage) */
  editPage?: string;
  /** First row headers (CT_WebPr @firstRow) */
  firstRow?: boolean;
  /** POST request (CT_WebPr @post) */
  post?: boolean;
  /** Text dates (CT_WebPr @textDates) */
  textDates?: boolean;
  /** XL97 format (CT_WebPr @xl97) */
  xl97?: boolean;
  /** Text fields configuration */
  textFields?: TextFieldOptions[];
}

export interface TextPrOptions {
  /** Text fields */
  textFields?: TextFieldOptions[];
  /** Code page */
  codePage?: number;
  /** Character set */
  characterSet?: string;
  /** Source file */
  sourceFile?: string;
  /** Delimited */
  delimited?: boolean;
  /** Tab delimiter */
  tab?: boolean;
  /** Space delimiter */
  space?: boolean;
  /** Comma delimiter */
  comma?: boolean;
  /** Semicolon delimiter */
  semicolon?: boolean;
  /** Custom delimiter character */
  custom?: string;
  /** Decimal character */
  decimal?: string;
  /** Thousands separator */
  thousands?: string;
  /** Trailing minus */
  trailingMinus?: boolean;
  /** Delimiter character (CT_TextPr @delimiter) */
  delimiter?: string;
  /** File type (CT_TextPr @fileType) */
  fileType?: "mac" | "win" | "dos";
  /** First row headers (CT_TextPr @firstRow) */
  firstRow?: boolean;
  /** Qualifier (CT_TextPr @qualifier) */
  qualifier?: "doubleQuote" | "singleQuote" | "none";
}

export interface TextFieldOptions {
  /** Column index (1-based) */
  type: number;
  /** Field data type: "auto" | "text" | "MDY" | "DMY" | "YMD" | "MYD" | "DYM" | "YDM" | "skip" | "ELAPSED" | "SYS_DATE" | ... */
  dataType?: string;
}

export interface ParameterOptions {
  /** Parameter name (required) */
  name: string;
  /** SQL data type */
  sqlType?: number;
  /** Character set */
  characterSet?: string;
  /** String value */
  stringValue?: string;
  /** Integer value */
  integerValue?: number;
  /** Boolean value */
  booleanValue?: boolean;
  /** Refresh on load */
  refreshOnChange?: boolean;
  /** Prompt user */
  prompt?: boolean;
  /** Integer parameter (CT_Parameter @integer) */
  integer?: boolean;
  /** Cell reference for parameter value */
  reference?: string;
  /** Parameter type: "prompt" | "value" | "cell" */
  parameterType?: string;
}

export interface ConnectionOptions {
  /** Unique connection ID (required) */
  id: number;
  /** Connection name */
  name?: string;
  /** Connection type: 1=ODBC, 2=DAO, 3=OLE DB, 4=web, 5=text, 6=ADO, 7=DSP */
  type?: number;
  /** Refresh on load */
  refreshOnLoad?: boolean;
  /** Refreshed version */
  refreshedVersion?: number;
  /** Background refresh */
  backgroundRefresh?: boolean;
  /** Save data */
  saveData?: boolean;
  /** Save password */
  savePassword?: boolean;
  /** Connection description */
  description?: string;
  /** Credentials method: "integrated" | "none" | "stored" | "prompt" */
  credentials?: string;
  /** Refresh interval (CT_Connection @interval) */
  interval?: number;
  /** Keep alive (CT_Connection @keepAlive) */
  keepAlive?: boolean;
  /** New connection (CT_Connection @new) */
  new?: boolean;
  /** ODC file path (CT_Connection @odcFile) */
  odcFile?: string;
  /** Only use connection file (CT_Connection @onlyUseConnectionFile) */
  onlyUseConnectionFile?: boolean;
  /** Reconnection method (CT_Connection @reconnectionMethod) */
  reconnectionMethod?: number;
  /** Single sign-on ID (CT_Connection @singleSignOnId) */
  singleSignOnId?: string;
  /** OLE DB properties */
  dbPr?: DbPrOptions;
  /** Web query properties */
  webPr?: WebPrOptions;
  /** Text import properties */
  textPr?: TextPrOptions;
  /** Parameters */
  parameters?: ParameterOptions[];
}
