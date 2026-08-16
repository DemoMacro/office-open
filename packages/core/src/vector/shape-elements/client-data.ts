/**
 * Spreadsheet drawing x:ClientData — the form-control/comment anchor payload
 * every xlsx VML shape carries. One required ObjectType attribute plus a
 * repeating choice of 67 leaf child elements, each holding a single text
 * value (TrueFalseBlank flag, integer, or string).
 *
 * Field names keep the XML PascalCase spelling — they mirror the child
 * elements one-to-one (same decision as o:OLEObject's Type/ProgID).
 *
 * Reference: ISO/IEC 29500-4, vml-spreadsheetDrawing.xsd, CT_ClientData.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { VmlTrueFalseBlank } from "../attributes";
import { parseVmlTrueFalseBlank, stringifyVmlTrueFalseBlank } from "./office-elements";

/** ST_ObjectType — the legacy form-control kinds. */
export type VmlClientObjectType =
  | "Button"
  | "Checkbox"
  | "Dialog"
  | "Drop"
  | "Edit"
  | "GBox"
  | "Label"
  | "LineA"
  | "List"
  | "Movie"
  | "Note"
  | "Pict"
  | "Radio"
  | "RectA"
  | "Scroll"
  | "Spin"
  | "Shape"
  | "Group"
  | "Rect";

/** x:ClientData options (CT_ClientData). */
export interface VmlClientDataOptions {
  objectType: VmlClientObjectType;
  // ST_TrueFalseBlank children
  MoveWithCells?: VmlTrueFalseBlank;
  SizeWithCells?: VmlTrueFalseBlank;
  Locked?: VmlTrueFalseBlank;
  DefaultSize?: VmlTrueFalseBlank;
  PrintObject?: VmlTrueFalseBlank;
  Disabled?: VmlTrueFalseBlank;
  AutoFill?: VmlTrueFalseBlank;
  AutoLine?: VmlTrueFalseBlank;
  AutoPict?: VmlTrueFalseBlank;
  LockText?: VmlTrueFalseBlank;
  JustLastX?: VmlTrueFalseBlank;
  SecretEdit?: VmlTrueFalseBlank;
  Default?: VmlTrueFalseBlank;
  Help?: VmlTrueFalseBlank;
  Cancel?: VmlTrueFalseBlank;
  Dismiss?: VmlTrueFalseBlank;
  Visible?: VmlTrueFalseBlank;
  RowHidden?: VmlTrueFalseBlank;
  ColHidden?: VmlTrueFalseBlank;
  MultiLine?: VmlTrueFalseBlank;
  VScroll?: VmlTrueFalseBlank;
  ValidIds?: VmlTrueFalseBlank;
  NoThreeD2?: VmlTrueFalseBlank;
  Colored?: VmlTrueFalseBlank;
  NoThreeD?: VmlTrueFalseBlank;
  FirstButton?: VmlTrueFalseBlank;
  Horiz?: VmlTrueFalseBlank;
  MapOCX?: VmlTrueFalseBlank;
  Camera?: VmlTrueFalseBlank;
  RecalcAlways?: VmlTrueFalseBlank;
  AutoScale?: VmlTrueFalseBlank;
  DDE?: VmlTrueFalseBlank;
  UIObj?: VmlTrueFalseBlank;
  // integer children
  Accel?: number;
  Accel2?: number;
  Row?: number;
  Column?: number;
  VTEdit?: number;
  WidthMin?: number;
  Sel?: number;
  DropLines?: number;
  Checked?: number;
  Val?: number;
  Min?: number;
  Max?: number;
  Inc?: number;
  Page?: number;
  Dx?: number;
  ScriptLanguage?: number;
  ScriptLocation?: number;
  // string children
  /** Cell anchor, e.g. "\n_Frmd4c1s387\n1\n15\n90\n1\n12\n65\n85". */
  Anchor?: string;
  FmlaMacro?: string;
  TextHAlign?: string;
  TextVAlign?: string;
  FmlaRange?: string;
  SelType?: string;
  MultiSel?: string;
  LCT?: string;
  ListItem?: string;
  DropStyle?: string;
  FmlaLink?: string;
  FmlaPict?: string;
  FmlaGroup?: string;
  ScriptText?: string;
  ScriptExtended?: string;
  FmlaTxbx?: string;
  /** CF — clipboard format name. */
  CF?: string;
}

const TRUE_FALSE_BLANK_FIELDS = [
  "MoveWithCells",
  "SizeWithCells",
  "Locked",
  "DefaultSize",
  "PrintObject",
  "Disabled",
  "AutoFill",
  "AutoLine",
  "AutoPict",
  "LockText",
  "JustLastX",
  "SecretEdit",
  "Default",
  "Help",
  "Cancel",
  "Dismiss",
  "Visible",
  "RowHidden",
  "ColHidden",
  "MultiLine",
  "VScroll",
  "ValidIds",
  "NoThreeD2",
  "Colored",
  "NoThreeD",
  "FirstButton",
  "Horiz",
  "MapOCX",
  "Camera",
  "RecalcAlways",
  "AutoScale",
  "DDE",
  "UIObj",
] as const;

const INTEGER_FIELDS = [
  "Accel",
  "Accel2",
  "Row",
  "Column",
  "VTEdit",
  "WidthMin",
  "Sel",
  "DropLines",
  "Checked",
  "Val",
  "Min",
  "Max",
  "Inc",
  "Page",
  "Dx",
  "ScriptLanguage",
  "ScriptLocation",
] as const;

const STRING_FIELDS = [
  "Anchor",
  "FmlaMacro",
  "TextHAlign",
  "TextVAlign",
  "FmlaRange",
  "SelType",
  "MultiSel",
  "LCT",
  "ListItem",
  "DropStyle",
  "FmlaLink",
  "FmlaPict",
  "FmlaGroup",
  "ScriptText",
  "ScriptExtended",
  "FmlaTxbx",
  "CF",
] as const;

const TRUE_FALSE_BLANK_FIELD_SET = new Set<string>(TRUE_FALSE_BLANK_FIELDS);
const INTEGER_FIELD_SET = new Set<string>(INTEGER_FIELDS);
const STRING_FIELD_SET = new Set<string>(STRING_FIELDS);

/** Serialize x:ClientData. */
export function stringifyVmlClientData(opts: VmlClientDataOptions): string {
  const children: string[] = [];
  for (const field of TRUE_FALSE_BLANK_FIELDS) {
    const value = opts[field];
    if (value !== undefined) {
      children.push(`<x:${field}>${stringifyVmlTrueFalseBlank(value)}</x:${field}>`);
    }
  }
  for (const field of INTEGER_FIELDS) {
    const value = opts[field];
    if (value !== undefined) children.push(`<x:${field}>${value}</x:${field}>`);
  }
  for (const field of STRING_FIELDS) {
    const value = opts[field];
    if (value !== undefined) children.push(`<x:${field}>${escapeXml(value)}</x:${field}>`);
  }
  return `<x:ClientData ObjectType="${opts.objectType}">${children.join("")}</x:ClientData>`;
}

/** Parse an x:ClientData element. */
export function parseVmlClientData(el: XmlElement): VmlClientDataOptions {
  const out: VmlClientDataOptions = {
    objectType: String(el.attributes?.ObjectType ?? "") as VmlClientObjectType,
  };
  const extras: Record<string, unknown> = {};
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    const local = child.name?.startsWith("x:") ? child.name.slice(2) : null;
    if (local === null) continue;
    const text = (child.elements ?? []).map((c) => String(c.text ?? "")).join("");
    if (TRUE_FALSE_BLANK_FIELD_SET.has(local)) {
      extras[local] = parseVmlTrueFalseBlank(text);
    } else if (INTEGER_FIELD_SET.has(local)) {
      extras[local] = Number(text);
    } else if (STRING_FIELD_SET.has(local)) {
      extras[local] = text;
    }
  }
  return Object.assign(out, extras);
}
