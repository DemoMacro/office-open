/**
 * XLSX patching — replace placeholders in existing .xlsx files.
 *
 * Unlike DOCX/PPTX (paragraph-based), XLSX patching targets cell values:
 * - String-valued patches substitute the placeholder text in place (shared
 *   strings table or inline strings).
 * - Non-string-valued patches (number / boolean / Date) whose placeholder is
 *   the entire cell content rewrite the cell to a typed value with the correct
 *   cell type (`t`), so the result is a real number/boolean/date cell — not a
 *   stringified copy.
 *
 * @module
 */
import {
  OoxmlMimeType,
  applyCorePropertiesOverride,
  strFromU8,
  toJson,
  unzipSync,
  zipAndConvert,
} from "@office-open/core";
import type {
  BasePatchOptions,
  CorePropertiesOptions,
  OutputByType,
  OutputType,
} from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { js2xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/** Scalar cell value accepted by {@link patchWorkbook}. */
export type ScalarValue = string | number | boolean | Date;

/** A patch value is just a scalar. */
export type Patch = ScalarValue;

export interface PatchWorkbookOptions<
  T extends OutputType = OutputType,
> extends BasePatchOptions<T> {
  /** Placeholder substitutions: `{{key}}` (per delimiters) → value. */
  placeholders?: Readonly<Record<string, ScalarValue>>;
  /** Literal find/replace: the find string → value (no delimiters added). */
  findReplace?: Readonly<Record<string, ScalarValue>>;
  /** Core-properties metadata override (merged over the existing docProps/core.xml). */
  coreProperties?: Partial<CorePropertiesOptions>;
}

// Excel serial-date epoch: 1899-12-30 accounts for the spurious 1900 leap year.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

const toExcelSerial = (date: Date): number => (date.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY;

/** Human-readable form for string-substitution (mixed-text) fallback. */
const toDisplayString = (value: Exclude<ScalarValue, string>): string => {
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value instanceof Date) return value.toISOString().split("T")[0];
  return String(value);
};

/**
 * Patch an existing .xlsx workbook by replacing placeholder text in cells.
 *
 * @publicApi
 */
export const patchWorkbook = async <T extends OutputType = OutputType>({
  outputType,
  data,
  placeholders,
  findReplace,
  coreProperties,
  placeholderDelimiters = { start: "{{", end: "}}" } as const,
}: PatchWorkbookOptions<T>): Promise<OutputByType[T]> => {
  const { start, end } = placeholderDelimiters;
  if (!start.trim() || !end.trim()) {
    throw new Error("Both start and end delimiters must be non-empty strings.");
  }

  const zipContent = unzipSync(toUint8Array(data));

  const xmlMap = new Map<string, Element>();
  const binaryMap = new Map<string, Uint8Array>();

  // Separate XML files from binary files
  for (const [key, value] of Object.entries(zipContent)) {
    if (key.endsWith(".xml") || key.endsWith(".rels")) {
      xmlMap.set(key, toJson(strFromU8(value)));
    } else {
      binaryMap.set(key, value);
    }
  }

  // Unified patch map: placeholders are delimiter-wrapped, findReplace keys are
  // literal strings. Both feed the same Pass A (typed cells) / Pass B (text) engine.
  const patchMap = new Map<string, ScalarValue>();
  if (placeholders) {
    for (const [key, value] of Object.entries(placeholders)) {
      patchMap.set(`${start}${key}${end}`, value);
    }
  }
  if (findReplace) {
    for (const [key, value] of Object.entries(findReplace)) {
      patchMap.set(key, value);
    }
  }

  const sharedStrings = xmlMap.get("xl/sharedStrings.xml");
  const sharedTextByIndex = buildSharedTextIndex(sharedStrings);

  const worksheetKeys = Object.keys(zipContent).filter(
    (k) => k.startsWith("xl/worksheets/sheet") && k.endsWith(".xml"),
  );

  // Pass A: rewrite pure placeholder cells to typed values (number / boolean / Date)
  for (const wsKey of worksheetKeys) {
    const ws = xmlMap.get(wsKey);
    if (ws) rewriteTypedCells(ws, patchMap, sharedTextByIndex);
  }

  // Pass B: text substitution in shared strings + inline strings
  // (string-valued patches, plus non-string patches embedded in mixed text)
  if (sharedStrings) patchSharedStrings(sharedStrings, patchMap);
  for (const wsKey of worksheetKeys) {
    const ws = xmlMap.get(wsKey);
    if (!ws) continue;
    patchWorksheetInlineStrings(ws, patchMap);
    // Pass C: print headers/footers (oddHeader/oddFooter/… literal text)
    patchHeaderFooter(ws, patchMap);
  }

  // Rebuild ZIP
  const files: Record<string, Uint8Array> = {};
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  for (const [key, value] of xmlMap) {
    files[key] =
      key === "docProps/core.xml" && coreProperties
        ? encoder.encode(XML_DECL + applyCorePropertiesOverride(value, coreProperties))
        : encoder.encode(js2xml(value));
  }
  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.XLSX);
};

/** Map each `<si>` shared-string index to its flattened text. */
function buildSharedTextIndex(sst: Element | undefined): Map<number, string> {
  const map = new Map<number, string>();
  if (!sst) return map;
  const root = sst.name ? sst : sst.elements?.[0];
  let idx = 0;
  for (const si of root?.elements ?? []) {
    if (si.name !== "si") continue;
    map.set(idx, flattenText(si));
    idx++;
  }
  return map;
}

/**
 * Pass A — for each cell whose entire text is a non-string placeholder, rewrite
 * the cell to a typed value cell (`<v>` with the correct `t`).
 */
function rewriteTypedCells(
  ws: Element,
  patchMap: Map<string, ScalarValue>,
  sharedTextByIndex: Map<number, string>,
): void {
  const root = ws.name ? ws : ws.elements?.[0];
  const sheetData = root && findLocalChild(root, "sheetData");
  if (!sheetData) return;

  for (const row of sheetData.elements ?? []) {
    if (localName(row) !== "row") continue;
    for (const cell of row.elements ?? []) {
      if (localName(cell) !== "c") continue;
      const text = effectiveCellText(cell, sharedTextByIndex);
      if (text === undefined) continue;

      for (const [placeholder, value] of patchMap) {
        if (typeof value === "string") continue; // Pass B handles string values
        if (text === placeholder) {
          rewriteCellToValue(cell, value);
          break; // a pure cell holds a single placeholder
        }
      }
    }
  }
}

/** Resolve the displayed text of a string/inline-string cell, else undefined. */
function effectiveCellText(
  cell: Element,
  sharedTextByIndex: Map<number, string>,
): string | undefined {
  const t = cell.attributes?.["t"];
  if (t === "inlineStr") {
    const isEl = findLocalChild(cell, "is");
    return isEl ? flattenText(isEl) : undefined;
  }
  if (t === "s") {
    const v = findLocalChild(cell, "v");
    const raw = v?.elements?.[0]?.text;
    const idx = Number(raw);
    return Number.isNaN(idx) ? undefined : sharedTextByIndex.get(idx);
  }
  return undefined;
}

/** Rewrite a cell element to a typed value (`<v>` with the correct `t`). */
function rewriteCellToValue(cell: Element, value: Exclude<ScalarValue, string>): void {
  const attrs = cell.attributes ?? (cell.attributes = {});
  let vText: string;
  if (typeof value === "boolean") {
    attrs["t"] = "b";
    vText = value ? "1" : "0";
  } else if (value instanceof Date) {
    delete attrs["t"]; // serial date is numeric
    vText = String(toExcelSerial(value));
  } else {
    delete attrs["t"]; // plain numeric
    vText = String(value);
  }
  cell.elements = [{ type: "element", name: "v", elements: [{ type: "text", text: vText }] }];
}

/** Pass B — substitute placeholders in shared-string `<si>` text. */
function patchSharedStrings(sst: Element, patchMap: Map<string, ScalarValue>): void {
  const root = sst.name ? sst : sst.elements?.[0];
  if (!root) return;

  for (const si of root.elements ?? []) {
    if (si.name !== "si") continue;

    // Simple: <si><t>text</t></si>
    const t = findLocalChild(si, "t");
    if (t) {
      patchTextElement(t, patchMap);
    } else {
      // Rich text: <si><r><t>text</t></r>...</si>
      for (const r of si.elements ?? []) {
        if (r.name !== "r") continue;
        const rt = findLocalChild(r, "t");
        if (rt) patchTextElement(rt, patchMap);
      }
    }
  }
}

/** Pass B — substitute placeholders in worksheet inline strings. */
function patchWorksheetInlineStrings(ws: Element, patchMap: Map<string, ScalarValue>): void {
  const root = ws.name ? ws : ws.elements?.[0];
  if (!root) return;

  const sheetData = findLocalChild(root, "sheetData");
  if (!sheetData) return;

  for (const row of sheetData.elements ?? []) {
    if (localName(row) !== "row") continue;
    for (const cell of row.elements ?? []) {
      if (localName(cell) !== "c") continue;

      // Inline strings: <c t="inlineStr"><is><t>text</t></is></c>
      const isEl = findLocalChild(cell, "is");
      if (isEl) {
        const t = findLocalChild(isEl, "t");
        if (t) patchTextElement(t, patchMap);
      }
    }
  }
}

const HEADER_FOOTER_NAMES = new Set([
  "oddHeader",
  "oddFooter",
  "evenHeader",
  "evenFooter",
  "firstHeader",
  "firstFooter",
]);

/**
 * Pass C — substitute placeholders in worksheet print header/footer text.
 * Header/footer children (oddHeader/oddFooter/…) carry their text directly
 * (not wrapped in `<t>`), so this is a substring substitution on the element's
 * own text node.
 */
function patchHeaderFooter(ws: Element, patchMap: Map<string, ScalarValue>): void {
  const root = ws.name ? ws : ws.elements?.[0];
  if (!root) return;
  const headerFooter = findLocalChild(root, "headerFooter");
  if (!headerFooter) return;

  for (const part of headerFooter.elements ?? []) {
    if (part.type !== "element" || !HEADER_FOOTER_NAMES.has(localName(part))) continue;
    const textNode = part.elements?.[0];
    if (!textNode || typeof textNode.text !== "string") continue;

    let replaced = textNode.text;
    for (const [placeholder, value] of patchMap) {
      if (!replaced.includes(placeholder)) continue;
      const display = typeof value === "string" ? value : toDisplayString(value);
      replaced = replaced.split(placeholder).join(display);
    }
    if (replaced !== textNode.text) {
      part.elements![0] = { ...textNode, text: replaced };
    }
  }
}

/** Substitute any placeholder substring in a `<t>` element's text. */
function patchTextElement(tEl: Element, patchMap: Map<string, ScalarValue>): void {
  const textNode = tEl.elements?.[0];
  const text = textNode?.text;
  if (typeof text !== "string") return;

  let replaced = text;
  for (const [placeholder, value] of patchMap) {
    if (!replaced.includes(placeholder)) continue;
    const display = typeof value === "string" ? value : toDisplayString(value);
    replaced = replaced.split(placeholder).join(display);
  }
  if (replaced !== text && tEl.elements) {
    tEl.elements[0] = { ...textNode, text: replaced };
  }
}

/** Flatten all descendant text of an element into a single string. */
function flattenText(el: Element): string {
  let out = "";
  for (const child of el.elements ?? []) {
    if (child.type === "text") {
      out += child.text ?? "";
    } else if (child.type === "element") {
      out += flattenText(child);
    }
  }
  return out;
}

function localName(el: Element): string {
  const name = el.name ?? "";
  const colonIdx = name.indexOf(":");
  return colonIdx >= 0 ? name.slice(colonIdx + 1) : name;
}

function findLocalChild(parent: Element, name: string): Element | undefined {
  return (parent.elements ?? []).find((el) => localName(el) === name);
}
