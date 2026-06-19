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
  appendRelationship,
  applyCorePropertiesOverride,
  getNextRelationshipIndex,
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
import type { ReadContext } from "@office-open/core/descriptor";
import { js2xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { SharedStrings, sharedStringsDesc } from "@parts/shared-strings";
import { buildWorksheetXml } from "@parts/worksheet";
import type { WorksheetOptions } from "@parts/worksheet";

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
  /**
   * Worksheet collection edits. The shared-strings table is rebuilt from the
   * template first, so appended/replaced worksheets continue indexing strings
   * at the correct offset. Cell styles are not supported on appended rows —
   * reference only styles already present in the template.
   */
  worksheets?: {
    /** Replace worksheets keyed by sheet name (as declared in workbook.xml). */
    replace?: Readonly<Record<string, WorksheetOptions>>;
    /** Append worksheets after the last existing sheet. */
    append?: Readonly<WorksheetOptions[]>;
  };
}

// Excel serial-date epoch: 1899-12-30 accounts for the spurious 1900 leap year.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKSHEET_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

/**
 * sharedStringsDesc.parse ignores its context, but its signature requires a
 * {@link ReadContext}. This stub satisfies the type without pulling in parse
 * machinery the patch path does not need.
 */
const STUB_READ_CTX: ReadContext = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
};

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
  worksheets,
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

  // Worksheet collection edits — rebuild shared strings first so appended/
  // replaced worksheets continue indexing at the right offset, then serialize.
  if (worksheets) {
    const sharedStrings = rebuildSharedStrings(xmlMap);
    if (worksheets.replace) {
      for (const [name, wsOpts] of Object.entries(worksheets.replace)) {
        replaceWorksheetInMap(xmlMap, name, wsOpts, sharedStrings);
      }
    }
    if (worksheets.append) {
      for (const wsOpts of worksheets.append) {
        appendWorksheetToMap(xmlMap, wsOpts, sharedStrings);
      }
    }
    rewriteSharedStrings(xmlMap, sharedStrings);
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

// ── Worksheet collection edits (append / replace) ──

/** Unwrap the document element: a parsed part is `{ declaration, elements: [<root>] }`. */
function rootOf(doc: Element | undefined): Element | undefined {
  return doc && (doc.name ? doc : doc.elements?.[0]);
}

/** Build a bare element node with the given attributes. */
function makeElement(name: string, attributes: Record<string, string> = {}): Element {
  return { type: "element", name, attributes, elements: [] };
}

/** Empty `<Relationships>` part (wrapped for js2xml serialization). */
function createRelationshipFile(): Element {
  return {
    declaration: { attributes: { encoding: "UTF-8", standalone: "yes", version: "1.0" } },
    elements: [
      {
        attributes: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
        elements: [],
        name: "Relationships",
        type: "element",
      },
    ],
  };
}

/** Next 1-based worksheet file number given existing worksheet parts. */
function nextWorksheetNumber(xmlMap: Map<string, Element>): number {
  let maxN = 0;
  for (const key of xmlMap.keys()) {
    const m = key.match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
    if (m) maxN = Math.max(maxN, Number(m[1]));
  }
  return maxN + 1;
}

/** Largest numeric `sheetId` among existing `<sheet>` entries. */
function maxSheetId(sheetsEl: Element | undefined): number {
  let maxId = 0;
  for (const child of sheetsEl?.elements ?? []) {
    if (localName(child) !== "sheet") continue;
    const id = Number(child.attributes?.["sheetId"]);
    if (Number.isFinite(id)) maxId = Math.max(maxId, id);
  }
  return maxId;
}

/** Resolve a relationship `rId` to its Target via a rels part root. */
function resolveRelTarget(relsRoot: Element | undefined, rId: string): string | undefined {
  for (const child of relsRoot?.elements ?? []) {
    if (localName(child) === "Relationship" && String(child.attributes?.["Id"]) === rId) {
      const target = child.attributes?.["Target"];
      return target !== undefined ? String(target) : undefined;
    }
  }
  return undefined;
}

/** Rebuild the shared-strings table from the template so new sheets extend it. */
function rebuildSharedStrings(xmlMap: Map<string, Element>): SharedStrings {
  const ss = new SharedStrings();
  const sstRoot = rootOf(xmlMap.get("xl/sharedStrings.xml"));
  if (sstRoot) {
    ss.loadEntries(sharedStringsDesc.parse(sstRoot, STUB_READ_CTX).entries);
  }
  return ss;
}

/** Re-serialize the shared-strings part (called after worksheets are built). */
function rewriteSharedStrings(xmlMap: Map<string, Element>, ss: SharedStrings): void {
  if (ss.count === 0) return;
  xmlMap.set(
    "xl/sharedStrings.xml",
    toJson(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${ss.serialize()}`),
  );
}

/**
 * Append a worksheet: serialize it (registering strings into the shared-strings
 * table), then wire workbook `<sheets>` + workbook rels + content types.
 */
function appendWorksheetToMap(
  xmlMap: Map<string, Element>,
  wsOpts: WorksheetOptions,
  sharedStrings: SharedStrings,
): void {
  const newN = nextWorksheetNumber(xmlMap);
  const sheetPath = `xl/worksheets/sheet${newN}.xml`;
  const sheetName = wsOpts.name ?? `Sheet${newN}`;

  // Worksheet part (no XML declaration — matches generated worksheet parts)
  xmlMap.set(sheetPath, toJson(buildWorksheetXml(wsOpts, { sharedStrings })));

  // workbook.xml <sheets> + workbook.xml.rels (the new rId ties them)
  const wbRoot = rootOf(xmlMap.get("xl/workbook.xml"));
  const wbSheets = wbRoot && findLocalChild(wbRoot, "sheets");
  const wbRels = xmlMap.get("xl/_rels/workbook.xml.rels") ?? createRelationshipFile();
  xmlMap.set("xl/_rels/workbook.xml.rels", wbRels);
  const newRId = getNextRelationshipIndex(wbRels);

  if (wbSheets) {
    const els = wbSheets.elements ?? (wbSheets.elements = []);
    els.push(
      makeElement("sheet", {
        name: sheetName,
        sheetId: String(maxSheetId(wbSheets) + 1),
        "r:id": `rId${newRId}`,
      }),
    );
  }
  appendRelationship(wbRels, newRId, WORKSHEET_REL_TYPE, `worksheets/sheet${newN}.xml`);

  // [Content_Types].xml Override
  const typesRoot = rootOf(xmlMap.get("[Content_Types].xml"));
  if (typesRoot) {
    const els = typesRoot.elements ?? (typesRoot.elements = []);
    els.push(
      makeElement("Override", {
        PartName: `/${sheetPath}`,
        ContentType: WORKSHEET_CONTENT_TYPE,
      }),
    );
  }
}

/**
 * Replace a worksheet by name: resolve the name to its part via workbook rels,
 * then rewrite only that part's content. Sheet identity, rId, and content types
 * are unchanged.
 */
function replaceWorksheetInMap(
  xmlMap: Map<string, Element>,
  name: string,
  wsOpts: WorksheetOptions,
  sharedStrings: SharedStrings,
): void {
  const wbRoot = rootOf(xmlMap.get("xl/workbook.xml"));
  const wbSheets = wbRoot && findLocalChild(wbRoot, "sheets");
  const sheetEls = (wbSheets?.elements ?? []).filter((e) => localName(e) === "sheet");
  const target = sheetEls.find((e) => String(e.attributes?.["name"]) === name);
  if (!target?.attributes?.["r:id"]) {
    throw new Error(`patchWorkbook: no worksheet named "${name}" to replace`);
  }
  const relTarget = resolveRelTarget(
    rootOf(xmlMap.get("xl/_rels/workbook.xml.rels")),
    String(target.attributes["r:id"]),
  );
  if (!relTarget) {
    throw new Error(`patchWorkbook: worksheet "${name}" relationship target not found`);
  }
  xmlMap.set(`xl/${relTarget}`, toJson(buildWorksheetXml(wsOpts, { sharedStrings })));
}
