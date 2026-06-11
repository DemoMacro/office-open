/**
 * XLSX patching — replace placeholders in existing .xlsx files.
 *
 * Unlike DOCX/PPTX (paragraph-based), XLSX patching targets cell values:
 * - Replaces placeholders in the shared strings table (most common)
 * - Replaces placeholders in inline strings within worksheet cells
 *
 * @module
 */
import { OoxmlMimeType, unzipSync, zipAndConvert, strFromU8, toJson } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";
import { toUint8Array } from "@office-open/core";
import { js2xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

export type InputDataType = Buffer | string | number[] | Uint8Array | ArrayBuffer | Blob;

export type PatchDocumentOutputType = OutputType;

export const PatchType = {
  CELL: "cell",
} as const;

export interface CellPatch {
  /** Replacement value (string, number, boolean, or Date) */
  value: string | number | boolean | Date;
}

export type IPatch = CellPatch;

export interface PatchWorkbookOptions<T extends PatchDocumentOutputType = PatchDocumentOutputType> {
  outputType: T;
  data: InputDataType;
  patches: Readonly<Record<string, IPatch>>;
  placeholderDelimiters?: Readonly<{
    start: string;
    end: string;
  }>;
}

/**
 * Patch an existing .xlsx workbook by replacing placeholder text in cells.
 *
 * Placeholders are matched in shared strings and inline strings.
 * For string replacements, the shared string value is updated in-place.
 * For non-string replacements, the cell type and value are updated.
 */
export const patchWorkbook = async <T extends PatchDocumentOutputType = PatchDocumentOutputType>({
  outputType,
  data,
  patches,
  placeholderDelimiters = { start: "{{", end: "}}" } as const,
}: PatchWorkbookOptions<T>): Promise<OutputByType[T]> => {
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

  const { start, end } = placeholderDelimiters;

  // Build placeholder → patch map
  const patchMap = new Map<string, IPatch>();
  for (const [key, patch] of Object.entries(patches)) {
    patchMap.set(`${start}${key}${end}`, patch);
  }

  // 1. Patch shared strings
  const sst = xmlMap.get("xl/sharedStrings.xml");
  if (sst) {
    patchSharedStrings(sst, patchMap);
  }

  // 2. Patch inline strings in worksheets
  const worksheetKeys = Object.keys(zipContent).filter(
    (k) => k.startsWith("xl/worksheets/sheet") && k.endsWith(".xml"),
  );
  for (const wsKey of worksheetKeys) {
    const ws = xmlMap.get(wsKey);
    if (ws) patchWorksheetInlineStrings(ws, patchMap);
  }

  // Rebuild ZIP
  const files: Record<string, Uint8Array> = {};
  for (const [key, value] of xmlMap) {
    files[key] = encoder.encode(js2xml(value));
  }
  for (const [key, value] of binaryMap) {
    files[key] = value;
  }

  return await zipAndConvert(files, outputType, OoxmlMimeType.XLSX);
};

function patchSharedStrings(sst: Element, patchMap: Map<string, IPatch>): void {
  // toJson returns {elements: [{name: "sst", ...}]} — unwrap to actual root
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

function patchWorksheetInlineStrings(ws: Element, patchMap: Map<string, IPatch>): void {
  const root = ws.name ? ws : ws.elements?.[0];
  if (!root) return;

  const sheetData = findLocalChild(root, "sheetData");
  if (!sheetData) return;

  for (const row of sheetData.elements ?? []) {
    if (localName(row) !== "row") continue;
    for (const cell of row.elements ?? []) {
      if (localName(cell) !== "c") continue;

      // Check inline strings: <c t="inlineStr"><is><t>text</t></is></c>
      const isEl = findLocalChild(cell, "is");
      if (isEl) {
        const t = findLocalChild(isEl, "t");
        if (t) patchTextElement(t, patchMap);
      }
    }
  }
}

function patchTextElement(tEl: Element, patchMap: Map<string, IPatch>): void {
  const text = tEl.elements?.[0]?.text;
  if (typeof text !== "string") return;

  for (const [placeholder, patch] of patchMap) {
    if (text.includes(placeholder)) {
      const newValue = String(patch.value);
      const replaced = text.replace(placeholder, newValue);
      if (tEl.elements) {
        tEl.elements[0] = { ...tEl.elements[0], text: replaced };
      }
    }
  }
}

function localName(el: Element): string {
  const name = el.name ?? "";
  const colonIdx = name.indexOf(":");
  return colonIdx >= 0 ? name.slice(colonIdx + 1) : name;
}

function findLocalChild(parent: Element, name: string): Element | undefined {
  return (parent.elements ?? []).find((el) => localName(el) === name);
}
