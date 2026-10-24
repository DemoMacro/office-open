/**
 * sheetData row scanner — the single parse implementation for worksheet rows.
 *
 * `parseWorkbook` requests `deferElements: ["sheetData"]` when reading
 * worksheet parts, so the XML parser captures the container's inner XML
 * verbatim (`Element.raw`) instead of materializing millions of row/cell
 * nodes (~1.75 GB of allocations on a 100k×20 sheet). This scanner walks
 * that string directly into RowOptions. Semantics are field-for-field
 * identical to the former Element-tree walk it replaced.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { attr, findChild, parse, textOf, unescapeXml } from "@office-open/xml";

import { parseRPr } from "../shared-strings";
import type {
  CellOptions,
  FormulaOptions,
  RichTextOptions,
  RichTextRunOptions,
  RowOptions,
} from "./types";

// ── Tag scanning primitives ──

/** Index of the next `<name` whose following char is a tag boundary, or -1. */
function indexOfTag(src: string, name: string, from: number, limit: number): number {
  const open = `<${name}`;
  const nameLen = name.length;
  let p = from;
  for (;;) {
    const idx = src.indexOf(open, p);
    if (idx === -1 || idx >= limit) return -1;
    const after = src.charCodeAt(idx + 1 + nameLen);
    if (
      after === 0x20 ||
      after === 0x09 ||
      after === 0x0a ||
      after === 0x0d ||
      after === 0x2f ||
      after === 0x3e
    ) {
      return idx;
    }
    p = idx + 1 + nameLen;
  }
}

/**
 * Index of the tag's closing `>`, honoring quoted attribute values (a quoted
 * `>` is not a tag end), or -1.
 */
function findTagEnd(src: string, from: number): number {
  let i = from;
  const len = src.length;
  while (i < len) {
    const c = src.charCodeAt(i);
    if (c === 0x22) {
      const close = src.indexOf('"', i + 1);
      if (close === -1) return -1;
      i = close + 1;
      continue;
    }
    if (c === 0x27) {
      const close = src.indexOf("'", i + 1);
      if (close === -1) return -1;
      i = close + 1;
      continue;
    }
    if (c === 0x3e) return i;
    i++;
  }
  return -1;
}

/**
 * Does raw XML contain constructs whose text can look like a tag? One scan
 * result is reused for an entire sheetData slice so the common OOXML case can
 * retain the direct close-tag search.
 */
function hasEscapedMarkup(src: string): boolean {
  return src.includes("<![CDATA[") || src.includes("<!--") || src.includes("<?");
}

/**
 * Fast path for schema-valid sheetData. These container/leaf names never nest
 * in SML, so the original direct search remains exact. Inputs containing
 * CDATA, comments, or processing instructions use the markup-aware scanner.
 */
function indexOfSimpleCloseTag(src: string, name: string, from: number, limit: number): number {
  const closeTag = `</${name}>`;
  const closeIndex = src.indexOf(closeTag, from);
  return closeIndex === -1 || closeIndex >= limit ? -1 : closeIndex;
}

/** Find a matching close tag without treating markup-internal text as XML. */
function indexOfCloseTag(
  src: string,
  name: string,
  from: number,
  limit = src.length,
  markupAware = true,
): number {
  if (!markupAware) return indexOfSimpleCloseTag(src, name, from, limit);

  const close = `</${name}`;
  let cursor = from;
  let depth = 1;

  while (cursor < limit) {
    const tagOpen = src.indexOf("<", cursor);
    if (tagOpen === -1 || tagOpen >= limit) return -1;

    if (src.startsWith("<![CDATA[", tagOpen)) {
      const end = src.indexOf("]]>", tagOpen + 9);
      if (end === -1 || end >= limit) return -1;
      cursor = end + 3;
      continue;
    }
    if (src.startsWith("<!--", tagOpen)) {
      const end = src.indexOf("-->", tagOpen + 4);
      if (end === -1 || end >= limit) return -1;
      cursor = end + 3;
      continue;
    }
    if (src.startsWith("<?", tagOpen)) {
      const end = src.indexOf("?>", tagOpen + 2);
      if (end === -1 || end >= limit) return -1;
      cursor = end + 2;
      continue;
    }

    const isClosing = src.startsWith(close, tagOpen);
    const nameStart = tagOpen + (isClosing ? 2 : 1);
    let nameEnd = nameStart;
    while (
      nameEnd < limit &&
      src.charCodeAt(nameEnd) !== 0x20 &&
      src.charCodeAt(nameEnd) !== 0x09 &&
      src.charCodeAt(nameEnd) !== 0x0a &&
      src.charCodeAt(nameEnd) !== 0x0d &&
      src.charCodeAt(nameEnd) !== 0x2f &&
      src.charCodeAt(nameEnd) !== 0x3e
    ) {
      nameEnd++;
    }
    if (src.slice(nameStart, nameEnd) !== name) {
      const end = findTagEnd(src, nameStart);
      if (end === -1 || end >= limit) return -1;
      cursor = end + 1;
      continue;
    }

    const end = findTagEnd(src, nameEnd);
    if (end === -1 || end >= limit) return -1;
    if (isClosing) {
      depth--;
      if (depth === 0) return tagOpen;
    } else if (src.charCodeAt(end - 1) !== 0x2f) {
      depth++;
    }
    cursor = end + 1;
  }
  return -1;
}

/** Visit `name="value"` pairs inside a tag (from = after tag name, to = before `>`/`/>`). */
function scanAttrs(
  src: string,
  from: number,
  to: number,
  visit: (name: string, value: string) => void,
): void {
  let i = from;
  while (i < to) {
    while (i < to) {
      const c = src.charCodeAt(i);
      if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) break;
      i++;
    }
    if (i >= to) return;
    const nameStart = i;
    while (i < to && src.charCodeAt(i) !== 0x3d) i++;
    if (i >= to) return;
    const name = src.slice(nameStart, i);
    i++; // '='
    while (i < to) {
      const c = src.charCodeAt(i);
      if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) break;
      i++;
    }
    const quote = src.charCodeAt(i);
    if (quote !== 0x22 && quote !== 0x27) return;
    i++;
    const valueStart = i;
    while (i < to && src.charCodeAt(i) !== quote) i++;
    visit(name, unescapeXml(src.slice(valueStart, i)));
    i++;
  }
}

/** parseOnOff truthiness for on/off attribute values ("1"/"true"/"on" plus the
 *  single-letter t/f Word 2007 wrote). */
function isOn(value: string): boolean {
  const lower = value.length <= 5 ? value.toLowerCase() : value;
  return lower === "1" || lower === "true" || lower === "on" || lower === "t";
}

function isOff(value: string): boolean {
  const lower = value.length <= 5 ? value.toLowerCase() : value;
  return lower === "0" || lower === "false" || lower === "off" || lower === "f";
}

/**
 * Text content of a leaf element's inner XML slice: text is unescaped and
 * CDATA sections carry their inner characters verbatim. Self-closing / empty
 * content yields "".
 */
function textContent(src: string, from: number, to: number): string {
  if (to <= from) return "";
  const text = src.slice(from, to);
  if (!text.includes("<")) return unescapeXml(text);

  // Outside CDATA, character data is escaped; CDATA content is literal. Copy
  // markup-free text separately from CDATA so `<![CDATA[&amp;]]>` remains
  // `&amp;`; comments and processing instructions contribute no text.
  let result = "";
  let cursor = 0;
  while (cursor < text.length) {
    const start = text.indexOf("<", cursor);
    if (start === -1) {
      result += unescapeXml(text.slice(cursor));
      break;
    }
    result += unescapeXml(text.slice(cursor, start));

    if (text.startsWith("<![CDATA[", start)) {
      const end = text.indexOf("]]>", start + 9);
      if (end === -1) throw new Error("Malformed XML: unclosed CDATA section.");
      result += text.slice(start + 9, end);
      cursor = end + 3;
      continue;
    }
    if (text.startsWith("<!--", start)) {
      const end = text.indexOf("-->", start + 4);
      if (end === -1) throw new Error("Malformed XML: unclosed comment.");
      cursor = end + 3;
      continue;
    }
    if (text.startsWith("<?", start)) {
      const end = text.indexOf("?>", start + 2);
      if (end === -1) throw new Error("Malformed XML: unclosed processing instruction.");
      cursor = end + 2;
      continue;
    }

    throw new Error("Malformed XML: unexpected markup in text content.");
  }
  return result;
}

/**
 * Content of an `<is>` slice: the first direct `<t>` (with its xml:space form)
 * and any direct CT_RElt `<r>` rich runs.
 */
function inlineStringText(
  src: string,
  from: number,
  to: number,
  markupAware: boolean,
):
  | {
      text?: string;
      spaceRaw?: "default" | "preserve";
      runs?: RichTextRunOptions[];
    }
  | undefined {
  let text: string | undefined;
  let spaceRaw: "default" | "preserve" | undefined;
  let runs: RichTextRunOptions[] | undefined;
  let q = from;
  while (q < to) {
    const lt = src.indexOf("<", q);
    if (lt === -1 || lt >= to) break;
    const tEnd = findTagEnd(src, lt + 1);
    if (tEnd === -1 || tEnd >= to) break;
    const afterLt = lt + 1;
    let nameEnd = afterLt;
    while (nameEnd < tEnd) {
      const c = src.charCodeAt(nameEnd);
      if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0x2f) break;
      nameEnd++;
    }
    const name = src.slice(afterLt, nameEnd);
    if (name === "t") {
      scanAttrs(
        src,
        nameEnd,
        src.charCodeAt(tEnd - 1) === 0x2f ? tEnd - 1 : tEnd,
        (aName, value) => {
          if (aName === "xml:space" && (value === "preserve" || value === "default")) {
            spaceRaw = value;
          }
        },
      );
      if (src.charCodeAt(tEnd - 1) === 0x2f) {
        text = "";
        q = tEnd + 1;
        continue;
      }
      const close = indexOfCloseTag(src, "t", tEnd + 1, to, markupAware);
      text = close === -1 ? "" : textContent(src, tEnd + 1, close);
      if (close === -1) break;
      q = close + 4;
      continue;
    }
    if (name === "r") {
      // Rich-text run: its <t> is nested, not a direct child — tree-parse the
      // run slice (rare: only cells with formatted inline text pay this cost).
      const close = indexOfCloseTag(src, "r", tEnd + 1, to, markupAware);
      if (close === -1) break;
      // parse() wraps the slice in a synthetic document node; the <r> root is
      // its first child.
      const runEl = parse(src.slice(lt, close + 4)).elements?.[0];
      if (!runEl) {
        q = close + 4;
        continue;
      }
      const runT = findChild(runEl, "t");
      if (runT) {
        const run: RichTextRunOptions = { text: textOf(runT) ?? "" };
        if (attr(runT, "xml:space") !== undefined) run.textSpaceRaw = attr(runT, "xml:space");
        const rPrEl = findChild(runEl, "rPr");
        if (rPrEl) {
          run.properties = parseRPr(rPrEl);
          run.propertiesPresent = true;
        }
        (runs ??= []).push(run);
      }
      q = close + 4;
      continue;
    }
    q = tEnd + 1;
  }
  if (text === undefined && runs === undefined) return undefined;
  return {
    ...(text !== undefined ? { text } : {}),
    ...(spaceRaw !== undefined ? { spaceRaw } : {}),
    ...(runs ? { runs } : {}),
  };
}

// ── Row scanner ──

/**
 * Scan `sheetData` inner XML into RowOptions. Cell/row styles stay as raw
 * cellXfs indices (the SDK's StyleIndex model) — the style table travels
 * alongside on WorkbookOptions and the compiler adopts it wholesale, so the
 * indices resolve exactly as they did in the source file.
 */
export function parseSheetDataRows(
  raw: string,
  strings: (string | RichTextOptions)[],
): RowOptions[] {
  const rows: RowOptions[] = [];
  const len = raw.length;
  let pos = 0;
  const markupAware = hasEscapedMarkup(raw);

  for (;;) {
    const rowOpen = indexOfTag(raw, "row", pos, len);
    if (rowOpen === -1) break;
    const tagEnd = findTagEnd(raw, rowOpen + 4);
    if (tagEnd === -1) break;
    const selfClosing = raw.charCodeAt(tagEnd - 1) === 0x2f;
    const attrEnd = selfClosing ? tagEnd - 1 : tagEnd;

    const row: RowOptions = {};
    let rowClose = tagEnd;
    let rowStyleIdx: number | undefined;
    scanAttrs(raw, rowOpen + 4, attrEnd, (name, value) => {
      switch (name) {
        case "r": {
          const n = Number(value);
          if (!isNaN(n)) row.rowNumber = n;
          break;
        }
        case "ht": {
          const n = Number(value);
          if (!isNaN(n)) row.height = n;
          break;
        }
        case "hidden":
          {
            const hidden = isOn(value);
            if (hidden !== undefined) row.hidden = hidden;
          }
          break;
        case "spans":
          row.spans = value;
          break;
        case "customFormat":
          if (value !== undefined) {
            const customFormat = parseOnOff(value);
            if (customFormat !== undefined) row.customFormat = customFormat;
            row.customFormatDeclared = true;
          }
          break;
        case "customHeight": {
          const customHeight = parseOnOff(value);
          if (customHeight !== undefined) row.customHeight = customHeight;
          break;
        }
        case "thickTop":
          if (isOn(value)) row.thickTop = true;
          break;
        case "thickBot":
          if (isOn(value)) row.thickBot = true;
          break;
        case "ph":
          if (isOn(value)) row.phonetic = true;
          break;
        case "outlineLevel": {
          const n = Number(value);
          if (!isNaN(n)) row.outlineLevel = n;
          break;
        }
        case "collapsed":
          {
            const collapsed = isOn(value);
            if (collapsed !== undefined) row.collapsed = collapsed;
          }
          break;
        case "s": {
          const n = Number(value);
          if (!isNaN(n)) rowStyleIdx = n;
          break;
        }
        case "x14ac:dyDescent": {
          const n = Number(value);
          if (!isNaN(n)) row.dyDescent = n;
          break;
        }
      }
    });
    if (rowStyleIdx !== undefined) {
      row.style = rowStyleIdx;
    }
    if (row.customFormat !== undefined) row.customFormatDeclared = true;
    else row.customFormatDeclared = false;
    const cells: CellOptions[] = [];
    if (!selfClosing) {
      rowClose = indexOfCloseTag(raw, "row", tagEnd + 1, len, markupAware);
      if (rowClose === -1) break;
      let cp = tagEnd + 1;
      for (;;) {
        const cellOpen = indexOfTag(raw, "c", cp, rowClose);
        if (cellOpen === -1) break;
        const cellTagEnd = findTagEnd(raw, cellOpen + 2);
        if (cellTagEnd === -1) break;
        const cellSelfClosing = raw.charCodeAt(cellTagEnd - 1) === 0x2f;
        const cell: CellOptions = {};
        let type: string | undefined;
        let styleIdx: number | undefined;
        let styleRawValue: string | undefined;

        // Inline attribute scan for the five cell attributes (r/t/s/cm/vm).
        // scanAttrs + closure costs one closure allocation and one name-string
        // slice per attribute on a 2M-cell sheet; dispatching on name length +
        // first char keeps the hot attributes allocation-free (other names are
        // skipped, same as the switch default they replace).
        let ap = cellOpen + 2;
        const attrLimit = cellSelfClosing ? cellTagEnd - 1 : cellTagEnd;
        while (ap < attrLimit) {
          while (ap < attrLimit) {
            const c = raw.charCodeAt(ap);
            if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) break;
            ap++;
          }
          if (ap >= attrLimit) break;
          const aNameStart = ap;
          while (ap < attrLimit && raw.charCodeAt(ap) !== 0x3d) ap++;
          if (ap >= attrLimit) break;
          const aNameLen = ap - aNameStart;
          ap++; // '='
          while (ap < attrLimit) {
            const c = raw.charCodeAt(ap);
            if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) break;
            ap++;
          }
          const aQuote = raw.charCodeAt(ap);
          if (aQuote !== 0x22 && aQuote !== 0x27) break;
          ap++;
          const aValueStart = ap;
          while (ap < attrLimit && raw.charCodeAt(ap) !== aQuote) ap++;
          const aFirst = raw.charCodeAt(aNameStart);
          if (aNameLen === 1) {
            if (aFirst === 0x72 /* r */) {
              cell.reference = unescapeXml(raw.slice(aValueStart, ap));
            } else if (aFirst === 0x74 /* t */) {
              type = unescapeXml(raw.slice(aValueStart, ap));
            } else if (aFirst === 0x73 /* s */) {
              const lexical = unescapeXml(raw.slice(aValueStart, ap));
              const n = Number(lexical);
              if (lexical !== "" && !isNaN(n)) styleIdx = n;
              else styleRawValue = lexical;
            }
          } else if (aNameLen === 2 && aFirst === 0x63 /* cm */) {
            const n = Number(unescapeXml(raw.slice(aValueStart, ap)));
            if (!isNaN(n)) cell.cellMetadataId = n;
          } else if (aNameLen === 2 && aFirst === 0x76 /* vm */) {
            const n = Number(unescapeXml(raw.slice(aValueStart, ap)));
            if (!isNaN(n)) cell.valueMetadataId = n;
          } else if (aNameLen === 2 && aFirst === 0x70 /* ph */) {
            if (isOn(unescapeXml(raw.slice(aValueStart, ap)))) cell.phonetic = true;
          }
          ap++;
        }

        if (styleIdx !== undefined) {
          cell.style = styleIdx;
        } else if (styleRawValue !== undefined) {
          cell.styleRaw = styleRawValue;
        }

        let cellEnd = cellTagEnd + 1;
        if (!cellSelfClosing) {
          const cellClose = indexOfCloseTag(raw, "c", cellTagEnd + 1, rowClose, markupAware);
          if (cellClose === -1) break;
          cellEnd = cellClose + 4;
          let vText: string | undefined;
          let vNum: number | undefined;
          let vRaw: string | undefined;
          let vTypeRaw: string | undefined;
          let vSpaceRaw: "default" | "preserve" | undefined;
          let inline:
            | {
                text?: string;
                spaceRaw?: "default" | "preserve";
                runs?: RichTextRunOptions[];
              }
            | undefined;
          let hasInline = false;
          let formula: FormulaOptions | undefined;

          let p = cellTagEnd + 1;
          while (p < cellClose) {
            const lt = raw.indexOf("<", p);
            if (lt === -1 || lt >= cellClose) break;
            const tEnd = findTagEnd(raw, lt + 1);
            if (tEnd === -1) break;
            let nameEnd = lt + 2;
            while (nameEnd < tEnd) {
              const c = raw.charCodeAt(nameEnd);
              if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0x2f) break;
              nameEnd++;
            }
            // Dispatch on tag-name length + first char: v/is/f cover the
            // children of virtually every cell, so the common path never
            // allocates a name string.
            const nameLen = nameEnd - (lt + 1);
            const first = raw.charCodeAt(lt + 1);
            if (nameLen === 1 && first === 0x76 /* v */) {
              const vSelfClosing = raw.charCodeAt(tEnd - 1) === 0x2f;
              scanAttrs(raw, nameEnd, vSelfClosing ? tEnd - 1 : tEnd, (name, value) => {
                if (name === "xml:space" && (value === "preserve" || value === "default"))
                  vSpaceRaw = value;
                else if (name === "t") vTypeRaw = value;
              });
              if (vSelfClosing) {
                if (vText === undefined) vText = "";
              } else {
                const close = indexOfCloseTag(raw, "v", tEnd + 1, cellClose, markupAware);
                if (close === -1) break;
                if (vText === undefined && vNum === undefined) {
                  // Fast path: pure-digit <v> content (about half the cells in
                  // data-heavy sheets) parses straight from char codes — no
                  // slice, no entity scan, no Number() re-scan. Up to 15
                  // digits is always exact in float64 (same bound as
                  // nativeTypeValue); anything else takes the string path.
                  const vStart = tEnd + 1;
                  const vLen = close - vStart;
                  let n = 0;
                  let allDigits = vLen > 0 && vLen <= 15;
                  if (allDigits) {
                    for (let q = vStart; q < close; q++) {
                      const d = raw.charCodeAt(q) - 0x30;
                      if (d < 0 || d > 9) {
                        allDigits = false;
                        break;
                      }
                      n = n * 10 + d;
                    }
                  }
                  if (allDigits) {
                    vNum = n;
                    vRaw = raw.slice(vStart, close);
                  } else {
                    vText = textContent(raw, vStart, close);
                    vRaw = vText;
                  }
                }
                p = close + 4;
                continue;
              }
            } else if (nameLen === 2 && first === 0x69 /* is */) {
              hasInline = true;
              const close = indexOfCloseTag(raw, "is", tEnd + 1, cellClose, markupAware);
              const isEnd = close === -1 ? cellClose : close;
              if (inline === undefined) {
                inline = inlineStringText(raw, tEnd + 1, isEnd, markupAware);
              }
              if (close === -1) break;
              p = close + 5;
              continue;
            } else if (nameLen === 1 && first === 0x66 /* f */) {
              const fSelfClosing = raw.charCodeAt(tEnd - 1) === 0x2f;
              const f: FormulaOptions = { formula: "" };
              scanAttrs(raw, nameEnd, fSelfClosing ? tEnd - 1 : tEnd, (n3, v3) => {
                switch (n3) {
                  case "t":
                    if (v3 !== "normal") f.type = v3 as FormulaOptions["type"];
                    break;
                  case "ref":
                    f.reference = v3;
                    break;
                  case "si": {
                    const n = Number(v3);
                    if (!isNaN(n)) f.sharedIndex = n;
                    break;
                  }
                  case "aca": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.aca = v;
                    break;
                  }
                  case "dt2D": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.dt2D = v;
                    break;
                  }
                  case "dtr": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.dtr = v;
                    break;
                  }
                  case "del1": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.del1 = v;
                    break;
                  }
                  case "del2": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.del2 = v;
                    break;
                  }
                  case "r1":
                    f.inputCell1 = v3;
                    break;
                  case "r2":
                    f.inputCell2 = v3;
                    break;
                  case "ca": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.calculateCell = v;
                    break;
                  }
                  case "bx": {
                    const v = isOn(v3) ? true : isOff(v3) ? false : undefined;
                    if (v !== undefined) f.arrayContext = v;
                    break;
                  }
                }
              });
              if (fSelfClosing) {
                p = tEnd + 1;
              } else {
                const close = indexOfCloseTag(raw, "f", tEnd + 1, cellClose, markupAware);
                if (close === -1) break;
                f.formula = textContent(raw, tEnd + 1, close);
                p = close + 4;
              }
              if (formula === undefined) formula = f;
              continue;
            }
            p = tEnd + 1;
          }

          // Cell value — resolution order matches the former tree walk.
          if (type === "n" || type === "d") cell.typeRaw = type;
          if (type === "s" && (vText !== undefined || vNum !== undefined)) {
            const idx = vNum !== undefined ? vNum : parseInt(vText!, 10);
            // An index that parses but has no table entry still round-trips —
            // the source may simply lack the SST part.
            if (Number.isNaN(idx)) {
              cell.value = "";
            } else {
              const sstValue = strings[idx];
              cell.value = sstValue ?? "";
              cell.sharedIndex = idx;
            }
          } else if (type === "b" && (vText !== undefined || vNum !== undefined)) {
            cell.value = vNum !== undefined ? vNum === 1 : isOn(vText!);
          } else if (type === "e" && (vText !== undefined || vNum !== undefined)) {
            cell.error = vNum !== undefined ? String(vNum) : vText!;
          } else if (type === "str" && (vText !== undefined || vNum !== undefined)) {
            cell.value = vText ?? String(vNum);
            cell.typeRaw = "str";
          } else if (type === "d" && (vText !== undefined || vNum !== undefined)) {
            cell.value = vNum !== undefined ? String(vNum) : vText!;
          } else if (type === "inlineStr" && hasInline) {
            cell.inline = true;
            if (vRaw !== undefined) cell.inlineValueRaw = vRaw;
            else if (vNum !== undefined) cell.inlineValueRaw = String(vNum);
            else if (vText !== undefined) cell.inlineValueRaw = vText;
            if (inline?.runs && inline.runs.length > 0) {
              const entry: RichTextOptions = { runs: inline.runs };
              if (inline.text !== undefined) {
                entry.text = inline.text;
                if (inline.spaceRaw !== undefined) entry.textSpaceRaw = inline.spaceRaw;
                else if (/^\s|\s$/.test(inline.text)) entry.textSpaceRaw = "default";
              }
              cell.value = entry;
            } else {
              const text = inline?.text ?? "";
              // The explicit source xml:space form is significant even without
              // outer whitespace (producers emit it unconditionally); carry it
              // so stringify mirrors the source attribute.
              const spaceRaw =
                inline?.spaceRaw ?? (/^\s|\s$/.test(text) ? ("default" as const) : undefined);
              if (spaceRaw !== undefined) {
                cell.value = { text, textSpaceRaw: spaceRaw };
              } else {
                cell.value = text;
              }
            }
          } else if (vNum !== undefined) {
            cell.value = vNum;
          } else if (vText !== undefined) {
            const num = vText === "" ? Number.NaN : Number(vText);
            cell.value = isNaN(num) ? vText : num;
          }
          if (
            type !== undefined &&
            type !== "n" &&
            type !== "d" &&
            cell.value === undefined &&
            cell.error === undefined &&
            cell.inline === undefined &&
            cell.sharedIndex === undefined
          ) {
            cell.typeRaw = type;
          }
          if (formula !== undefined && type !== undefined && type !== "n" && type !== "d") {
            cell.typeRaw = type;
          }
          if (formula !== undefined) cell.formula = formula;
          if (vTypeRaw !== undefined) cell.cachedValueTypeRaw = vTypeRaw;
          if (
            (type === "b" && (vText === "true" || vText === "false")) ||
            (vSpaceRaw === "preserve" &&
              vText !== undefined &&
              (type === undefined || type === "str" || type === "d"))
          )
            cell.valueRaw = vText;
          if (vRaw !== undefined && (type === undefined || type === "n")) cell.valueRaw = vRaw;
        } else if (type !== undefined) {
          cell.typeRaw = type;
        }

        cells.push(cell);
        cp = cellEnd;
      }
    }

    row.cells = cells;
    rows.push(row);
    pos = selfClosing ? tagEnd + 1 : rowClose + 6;
  }

  return rows;
}
