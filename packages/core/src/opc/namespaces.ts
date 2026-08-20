/**
 * Namespace URI → canonical prefix table for parse-time normalization.
 *
 * Producers are free to bind any prefix to a namespace (XML namespaces say
 * the prefix is arbitrary), and real-world generators do: pandoc writes
 * `ns0:`, ClosedXML binds the spreadsheet namespace to `x:`. Everything
 * downstream of the parser — descriptor `findChild("w:p")` calls, registry
 * path matching — addresses elements by the canonical prefixes this library
 * itself emits, so non-canonical sources parsed blind.
 *
 * Feeding this table to the XML parser's `normalizeNamespaces` rewrites every
 * element and attribute name onto the canonical prefix (an empty value means
 * the canonical form is unprefixed — the sml and OPC namespaces), while
 * unknown URIs keep their source prefix and `xmlns` declarations follow the
 * rewrite so serialized output stays self-consistent.
 *
 * The table mirrors the prefixes this library's own stringify output uses —
 * when a new namespace appears in a part template, add it here in the same
 * change. Strict-OOXML variants (`purl.oclc.org`) are deliberately absent:
 * mapping them would silently convert strict packages to transitional.
 */

const SCHEMAS_OPENXML = "http://schemas.openxmlformats.org";
const SCHEMAS_MS = "http://schemas.microsoft.com";

/**
 * Obsolete namespace URIs accepted by Office and the Open XML SDK, mapped to
 * their final standardized forms. A round-trip must normalize passthrough XML
 * too: rebuilding modeled parts with the final URI while carrying an obsolete
 * URI verbatim in a raw part creates a mixed-dialect package that Word rejects.
 *
 * Keep this aligned with Open XML SDK's OpenXmlNamespaceResolver
 * `_extendedNamespaces` table.
 */
export const OOXML_OBSOLETE_NAMESPACE_ALIASES: Readonly<Record<string, string>> = {
  [`${SCHEMAS_OPENXML}/wordprocessingml/2006/3/main`]: `${SCHEMAS_OPENXML}/wordprocessingml/2006/main`,
  [`${SCHEMAS_OPENXML}/wordprocessingml/2006/5/main`]: `${SCHEMAS_OPENXML}/wordprocessingml/2006/main`,
  [`${SCHEMAS_OPENXML}/wordprocessingml/2006/6/main`]: `${SCHEMAS_OPENXML}/wordprocessingml/2006/main`,
  [`${SCHEMAS_OPENXML}/spreadsheetml/2006/5/main`]: `${SCHEMAS_OPENXML}/spreadsheetml/2006/main`,
  [`${SCHEMAS_OPENXML}/spreadsheetml/2006/7/main`]: `${SCHEMAS_OPENXML}/spreadsheetml/2006/main`,
  [`${SCHEMAS_OPENXML}/presentationml/2006/3/main`]: `${SCHEMAS_OPENXML}/presentationml/2006/main`,
  [`${SCHEMAS_OPENXML}/drawingml/2006/3/main`]: `${SCHEMAS_OPENXML}/drawingml/2006/main`,
  [`${SCHEMAS_MS}/office/word/2010/11/wordml`]: `${SCHEMAS_MS}/office/word/2012/wordml`,
};

/**
 * Normalize obsolete OOXML namespace URI bytes without reparsing or
 * reserializing the XML. Byte-level replacement preserves every other source
 * detail (declaration, whitespace, prefix choices, attribute order, and line
 * endings), which is required for passthrough parts.
 */
export function normalizeObsoleteNamespaceAliases(data: Uint8Array): Uint8Array {
  let result = data;
  for (const [obsolete, finalUri] of Object.entries(OOXML_OBSOLETE_NAMESPACE_ALIASES)) {
    result = replaceAsciiBytes(result, obsolete, finalUri);
  }
  return result;
}

function isNamespaceDeclarationValue(
  data: Uint8Array,
  valueOffset: number,
  valueLength: number,
): boolean {
  const quote = data[valueOffset - 1];
  if ((quote !== 34 && quote !== 39) || data[valueOffset + valueLength] !== quote) return false;
  let i = valueOffset - 2;
  while (i >= 0 && (data[i] === 32 || data[i] === 9 || data[i] === 10 || data[i] === 13)) i--;
  if (data[i] !== 61) return false; // =
  i--;
  while (i >= 0 && (data[i] === 32 || data[i] === 9 || data[i] === 10 || data[i] === 13)) i--;
  const nameEnd = i + 1;
  while (
    i >= 0 &&
    ((data[i]! >= 65 && data[i]! <= 90) ||
      (data[i]! >= 97 && data[i]! <= 122) ||
      (data[i]! >= 48 && data[i]! <= 57) ||
      data[i] === 58 ||
      data[i] === 45 ||
      data[i] === 95)
  ) {
    i--;
  }
  const name = new TextDecoder().decode(data.subarray(i + 1, nameEnd));
  return name === "xmlns" || name.startsWith("xmlns:");
}

function replaceAsciiBytes(data: Uint8Array, search: string, replacement: string): Uint8Array {
  const needle = new TextEncoder().encode(search);
  const positions: number[] = [];
  for (let i = 0; i <= data.length - needle.length; ) {
    let matches = true;
    for (let j = 0; j < needle.length; j++) {
      if (data[i + j] !== needle[j]) {
        matches = false;
        break;
      }
    }
    if (matches && isNamespaceDeclarationValue(data, i, needle.length)) {
      positions.push(i);
      i += needle.length;
    } else {
      i++;
    }
  }
  if (positions.length === 0) return data;

  const value = new TextEncoder().encode(replacement);
  const out = new Uint8Array(data.length + positions.length * (value.length - needle.length));
  let sourceOffset = 0;
  let targetOffset = 0;
  for (const position of positions) {
    const prefix = data.subarray(sourceOffset, position);
    out.set(prefix, targetOffset);
    targetOffset += prefix.length;
    out.set(value, targetOffset);
    targetOffset += value.length;
    sourceOffset = position + needle.length;
  }
  out.set(data.subarray(sourceOffset), targetOffset);
  return out;
}

export const OOXML_CANONICAL_PREFIXES: Readonly<Record<string, string>> = {
  // Obsolete aliases resolve through their final namespace's canonical prefix.
  ...Object.fromEntries(
    Object.entries(OOXML_OBSOLETE_NAMESPACE_ALIASES).map(([oldUri, finalUri]) => [
      oldUri,
      finalUri === `${SCHEMAS_OPENXML}/spreadsheetml/2006/main`
        ? ""
        : finalUri === `${SCHEMAS_OPENXML}/presentationml/2006/main`
          ? "p"
          : finalUri === `${SCHEMAS_OPENXML}/drawingml/2006/main`
            ? "a"
            : finalUri === `${SCHEMAS_OPENXML}/wordprocessingml/2006/main`
              ? "w"
              : "w15",
    ]),
  ),

  // ── Main format namespaces ──
  [`${SCHEMAS_OPENXML}/wordprocessingml/2006/main`]: "w",
  [`${SCHEMAS_OPENXML}/presentationml/2006/main`]: "p",
  // spreadsheetml canonical form is the default (unprefixed) namespace
  [`${SCHEMAS_OPENXML}/spreadsheetml/2006/main`]: "",
  [`${SCHEMAS_OPENXML}/drawingml/2006/main`]: "a",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/relationships`]: "r",
  [`${SCHEMAS_OPENXML}/markup-compatibility/2006`]: "mc",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/math`]: "m",

  // ── DrawingML variants ──
  [`${SCHEMAS_OPENXML}/drawingml/2006/wordprocessingDrawing`]: "wp",
  [`${SCHEMAS_OPENXML}/drawingml/2006/spreadsheetDrawing`]: "xdr",
  [`${SCHEMAS_OPENXML}/drawingml/2006/chart`]: "c",
  [`${SCHEMAS_OPENXML}/drawingml/2006/chartDrawing`]: "cdr",
  [`${SCHEMAS_OPENXML}/drawingml/2006/diagram`]: "dgm",
  [`${SCHEMAS_OPENXML}/drawingml/2006/lockedCanvas`]: "lc",
  [`${SCHEMAS_OPENXML}/drawingml/2006/picture`]: "pic",

  // ── OPC package namespaces (canonical form: unprefixed) ──
  [`${SCHEMAS_OPENXML}/package/2006/relationships`]: "",
  [`${SCHEMAS_OPENXML}/package/2006/content-types`]: "",
  [`${SCHEMAS_OPENXML}/package/2006/metadata/core-properties`]: "",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/extended-properties`]: "",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/custom-properties`]: "",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/docPropsVTypes`]: "vt",
  [`${SCHEMAS_OPENXML}/officeDocument/2006/bibliography`]: "b",

  // ── VML family ──
  "urn:schemas-microsoft-com:vml": "v",
  "urn:schemas-microsoft-com:office:office": "o",
  "urn:schemas-microsoft-com:office:word": "w10",
  "urn:schemas-microsoft-com:office:excel": "x",
  "urn:schemas-microsoft-com:office:powerpoint": "pvml",

  // ── Word extensions ──
  [`${SCHEMAS_MS}/office/word/2006/wordml`]: "wne",
  [`${SCHEMAS_MS}/office/word/2010/wordml`]: "w14",
  [`${SCHEMAS_MS}/office/word/2010/11/wordml`]: "w15",
  [`${SCHEMAS_MS}/office/word/2012/wordml`]: "w15",
  [`${SCHEMAS_MS}/office/word/2016/wordml/cid`]: "w16cid",
  [`${SCHEMAS_MS}/office/word/2018/wordml`]: "w16",
  [`${SCHEMAS_MS}/office/word/2018/wordml/cex`]: "w16cex",
  [`${SCHEMAS_MS}/office/word/2020/wordml/sdtdatahash`]: "w16sdtdh",
  [`${SCHEMAS_MS}/office/word/2015/wordml/symex`]: "w16se",
  [`${SCHEMAS_MS}/office/word/2010/wordprocessingDrawing`]: "wp14",
  [`${SCHEMAS_MS}/office/word/2010/wordprocessingCanvas`]: "wpc",
  [`${SCHEMAS_MS}/office/word/2010/wordprocessingGroup`]: "wpg",
  [`${SCHEMAS_MS}/office/word/2010/wordprocessingInk`]: "wpi",
  [`${SCHEMAS_MS}/office/word/2010/wordprocessingShape`]: "wps",

  // ── Drawing extensions ──
  [`${SCHEMAS_MS}/office/drawing/2010/main`]: "a14",
  [`${SCHEMAS_MS}/office/drawing/2014/main`]: "a16",
  [`${SCHEMAS_MS}/office/drawing/2016/ink`]: "aink",
  [`${SCHEMAS_MS}/office/drawing/2017/model3d`]: "am3d",
  [`${SCHEMAS_MS}/office/drawing/2016/SVG/main`]: "asvg",
  [`${SCHEMAS_MS}/office/drawing/2007/8/2/chart`]: "c14",
  [`${SCHEMAS_MS}/office/drawing/2008/diagram`]: "dsp",
  [`${SCHEMAS_MS}/office/drawing/2014/chartex`]: "cx",
  [`${SCHEMAS_MS}/office/drawing/2015/9/8/chartex`]: "cx1",
  [`${SCHEMAS_MS}/office/drawing/2015/10/21/chartex`]: "cx2",
  [`${SCHEMAS_MS}/office/drawing/2016/5/9/chartex`]: "cx3",
  [`${SCHEMAS_MS}/office/drawing/2016/5/10/chartex`]: "cx4",
  [`${SCHEMAS_MS}/office/drawing/2016/5/11/chartex`]: "cx5",
  [`${SCHEMAS_MS}/office/drawing/2016/5/12/chartex`]: "cx6",
  [`${SCHEMAS_MS}/office/drawing/2016/5/13/chartex`]: "cx7",
  [`${SCHEMAS_MS}/office/drawing/2016/5/14/chartex`]: "cx8",
  [`${SCHEMAS_MS}/office/2019/extlst`]: "oel",

  // ── PowerPoint extensions ──
  [`${SCHEMAS_MS}/office/powerpoint/2010/main`]: "p14",
  [`${SCHEMAS_MS}/office/powerpoint/2015/main`]: "p15",

  // ── Excel extensions ──
  [`${SCHEMAS_MS}/office/spreadsheetml/2009/9/main`]: "x14",
  [`${SCHEMAS_MS}/office/spreadsheetml/2009/9/ac`]: "x14ac",
  [`${SCHEMAS_MS}/office/spreadsheetml/2010/11/main`]: "x15",
  [`${SCHEMAS_MS}/office/spreadsheetml/2010/11/ac`]: "x15ac",
  [`${SCHEMAS_MS}/office/spreadsheetml/2014/revision`]: "xr",
  [`${SCHEMAS_MS}/office/spreadsheetml/2015/revision2`]: "xr2",
  [`${SCHEMAS_MS}/office/spreadsheetml/2016/revision3`]: "xr3",
  [`${SCHEMAS_MS}/office/spreadsheetml/2016/revision6`]: "xr6",
  [`${SCHEMAS_MS}/office/spreadsheetml/2016/revision10`]: "xr10",
};
