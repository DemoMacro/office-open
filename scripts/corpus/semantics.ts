import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import {
  convertToEmu,
  convertToPt,
  convertToTwip,
  ParsedArchive,
  type ParsedArchive as ParsedArchiveType,
} from "@office-open/core";
import { OOXML_CANONICAL_PREFIXES } from "@office-open/core/opc";
import { unzipSync } from "fflate";

import { parse, type Element } from "../../packages/xml/dist/index.mjs";

export type SemanticPartKind = "xml" | "relationship" | "content-types" | "binary";

export type SemanticDiffCategory =
  | "missing-part"
  | "added-part"
  | "binary"
  | "element"
  | "attribute"
  | "text"
  | "child-order"
  | "child";

export interface SemanticPartDiff {
  path: string;
  kind: SemanticPartKind;
  category?: SemanticDiffCategory;
  xpath?: string;
  detail?: string;
}

interface CanonicalNode {
  name: string;
  attributes: Record<string, string>;
  text: string;
  children: CanonicalNode[];
}

const UNORDERED_PART_PATHS = new Set([
  "docProps/app.xml",
  "docProps/core.xml",
  "docProps/custom.xml",
]);

const IGNORED_ATTRIBUTES = new Set([
  "mc:Ignorable",
  "mc:MustUnderstand",
  "mc:PreserveAttributes",
  "mc:ProcessContent",
]);

/** Package Defaults the writer must add for opaque binaries it preserves. */
const WRITER_DEFAULT_EXTENSION_DEFAULTS = new Map([
  ["bin", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"],
]);

/** Extended/core-properties boolean elements whose text is xsd:boolean. */
const BOOLEAN_TEXT_ELEMENTS = new Set([
  "ScaleCrop",
  "LinksUpToDate",
  "SharedDoc",
  "HyperlinksChanged",
  "o:LockedField",
]);

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

/** OOXML toggle elements whose omitted val means semantic true. */
function createOnOffElements(): Set<string> {
  const schemaDir = path.join(ROOT_DIR, "ooxml-schemas", "transitional");
  const schemas = [
    { file: "wml.xsd", prefix: "w" },
    { file: "shared-math.xsd", prefix: "m" },
  ];
  const elements = new Set<string>();
  for (const schema of schemas) {
    const root = parse(readFileSync(path.join(schemaDir, schema.file), "utf8")).elements?.find(
      (candidate): candidate is Element =>
        candidate.type === "element" && !(candidate.name ?? "").startsWith("!"),
    );
    const visit = (element: Element | undefined): void => {
      if (!element || element.type !== "element") return;
      if (element.attributes?.type === "CT_OnOff" && element.attributes?.name) {
        elements.add(`${schema.prefix}:${element.attributes.name}`);
      }
      for (const child of element.elements ?? []) visit(child);
    };
    visit(root);
  }
  return elements;
}

const ON_OFF_ELEMENTS = createOnOffElements();
ON_OFF_ELEMENTS.add("w:active");
ON_OFF_ELEMENTS.add("wne:active");

/** XML attributes whose XSD default is emitted explicitly by the writers. */
const DEFAULT_ATTRIBUTES = new Map<string, Record<string, string>>([
  [
    "w:pgMar",
    {
      "w:header": "851",
      "w:footer": "992",
      "w:gutter": "0",
    },
  ],
  [
    "wp:inline",
    {
      distT: "0",
      distB: "0",
      distL: "0",
      distR: "0",
    },
  ],
  [
    "f",
    {
      aca: "0",
      bx: "0",
      ca: "0",
      del1: "0",
      del2: "0",
      dt2D: "0",
      dtr: "0",
    },
  ],
  ["c", { ph: "0" }],
  [
    "outlinePr",
    {
      applyStyles: "0",
      showOutlineSymbols: "1",
      summaryBelow: "1",
      summaryRight: "1",
    },
  ],
  [
    "definedName",
    {
      function: "0",
      hidden: "0",
      publishToServer: "0",
      vbProcedure: "0",
      workbookParameter: "0",
      xlm: "0",
    },
  ],
  [
    "col",
    {
      bestFit: "0",
      phonetic: "0",
      outlineLevel: "0",
      collapsed: "0",
      hidden: "0",
    },
  ],
  [
    "sheetProtection",
    {
      objects: "0",
      scenarios: "0",
    },
  ],
]);

/** Container elements whose `@count` is derivable from their children — the
 * digest compares children structurally, so the redundant attribute drops. */
const DERIVED_COUNT_ELEMENTS = new Set([
  "borders",
  "cellStyles",
  "cellStyleXfs",
  "cellXfs",
  "dxfs",
  "fills",
  "fonts",
  "numFmts",
  "tableStyle",
  "tableStyles",
]);

/** XSD element-only style containers; stray producer text is not style data. */
const ELEMENT_ONLY_STYLE_CONTAINERS = new Set([
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "dxfs",
  "fills",
  "fonts",
  "numFmts",
]);

/** XSD-false-default attributes sources may write explicitly (`0`/`false`);
 * omission in the rebuilt part is semantically identical. */
const SOURCE_FALSE_DEFAULT_ATTRIBUTES = new Set([
  // AG_RevData @ua (undo) / @ra (rejected) across all revision elements.
  "ua",
  "ra",
  // CT_Style @default/@customStyle are optional CT_OnOff attributes; explicit
  // false has the same style-selection semantics as omission.
  "w:default",
  "w:customStyle",
]);

/** Document-unique revision markers; canonical pairing makes ids positional. */
const REVISION_ELEMENTS = new Set([
  "w:ins",
  "w:del",
  "w:cellIns",
  "w:cellDel",
  "w:cellMerge",
  "w:pPrChange",
  "w:rPrChange",
  "w:tblPrChange",
  "w:tblGridChange",
  "w:trPrChange",
  "w:tcPrChange",
  "w:sectPrChange",
]);

const LEGACY_OFFICE_URI_PREFIX = "http://schemas.microsoft.com/office/2006/relationships/";
const STRICT_URI_PREFIX = "http://purl.oclc.org/ooxml/";
const TRANSITIONAL_URI_PREFIX = "http://schemas.openxmlformats.org/";

const TRANSITIONAL_URI_ALIASES = new Map([
  ["docPropsApp", "officeDocument/2006/relationships/extended-properties"],
  // O12 templates write `…/xlChart` for chart companion relationships.
  ["officeDocument/2006/relationships/xlChart", "officeDocument/2006/relationships/chart"],
  [
    "officeDocument/2006/relationships/extendedProperties",
    "officeDocument/2006/relationships/extended-properties",
  ],
]);

const VERSIONED_TRANSITIONAL_PREFIXES = [
  "drawingml/",
  "presentationml/",
  "spreadsheetml/",
  "wordprocessingml/",
] as const;

function canonicalAttributeValue(name: string, value: string, elementName?: string): string {
  if (
    (elementName === "w:sz" || elementName === "w:szCs") &&
    name === "w:val" &&
    /^[+-]?\d+(?:\.\d+)?(?:mm|cm|in|pt|pc|pi|px)$/.test(value)
  ) {
    return String(convertToPt(value as `${number}pt`) * 2);
  }
  // Percentage-valued CT_TblWidth attributes use fiftieths of a percent when
  // writers omit the literal; normalize before the generic percent scaler.
  if (
    elementName !== undefined &&
    ["w:tblW", "w:tcW", "w:tblCellSpacing", "w:tblInd"].includes(elementName) &&
    name === "w:w" &&
    value.endsWith("%") &&
    Number.isFinite(Number(value.slice(0, -1)))
  ) {
    return String(Number(value.slice(0, -1)) * 50);
  }
  if (value.endsWith("%") && Number.isFinite(Number(value.slice(0, -1)))) {
    return String(Number(value.slice(0, -1)) * 1000);
  }
  if (value === "") return value;
  if (value === "on" || value === "true") return "1";
  if (value === "off" || value === "false") return "0";
  if (value === "t") return "1";
  if (value === "f") return "0";
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) && Number.isFinite(Number(value))) {
    return String(Number(value));
  }
  if (
    ((elementName === "wp:simplePos" && (name === "x" || name === "y")) ||
      (elementName === "a:off" && (name === "x" || name === "y")) ||
      (elementName === "a:ext" && (name === "cx" || name === "cy"))) &&
    /^[+-]?\d+(?:\.\d+)?(?:mm|cm|in|pt|pc|pi|px)$/.test(value)
  ) {
    return String(
      convertToEmu(value as `${number}${"mm" | "cm" | "in" | "pt" | "pc" | "pi" | "px"}`),
    );
  }
  if (
    (elementName === "wp:simplePos" && (name === "x" || name === "y")) ||
    (elementName === "w:pgSz" && (name === "w:w" || name === "w:h")) ||
    (elementName === "w:pgMar" && name.startsWith("w:")) ||
    (elementName === "w:ind" &&
      ["w:start", "w:end", "w:left", "w:right", "w:hanging", "w:firstLine"].includes(name))
  ) {
    if (/^[+-]?\d+(?:\.\d+)?(?:mm|cm|in|pt|pc|pi|px)$/.test(value)) {
      return String(
        convertToTwip(value as `${number}${"mm" | "cm" | "in" | "pt" | "pc" | "pi" | "px"}`),
      );
    }
  }
  if (elementName === "a:buSzPct" && name === "val") {
    const percent = value.endsWith("%")
      ? String(Number(value.slice(0, -1)) * 1000)
      : String(Number(value) / 1000);
    return Number.isFinite(Number(percent)) ? String(Number(percent)) : value;
  }
  if (name === "ht" && Number.isFinite(Number(value))) return String(Number(value));
  if (name === "Type" || name === "uri" || name === "Namespace") {
    if (value.startsWith(STRICT_URI_PREFIX)) {
      const path = value.slice(STRICT_URI_PREFIX.length);
      const versionedPath = VERSIONED_TRANSITIONAL_PREFIXES.some(
        (prefix) =>
          path.startsWith(prefix) && !/^0*(?:\d+\.)*\d+\//.test(path.slice(prefix.length)),
      )
        ? path.replace(/^([^/]+)\//, "$1/2006/")
        : path;
      const transitionalPath = versionedPath.replace(
        "officeDocument/relationships/",
        "officeDocument/2006/relationships/",
      );
      const alias = TRANSITIONAL_URI_ALIASES.get(transitionalPath);
      return `${TRANSITIONAL_URI_PREFIX}${alias ?? transitionalPath}`;
    }
    const legacyPath = value.startsWith(LEGACY_OFFICE_URI_PREFIX)
      ? value.slice(LEGACY_OFFICE_URI_PREFIX.length)
      : undefined;
    if (legacyPath !== undefined) {
      // O12 template aliases for canonical transitional relationship types.
      const legacyAliases: Record<string, string> = {
        xlChart: "chart",
        xlWorksheet: "worksheet",
        xlSharedStrings: "sharedStrings",
        xlStyles: "styles",
        xlCalcChain: "calcChain",
        xlVolatileDependencies: "volatileDependencies",
        xlPrinterSettings: "printerSettings",
      };
      const aliasTarget = legacyAliases[legacyPath];
      if (aliasTarget) {
        return `${TRANSITIONAL_URI_PREFIX}officeDocument/2006/relationships/${aliasTarget}`;
      }
      const alias = TRANSITIONAL_URI_ALIASES.get(legacyPath);
      if (alias) return `${TRANSITIONAL_URI_PREFIX}${alias}`;
      return `${TRANSITIONAL_URI_PREFIX}officeDocument/2006/relationships/${legacyPath}`;
    }
    return value;
  }
  return value;
}

function canonicalBooleanText(value: string): string | undefined {
  if (value === "1" || value === "true" || value === "on" || value === "t") return "1";
  if (value === "0" || value === "false" || value === "off" || value === "f") return "0";
  return undefined;
}

function canonicalElementName(name: string, partPath: string): string {
  if (partPath.startsWith("xl/sharedStrings.xml") && name === "sstItem") return "si";
  if (name === "w16se:symEx") return "w16se:sym";
  if (!partPath.startsWith("docProps/core.xml") && !partPath.startsWith("docProps/app.xml")) {
    return name;
  }
  const localName = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
  const lower = localName.toLowerCase();
  return LEGACY_CORE_PROPERTY_NAMES.get(lower) ?? lower;
}

/**
 * Older producers write created/modified with legacy lowercase element names
 * in docProps/core.xml; both spellings carry the same semantics.
 */
const LEGACY_CORE_PROPERTY_NAMES = new Map([
  ["datecreated", "created"],
  ["datemodified", "modified"],
]);

function relationshipOwnerPath(relsPath: string): string {
  if (relsPath === "_rels/.rels" || relsPath.startsWith("_rels/")) return "";
  const ownerPath = relsPath.replace(/\/_rels\/(?=[^/]+\.rels$)/, "/").replace(/\/_rels$/, "");
  return ownerPath.endsWith(".rels") ? ownerPath.slice(0, -".rels".length) : ownerPath;
}

const ALWAYS_REBUILT_PARTS = new Set(["[Content_Types].xml", "_rels/.rels"]);
const ALWAYS_REBUILT_PARTS_LOWER = new Map(
  [...ALWAYS_REBUILT_PARTS].map((path) => [path.toLowerCase(), path] as const),
);

/**
 * Read ZIP entries with the same OPC collision policy as `ParsedArchive`:
 * the first central-directory entry wins. fflate's plain object result is
 * last-write-wins, which makes the comparator disagree with the parser on
 * adversarial archives containing duplicate local entries.
 */
function unzipOpcEntries(data: Uint8Array): Record<string, Uint8Array> {
  try {
    return unzipSync(new ParsedArchive(data).save());
  } catch {
    return unzipSync(data);
  }
}

/**
 * Parts no relationship reaches, walking the closure from `_rels/.rels`.
 * These are orphaned package members: the document model has no anchor for
 * them, so they travel verbatim instead of masking a model gap.
 */
export function orphanedPackageParts(source: Uint8Array): Set<string> {
  let entries = unzipOpcEntries(source);
  const files = new Map<string, Uint8Array>(
    Object.entries(entries).map(([name, bytes]) => [name.replaceAll("\\", "/"), bytes]),
  );
  const relsCache = new Map<string, Element | undefined>();
  const readRels = (relsPath: string): Element | undefined => {
    if (!relsCache.has(relsPath)) {
      const bytes = files.get(relsPath);
      relsCache.set(relsPath, bytes ? parseCanonicalXml(decodeXmlBytes(bytes)) : undefined);
    }
    return relsCache.get(relsPath);
  };
  const reachable = new Set<string>();
  const queue: string[] = [""];
  while (queue.length > 0) {
    const owner = queue.shift()!;
    const relsPath = owner === "" ? "_rels/.rels" : relationshipPath(owner);
    for (const relationship of readRels(relsPath)?.elements ?? []) {
      if (relationship.name !== "Relationship") continue;
      if (relationship.attributes?.["TargetMode"] === "External") continue;
      const target = relationship.attributes?.["Target"];
      if (!target) continue;
      const resolved = resolveRelationshipTarget(owner, target);
      if (reachable.has(resolved)) continue;
      reachable.add(resolved);
      queue.push(resolved);
    }
  }
  const orphaned = new Set<string>();
  for (const partPath of files.keys()) {
    if (ALWAYS_REBUILT_PARTS.has(partPath)) continue;
    const ownerPart = partPath.endsWith(".rels") ? relationshipOwnerPath(partPath) : partPath;
    if (!reachable.has(ownerPart)) orphaned.add(partPath);
  }
  return orphaned;
}

/**
 * An empty relationships part carries no semantics, so a generator may omit it
 * (matching Office's normalized output) without losing package information.
 */
function isEmptyRelationshipsPart(data: Uint8Array | undefined): boolean {
  if (!data) return true;
  return !/<Relationship[\s/>]/.test(new TextDecoder().decode(data));
}

function canonicalNode(
  element: Element,
  path = "",
  references?: Map<string, string>,
): CanonicalNode {
  const name = element.name ?? "";
  const childPath = `${path}/${name}`;
  const canonicalAttributeName = (attributeName: string): string => {
    // O12 templates write `@defaultWidth` (1/256 char units); the standard
    // `@width` is in character units.
    if ((name === "col" || name.endsWith(":col")) && attributeName === "defaultWidth") {
      return "width";
    }
    // O12 templates write `relId` for the chart relationship reference.
    if (attributeName === "relId") return "r:id";
    // Producers (LibreOffice) write the cellStyle attribute with a capital I;
    // the XSD name is `builtinId`.
    if ((name === "cellStyle" || name.endsWith(":cellStyle")) && attributeName === "builtInId") {
      return "builtinId";
    }
    // O12 templates write `@range`; the XSD attribute is `@ref`.
    if (name === "dimension" && attributeName === "range") return "ref";
    if (name === "sst" && attributeName === "totalCount") return "count";
    if ((name === "w16se:sym" || name === "w16se:symEx") && attributeName === "w16se:char")
      return "w:char";
    if ((name === "w16se:sym" || name === "w16se:symEx") && attributeName === "w16se:font")
      return "w:font";
    // LibreOffice writes the legacy unprefixed spelling on style w:ind; Word
    // reads it as the OOXML w:left attribute.
    if (name === "w:ind" && attributeName === "left") return "w:left";
    // Lenient VML producers write the office drawing attribute unprefixed;
    // our writers always use the canonical o: form.
    if (attributeName === "detectmouseclick") return "o:detectmouseclick";
    return attributeName;
  };
  if (name === "mc:AlternateContent") {
    const branches = element.elements ?? [];
    const branch =
      branches.find(
        (child) =>
          child.name === "mc:Choice" &&
          (child.elements ?? []).some((content) => content.type === "element"),
      ) ??
      branches.find(
        (child) =>
          child.name === "mc:Fallback" &&
          (child.elements ?? []).some((content) => content.type === "element"),
      );
    const content = branch?.elements?.find((child) => child.type === "element");
    if (content) return canonicalNode(content, `${childPath}/${branch?.name}`, references);
  }
  if (name === "mc:Choice") {
    const content = (element.elements ?? []).find((child) => child.type === "element");
    if (content) return canonicalNode(content, childPath, references);
  }
  // Lenient producers write the section type as a schema-invalid @w:type
  // attribute; Word reads it the same as the <w:type w:val/> child. Normalize
  // to the child form so both spellings compare equal.
  if ((name === "w:sectPr" || name.endsWith(":sectPr")) && element.attributes?.["w:type"]) {
    const sectionType = element.attributes["w:type"]!;
    const withoutAttr = { ...element, attributes: { ...element.attributes } };
    delete withoutAttr.attributes!["w:type"];
    const hasTypeChild = (withoutAttr.elements ?? []).some((child) => {
      const childName = child.name ?? "";
      return childName === "w:type" || childName.endsWith(":type");
    });
    if (!hasTypeChild) {
      withoutAttr.elements = [
        ...(withoutAttr.elements ?? []),
        {
          type: "element" as const,
          name: "w:type",
          attributes: { "w:val": sectionType },
          cdata: undefined,
        },
      ];
    }
    return canonicalNode(withoutAttr, path, references);
  }
  const isRelationship = name === "Relationship" && path.includes("_rels/");
  const relsPath = name === "Relationship" ? path.slice(0, path.lastIndexOf("/")) : path;
  const ownerPath = isRelationship ? relationshipOwnerPath(relsPath) : path;
  const attributes: Record<string, string> = Object.fromEntries(
    Object.entries(element.attributes ?? {})
      .filter(([attributeName]) => attributeName !== "xmlns" && !attributeName.startsWith("xmlns:"))
      .filter(([attributeName]) => !(isRelationship && attributeName === "Id"))
      .filter(([attributeName]) => !IGNORED_ATTRIBUTES.has(attributeName))
      .filter(
        ([attributeName]) =>
          !(
            (name === "c:f" ||
              name.endsWith(":f") ||
              name === "c:formatCode" ||
              name.endsWith(":formatCode")) &&
            attributeName === "xml:space"
          ),
      )
      // O12 sheetView `@active` is an unused selection hint absent from the
      // transitional XSD; writers never re-emit it.
      .filter(
        ([attributeName]) =>
          !((name === "sheetView" || name.endsWith(":sheetView")) && attributeName === "active"),
      )
      // Producers (LibreOffice) write empty-valued side attributes on border
      // elements; CT_Border has no such attributes and an empty value carries
      // no semantic content, so both sides drop them.
      .filter(
        ([attributeName, value]) =>
          !(
            (name === "border" || name.endsWith(":border")) &&
            ["left", "right", "top", "bottom", "diagonal"].includes(attributeName) &&
            value === ""
          ),
      )
      .map(([rawAttributeName, value]) => {
        const attributeName = canonicalAttributeName(rawAttributeName);
        return [
          attributeName,
          references && attributeName.startsWith("r:")
            ? (references.get(String(value ?? "")) ??
              canonicalAttributeValue(attributeName, String(value ?? ""), name))
            : isRelationship &&
                attributeName === "Target" &&
                element.attributes?.TargetMode !== "External"
              ? resolveRelationshipTarget(ownerPath, String(value ?? ""))
              : canonicalAttributeValue(attributeName, String(value ?? ""), name),
        ] as const;
      })
      .sort(([left], [right]) => left.localeCompare(right)),
  );
  if (ON_OFF_ELEMENTS.has(name)) {
    const valueAttribute = name.startsWith("m:") ? "m:val" : "w:val";
    if (!(valueAttribute in attributes) || attributes[valueAttribute] === "")
      attributes[valueAttribute] = "1";
  }
  for (const [attributeName, value] of Object.entries(DEFAULT_ATTRIBUTES.get(name) ?? {})) {
    if (!(attributeName in attributes)) attributes[attributeName] = value;
  }
  let orderedAttributes = Object.fromEntries(
    Object.entries(attributes).sort(([left], [right]) => left.localeCompare(right)),
  );
  // O12 templates write `@defaultWidth` in 1/256 char units; normalize to the
  // standard character-unit value so both forms compare equal.
  if (name === "col" && orderedAttributes.width !== undefined) {
    const width = Number(orderedAttributes.width);
    if (Number.isFinite(width) && width > 256) {
      orderedAttributes.width = String(Math.round((width / 256) * 1000) / 1000);
    }
  }
  // O12 templates reference charts via `a:chart` inside a legacy ChartML
  // graphicData URI; the writer emits the canonical `c:chart` form.
  if ((name.split(":").pop() ?? name) === "graphicData") {
    const uri = element.attributes?.uri;
    if (typeof uri === "string" && uri.includes("ChartML")) {
      orderedAttributes.uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    }
  }
  const localName = name.split(":").pop() ?? name;
  if (localName === "color" || localName.endsWith("Color")) {
    const legacyType = orderedAttributes.type;
    const legacyValue = orderedAttributes.val;
    if (legacyType === "theme" && legacyValue !== undefined) orderedAttributes.theme = legacyValue;
    if ((legacyType === "indexed" || legacyType === "icv") && legacyValue !== undefined) {
      orderedAttributes.indexed = legacyValue;
    }
    if (legacyType === "rgb" && legacyValue !== undefined) orderedAttributes.rgb = legacyValue;
    if (legacyType !== undefined && legacyValue !== undefined) {
      delete orderedAttributes.type;
      delete orderedAttributes.val;
    }
    orderedAttributes = Object.fromEntries(
      Object.entries(orderedAttributes).sort(([left], [right]) => left.localeCompare(right)),
    );
  }
  if (DERIVED_COUNT_ELEMENTS.has(name)) delete orderedAttributes.count;
  const elementText = (element.elements ?? [])
    .filter((child) => child.type === "text" || child.type === "cdata")
    .map((child) => String(child.text ?? child.cdata ?? ""))
    .join("");
  const rawText = ELEMENT_ONLY_STYLE_CONTAINERS.has(name) ? "" : elementText;
  const text =
    element.name === "v" && rawText !== "" && Number.isFinite(Number(rawText))
      ? String(Number(rawText))
      : element.name?.startsWith("vt:") &&
          /^(?:\d{4}-\d{2}-\d{2})[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?$/.test(
            rawText,
          )
        ? new Date(rawText).toISOString()
        : BOOLEAN_TEXT_ELEMENTS.has(element.name ?? "")
          ? (canonicalBooleanText(rawText) ?? rawText)
          : rawText;
  const mappedChildren = (element.elements ?? [])
    .filter((child): child is Element => child.type === "element")
    .filter((child) => !(child.name ?? "").startsWith("!"))
    // `c:lastLayout` is a render cache that Office recomputes; the
    // transitional XSD does not model it and writers never re-emit it.
    .filter((child) => child.name !== "c:lastLayout" && child.name !== "c:lastLayoutOuter")
    .map((child) => {
      const node = canonicalNode(child, childPath, references);
      // O12 templates write `a:chart` (main DrawingML ns) instead of `c:chart`
      // for the chart reference inside legacy ChartML graphicData.
      if ((name.split(":").pop() ?? name) === "graphicData" && node.name?.endsWith(":chart")) {
        return { ...node, name: "c:chart" };
      }
      // Comment-pr compatibility fallbacks omit the spreadsheetDrawing prefix
      // on anchor corners; the canonical CT_ObjectAnchor uses xdr:from/to.
      if ((name.split(":").pop() ?? name) === "anchor") {
        const localChildName = node.name?.split(":").pop() ?? node.name;
        if (localChildName === "from" || localChildName === "to") {
          return { ...node, name: `xdr:${localChildName}` };
        }
      }
      return node;
    });
  // Legacy documents may carry bare text directly inside w:r. Word treats it
  // as run text and normalizes it into w:t on save, so the digest normalizes
  // it the same way instead of comparing structurally different encodings.
  const bareRunText =
    name === "w:r" && rawText.trim() !== "" && !mappedChildren.some((child) => child.name === "w:t")
      ? rawText.trim()
      : "";
  if (bareRunText !== "") {
    mappedChildren.push({ name: "w:t", attributes: {}, text: bareRunText, children: [] });
  }
  // Some producers park style-level flags in the paragraph-properties
  // container. Promote that copy to the canonical style location before
  // removing the schema-invalid stray.
  if (name === "w:style") {
    const styleQFormat = mappedChildren.find((child) => child.name === "w:qFormat");
    const paragraphProperties = mappedChildren.find((child) => child.name === "w:pPr");
    const parkedQFormat = paragraphProperties?.children.find((child) => child.name === "w:qFormat");
    if (paragraphProperties && parkedQFormat) {
      const promoted = !styleQFormat;
      const paragraphIndex = mappedChildren.indexOf(paragraphProperties);
      const cleaned = {
        ...paragraphProperties,
        children: paragraphProperties.children.filter((child) => child !== parkedQFormat),
      };
      mappedChildren[paragraphIndex] = cleaned;
      if (promoted) mappedChildren.push(parkedQFormat);
    }
  }
  // CT_StrData/CT_NumData require c:ptCount; some producers omit it. The
  // writer always emits it, so the digest fills the default when missing.
  if (
    localName === "strCache" ||
    localName === "numCache" ||
    localName === "strLit" ||
    localName === "numLit" ||
    localName === "multiLvlStrCache"
  ) {
    const pointCount = mappedChildren.find((child) => child.name === "c:ptCount");
    if (!pointCount) {
      const pointCountValue = mappedChildren.filter((child) =>
        (child.name ?? "").endsWith(":pt"),
      ).length;
      mappedChildren.unshift({
        name: "c:ptCount",
        attributes: { val: String(pointCountValue) },
        text: "",
        children: [],
      });
    }
  }
  // Word tolerates a stray nested empty w:pPr inside w:pPr (schema-invalid
  // legacy input) by ignoring it — normalize it out instead of comparing
  // structurally invalid source markup against our valid output.
  // Legacy producers can wrap run content in a schema-invalid nested w:r.
  // Word flattens that wrapper on read; preserve inner properties/content but
  // not the wrapper itself when comparing normalized semantic output.
  const flattenedRunChildren =
    name === "w:r"
      ? mappedChildren.flatMap((child) => {
          if (child.name !== "w:r") return [child];
          const nestedProperties = child.children.filter((item) => item.name === "w:rPr");
          const nestedContent = child.children.filter((item) => item.name !== "w:rPr");
          return mappedChildren.some((item) => item.name === "w:rPr") ||
            nestedProperties.length === 0
            ? nestedContent
            : [...nestedProperties, ...nestedContent];
        })
      : mappedChildren;
  const childrenOut =
    name === "w:pPr"
      ? flattenedRunChildren.filter(
          (child) => !(child.name === "w:pPr" && child.children.length === 0 && child.text === ""),
        )
      : flattenedRunChildren;
  // CT_Tbl requires tblPr, CT_Tc requires a block-level child, and
  // CT_PPrChange requires its historical pPr. The canonical writer fills
  // these producer omissions with the schema-required empty structure.
  if (name === "w:tbl" && !childrenOut.some((child) => child.name === "w:tblPr")) {
    childrenOut.unshift({ name: "w:tblPr", attributes: {}, text: "", children: [] });
  }
  if (
    name === "w:tc" &&
    !childrenOut.some(
      (child) => child.name === "w:p" || child.name === "w:tbl" || child.name === "w:sdt",
    )
  ) {
    childrenOut.push({ name: "w:p", attributes: {}, text: "", children: [] });
  }
  if (name === "w:pPrChange" && !childrenOut.some((child) => child.name === "w:pPr")) {
    childrenOut.push({ name: "w:pPr", attributes: {}, text: "", children: [] });
  }
  // Word also tolerates duplicate w:rPr children inside w:r by applying the
  // last one — normalize to the same precedence instead of comparing the
  // structurally invalid source markup against our valid output.
  if (name === "w:r") {
    const rPrIndices = childrenOut
      .map((child, index) => (child.name === "w:rPr" ? index : -1))
      .filter((index) => index >= 0);
    if (rPrIndices.length > 1) {
      const keep = rPrIndices[rPrIndices.length - 1]!;
      const dropped = new Set(rPrIndices.filter((index) => index !== keep));
      const deduped = childrenOut.filter((_, index) => !dropped.has(index));
      childrenOut.length = 0;
      childrenOut.push(...deduped);
    }
  }
  // CT_SdtDocPartObj allows one gallery and one unique flag. Some producers
  // duplicate both; Word applies the first pair, matching our typed parser.
  if (name === "w:docPartObj") {
    const seen = new Set<string>();
    const deduped = childrenOut.filter((child) => {
      if (child.name !== "w:docPartGallery" && child.name !== "w:docPartUnique") return true;
      if (seen.has(child.name)) return false;
      seen.add(child.name);
      return true;
    });
    childrenOut.length = 0;
    childrenOut.push(...deduped);
  }
  // Some legacy producers wrap the entire numbering-level run-property set in
  // a second empty w:rPr. Unwrap the redundant container because Word resolves
  // to the inner properties and our typed writer emits the normalized form.
  if (name === "w:rPr") {
    const nested = mappedChildren.filter((child) => child.name === "w:rPr");
    if (nested.length === 1) {
      mappedChildren.length = 0;
      mappedChildren.push(...nested[0]!.children);
    }
  }
  return {
    name: canonicalElementName(name, path),
    attributes: orderedAttributes,
    text:
      bareRunText !== ""
        ? ""
        : (element.attributes?.["xml:space"] === "preserve" ||
              ["w:t", "w:delText", "w:instrText", "w:delInstrText", "m:t"].includes(name)) &&
            mappedChildren.length === 0
          ? text
          : text.trim(),
    children: (() => {
      if (
        name === "text" &&
        childrenOut.length === 1 &&
        childrenOut[0].name === "t" &&
        childrenOut[0].text === "" &&
        childrenOut[0].children.length === 0 &&
        Object.keys(childrenOut[0].attributes).length === 0
      ) {
        return [];
      }
      return childrenOut;
    })(),
  };
}

function sortUnorderedChildren(
  node: CanonicalNode,
  path: string,
  references?: Map<string, string>,
  ignorablePrefixes?: ReadonlySet<string>,
): CanonicalNode {
  const childPath = `${path}/${node.name}`;
  let children = node.children.map((child) =>
    sortUnorderedChildren(child, childPath, references, ignorablePrefixes),
  );
  const localName = (node.name.split(":").pop() ?? node.name).toLowerCase();
  // LibreOffice can leave schema-invalid direct properties on document
  // default containers. Valid CT_PPrDefault/CT_RPrDefault carry only the
  // nested properties container, so ignore the stray legacy sibling rather
  // than inventing a non-OOXML model for producer debris.
  if (localName === "pprdefault" || localName === "rprdefault") {
    children = children.filter(
      (child) => child.name === (localName === "pprdefault" ? "w:pPr" : "w:rPr"),
    );
  }
  if (localName === "tblcellmar") {
    const seen = new Set<string>();
    children = children.filter((child) => {
      const key = child.name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
  // Property containers are strict XSD sequences, but their child order is
  // representation only: canonical parsers flatten them and writers rebuild
  // the valid sequence. Sort these comparison-only containers so damaged or
  // lenient producer order does not mask real content differences.
  if (
    ["ppr", "numpr", "compat", "num", "lvl", "docpartpr", "footnotes", "endnotes"].includes(
      localName,
    )
  ) {
    if (localName === "footnotes" || localName === "endnotes") {
      const noteId = (child: CanonicalNode): number => {
        const value = Number(child.attributes["w:id"] ?? child.attributes.id ?? Number.NaN);
        return Number.isFinite(value) ? value : Number.MAX_SAFE_INTEGER;
      };
      children.sort((left, right) => noteId(left) - noteId(right));
      return { ...node, children };
    }
    const fingerprints = new Map(
      children.map((child) => [child, semanticChildFingerprint(child, ignorablePrefixes)]),
    );
    children = [...children].sort((left, right) =>
      (fingerprints.get(left) ?? "").localeCompare(fingerprints.get(right) ?? ""),
    );
    if (localName === "footnotes" || localName === "endnotes") {
      return { ...node, children };
    }
  }
  // OPC relationships are a set of graph edges. Distinct ids are an XML
  // addressing detail; two edges with the same type, mode, and target are
  // semantically redundant and may be normalized to one shared relationship.
  if (path.endsWith(".rels") && localName === "relationships") {
    const seen = new Set(
      children.map((child) => semanticChildFingerprint(child, ignorablePrefixes)),
    );
    if (seen.size < children.length) {
      const kept = new Set<string>();
      children = children.filter((child) => {
        const fingerprint = semanticChildFingerprint(child, ignorablePrefixes);
        if (kept.has(fingerprint)) return false;
        kept.add(fingerprint);
        return true;
      });
    }
  }
  // Workbook and stylesheet roots are strict XSD sequences. Normalize legacy
  // producer order before comparison; child identity and content still differ.
  if (localName === "abstractnum") {
    const order = ["nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl"];
    const orderIndex = new Map(order.map((name, index) => [name, index]));
    children = [...children].sort((left, right) => {
      const leftName = left.name.split(":").pop() ?? left.name;
      const rightName = right.name.split(":").pop() ?? right.name;
      return (orderIndex.get(leftName) ?? 99) - (orderIndex.get(rightName) ?? 99);
    });
    return { ...node, children };
  }
  if (localName === "workbook" || localName === "styles" || localName === "stylesheet") {
    const order =
      localName === "workbook"
        ? [
            "fileVersion",
            "fileSharing",
            "workbookPr",
            "workbookProtection",
            "bookViews",
            "sheets",
            "functionGroups",
            "externalReferences",
            "definedNames",
            "calcPr",
            "oleSize",
            "customWorkbookViews",
            "pivotCaches",
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
          ]
        : [
            "numFmts",
            "fonts",
            "fills",
            "borders",
            "cellStyleXfs",
            "cellXfs",
            "cellStyles",
            "dxfs",
            "tableStyles",
            "colors",
            "extLst",
          ];
    const orderIndex = new Map(order.map((name, index) => [name, index]));
    children = [...children].sort((left, right) => {
      const leftName = left.name.split(":").pop() ?? left.name;
      const rightName = right.name.split(":").pop() ?? right.name;
      return (orderIndex.get(leftName) ?? 99) - (orderIndex.get(rightName) ?? 99);
    });
    return { ...node, children };
  }
  if (localName === "worksheet") {
    const order = [
      "sheetPr",
      "dimension",
      "sheetViews",
      "sheetFormatPr",
      "cols",
      "sheetData",
      "sheetCalcPr",
      "sheetProtection",
      "protectedRanges",
      "scenarios",
      "autoFilter",
      "sortState",
      "dataConsolidate",
      "customSheetViews",
      "mergeCells",
      "phoneticPr",
      "conditionalFormatting",
      "dataValidations",
      "hyperlinks",
      "printOptions",
      "pageMargins",
      "pageSetup",
      "headerFooter",
      "rowBreaks",
      "colBreaks",
      "customProperties",
      "cellWatches",
      "ignoredErrors",
      "smartTags",
      "drawing",
      "drawingHF",
      "picture",
      "oleObjects",
      "controls",
      "webPublishItems",
      "tableParts",
      "extLst",
    ];
    const orderIndex = new Map(order.map((name) => [name, order.indexOf(name)]));
    children = [...children].sort((left, right) => {
      const leftName = left.name.split(":").pop() ?? left.name;
      const rightName = right.name.split(":").pop() ?? right.name;
      return (
        (orderIndex.get(leftName) ?? order.length) - (orderIndex.get(rightName) ?? order.length)
      );
    });
    return { ...node, children };
  }
  // Some producers emit CT_DPr children in a schema-invalid order. The XSD
  // order is semantic for OOXML, so compare against its canonical sequence.
  if (localName === "dpr") {
    const order = new Map([
      ["m:begChr", 0],
      ["m:sepChr", 1],
      ["m:endChr", 2],
      ["m:grow", 3],
      ["m:shp", 4],
      ["m:ctrlPr", 5],
    ]);
    children = [...children].sort(
      (left, right) => (order.get(left.name) ?? 99) - (order.get(right.name) ?? 99),
    );
    return { ...node, children };
  }
  // CT_DocPartPr is an xsd:all container, while producers may write its
  // optional children in any order. Canonicalize only comparison order.
  if (localName === "docpartpr") {
    const order = ["name", "style", "category", "types", "behaviors", "description", "guid"];
    const orderIndex = new Map(order.map((name, index) => [name, index]));
    children = [...children].sort(
      (left, right) =>
        (orderIndex.get(left.name.split(":").pop() ?? left.name) ?? 99) -
        (orderIndex.get(right.name.split(":").pop() ?? right.name) ?? 99),
    );
    return { ...node, children };
  }
  // Border containers are XSD sequences, but damaged producers may emit
  // optional sides out of order. Canonicalize only comparison order while
  // retaining every side and its full content.
  if (["pbdr", "tcborders", "tblborders", "tblcellmar"].includes(localName)) {
    const order = new Map(
      [
        "top",
        "start",
        "left",
        "bottom",
        "end",
        "right",
        "insideh",
        "insidev",
        "between",
        "bar",
      ].map((name, index) => [name, index]),
    );
    children = [...children].sort(
      (left, right) =>
        (order.get((left.name.split(":").pop() ?? left.name).toLowerCase()) ?? 99) -
        (order.get((right.name.split(":").pop() ?? right.name).toLowerCase()) ?? 99),
    );
    return { ...node, children };
  }
  // Some legacy producers reorder the CT_Tcg sequence. Keep schema-valid
  // writer order canonical while still comparing every child and its content.
  if (localName === "tcg") {
    const order = ["keymaps", "keymapsBad", "toolbars", "acds"];
    const orderIndex = new Map(order.map((name, index) => [name.toLowerCase(), index]));
    children = [...children].sort(
      (left, right) =>
        (orderIndex.get((left.name.split(":").pop() ?? left.name).toLowerCase()) ?? 99) -
        (orderIndex.get((right.name.split(":").pop() ?? right.name).toLowerCase()) ?? 99),
    );
    return { ...node, children };
  }
  // Damaged legacy producers can repeat the singleton cell-properties
  // container. Word applies the final declaration, so canonicalize both sides
  // to that last-wins value instead of preserving the invalid duplicate.
  if (localName === "tc" && children.filter((child) => child.name === "w:tcPr").length > 1) {
    children = children.filter(
      (child, index) =>
        child.name !== "w:tcPr" ||
        index === children.findLastIndex((item) => item.name === "w:tcPr"),
    );
    return { ...node, children };
  }
  // Some chart producers emit CT_DLbls children outside the XSD sequence. The
  // writer rebuilds the valid sequence, so normalize only ordering before the
  // semantic comparison.
  if (localName === "dlbls") {
    const order = new Map([
      ["dLbl", 0],
      ["numFmt", 1],
      ["spPr", 2],
      ["txPr", 3],
      ["dLblPos", 4],
      ["showLegendKey", 5],
      ["showVal", 6],
      ["showCatName", 7],
      ["showSerName", 8],
      ["showPercent", 9],
      ["showBubbleSize", 10],
      ["separator", 11],
      ["showLeaderLines", 12],
      ["leaderLines", 13],
      ["extLst", 14],
    ]);
    children = [...children].sort((left, right) => {
      const leftName = left.name.split(":").pop() ?? left.name;
      const rightName = right.name.split(":").pop() ?? right.name;
      return (order.get(leftName) ?? 99) - (order.get(rightName) ?? 99);
    });
    return { ...node, children };
  }
  // Style and numbering roots are ordered containers, but duplicate keyed
  // definitions are replacement semantics rather than distinct list items.
  if (localName === "styles" || localName === "numbering") {
    const identityByChildName = new Map([
      ["w:style", "w:styleId"],
      ["w:abstractNum", "w:abstractNumId"],
      ["w:num", "w:numId"],
      ["w:numPicBullet", "w:numPicBulletId"],
    ]);
    const keepLastIndex = new Map<string, number>();
    children.forEach((child, index) => {
      const identity = identityByChildName.get(child.name);
      const value = identity === undefined ? undefined : child.attributes[identity];
      if (identity === undefined || value === undefined) return;
      keepLastIndex.set(`${child.name}\0${value}`, index);
    });
    if (keepLastIndex.size > 0) {
      children = children.filter((child, index) => {
        const identity = identityByChildName.get(child.name);
        const value = identity === undefined ? undefined : child.attributes[identity];
        if (identity === undefined || value === undefined) return true;
        return keepLastIndex.get(`${child.name}\0${value}`) === index;
      });
    }
    return { ...node, children };
  }
  const unorderedRoot =
    UNORDERED_PART_PATHS.has(path) || path === "[Content_Types].xml" || path.endsWith(".rels");
  if (
    unorderedRoot ||
    localName === "footnotes" ||
    localName === "endnotes" ||
    localName === "fonts" ||
    localName === "font" ||
    localName === "styles" ||
    localName === "docpartpr" ||
    localName === "settings" ||
    localName === "compat" ||
    localName === "ppr" ||
    localName === "rpr" ||
    localName === "gslst" ||
    localName === "schemeclr" ||
    localName === "ser" ||
    localName === "chart" ||
    localName === "catax" ||
    localName === "title" ||
    localName === "piechart" ||
    localName === "valax" ||
    localName === "dataax" ||
    localName === "serax" ||
    localName === "dateax" ||
    localName === "sectpr" ||
    // CT_TcPr/CT_TrPr children are xsd:choice maxOccurs=unbounded — order-free.
    localName === "tcpr" ||
    localName === "trpr"
  ) {
    // CT_PPrBase/CT_RPr/CT_Settings children are schema singletons; Word
    // applies the last duplicate. Normalize the same way instead of
    // comparing structurally invalid duplicate source markup against the
    // single element our model round-trips.
    if (localName === "ppr" || localName === "rpr" || localName === "settings") {
      const compatIndexes = children
        .map((child, index) => (child.name === "w:compat" ? index : -1))
        .filter((index) => index >= 0);
      if (compatIndexes.length > 1) {
        const mergedCompat: CanonicalNode = {
          ...children[compatIndexes[0]!]!,
          children: compatIndexes.flatMap((index) => children[index]!.children),
        };
        const dropped = new Set(compatIndexes.slice(1));
        children[compatIndexes[0]!] = mergedCompat;
        children = children.filter((_, index) => !dropped.has(index));
      }
      const keepLastIndex = new Map<string, number>();
      children.forEach((child, index) => {
        if (child.name) keepLastIndex.set(child.name, index);
      });
      if (keepLastIndex.size < children.length) {
        children = children.filter(
          (child, index) => child.name !== undefined && keepLastIndex.get(child.name) === index,
        );
      }
    }
    // Pandoc (and other producers) can emit the same w:style twice with the
    // same type + styleId; Word applies the last definition. Normalize the
    // duplicate tail the same way instead of comparing the unreferenced
    // earlier copies as unmatched source entries.
    if (localName === "styles") {
      const keepLastIndex = new Map<string, number>();
      children.forEach((child, index) => {
        if (child.name !== "w:style") return;
        keepLastIndex.set(
          `${child.attributes["w:type"] ?? ""}\0${child.attributes["w:styleId"] ?? ""}`,
          index,
        );
      });
      if (keepLastIndex.size < children.filter((child) => child.name === "w:style").length) {
        children = children.filter(
          (child, index) =>
            child.name !== "w:style" ||
            keepLastIndex.get(
              `${child.attributes["w:type"] ?? ""}\0${child.attributes["w:styleId"] ?? ""}`,
            ) === index,
        );
      }
    }
    // Word folds a schema-invalid settings-root w:compatSetting into the
    // w:compat element on save — normalize the stray entry the same way
    // instead of comparing structurally invalid source markup against our
    // valid output.
    if (localName === "settings") {
      const stray = children.filter((child) => child.name === "w:compatSetting");
      if (stray.length > 0) {
        const compatIndex = children.findIndex((child) => child.name === "w:compat");
        if (compatIndex >= 0) {
          const compat = children[compatIndex]!;
          const known = new Set(
            compat.children
              .filter((child) => child.name === "w:compatSetting")
              .map((child) => child.attributes["w:name"]),
          );
          const moved = stray.filter((child) => !known.has(child.attributes["w:name"]));
          const remaining = children.filter((child) => child.name !== "w:compatSetting");
          const mergedIndex = remaining.findIndex((child) => child.name === "w:compat");
          const mergedCompat =
            moved.length > 0 ? { ...compat, children: [...compat.children, ...moved] } : compat;
          mergedCompat.children.sort(
            (left, right) =>
              childSortKey(left, ignorablePrefixes).localeCompare(
                childSortKey(right, ignorablePrefixes),
              ) ||
              semanticChildFingerprint(left, ignorablePrefixes).localeCompare(
                semanticChildFingerprint(right, ignorablePrefixes),
              ),
          );
          remaining[mergedIndex] = mergedCompat;
          children = remaining;
        }
      }
    }
    children.sort(
      (left, right) =>
        childSortKey(left, ignorablePrefixes).localeCompare(
          childSortKey(right, ignorablePrefixes),
        ) ||
        semanticChildFingerprint(left, ignorablePrefixes).localeCompare(
          semanticChildFingerprint(right, ignorablePrefixes),
        ),
    );
    return {
      ...node,
      children,
    };
  }
  return { ...node, children };
}

function childFingerprint(node: CanonicalNode): string {
  const hash = createHash("sha256");
  hash.update(`${node.name}\0${stableAttributes(node.attributes)}\0${JSON.stringify(node.text)}`);
  for (const child of node.children) hash.update(childFingerprint(child));
  return hash.digest("hex");
}

function writerDefaultOmittedAttributes(
  node: CanonicalNode,
  ignorablePrefixes: ReadonlySet<string> | undefined,
): Record<string, string> {
  const localName = node.name.split(":").pop() ?? node.name;
  const defaults =
    WRITER_EXPLICIT_DEFAULTS.get(node.name) ?? WRITER_EXPLICIT_DEFAULTS.get(localName);
  return Object.fromEntries(
    Object.entries(node.attributes).filter(
      ([name, value]) =>
        !isIgnorableForeignAttr(name, ignorablePrefixes) &&
        name !== "xml:space" &&
        !(defaults && defaults[name] === value) &&
        !(SOURCE_FALSE_DEFAULT_ATTRIBUTES.has(name) && (value === "0" || value === "false")),
    ),
  );
}

function semanticNodeText(node: CanonicalNode): string {
  const isTextElement = ["w:t", "w:delText", "w:instrText", "w:delInstrText", "m:t"].includes(
    node.name,
  );
  return isTextElement && node.text.trim() === "" ? "" : node.text;
}

export function semanticChildFingerprint(
  node: CanonicalNode,
  ignorablePrefixes: ReadonlySet<string> | undefined,
): string {
  const hash = createHash("sha256");
  const allAttributes = writerDefaultOmittedAttributes(node, ignorablePrefixes);
  const attributes = Object.fromEntries(
    Object.entries(allAttributes).filter(
      ([name]) => !(name === "w:id" && REVISION_ELEMENTS.has(node.name)),
    ),
  );
  hash.update(
    `${node.name}\0${stableAttributes(attributes)}\0${JSON.stringify(semanticNodeText(node))}`,
  );
  for (const child of node.children) {
    if (isIgnorableForeignAttr(child.name, ignorablePrefixes)) continue;
    hash.update(semanticChildFingerprint(child, ignorablePrefixes));
  }
  return hash.digest("hex");
}

/**
 * Stable ordering key for unordered containers. Recursive normalized
 * fingerprints pair identical subtrees first; when a subtree differs, the
 * child comparator reports the real content diff instead of a phantom reorder.
 */
function childSortKey(node: CanonicalNode, ignorablePrefixes?: ReadonlySet<string>): string {
  const localName = node.name.split(":").pop() ?? node.name;
  if (localName === "ser" || localName === "dPt") {
    const identity = node.children.find(
      (child) => (child.name.split(":").pop() ?? child.name) === "idx",
    );
    if (identity) return `${node.name}:idx=${identity.attributes["val"] ?? ""}`;
  }
  const hash = createHash("sha256");
  const attributes = writerDefaultOmittedAttributes(node, ignorablePrefixes);
  if (REVISION_ELEMENTS.has(node.name)) delete attributes["w:id"];
  hash.update(
    `${node.name}\0${stableAttributes(attributes)}\0${JSON.stringify(semanticNodeText(node))}`,
  );
  for (const child of node.children) {
    const childAttributes = Object.fromEntries(
      Object.entries(child.attributes).filter(
        ([name]) => !isIgnorableForeignAttr(name, ignorablePrefixes),
      ),
    );
    if (REVISION_ELEMENTS.has(child.name)) delete childAttributes["w:id"];
    hash.update(semanticChildFingerprint(child, ignorablePrefixes));
  }
  return hash.digest("hex");
}

function stableAttributes(attributes: Record<string, string>): string {
  return JSON.stringify(
    Object.fromEntries(
      Object.entries(attributes).sort(([left], [right]) =>
        left < right ? -1 : left > right ? 1 : 0,
      ),
    ),
  );
}

function canonicalAttributes(element: Element): string {
  return Object.entries(element.attributes ?? {})
    .filter(([name]) => name !== "xmlns" && !name.startsWith("xmlns:"))
    .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))
    .map(([name, value]) => `${name}\0${value ?? ""}`)
    .join("\u0001");
}

function canonicalText(element: Element): string {
  const rawText = (element.elements ?? [])
    .filter((child) => child.type === "text" || child.type === "cdata")
    .map((child) => String(child.text ?? child.cdata ?? ""))
    .join("");
  const text =
    element.name === "v" && rawText !== "" && Number.isFinite(Number(rawText))
      ? String(Number(rawText))
      : rawText;
  const isTextElement = ["w:t", "w:delText", "w:instrText", "w:delInstrText", "m:t"].includes(
    element.name,
  );
  return element.attributes?.["xml:space"] === "preserve" || isTextElement ? text : text.trim();
}

function decodeXmlBytes(bytes: Uint8Array): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
    return new TextDecoder("utf-8").decode(bytes.subarray(3));
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe)
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff)
    return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  const ascii = Buffer.from(bytes.subarray(0, 120)).toString("latin1");
  const declaration = /^<\?xml\s+[^>]*encoding=["']([^"']+)["']/i.exec(ascii);
  if (declaration) {
    const encoding = declaration[1]!.toLowerCase();
    if (encoding === "utf-16" || encoding === "utf16") {
      if (bytes[0] === 0 && bytes[1] !== 0) return new TextDecoder("utf-16be").decode(bytes);
      return new TextDecoder("utf-16le").decode(bytes);
    }
  }
  return new TextDecoder("utf-8").decode(bytes);
}

export function canonicalXmlDigest(element: Element | undefined): string {
  if (!element || element.type !== "element")
    return createHash("sha256").update("empty").digest("hex");
  const hash = createHash("sha256");
  const visit = (node: Element): void => {
    hash.update(`<${node.name ?? ""}|${canonicalAttributes(node)}|${canonicalText(node)}>`);
    for (const child of node.elements ?? []) {
      if (child.type === "element") visit(child);
    }
  };
  visit(element);
  return hash.digest("hex");
}

export function canonicalXmlNodes(
  element: Element | undefined,
  partPath = "",
  references?: Map<string, string>,
): CanonicalNode | undefined {
  const ignorablePrefixes = new Set(
    (element?.attributes?.["mc:Ignorable"] ?? "")
      .split(/[\s,]+/)
      .map((prefix) => prefix.trim())
      .filter(Boolean),
  );
  const root =
    element && element.type === "element"
      ? canonicalNode(element, partPath, references)
      : undefined;
  return root && sortUnorderedChildren(root, partPath, references, ignorablePrefixes);
}

function attributeKey(node: CanonicalNode): string {
  return Object.entries(node.attributes)
    .map(([name, value]) => `${name}=${JSON.stringify(value)}`)
    .sort()
    .join(",");
}

function isIgnorableForeignAttr(name: string, prefixes: ReadonlySet<string> | undefined): boolean {
  if (!prefixes || prefixes.size === 0) return false;
  const colon = name.indexOf(":");
  return colon > 0 && prefixes.has(name.slice(0, colon));
}

function childKey(node: CanonicalNode): string {
  return childFingerprint(node);
}

function childDiagnostic(node: CanonicalNode): string {
  const attributes = Object.entries(node.attributes).map(
    ([name, value]) => `${name}=${JSON.stringify(value)}`,
  );
  const childNames = [...new Set(node.children.map((child) => child.name))];
  return [
    node.name,
    attributes.length > 0 ? `{ ${attributes.join(", ")} }` : undefined,
    node.text !== "" ? `text=${JSON.stringify(node.text)}` : undefined,
    childNames.length > 0 ? `children=[${childNames.join(", ")}]` : undefined,
  ]
    .filter(Boolean)
    .join(" ");
}

function compareNodes(
  path: string,
  source: CanonicalNode | undefined,
  output: CanonicalNode | undefined,
  ignorablePrefixes?: ReadonlySet<string>,
  packageParts?: { source: ReadonlySet<string>; output: ReadonlySet<string> },
): SemanticPartDiff["detail"][] {
  if (!source && !output) return [];
  const location = path || "/";
  // Minimal source packages may omit standard companion parts that every
  // writer emits; the generated package is a semantic superset.
  if (!source && WRITER_DEFAULT_PARTS.has(path)) return [];
  if (!source) return [{ category: "added-part", xpath: location, detail: "output-only element" }];
  if (!output) return [{ category: "element", xpath: location, detail: "source-only element" }];
  const diffs: NonNullable<SemanticPartDiff["detail"]>[] = [];
  // A canonical writer renumbers document-unique revision markers when
  // restructuring source XML. Children pair positionally after semantic
  // matching, so positional ids carry no independent document semantics.
  const sourceRevisionId = REVISION_ELEMENTS.has(source.name)
    ? source.attributes["w:id"]
    : undefined;
  const outputRevisionId = REVISION_ELEMENTS.has(output.name)
    ? output.attributes["w:id"]
    : undefined;
  if (
    sourceRevisionId !== undefined &&
    outputRevisionId !== undefined &&
    sourceRevisionId !== outputRevisionId
  ) {
    source = { ...source, attributes: { ...source.attributes, "w:id": outputRevisionId } };
    output = { ...output, attributes: { ...output.attributes, "w:id": outputRevisionId } };
  }
  if (source.name !== output.name) {
    diffs.push({
      category: "element",
      xpath: location,
      detail: `${source.name} -> ${output.name}`,
    });
  }
  for (const [name, value] of Object.entries(source.attributes)) {
    const next = output.attributes[name];
    if (next === undefined) {
      if (name === "xml:space") continue;
      if (SOURCE_FALSE_DEFAULT_ATTRIBUTES.has(name) && (value === "0" || value === "false"))
        continue;
      // MCE: a consumer must ignore attributes from mc:Ignorable namespaces it
      // does not understand — our writer legitimately drops unknown foreign
      // attributes, so a source-only foreign attribute is not a model gap.
      if (!isIgnorableForeignAttr(name, ignorablePrefixes)) {
        diffs.push({
          category: "attribute",
          xpath: `${location}/@${name}`,
          detail: "source-only",
        });
      }
    } else if (!semanticAttributeValueEqual(value, next)) {
      diffs.push({
        category: "attribute",
        xpath: `${location}/@${name}`,
        detail: "value mismatch",
      });
    }
  }
  for (const name of Object.keys(output.attributes)) {
    if (!(name in source.attributes)) {
      // Writer-side lexical encoding: text elements carry xml:space so exact
      // whitespace survives XML whitespace normalization in consumers.
      if (name === "xml:space" && output.attributes[name] === "preserve") continue;
      // Writer-side XSD defaults: a minimal source omits the attribute but
      // the writer always emits it with its schema default value.
      if (output.attributes[name] === WRITER_EXPLICIT_DEFAULTS.get(name)?.[name]) continue;
      const localName = source.name.split(":").pop() ?? source.name;
      const elementDefaults =
        WRITER_EXPLICIT_DEFAULTS.get(source.name) ?? WRITER_EXPLICIT_DEFAULTS.get(localName);
      if (elementDefaults && elementDefaults[name] === output.attributes[name]) continue;
      diffs.push({
        category: "attribute",
        xpath: `${location}/@${name}`,
        detail: "output-only default",
      });
    }
  }
  const whitespaceTextEquivalent =
    source.name === output.name &&
    ["w:t", "w:delText", "w:instrText", "w:delInstrText", "m:t"].includes(source.name) &&
    source.text.trim() === "" &&
    output.text.trim() === "";
  if (!whitespaceTextEquivalent && source.text !== output.text) {
    diffs.push({ category: "text", xpath: location, detail: "text mismatch" });
  }
  if (
    sourceRevisionId !== undefined &&
    outputRevisionId !== undefined &&
    sourceRevisionId !== outputRevisionId
  ) {
  }
  const visibleChild = (child: CanonicalNode): boolean =>
    !isIgnorableForeignAttr(child.name, ignorablePrefixes);
  let sourceChildren = source.children.filter(visibleChild).map((child) => ({
    key: semanticChildFingerprint(child, ignorablePrefixes),
    child,
  }));
  let outputChildren = output.children.filter(visibleChild).map((child) => ({
    key: semanticChildFingerprint(child, ignorablePrefixes),
    child,
  }));
  // Glossary building blocks have stable user-visible identity (`w:name`);
  // pair by identity before comparing full subtrees so an unchanged entry
  // that moved cannot mask the actual changed entry behind a reorder report.
  const docPartIdentity = (child: CanonicalNode): string | undefined => {
    if (source.name !== "w:docParts") return undefined;
    const name = child.children
      .find((candidate) => candidate.name === "w:docPartPr")
      ?.children.find((candidate) => candidate.name === "w:name");
    return name?.attributes["w:val"];
  };
  const docPartMatches = new Map<number, number>();
  if (source.name === "w:docParts") {
    const outputIdentities = new Map<string, number[]>();
    outputChildren.forEach((entry) => {
      const identity = docPartIdentity(entry.child);
      if (identity !== undefined) {
        const indexes = outputIdentities.get(identity) ?? [];
        indexes.push(outputChildren.indexOf(entry));
        outputIdentities.set(identity, indexes);
      }
    });
    sourceChildren.forEach((sourceEntry, sourceIndex) => {
      const identity = docPartIdentity(sourceEntry.child);
      if (identity === undefined) return;
      const outputIndex = outputIdentities.get(identity)?.shift();
      if (outputIndex !== undefined) docPartMatches.set(sourceIndex, outputIndex);
    });
  }
  // Producers (pandoc) can emit the same w:style twice with the same
  // type + styleId; Word applies the last definition. Drop the superseded
  // copies on both sides so the unreferenced duplicates do not compare as
  // unmatched entries.
  if ((source.name.split(":").pop() ?? source.name) === "styles") {
    // XSD requires both attributes; a typeless style has no style identity for
    // Word to resolve and is dropped by valid writers.
    sourceChildren = sourceChildren.filter(
      ({ child }) =>
        child.name !== "w:style" ||
        (child.attributes["w:type"] !== undefined && child.attributes["w:styleId"] !== undefined),
    );
    const keepLastIndex = new Map<string, number>();
    sourceChildren.forEach(({ child }, index) => {
      if (child.name !== "w:style") return;
      keepLastIndex.set(
        `${child.attributes["w:type"] ?? ""}\0${child.attributes["w:styleId"] ?? ""}`,
        index,
      );
    });
    if (
      keepLastIndex.size < sourceChildren.filter(({ child }) => child.name === "w:style").length
    ) {
      sourceChildren = sourceChildren.filter(({ child }, index) => {
        if (child.name !== "w:style") return true;
        return (
          keepLastIndex.get(
            `${child.attributes["w:type"] ?? ""}\0${child.attributes["w:styleId"] ?? ""}`,
          ) === index
        );
      });
    }
  }
  if (path === "[Content_Types].xml") {
    // The writer adds standard companion parts a minimal source omits (see
    // WRITER_DEFAULT_PARTS); their Override declarations are required for OPC
    // validity and are not table drift when the source lacked the part.
    outputChildren = outputChildren.filter(({ child }) => {
      if (child.name !== "Override") return true;
      const partName = (child.attributes?.PartName ?? "").replace(/^\//, "");
      if (!isWriterDefaultPart(partName)) return true;
      return sourceChildren.some(({ child: sourceChild }) => {
        return (
          sourceChild.name === "Override" &&
          sourceChild.attributes?.PartName === child.attributes?.PartName
        );
      });
    });
    outputChildren = outputChildren.filter(({ child }) => {
      if (child.name !== "Override") return true;
      const partName = String(child.attributes?.PartName ?? "");
      const packagePath = partName.replace(/^\//, "");
      // Both packages omit the part: the source declaration was orphaned and
      // the round-tripped declaration carries no payload.
      if (
        packageParts !== undefined &&
        !packageParts.source.has(packagePath) &&
        !packageParts.output.has(packagePath)
      )
        return false;
      const sourceOverride = sourceChildren.some(({ child: sourceChild }) => {
        return (
          sourceChild.name === "Override" &&
          String(sourceChild.attributes?.PartName ?? "").toLowerCase() === partName.toLowerCase() &&
          sourceChild.attributes?.ContentType !== undefined
        );
      });
      // A writer may synthesize a standard part absent from a minimal source;
      // its required Override is a semantic superset, not drift.
      if (sourceOverride) return true;
      return (
        packageParts !== undefined &&
        !packageParts.source.has(packagePath) &&
        packageParts.output.has(packagePath)
      );
    });
    // Drop an Override for a part that neither package carries. The source
    // declaration was orphaned before parsing and has no semantic payload to
    // preserve; filtering it out must not count as content-type drift.
    sourceChildren = sourceChildren.filter(({ child }) => {
      if (child.name !== "Override" || packageParts === undefined) return true;
      const partName = String(child.attributes?.PartName ?? "").replace(/^\//, "");
      return packageParts.source.has(partName) || packageParts.output.has(partName);
    });
    // Writer-emitted opaque binaries (printerSettings) need a package-level
    // Default; a minimal source that shipped the binary without declaring it
    // stays readable, and the output's required Default is not drift.
    outputChildren = outputChildren.filter(({ child }) => {
      if (child.name !== "Default") return true;
      const extension = String(child.attributes?.Extension ?? "");
      const contentType = String(child.attributes?.ContentType ?? "");
      const required = WRITER_DEFAULT_EXTENSION_DEFAULTS.get(extension);
      if (required === undefined || required !== contentType) return true;
      return sourceChildren.some(
        ({ child: sourceChild }) =>
          sourceChild.name === "Default" &&
          String(sourceChild.attributes?.Extension ?? "") === extension,
      );
    });
  }
  const sourceCounts = new Map(sourceChildren.map(({ key }) => [key, 0]));
  for (const { key } of sourceChildren) sourceCounts.set(key, (sourceCounts.get(key) ?? 0) + 1);
  const outputCounts = new Map(outputChildren.map(({ key }) => [key, 0]));
  for (const { key } of outputChildren) outputCounts.set(key, (outputCounts.get(key) ?? 0) + 1);
  const sourceMatched = sourceChildren.map(() => false);
  const outputMatched = outputChildren.map(() => false);
  for (const [sourceIndex, outputIndex] of docPartMatches) {
    sourceMatched[sourceIndex] = true;
    outputMatched[outputIndex] = true;
  }
  for (const [sourceIndex, { key }] of sourceChildren.entries()) {
    const outputIndex = outputChildren.findIndex(
      ({ key: outputKey }, index) => !outputMatched[index] && outputKey === key,
    );
    if (outputIndex !== -1) {
      sourceMatched[sourceIndex] = true;
      outputMatched[outputIndex] = true;
    }
  }

  const alignedOutputChildren =
    docPartMatches.size === 0
      ? outputChildren
      : sourceChildren.map((sourceEntry, sourceIndex) => {
          const matchedOutputIndex = docPartMatches.get(sourceIndex);
          if (matchedOutputIndex !== undefined) return outputChildren[matchedOutputIndex];
          const keyOutputIndex = outputChildren.findIndex(
            ({ key }, index) => !outputMatched[index] && key === sourceEntry.key,
          );
          if (keyOutputIndex !== -1) return outputChildren[keyOutputIndex];
          const nameOutputIndex = outputChildren.findIndex(
            ({ child }, index) => !outputMatched[index] && child.name === sourceEntry.child.name,
          );
          return nameOutputIndex === -1 ? undefined : outputChildren[nameOutputIndex];
        });

  for (const [sourceIndex, { child }] of sourceChildren.entries()) {
    if (sourceMatched[sourceIndex]) continue;
    const outputIndex = outputChildren.findIndex(
      ({ child: outputChild }, index) => !outputMatched[index] && outputChild.name === child.name,
    );
    if (outputIndex === -1) continue;
    sourceMatched[sourceIndex] = true;
    outputMatched[outputIndex] = true;
    const childDiffs = compareNodes(
      `${location}/${child.name}[${sourceIndex + 1}]`,
      child,
      outputChildren[outputIndex]!.child,
      ignorablePrefixes,
    );
    if (childDiffs.length) return childDiffs;
  }

  const sourceOnlyKeys = sourceChildren
    .map(({ key }, index) => (sourceMatched[index] ? undefined : key))
    .filter((key): key is string => key !== undefined);
  const outputOnlyKeys = outputChildren
    .map(({ key }, index) => (outputMatched[index] ? undefined : key))
    .filter((key): key is string => key !== undefined);
  if (sourceOnlyKeys.length || outputOnlyKeys.length) {
    for (const key of sourceOnlyKeys) {
      const child = sourceChildren.find(({ key: childKey }) => childKey === key)!.child;
      diffs.push({
        category: "child",
        xpath: `${location}/${child.name}`,
        detail: `source-only ${childDiagnostic(child)}`,
      });
    }
    for (const key of outputOnlyKeys) {
      const child = outputChildren.find(({ key: childKey }) => childKey === key)!.child;
      diffs.push({
        category: "child",
        xpath: `${location}/${child.name}`,
        detail: `output-only ${childDiagnostic(child)}`,
      });
    }
  } else if (
    sourceChildren.some(({ child }, index) => {
      const outputChild = alignedOutputChildren[index]?.child;
      return (
        !outputChild ||
        semanticChildFingerprint(child, ignorablePrefixes) !==
          semanticChildFingerprint(outputChild, ignorablePrefixes)
      );
    })
  ) {
    const mismatchIndex = sourceChildren.findIndex(({ child }, index) => {
      const outputChild = alignedOutputChildren[index]?.child;
      return (
        !outputChild ||
        semanticChildFingerprint(child, ignorablePrefixes) !==
          semanticChildFingerprint(outputChild, ignorablePrefixes)
      );
    });
    const child = sourceChildren[mismatchIndex]!.child;
    diffs.push({
      category: "child-order",
      xpath: `${location}/${child.name}[${mismatchIndex + 1}]`,
      detail: `child fingerprint mismatch: ${childDiagnostic(child)} != ${childDiagnostic(alignedOutputChildren[mismatchIndex]?.child)}`,
    });
  }
  if (diffs.length) return diffs;
  for (let index = 0; index < sourceChildren.length; index++) {
    const childDiffs = compareNodes(
      `${location}/${sourceChildren[index]!.child.name}[${index + 1}]`,
      sourceChildren[index]!.child,
      alignedOutputChildren[index]?.child,
      ignorablePrefixes,
    );
    if (childDiffs.length) return childDiffs;
  }
  return [];
}

/**
 * ST_Percentage permits both scaled integers (`50000` = 50%) and percent
 * literals (`50%`); writers may choose either equivalent serialization.
 */
function semanticAttributeValueEqual(left: string, right: string): boolean {
  if (left === right) return true;
  if (left === "" || right === "") return false;
  const scaled = (value: string) =>
    value.endsWith("%") ? Number(value.slice(0, -1)) / 100 : Number(value);
  if (!Number.isFinite(scaled(left)) || !Number.isFinite(scaled(right))) return false;
  return scaled(left) === scaled(right);
}

/** Parts the writers always emit even when a minimal source omits them. */
const WRITER_DEFAULT_PARTS = new Set([
  "docProps/core.xml",
  "docProps/app.xml",
  "word/fontTable.xml",
  "word/theme/theme1.xml",
  "word/settings.xml",
  "word/styles.xml",
  "word/numbering.xml",
  "word/webSettings.xml",
  "xl/theme/theme1.xml",
  "xl/styles.xml",
  "xl/sharedStrings.xml",
  "ppt/theme/theme1.xml",
  "ppt/presProps.xml",
  "ppt/viewProps.xml",
  "ppt/tableStyles.xml",
  "xl/workbook.xml",
]);

/** Writer-required generated members whose source declarations may be absent. */
function isWriterDefaultPart(partName: string): boolean {
  return (
    WRITER_DEFAULT_PARTS.has(partName) ||
    /^xl\/worksheets\/sheet\d+\.xml$/.test(partName) ||
    /^xl\/chartsheets\/sheet\d+\.xml$/.test(partName)
  );
}

/**
 * XSD defaults the writers emit explicitly on elements where the source may
 * omit the attribute. Keyed by local element name.
 */
const WRITER_EXPLICIT_DEFAULTS = new Map<string, Record<string, string>>([
  ["w:qFormat", { "w:val": "1" }],
  ["functionGroups", { builtInGroupCount: "16" }],
  ["w:lvlJc", { "w:val": "start" }],
  ["w:pgMar", { "w:header": "720", "w:footer": "720", "w:gutter": "0" }],
  ["w:ins", { "w:date": "" }],
  ["w:del", { "w:date": "" }],
  ["Relationship", { TargetMode: "Internal" }],
]);

export function explainSemanticPartDiff(
  partPath: string,
  source: Uint8Array,
  output: Uint8Array | undefined,
  references?: {
    source?: Map<string, string> | undefined;
    output?: Map<string, string> | undefined;
  },
  packageParts?: { source: ReadonlySet<string>; output: ReadonlySet<string> },
): SemanticPartDiff[] {
  const kind = semanticPartKind(partPath);
  if (!output)
    return [{ path: partPath, kind, category: "missing-part", detail: "output missing" }];
  if (kind === "binary") {
    const sourceHash = createHash("sha256").update(source).digest("hex");
    const outputHash = createHash("sha256").update(output).digest("hex");
    return sourceHash === outputHash
      ? []
      : [{ path: partPath, kind, category: "binary", detail: "bytes differ" }];
  }
  const sourceXml = parseCanonicalXml(decodeXmlBytes(source));
  const outputXml = parseCanonicalXml(decodeXmlBytes(output));
  const ignorablePrefixes = new Set(
    (sourceXml?.attributes?.["mc:Ignorable"] ?? "")
      .split(/[\s,]+/)
      .map((prefix) => prefix.trim())
      .filter(Boolean),
  );
  const sourceNode = canonicalXmlNodes(sourceXml, partPath, references?.source);
  const outputNode = canonicalXmlNodes(outputXml, partPath, references?.output);
  const details = compareNodes(partPath, sourceNode, outputNode, ignorablePrefixes, packageParts);
  return details.map((detail) => ({ path: partPath, kind, ...detail }));
}

function relationshipPath(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  const directory = separator === -1 ? "" : partPath.slice(0, separator);
  const fileName = separator === -1 ? partPath : partPath.slice(separator + 1);
  return `${directory ? `${directory}/` : ""}_rels/${fileName}.rels`;
}

function resolveRelationshipTarget(ownerPath: string, target: string): string {
  if (/^[a-z]+:\/\//i.test(target)) return target;
  if (target.startsWith("/")) return target.slice(1);
  const separator = ownerPath.lastIndexOf("/");
  const directory = separator === -1 ? "." : ownerPath.slice(0, separator);
  return path.posix.normalize(path.posix.join(directory, target));
}

function relationshipReferences(
  archive: Record<string, Uint8Array>,
  partPath: string,
): Map<string, string> | undefined {
  const rels = archive[relationshipPath(partPath)];
  if (!rels) return undefined;
  const root = parseCanonicalXml(decodeXmlBytes(rels));
  const references = new Map<string, string>();
  const representativeIds = new Map<string, string>();
  for (const relationship of root?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const id = relationship.attributes?.["Id"];
    const type = relationship.attributes?.["Type"];
    const target = relationship.attributes?.["Target"];
    if (id === undefined || type === undefined || target === undefined) continue;
    const mode = relationship.attributes?.["TargetMode"] ?? "Internal";
    const normalizedTarget =
      mode === "External" ? target : resolveRelationshipTarget(partPath, target);
    const edge = `${canonicalAttributeValue("Type", String(type))}|${mode}|${normalizedTarget}`;
    const representativeId = representativeIds.get(edge);
    if (representativeId !== undefined) {
      references.set(String(id), references.get(representativeId)!);
      continue;
    }
    representativeIds.set(edge, String(id));
    references.set(String(id), edge);
  }
  return references.size > 0 ? references : undefined;
}

function sortedElementXml(
  element: Element | undefined,
  childKey: (child: Element) => string,
): string {
  if (!element || element.type !== "element") return "empty";
  const children = (element.elements ?? [])
    .filter((child) => child.type === "element")
    .sort((left, right) => childKey(left).localeCompare(childKey(right)))
    .map((child) => `<${child.name ?? ""}|${canonicalAttributes(child)}|${canonicalText(child)}>`);
  return `${element.name ?? ""}|${canonicalAttributes(element)}|${canonicalText(element)}|${children.join("")}`;
}

export function canonicalRelationshipsDigest(element: Element | undefined): string {
  return createHash("sha256")
    .update(
      sortedElementXml(element, (child) =>
        ["Id", "Type", "Target", "TargetMode"]
          .map((name) => String(child.attributes?.[name] ?? ""))
          .join("\0"),
      ),
    )
    .digest("hex");
}

export function canonicalContentTypesDigest(element: Element | undefined): string {
  return createHash("sha256")
    .update(
      sortedElementXml(element, (child) =>
        child.name === "Default"
          ? `Default\0${String(child.attributes?.Extension ?? "")}`
          : `Override\0${String(child.attributes?.PartName ?? "")}`,
      ),
    )
    .digest("hex");
}

export function parseCanonicalXml(xml: string): Element | undefined {
  return parse(xml, {
    ignoreDeclaration: true,
    normalizeNamespaces: OOXML_CANONICAL_PREFIXES,
  }).elements?.find(
    (candidate): candidate is Element =>
      candidate.type === "element" && !(candidate.name ?? "").startsWith("!"),
  );
}

export function semanticPartKind(path: string): SemanticPartKind {
  if (path === "[Content_Types].xml") return "content-types";
  if (path.endsWith(".rels")) return "relationship";
  // OPC 3.0 package-level core properties are core-properties XML despite
  // the .psmdcp extension — compare semantically, not byte-wise.
  if (path.endsWith(".psmdcp")) return "xml";
  if (path.toLowerCase().endsWith(".xml")) return "xml";
  return "binary";
}

export function semanticPartDiff(
  path: string,
  source: Uint8Array,
  output: Uint8Array | undefined,
): SemanticPartDiff | undefined {
  const kind = semanticPartKind(path);
  if (kind === "binary") {
    const sourceHash = createHash("sha256").update(source).digest("hex");
    const outputHash = output
      ? createHash("sha256").update(output).digest("hex")
      : createHash("sha256").digest("hex");
    return sourceHash === outputHash ? undefined : { path, kind };
  }
  return explainSemanticPartDiff(path, source, output)[0];
}

export function archiveSemanticDiffDetails(
  source: Uint8Array,
  output: Uint8Array,
): SemanticPartDiff[] {
  // OPC part names use forward slashes; normalize backslash-separated entry
  // names the same way the parser does so non-conformant producers compare
  // against their canonical round-trip output.
  const normalizeArchive = (data: Uint8Array): Record<string, Uint8Array> => {
    const raw = unzipOpcEntries(data);
    const out: Record<string, Uint8Array> = {};
    for (const [name, bytes] of Object.entries(raw)) {
      const normalized = name.replace(/\\/g, "/");
      // OPC part names are case-insensitive; producers may spell the package
      // metadata files with different casing than the writer's canonical form.
      const canonicalName = ALWAYS_REBUILT_PARTS_LOWER.get(normalized.toLowerCase()) ?? normalized;
      out[canonicalName] = bytes;
    }
    return out;
  };
  const sourceArchive = normalizeArchive(source);
  const outputArchive = normalizeArchive(output);
  const paths = new Set([...Object.keys(sourceArchive), ...Object.keys(outputArchive)]);
  const packageParts = {
    source: new Set(Object.keys(sourceArchive)),
    output: new Set(Object.keys(outputArchive)),
  };
  const diffs: SemanticPartDiff[] = [];
  for (const path of [...paths].sort()) {
    if (path.endsWith("/")) continue;
    if (
      path.endsWith(".rels") &&
      isEmptyRelationshipsPart(sourceArchive[path]) &&
      isEmptyRelationshipsPart(outputArchive[path])
    )
      continue;
    const references =
      /\.xml$/i.test(path) && !path.endsWith(".rels")
        ? {
            source: relationshipReferences(sourceArchive, path),
            output: relationshipReferences(outputArchive, path),
          }
        : undefined;
    diffs.push(
      ...explainSemanticPartDiff(
        path,
        sourceArchive[path] ?? new Uint8Array(),
        outputArchive[path],
        references,
        packageParts,
      ),
    );
  }
  return diffs;
}

export function assertEncryptedContainerRoundTrip(source: Uint8Array, output: Uint8Array): void {
  const encrypted = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (source.length < encrypted.length) return;
  const isEncrypted = encrypted.every((byte, index) => source[index] === byte);
  if (!isEncrypted) return;
  if (source.length !== output.length || !source.every((byte, index) => output[index] === byte))
    throw new Error("encrypted DOCX container was not re-emitted verbatim");
}

export function archiveSemanticDiffs(source: Uint8Array, output: Uint8Array): SemanticPartDiff[] {
  const sourceArchive = unzipOpcEntries(source);
  const outputArchive = unzipOpcEntries(output);
  const paths = new Set([...Object.keys(sourceArchive), ...Object.keys(outputArchive)]);
  const diffs: SemanticPartDiff[] = [];
  for (const path of [...paths].sort()) {
    if (path.endsWith("/")) continue;
    if (
      path.endsWith(".rels") &&
      isEmptyRelationshipsPart(sourceArchive[path]) &&
      isEmptyRelationshipsPart(outputArchive[path])
    )
      continue;
    const diff = semanticPartDiff(
      path,
      sourceArchive[path] ?? new Uint8Array(),
      outputArchive[path],
    );
    if (diff) diffs.push(diff);
  }
  return diffs;
}

/** Transitional baseline comparison while strict semantic repair is in progress. */
export function archiveTagDiffs(source: Uint8Array, output: Uint8Array): string[] {
  const sourceArchive = unzipOpcEntries(source);
  const outputArchive = unzipOpcEntries(output);
  const decoder = new TextDecoder();
  const tagCounts = (xml: string): Map<string, number> => {
    const counts = new Map<string, number>();
    for (const match of xml.matchAll(/<(?:[\w-]+:)?([\w-]+)[ >/]/g)) {
      counts.set(match[1]!, (counts.get(match[1]!) ?? 0) + 1);
    }
    return counts;
  };
  const differs = (sourceXml: string, outputXml: string): boolean => {
    const sourceCounts = tagCounts(sourceXml);
    const outputCounts = tagCounts(outputXml);
    for (const [name, count] of sourceCounts) if (outputCounts.get(name) !== count) return true;
    for (const [name, count] of outputCounts) if (sourceCounts.get(name) !== count) return true;
    return false;
  };
  const diffs: string[] = [];
  for (const path of Object.keys(sourceArchive)) {
    if (!path.endsWith(".xml") && !path.endsWith(".rels")) continue;
    if (path.includes("theme")) continue;
    if (
      path.endsWith(".rels") &&
      isEmptyRelationshipsPart(sourceArchive[path]) &&
      isEmptyRelationshipsPart(outputArchive[path])
    )
      continue;
    const sourceXml = decoder.decode(sourceArchive[path]!);
    const outputXml = decoder.decode(outputArchive[path] ?? new Uint8Array());
    if (sourceXml !== outputXml && differs(sourceXml, outputXml)) diffs.push(path);
  }
  return diffs;
}

export type CorpusFailureKind =
  | "valid"
  | "encrypted"
  | "invalid-zip"
  | "invalid-cfb"
  | "truncated-xml"
  | "invalid-package";

export function classifyPackageFailure(error: unknown): Exclude<CorpusFailureKind, "valid"> {
  const message = String(error).toLowerCase();
  if (message.includes("crc")) return "invalid-zip";
  if (message.includes("invalid zip data") || message.includes("end of central directory"))
    return "invalid-zip";
  if (message.includes("literal/lengths") || message.includes("invalid stored block"))
    return "invalid-zip";
  if (message.includes("cfb") || message.includes("compound file")) return "invalid-cfb";
  if (message.includes("unexpected end") || message.includes("unterminated"))
    return "truncated-xml";
  if (message.includes("not found")) return "invalid-package";
  return "invalid-package";
}

export function primaryPackagePart(archive: Pick<ParsedArchive, "get">, fallback: string): string {
  const relationships = archive.get("_rels/.rels");
  for (const relationship of relationships?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const type = String(relationship.attributes?.Type ?? "");
    if (!type.endsWith("/officeDocument")) continue;
    if (relationship.attributes?.TargetMode === "External") continue;
    const target = String(relationship.attributes?.Target ?? "");
    if (!target) continue;
    const resolved = target.startsWith("/") ? target.slice(1) : target.replaceAll("\\", "/");
    if (archive.get(resolved)) return resolved;
  }
  if (archive.get(fallback)) return fallback;
  throw new Error(`${fallback} package primary part not found`);
}

const PRIMARY_PART_ROOTS = {
  docx: "document",
  pptx: "presentation",
  xlsx: "workbook",
} as const;

/** XSD singletons commonly duplicated by damaged or lenient producers. */
const SINGLETON_CHILDREN = new Map<string, ReadonlySet<string>>([
  [
    "w:settings",
    new Set([
      "w:writeProtection",
      "w:view",
      "w:zoom",
      "w:defaultTabStop",
      "w:characterSpacingControl",
      "w:compat",
      "w:docVars",
      "w:rsids",
      "w:themeFontLang",
      "w:clrSchemeMapping",
      "w:shapeDefaults",
      "w:decimalSymbol",
      "w:listSeparator",
    ]),
  ],
  [
    "w:sectPr",
    new Set([
      "w:pgSz",
      "w:pgMar",
      "w:paperSrc",
      "w:pgBorders",
      "w:lnNumType",
      "w:pgNumType",
      "w:cols",
      "w:formProt",
      "w:vAlign",
      "w:noEndnote",
      "w:titlePg",
      "w:textDirection",
      "w:bidi",
      "w:rtlGutter",
      "w:docGrid",
      "w:printerSettings",
      "w:sectPrChange",
    ]),
  ],
]);

function rejectDuplicateSingletonChildren(partPath: string, element: Element | undefined): void {
  const singletonNames = SINGLETON_CHILDREN.get(element?.name ?? "");
  if (!singletonNames) return;
  const seen = new Set<string>();
  for (const child of element?.elements ?? []) {
    if (child.type !== "element" || !singletonNames.has(child.name ?? "")) continue;
    if (seen.has(child.name)) {
      throw new Error(`${partPath} duplicates ${child.name} in ${element!.name}`);
    }
    seen.add(child.name);
  }
}

function rejectKnownInvalidStructures(partPath: string, root: Element): void {
  const stack = [root];
  while (stack.length > 0) {
    const element = stack.pop()!;
    if ((element.name?.split(":").pop() ?? element.name) === "worksheet") {
      const prefix = element.name?.includes(":") ? element.name.split(":")[0] : undefined;
      const namespace = prefix
        ? element.attributes?.[`xmlns:${prefix}`]
        : element.attributes?.xmlns;
      if (namespace === "http://schemas.microsoft.com/office/excel/2005/8/worksheet") continue;
      const order = [
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "mergeCells",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "drawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
      ];
      const orderIndex = new Map(order.map((name, index) => [name, index]));
      let previousName: string | undefined;
      let previousIndex = -1;
      for (const child of element.elements ?? []) {
        if (child.type !== "element") continue;
        const childName = child.name?.split(":").pop() ?? child.name;
        const childIndex = orderIndex.get(childName);
        if (childIndex === undefined) continue;
        if (childIndex < previousIndex) {
          throw new Error(
            `${partPath} has worksheet children out of XSD order: ${childName} after ${previousName}`,
          );
        }
        previousName = childName;
        previousIndex = childIndex;
      }
    }
    if (
      element.name === "w:pPr" &&
      (element.elements ?? []).some((child) => child.name === "w:numId")
    )
      throw new Error(`${partPath} places w:numId directly in w:pPr`);
    if (element.name === "w:body") {
      const children = (element.elements ?? []).filter((child) => child.type === "element");
      const last = children.at(-1);
      if (children.some((child) => child.name === "w:sectPr") && last?.name !== "w:sectPr")
        throw new Error(`${partPath} has a body-level w:sectPr before trailing content`);
    }
    if (element.name === "w:r" && (element.elements ?? []).some((child) => child.name === "w:r"))
      throw new Error(`${partPath} nests w:r directly inside w:r`);
    stack.push(
      ...(element.elements ?? []).filter((child): child is Element => child.type === "element"),
    );
  }
}

export function validatePrimaryPackagePart(
  archive: Pick<ParsedArchive, "get" | "getRaw">,
  format: keyof typeof PRIMARY_PART_ROOTS,
  fallback: string,
): string {
  for (const partPath of archive.keys()) {
    if (!/\.xml$|\.rels$/i.test(partPath)) continue;
    const raw = archive.getRaw(partPath);
    if (raw !== undefined) {
      const prologue = new TextDecoder().decode(raw.slice(0, 1024));
      if (/<!DOCTYPE/i.test(prologue))
        throw new Error(`${partPath} must not contain a DOCTYPE declaration`);
    }
    const root = archive.get(partPath);
    if (!root) throw new Error(`${partPath} package part has no XML root`);
    rejectKnownInvalidStructures(partPath, root);
    const stack = [root];
    while (stack.length > 0) {
      const element = stack.pop()!;
      rejectDuplicateSingletonChildren(partPath, element);
      if (
        (element.name === "w:p" || element.name === "w:r") &&
        ((element.elements ?? []).some(
          (child) =>
            (child.type === "text" || child.name === undefined) &&
            String(child.text ?? "").trim() !== "",
        ) ||
          String(element.text ?? "").trim() !== "")
      ) {
        throw new Error(`${partPath} has raw text directly inside ${element.name}`);
      }
      if (
        (element.name === "c:val" || element.name.endsWith(":val")) &&
        (element.elements?.length ?? 0) === 0
      ) {
        throw new Error(`${partPath} has an empty chart value slot`);
      }
      stack.push(
        ...(element.elements ?? []).filter(
          (child): child is Element => child.type === "element" && child.name !== undefined,
        ),
      );
    }
    if (partPath.toLowerCase().endsWith(".rels")) {
      if (root.name !== "Relationships")
        throw new Error(`${partPath} relationships part has an invalid root`);
      for (const relationship of root.elements ?? []) {
        if (relationship.name !== "Relationship") continue;
        for (const attributeName of ["Id", "Type", "Target"] as const) {
          if (!relationship.attributes?.[attributeName])
            throw new Error(`${partPath} relationship is missing ${attributeName}`);
        }
      }
    }
  }
  const partPath = primaryPackagePart(archive, fallback);
  const root = archive.get(partPath);
  const rootName = root?.name?.split(":").pop();
  if (rootName !== PRIMARY_PART_ROOTS[format])
    throw new Error(`${partPath} primary part is not a ${format} document`);
  return partPath;
}
