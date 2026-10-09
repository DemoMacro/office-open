import { createHash } from "node:crypto";
import * as path from "node:path";

import { unzipSync } from "fflate";

import {
  OOXML_CANONICAL_PREFIXES,
  convertToTwip,
  type ParsedArchive,
} from "../../packages/core/dist/index.mjs";
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

/** Extended/core-properties boolean elements whose text is xsd:boolean. */
const BOOLEAN_TEXT_ELEMENTS = new Set([
  "ScaleCrop",
  "LinksUpToDate",
  "SharedDoc",
  "HyperlinksChanged",
  "o:LockedField",
]);

/** OOXML toggle elements whose omitted w:val means semantic true. */
const ON_OFF_ELEMENTS = new Set([
  "w:b",
  "w:bCs",
  "w:i",
  "w:iCs",
  "w:noProof",
  "w:vanish",
  "w:webHidden",
  "w:bidi",
  "w:keepNext",
  "w:keepLines",
  "w:pageBreakBefore",
  "w:widowControl",
  "w:suppressLineNumbers",
  "w:suppressAutoHyphens",
  "w:kinsoku",
  "w:wordWrap",
  "w:overflowPunct",
  "w:topLinePunct",
  "w:autoSpaceDE",
  "w:autoSpaceDN",
  "w:snapToGrid",
  "w:cantSplit",
  "w:tblHeader",
  "w:hidden",
  "m:grow",
  "m:subHide",
  "m:supHide",
  "m:aln",
  "m:plcHide",
]);

/** XML attributes whose XSD default is emitted explicitly by the writers. */
const DEFAULT_ATTRIBUTES = new Map<string, Record<string, string>>([
  [
    "w:pgMar",
    {
      "w:header": "851",
      "w:footer": "992",
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
  if (value === "on" || value === "true") return "1";
  if (value === "off" || value === "false") return "0";
  if (value === "t") return "1";
  if (value === "f") return "0";
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) && Number.isFinite(Number(value))) {
    return String(Number(value));
  }
  if (
    (elementName === "w:pgSz" && (name === "w:w" || name === "w:h")) ||
    (elementName === "w:pgMar" && name.startsWith("w:"))
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

const PACKAGE_XML_PART = /\.(?:xml|rels|vml)$/i;

/**
 * XML-shaped parts no relationship reaches, walking the closure from
 * `_rels/.rels`. These are orphaned package members: the document model has no
 * anchor for them, so they travel verbatim instead of masking a model gap.
 */
export function orphanedPackageMembers(source: Uint8Array): Set<string> {
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(source);
  } catch {
    return new Set();
  }
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
    if (!PACKAGE_XML_PART.test(partPath)) continue;
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
    // O12 templates write `@range`; the XSD attribute is `@ref`.
    if (name === "dimension" && attributeName === "range") return "ref";
    if (name === "sst" && attributeName === "totalCount") return "count";
    if ((name === "w16se:sym" || name === "w16se:symEx") && attributeName === "w16se:char")
      return "w:char";
    if ((name === "w16se:sym" || name === "w16se:symEx") && attributeName === "w16se:font")
      return "w:font";
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
      // O12 sheetView `@active` is an unused selection hint absent from the
      // transitional XSD; writers never re-emit it.
      .filter(
        ([attributeName]) =>
          !((name === "sheetView" || name.endsWith(":sheetView")) && attributeName === "active"),
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
    if (!(valueAttribute in attributes)) attributes[valueAttribute] = "1";
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
  const rawText = (element.elements ?? [])
    .filter((child) => child.type === "text" || child.type === "cdata")
    .map((child) => String(child.text ?? child.cdata ?? ""))
    .join("");
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
  const childrenOut =
    name === "w:pPr"
      ? mappedChildren.filter(
          (child) => !(child.name === "w:pPr" && child.children.length === 0 && child.text === ""),
        )
      : mappedChildren;
  // Word also tolerates duplicate w:rPr children inside w:r by applying the
  // last one — normalize to the same precedence instead of comparing the
  // structurally invalid source markup against our valid output.
  if (name === "w:r") {
    const rPrIndices = mappedChildren
      .map((child, index) => (child.name === "w:rPr" ? index : -1))
      .filter((index) => index >= 0);
    if (rPrIndices.length > 1) {
      const keep = rPrIndices[rPrIndices.length - 1]!;
      const dropped = new Set(rPrIndices.filter((index) => index !== keep));
      const deduped = mappedChildren.filter((_, index) => !dropped.has(index));
      childrenOut.length = 0;
      childrenOut.push(...deduped);
    }
  }
  return {
    name: canonicalElementName(name, path),
    attributes: orderedAttributes,
    text:
      bareRunText !== ""
        ? ""
        : element.attributes?.["xml:space"] === "preserve"
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
    localName === "styles" ||
    localName === "docparts" ||
    localName === "docpartpr" ||
    localName === "settings" ||
    localName === "ppr" ||
    localName === "rpr" ||
    localName === "gslst" ||
    localName === "schemeclr" ||
    localName === "ser" ||
    localName === "chart" ||
    localName === "catax" ||
    localName === "valax" ||
    localName === "dataax" ||
    localName === "serax" ||
    localName === "dateax" ||
    localName === "worksheet" ||
    localName === "dialogsheet" ||
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
          remaining[mergedIndex] =
            moved.length > 0 ? { ...compat, children: [...compat.children, ...moved] } : compat;
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
  hash.update(`${node.name}\0${JSON.stringify(node.attributes)}\0${JSON.stringify(node.text)}`);
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
        !isIgnorableForeignAttr(name, ignorablePrefixes) && !(defaults && defaults[name] === value),
    ),
  );
}

export function semanticChildFingerprint(
  node: CanonicalNode,
  ignorablePrefixes: ReadonlySet<string> | undefined,
): string {
  const hash = createHash("sha256");
  const attributes = writerDefaultOmittedAttributes(node, ignorablePrefixes);
  hash.update(`${node.name}\0${JSON.stringify(attributes)}\0${JSON.stringify(node.text)}`);
  for (const child of node.children)
    hash.update(semanticChildFingerprint(child, ignorablePrefixes));
  return hash.digest("hex");
}

/**
 * Stable ordering key for unordered containers: name + attributes + text plus
 * one level of child identity. Deep subtree differences (e.g. a missing
 * c:extLst inside a c:dPt) must not reorder the container, or the comparator
 * would pair unrelated siblings and report phantom diffs.
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
  hash.update(`${node.name}\0${JSON.stringify(attributes)}\0${JSON.stringify(node.text)}`);
  for (const child of node.children) {
    const childAttributes = Object.fromEntries(
      Object.entries(child.attributes).filter(
        ([name]) => !isIgnorableForeignAttr(name, ignorablePrefixes),
      ),
    );
    hash.update(`${child.name}\0${JSON.stringify(childAttributes)}\0${JSON.stringify(child.text)}`);
  }
  return hash.digest("hex");
}

function canonicalAttributes(element: Element): string {
  return Object.entries(element.attributes ?? {})
    .filter(([name]) => name !== "xmlns" && !name.startsWith("xmlns:"))
    .sort(([left], [right]) => left.localeCompare(right))
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
  return element.attributes?.["xml:space"] === "preserve" ? text : text.trim();
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
): SemanticPartDiff["detail"][] {
  if (!source && !output) return [];
  const location = path || "/";
  // Minimal source packages may omit standard companion parts that every
  // writer emits; the generated package is a semantic superset.
  if (!source && WRITER_DEFAULT_PARTS.has(path)) return [];
  if (!source) return [{ category: "added-part", xpath: location, detail: "output-only element" }];
  if (!output) return [{ category: "element", xpath: location, detail: "source-only element" }];
  const diffs: NonNullable<SemanticPartDiff["detail"]>[] = [];
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
    } else if (next !== value) {
      diffs.push({
        category: "attribute",
        xpath: `${location}/@${name}`,
        detail: "value mismatch",
      });
    }
  }
  for (const name of Object.keys(output.attributes)) {
    if (!(name in source.attributes)) {
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
  if (source.text !== output.text) {
    diffs.push({ category: "text", xpath: location, detail: "text mismatch" });
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
  // Producers (pandoc) can emit the same w:style twice with the same
  // type + styleId; Word applies the last definition. Drop the superseded
  // copies on both sides so the unreferenced duplicates do not compare as
  // unmatched entries.
  if ((source.name.split(":").pop() ?? source.name) === "styles") {
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
      if (!WRITER_DEFAULT_PARTS.has(partName)) return true;
      return sourceChildren.some(({ child: sourceChild }) => {
        return (
          sourceChild.name === "Override" &&
          sourceChild.attributes?.PartName === child.attributes?.PartName
        );
      });
    });
  }
  const sourceCounts = new Map(sourceChildren.map(({ key }) => [key, 0]));
  for (const { key } of sourceChildren) sourceCounts.set(key, (sourceCounts.get(key) ?? 0) + 1);
  const outputCounts = new Map(outputChildren.map(({ key }) => [key, 0]));
  for (const { key } of outputChildren) outputCounts.set(key, (outputCounts.get(key) ?? 0) + 1);
  const sourceMatched = sourceChildren.map(() => false);
  const outputMatched = outputChildren.map(() => false);
  for (const [sourceIndex, { key }] of sourceChildren.entries()) {
    const outputIndex = outputChildren.findIndex(
      ({ key: outputKey }, index) => !outputMatched[index] && outputKey === key,
    );
    if (outputIndex !== -1) {
      sourceMatched[sourceIndex] = true;
      outputMatched[outputIndex] = true;
    }
  }

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
    sourceChildren.some(
      ({ child }, index) =>
        semanticChildFingerprint(child, ignorablePrefixes) !==
        semanticChildFingerprint(outputChildren[index]!.child, ignorablePrefixes),
    )
  ) {
    diffs.push({ category: "child-order", xpath: location, detail: "children reordered" });
  }
  if (diffs.length) return diffs;
  for (let index = 0; index < sourceChildren.length; index++) {
    const childDiffs = compareNodes(
      `${location}/${sourceChildren[index]!.child.name}[${index + 1}]`,
      sourceChildren[index]!.child,
      outputChildren[index]?.child,
      ignorablePrefixes,
    );
    if (childDiffs.length) return childDiffs;
  }
  return [];
}

/** Parts the writers always emit even when a minimal source omits them. */
const WRITER_DEFAULT_PARTS = new Set([
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
]);

/**
 * XSD defaults the writers emit explicitly on elements where the source may
 * omit the attribute. Keyed by local element name.
 */
const WRITER_EXPLICIT_DEFAULTS = new Map<string, Record<string, string>>([
  ["functionGroups", { builtInGroupCount: "16" }],
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
  const details = compareNodes(partPath, sourceNode, outputNode, ignorablePrefixes);
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
  for (const relationship of root?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const id = relationship.attributes?.["Id"];
    const type = relationship.attributes?.["Type"];
    const target = relationship.attributes?.["Target"];
    if (id === undefined || type === undefined || target === undefined) continue;
    const mode = relationship.attributes?.["TargetMode"] ?? "Internal";
    const normalizedTarget =
      mode === "External" ? target : resolveRelationshipTarget(partPath, target);
    references.set(
      String(id),
      `${canonicalAttributeValue("Type", String(type))}|${mode}|${normalizedTarget}`,
    );
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
  }).elements?.find((candidate): candidate is Element => candidate.type === "element");
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
    const raw = unzipSync(data);
    const out: Record<string, Uint8Array> = {};
    for (const [name, bytes] of Object.entries(raw)) out[name.replace(/\\/g, "/")] = bytes;
    return out;
  };
  const sourceArchive = normalizeArchive(source);
  const outputArchive = normalizeArchive(output);
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
  const sourceArchive = unzipSync(source);
  const outputArchive = unzipSync(output);
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
  const sourceArchive = unzipSync(source);
  const outputArchive = unzipSync(output);
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
