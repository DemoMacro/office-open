/**
 * Declarative field-consistency spec for descriptors.
 *
 * The `CustomDescriptor<T>` closure carries no field metadata, so the three
 * places that must agree — the Options interface, the stringify body, and the
 * parse body — drift silently (reflection dropping 9/14 props, styles
 * round-trip loss, core-properties 8-interface/10-write/8-parse inflation).
 *
 * Each entry declares the three field sets in a single semantic dimension
 * (Options field names, not XML tags), so {@link diffTagSets} can surface the
 * asymmetries directly:
 *   - F1 write-loss  — declared on the interface but never written
 *   - F2 write-only  — written but absent from the interface (inflation)
 *   - F3 parse-loss  — written but never restored on parse (round-trip loss)
 *   - F5 parse-only  — restored but never written
 *
 * @module
 */

export interface DescriptorFieldSpec {
  /** Stable identifier, matched to the descriptor under test. */
  id: string;
  /** Options interface name (documentation / assertion messages). */
  optionsInterface: string;
  /** Fields declared on the Options interface — the contract. */
  interfaceFields: readonly string[];
  /** Fields the stringify path actually emits; may carry inflation not on the interface. */
  writeFields: readonly string[];
  /** Fields the parse path actually restores. */
  parseFields: readonly string[];
  /** Expected child order when the XSD mandates sequence; omitted when order-independent. */
  order?: readonly string[];
  /** Sample exercising the interface fields, for F4 round-trip deep-equality. */
  sampleOptions: unknown;
  /**
   * Interface fields excluded from the contract comparison — input-side sugar
   * (thematicBreak→border, rightTabStop→tabStops) and control flags
   * (includeIfEmpty) that map field→XML but XML→a different field, breaking the
   * 1:1 round-trip assumption. The interface-drift test asserts
   * `interfaceFields === (live interface fields − excludeFields)`.
   */
  excludeFields?: readonly string[];
  notes?: string;
}

/**
 * Curated high-value targets — each maps to a known historical drift.
 * `settings` is intentionally absent: its round-trip is rawXml-based (byte
 * passthrough), so field-level diff would mislabel every field as F2.
 */
export const FIELD_SPECS: readonly DescriptorFieldSpec[] = [
  {
    id: "core-properties",
    optionsInterface: "CorePropertiesOptions",
    // 15 content fields — created/modified + the five CT_CoreProperties
    // string fields (category/contentStatus/identifier/language/version).
    // defaultNamespace is a round-trip form flag (prefix vs default-ns
    // binding), not a serialized content field.
    excludeFields: ["defaultNamespace"],
    interfaceFields: [
      "title",
      "subject",
      "creator",
      "keywords",
      "description",
      "lastModifiedBy",
      "revision",
      "lastPrinted",
      "created",
      "modified",
      "category",
      "contentStatus",
      "identifier",
      "language",
      "version",
    ],
    // stringify emits all 15 fields; created/modified default to now when absent.
    writeFields: [
      "title",
      "subject",
      "creator",
      "keywords",
      "description",
      "lastModifiedBy",
      "revision",
      "lastPrinted",
      "created",
      "modified",
      "category",
      "contentStatus",
      "identifier",
      "language",
      "version",
    ],
    // parse reads back all 15 interface fields.
    parseFields: [
      "title",
      "subject",
      "creator",
      "keywords",
      "description",
      "lastModifiedBy",
      "revision",
      "lastPrinted",
      "created",
      "modified",
      "category",
      "contentStatus",
      "identifier",
      "language",
      "version",
    ],
    sampleOptions: {
      title: "T",
      subject: "S",
      creator: "C",
      keywords: "k1,k2",
      description: "D",
      lastModifiedBy: "LMB",
      revision: 3,
      lastPrinted: "2024-01-02T03:04:05Z",
      created: "2018-05-11T07:02:00Z",
      modified: "2026-04-10T02:02:36Z",
      category: "cat",
      contentStatus: "draft",
      identifier: "id-1",
      language: "en-US",
      version: "1.2",
    },
    notes:
      "created/modified round-tripped verbatim from dcterms:created/modified; the five string fields map CT_CoreProperties elements 1:1.",
  },
  {
    id: "paragraph-properties",
    optionsInterface: "ParagraphPropertiesOptions",
    interfaceFields: [
      "heading",
      "style",
      "bullet",
      "run",
      "border",
      "shading",
      "numbering",
      "alignment",
      "bidirectional",
      "pageBreakBefore",
      "tabStops",
      "widowControl",
      "contextualSpacing",
      "indent",
      "spacing",
      "keepNext",
      "keepLines",
      "frame",
      "suppressLineNumbers",
      "wordWrap",
      "overflowPunctuation",
      "autoSpaceEastAsianText",
      "suppressOverlap",
      "suppressAutoHyphens",
      "adjustRightInd",
      "snapToGrid",
      "mirrorIndents",
      "kinsoku",
      "topLinePunct",
      "autoSpaceDE",
      "textAlignment",
      "textboxTightWrap",
      "textDirection",
      "outlineLevel",
      "divId",
      "cnfStyle",
      "revision",
    ],
    // stringify writes every interface field (pPrChange is revision's XML).
    writeFields: [
      "heading",
      "style",
      "bullet",
      "run",
      "border",
      "shading",
      "numbering",
      "alignment",
      "bidirectional",
      "pageBreakBefore",
      "tabStops",
      "widowControl",
      "contextualSpacing",
      "indent",
      "spacing",
      "keepNext",
      "keepLines",
      "frame",
      "suppressLineNumbers",
      "wordWrap",
      "overflowPunctuation",
      "autoSpaceEastAsianText",
      "suppressOverlap",
      "suppressAutoHyphens",
      "adjustRightInd",
      "snapToGrid",
      "mirrorIndents",
      "kinsoku",
      "topLinePunct",
      "autoSpaceDE",
      "textAlignment",
      "textboxTightWrap",
      "textDirection",
      "outlineLevel",
      "divId",
      "cnfStyle",
      "revision",
    ],
    // parse covers every write field (round-trip complete).
    parseFields: [
      "heading",
      "style",
      "bullet",
      "run",
      "border",
      "shading",
      "numbering",
      "alignment",
      "bidirectional",
      "pageBreakBefore",
      "tabStops",
      "widowControl",
      "contextualSpacing",
      "indent",
      "spacing",
      "keepNext",
      "keepLines",
      "frame",
      "suppressLineNumbers",
      "wordWrap",
      "overflowPunctuation",
      "autoSpaceEastAsianText",
      "suppressOverlap",
      "suppressAutoHyphens",
      "adjustRightInd",
      "snapToGrid",
      "mirrorIndents",
      "kinsoku",
      "topLinePunct",
      "autoSpaceDE",
      "textAlignment",
      "textboxTightWrap",
      "textDirection",
      "outlineLevel",
      "divId",
      "cnfStyle",
      "revision",
    ],
    order: [
      "pStyle",
      "keepNext",
      "keepLines",
      "pageBreakBefore",
      "framePr",
      "widowControl",
      "numPr",
      "suppressLineNumbers",
      "pBdr",
      "shd",
      "tabs",
      "suppressAutoHyphens",
      "kinsoku",
      "wordWrap",
      "overflowPunct",
      "topLinePunct",
      "autoSpaceDE",
      "autoSpaceDN",
      "bidi",
      "adjustRightInd",
      "snapToGrid",
      "spacing",
      "ind",
      "contextualSpacing",
      "mirrorIndents",
      "suppressOverlap",
      "jc",
      "textDirection",
      "textAlignment",
      "textboxTightWrap",
      "outlineLvl",
      "divId",
      "cnfStyle",
      "rPr",
      "pPrChange",
    ],
    sampleOptions: {
      heading: "Heading1",
      alignment: "center",
      spacing: { before: 240, after: 120, line: 278, lineRule: "auto" },
      indent: { left: 720, firstLine: 360 },
      keepNext: true,
      keepLines: true,
      widowControl: true,
      contextualSpacing: true,
      border: { bottom: { style: "single", color: "FF0000", size: 6, space: 1 } },
      shading: { fill: "FFFF00", type: "clear" },
      tabStops: [{ position: 720, type: "left" }],
      numbering: { reference: "list_1", level: 0 },
      outlineLevel: 1,
      textAlignment: "center",
    },
    excludeFields: ["thematicBreak", "rightTabStop", "leftTabStop", "includeIfEmpty"],
    notes:
      "Input-side sugar excluded from the field sets: thematicBreak (→pBdr/border), " +
      "rightTabStop/leftTabStop (→tabs/tabStops), includeIfEmpty (control flag) — these map field→XML but XML→a different field, breaking the 1:1 round-trip assumption.",
  },
];

export function findFieldSpec(id: string): DescriptorFieldSpec | undefined {
  return FIELD_SPECS.find((s) => s.id === id);
}
