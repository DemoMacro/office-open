import {
  DATABASE_SCHEMA_ELEMENTS,
  DRAWING_VOCABULARY_SCHEMA_ELEMENTS,
  FORM_SCHEMA_ELEMENTS,
  MATH_SCHEMA_ELEMENTS,
} from "../../packages/odf/src/index";
import type { SchemaElementDescriptor } from "./odf-rng";

export type OdfCodecClassification = "canonical" | "subdocument" | "unsupported" | "generic-only";
export type OdfPart = "content.xml" | "styles.xml" | "meta.xml" | "META-INF/manifest.xml";

export interface OdfCodecOwner {
  package: string;
  module: string;
  export: string | null;
}

export interface OdfCodecEntry {
  id: string;
  classification: OdfCodecClassification;
  owner: OdfCodecOwner;
  parts: readonly OdfPart[];
  /** Explicitly owned element/attribute names; ownership is never inferred from RNG. */
  schemaElements?: readonly { name: string; attributes: readonly string[] }[];
  roundTrip?: readonly [generate: string, parse: string];
  testId: string;
  fixtureKey: string;
  negativeCapabilities?: readonly { element: string; testId: string }[];
  rawFallbacks?: readonly string[];
}

const CONTENT_PART: readonly OdfPart[] = ["content.xml"];
const DOCUMENT_PARTS: readonly OdfPart[] = ["content.xml", "styles.xml", "meta.xml"];

const ODT_CONTENT: readonly SchemaElementDescriptor[] = [
  { name: "office:document-content", attributes: [] },
  { name: "office:text", attributes: [] },
  { name: "text:p", attributes: ["text:style-name"] },
  { name: "text:span", attributes: ["text:style-name"] },
  { name: "table:table", attributes: ["table:name"] },
  { name: "table:table-row", attributes: [] },
  { name: "table:table-cell", attributes: ["office:value-type"] },
];
const ODT_STYLES: readonly SchemaElementDescriptor[] = [
  { name: "office:document-styles", attributes: ["office:version"] },
  { name: "office:styles", attributes: [] },
  { name: "style:style", attributes: ["style:name", "style:family"] },
  { name: "style:paragraph-properties", attributes: ["fo:text-align"] },
  { name: "style:text-properties", attributes: ["fo:font-weight", "fo:font-style"] },
];
const ODT_META: readonly SchemaElementDescriptor[] = [
  { name: "office:document-meta", attributes: ["office:version"] },
  { name: "office:meta", attributes: [] },
  { name: "dc:title", attributes: [] },
  { name: "dc:creator", attributes: [] },
  { name: "dc:date", attributes: [] },
];
const ODS_CONTENT: readonly SchemaElementDescriptor[] = [
  { name: "office:document-content", attributes: [] },
  { name: "office:spreadsheet", attributes: [] },
  { name: "table:table", attributes: ["table:name"] },
  { name: "table:table-column", attributes: ["table:number-columns-repeated"] },
  { name: "table:table-row", attributes: ["table:style-name"] },
  { name: "table:table-cell", attributes: ["office:value-type", "office:value", "table:formula"] },
  { name: "draw:frame", attributes: ["draw:name", "svg:x", "svg:y", "svg:width", "svg:height"] },
  { name: "draw:image", attributes: ["xlink:href"] },
  { name: "draw:object", attributes: ["xlink:href"] },
];
const ODS_STYLES: readonly SchemaElementDescriptor[] = [
  { name: "office:document-styles", attributes: ["office:version"] },
  { name: "office:styles", attributes: [] },
  { name: "style:style", attributes: ["style:name", "style:family"] },
  { name: "style:table-row-properties", attributes: ["style:row-height"] },
  { name: "style:paragraph-properties", attributes: ["fo:line-height"] },
];
const ODS_META: readonly SchemaElementDescriptor[] = ODT_META;
const ODP_CONTENT: readonly SchemaElementDescriptor[] = [
  { name: "office:document-content", attributes: [] },
  { name: "office:presentation", attributes: [] },
  { name: "draw:page", attributes: ["draw:name"] },
  { name: "draw:frame", attributes: ["svg:x", "svg:y", "svg:width", "svg:height"] },
  { name: "draw:text-box", attributes: [] },
  { name: "presentation:notes", attributes: [] },
  { name: "text:p", attributes: ["text:style-name"] },
];
const ODP_STYLES: readonly SchemaElementDescriptor[] = [
  { name: "office:document-styles", attributes: ["office:version"] },
  { name: "office:styles", attributes: [] },
  { name: "style:master-page", attributes: ["style:name", "style:page-layout-name"] },
];
const ODP_META: readonly SchemaElementDescriptor[] = ODT_META;

export const ODF_CODEC_REGISTRY: readonly OdfCodecEntry[] = [
  {
    id: "manifest",
    classification: "subdocument",
    owner: { package: "odf", module: "src/container/manifest.ts", export: "parseManifestOptions" },
    parts: ["META-INF/manifest.xml"],
    schemaElements: [
      { name: "manifest:manifest", attributes: ["manifest:version"] },
      {
        name: "manifest:file-entry",
        attributes: [
          "manifest:full-path",
          "manifest:media-type",
          "manifest:preferred-view-mode",
          "manifest:size",
          "manifest:version",
        ],
      },
      {
        name: "manifest:encryption-data",
        attributes: ["manifest:checksum", "manifest:checksum-type"],
      },
      {
        name: "manifest:algorithm",
        attributes: ["manifest:algorithm-name", "manifest:initialisation-vector"],
      },
      {
        name: "manifest:key-derivation",
        attributes: [
          "manifest:iteration-count",
          "manifest:key-derivation-name",
          "manifest:key-size",
          "manifest:salt",
        ],
      },
      {
        name: "manifest:start-key-generation",
        attributes: ["manifest:key-size", "manifest:start-key-generation-name"],
      },
      { name: "manifest:keyinfo", attributes: [] },
      { name: "manifest:encrypted-key", attributes: [] },
      { name: "manifest:encryption-method", attributes: ["manifest:PGPAlgorithm"] },
      { name: "manifest:PGPData", attributes: [] },
      { name: "manifest:PGPKeyID", attributes: [] },
      { name: "manifest:PGPKeyPacket", attributes: [] },
      { name: "manifest:CipherData", attributes: [] },
      { name: "manifest:CipherValue", attributes: [] },
    ],
    roundTrip: ["manifestOptionsXml", "parseManifestOptions"],
    testId: "manifest semantic round trip",
    fixtureKey: "manifest-package-entry",
    negativeCapabilities: [
      { element: "office:unknown", testId: "manifest rejects foreign element" },
    ],
  },
  {
    id: "drawing-vocabulary",
    classification: "subdocument",
    owner: {
      package: "odf",
      module: "src/runtime/xml-vocabulary.ts",
      export: "parseDrawingVocabulary",
    },
    parts: CONTENT_PART,
    schemaElements: [
      { name: "svg:desc", attributes: [] },
      { name: "svg:title", attributes: [] },
      { name: "math:math", attributes: [] },
    ],
    roundTrip: ["drawingVocabularyXml", "parseDrawingVocabulary"],
    testId: "drawing vocabulary semantic round trip",
    fixtureKey: "drawing-vocabulary-desc",
    negativeCapabilities: [
      { element: "draw:unknown", testId: "drawing vocabulary rejects foreign element" },
    ],
  },
  {
    id: "odt",
    classification: "canonical",
    owner: { package: "odt", module: "src/index.ts", export: "generateDocument" },
    parts: DOCUMENT_PARTS,
    schemaElements: [...ODT_CONTENT, ...ODT_STYLES, ...ODT_META],
    roundTrip: ["generateDocument", "parseDocument"],
    testId: "ODT document semantic round trip",
    fixtureKey: "odt-paragraph",
    negativeCapabilities: [
      { element: "office:unknown", testId: "ODT rejects foreign body element" },
    ],
  },
  {
    id: "ods",
    classification: "canonical",
    owner: { package: "ods", module: "src/index.ts", export: "generateWorkbook" },
    parts: DOCUMENT_PARTS,
    schemaElements: [...ODS_CONTENT, ...ODS_STYLES, ...ODS_META],
    roundTrip: ["generateWorkbook", "parseWorkbook"],
    testId: "ODS workbook semantic round trip",
    fixtureKey: "ods-cell",
    negativeCapabilities: [
      { element: "office:unknown", testId: "ODS rejects foreign body element" },
    ],
  },
  {
    id: "odp",
    classification: "canonical",
    owner: { package: "odp", module: "src/index.ts", export: "generatePresentation" },
    parts: DOCUMENT_PARTS,
    schemaElements: [...ODP_CONTENT, ...ODP_STYLES, ...ODP_META],
    roundTrip: ["generatePresentation", "parsePresentation"],
    testId: "ODP presentation semantic round trip",
    fixtureKey: "odp-notes",
    negativeCapabilities: [
      { element: "draw:unknown", testId: "ODP rejects foreign slide element" },
    ],
  },
  {
    id: "ods-style-overlay",
    classification: "subdocument",
    owner: { package: "ods", module: "src/index.ts", export: "generateWorkbook" },
    parts: ["content.xml", "styles.xml"],
    schemaElements: [
      { name: "style:table-row-properties", attributes: ["style:row-height"] },
      { name: "style:paragraph-properties", attributes: ["fo:line-height"] },
    ],
    roundTrip: ["generateWorkbook", "parseWorkbook"],
    testId: "ODS raw style overlay semantic round trip",
    fixtureKey: "percent-length-overlay",
    rawFallbacks: ["style:table-row-properties", "style:paragraph-properties"],
  },
  {
    id: "ods-linked-graphics",
    classification: "subdocument",
    owner: { package: "ods", module: "src/index.ts", export: "generateWorkbook" },
    parts: CONTENT_PART,
    schemaElements: [
      {
        name: "draw:image",
        attributes: ["xlink:actuate", "xlink:href", "xlink:show", "xlink:type"],
      },
    ],
    roundTrip: ["generateWorkbook", "parseWorkbook"],
    testId: "ODS linked graphic semantic round trip",
    fixtureKey: "linked-graphics",
  },
  {
    id: "ods-embedded-chart",
    classification: "subdocument",
    owner: { package: "ods", module: "src/index.ts", export: "generateWorkbook" },
    parts: CONTENT_PART,
    schemaElements: [
      {
        name: "draw:frame",
        attributes: ["svg:height", "svg:width", "svg:x", "svg:y"],
      },
      { name: "draw:object", attributes: ["xlink:href"] },
    ],
    roundTrip: ["generateWorkbook", "parseWorkbook"],
    testId: "ODS embedded chart semantic round trip",
    fixtureKey: "chart-embedded-object",
  },
  {
    id: "manifest-encrypted",
    classification: "subdocument",
    owner: { package: "odf", module: "src/container/manifest.ts", export: "parseManifestOptions" },
    parts: ["META-INF/manifest.xml"],
    schemaElements: [
      {
        name: "manifest:encryption-data",
        attributes: ["manifest:checksum", "manifest:checksum-type"],
      },
      {
        name: "manifest:algorithm",
        attributes: ["manifest:algorithm-name", "manifest:initialisation-vector"],
      },
      {
        name: "manifest:start-key-generation",
        attributes: ["manifest:key-size", "manifest:start-key-generation-name"],
      },
      {
        name: "manifest:key-derivation",
        attributes: [
          "manifest:iteration-count",
          "manifest:key-derivation-name",
          "manifest:key-size",
          "manifest:salt",
        ],
      },
      { name: "manifest:encrypted-key", attributes: [] },
      { name: "manifest:encryption-method", attributes: ["manifest:PGPAlgorithm"] },
    ],
    roundTrip: ["manifestOptionsXml", "parseManifestOptions"],
    testId: "encrypted manifest semantic round trip",
    fixtureKey: "encrypted-manifest",
  },
  {
    id: "database-formula-embedded-object",
    classification: "subdocument",
    owner: { package: "odf", module: "src/database/database.ts", export: "parseDatabaseDocument" },
    parts: CONTENT_PART,
    schemaElements: [
      { name: "table:table-cell", attributes: ["table:formula"] },
      { name: "draw:frame", attributes: ["svg:height", "svg:width"] },
      { name: "draw:object", attributes: ["xlink:href"] },
    ],
    roundTrip: ["generateDatabaseDocument", "parseDatabaseDocument"],
    testId: "database formula and embedded object semantic round trip",
    fixtureKey: "database-formula-embedded-object",
  },
  {
    id: "chart",
    classification: "subdocument",
    owner: { package: "odf", module: "src/chart/chart.ts", export: "parseChartDocument" },
    parts: CONTENT_PART,
    schemaElements: [
      { name: "office:document-content", attributes: ["office:version"] },
      { name: "office:chart", attributes: [] },
      { name: "chart:chart", attributes: ["chart:class"] },
      { name: "chart:title", attributes: [] },
      { name: "chart:plot-area", attributes: ["chart:style-name"] },
      { name: "chart:series", attributes: ["chart:values-cell-range-address"] },
    ],
    roundTrip: ["generateChartDocument", "parseChartDocument"],
    testId: "chart subdocument semantic round trip",
    fixtureKey: "chart-column-series",
    negativeCapabilities: [{ element: "chart:foo", testId: "chart rejects foreign element" }],
  },
  {
    id: "database",
    classification: "subdocument",
    owner: { package: "odf", module: "src/database/database.ts", export: "parseDatabaseDocument" },
    parts: CONTENT_PART,
    schemaElements: [
      { name: "office:document-content", attributes: ["office:version"] },
      { name: "office:database", attributes: [] },
      { name: "db:data-source", attributes: [] },
      { name: "db:connection-data", attributes: [] },
      { name: "db:database-description", attributes: [] },
    ],
    roundTrip: ["generateDatabaseDocument", "parseDatabaseDocument"],
    testId: "database subdocument semantic round trip",
    fixtureKey: "database-connection",
    negativeCapabilities: [{ element: "db:unknown", testId: "database rejects foreign element" }],
  },
  {
    id: "forms-and-xforms",
    classification: "subdocument",
    owner: { package: "odf", module: "src/forms/forms.ts", export: "parseOfficeForms" },
    parts: CONTENT_PART,
    schemaElements: [
      { name: "office:forms", attributes: [] },
      { name: "form:form", attributes: ["form:name"] },
      { name: "form:grid", attributes: ["form:id"] },
      { name: "form:text", attributes: ["form:id"] },
    ],
    roundTrip: ["officeFormsXml", "parseOfficeForms"],
    testId: "form controls and XForms semantic round trip",
    fixtureKey: "form-grid-xforms",
    negativeCapabilities: [{ element: "form:unknown", testId: "forms reject foreign element" }],
  },
] as const;
