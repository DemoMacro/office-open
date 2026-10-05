import {
  DATABASE_SCHEMA_ELEMENTS,
  DRAWING_VOCABULARY_SCHEMA_ELEMENTS,
  FORM_SCHEMA_ELEMENTS,
  MATH_SCHEMA_ELEMENTS,
  MANIFEST_SCHEMA_ELEMENTS,
} from "../../packages/odf/src/index";
import { rngElementDescriptors } from "./odf-rng";

const ODF_SCHEMA = ["odf-schemas/OpenDocument-v1.3-schema.rng"] as const;

export type OdfCodecClassification = "canonical" | "subdocument" | "unsupported" | "generic-only";

export interface OdfCodecOwner {
  package: string;
  module: string;
  export: string | null;
}

export interface OdfCodecEntry {
  id: string;
  classification: OdfCodecClassification;
  owner: OdfCodecOwner;
  schemaPrefixes: readonly string[];
  schemaElements?: readonly { name: string; attributes: readonly string[] }[];
  roundTrip?: readonly [generate: string, parse: string];
  testId: string;
  fixtureKey: string;
  negativeCapabilities?: readonly { element: string; testId: string }[];
}

export const ODF_CODEC_REGISTRY: readonly OdfCodecEntry[] = [
  {
    id: "manifest",
    classification: "subdocument",
    owner: { package: "odf", module: "src/container/manifest.ts", export: "parseManifestOptions" },
    schemaPrefixes: ["manifest"],
    schemaElements: MANIFEST_SCHEMA_ELEMENTS,
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
    schemaPrefixes: ["dr3d", "math", "svg"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["dr3d", "math", "svg"]),
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
    schemaPrefixes: ["config", "dc", "meta", "office", "style", "table", "text"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, [
      "config",
      "dc",
      "meta",
      "office",
      "style",
      "table",
      "text",
    ]),
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
    schemaPrefixes: ["number", "office", "style", "table", "text"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, [
      "number",
      "office",
      "style",
      "table",
      "text",
    ]),
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
    schemaPrefixes: ["anim", "draw", "office", "presentation", "style", "text"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, [
      "anim",
      "draw",
      "office",
      "presentation",
      "style",
      "text",
    ]),
    roundTrip: ["generatePresentation", "parsePresentation"],
    testId: "ODP presentation semantic round trip",
    fixtureKey: "odp-notes",
    negativeCapabilities: [
      { element: "draw:unknown", testId: "ODP rejects foreign slide element" },
    ],
  },
  {
    id: "chart",
    classification: "subdocument",
    owner: { package: "odf", module: "src/chart/chart.ts", export: "parseChartDocument" },
    schemaPrefixes: ["chart"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["chart"]),
    roundTrip: ["generateChartDocument", "parseChartDocument"],
    testId: "chart subdocument semantic round trip",
    fixtureKey: "chart-column-series",
    negativeCapabilities: [{ element: "chart:foo", testId: "chart rejects foreign element" }],
  },
  {
    id: "database",
    classification: "subdocument",
    owner: { package: "odf", module: "src/database/database.ts", export: "parseDatabaseDocument" },
    schemaPrefixes: ["db"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["db"]),
    roundTrip: ["generateDatabaseDocument", "parseDatabaseDocument"],
    testId: "database subdocument semantic round trip",
    fixtureKey: "database-connection",
    negativeCapabilities: [{ element: "db:unknown", testId: "database rejects foreign element" }],
  },
  {
    id: "forms-and-xforms",
    classification: "subdocument",
    owner: { package: "odf", module: "src/forms/forms.ts", export: "parseOfficeForms" },
    schemaPrefixes: ["form", "script", "xforms"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["form", "script", "xforms"]),
    roundTrip: ["officeFormsXml", "parseOfficeForms"],
    testId: "form controls and XForms semantic round trip",
    fixtureKey: "form-grid-xforms",
    negativeCapabilities: [{ element: "form:unknown", testId: "forms reject foreign element" }],
  },
] as const;
