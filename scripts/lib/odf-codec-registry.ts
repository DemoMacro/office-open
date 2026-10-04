import { MANIFEST_SCHEMA_ELEMENTS } from "../../packages/ocf/src/manifest-model";
import { DATABASE_SCHEMA_ELEMENTS } from "../../packages/odf-schema/src/db";
import { FORM_SCHEMA_ELEMENTS } from "../../packages/odf-schema/src/forms";
import {
  DRAWING_VOCABULARY_SCHEMA_ELEMENTS,
  MATH_SCHEMA_ELEMENTS,
} from "../../packages/odf-schema/src/xml-vocabulary";
import { rngElementDescriptors } from "./odf-rng-capabilities";

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
}

export const ODF_CODEC_REGISTRY: readonly OdfCodecEntry[] = [
  {
    id: "manifest",
    classification: "subdocument",
    owner: { package: "ocf", module: "src/manifest-model.ts", export: "parseManifestOptions" },
    schemaPrefixes: ["manifest"],
    schemaElements: MANIFEST_SCHEMA_ELEMENTS,
    roundTrip: ["manifestOptionsXml", "parseManifestOptions"],
  },
  {
    id: "drawing-vocabulary",
    classification: "subdocument",
    owner: {
      package: "odf-schema",
      module: "src/xml-vocabulary.ts",
      export: "parseDrawingVocabulary",
    },
    schemaPrefixes: ["dr3d", "math", "svg"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["dr3d", "math", "svg"]),
    roundTrip: ["drawingVocabularyXml", "parseDrawingVocabulary"],
  },
  {
    id: "odt",
    classification: "canonical",
    owner: { package: "odt", module: "src/index.ts", export: "generateOdt" },
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
    roundTrip: ["generateOdt", "parseOdt"],
  },
  {
    id: "ods",
    classification: "canonical",
    owner: { package: "ods", module: "src/index.ts", export: "generateOds" },
    schemaPrefixes: ["number", "office", "style", "table", "text"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, [
      "number",
      "office",
      "style",
      "table",
      "text",
    ]),
    roundTrip: ["generateOds", "parseOds"],
  },
  {
    id: "odp",
    classification: "canonical",
    owner: { package: "odp", module: "src/index.ts", export: "generateOdp" },
    schemaPrefixes: ["anim", "draw", "office", "presentation", "style", "text"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, [
      "anim",
      "draw",
      "office",
      "presentation",
      "style",
      "text",
    ]),
    roundTrip: ["generateOdp", "parseOdp"],
  },
  {
    id: "chart",
    classification: "subdocument",
    owner: { package: "odf-schema", module: "src/chart.ts", export: "parseChartDocument" },
    schemaPrefixes: ["chart"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["chart"]),
    roundTrip: ["generateChartDocument", "parseChartDocument"],
  },
  {
    id: "database",
    classification: "subdocument",
    owner: { package: "odf-schema", module: "src/db.ts", export: "parseDatabaseDocument" },
    schemaPrefixes: ["db"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["db"]),
    roundTrip: ["generateDatabaseDocument", "parseDatabaseDocument"],
  },
  {
    id: "forms-and-xforms",
    classification: "subdocument",
    owner: { package: "odf-schema", module: "src/forms.ts", export: "parseOfficeForms" },
    schemaPrefixes: ["form", "script", "xforms"],
    schemaElements: rngElementDescriptors(ODF_SCHEMA, ["form", "script", "xforms"]),
    roundTrip: ["officeFormsXml", "parseOfficeForms"],
  },
] as const;
