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
  roundTrip?: readonly [generate: string, parse: string];
}

export const ODF_CODEC_REGISTRY: readonly OdfCodecEntry[] = [
  {
    id: "odt",
    classification: "canonical",
    owner: { package: "odt", module: "src/index.ts", export: "generateOdt" },
    schemaPrefixes: ["config", "dc", "meta", "office", "style", "table", "text"],
    roundTrip: ["generateOdt", "parseOdt"],
  },
  {
    id: "ods",
    classification: "canonical",
    owner: { package: "ods", module: "src/index.ts", export: "generateOds" },
    schemaPrefixes: ["number", "office", "style", "table", "text"],
    roundTrip: ["generateOds", "parseOds"],
  },
  {
    id: "odp",
    classification: "canonical",
    owner: { package: "odp", module: "src/index.ts", export: "generateOdp" },
    schemaPrefixes: ["anim", "draw", "office", "presentation", "style", "text"],
    roundTrip: ["generateOdp", "parseOdp"],
  },
  {
    id: "chart",
    classification: "subdocument",
    owner: { package: "odf-schema", module: "src/chart.ts", export: "parseChartDocument" },
    schemaPrefixes: ["chart"],
    roundTrip: ["generateChartDocument", "parseChartDocument"],
  },
  {
    id: "database",
    classification: "subdocument",
    owner: { package: "odf-schema", module: "src/database.ts", export: "parseDatabaseDocument" },
    schemaPrefixes: ["db"],
    roundTrip: ["generateDatabaseDocument", "parseDatabaseDocument"],
  },
  {
    id: "generic-node",
    classification: "generic-only",
    owner: { package: "odf-schema", module: "src/node.ts", export: "OdfXmlNode" },
    schemaPrefixes: ["dr3d", "manifest", "math", "svg"],
  },
  {
    id: "forms-and-xforms",
    classification: "unsupported",
    owner: { package: "odf-schema", module: "src/forms.ts", export: null },
    schemaPrefixes: ["form", "script", "xforms"],
  },
] as const;
