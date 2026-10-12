import type { ModelOwnershipEntry } from "./types";

const OWNER = "packages/core/src/vector/shapes.ts:stringifyVmlShapeChild";
const RUNTIME_OWNER = "packages/core/src/vector/shape-elements";
const EVIDENCE = [
  "packages/core/src/vector/vector.spec.ts",
  "packages/docx/src/parts/pict/pict.spec.ts",
] as const;

function modeled(localName: string, canonicalOptionPath: string): ModelOwnershipEntry {
  return {
    domain: "vml-main",
    qname: `v:${localName}`,
    canonicalOptionPath,
    classification: "planned",
    classificationReason:
      "Canonical VML parse/stringify may own this element, but executable semantic round-trip evidence is pending.",
    owner: {
      stringify: localName === "fill" || localName === "path" ? RUNTIME_OWNER : OWNER,
      parse: localName === "fill" || localName === "path" ? RUNTIME_OWNER : OWNER,
      compiler: "DOCX pict and XLSX/PPTX VML writers rebuild the typed shape tree.",
      publicApi: canonicalOptionPath,
    },
    evidence: { static: EVIDENCE, executable: EVIDENCE, rawOnly: false },
  };
}

export const VML_MAIN_OWNERSHIP: readonly ModelOwnershipEntry[] = [
  modeled("arc", "VmlArcOptions"),
  modeled("background", "VmlBackgroundOptions"),
  modeled("curve", "VmlCurveOptions"),
  modeled("f", "VmlFormulasOptions.equations"),
  modeled("fill", "VmlFillOptions"),
  modeled("formulas", "VmlFormulasOptions"),
  modeled("group", "VmlGroupOptions"),
  modeled("h", "VmlHandlesOptions.handles"),
  modeled("handles", "VmlHandlesOptions"),
  modeled("image", "VmlImageOptions"),
  modeled("imagedata", "VmlImageDataOptions"),
  modeled("line", "VmlLineOptions"),
  modeled("oval", "VmlOvalOptions"),
  modeled("path", "VmlPathOptions"),
  modeled("polyline", "VmlPolylineOptions"),
  modeled("rect", "VmlRectOptions"),
  modeled("roundrect", "VmlRoundRectOptions"),
  modeled("shadow", "VmlShadowOptions"),
  modeled("shape", "VmlShapeOptions"),
  modeled("shapetype", "VmlShapetypeOptions"),
  modeled("stroke", "VmlStrokeOptions"),
  modeled("textbox", "VmlTextboxOptions"),
  modeled("textpath", "VmlTextPathOptions"),
];
