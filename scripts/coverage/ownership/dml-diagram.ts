import type { ModelOwner, ModelOwnershipEntry } from "./types";

const layoutOwner: ModelOwner = {
  stringify: "packages/core/src/smartart/layout-definition.ts:stringifyLayoutDefinitionPart",
  parse: "packages/core/src/smartart/layout-definition.ts:parseLayoutDefinition",
  compiler: "SmartArtCollection via format drawing compilers",
  publicApi: "SmartArtOptions.layout",
};

const colorOwner: ModelOwner = {
  stringify: "packages/core/src/smartart/color-definition.ts:stringifyColorDefinitionPart",
  parse: "packages/core/src/smartart/color-definition.ts:parseColorDefinition",
  compiler: "SmartArtCollection via format drawing compilers",
  publicApi: "SmartArtOptions.color",
};

const styleOwner: ModelOwner = {
  stringify: "packages/core/src/smartart/style-definition.ts:stringifyStyleDefinitionPart",
  parse: "packages/core/src/smartart/style-definition.ts:parseStyleDefinition",
  compiler: "SmartArtCollection via format drawing compilers",
  publicApi: "SmartArtOptions.style",
};

const dataModelOwner: ModelOwner = {
  stringify: "packages/core/src/smartart/data-model/data-model.ts:stringifyDataModelOptions",
  parse: "packages/core/src/smartart/data-model/data-model.ts:parseDataModelOptions",
  compiler: "SmartArtCollection via format drawing parsers",
  publicApi: "SmartArtOptions.dataModel",
};

const relationshipOwner: ModelOwner = {
  stringify: "packages/core/src/drawing/diagram/diagram-descriptors.ts:diagramRelationshipIdsDesc",
  parse: "packages/core/src/drawing/diagram/diagram-descriptors.ts:diagramRelationshipIdsDesc",
  compiler: "format drawing compilers rebuild SmartArt relationships",
  publicApi: "SmartArt relationship paths",
};

const presentationVariablesOwner: ModelOwner = {
  stringify:
    "packages/core/src/drawing/diagram/diagram-descriptors.ts:presentationLayoutVariablesDesc",
  parse: "packages/core/src/drawing/diagram/diagram-descriptors.ts:presentationLayoutVariablesDesc",
  compiler: "presentationLayoutVariablesDesc",
  publicApi: "PresentationLayoutVariablesOptions",
};

const extensionOwner: ModelOwner = {
  stringify: "packages/core/src/drawing/diagram/diagram-props.ts:stringifyDiagramExtensions",
  parse: "packages/core/src/drawing/diagram/diagram-props.ts:parseDiagramExtensions",
  compiler: "SmartArt definition and data-model writers",
  publicApi: "DiagramExtensionOptions",
};

function modeledDefinitionHeaders(
  qname: string,
  descriptorName: string,
  optionName: string,
  caseId: string,
): ModelOwnershipEntry {
  const reference = `packages/core/src/drawing/diagram/headers.ts:${descriptorName}`;
  return {
    domain: "dml-diagram",
    qname,
    canonicalOptionPath: optionName,
    classification: "modeled",
    classificationReason: "Canonical typed parse and stringify both own this header surface.",
    owner: {
      stringify: reference,
      parse: reference,
      compiler: "Standalone diagram definition-header descriptors",
      publicApi: optionName,
    },
    evidence: {
      static: DIAGRAM_DESCRIPTOR_EVIDENCE,
      executable: [caseId],
      rawOnly: false,
    },
  };
}

const DIAGRAM_DESCRIPTOR_EVIDENCE = [
  "packages/core/src/drawing/diagram/diagram-descriptors.spec.ts",
];

const DATA_MODEL_EVIDENCE = [
  "packages/core/src/smartart/data-model/data-model.spec.ts",
  "packages/docx/src/parts/drawing/smartart-opaque.spec.ts",
  "packages/pptx/src/parts/descriptors/smartart.spec.ts",
  "packages/pptx/src/parts/smartart-opaque.spec.ts",
  "packages/xlsx/src/smartart.spec.ts",
];

const LAYOUT_EVIDENCE = ["packages/core/src/smartart/layout-definition.spec.ts"];
const DATA_MODEL_CASE = "dml-diagram.data-model.round-trip";
const LAYOUT_CASE = "dml-diagram.layout-definition.round-trip";
const STYLE_CASE = "dml-diagram.style-definition.round-trip";
const COLOR_CASE = "dml-diagram.color-definition.round-trip";

function modeledDescriptor(qname: string): ModelOwnershipEntry {
  return {
    domain: "dml-diagram",
    qname,
    canonicalOptionPath: "PresentationLayoutVariablesOptions",
    classification: "modeled",
    classificationReason: "Canonical typed parse and stringify both own this element.",
    owner: presentationVariablesOwner,
    evidence: {
      static: DIAGRAM_DESCRIPTOR_EVIDENCE,
      executable: [DATA_MODEL_CASE],
      rawOnly: false,
    },
  };
}

function modeledExtensions(): ModelOwnershipEntry {
  return {
    domain: "dml-diagram",
    qname: "dgm:extLst",
    canonicalOptionPath: "DiagramExtensionOptions.content",
    classification: "modeled",
    classificationReason: "Foreign-namespace extension content is typed and round-trips.",
    owner: extensionOwner,
    evidence: {
      static: DIAGRAM_DESCRIPTOR_EVIDENCE,
      executable: [DATA_MODEL_CASE],
      rawOnly: false,
    },
  };
}

function modeled(
  qnames: readonly string[],
  owner: ModelOwner,
  canonicalOptionPath: string,
  evidence: readonly string[],
  executable: readonly string[] = evidence,
): ModelOwnershipEntry[] {
  return qnames.map((qname) => ({
    domain: "dml-diagram",
    qname,
    canonicalOptionPath,
    classification: "modeled",
    classificationReason: "Canonical typed parse and stringify both own this element.",
    owner,
    evidence: { static: evidence, executable, rawOnly: false },
  }));
}

function planned(
  qnames: readonly string[],
  owner: ModelOwner,
  canonicalOptionPath: string,
  evidence: readonly string[],
): ModelOwnershipEntry[] {
  return qnames.map((qname) => ({
    domain: "dml-diagram",
    qname,
    canonicalOptionPath,
    classification: "planned",
    classificationReason:
      "Canonical implementation exists, but package-level executable evidence is still pending.",
    owner,
    evidence: { static: evidence, executable: [], rawOnly: false },
  }));
}

export const DML_DIAGRAM_OWNERSHIP: readonly ModelOwnershipEntry[] = [
  ...modeled(
    [
      "dgm:layoutDef",
      "dgm:title",
      "dgm:catLst",
      "dgm:cat",
      "dgm:layoutNode",
      "dgm:alg",
      "dgm:param",
      "dgm:shape",
      "dgm:presOf",
      "dgm:constrLst",
      "dgm:constr",
      "dgm:ruleLst",
      "dgm:rule",
      "dgm:varLst",
      "dgm:orgChart",
      "dgm:chMax",
      "dgm:bulletEnabled",
      "dgm:dir",
      "dgm:resizeHandles",
      "dgm:adjLst",
      "dgm:adj",
    ],
    layoutOwner,
    "SmartArtOptions.layout",
    LAYOUT_EVIDENCE,
    [LAYOUT_CASE],
  ),
  ...planned(
    [
      "dgm:desc",
      "dgm:chPref",
      "dgm:hierBranch",
      "dgm:animOne",
      "dgm:animLvl",
      "dgm:forEach",
      "dgm:choose",
      "dgm:if",
      "dgm:else",
      "dgm:sampData",
      "dgm:styleData",
      "dgm:clrData",
    ],
    layoutOwner,
    "SmartArtOptions.layout",
    LAYOUT_EVIDENCE,
  ),
  ...modeled(
    ["dgm:colorsDef", "dgm:styleLbl", "dgm:fillClrLst", "dgm:linClrLst", "dgm:txFillClrLst"],
    colorOwner,
    "SmartArtOptions.color",
    LAYOUT_EVIDENCE,
    [COLOR_CASE],
  ),
  ...planned(
    ["dgm:effectClrLst", "dgm:txLinClrLst", "dgm:txEffectClrLst"],
    colorOwner,
    "SmartArtOptions.color",
    LAYOUT_EVIDENCE,
  ),
  ...modeled(["dgm:styleDef", "dgm:style"], styleOwner, "SmartArtOptions.style", LAYOUT_EVIDENCE, [
    STYLE_CASE,
  ]),
  ...planned(
    ["dgm:scene3d", "dgm:sp3d", "dgm:txPr"],
    styleOwner,
    "SmartArtOptions.style",
    LAYOUT_EVIDENCE,
  ),
  ...modeled(
    [
      "dgm:dataModel",
      "dgm:ptLst",
      "dgm:cxnLst",
      "dgm:pt",
      "dgm:cxn",
      "dgm:prSet",
      "dgm:spPr",
      "dgm:t",
      "dgm:bg",
      "dgm:whole",
    ],
    dataModelOwner,
    "SmartArtOptions.dataModel",
    DATA_MODEL_EVIDENCE,
    [DATA_MODEL_CASE],
  ),
  ...modeled(
    ["dgm:relIds"],
    relationshipOwner,
    "SmartArt relationship identity",
    [
      "packages/core/src/drawing/diagram/diagram-descriptors.spec.ts",
      "packages/docx/src/parts/drawing/smartart-opaque.spec.ts",
      "packages/pptx/src/parts/smartart-opaque.spec.ts",
    ],
    ["dml-diagram.relationship-ids.round-trip"],
  ),
  modeledDefinitionHeaders(
    "dgm:colorsDefHdr",
    "colorsDefinitionHeaderDesc",
    "ColorsDefinitionHeaderOptions",
    "dml-diagram.colors-definition-header.round-trip",
  ),
  modeledDefinitionHeaders(
    "dgm:colorsDefHdrLst",
    "colorsDefinitionHeaderListDesc",
    "ColorsDefinitionHeaderListOptions",
    "dml-diagram.colors-definition-header.round-trip",
  ),
  modeledDefinitionHeaders(
    "dgm:layoutDefHdr",
    "layoutDefinitionHeaderDesc",
    "LayoutDefinitionHeaderOptions",
    "dml-diagram.layout-definition-header.round-trip",
  ),
  modeledDefinitionHeaders(
    "dgm:layoutDefHdrLst",
    "layoutDefinitionHeaderListDesc",
    "LayoutDefinitionHeaderListOptions",
    "dml-diagram.layout-definition-header.round-trip",
  ),
  modeledDefinitionHeaders(
    "dgm:styleDefHdr",
    "styleDefinitionHeaderDesc",
    "StyleDefinitionHeaderOptions",
    "dml-diagram.style-definition-header.round-trip",
  ),
  modeledDefinitionHeaders(
    "dgm:styleDefHdrLst",
    "styleDefinitionHeaderListDesc",
    "StyleDefinitionHeaderListOptions",
    "dml-diagram.style-definition-header.round-trip",
  ),
  modeledDescriptor("dgm:presLayoutVars"),
  modeledExtensions(),
];
