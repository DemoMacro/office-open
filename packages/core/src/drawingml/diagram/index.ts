/**
 * DrawingML diagram (SmartArt) elements — dgm: namespace.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 *
 * @module
 */

// Layout variable property set elements
export {
  createAdj,
  createAdjLst,
  createAnimLvl,
  createAnimOne,
  createChMax,
  createChPref,
  createOrgChart,
  createHierBranch,
  createPresLayoutVars,
  AnimLevelValue,
  AnimOneValue,
  HierBranchStyle,
} from "./layout-vars";
export type {
  AdjOptions,
  AdjLstOptions,
  AnimLvlOptions,
  AnimOneOptions,
  ChMaxOptions,
  ChPrefOptions,
  OrgChartOptions,
  HierBranchOptions,
  PresLayoutVarsOptions,
} from "./layout-vars";

// Definition headers and lists
export {
  createColorsDefHdr,
  createColorsDefHdrLst,
  createLayoutDefHdr,
  createLayoutDefHdrLst,
  createStyleDefHdr,
  createStyleDefHdrLst,
} from "./headers";
export type {
  DiagramNameOptions,
  DiagramDescriptionOptions,
  DiagramCategoryOptions,
  ColorsDefHdrOptions,
  ColorsDefHdrLstOptions,
  LayoutDefHdrOptions,
  LayoutDefHdrLstOptions,
  StyleDefHdrOptions,
  StyleDefHdrLstOptions,
} from "./headers";

// Style, color lists, style label
export {
  createDiagramStyle,
  createFillClrLst,
  createLinClrLst,
  createEffectClrLst,
  createTxFillClrLst,
  createTxLinClrLst,
  createTxEffectClrLst,
  createStyleLbl,
  ColorMethod,
  HueDirection,
  StyleMatrixIndex,
  FontCollectionIndex,
} from "./diagram-style";
export type {
  DiagramStyleOptions,
  ColorListOptions,
  DiagramStyleLblOptions,
} from "./diagram-style";

// Relationship IDs
export { createDiagramRelIds } from "./diagram-rel";
export type { DiagramRelIdsOptions } from "./diagram-rel";

// Extension list, sp3d, txPr
export { createDiagramExtLst, createDiagramSp3d, createDiagramTxPr } from "./diagram-props";
export type {
  DiagramExtLstOptions,
  DiagramExtensionOptions,
  DiagramTextPropsOptions,
} from "./diagram-props";
