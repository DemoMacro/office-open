/**
 * DrawingML diagram (SmartArt) elements — dgm: namespace.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 *
 * @module
 */

// Layout variable property set elements
export {
  createAdjust,
  createAdjustList,
  createAnimationLevel,
  createAnimateOneByOne,
  createMaxChildren,
  createPreferredChildren,
  createOrgChart,
  createHierBranch,
  createPresentationLayoutVariables,
  AnimationLevelValue,
  AnimateOneByOneValue,
  HierBranchStyle,
} from "./layout-vars";
export type {
  AdjustOptions,
  AdjustListOptions,
  AnimationLevelOptions,
  AnimateOneByOneOptions,
  MaxChildrenOptions,
  PreferredChildrenOptions,
  OrgChartOptions,
  HierBranchOptions,
  PresentationLayoutVariablesOptions,
} from "./layout-vars";

// Definition headers and lists
export {
  createColorsDefinitionHeader,
  createColorsDefinitionHeaderList,
  createLayoutDefinitionHeader,
  createLayoutDefinitionHeaderList,
  createStyleDefinitionHeader,
  createStyleDefinitionHeaderList,
} from "./headers";
export type {
  DiagramNameOptions,
  DiagramDescriptionOptions,
  DiagramCategoryOptions,
  ColorsDefinitionHeaderOptions,
  ColorsDefinitionHeaderListOptions,
  LayoutDefinitionHeaderOptions,
  LayoutDefinitionHeaderListOptions,
  StyleDefinitionHeaderOptions,
  StyleDefinitionHeaderListOptions,
} from "./headers";

// Style, color lists, style label
export {
  createDiagramStyle,
  createFillColorList,
  createLineColorList,
  createEffectColorList,
  createTextFillColorList,
  createTextLineColorList,
  createTextEffectColorList,
  createStyleLabel,
  ColorMethod,
  HueDirection,
  StyleMatrixIndex,
  FontCollectionIndex,
} from "./diagram-style";
export type {
  DiagramStyleOptions,
  ColorListOptions,
  DiagramStyleLabelOptions,
} from "./diagram-style";

// Relationship IDs
export { createDiagramRelationshipIds } from "./diagram-rel";
export type { DiagramRelationshipIdsOptions } from "./diagram-rel";

// Extension list, shape3D, text properties
export {
  createDiagramExtensionList,
  createDiagramShape3D,
  createDiagramTextProperties,
} from "./diagram-props";
export type {
  DiagramExtensionListOptions,
  DiagramExtensionOptions,
  DiagramTextPropertiesOptions,
} from "./diagram-props";
