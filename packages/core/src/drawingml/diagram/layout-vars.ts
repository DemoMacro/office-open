/**
 * Diagram layout variable property set elements.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd
 * CT_LayoutVariablePropertySet and its child elements.
 *
 * @module
 */
import { element } from "@office-open/xml";

// ---------------------------------------------------------------------------
// dgm:adj — adjustment handle (CT_Adj)
// ---------------------------------------------------------------------------

export interface AdjustOptions {
  /** 1-based index (required) */
  idx: number;
  /** Adjustment value (required) */
  val: number;
}

/** Creates a dgm:adj element. */
export const createAdjust = (options: AdjustOptions): string =>
  `<dgm:adj idx="${options.idx}" val="${options.val}"/>`;

// ---------------------------------------------------------------------------
// dgm:animLvl — animation level (CT_AnimLvl)
// ---------------------------------------------------------------------------

export const AnimationLevelValue = {
  NONE: "none",
  LEVEL: "lvl",
  CENTER: "ctr",
} as const;

export interface AnimationLevelOptions {
  val?: (typeof AnimationLevelValue)[keyof typeof AnimationLevelValue];
}

/** Creates a dgm:animLvl element. */
export const createAnimationLevel = (options?: AnimationLevelOptions): string =>
  options?.val !== undefined ? `<dgm:animLvl val="${options.val}"/>` : "<dgm:animLvl/>";

// ---------------------------------------------------------------------------
// dgm:animOne — animation one-by-one (CT_AnimOne)
// ---------------------------------------------------------------------------

export const AnimateOneByOneValue = {
  NONE: "none",
  ONE: "one",
  BRANCH: "branch",
} as const;

export interface AnimateOneByOneOptions {
  val?: (typeof AnimateOneByOneValue)[keyof typeof AnimateOneByOneValue];
}

/** Creates a dgm:animOne element. */
export const createAnimateOneByOne = (options?: AnimateOneByOneOptions): string =>
  options?.val !== undefined ? `<dgm:animOne val="${options.val}"/>` : "<dgm:animOne/>";

// ---------------------------------------------------------------------------
// dgm:chMax — max children constraint (CT_ChildMax)
// ---------------------------------------------------------------------------

export interface MaxChildrenOptions {
  val?: number;
}

/** Creates a dgm:chMax element. */
export const createMaxChildren = (options?: MaxChildrenOptions): string =>
  options?.val !== undefined ? `<dgm:chMax val="${options.val}"/>` : "<dgm:chMax/>";

// ---------------------------------------------------------------------------
// dgm:chPref — preferred children count (CT_ChildPref)
// ---------------------------------------------------------------------------

export interface PreferredChildrenOptions {
  val?: number;
}

/** Creates a dgm:chPref element. */
export const createPreferredChildren = (options?: PreferredChildrenOptions): string =>
  options?.val !== undefined ? `<dgm:chPref val="${options.val}"/>` : "<dgm:chPref/>";

// ---------------------------------------------------------------------------
// dgm:orgChart — organization chart flag (CT_OrgChart)
// ---------------------------------------------------------------------------

export interface OrgChartOptions {
  val?: boolean;
}

/** Creates a dgm:orgChart element. */
export const createOrgChart = (options?: OrgChartOptions): string =>
  options?.val !== undefined ? `<dgm:orgChart val="${options.val}"/>` : "<dgm:orgChart/>";

// ---------------------------------------------------------------------------
// dgm:hierBranch — hierarchy branch style (CT_HierBranchStyle)
// ---------------------------------------------------------------------------

export const HierBranchStyle = {
  LEFT: "l",
  RIGHT: "r",
  HANGING: "hang",
  STANDARD: "std",
  INITIAL: "init",
} as const;

export interface HierBranchOptions {
  val?: (typeof HierBranchStyle)[keyof typeof HierBranchStyle];
}

/** Creates a dgm:hierBranch element. */
export const createHierBranch = (options?: HierBranchOptions): string =>
  options?.val !== undefined ? `<dgm:hierBranch val="${options.val}"/>` : "<dgm:hierBranch/>";

// ---------------------------------------------------------------------------
// dgm:presLayoutVars — presentation layout variables (CT_LayoutVariablePropertySet)
// ---------------------------------------------------------------------------

export interface PresentationLayoutVariablesOptions {
  orgChart?: OrgChartOptions;
  maxChildren?: MaxChildrenOptions;
  preferredChildren?: PreferredChildrenOptions;
  animateOneByOne?: AnimateOneByOneOptions;
  animationLevel?: AnimationLevelOptions;
  hierBranch?: HierBranchOptions;
}

/**
 * Creates a dgm:presLayoutVars element (CT_LayoutVariablePropertySet).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LayoutVariablePropertySet">
 *   <xsd:sequence>
 *     <xsd:element name="orgChart" type="CT_OrgChart" minOccurs="0"/>
 *     <xsd:element name="chMax" type="CT_ChildMax" minOccurs="0"/>
 *     <xsd:element name="chPref" type="CT_ChildPref" minOccurs="0"/>
 *     <xsd:element name="bulletEnabled" type="CT_BulletEnabled" minOccurs="0"/>
 *     <xsd:element name="dir" type="CT_Direction" minOccurs="0"/>
 *     <xsd:element name="hierBranch" type="CT_HierBranchStyle" minOccurs="0"/>
 *     <xsd:element name="animOne" type="CT_AnimOne" minOccurs="0"/>
 *     <xsd:element name="animLvl" type="CT_AnimLvl" minOccurs="0"/>
 *     <xsd:element name="resizeHandles" type="CT_ResizeHandles" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createPresentationLayoutVariables = (
  options?: PresentationLayoutVariablesOptions,
): string => {
  const children: string[] = [];
  if (options?.orgChart) children.push(createOrgChart(options.orgChart));
  if (options?.maxChildren) children.push(createMaxChildren(options.maxChildren));
  if (options?.preferredChildren) children.push(createPreferredChildren(options.preferredChildren));
  if (options?.animateOneByOne) children.push(createAnimateOneByOne(options.animateOneByOne));
  if (options?.animationLevel) children.push(createAnimationLevel(options.animationLevel));
  if (options?.hierBranch) children.push(createHierBranch(options.hierBranch));

  return element("dgm:presLayoutVars", undefined, children);
};

// ---------------------------------------------------------------------------
// dgm:adjLst — adjustment list (CT_AdjLst)
// ---------------------------------------------------------------------------

export interface AdjustListOptions {
  adjustments?: readonly AdjustOptions[];
}

/** Creates a dgm:adjLst element containing dgm:adj children. */
export const createAdjustList = (options?: AdjustListOptions): string => {
  const children: string[] = [];
  if (options?.adjustments) {
    for (const a of options.adjustments) {
      children.push(createAdjust(a));
    }
  }
  return element("dgm:adjLst", undefined, children);
};
