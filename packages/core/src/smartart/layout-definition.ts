/**
 * SmartArt layout definition (dgm:layoutDef, CT_DiagramDefinition) — the
 * authoring-layer model describing how a diagram arranges its shapes.
 *
 * Bidirectional: stringify builds the layoutDef part body, parse reads one
 * back into the same options shape, so custom layouts round-trip intact.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd (CT_DiagramDefinition and the
 * CT_LayoutNode tree it contains).
 *
 * @module
 */

import { attr, attrNum, attrs, findChild } from "@office-open/xml";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import { stringify as stringifyInnerXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import type {
  DiagramCategoryOptions,
  DiagramDescriptionOptions,
  DiagramNameOptions,
} from "../drawing/diagram/headers";
import type {
  AdjustOptions,
  AnimateOneByOne,
  AnimationLevel,
  HierBranch,
} from "../drawing/diagram/layout-vars";
import { parseDataModelOptions, stringifyDataModelOptions } from "./data-model";
import type { DataModelOptions } from "./data-model";

// ── Enumerations (ST_* from dml-diagram.xsd, verbatim token values) ──

/** ST_AlgorithmType (dgm:alg `@type`). */
export type AlgorithmType =
  | "composite"
  | "conn"
  | "cycle"
  | "hierChild"
  | "hierRoot"
  | "pyra"
  | "lin"
  | "sp"
  | "tx"
  | "snake";

/** ST_ConstraintRelationship (dgm:constr `@for`/`@refFor`). */
export type ConstraintRelationship = "self" | "ch" | "des";

/** ST_ConstraintType (dgm:constr/`@type`/`@refType`, dgm:rule `@type`). */
export type ConstraintType =
  | "none"
  | "alignOff"
  | "begMarg"
  | "bendDist"
  | "begPad"
  | "b"
  | "bMarg"
  | "bOff"
  | "ctrX"
  | "ctrXOff"
  | "ctrY"
  | "ctrYOff"
  | "connDist"
  | "diam"
  | "endMarg"
  | "endPad"
  | "h"
  | "hArH"
  | "hOff"
  | "l"
  | "lMarg"
  | "lOff"
  | "r"
  | "rMarg"
  | "rOff"
  | "primFontSz"
  | "pyraAcctRatio"
  | "secFontSz"
  | "sibSp"
  | "secSibSp"
  | "sp"
  | "stemThick"
  | "t"
  | "tMarg"
  | "tOff"
  | "userA"
  | "userB"
  | "userC"
  | "userD"
  | "userE"
  | "userF"
  | "userG"
  | "userH"
  | "userI"
  | "userJ"
  | "userK"
  | "userL"
  | "userM"
  | "userN"
  | "userO"
  | "userP"
  | "userQ"
  | "userR"
  | "userS"
  | "userT"
  | "userU"
  | "userV"
  | "userW"
  | "userX"
  | "userY"
  | "userZ"
  | "w"
  | "wArH"
  | "wOff";

/** ST_BoolOperator (dgm:constr `@op`). */
export type ConstraintOperator = "none" | "equ" | "gte" | "lte";

/** ST_ChildOrderType (dgm:layoutNode `@chOrder`). */
export type ChildOrder = "b" | "t";

/** ST_FunctionType (dgm:if `@func`). */
export type ConditionFunction =
  | "cnt"
  | "pos"
  | "revPos"
  | "posEven"
  | "posOdd"
  | "var"
  | "depth"
  | "maxDepth";

/** ST_FunctionOperator (dgm:if `@op`). */
export type ConditionOperator = "equ" | "neq" | "gt" | "lt" | "gte" | "lte";

/** ST_Direction (dgm:dir `@val`). */
export type DiagramDirection = "norm" | "rev";

/** ST_ResizeHandlesStr (dgm:resizeHandles `@val`). */
export type ResizeHandles = "exact" | "rel";

// ST_FunctionArgument / ST_FunctionValue are unions of every layout-parameter
// enum plus numbers; ST_AxisTypes / ST_ElementTypes / ST_Booleans / ST_Ints are
// space-separated lists — all surface as plain string.

// dgm:title / dgm:desc reuse DiagramNameOptions / DiagramDescriptionOptions
// (same CT_Name shape as the headers), dgm:cat reuses DiagramCategoryOptions,
// and dgm:adjLst reuses AdjustOptions — all from drawing/diagram.

// ── Shared building blocks ──

/** dgm:sampData / dgm:styleData / dgm:clrData (CT_SampleData). */
export interface SampleDataOptions {
  /** Use the data model of the containing layout as sample (dgm:`@useDef`). */
  useDefault?: boolean;
  dataModel?: DataModelOptions;
}

/**
 * AG_IteratorAttributes — the traversal selector shared by presOf, forEach,
 * and when-branches. List-typed XSD attributes (axis, ptType, hideLastTrans,
 * st, cnt, step) surface as strings; single values read fine too.
 */
export interface DiagramIterationOptions {
  /** Selection axes, space-separated (dgm:`@axis`, ST_AxisTypes). */
  axis?: string;
  /** Selected point kinds, space-separated (dgm:`@ptType`, ST_ElementTypes). */
  pointType?: string;
  /** Hide the last transition point (dgm:`@hideLastTrans`, ST_Booleans). */
  hideLastTransition?: string;
  /** Start index (dgm:`@st`, ST_Ints). */
  start?: string;
  /** Iteration count (dgm:`@cnt`, ST_UnsignedInts). */
  count?: string;
  /** Index step (dgm:`@step`, ST_Ints). */
  step?: string;
}

/** dgm:alg (CT_Algorithm). */
export interface AlgorithmOptions {
  /** Layout algorithm (dgm:alg `@type`). */
  type: AlgorithmType;
  /** Revision counter (dgm:alg `@rev`). */
  revision?: number;
  /** Algorithm parameters (dgm:param*). */
  parameters?: AlgorithmParameterOptions[];
}

/** dgm:param (CT_Parameter) — inlined in AlgorithmOptions.parameters. */
export interface AlgorithmParameterOptions {
  /** Parameter name (dgm:param `@type`, ST_ParameterId). */
  type: string;
  /**
   * Parameter value (dgm:param `@val`, ST_ParameterVal — a union of direction
   * tokens, numbers, and booleans, so it stays polymorphic).
   */
  value: string | number | boolean;
}

/** dgm:shape (CT_Shape) — the geometry template a layout node renders as. */
export interface LayoutShapeOptions {
  /** Shape rotation as a raw double (dgm:shape `@rot`). */
  rotation?: number;
  /** Preset geometry name or "none"/"conn" (dgm:shape `@type`, ST_LayoutShapeType). */
  type?: string;
  /** Preview picture relationship id (dgm:shape `@r:blip`). */
  blip?: string;
  /** Z-order offset (dgm:shape `@zOrderOff`). */
  zOrderOffset?: number;
  /** Hide the shape geometry, keep the text (dgm:shape `@hideGeom`). */
  hideGeometry?: boolean;
  /** Lock text entry (dgm:shape `@lkTxEntry`). */
  lockTextEntry?: boolean;
  /** Use the picture placeholder (dgm:shape `@blipPhldr`). */
  blipPlaceholder?: boolean;
  /** Geometry adjustment values (dgm:adjLst). */
  adjustments?: AdjustOptions[];
}

/** dgm:constr (CT_Constraint). */
export interface ConstraintOptions {
  /** Constrained dimension (dgm:constr `@type`). */
  type: ConstraintType;
  /** Related point the constraint applies to (dgm:constr `@for`). */
  for?: ConstraintRelationship;
  /** Layout-node name the constraint applies to (dgm:constr `@forName`). */
  forName?: string;
  /** Point kinds the constraint applies to (dgm:constr `@ptType`). */
  pointType?: string;
  /** Referenced dimension (dgm:constr `@refType`). */
  referenceType?: ConstraintType;
  /** Relationship of the referenced point (dgm:constr `@refFor`). */
  referenceFor?: ConstraintRelationship;
  /** Layout-node name of the referenced point (dgm:constr `@refForName`). */
  referenceForName?: string;
  /** Point kinds of the referenced point (dgm:constr `@refPtType`). */
  referencePointType?: string;
  /** How the reference combines with val (dgm:constr `@op`). */
  operation?: ConstraintOperator;
  /** Absolute value (dgm:constr `@val`). */
  value?: number;
  /** Factor applied to the reference (dgm:constr `@fact`). */
  factor?: number;
}

/** dgm:rule (CT_NumericRule). */
export interface LayoutRuleOptions {
  /** Rule dimension (dgm:rule `@type`). */
  type: ConstraintType;
  for?: ConstraintRelationship;
  forName?: string;
  pointType?: string;
  /** Preferred value (dgm:rule `@val`). */
  value?: number;
  /** Multiplied factor (dgm:rule `@fact`). */
  factor?: number;
  /** Upper bound (dgm:rule `@max`). */
  maximum?: number;
}

/** dgm:varLst (CT_LayoutVariablePropertySet). */
export interface VariableListOptions {
  /** This is an organization chart (dgm:orgChart `@val`). */
  organizationChart?: boolean;
  /** Maximum children per parent (dgm:chMax `@val`). */
  childMaximum?: number;
  /** Preferred children per parent (dgm:chPref `@val`). */
  childPreferred?: number;
  /** Bullets enabled on text nodes (dgm:bulletEnabled `@val`). */
  bulletEnabled?: boolean;
  /** Reading direction (dgm:dir `@val`). */
  direction?: DiagramDirection;
  /** Branch style for hierarchy layouts (dgm:hierBranch `@val`). */
  hierarchyBranchStyle?: HierBranch;
  /** Animate siblings one by one or as a branch (dgm:animOne `@val`). */
  animateOne?: AnimateOneByOne;
  /** Animate by level or from center (dgm:animLvl `@val`). */
  animateLevel?: AnimationLevel;
  /** Resize-handle behavior (dgm:resizeHandles `@val`). */
  resizeHandles?: ResizeHandles;
}

/** One alternative inside a choose — dgm:if (CT_When); `if` is a reserved word. */
export interface WhenOptions extends DiagramIterationOptions {
  name?: string;
  /** Selection function (dgm:if `@func`). */
  function: ConditionFunction;
  /** Function argument, usually a variable name (dgm:if `@arg`). */
  argument?: string;
  /** Comparison operator (dgm:if `@op`). */
  operator: ConditionOperator;
  /** Comparison value (dgm:if `@val`, ST_FunctionValue union). */
  value: string;
  children?: LayoutNodeChild[];
}

/** dgm:else (CT_Otherwise). */
export interface OtherwiseOptions {
  name?: string;
  children?: LayoutNodeChild[];
}

/** dgm:choose (CT_Choose): pick the first matching condition. */
export interface ChooseOptions {
  name?: string;
  /** Alternatives in priority order (dgm:if*, CT_When). */
  conditions: WhenOptions[];
  /** Fallback branch (dgm:else, CT_Otherwise). */
  otherwise?: OtherwiseOptions;
}

/**
 * The unordered, repeatable choice content of dgm:layoutNode / dgm:forEach /
 * dgm:if / dgm:else, as single-key objects so both order and repetition
 * round-trip (same pattern as paragraph children in the format packages).
 */
export type LayoutNodeChild =
  | { algorithm: AlgorithmOptions }
  | { shape: LayoutShapeOptions }
  | { presentationOf: DiagramIterationOptions }
  | { constraints: ConstraintOptions[] }
  | { rules: LayoutRuleOptions[] }
  | { variables: VariableListOptions }
  | { forEach: ForEachOptions }
  | { layoutNode: LayoutNodeOptions }
  | { choose: ChooseOptions };

/** dgm:layoutNode (CT_LayoutNode) — one node of the layout tree. */
export interface LayoutNodeOptions {
  /** Unique name referenced by constraints and moveWith (dgm:`@name`). */
  name?: string;
  /** Formatting slot from styleDef/colorsDef (dgm:`@styleLbl`). */
  styleLabel?: string;
  /** Child ordering: bottom-first "b" or top-first "t" (dgm:`@chOrder`). */
  childOrder?: ChildOrder;
  /** Move with another named layout node (dgm:`@moveWith`). */
  moveWith?: string;
  children?: LayoutNodeChild[];
  /** Raw a:extLst inner XML — verbatim round-trip. */
  ext?: string;
}

/** dgm:forEach (CT_ForEach) — apply children to each selected point. */
export interface ForEachOptions extends DiagramIterationOptions {
  name?: string;
  /** Reference to a named layout node instead of inline children (dgm:`@ref`). */
  reference?: string;
  children?: LayoutNodeChild[];
}

/** dgm:layoutDef (CT_DiagramDefinition) — the layoutDef part root. */
export interface LayoutDefinitionOptions {
  /** Layout identity URI (dgm:layoutDef `@uniqueId`). */
  uniqueId?: string;
  /** Minimum Office version (dgm:layoutDef `@minVer`). */
  minVer?: string;
  /** Default quick-style label (dgm:layoutDef `@defStyle`). */
  defaultStyle?: string;
  /** Localized titles (dgm:title*). */
  titles?: DiagramNameOptions[];
  /** Localized descriptions (dgm:desc*). */
  descriptions?: DiagramDescriptionOptions[];
  /** Gallery categories (dgm:catLst). */
  categories?: DiagramCategoryOptions[];
  /** Sample data shown in the picker (dgm:sampData). */
  sampleData?: SampleDataOptions;
  /** Sample data styled for the quick style preview (dgm:styleData). */
  styleData?: SampleDataOptions;
  /** Sample data colored for the colors preview (dgm:clrData). */
  colorData?: SampleDataOptions;
  /** The layout tree (dgm:layoutNode, required). */
  layoutNode: LayoutNodeOptions;
  /** Raw a:extLst inner XML — verbatim round-trip. */
  ext?: string;
}

// ── Stringify ──

function emitParameterValue(value: string | number | boolean): string {
  return typeof value === "boolean" ? (value ? "1" : "0") : String(value);
}

function stringifyAlgorithm(o: AlgorithmOptions): string {
  const params = (o.parameters ?? [])
    .map((p) => attrs({ type: p.type, val: emitParameterValue(p.value) }))
    .map((a) => `<dgm:param${a}/>`)
    .join("");
  return `<dgm:alg${attrs({ type: o.type, rev: o.revision })}>${params}</dgm:alg>`;
}

function stringifyShape(o: LayoutShapeOptions): string {
  const adjustments = (o.adjustments ?? [])
    .map((a) => `<dgm:adj${attrs({ idx: a.idx, val: a.val })}/>`)
    .join("");
  const adjLst = o.adjustments ? `<dgm:adjLst>${adjustments}</dgm:adjLst>` : "";
  const attrStr = attrs({
    rot: o.rotation,
    type: o.type,
    "r:blip": o.blip,
    zOrderOff: o.zOrderOffset,
    hideGeom: booleanAttr(o.hideGeometry),
    lkTxEntry: booleanAttr(o.lockTextEntry),
    blipPhldr: booleanAttr(o.blipPlaceholder),
  });
  return `<dgm:shape${attrStr}>${adjLst}</dgm:shape>`;
}

function stringifyConstraints(list: readonly ConstraintOptions[]): string {
  const items = list
    .map((c) => {
      const base = attrs({
        type: c.type,
        for: c.for,
        forName: c.forName,
        ptType: c.pointType,
        refType: c.referenceType,
        refFor: c.referenceFor,
        refForName: c.referenceForName,
        refPtType: c.referencePointType,
        op: c.operation,
        val: c.value,
      });
      const fact = c.factor !== undefined ? ` fact="${c.factor}"` : "";
      return `<dgm:constr${base}${fact}/>`;
    })
    .join("");
  return `<dgm:constrLst>${items}</dgm:constrLst>`;
}

function stringifyRules(list: readonly LayoutRuleOptions[]): string {
  const items = list
    .map((r) => {
      const base = attrs({
        type: r.type,
        for: r.for,
        forName: r.forName,
        ptType: r.pointType,
        val: r.value,
        max: r.maximum,
      });
      const fact = r.factor !== undefined ? ` fact="${r.factor}"` : "";
      return `<dgm:rule${base}${fact}/>`;
    })
    .join("");
  return `<dgm:ruleLst>${items}</dgm:ruleLst>`;
}

function stringifyVariables(o: VariableListOptions): string {
  const emit = (tag: string, value: string | number | boolean | undefined) =>
    value === undefined ? "" : `<dgm:${tag}${attrs({ val: value })}/>`;
  return (
    "<dgm:varLst>" +
    emit("orgChart", booleanAttr(o.organizationChart)) +
    emit("chMax", o.childMaximum) +
    emit("chPref", o.childPreferred) +
    emit("bulletEnabled", booleanAttr(o.bulletEnabled)) +
    emit("dir", o.direction) +
    emit("hierBranch", o.hierarchyBranchStyle) +
    emit("animOne", o.animateOne) +
    emit("animLvl", o.animateLevel) +
    emit("resizeHandles", o.resizeHandles) +
    "</dgm:varLst>"
  );
}

/** Boolean options emit "1"/"0" when set; undefined stays absent. */
function booleanAttr(value: boolean | undefined): string | number | undefined {
  return value === undefined ? undefined : value ? "1" : "0";
}

function selfClose(tag: string, record: Record<string, string | number | undefined>): string {
  return `<${tag}${attrs(record)}/>`;
}

function stringifyChildren(children: readonly LayoutNodeChild[] | undefined): string {
  let out = "";
  for (const child of children ?? []) {
    if ("algorithm" in child) out += stringifyAlgorithm(child.algorithm);
    else if ("shape" in child) out += stringifyShape(child.shape);
    else if ("presentationOf" in child)
      out += selfClose("dgm:presOf", {
        axis: child.presentationOf.axis,
        ptType: child.presentationOf.pointType,
        hideLastTrans: child.presentationOf.hideLastTransition,
        st: child.presentationOf.start,
        cnt: child.presentationOf.count,
        step: child.presentationOf.step,
      });
    else if ("constraints" in child) out += stringifyConstraints(child.constraints);
    else if ("rules" in child) out += stringifyRules(child.rules);
    else if ("variables" in child) out += stringifyVariables(child.variables);
    else if ("forEach" in child) out += stringifyForEach(child.forEach);
    else if ("layoutNode" in child) out += stringifyLayoutNode(child.layoutNode);
    else if ("choose" in child) out += stringifyChoose(child.choose);
  }
  return out;
}

function stringifyLayoutNode(o: LayoutNodeOptions): string {
  const body = stringifyChildren(o.children) + (o.ext ? `<a:extLst>${o.ext}</a:extLst>` : "");
  return `<dgm:layoutNode${attrs({
    name: o.name,
    styleLbl: o.styleLabel,
    chOrder: o.childOrder,
    moveWith: o.moveWith,
  })}>${body}</dgm:layoutNode>`;
}

function stringifyForEach(o: ForEachOptions): string {
  const attrStr = attrs({ name: o.name, ref: o.reference, ...iterationRecord(o) });
  return `<dgm:forEach${attrStr}>${stringifyChildren(o.children)}</dgm:forEach>`;
}

function stringifyWhen(o: WhenOptions): string {
  const attrStr = attrs({
    name: o.name,
    ...iterationRecord(o),
    func: o.function,
    arg: o.argument,
    op: o.operator,
    val: o.value,
  });
  return `<dgm:if${attrStr}>${stringifyChildren(o.children)}</dgm:if>`;
}

function stringifyOtherwise(o: OtherwiseOptions): string {
  return `<dgm:else${attrs({ name: o.name })}>${stringifyChildren(o.children)}</dgm:else>`;
}

function stringifyChoose(o: ChooseOptions): string {
  const body =
    o.conditions.map(stringifyWhen).join("") + (o.otherwise ? stringifyOtherwise(o.otherwise) : "");
  return `<dgm:choose${attrs({ name: o.name })}>${body}</dgm:choose>`;
}

function iterationRecord(o: DiagramIterationOptions): Record<string, string | undefined> {
  return {
    axis: o.axis,
    ptType: o.pointType,
    hideLastTrans: o.hideLastTransition,
    st: o.start,
    cnt: o.count,
    step: o.step,
  };
}

function stringifyLocalizedString(
  tag: string,
  o: DiagramNameOptions | DiagramDescriptionOptions,
): string {
  return `<dgm:${tag}${attrs({ lang: o.lang, val: o.val })}/>`;
}

function stringifySampleData(tag: string, o: SampleDataOptions | undefined): string {
  if (!o) return "";
  const body = o.dataModel ? stringifyDataModelOptions(o.dataModel) : "";
  return `<dgm:${tag}${attrs({ useDef: booleanAttr(o.useDefault) })}>${body}</dgm:${tag}>`;
}

/**
 * Serialize a layout definition to the dgm:layoutDef element (no XML
 * declaration, no namespace declarations — the part wrapper adds both).
 */
export function stringifyLayoutDefinition(o: LayoutDefinitionOptions): string {
  const body =
    (o.titles ?? []).map((t) => stringifyLocalizedString("title", t)).join("") +
    (o.descriptions ?? []).map((d) => stringifyLocalizedString("desc", d)).join("") +
    (o.categories?.length
      ? `<dgm:catLst>${o.categories
          .map((c) => selfClose("dgm:cat", { type: c.type, pri: c.pri }))
          .join("")}</dgm:catLst>`
      : "") +
    stringifySampleData("sampData", o.sampleData) +
    stringifySampleData("styleData", o.styleData) +
    stringifySampleData("clrData", o.colorData) +
    stringifyLayoutNode(o.layoutNode) +
    (o.ext ? `<a:extLst>${o.ext}</a:extLst>` : "");
  return `<dgm:layoutDef${attrs({
    uniqueId: o.uniqueId,
    minVer: o.minVer,
    defStyle: o.defaultStyle,
  })}>${body}</dgm:layoutDef>`;
}

const DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/**
 * Inject the part-level dgm/a namespace declarations into the root element of
 * a serialized layoutDef/styleDef/colorsDef body.
 */
export function withDiagramNamespaces(xml: string): string {
  return xml.replace(
    /^<dgm:(layoutDef|styleDef|colorsDef)/,
    `<dgm:$1 xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}"`,
  );
}

// ── Parse ──

function readIteration(el: Element): DiagramIterationOptions {
  const result: Partial<DiagramIterationOptions> = {};
  const axis = attr(el, "axis");
  if (axis !== undefined) result.axis = axis;
  const ptType = attr(el, "ptType");
  if (ptType !== undefined) result.pointType = ptType;
  const hideLastTrans = attr(el, "hideLastTrans");
  if (hideLastTrans !== undefined) result.hideLastTransition = hideLastTrans;
  const st = attr(el, "st");
  if (st !== undefined) result.start = st;
  const cnt = attr(el, "cnt");
  if (cnt !== undefined) result.count = cnt;
  const step = attr(el, "step");
  if (step !== undefined) result.step = step;
  return result as DiagramIterationOptions;
}

function parseAlgorithm(el: Element): AlgorithmOptions {
  const result: Partial<AlgorithmOptions> = {
    type: (attr(el, "type") ?? "composite") as AlgorithmType,
  };
  const rev = attrNum(el, "rev");
  if (rev !== undefined) result.revision = rev;
  const parameters: AlgorithmParameterOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "dgm:param") continue;
    const raw = attr(child, "val") ?? "";
    const numeric = Number(raw);
    parameters.push({
      type: attr(child, "type") ?? "",
      value:
        raw === "1" || raw === "0"
          ? raw === "1"
          : Number.isFinite(numeric) && raw.trim() !== ""
            ? numeric
            : raw,
    });
  }
  if (parameters.length) result.parameters = parameters;
  return result as AlgorithmOptions;
}

function parseShape(el: Element): LayoutShapeOptions {
  const result: Partial<LayoutShapeOptions> = {};
  const rot = attrNum(el, "rot");
  if (rot !== undefined) result.rotation = rot;
  const type = attr(el, "type");
  if (type !== undefined && type !== "none") result.type = type;
  const blip = attr(el, "r:blip");
  if (blip) result.blip = blip;
  const zOrderOff = attrNum(el, "zOrderOff");
  if (zOrderOff !== undefined) result.zOrderOffset = zOrderOff;
  const hideGeom = attr(el, "hideGeom");
  if (hideGeom !== undefined) result.hideGeometry = hideGeom === "1" || hideGeom === "true";
  const lkTxEntry = attr(el, "lkTxEntry");
  if (lkTxEntry !== undefined) result.lockTextEntry = lkTxEntry === "1" || lkTxEntry === "true";
  const blipPhldr = attr(el, "blipPhldr");
  if (blipPhldr !== undefined) result.blipPlaceholder = blipPhldr === "1" || blipPhldr === "true";
  const adjLst = findChild(el, "dgm:adjLst");
  if (adjLst) {
    const adjustments: AdjustOptions[] = [];
    for (const adj of adjLst.elements ?? []) {
      if (adj.name !== "dgm:adj") continue;
      const idx = attrNum(adj, "idx");
      const val = attrNum(adj, "val");
      if (idx !== undefined && val !== undefined) adjustments.push({ idx, val });
    }
    if (adjustments.length) result.adjustments = adjustments;
  }
  return result as LayoutShapeOptions;
}

function parseConstraint(el: Element): ConstraintOptions {
  const result: Partial<ConstraintOptions> = {
    type: (attr(el, "type") ?? "none") as ConstraintType,
  };
  const forRel = attr(el, "for");
  if (forRel !== undefined) result.for = forRel as ConstraintRelationship;
  const forName = attr(el, "forName");
  if (forName !== undefined && forName !== "") result.forName = forName;
  const ptType = attr(el, "ptType");
  if (ptType !== undefined && ptType !== "all") result.pointType = ptType;
  const refType = attr(el, "refType");
  if (refType !== undefined && refType !== "none") result.referenceType = refType as ConstraintType;
  const refFor = attr(el, "refFor");
  if (refFor !== undefined && refFor !== "self")
    result.referenceFor = refFor as ConstraintRelationship;
  const refForName = attr(el, "refForName");
  if (refForName !== undefined && refForName !== "") result.referenceForName = refForName;
  const refPtType = attr(el, "refPtType");
  if (refPtType !== undefined && refPtType !== "all") result.referencePointType = refPtType;
  const op = attr(el, "op");
  if (op !== undefined && op !== "none") result.operation = op as ConstraintOperator;
  const val = finiteAttr(el, "val");
  if (val !== undefined) result.value = val;
  const fact = finiteAttr(el, "fact");
  if (fact !== undefined && fact !== 1) result.factor = fact;
  return result as ConstraintOptions;
}

function parseRule(el: Element): LayoutRuleOptions {
  const result: Partial<LayoutRuleOptions> = {
    type: (attr(el, "type") ?? "none") as ConstraintType,
  };
  const forRel = attr(el, "for");
  if (forRel !== undefined && forRel !== "self") result.for = forRel as ConstraintRelationship;
  const forName = attr(el, "forName");
  if (forName !== undefined && forName !== "") result.forName = forName;
  const ptType = attr(el, "ptType");
  if (ptType !== undefined && ptType !== "all") result.pointType = ptType;
  const val = finiteAttr(el, "val");
  if (val !== undefined) result.value = val;
  const fact = finiteAttr(el, "fact");
  if (fact !== undefined) result.factor = fact;
  const max = finiteAttr(el, "max");
  if (max !== undefined) result.maximum = max;
  return result as LayoutRuleOptions;
}

/** Numeric attribute read; the XSD NaN default comes back as undefined. */
function finiteAttr(el: Element, name: string): number | undefined {
  const raw = attr(el, name);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function parseVariables(el: Element): VariableListOptions {
  const result: Partial<VariableListOptions> = {};
  const readBool = (tag: string) => {
    for (const child of el.elements ?? []) {
      if (child.name !== `dgm:${tag}`) continue;
      const v = attr(child, "val") ?? "1";
      return v === "1" || v === "true";
    }
    return undefined;
  };
  const readEnum = (tag: string) => {
    for (const child of el.elements ?? []) {
      if (child.name !== `dgm:${tag}`) continue;
      return attr(child, "val") ?? "";
    }
    return undefined;
  };
  const readNumber = (tag: string) => {
    for (const child of el.elements ?? []) {
      if (child.name !== `dgm:${tag}`) continue;
      const v = attrNum(child, "val");
      return v === -1 ? undefined : v;
    }
    return undefined;
  };
  const orgChart = readBool("orgChart");
  if (orgChart !== undefined) result.organizationChart = orgChart;
  const chMax = readNumber("chMax");
  if (chMax !== undefined) result.childMaximum = chMax;
  const chPref = readNumber("chPref");
  if (chPref !== undefined) result.childPreferred = chPref;
  const bulletEnabled = readBool("bulletEnabled");
  if (bulletEnabled !== undefined) result.bulletEnabled = bulletEnabled;
  const direction = readEnum("dir");
  if (direction) result.direction = direction as DiagramDirection;
  const hierBranch = readEnum("hierBranch");
  if (hierBranch) result.hierarchyBranchStyle = hierBranch as HierBranch;
  const animOne = readEnum("animOne");
  if (animOne) result.animateOne = animOne as AnimateOneByOne;
  const animLvl = readEnum("animLvl");
  if (animLvl) result.animateLevel = animLvl as AnimationLevel;
  const resizeHandles = readEnum("resizeHandles");
  if (resizeHandles) result.resizeHandles = resizeHandles as ResizeHandles;
  return result as VariableListOptions;
}

function parseChildren(el: Element): LayoutNodeChild[] | undefined {
  const children: LayoutNodeChild[] = [];
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "dgm:alg":
        children.push({ algorithm: parseAlgorithm(child) });
        break;
      case "dgm:shape":
        children.push({ shape: parseShape(child) });
        break;
      case "dgm:presOf":
        children.push({ presentationOf: readIteration(child) });
        break;
      case "dgm:constrLst":
        children.push({
          constraints: (child.elements ?? [])
            .filter((c) => c.name === "dgm:constr")
            .map(parseConstraint),
        });
        break;
      case "dgm:ruleLst":
        children.push({
          rules: (child.elements ?? []).filter((c) => c.name === "dgm:rule").map(parseRule),
        });
        break;
      case "dgm:varLst":
        children.push({ variables: parseVariables(child) });
        break;
      case "dgm:forEach":
        children.push({ forEach: parseForEach(child) });
        break;
      case "dgm:layoutNode":
        children.push({ layoutNode: parseLayoutNode(child) });
        break;
      case "dgm:choose":
        children.push({ choose: parseChoose(child) });
        break;
      default:
        break;
    }
  }
  return children.length ? children : undefined;
}

function parseLayoutNode(el: Element): LayoutNodeOptions {
  const result: Partial<LayoutNodeOptions> = {};
  const name = attr(el, "name");
  if (name) result.name = name;
  const styleLbl = attr(el, "styleLbl");
  if (styleLbl) result.styleLabel = styleLbl;
  const chOrder = attr(el, "chOrder");
  if (chOrder !== undefined && chOrder !== "b") result.childOrder = chOrder as ChildOrder;
  const moveWith = attr(el, "moveWith");
  if (moveWith) result.moveWith = moveWith;
  const children = parseChildren(el);
  if (children) result.children = children;
  const extLst = findChild(el, "a:extLst");
  if (extLst) result.ext = stringifyInnerXml(extLst);
  return result as LayoutNodeOptions;
}

function parseForEach(el: Element): ForEachOptions {
  const result: Partial<ForEachOptions> = readIteration(el);
  const name = attr(el, "name");
  if (name) result.name = name;
  const ref = attr(el, "ref");
  if (ref) result.reference = ref;
  const children = parseChildren(el);
  if (children) result.children = children;
  return result as ForEachOptions;
}

function parseWhen(el: Element): WhenOptions {
  const result: Partial<WhenOptions> = {
    ...readIteration(el),
    function: (attr(el, "func") ?? "cnt") as ConditionFunction,
    operator: (attr(el, "op") ?? "equ") as ConditionOperator,
    value: attr(el, "val") ?? "",
  };
  const name = attr(el, "name");
  if (name) result.name = name;
  const arg = attr(el, "arg");
  if (arg !== undefined && arg !== "none") result.argument = arg;
  const children = parseChildren(el);
  if (children) result.children = children;
  return result as WhenOptions;
}

function parseChoose(el: Element): ChooseOptions {
  const conditions: WhenOptions[] = [];
  let otherwise: OtherwiseOptions | undefined;
  for (const child of el.elements ?? []) {
    if (child.name === "dgm:if") conditions.push(parseWhen(child));
    else if (child.name === "dgm:else") otherwise = parseOtherwise(child);
  }
  const result: Partial<ChooseOptions> = { conditions };
  const name = attr(el, "name");
  if (name) result.name = name;
  if (otherwise) result.otherwise = otherwise;
  return result as ChooseOptions;
}

function parseOtherwise(el: Element): OtherwiseOptions {
  const result: Partial<OtherwiseOptions> = {};
  const name = attr(el, "name");
  if (name) result.name = name;
  const children = parseChildren(el);
  if (children) result.children = children;
  return result as OtherwiseOptions;
}

function parseLocalizedStrings(
  el: Element,
  tag: string,
): DiagramNameOptions[] | DiagramDescriptionOptions[] | undefined {
  const out: (DiagramNameOptions | DiagramDescriptionOptions)[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== `dgm:${tag}`) continue;
    const val = attr(child, "val") ?? "";
    const lang = attr(child, "lang");
    out.push(lang ? { lang, val } : { val });
  }
  return out.length ? (out as DiagramNameOptions[]) : undefined;
}

function parseSampleData(el: Element | undefined): SampleDataOptions | undefined {
  if (!el) return undefined;
  const result: Partial<SampleDataOptions> = {};
  const useDef = attr(el, "useDef");
  if (useDef !== undefined) result.useDefault = useDef === "1" || useDef === "true";
  const dataModel = parseDataModelOptions(findChild(el, "dgm:dataModel"));
  if (dataModel) result.dataModel = dataModel;
  return Object.keys(result).length ? (result as SampleDataOptions) : {};
}

/** Parse a dgm:layoutDef element into options. */
export function parseLayoutDefinition(el: Element): LayoutDefinitionOptions {
  const root = el.name === "dgm:layoutDef" ? el : (findChild(el, "dgm:layoutDef") ?? el);
  const result: Partial<LayoutDefinitionOptions> = {};
  const uniqueId = attr(root, "uniqueId");
  if (uniqueId) result.uniqueId = uniqueId;
  const minVer = attr(root, "minVer");
  if (minVer) result.minVer = minVer;
  const defStyle = attr(root, "defStyle");
  if (defStyle) result.defaultStyle = defStyle;
  const titles = parseLocalizedStrings(root, "title");
  if (titles) result.titles = titles;
  const descriptions = parseLocalizedStrings(root, "desc");
  if (descriptions) result.descriptions = descriptions;
  const catLst = findChild(root, "dgm:catLst");
  if (catLst) {
    const categories = (catLst.elements ?? [])
      .filter((c) => c.name === "dgm:cat")
      .map((c) => ({ type: attr(c, "type") ?? "", pri: attrNum(c, "pri") ?? 0 }));
    if (categories.length) result.categories = categories;
  }
  const sampleData = parseSampleData(findChild(root, "dgm:sampData"));
  if (sampleData) result.sampleData = sampleData;
  const styleData = parseSampleData(findChild(root, "dgm:styleData"));
  if (styleData) result.styleData = styleData;
  const colorData = parseSampleData(findChild(root, "dgm:clrData"));
  if (colorData) result.colorData = colorData;
  const layoutNode = findChild(root, "dgm:layoutNode");
  if (layoutNode) result.layoutNode = parseLayoutNode(layoutNode);
  const extLst = findChild(root, "a:extLst");
  if (extLst) result.ext = stringifyInnerXml(extLst);
  return result as LayoutDefinitionOptions;
}

/** layoutDef part descriptor — stringify emits the element, parse reads it. */
export const layoutDefDesc: CustomDescriptor<LayoutDefinitionOptions> = {
  kind: "custom",
  stringify: (opts: LayoutDefinitionOptions) => stringifyLayoutDefinition(opts),
  parse: (el: Element) => parseLayoutDefinition(el),
};

/** Full layoutDef part body: XML declaration + namespaces + the element. */
export function stringifyLayoutDefinitionPart(o: LayoutDefinitionOptions): string {
  return OOXML_XML_DECLARATION + withDiagramNamespaces(stringifyLayoutDefinition(o));
}
