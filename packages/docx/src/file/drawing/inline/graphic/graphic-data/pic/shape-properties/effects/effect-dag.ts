/**
 * Effect container (effectDag) for DrawingML shapes.
 *
 * Provides CT_EffectContainer — a directed acyclic graph (DAG) of effects
 * supporting 28 effect types including alpha/color operations, nested containers,
 * and all effects from CT_EffectList.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_EffectContainer, EG_Effect
 *
 * @module
 */

export { createEffectDag } from "@office-open/core/drawingml";
export type { EffectDagOptions, EffectContainerType } from "@office-open/core/drawingml";
