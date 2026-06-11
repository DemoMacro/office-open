/**
 * Master slide child coercion.
 *
 * @module
 */
import type { MasterChild } from "@shared/file";

/**
 * Build master slide children XML from JSON options.
 *
 * Currently only supports empty children (master slides with no custom shapes).
 * Custom shapes on master slides should use the descriptor pipeline directly.
 */
export function buildMasterChildrenXml(children?: MasterChild[]): string {
  if (!children || children.length === 0) return "";
  // Master slide custom children with shape class instances are no longer supported.
  // Use the descriptor pipeline (slideMasterDesc) for master slides with custom shapes.
  return "";
}
