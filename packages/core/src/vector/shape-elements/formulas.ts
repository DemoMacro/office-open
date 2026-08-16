/**
 * v:formulas element — CT_Formulas (list of v:f equation definitions).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Formulas / CT_F.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

/** v:formulas options (CT_Formulas). */
export interface VmlFormulasOptions {
  /** Equation strings, each becoming a `<v:f eqn="…"/>` child. */
  equations?: string[];
}

/** Serialize v:formulas. */
export function stringifyVmlFormulas(opts: VmlFormulasOptions): string {
  const children = (opts.equations ?? []).map((eqn) => `<v:f eqn="${escapeXml(eqn)}"/>`).join("");
  return `<v:formulas>${children}</v:formulas>`;
}

/** Parse a v:formulas element. */
export function parseVmlFormulas(el: XmlElement): VmlFormulasOptions {
  const equations: string[] = [];
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "v:f") {
      const eqn = child.attributes?.eqn;
      if (eqn !== undefined) equations.push(String(eqn));
    }
  }
  return equations.length > 0 ? { equations } : {};
}
