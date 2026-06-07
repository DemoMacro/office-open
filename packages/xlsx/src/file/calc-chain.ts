/**
 * Calculation Chain types.
 *
 * Reference: OOXML transitional, sml.xsd, CT_CalcChain / CT_CalcCell
 *
 * @module
 */

export interface CalcCell {
  /** Cell reference, e.g. "A1" */
  reference: string;
  /** Sheet index (1-based) */
  sheetIndex: number;
  /** Array formula */
  array?: boolean;
}
