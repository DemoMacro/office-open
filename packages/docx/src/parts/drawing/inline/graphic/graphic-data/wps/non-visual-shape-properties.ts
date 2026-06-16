export interface NonVisualShapePropertiesOptions {
  /** Drawing id (wps:cNvPr @id). */
  id?: number;
  /** Drawing name (wps:cNvPr @name). */
  name?: string;
  /** Alt-text description (wps:cNvPr @descr). */
  description?: string;
  /** Alt-text title (wps:cNvPr @title). */
  title?: string;
  /** Shape type marker (wps:cNvSpPr @txBox) — "1" marks a text box. */
  txBox?: string;
  /**
   * Connector shape marker — when true the shape emits wps:cNvCnPr (connector
   * non-visual properties) instead of wps:cNvSpPr. Per the CT_WordprocessingShape
   * choice these are mutually exclusive.
   */
  connector?: boolean;
}

export const createNonVisualShapeProperties = (
  options: NonVisualShapePropertiesOptions = { txBox: "1" },
): string => `<wps:cNvSpPr txBox="${options.txBox ?? "1"}"/>`;
