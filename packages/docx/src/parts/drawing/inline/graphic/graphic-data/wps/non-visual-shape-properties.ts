export interface NonVisualShapePropertiesOptions {
  txBox: string;
}

export const createNonVisualShapeProperties = (
  options: NonVisualShapePropertiesOptions = { txBox: "1" },
): string => `<wps:cNvSpPr txBox="${options.txBox}"/>`;
