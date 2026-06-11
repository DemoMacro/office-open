/**
 * CT_DataModel XML stringifier — assembles complete SmartArt data model XML.
 *
 * @module
 */

/**
 * Build the complete dgm:dataModel XML from pre-serialized point and connection strings.
 */
export function stringifyDataModel(
  points: readonly string[],
  connections: readonly string[],
): string {
  return [
    '<dgm:dataModel xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">',
    `<dgm:ptLst>${points.join("")}</dgm:ptLst>`,
    `<dgm:cxnLst>${connections.join("")}</dgm:cxnLst>`,
    "<dgm:bg/>",
    "<dgm:whole/>",
    "</dgm:dataModel>",
  ].join("");
}
