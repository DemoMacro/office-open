import type { IXmlableObject } from "@file/xml-components";
import { XmlComponent } from "@file/xml-components";
import { xsdVerticalMergeRev } from "@office-open/core";

import type { ChangedAttributesProperties } from "../track-revision";

/**
 * Vertical merge revision types.
 */
export const VerticalMergeRevisionType = {
  /**
   * Cell that is merged with upper one.
   */
  CONTINUE: "continue",
  /**
   * Cell that is starting the vertical merge.
   */
  RESTART: "restart",
} as const;

export type ICellMergeAttributes = ChangedAttributesProperties & {
  readonly verticalMerge?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
  readonly verticalMergeOriginal?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
};

export function buildCellMergeObj(options: ICellMergeAttributes): IXmlableObject {
  const attrs: Record<string, string | number> = {
    "w:author": options.author,
    "w:date": options.date,
    "w:id": options.id,
  };
  if (options.verticalMerge !== undefined) {
    attrs["w:vMerge"] = xsdVerticalMergeRev.to(options.verticalMerge);
  }
  if (options.verticalMergeOriginal !== undefined) {
    attrs["w:vMergeOrig"] = xsdVerticalMergeRev.to(options.verticalMergeOriginal);
  }
  return { "w:cellMerge": { _attr: attrs } };
}

export class CellMerge extends XmlComponent {
  public constructor(options: ICellMergeAttributes) {
    super("w:cellMerge");

    const attrs: Record<string, string | number> = {
      "w:author": options.author,
      "w:date": options.date,
      "w:id": options.id,
    };
    if (options.verticalMerge !== undefined) {
      attrs["w:vMerge"] = xsdVerticalMergeRev.to(options.verticalMerge);
    }
    if (options.verticalMergeOriginal !== undefined) {
      attrs["w:vMergeOrig"] = xsdVerticalMergeRev.to(options.verticalMergeOriginal);
    }
    this.root.push({ _attr: attrs });
  }
}
