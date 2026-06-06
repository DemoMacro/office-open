import { BuilderElement, type XmlComponent } from "../../xml-components";

export interface LockedCanvasOptions {
  readonly id?: number;
  readonly name?: string;
}

/**
 * lc:lockedCanvas — A locked canvas group shape.
 * Type: CT_GvmlGroupShape (nvGrpSpPr + grpSpPr + child shapes).
 */
export class LockedCanvas extends BuilderElement {
  public constructor(options: LockedCanvasOptions = {}) {
    const id = options.id ?? 1;
    const name = options.name ?? "Locked Canvas";
    super({
      name: "lc:lockedCanvas",
      attributes: {
        "xmlns:lc": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
      },
      children: [
        {
          name: "lc:nvGrpSpPr",
          children: [{ name: "lc:cNvPr", attributes: { id, name } }, { name: "lc:cNvGrpSpPr" }],
        },
        { name: "lc:grpSpPr" },
      ],
    });
  }
}
