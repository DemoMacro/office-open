import { BuilderElement } from "@file/xml-components";

/**
 * a:prstGeom — Preset geometry for shapes (rect, ellipse, triangle, etc.).
 */
export class PresetGeometry extends BuilderElement<{ readonly prst: string }> {
    public constructor(prst: string, avLst?: BuilderElement) {
        super({
            name: "a:prstGeom",
            attributes: { prst: { key: "prst", value: prst } },
            children: avLst ? [avLst] : [new BuilderElement({ name: "a:avLst" })],
        });
    }
}
