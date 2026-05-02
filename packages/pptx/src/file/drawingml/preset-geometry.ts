import { BuilderElement, type XmlComponent } from "@file/xml-components";

/**
 * a:prstGeom — Preset geometry for shapes (rect, ellipse, triangle, etc.).
 */
export class PresetGeometry extends BuilderElement<{ readonly prst: string }> {
    public constructor(prst: string, avLst?: XmlComponent) {
        super({
            name: "a:prstGeom",
            attributes: { prst: { key: "prst", value: prst } },
            children: avLst ? [avLst] : undefined,
        });
    }
}
