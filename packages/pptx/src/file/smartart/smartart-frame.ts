import { Transform2D } from "@file/drawingml/transform-2d";
import { BuilderElement, NextAttributeComponent, type IContext, type IXmlableObject, XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

import type { SmartArtCollection } from "./smartart-collection";
import { createDataModel, type ITreeNode } from "./tree-to-model";

let nextSmartArtFrameId = 1024;

export interface ISmartArtFrameOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly nodes: readonly ITreeNode[];
    readonly name?: string;
    /** Layout ID (e.g. "default", "process1", "hierarchy1") */
    readonly layout?: string;
    /** Quick style ID (e.g. "simple1", "moderate1") */
    readonly style?: string;
    /** Color transform ID (e.g. "accent1_2", "colorful1") */
    readonly color?: string;
}

/**
 * p:graphicFrame — Slide-level graphic frame wrapping a SmartArt diagram.
 */
export class SmartArtFrame extends XmlComponent {
    private readonly smartArtKey: string;
    private readonly dataModel: ReturnType<typeof createDataModel>;
    private readonly layoutId: string;
    private readonly styleId: string;
    private readonly colorId: string;

    public constructor(options: ISmartArtFrameOptions) {
        super("p:graphicFrame");

        this.smartArtKey = `smartart_${nextSmartArtFrameId++}`;
        this.layoutId = options.layout ?? "default";
        this.styleId = options.style ?? "simple1";
        this.colorId = options.color ?? "accent1_2";
        this.dataModel = createDataModel(options.nodes, this.layoutId, this.styleId, this.colorId);

        const id = nextSmartArtFrameId++;
        const name = options.name ?? `Diagram ${id}`;

        this.root.push(new GraphicFrameNonVisual(id, name));
        this.root.push(
            new Transform2D(
                {
                    x: pixelsToEmus(options.x ?? 0),
                    y: pixelsToEmus(options.y ?? 0),
                    width: pixelsToEmus(options.width ?? 0),
                    height: pixelsToEmus(options.height ?? 0),
                },
                "p",
            ),
        );
        this.root.push(new SmartArtGraphic(this.smartArtKey));
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const file = context.fileData as { SmartArts: SmartArtCollection };
        if (file?.SmartArts) {
            file.SmartArts.addSmartArt(this.smartArtKey, {
                key: this.smartArtKey,
                dataModel: this.dataModel,
                layout: this.layoutId,
                style: this.styleId,
                color: this.colorId,
            });
        }
        return super.prepForXml(context);
    }
}

class GraphicFrameNonVisual extends XmlComponent {
    public constructor(id: number, name: string) {
        super("p:nvGraphicFramePr");
        this.root.push(
            new BuilderElement({
                name: "p:cNvPr",
                attributes: {
                    id: { key: "id", value: id },
                    name: { key: "name", value: name },
                },
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "p:cNvGraphicFramePr",
                children: [
                    new BuilderElement({
                        name: "a:graphicFrameLocks",
                        attributes: { noGrp: { key: "noGrp", value: 1 } },
                    }),
                ],
            }),
        );
        this.root.push(new BuilderElement({ name: "p:nvPr" }));
    }
}

class SmartArtGraphic extends XmlComponent {
    public constructor(smartArtKey: string) {
        super("a:graphic");
        this.root.push(new SmartArtGraphicData(smartArtKey));
    }
}

class SmartArtGraphicData extends XmlComponent {
    public constructor(smartArtKey: string) {
        super("a:graphicData");
        this.root.push(
            new NextAttributeComponent({
                uri: {
                    key: "uri",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
            }),
        );
        this.root.push(new DiagramRelIds(smartArtKey));
    }
}

class DiagramRelIds extends XmlComponent {
    public constructor(smartArtKey: string) {
        super("dgm:relIds");
        this.root.push(
            new NextAttributeComponent({
                xmlnsDgm: {
                    key: "xmlns:dgm",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                xmlnsR: {
                    key: "xmlns:r",
                    value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                },
                rDm: { key: "r:dm", value: `{smartart:${smartArtKey}}` },
                rLo: { key: "r:lo", value: `{smartart-lo:${smartArtKey}}` },
                rQs: { key: "r:qs", value: `{smartart-qs:${smartArtKey}}` },
                rCs: { key: "r:cs", value: `{smartart-cs:${smartArtKey}}` },
            }),
        );
    }
}
