import type { HyperlinkOptions } from "@file/drawing/doc-properties/doc-properties";
import { createWpsShape } from "@file/drawing/inline/graphic/graphic-data/wps/wps-shape";
import type {
    ChartMediaData,
    IExtendedMediaData,
    IMediaData,
    IMediaDataTransformation,
    SmartArtMediaData,
    WpgMediaData,
} from "@file/media";
import { NextAttributeComponent, XmlComponent } from "@file/xml-components";

import { GraphicDataAttributes } from "./graphic-data-attribute";
import { Pic } from "./pic";
import type { BlipEffectsOptions } from "./pic/blip/blip-effects";
import type { TileOptions } from "./pic/blip/tile";
import type { EffectListOptions } from "./pic/shape-properties/effects/effect-list";
import type { OutlineOptions } from "./pic/shape-properties/outline/outline";
import type { SolidFillOptions } from "./pic/shape-properties/outline/solid-fill";
import { createWpgGroup } from "./wpg/wpg-group";

/**
 * Represents graphical data within a DrawingML graphic element.
 *
 * GraphicData contains the actual picture, chart, or other graphical
 * content referenced by a graphic element. It uses a URI to identify
 * the type of content being stored.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GraphicalObjectData">
 *   <xsd:sequence>
 *     <xsd:any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="uri" type="xsd:token" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const graphicData = new GraphicData({
 *   mediaData: imageData,
 *   transform: transformation,
 *   outline: { color: "000000", width: 9525 }
 * });
 * ```
 */
export class GraphicData extends XmlComponent {
    // Private readonly pic: Pic;

    public constructor({
        mediaData,
        transform,
        outline,
        solidFill,
        effects,
        blipEffects,
        tile,
        hyperlink,
    }: {
        readonly mediaData: IExtendedMediaData;
        readonly transform: IMediaDataTransformation;
        readonly outline?: OutlineOptions;
        readonly solidFill?: SolidFillOptions;
        readonly effects?: EffectListOptions;
        readonly blipEffects?: BlipEffectsOptions;
        readonly tile?: TileOptions;
        readonly hyperlink?: HyperlinkOptions;
    }) {
        super("a:graphicData");

        if (mediaData.type === "wps") {
            this.root.push(
                new GraphicDataAttributes({
                    uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                }),
            );
            const wps = createWpsShape({
                ...mediaData.data,
                outline,
                solidFill,
                transformation: transform,
            });
            this.root.push(wps);
        } else if (mediaData.type === "wpg") {
            this.root.push(
                new GraphicDataAttributes({
                    uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                }),
            );
            const md = mediaData as WpgMediaData;
            const children = md.children.map((child) => {
                if (child.type === "wps") {
                    return createWpsShape({
                        ...child.data,
                        outline: child.outline,
                        solidFill: child.solidFill,
                        transformation: child.transformation,
                    });
                } else {
                    return new Pic({
                        mediaData: child,
                        outline: child.outline,
                        transform: child.transformation,
                    });
                }
            });
            // Const wps = new WpsShape({ ...mediaData.data, transformation: transform, outline, solidFill });
            const wpg = createWpgGroup({
                children,
                transformation: transform,
                chOff: md.chOff,
                chExt: md.chExt,
                solidFill: md.solidFill,
                effects: md.effects,
            });
            this.root.push(wpg);
        } else if (mediaData.type === "chart") {
            this.root.push(
                new GraphicDataAttributes({
                    uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
                }),
            );
            const md = mediaData as ChartMediaData;
            const chartRef = new (class extends XmlComponent {
                public constructor() {
                    super("c:chart");
                }
            })();
            chartRef["root"].push(
                new NextAttributeComponent({
                    xmlnsC: {
                        key: "xmlns:c",
                        value: "http://schemas.openxmlformats.org/drawingml/2006/chart",
                    },
                    xmlnsR: {
                        key: "xmlns:r",
                        value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    },
                    rId: { key: "r:id", value: `rId{chart:${md.chartKey}}` },
                }),
            );
            this.root.push(chartRef);
        } else if (mediaData.type === "smartart") {
            this.root.push(
                new GraphicDataAttributes({
                    uri: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                }),
            );
            const md = mediaData as SmartArtMediaData;
            const relIds = new (class extends XmlComponent {
                public constructor() {
                    super("dgm:relIds");
                }
            })();
            relIds["root"].push(
                new NextAttributeComponent({
                    xmlnsDgm: {
                        key: "xmlns:dgm",
                        value: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                    },
                    xmlnsR: {
                        key: "xmlns:r",
                        value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    },
                    rCs: { key: "r:cs", value: `rId{smartart-cs:${md.smartArtKey}}` },
                    rDm: { key: "r:dm", value: `rId{smartart:${md.smartArtKey}}` },
                    rLo: { key: "r:lo", value: `rId{smartart-lo:${md.smartArtKey}}` },
                    rQs: { key: "r:qs", value: `rId{smartart-qs:${md.smartArtKey}}` },
                }),
            );
            this.root.push(relIds);
        } else {
            this.root.push(
                new GraphicDataAttributes({
                    uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
                }),
            );
            const md = mediaData as IMediaData;
            const pic = new Pic({
                blipEffects,
                effects,
                hyperlink,
                mediaData: md,
                outline,
                solidFill,
                tile,
                transform,
            });
            this.root.push(pic);
        }

        // If (mediaData.type !== "wps") {
        //     Const pic = new Pic({ mediaData, transform, outline });
        //     This.root.push(pic);
        // } else {
        //     Const wps = new WpsShape({ ...mediaData.data, transformation: transform, outline, solidFill });
        //     This.root.push(wps);
        // }
    }
}
