import type { HyperlinkOptions } from "@file/drawing/doc-properties/doc-properties";
import { createWpsShape } from "@file/drawing/inline/graphic/graphic-data/wps/wps-shape";
import type {
  ChartMediaData,
  IExtendedMediaData,
  IMediaData,
  MediaDataTransformation,
  SmartArtMediaData,
  WpgMediaData,
} from "@file/media";
import { XmlComponent } from "@file/xml-components";
import type { FillOptions } from "@office-open/core/drawingml";

import { Pic } from "./pic";
import type { BlipEffectsOptions } from "./pic/blip/blip-effects";
import type { TileOptions } from "./pic/blip/tile";
import type { EffectListOptions } from "./pic/shape-properties/effects/effect-list";
import type { OutlineOptions } from "./pic/shape-properties/outline/outline";
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
  // Private pic: Pic;

  public constructor({
    mediaData,
    transform,
    outline,
    fill,
    effects,
    blipEffects,
    tile,
    hyperlink,
  }: {
    mediaData: IExtendedMediaData;
    transform: MediaDataTransformation;
    outline?: OutlineOptions;
    fill?: FillOptions;
    effects?: EffectListOptions;
    blipEffects?: BlipEffectsOptions;
    tile?: TileOptions;
    hyperlink?: HyperlinkOptions;
  }) {
    super("a:graphicData");

    if (mediaData.type === "wps") {
      this.root.push({
        _attr: { uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" },
      });
      const wps = createWpsShape({
        ...mediaData.data,
        outline,
        fill,
        transformation: transform,
      });
      this.root.push(wps);
    } else if (mediaData.type === "wpg") {
      this.root.push({
        _attr: { uri: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" },
      });
      const md = mediaData as WpgMediaData;
      const children = md.children.map((child) => {
        if (child.type === "wps") {
          return createWpsShape({
            ...child.data,
            outline: child.outline,
            fill: child.fill,
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
        fill: md.fill,
        effects: md.effects,
      });
      this.root.push(wpg);
    } else if (mediaData.type === "chart") {
      this.root.push({ _attr: { uri: "http://schemas.openxmlformats.org/drawingml/2006/chart" } });
      const md = mediaData as ChartMediaData;
      const chartRef = new (class extends XmlComponent {
        public constructor() {
          super("c:chart");
        }
      })();
      chartRef["root"].push({
        _attr: {
          "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "r:id": `{chart:${md.chartKey}}`,
        },
      });
      this.root.push(chartRef);
    } else if (mediaData.type === "smartart") {
      this.root.push({
        _attr: { uri: "http://schemas.openxmlformats.org/drawingml/2006/diagram" },
      });
      const md = mediaData as SmartArtMediaData;
      const relIds = new (class extends XmlComponent {
        public constructor() {
          super("dgm:relIds");
        }
      })();
      relIds["root"].push({
        _attr: {
          "xmlns:dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "r:cs": `{smartart-cs:${md.smartArtKey}}`,
          "r:dm": `{smartart:${md.smartArtKey}}`,
          "r:lo": `{smartart-lo:${md.smartArtKey}}`,
          "r:qs": `{smartart-qs:${md.smartArtKey}}`,
        },
      });
      this.root.push(relIds);
    } else {
      this.root.push({
        _attr: { uri: "http://schemas.openxmlformats.org/drawingml/2006/picture" },
      });
      const md = mediaData as IMediaData;
      const pic = new Pic({
        blipEffects,
        effects,
        hyperlink,
        mediaData: md,
        outline,
        fill,
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
