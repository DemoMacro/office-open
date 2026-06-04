/**
 * Web Settings module for WordprocessingML documents.
 *
 * Web settings control how documents are rendered in web browsers,
 * including framesets, divs, encoding, and browser optimization.
 *
 * Reference: http://officeopenxml.com/WPwebSettings.php
 *
 * @module
 */
import {
  BuilderElement,
  XmlComponent,
  numberValObj,
  onOffObj,
  stringValObj,
} from "@file/xml-components";

import type { FramesetOptions } from "../frameset/frameset";
import { Frameset } from "../frameset/frameset";

/**
 * Border options for div elements.
 */
export interface DivBorderOptions {
  readonly top?: { readonly style: string; readonly color?: string; readonly size?: number };
  readonly left?: { readonly style: string; readonly color?: string; readonly size?: number };
  readonly bottom?: { readonly style: string; readonly color?: string; readonly size?: number };
  readonly right?: { readonly style: string; readonly color?: string; readonly size?: number };
}

/**
 * Options for a div element (CT_Div).
 */
export interface DivOptions {
  /** Unique div identifier (required by CT_Div/@id) */
  readonly id: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  /** Mark as HTML blockquote element */
  readonly blockQuote?: boolean;
  /** Mark as HTML body element */
  readonly bodyDiv?: boolean;
  readonly border?: DivBorderOptions;
  readonly children?: readonly DivOptions[];
}

/**
 * Screen size types for web settings target.
 */
export const TargetScreenSize = {
  SIZE_544X376: "544x376",
  SIZE_640X480: "640x480",
  SIZE_720X512: "720x512",
  SIZE_800X600: "800x600",
  SIZE_1024X768: "1024x768",
  SIZE_1152X882: "1152x882",
  SIZE_1152X900: "1152x900",
  SIZE_1280X1024: "1280x1024",
  SIZE_1600X1200: "1600x1200",
  SIZE_1800X1440: "1800x1440",
  SIZE_1920X1200: "1920x1200",
} as const;

/**
 * Options for web settings (CT_WebSettings).
 */
export interface WebSettingsOptions {
  /** Frameset definition for web layout */
  readonly frameset?: FramesetOptions;
  /** Div elements for HTML div formatting */
  readonly divs?: readonly DivOptions[];
  /** Character encoding for web output */
  readonly encoding?: string;
  /** Optimize document rendering for web browser */
  readonly optimizeForBrowser?: boolean;
  /** Rely on VML for graphics display */
  readonly relyOnVML?: boolean;
  /** Allow PNG image format in web output */
  readonly allowPNG?: boolean;
  /** Do not rely on CSS for formatting */
  readonly doNotRelyOnCSS?: boolean;
  /** Do not save as single web file */
  readonly doNotSaveAsSingleFile?: boolean;
  /** Do not organize supporting files in folders */
  readonly doNotOrganizeInFolder?: boolean;
  /** Do not use long file names for supporting files */
  readonly doNotUseLongFileNames?: boolean;
  /** Pixels per inch for web output */
  readonly pixelsPerInch?: number;
  /** Target screen size */
  readonly targetScreenSz?: (typeof TargetScreenSize)[keyof typeof TargetScreenSize] | string;
  /** Save smart tags as XML */
  readonly saveSmartTagsAsXml?: boolean;
}

/**
 * Build div border element.
 */
function buildDivBorder(options: DivBorderOptions): BuilderElement {
  const children: BuilderElement[] = [];
  if (options.top) {
    const attrs: { key: string; value: string | number }[] = [];
    attrs.push({ key: "w:val", value: options.top.style });
    if (options.top.color) attrs.push({ key: "w:color", value: options.top.color });
    if (options.top.size !== undefined) attrs.push({ key: "w:sz", value: options.top.size });
    children.push(new BuilderElement({ name: "w:top", attributes: attrs }));
  }
  if (options.left) {
    const attrs: { key: string; value: string | number }[] = [];
    attrs.push({ key: "w:val", value: options.left.style });
    if (options.left.color) attrs.push({ key: "w:color", value: options.left.color });
    if (options.left.size !== undefined) attrs.push({ key: "w:sz", value: options.left.size });
    children.push(new BuilderElement({ name: "w:left", attributes: attrs }));
  }
  if (options.bottom) {
    const attrs: { key: string; value: string | number }[] = [];
    attrs.push({ key: "w:val", value: options.bottom.style });
    if (options.bottom.color) attrs.push({ key: "w:color", value: options.bottom.color });
    if (options.bottom.size !== undefined) attrs.push({ key: "w:sz", value: options.bottom.size });
    children.push(new BuilderElement({ name: "w:bottom", attributes: attrs }));
  }
  if (options.right) {
    const attrs: { key: string; value: string | number }[] = [];
    attrs.push({ key: "w:val", value: options.right.style });
    if (options.right.color) attrs.push({ key: "w:color", value: options.right.color });
    if (options.right.size !== undefined) attrs.push({ key: "w:sz", value: options.right.size });
    children.push(new BuilderElement({ name: "w:right", attributes: attrs }));
  }
  return new BuilderElement({ name: "w:divBdr", children });
}

/**
 * Build a div element (CT_Div).
 */
function buildDiv(options: DivOptions): BuilderElement {
  const children: (BuilderElement | ReturnType<typeof numberValObj>)[] = [];

  if (options.blockQuote !== undefined) {
    children.push(stringValObj("w:blockQuote", options.blockQuote ? "on" : "off"));
  }
  if (options.bodyDiv !== undefined) {
    children.push(stringValObj("w:bodyDiv", options.bodyDiv ? "on" : "off"));
  }
  children.push(numberValObj("w:marLeft", options.marginLeft));
  children.push(numberValObj("w:marRight", options.marginRight));
  children.push(numberValObj("w:marTop", options.marginTop));
  children.push(numberValObj("w:marBottom", options.marginBottom));
  if (options.border) {
    children.push(buildDivBorder(options.border));
  }
  if (options.children && options.children.length > 0) {
    children.push(
      new BuilderElement({
        name: "w:divsChild",
        children: options.children.map(buildDiv),
      }),
    );
  }

  return new BuilderElement({
    name: "w:div",
    attributes: { id: { key: "w:id", value: options.id } },
    children,
  });
}

/**
 * Build divs container (CT_Divs).
 */
function buildDivs(divs: readonly DivOptions[]): BuilderElement {
  return new BuilderElement({
    name: "w:divs",
    children: divs.map(buildDiv),
  });
}

/**
 * Represents the web settings part (word/webSettings.xml).
 *
 * Web settings control how a document is displayed and formatted
 * when saved as a web page, including framesets, div elements,
 * encoding, and browser optimization settings.
 *
 * Reference: http://officeopenxml.com/WPwebSettings.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_WebSettings">
 *   <xsd:sequence>
 *     <xsd:element name="frameset" type="CT_Frameset" minOccurs="0"/>
 *     <xsd:element name="divs" type="CT_Divs" minOccurs="0"/>
 *     <xsd:element name="encoding" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="optimizeForBrowser" type="CT_OptimizeForBrowser" minOccurs="0"/>
 *     <xsd:element name="relyOnVML" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="allowPNG" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="doNotRelyOnCSS" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="doNotSaveAsSingleFile" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="doNotOrganizeInFolder" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="doNotUseLongFileNames" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="pixelsPerInch" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="targetScreenSz" type="CT_TargetScreenSz" minOccurs="0"/>
 *     <xsd:element name="saveSmartTagsAsXml" type="CT_OnOff" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export class WebSettings extends XmlComponent {
  public constructor(options?: WebSettingsOptions) {
    super("w:webSettings");
    this.root.push({
      _attr: {
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
      },
    });

    if (options === undefined) return;

    if (options.frameset !== undefined) {
      this.root.push(new Frameset(options.frameset));
    }

    if (options.divs !== undefined && options.divs.length > 0) {
      this.root.push(buildDivs(options.divs));
    }

    if (options.encoding !== undefined) {
      this.root.push(stringValObj("w:encoding", options.encoding));
    }

    if (options.optimizeForBrowser !== undefined) {
      this.root.push(onOffObj("w:optimizeForBrowser", options.optimizeForBrowser));
    }

    if (options.relyOnVML !== undefined) {
      this.root.push(onOffObj("w:relyOnVML", options.relyOnVML));
    }

    if (options.allowPNG !== undefined) {
      this.root.push(onOffObj("w:allowPNG", options.allowPNG));
    }

    if (options.doNotRelyOnCSS !== undefined) {
      this.root.push(onOffObj("w:doNotRelyOnCSS", options.doNotRelyOnCSS));
    }

    if (options.doNotSaveAsSingleFile !== undefined) {
      this.root.push(onOffObj("w:doNotSaveAsSingleFile", options.doNotSaveAsSingleFile));
    }

    if (options.doNotOrganizeInFolder !== undefined) {
      this.root.push(onOffObj("w:doNotOrganizeInFolder", options.doNotOrganizeInFolder));
    }

    if (options.doNotUseLongFileNames !== undefined) {
      this.root.push(onOffObj("w:doNotUseLongFileNames", options.doNotUseLongFileNames));
    }

    if (options.pixelsPerInch !== undefined) {
      this.root.push(numberValObj("w:pixelsPerInch", options.pixelsPerInch));
    }

    if (options.targetScreenSz !== undefined) {
      this.root.push(stringValObj("w:targetScreenSz", options.targetScreenSz));
    }

    if (options.saveSmartTagsAsXml !== undefined) {
      this.root.push(onOffObj("w:saveSmartTagsAsXml", options.saveSmartTagsAsXml));
    }
  }
}
