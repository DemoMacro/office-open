/**
 * Document background module for WordprocessingML documents.
 *
 * This module provides functionality for setting document background colors,
 * theme-based backgrounds, and background images.
 *
 * Background images use VML (Vector Markup Language) via `v:background` and `v:fill`,
 * which is how Word implements page background images. The `v:fill` element references
 * the image through a relationship (`r:id`).
 *
 * Reference: http://officeopenxml.com/WPdocument.php
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { hashedId } from "@util/convenience-functions";
import { hexColorValue, uCharHexNumber } from "@util/values";
import type { DataType } from "undio";
import { toUint8Array } from "undio";

/**
 * Attributes for the document background element.
 *
 * ## XSD Schema (ST_ThemeColor)
 * ```xml
 * <xsd:simpleType name="ST_ThemeColor">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="dark1"/>
 *     <xsd:enumeration value="light1"/>
 *     <xsd:enumeration value="dark2"/>
 *     <xsd:enumeration value="light2"/>
 *     <xsd:enumeration value="accent1"/>
 *     <xsd:enumeration value="accent2"/>
 *     <xsd:enumeration value="accent3"/>
 *     <xsd:enumeration value="accent4"/>
 *     <xsd:enumeration value="accent5"/>
 *     <xsd:enumeration value="accent6"/>
 *     <xsd:enumeration value="hyperlink"/>
 *     <xsd:enumeration value="followedHyperlink"/>
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="background1"/>
 *     <xsd:enumeration value="text1"/>
 *     <xsd:enumeration value="background2"/>
 *     <xsd:enumeration value="text2"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @internal
 */
export class DocumentBackgroundAttributes extends XmlAttributeComponent<{
    readonly color?: string;
    readonly themeColor?: string;
    readonly themeShade?: string;
    readonly themeTint?: string;
}> {
    protected readonly xmlKeys = {
        color: "w:color",
        themeColor: "w:themeColor",
        themeShade: "w:themeShade",
        themeTint: "w:themeTint",
    };
}

/**
 * Image options for document background.
 *
 * Specifies the image data and type for a background image.
 * The image is rendered as a full-page background using VML `v:fill`.
 */
export interface IBackgroundImageOptions {
    /** Raw image data (Uint8Array, base64 string, etc.) */
    readonly data: DataType;
    /** Image format type */
    readonly type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}

/**
 * Options for creating a document background.
 *
 * @see {@link DocumentBackground}
 */
export interface IDocumentBackgroundOptions {
    /** Background color in hex format (e.g., "FF0000" for red) */
    readonly color?: string;
    /** Theme color name (e.g., "accent1", "dark1") */
    readonly themeColor?: string;
    /** Theme shade value (darkens the theme color) */
    readonly themeShade?: string;
    /** Theme tint value (lightens the theme color) */
    readonly themeTint?: string;
    /** Background image rendered as a full-page VML fill */
    readonly image?: IBackgroundImageOptions;
}

interface ImageData {
    readonly fileName: string;
    readonly data: Uint8Array;
    readonly type: string;
}

/**
 * Represents a document background in a WordprocessingML document.
 *
 * The background element specifies the background color or theme color
 * for the document, and optionally a background image.
 *
 * Background images use VML `v:background` with `v:fill` to create a true
 * page background (not just a drawing behind text). This is the same approach
 * Word uses for page background images.
 *
 * Reference: http://officeopenxml.com/WPdocument.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Background">
 *   <xsd:sequence>
 *     <xsd:element name="drawing" type="CT_Drawing" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="color" type="ST_HexColor" use="optional" default="auto"/>
 *   <xsd:attribute name="themeColor" type="ST_ThemeColor" use="optional"/>
 *   <xsd:attribute name="themeTint" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="themeShade" type="ST_UcharHexNumber" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * Note: Background images use VML (`v:background`/`v:fill`) rather than DrawingML
 * (`w:drawing`/`wp:anchor`), as VML is how Word implements true page backgrounds.
 *
 * @example
 * ```typescript
 * new DocumentBackground({ color: "FFFF00" }); // Yellow background
 * new DocumentBackground({ themeColor: "accent1" }); // Theme accent color
 * new DocumentBackground({
 *   color: "E8E8E8",
 *   image: { data: fs.readFileSync("bg.png"), type: "png" },
 * });
 * ```
 */
export class DocumentBackground extends XmlComponent {
    private readonly imageData?: ImageData;

    public constructor(options: IDocumentBackgroundOptions) {
        super("w:background");

        this.root.push(
            new DocumentBackgroundAttributes({
                color: options.color === undefined ? undefined : hexColorValue(options.color),
                themeColor: options.themeColor,
                themeShade:
                    options.themeShade === undefined
                        ? undefined
                        : uCharHexNumber(options.themeShade),
                themeTint:
                    options.themeTint === undefined ? undefined : uCharHexNumber(options.themeTint),
            }),
        );

        if (options.image) {
            const rawData = toUint8Array(options.image.data) as Uint8Array;
            const hash = hashedId(rawData);
            this.imageData = {
                data: rawData,
                fileName: `${hash}.${options.image.type}`,
                type: options.image.type,
            };
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        if (this.imageData) {
            // Register the image with the media collection for packaging
            context.file.Media.addImage(this.imageData.fileName, {
                type: this.imageData.type as
                    | "jpg"
                    | "png"
                    | "gif"
                    | "bmp"
                    | "tif"
                    | "ico"
                    | "emf"
                    | "wmf",
                data: this.imageData.data,
                fileName: this.imageData.fileName,
                transformation: {
                    emus: { x: 0, y: 0 },
                    pixels: { x: 0, y: 0 },
                },
            });

            // VML background with fill referencing the image via relationship
            // Word uses v:background + v:fill for true page background images
            this.root.push(
                new BuilderElement({
                    attributes: {
                        id: { key: "id", value: "_x0000_s1025" },
                    },
                    children: [
                        new BuilderElement({
                            attributes: {
                                id: {
                                    key: "r:id",
                                    value: `{${this.imageData.fileName}}`,
                                },
                                title: {
                                    key: "o:title",
                                    value: this.imageData.fileName.split(".")[0],
                                },
                                recolor: { key: "recolor", value: "t" },
                                type: { key: "type", value: "frame" },
                            },
                            name: "v:fill",
                        }),
                    ],
                    name: "v:background",
                }),
            );
        }

        return super.prepForXml(context);
    }
}
