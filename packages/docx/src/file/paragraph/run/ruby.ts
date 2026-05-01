/**
 * Ruby annotation module for WordprocessingML documents.
 *
 * Ruby annotations are small text rendered above or below base text,
 * commonly used for pronunciation guides in East Asian languages (furigana
 * in Japanese, pinyin in Chinese, etc.).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Ruby
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { Text } from "./run-components/text";

/**
 * Ruby alignment options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_RubyAlign">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="center"/>
 *     <xsd:enumeration value="distributeLetter"/>
 *     <xsd:enumeration value="distributeSpace"/>
 *     <xsd:enumeration value="left"/>
 *     <xsd:enumeration value="right"/>
 *     <xsd:enumeration value="rightVertical"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const RubyAlign = {
    /** Center alignment */
    CENTER: "center",
    /** Distribute letters evenly */
    DISTRIBUTE_LETTER: "distributeLetter",
    /** Distribute space evenly */
    DISTRIBUTE_SPACE: "distributeSpace",
    /** Left alignment */
    LEFT: "left",
    /** Right alignment */
    RIGHT: "right",
    /** Right vertical alignment */
    RIGHT_VERTICAL: "rightVertical",
} as const;

/**
 * Options for a Ruby annotation.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Ruby">
 *   <xsd:sequence>
 *     <xsd:element name="rubyPr" type="CT_RubyPr"/>
 *     <xsd:element name="rt" type="CT_RubyContent"/>
 *     <xsd:element name="rubyBase" type="CT_RubyContent"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export interface RubyOptions {
    /** Ruby annotation text (e.g., furigana, pinyin) */
    readonly text: string;
    /** Base text being annotated */
    readonly base: string;
    /**
     * Ruby alignment (defaults to "center").
     */
    readonly alignment?: keyof typeof RubyAlign;
    /**
     * Font size for the ruby annotation text in half-points (e.g., 20 = 10pt).
     *
     * Defaults to half the base text size if not specified.
     */
    readonly fontSize?: number;
    /**
     * Vertical offset for the ruby annotation in half-points.
     *
     * How far the annotation is raised above (or below) the base text.
     */
    readonly raise?: number;
    /**
     * Font size for the base text in half-points.
     *
     * Used to calculate the ruby annotation positioning.
     */
    readonly baseFontSize?: number;
    /**
     * Language identifier for the ruby annotation (e.g., "ja-JP").
     */
    readonly languageId?: string;
    /**
     * Whether the ruby annotation is dirty (needs recalculation).
     */
    readonly dirty?: boolean;
}

/**
 * Creates a ruby content element (w:rt or w:rubyBase).
 *
 * Ruby content contains a run with text.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RubyContent">
 *   <xsd:group ref="EG_RubyContent" minOccurs="0" maxOccurs="unbounded"/>
 * </xsd:complexType>
 * ```
 */
const createRubyContent = (name: string, text: string): XmlComponent =>
    new BuilderElement({
        children: [
            new BuilderElement({
                children: [new Text(text)],
                name: "w:r",
            }),
        ],
        name,
    });

/**
 * Creates a Ruby annotation element (w:ruby).
 *
 * Ruby annotations provide pronunciation guides for East Asian text,
 * commonly used for furigana in Japanese and pinyin in Chinese.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Ruby">
 *   <xsd:sequence>
 *     <xsd:element name="rubyPr" type="CT_RubyPr"/>
 *     <xsd:element name="rt" type="CT_RubyContent"/>
 *     <xsd:element name="rubyBase" type="CT_RubyContent"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Japanese furigana
 * createRuby({
 *   text: "かな",
 *   base: "漢字",
 *   alignment: "CENTER",
 *   fontSize: 20,
 *   raise: 20,
 *   languageId: "ja-JP",
 * });
 * ```
 */
export const createRuby = (options: RubyOptions): XmlComponent => {
    const align = options.alignment ?? "CENTER";
    const hps = options.fontSize ?? 20;
    const hpsRaise = options.raise ?? 20;
    const hpsBaseText = options.baseFontSize ?? 40;

    const rubyPr = new BuilderElement({
        children: [
            new BuilderElement<{ readonly val: string }>({
                attributes: { val: { key: "w:val", value: RubyAlign[align] } },
                name: "w:rubyAlign",
            }),
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: hps } },
                name: "w:hps",
            }),
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: hpsRaise } },
                name: "w:hpsRaise",
            }),
            new BuilderElement<{ readonly val: number }>({
                attributes: { val: { key: "w:val", value: hpsBaseText } },
                name: "w:hpsBaseText",
            }),
            new BuilderElement<{ readonly val: string }>({
                attributes: {
                    val: {
                        key: "w:val",
                        value: options.languageId ?? "ja-JP",
                    },
                },
                name: "w:lid",
            }),
            ...(options.dirty !== undefined
                ? [
                      new BuilderElement({
                          name: "w:dirty",
                      }),
                  ]
                : []),
        ],
        name: "w:rubyPr",
    });

    const rt = createRubyContent("w:rt", options.text);
    const rubyBase = createRubyContent("w:rubyBase", options.base);

    return new BuilderElement({
        children: [rubyPr, rt, rubyBase],
        name: "w:ruby",
    });
};
