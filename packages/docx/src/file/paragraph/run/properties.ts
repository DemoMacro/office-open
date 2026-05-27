/**
 * Run properties module for WordprocessingML documents.
 *
 * This module provides the run properties (rPr) element which specifies
 * the formatting applied to a run of text.
 *
 * Reference: https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
 * Section 17.3.2.21 (page 297)
 *
 * @module
 */
import { buildBorderObj } from "@file/border";
import type { BorderOptions } from "@file/border";
import { buildShadingObj } from "@file/shading";
import type { ShadingAttributesProperties } from "@file/shading";
import type { ChangedAttributesProperties } from "@file/track-revision/track-revision";
import { DeletionTrackChange } from "@file/track-revision/track-revision-components/deletion-track-change";
import { InsertionTrackChange } from "@file/track-revision/track-revision-components/insertion-track-change";
import type { IXmlableObject } from "@file/xml-components";
import { IgnoreIfEmptyXmlComponent, XmlComponent } from "@file/xml-components";
import {
  onOffObj,
  hpsMeasureObj,
  stringEnumValObj,
  stringValObj,
  numberValObj,
} from "@office-open/core";
import type { PositiveUniversalMeasure, UniversalMeasure } from "@util/values";

import { buildEastAsianLayoutObj } from "./east-asian-layout";
import type { EastAsianLayoutOptions } from "./east-asian-layout";
import { buildEmphasisMarkObj } from "./emphasis-mark";
import type { EmphasisMarkType } from "./emphasis-mark";
import {
  buildCharacterSpacingObj,
  buildColorObj,
  buildHighlightObj,
  buildHighlightComplexScriptObj,
} from "./formatting";
import type { ColorOptions } from "./formatting";
import { buildLanguageObj } from "./language";
import type { LanguageOptions } from "./language";
import { buildRunFontsObj } from "./run-fonts";
import type { FontAttributesProperties } from "./run-fonts";
import { buildUnderlineObj } from "./underline";
import type { UnderlineType } from "./underline";

interface RunFontReference {
  readonly name: string;
  readonly hint?: string;
}

/**
 * Text animation effect types.
 *
 * These effects specify animations that can be applied to text. Note that
 * these effects are deprecated and may not be supported by all applications.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TextEffect">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="blinkBackground"/>
 *     <xsd:enumeration value="lights"/>
 *     <xsd:enumeration value="antsBlack"/>
 *     <xsd:enumeration value="antsRed"/>
 *     <xsd:enumeration value="shimmer"/>
 *     <xsd:enumeration value="sparkle"/>
 *     <xsd:enumeration value="none"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const TextEffect = {
  /** Blinking background animation */
  BLINK_BACKGROUND: "blinkBackground",
  /** Lights animation effect */
  LIGHTS: "lights",
  /** Black marching ants animation */
  ANTS_BLACK: "antsBlack",
  /** Red marching ants animation */
  ANTS_RED: "antsRed",
  /** Shimmer animation effect */
  SHIMMER: "shimmer",
  /** Sparkle animation effect */
  SPARKLE: "sparkle",
  /** No text effect */
  NONE: "none",
} as const;

/**
 * Highlight color values for text highlighting.
 *
 * These colors specify the background highlight color that can be applied to text.
 *
 * Reference: http://officeopenxml.com/WPtextShading.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_HighlightColor">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="black"/>
 *     <xsd:enumeration value="blue"/>
 *     <xsd:enumeration value="cyan"/>
 *     <xsd:enumeration value="green"/>
 *     <xsd:enumeration value="magenta"/>
 *     <xsd:enumeration value="red"/>
 *     <xsd:enumeration value="yellow"/>
 *     <xsd:enumeration value="white"/>
 *     <xsd:enumeration value="darkBlue"/>
 *     <xsd:enumeration value="darkCyan"/>
 *     <xsd:enumeration value="darkGreen"/>
 *     <xsd:enumeration value="darkMagenta"/>
 *     <xsd:enumeration value="darkRed"/>
 *     <xsd:enumeration value="darkYellow"/>
 *     <xsd:enumeration value="darkGray"/>
 *     <xsd:enumeration value="lightGray"/>
 *     <xsd:enumeration value="none"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const HighlightColor = {
  /** Black highlight */
  BLACK: "black",
  /** Blue highlight */
  BLUE: "blue",
  /** Cyan highlight */
  CYAN: "cyan",
  /** Dark blue highlight */
  DARK_BLUE: "darkBlue",
  /** Dark cyan highlight */
  DARK_CYAN: "darkCyan",
  /** Dark gray highlight */
  DARK_GRAY: "darkGray",
  /** Dark green highlight */
  DARK_GREEN: "darkGreen",
  /** Dark magenta highlight */
  DARK_MAGENTA: "darkMagenta",
  /** Dark red highlight */
  DARK_RED: "darkRed",
  /** Dark yellow highlight */
  DARK_YELLOW: "darkYellow",
  /** Green highlight */
  GREEN: "green",
  /** Light gray highlight */
  LIGHT_GRAY: "lightGray",
  /** Magenta highlight */
  MAGENTA: "magenta",
  /** No highlight */
  NONE: "none",
  /** Red highlight */
  RED: "red",
  /** White highlight */
  WHITE: "white",
  /** Yellow highlight */
  YELLOW: "yellow",
} as const;

/**
 * Run style properties options.
 *
 * These properties define the formatting that can be applied to a run of text,
 * including font, size, bold, italic, underline, color, and other character formatting.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 */
export interface RunStylePropertiesOptions {
  readonly noProof?: boolean;
  readonly bold?: boolean;
  readonly boldComplexScript?: boolean;
  readonly italics?: boolean;
  readonly italicsComplexScript?: boolean;
  readonly underline?: {
    readonly color?: string;
    readonly type?: (typeof UnderlineType)[keyof typeof UnderlineType];
  };
  readonly effect?: (typeof TextEffect)[keyof typeof TextEffect];
  readonly emphasisMark?: {
    readonly type?: (typeof EmphasisMarkType)[keyof typeof EmphasisMarkType];
  };
  readonly color?: string | ColorOptions;
  readonly kern?: number | PositiveUniversalMeasure;
  readonly position?: UniversalMeasure;
  readonly size?: number | PositiveUniversalMeasure;
  readonly sizeComplexScript?: boolean | number | PositiveUniversalMeasure;
  readonly rightToLeft?: boolean;
  readonly smallCaps?: boolean;
  readonly allCaps?: boolean;
  readonly strike?: boolean;
  readonly doubleStrike?: boolean;
  readonly subScript?: boolean;
  readonly superScript?: boolean;
  readonly font?: string | RunFontReference | FontAttributesProperties;
  readonly highlight?: (typeof HighlightColor)[keyof typeof HighlightColor];
  readonly highlightComplexScript?: boolean | string;
  readonly characterSpacing?: number;
  readonly shading?: ShadingAttributesProperties;
  readonly emboss?: boolean;
  readonly imprint?: boolean;
  readonly revision?: IRunPropertiesChangeOptions;
  readonly language?: LanguageOptions;
  readonly border?: BorderOptions;
  readonly snapToGrid?: boolean;
  readonly vanish?: boolean;
  readonly specVanish?: boolean;
  readonly scale?: number;
  readonly math?: boolean;
  readonly outline?: boolean;
  readonly shadow?: boolean;
  readonly webHidden?: boolean;
  readonly fitText?: number;
  readonly complexScript?: boolean;
  readonly eastAsianLayout?: EastAsianLayoutOptions;
}

/**
 * Options for configuring run properties.
 *
 * Extends RunStylePropertiesOptions with a style reference.
 */
export type RunPropertiesOptions = {
  /** Reference to a character style by name */
  readonly style?: string;
} & RunStylePropertiesOptions;

/**
 * Options for run properties change tracking.
 *
 * Used for revision tracking when run properties have been modified.
 */
export type IRunPropertiesChangeOptions = {} & RunPropertiesOptions & ChangedAttributesProperties;

export type IParagraphRunPropertiesOptions = {
  readonly insertion?: ChangedAttributesProperties;
  readonly deletion?: ChangedAttributesProperties;
} & RunPropertiesOptions;

/**
 * Represents run properties (rPr) in a WordprocessingML document.
 *
 * Run properties specify all character-level formatting applied to text,
 * such as bold, italic, font, size, color, underline, and other text effects.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RPr">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_RPrContent" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * The EG_RPrBase group includes elements like rStyle, rFonts, b, bCs, i, iCs,
 * caps, smallCaps, strike, dstrike, outline, shadow, emboss, imprint, noProof,
 * snapToGrid, vanish, color, spacing, w, kern, position, sz, szCs, highlight,
 * u, effect, bdr, shd, vertAlign, rtl, em, lang, and more.
 */
export class RunProperties extends IgnoreIfEmptyXmlComponent {
  public constructor(options?: RunPropertiesOptions) {
    super("w:rPr");

    if (!options) {
      return;
    }

    const result = buildRunProperties(options);
    if (result) {
      const children = result["w:rPr"];
      if (Array.isArray(children)) {
        this.root.push(...children);
      }
    }
  }

  public push(item: XmlComponent): void {
    this.root.push(item);
  }
}

export class ParagraphRunProperties extends RunProperties {
  public constructor(options?: IParagraphRunPropertiesOptions) {
    super(options);

    if (options?.insertion) {
      this.push(new InsertionTrackChange(options.insertion));
    }

    if (options?.deletion) {
      this.push(new DeletionTrackChange(options.deletion));
    }
  }
}

/**
 * Represents a run properties change element for revision tracking.
 *
 * This element is used to track changes to run properties when revision
 * tracking is enabled in the document.
 */
export class RunPropertiesChange extends XmlComponent {
  public constructor(options: IRunPropertiesChangeOptions) {
    super("w:rPrChange");
    this.root.push({
      _attr: { "w:author": options.author, "w:date": options.date, "w:id": options.id },
    });
    this.addChildElement(new RunProperties(options as RunPropertiesOptions));
  }
}

/**
 * Build run properties (w:rPr) as IXmlableObject without allocating XmlComponent tree.
 * Used by Run.prepForXml for O(1) construction.
 */
export function buildRunProperties(options?: RunPropertiesOptions): IXmlableObject | undefined {
  if (!options) return undefined;

  const children: IXmlableObject[] = [];

  if (options.style) {
    children.push(stringValObj("w:rStyle", options.style));
  }

  if (options.font) {
    if (typeof options.font === "string") {
      children.push(buildRunFontsObj(options.font));
    } else if ("name" in options.font) {
      children.push(buildRunFontsObj(options.font.name, options.font.hint));
    } else {
      children.push(buildRunFontsObj(options.font));
    }
  }

  if (options.bold !== undefined) children.push(onOffObj("w:b", options.bold));

  if (
    (options.boldComplexScript === undefined && options.bold !== undefined) ||
    options.boldComplexScript
  ) {
    children.push(onOffObj("w:bCs", options.boldComplexScript ?? options.bold));
  }

  if (options.italics !== undefined) children.push(onOffObj("w:i", options.italics));

  if (
    (options.italicsComplexScript === undefined && options.italics !== undefined) ||
    options.italicsComplexScript
  ) {
    children.push(onOffObj("w:iCs", options.italicsComplexScript ?? options.italics));
  }

  if (options.smallCaps !== undefined) {
    children.push(onOffObj("w:smallCaps", options.smallCaps));
  } else if (options.allCaps !== undefined) {
    children.push(onOffObj("w:caps", options.allCaps));
  }

  if (options.strike !== undefined) children.push(onOffObj("w:strike", options.strike));
  if (options.doubleStrike !== undefined)
    children.push(onOffObj("w:dstrike", options.doubleStrike));
  if (options.emboss !== undefined) children.push(onOffObj("w:emboss", options.emboss));
  if (options.imprint !== undefined) children.push(onOffObj("w:imprint", options.imprint));
  if (options.outline !== undefined) children.push(onOffObj("w:outline", options.outline));
  if (options.shadow !== undefined) children.push(onOffObj("w:shadow", options.shadow));
  if (options.webHidden !== undefined) children.push(onOffObj("w:webHidden", options.webHidden));
  if (options.noProof !== undefined) children.push(onOffObj("w:noProof", options.noProof));
  if (options.snapToGrid !== undefined) children.push(onOffObj("w:snapToGrid", options.snapToGrid));

  if (options.vanish) {
    children.push(onOffObj("w:vanish", options.vanish));
  }

  if (options.color) {
    children.push(buildColorObj(options.color));
  }

  if (options.characterSpacing) {
    children.push(buildCharacterSpacingObj(options.characterSpacing));
  }

  if (options.scale !== undefined) {
    children.push(numberValObj("w:w", options.scale));
  }

  if (options.kern) {
    children.push(hpsMeasureObj("w:kern", options.kern));
  }

  if (options.position) {
    children.push(stringValObj("w:position", options.position));
  }

  if (options.size !== undefined) {
    children.push(hpsMeasureObj("w:sz", options.size));
  }
  const szCs =
    options.sizeComplexScript === undefined || options.sizeComplexScript === true
      ? options.size
      : options.sizeComplexScript;
  if (szCs) {
    children.push(hpsMeasureObj("w:szCs", szCs));
  }

  if (options.highlight) {
    children.push(buildHighlightObj(options.highlight));
  }
  if (options.highlightComplexScript === true) {
    if (options.highlight) {
      children.push(buildHighlightComplexScriptObj(options.highlight));
    }
  } else if (
    options.highlightComplexScript !== undefined &&
    options.highlightComplexScript !== false
  ) {
    children.push(buildHighlightComplexScriptObj(options.highlightComplexScript));
  }

  if (options.underline) {
    children.push(buildUnderlineObj(options.underline.type, options.underline.color));
  }

  if (options.effect) {
    children.push(stringValObj("w:effect", options.effect));
  }

  if (options.border) {
    children.push(buildBorderObj("w:bdr", options.border));
  }

  if (options.shading) {
    children.push(buildShadingObj(options.shading));
  }

  if (options.subScript) {
    children.push(stringEnumValObj("w:vertAlign", "subscript"));
  }

  if (options.superScript) {
    children.push(stringEnumValObj("w:vertAlign", "superscript"));
  }

  if (options.rightToLeft !== undefined) {
    children.push(onOffObj("w:rtl", options.rightToLeft));
  }

  if (options.emphasisMark) {
    children.push(buildEmphasisMarkObj(options.emphasisMark.type));
  }

  if (options.language) {
    children.push(buildLanguageObj(options.language));
  }

  if (options.specVanish) {
    children.push(onOffObj("w:specVanish", options.specVanish));
  }

  if (options.math) {
    children.push(onOffObj("w:oMath", options.math));
  }

  if (options.fitText !== undefined) {
    children.push(numberValObj("w:fitText", options.fitText));
  }

  if (options.complexScript !== undefined) {
    children.push(onOffObj("w:cs", options.complexScript));
  }

  if (options.eastAsianLayout) {
    children.push(buildEastAsianLayoutObj(options.eastAsianLayout));
  }

  if (options.revision) {
    const rev = options.revision;
    const { author: _, date: __, id: ___, ...originalProps } = rev;
    const rPrChangeChildren: (IXmlableObject | { _attr: Record<string, string | number> })[] = [
      { _attr: { "w:author": rev.author, "w:date": rev.date, "w:id": rev.id } },
    ];
    const innerRPr = buildRunProperties(originalProps as RunPropertiesOptions);
    if (innerRPr) rPrChangeChildren.push(innerRPr);
    children.push({ "w:rPrChange": rPrChangeChildren });
  }

  return children.length > 0 ? { "w:rPr": children } : undefined;
}
