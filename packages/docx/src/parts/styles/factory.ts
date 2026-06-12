/**
 * Factory module for creating default document styles.
 *
 * Creates styles matching Microsoft Word's default document template.
 * Pure string generation — no XmlComponent inheritance.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import type { IParagraphStylePropertiesOptions } from "@parts/paragraph/properties";
import type { RunStylePropertiesOptions } from "@parts/paragraph/run/properties";

import { stringifyParagraphProperties, stringifyRunProperties } from "../paragraph/stringify";
import type { StylesOptions } from "./styles";

// ── Style options interfaces ──

export interface DefaultStylesOptions {
  document?: DocumentDefaultsOptions;
  title?: IBaseParagraphStyleOptions;
  subtitle?: IBaseParagraphStyleOptions;
  heading1?: IBaseParagraphStyleOptions;
  heading2?: IBaseParagraphStyleOptions;
  heading3?: IBaseParagraphStyleOptions;
  heading4?: IBaseParagraphStyleOptions;
  heading5?: IBaseParagraphStyleOptions;
  heading6?: IBaseParagraphStyleOptions;
  heading7?: IBaseParagraphStyleOptions;
  heading8?: IBaseParagraphStyleOptions;
  heading9?: IBaseParagraphStyleOptions;
  strong?: IBaseParagraphStyleOptions;
  emphasis?: IBaseParagraphStyleOptions;
  listParagraph?: IBaseParagraphStyleOptions;
  quote?: IBaseParagraphStyleOptions;
  intenseQuote?: IBaseParagraphStyleOptions;
  hyperlink?: IBaseCharacterStyleOptions;
  footnoteReference?: IBaseCharacterStyleOptions;
  footnoteText?: IBaseParagraphStyleOptions;
  footnoteTextChar?: IBaseCharacterStyleOptions;
  endnoteReference?: IBaseCharacterStyleOptions;
  endnoteText?: IBaseParagraphStyleOptions;
  endnoteTextChar?: IBaseCharacterStyleOptions;
}

export interface DocumentDefaultsOptions {
  paragraph?: IParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
}

interface StyleOptions {
  name?: string;
  aliases?: string;
  basedOn?: string;
  next?: string;
  link?: string;
  autoRedefine?: boolean;
  uiPriority?: number;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  quickFormat?: boolean;
  locked?: boolean;
  personal?: boolean;
  personalCompose?: boolean;
  personalReply?: boolean;
}

export type IBaseParagraphStyleOptions = {
  paragraph?: IParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
} & StyleOptions & { id?: string };

export type IBaseCharacterStyleOptions = {
  run?: RunStylePropertiesOptions;
} & StyleOptions & { id?: string };

// ── String builders ──

/** Escape special XML characters. */
function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Build `<w:style>` XML for a paragraph style. */
export function stringifyParagraphStyle(
  opts: StyleOptions & {
    id: string;
    paragraph?: IParagraphStylePropertiesOptions;
    run?: RunStylePropertiesOptions;
  },
): string {
  const children: string[] = [];

  children.push(`<w:name w:val="${esc(opts.name ?? opts.id)}"/>`);
  if (opts.aliases) children.push(`<w:aliases w:val="${esc(opts.aliases)}"/>`);
  if (opts.basedOn) children.push(`<w:basedOn w:val="${esc(opts.basedOn)}"/>`);
  if (opts.next) children.push(`<w:next w:val="${esc(opts.next)}"/>`);
  if (opts.link) children.push(`<w:link w:val="${esc(opts.link)}"/>`);
  if (opts.autoRedefine) children.push("<w:autoRedefine/>");
  if (opts.uiPriority !== undefined) children.push(`<w:uiPriority w:val="${opts.uiPriority}"/>`);
  if (opts.semiHidden) children.push("<w:semiHidden/>");
  if (opts.unhideWhenUsed) children.push("<w:unhideWhenUsed/>");
  if (opts.quickFormat) children.push("<w:qFormat/>");
  if (opts.locked) children.push("<w:locked/>");
  if (opts.personal) children.push("<w:personal/>");
  if (opts.personalCompose) children.push("<w:personalCompose/>");
  if (opts.personalReply) children.push("<w:personalReply/>");

  const pPr = stringifyParagraphProperties(opts.paragraph).xml;
  if (pPr) children.push(pPr);
  const rPr = stringifyRunProperties(opts.run);
  if (rPr) children.push(rPr);

  return `<w:style w:type="paragraph" w:styleId="${esc(opts.id)}">${children.join("")}</w:style>`;
}

/** Build `<w:style>` XML for a character style. */
export function stringifyCharacterStyle(
  opts: StyleOptions & {
    id: string;
    run?: RunStylePropertiesOptions;
  },
): string {
  const children: string[] = [];

  children.push(`<w:name w:val="${esc(opts.name ?? opts.id)}"/>`);
  if (opts.aliases) children.push(`<w:aliases w:val="${esc(opts.aliases)}"/>`);
  if (opts.basedOn) children.push(`<w:basedOn w:val="${esc(opts.basedOn)}"/>`);
  if (opts.link) children.push(`<w:link w:val="${esc(opts.link)}"/>`);
  if (opts.autoRedefine) children.push("<w:autoRedefine/>");
  if (opts.uiPriority !== undefined) children.push(`<w:uiPriority w:val="${opts.uiPriority}"/>`);
  if (opts.semiHidden) children.push("<w:semiHidden/>");
  if (opts.unhideWhenUsed) children.push("<w:unhideWhenUsed/>");
  if (opts.locked) children.push("<w:locked/>");
  if (opts.personal) children.push("<w:personal/>");
  if (opts.personalCompose) children.push("<w:personalCompose/>");
  if (opts.personalReply) children.push("<w:personalReply/>");

  const rPr = stringifyRunProperties(opts.run);
  if (rPr) children.push(rPr);

  return `<w:style w:type="character" w:styleId="${esc(opts.id)}">${children.join("")}</w:style>`;
}

/** Build `<w:docDefaults>` XML matching Word's default settings. */
function stringifyDocDefaults(opts: DocumentDefaultsOptions): string {
  const children: string[] = [];

  // rPrDefault - Word default: Calibri 11pt (22 half-points), theme fonts
  const rPr = stringifyRunProperties(opts.run);
  if (rPr) {
    children.push(`<w:rPrDefault>${rPr}</w:rPrDefault>`);
  } else {
    // Match Word's default run properties exactly
    children.push(
      `<w:rPrDefault><w:rPr>` +
        `<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>` +
        `<w:kern w:val="2"/>` +
        `<w:sz w:val="22"/><w:szCs w:val="24"/>` +
        `<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>` +
        `<w14:ligatures w14:val="standardContextual"/>` +
        `</w:rPr></w:rPrDefault>`,
    );
  }

  // pPrDefault - Word default: spacing after 8pt (160 twips), line 1.16 (278 twips)
  const pPr = stringifyParagraphProperties(opts.paragraph).xml;
  if (pPr) {
    children.push(`<w:pPrDefault>${pPr}</w:pPrDefault>`);
  } else {
    // Match Word's default paragraph properties exactly
    children.push(
      `<w:pPrDefault><w:pPr>` +
        `<w:spacing w:after="160" w:line="278" w:lineRule="auto"/>` +
        `</w:pPr></w:pPrDefault>`,
    );
  }

  return `<w:docDefaults>${children.join("")}</w:docDefaults>`;
}

// ── Factory ──

let cachedDefaultStyles: StylesOptions | null = null;

/**
 * Factory for creating default document styles.
 *
 * Creates styles matching Microsoft Word's default document template.
 * Pure string generation — no XmlComponent inheritance.
 */
export class DefaultStylesFactory {
  public newInstance(options: DefaultStylesOptions = {}): StylesOptions {
    if (Object.keys(options).length === 0) {
      if (!cachedDefaultStyles) {
        cachedDefaultStyles = this.build({});
      }
      return cachedDefaultStyles;
    }
    return this.build(options);
  }

  private build(options: DefaultStylesOptions): StylesOptions {
    const importedStyles: { _raw: string }[] = [];

    // XML namespace attributes for styles root element (matching Word exactly)
    const initialAttributes: Record<string, string> = {
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
      "mc:Ignorable": "w14 w15",
    };

    importedStyles.push({ _raw: stringifyDocDefaults(options.document ?? {}) });

    // Latent styles - complete list from Word's default template
    // Only include styles that are NOT explicitly defined below
    importedStyles.push({
      _raw:
        `<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">` +
        `<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>` +
        `<w:lsdException w:name="Normal Table" w:semiHidden="1" w:uiPriority="99" w:unhideWhenUsed="1"/>` +
        `<w:lsdException w:name="No List" w:semiHidden="1" w:uiPriority="99" w:unhideWhenUsed="1"/>` +
        `<w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>` +
        `<w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>` +
        `<w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>` +
        `<w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>` +
        `<w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>` +
        `<w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>` +
        `<w:lsdException w:name="Revision" w:semiHidden="1"/>` +
        `</w:latentStyles>`,
    });

    // Built-in styles required by Word
    // Normal paragraph style (default for paragraphs)
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
        `<w:name w:val="Normal"/><w:qFormat/>` +
        `<w:pPr><w:widowControl w:val="0"/></w:pPr>` +
        `</w:style>`,
    });

    // heading 1-9 styles with proper formatting
    const headings = [
      {
        id: "Heading1",
        name: "heading 1",
        link: "Heading1Char",
        sz: "48",
        before: "480",
        after: "80",
        outlineLvl: "0",
      },
      {
        id: "Heading2",
        name: "heading 2",
        link: "Heading2Char",
        sz: "40",
        before: "160",
        after: "80",
        outlineLvl: "1",
      },
      {
        id: "Heading3",
        name: "heading 3",
        link: "Heading3Char",
        sz: "32",
        before: "160",
        after: "80",
        outlineLvl: "2",
      },
      {
        id: "Heading4",
        name: "heading 4",
        link: "Heading4Char",
        sz: "28",
        before: "80",
        after: "40",
        outlineLvl: "3",
      },
      {
        id: "Heading5",
        name: "heading 5",
        link: "Heading5Char",
        sz: "24",
        before: "80",
        after: "40",
        outlineLvl: "4",
      },
      {
        id: "Heading6",
        name: "heading 6",
        link: "Heading6Char",
        sz: undefined,
        before: "40",
        after: "0",
        outlineLvl: "5",
      },
      {
        id: "Heading7",
        name: "heading 7",
        link: "Heading7Char",
        sz: undefined,
        before: "40",
        after: "0",
        outlineLvl: "6",
      },
      {
        id: "Heading8",
        name: "heading 8",
        link: "Heading8Char",
        sz: undefined,
        before: undefined,
        after: "0",
        outlineLvl: "7",
      },
      {
        id: "Heading9",
        name: "heading 9",
        link: "Heading9Char",
        sz: undefined,
        before: undefined,
        after: "0",
        outlineLvl: "8",
      },
    ];

    for (const h of headings) {
      const pPrParts: string[] = [`<w:keepNext/>`, `<w:keepLines/>`];
      if (h.before || h.after) {
        const sp: string[] = [];
        if (h.before) sp.push(`w:before="${h.before}"`);
        if (h.after) sp.push(`w:after="${h.after}"`);
        pPrParts.push(`<w:spacing ${sp.join(" ")}/>`);
      }
      pPrParts.push(`<w:outlineLvl w:val="${h.outlineLvl}"/>`);

      const rPrParts: string[] = [];
      // heading 1-6 use accent1 color, 7-9 use text1 color
      if (parseInt(h.outlineLvl) < 6) {
        rPrParts.push(
          `<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>`,
        );
        rPrParts.push(`<w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/>`);
      } else {
        rPrParts.push(`<w:rFonts w:cstheme="majorBidi"/>`);
        rPrParts.push(`<w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/>`);
      }
      if (h.sz) {
        rPrParts.push(`<w:sz w:val="${h.sz}"/><w:szCs w:val="${h.sz}"/>`);
      }
      // heading 6-9 have bold
      if (parseInt(h.outlineLvl) >= 5) {
        rPrParts.push(`<w:b/><w:bCs/>`);
      }

      importedStyles.push({
        _raw:
          `<w:style w:type="paragraph" w:styleId="${h.id}">` +
          `<w:name w:val="${h.name}"/>` +
          `<w:basedOn w:val="Normal"/>` +
          `<w:next w:val="Normal"/>` +
          `<w:link w:val="${h.link}"/>` +
          `<w:uiPriority w:val="9"/>` +
          (parseInt(h.outlineLvl) > 0 ? `<w:semiHidden/><w:unhideWhenUsed/>` : "") +
          `<w:qFormat/>` +
          `<w:pPr>${pPrParts.join("")}</w:pPr>` +
          `<w:rPr>${rPrParts.join("")}</w:rPr>` +
          `</w:style>`,
      });

      // Linked character styles for headings
      importedStyles.push({
        _raw:
          `<w:style w:type="character" w:styleId="${h.link}">` +
          `<w:name w:val="${h.name} Char"/>` +
          `<w:basedOn w:val="DefaultParagraphFont"/>` +
          `<w:link w:val="${h.id}"/>` +
          `<w:uiPriority w:val="9"/>` +
          (parseInt(h.outlineLvl) > 0 ? `<w:semiHidden/>` : "") +
          `<w:rPr>${rPrParts.join("")}</w:rPr>` +
          `</w:style>`,
      });
    }

    // DefaultParagraphFont character style (default for runs)
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">` +
        `<w:name w:val="Default Paragraph Font"/>` +
        `<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>` +
        `</w:style>`,
    });

    // Normal Table style (default for tables)
    importedStyles.push({
      _raw:
        `<w:style w:type="table" w:default="1" w:styleId="NormalTable">` +
        `<w:name w:val="Normal Table"/>` +
        `<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>` +
        `<w:tblPr>` +
        `<w:tblInd w:w="0" w:type="dxa"/>` +
        `<w:tblCellMar>` +
        `<w:top w:w="0" w:type="dxa"/>` +
        `<w:left w:w="108" w:type="dxa"/>` +
        `<w:bottom w:w="0" w:type="dxa"/>` +
        `<w:right w:w="108" w:type="dxa"/>` +
        `</w:tblCellMar>` +
        `</w:tblPr>` +
        `</w:style>`,
    });

    // No List numbering style (default for numbering)
    importedStyles.push({
      _raw:
        `<w:style w:type="numbering" w:default="1" w:styleId="NoList">` +
        `<w:name w:val="No List"/>` +
        `<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>` +
        `</w:style>`,
    });

    // Title style
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:styleId="Title">` +
        `<w:name w:val="Title"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:link w:val="TitleChar"/>` +
        `<w:uiPriority w:val="10"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/><w:jc w:val="center"/></w:pPr>` +
        `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>` +
        `<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>` +
        `</w:style>`,
    });

    // Title Char style
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:styleId="TitleChar">` +
        `<w:name w:val="Title Char"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:link w:val="Title"/>` +
        `<w:uiPriority w:val="10"/>` +
        `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>` +
        `<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>` +
        `</w:style>`,
    });

    // Subtitle style
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:styleId="Subtitle">` +
        `<w:name w:val="Subtitle"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:link w:val="SubtitleChar"/>` +
        `<w:uiPriority w:val="11"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:jc w:val="center"/></w:pPr>` +
        `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>` +
        `<w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/>` +
        `<w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
        `</w:style>`,
    });

    // Subtitle Char style
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:styleId="SubtitleChar">` +
        `<w:name w:val="Subtitle Char"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:link w:val="Subtitle"/>` +
        `<w:uiPriority w:val="11"/>` +
        `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>` +
        `<w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/>` +
        `<w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
        `</w:style>`,
    });

    // List Paragraph style
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:styleId="ListParagraph">` +
        `<w:name w:val="List Paragraph"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:uiPriority w:val="34"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>` +
        `</w:style>`,
    });

    // Quote style
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:styleId="Quote">` +
        `<w:name w:val="Quote"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:link w:val="QuoteChar"/>` +
        `<w:uiPriority w:val="29"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:spacing w:before="160"/><w:jc w:val="center"/></w:pPr>` +
        `<w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/></w:rPr>` +
        `</w:style>`,
    });

    // Quote Char style
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:styleId="QuoteChar">` +
        `<w:name w:val="Quote Char"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:link w:val="Quote"/>` +
        `<w:uiPriority w:val="29"/>` +
        `<w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/></w:rPr>` +
        `</w:style>`,
    });

    // Intense Quote style
    importedStyles.push({
      _raw:
        `<w:style w:type="paragraph" w:styleId="IntenseQuote">` +
        `<w:name w:val="Intense Quote"/>` +
        `<w:basedOn w:val="Normal"/>` +
        `<w:next w:val="Normal"/>` +
        `<w:link w:val="IntenseQuoteChar"/>` +
        `<w:uiPriority w:val="30"/>` +
        `<w:qFormat/>` +
        `<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="0F4761" w:themeColor="accent1" w:themeShade="BF"/>` +
        `<w:bottom w:val="single" w:sz="4" w:space="10" w:color="0F4761" w:themeColor="accent1" w:themeShade="BF"/></w:pBdr>` +
        `<w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>` +
        `<w:rPr><w:i/><w:iCs/><w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/></w:rPr>` +
        `</w:style>`,
    });

    // Intense Quote Char style
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:styleId="IntenseQuoteChar">` +
        `<w:name w:val="Intense Quote Char"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:link w:val="IntenseQuote"/>` +
        `<w:uiPriority w:val="30"/>` +
        `<w:rPr><w:i/><w:iCs/><w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/></w:rPr>` +
        `</w:style>`,
    });

    // Hyperlink character style
    importedStyles.push({
      _raw: stringifyCharacterStyle({
        id: "Hyperlink",
        name: "Hyperlink",
        basedOn: "DefaultParagraphFont",
        semiHidden: true,
        unhideWhenUsed: true,
        run: { color: "0563C1", underline: { type: "single" } },
        ...options.hyperlink,
      }),
    });

    // Footnote Reference character style
    importedStyles.push({
      _raw: stringifyCharacterStyle({
        id: "FootnoteReference",
        name: "footnote reference",
        basedOn: "DefaultParagraphFont",
        semiHidden: true,
        unhideWhenUsed: true,
        run: { superScript: true },
        ...options.footnoteReference,
      }),
    });

    // Footnote Text paragraph style
    importedStyles.push({
      _raw: stringifyParagraphStyle({
        id: "FootnoteText",
        name: "footnote text",
        basedOn: "Normal",
        link: "FootnoteTextChar",
        semiHidden: true,
        uiPriority: 99,
        unhideWhenUsed: true,
        paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } },
        run: { size: 20 },
        ...options.footnoteText,
      }),
    });

    // Footnote Text Char character style
    importedStyles.push({
      _raw: stringifyCharacterStyle({
        id: "FootnoteTextChar",
        name: "Footnote Text Char",
        basedOn: "DefaultParagraphFont",
        link: "FootnoteText",
        semiHidden: true,
        run: { size: 20 },
        ...options.footnoteTextChar,
      }),
    });

    // Endnote Reference character style
    importedStyles.push({
      _raw: stringifyCharacterStyle({
        id: "EndnoteReference",
        name: "endnote reference",
        basedOn: "DefaultParagraphFont",
        semiHidden: true,
        unhideWhenUsed: true,
        run: { superScript: true },
        ...options.endnoteReference,
      }),
    });

    // Endnote Text paragraph style
    importedStyles.push({
      _raw: stringifyParagraphStyle({
        id: "EndnoteText",
        name: "endnote text",
        basedOn: "Normal",
        link: "EndnoteTextChar",
        semiHidden: true,
        uiPriority: 99,
        unhideWhenUsed: true,
        paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } },
        run: { size: 20 },
        ...options.endnoteText,
      }),
    });

    // Endnote Text Char character style
    importedStyles.push({
      _raw: stringifyCharacterStyle({
        id: "EndnoteTextChar",
        name: "Endnote Text Char",
        basedOn: "DefaultParagraphFont",
        link: "EndnoteText",
        semiHidden: true,
        run: { size: 20 },
        ...options.endnoteTextChar,
      }),
    });

    // Intense Reference character style
    importedStyles.push({
      _raw:
        `<w:style w:type="character" w:styleId="IntenseReference">` +
        `<w:name w:val="Intense Reference"/>` +
        `<w:basedOn w:val="DefaultParagraphFont"/>` +
        `<w:uiPriority w:val="32"/>` +
        `<w:qFormat/>` +
        `<w:rPr><w:b/><w:bCs/><w:smallCaps/><w:color w:val="0F4761" w:themeColor="accent1" w:themeShade="BF"/><w:spacing w:val="5"/></w:rPr>` +
        `</w:style>`,
    });

    return { importedStyles, initialAttributes };
  }
}
