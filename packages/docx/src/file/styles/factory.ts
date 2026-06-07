import { buildDocumentAttributes } from "../document/document-attributes";
/**
 * Factory module for creating default document styles.
 *
 * Provides a factory class that creates pre-configured styles for common document elements.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import { DocumentDefaults } from "./defaults";
import type { DocumentDefaultsOptions } from "./defaults";
import {
  EndnoteReferenceStyle,
  EndnoteText,
  EndnoteTextChar,
  FootnoteReferenceStyle,
  FootnoteText,
  FootnoteTextChar,
  Heading1Style,
  Heading2Style,
  Heading3Style,
  Heading4Style,
  Heading5Style,
  Heading6Style,
  HyperlinkStyle,
  ListParagraph,
  StrongStyle,
  TitleStyle,
} from "./style";
import type { IBaseCharacterStyleOptions, IBaseParagraphStyleOptions } from "./style";
import type { StylesOptions } from "./styles";

/**
 * Options for configuring default document styles.
 *
 * Allows customization of built-in styles for common document elements.
 *
 * @property document - Document-wide default formatting
 * @property title - Title paragraph style options
 * @property heading1 - Heading 1 paragraph style options
 * @property heading2 - Heading 2 paragraph style options
 * @property heading3 - Heading 3 paragraph style options
 * @property heading4 - Heading 4 paragraph style options
 * @property heading5 - Heading 5 paragraph style options
 * @property heading6 - Heading 6 paragraph style options
 * @property strong - Strong paragraph style options
 * @property listParagraph - List paragraph style options
 * @property hyperlink - Hyperlink character style options
 * @property footnoteReference - Footnote reference character style options
 * @property footnoteText - Footnote text paragraph style options
 * @property footnoteTextChar - Footnote text character style options
 */
export interface DefaultStylesOptions {
  /** Document-wide default formatting */
  document?: DocumentDefaultsOptions;
  /** Title paragraph style options */
  title?: IBaseParagraphStyleOptions;
  /** Heading 1 paragraph style options */
  heading1?: IBaseParagraphStyleOptions;
  /** Heading 2 paragraph style options */
  heading2?: IBaseParagraphStyleOptions;
  /** Heading 3 paragraph style options */
  heading3?: IBaseParagraphStyleOptions;
  /** Heading 4 paragraph style options */
  heading4?: IBaseParagraphStyleOptions;
  /** Heading 5 paragraph style options */
  heading5?: IBaseParagraphStyleOptions;
  /** Heading 6 paragraph style options */
  heading6?: IBaseParagraphStyleOptions;
  /** Strong paragraph style options */
  strong?: IBaseParagraphStyleOptions;
  /** List paragraph style options */
  listParagraph?: IBaseParagraphStyleOptions;
  /** Hyperlink character style options */
  hyperlink?: IBaseCharacterStyleOptions;
  /** Footnote reference character style options */
  footnoteReference?: IBaseCharacterStyleOptions;
  /** Footnote text paragraph style options */
  footnoteText?: IBaseParagraphStyleOptions;
  /** Footnote text character style options */
  footnoteTextChar?: IBaseCharacterStyleOptions;
  endnoteReference?: IBaseCharacterStyleOptions;
  endnoteText?: IBaseParagraphStyleOptions;
  endnoteTextChar?: IBaseCharacterStyleOptions;
}

let cachedDefaultStyles: StylesOptions | null = null;

/**
 * Factory for creating default document styles.
 *
 * This factory creates a complete set of default styles for common document elements
 * such as headings, hyperlinks, and footnotes. When no custom options are provided,
 * a cached instance is reused across documents.
 *
 * @example
 * ```typescript
 * // Create default styles with custom heading formatting
 * const factory = new DefaultStylesFactory();
 * const styles = factory.newInstance({
 *   heading1: {
 *     run: { color: "FF0000", size: 32 }
 *   },
 *   heading2: {
 *     run: { color: "00FF00", size: 26 }
 *   }
 * });
 * ```
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
    return {
      importedStyles: [
        new DocumentDefaults(options.document ?? {}),
        new TitleStyle({
          run: {
            size: 56,
          },
          ...options.title,
        }),
        new Heading1Style(options.heading1 || {}),
        new Heading2Style(options.heading2 || {}),
        new Heading3Style(options.heading3 || {}),
        new Heading4Style(options.heading4 || {}),
        new Heading5Style(options.heading5 || {}),
        new Heading6Style(options.heading6 || {}),
        new StrongStyle({
          run: {
            bold: true,
          },
          ...options.strong,
        }),
        new ListParagraph(options.listParagraph || {}),
        new HyperlinkStyle(options.hyperlink || {}),
        new FootnoteReferenceStyle(options.footnoteReference || {}),
        new FootnoteText(options.footnoteText || {}),
        new FootnoteTextChar(options.footnoteTextChar || {}),
        new EndnoteReferenceStyle(options.endnoteReference || {}),
        new EndnoteText(options.endnoteText || {}),
        new EndnoteTextChar(options.endnoteTextChar || {}),
      ],
      initialStyles: buildDocumentAttributes(["mc", "r", "w", "w14", "w15"], "w14 w15"),
    };
  }
}
