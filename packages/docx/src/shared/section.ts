import type { AltChunkOptions } from "@parts/alt-chunk/alt-chunk";
import type { CustomXmlBlockOptions } from "@parts/custom-xml";
/**
 * Section types for WordprocessingML documents.
 *
 * SectionChild — discriminated union for section/header/footer children.
 * SectionOptions — options for a document section.
 *
 * @module
 */
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties";
import type { MarkupRangeOptions, BookmarkStartOptions } from "@parts/paragraph/links/bookmark";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { SubDocOptions } from "@parts/sub-doc/sub-doc";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
import type { TableOfContentsOptions } from "@parts/table-of-contents/table-of-contents-properties";
import type { TableOptions } from "@parts/table/table";
import type { VmlShapeStyle } from "@parts/textbox/shape/shape";

/**
 * Discriminated union for section body, header, and footer children.
 *
 * Each variant uses a single-key discriminator:
 * - `{ paragraph: … }` → paragraph
 * - `{ table: … }`      → table
 * - `{ toc: … }`         → table of contents
 * - `{ textbox: … }`     → textbox
 * - `{ sdt: … }`         → structured document tag
 * - `{ altChunk: … }`    → alt chunk
 * - `{ subDoc: … }`      → sub document
 * - `{ customXml: … }`   → custom XML
 * - `{ bookmarkStart/End }` → body-level range markers (between paragraphs)
 */
export type SectionChild =
  | { paragraph: string | ParagraphOptions }
  | { table: TableOptions }
  | { toc: TableOfContentsOptions & { alias?: string } }
  | {
      textbox: Omit<ParagraphOptions, "style" | "children"> & {
        style?: VmlShapeStyle;
        children?: SectionChild[];
      };
    }
  | {
      sdt: {
        properties: SdtPropertiesOptions;
        children?: SectionChild[];
      };
    }
  | { altChunk: AltChunkOptions }
  | { subDoc: SubDocOptions }
  | { customXml: CustomXmlBlockOptions }
  | { bookmarkStart: BookmarkStartOptions }
  | { bookmarkEnd: MarkupRangeOptions }
  | { rawXml: string };

/**
 * Options for a document section.
 *
 * Each section can have its own headers, footers, and page properties.
 */
export interface SectionOptions {
  headers?: {
    default?: SectionChild[];
    first?: SectionChild[];
    even?: SectionChild[];
  };
  footers?: {
    default?: SectionChild[];
    first?: SectionChild[];
    even?: SectionChild[];
  };
  properties?: SectionPropertiesOptions;
  children: SectionChild[];
}
