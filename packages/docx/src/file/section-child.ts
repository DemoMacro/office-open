/**
 * SectionChild type definition for JSON-friendly API.
 *
 * Discriminated union for section/header/footer children.
 * Accepts both class instances (backward compat) and plain objects (JSON-friendly).
 *
 * @module
 */
import type { AltChunkOptions } from "./alt-chunk/alt-chunk";
import type { CustomXmlBlockOptions } from "./custom-xml";
import type { ParagraphOptions } from "./paragraph/paragraph";
import type { SubDocOptions } from "./sub-doc/sub-doc";
import type { SdtPropertiesOptions } from "./table-of-contents";
import type { TableOfContentsOptions } from "./table-of-contents/table-of-contents-properties";
import type { TableOptions } from "./table/table";
import type { VmlShapeStyle } from "./textbox/shape/shape";
import type { BaseXmlComponent } from "./xml-components";

/**
 * Discriminated union for section body, header, and footer children.
 *
 * Each plain-object variant uses a single-key discriminator so that
 * `coerceSectionChild` can map it to the correct class constructor.
 *
 * - Class instances (`BaseXmlComponent`) pass through unchanged (backward compat).
 * - `{ paragraph: … }` → `Paragraph`
 * - `{ table: … }`      → `Table`
 * - `{ toc: … }`         → `TableOfContents`
 * - `{ textbox: … }`     → `Textbox`
 * - `{ sdt: … }`         → `StructuredDocumentTagBlock`
 * - `{ altChunk: … }`    → `AltChunk`
 * - `{ subDoc: … }`      → `SubDoc`
 * - `{ customXml: … }`   → `CustomXmlBlock`
 */
export type SectionChild =
  | BaseXmlComponent
  | { paragraph: string | ParagraphOptions }
  | { table: TableOptions }
  | { toc: TableOfContentsOptions & { readonly alias?: string } }
  | {
      textbox: Omit<ParagraphOptions, "style" | "children"> & {
        readonly style?: VmlShapeStyle;
        readonly children?: readonly SectionChild[];
      };
    }
  | {
      sdt: {
        readonly properties: SdtPropertiesOptions;
        readonly children?: readonly SectionChild[];
      };
    }
  | { altChunk: AltChunkOptions }
  | { subDoc: SubDocOptions }
  | { customXml: CustomXmlBlockOptions };
