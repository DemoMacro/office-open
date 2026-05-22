/**
 * Coerce function for SectionChild type.
 *
 * Converts plain-object representations into their corresponding class
 * instances, enabling a JSON-friendly API.
 *
 * @module
 */
import { AltChunk } from "./alt-chunk/alt-chunk";
import type { FileChild } from "./file-child";
import { Paragraph } from "./paragraph/paragraph";
import { StructuredDocumentTagBlock } from "./sdt/sdt";
import type { SectionChild } from "./section-child";
import { SubDoc } from "./sub-doc/sub-doc";
import { TableOfContents } from "./table-of-contents/table-of-contents";
import { Table } from "./table/table";
import { Textbox } from "./textbox/textbox";
import { BaseXmlComponent } from "./xml-components";

/**
 * Coerces a `SectionChild` (plain object or class instance) into a `FileChild`.
 *
 * - Class instances pass through unchanged.
 * - Plain objects are mapped to their corresponding constructor.
 */
export function coerceSectionChild(child: SectionChild): FileChild {
  if (child instanceof BaseXmlComponent) return child as FileChild;
  if ("paragraph" in child) return new Paragraph(child.paragraph);
  if ("table" in child) return new Table(child.table);
  if ("toc" in child) {
    const { alias, ...options } = child.toc;
    return new TableOfContents(alias, options);
  }
  if ("textbox" in child) return new Textbox(child.textbox);
  if ("sdt" in child) return new StructuredDocumentTagBlock(child.sdt);
  if ("altChunk" in child) return new AltChunk(child.altChunk);
  if ("subDoc" in child) return new SubDoc(child.subDoc);
  throw new Error("Unknown section child type");
}
