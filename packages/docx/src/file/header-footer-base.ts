/**
 * Shared base class for header and footer XML components.
 *
 * @module
 */
import { InitializableXmlComponent } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import {
  buildDocumentAttributes,
  type DocumentAttributeNamespace,
} from "./document/document-attributes";
import type { FileChild } from "./file-child";

/**
 * Namespace keys used by header elements.
 * @internal
 */
export const HEADER_NAMESPACES: readonly DocumentAttributeNamespace[] = [
  "cx",
  "cx1",
  "cx2",
  "cx3",
  "cx4",
  "cx5",
  "cx6",
  "cx7",
  "cx8",
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "w16cid",
  "w16se",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
];

/**
 * Namespace keys used by footer elements.
 * @internal
 */
export const FOOTER_NAMESPACES: readonly DocumentAttributeNamespace[] = [
  "m",
  "mc",
  "o",
  "r",
  "v",
  "w",
  "w10",
  "w14",
  "w15",
  "wne",
  "wp",
  "wp14",
  "wpc",
  "wpg",
  "wpi",
  "wps",
];

/**
 * Base class for header (w:hdr) and footer (w:ftr) XML components.
 *
 * Both share the same structure: namespace attributes, reference ID,
 * and block-level child elements. This base class eliminates the
 * duplication between Header and Footer internal implementations.
 */
export abstract class HeaderFooterBase extends InitializableXmlComponent {
  private readonly refId: number;

  protected constructor(
    tagName: string,
    referenceNumber: number,
    namespaces: readonly DocumentAttributeNamespace[],
    initContent?: XmlComponent,
  ) {
    super(tagName, initContent);

    this.refId = referenceNumber;
    if (!initContent) {
      this.root.push(buildDocumentAttributes(namespaces));
    }
  }

  public get referenceId(): number {
    return this.refId;
  }

  public add(item: FileChild): void {
    this.root.push(item);
  }
}
