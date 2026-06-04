/**
 * Locked Canvas module for DrawingML.
 *
 * The locked canvas element (lc:lockedCanvas) represents a group of shapes
 * that cannot be individually selected or edited by the user.
 * It uses the same structure as a group shape (CT_GvmlGroupShape).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-lc_lockedCanvas-1.html
 *
 * @module
 */
import { XmlComponent } from "../../xml-components";
import type { IXmlableObject } from "../../xml-components";

/**
 * Options for creating a locked canvas element.
 */
export interface LockedCanvasOptions {
  /** Child elements of the locked canvas */
  readonly children?: readonly (IXmlableObject | XmlComponent)[];
}

/**
 * Represents a locked canvas element (lc:lockedCanvas).
 *
 * A locked canvas contains shapes that are rendered but cannot be
 * individually selected, moved, or edited.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-lc_lockedCanvas-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="lockedCanvas" type="a:CT_GvmlGroupShape"/>
 * ```
 */
export class LockedCanvas extends XmlComponent {
  public constructor(options?: LockedCanvasOptions) {
    super("lc:lockedCanvas");
    if (options?.children) {
      for (const child of options.children) {
        this.root.push(child);
      }
    }
  }
}
