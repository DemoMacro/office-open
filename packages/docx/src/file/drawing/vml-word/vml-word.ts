/**
 * VML WordprocessingDrawing elements (w10: namespace).
 *
 * These elements define VML-specific word processing drawing properties
 * such as text wrapping, border decorations, and anchor locking.
 *
 * Reference: ISO/IEC 29500-4, vml-wordprocessingDrawing.xsd
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

// ── Border Elements ──

export interface VmlWordBorderOptions {
  readonly type?: string;
  readonly width?: number;
  readonly shadow?: string;
}

/**
 * Creates a w10:bordertop element (Top Border).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="bordertop" type="CT_Border"/>
 * ```
 */
export const createW10BorderTop = (options?: VmlWordBorderOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.type ? { type: { key: "type", value: options.type } } : {}),
      ...(options?.width !== undefined ? { width: { key: "width", value: options.width } } : {}),
      ...(options?.shadow ? { shadow: { key: "shadow", value: options.shadow } } : {}),
    },
    name: "w10:bordertop",
  });

/**
 * Creates a w10:borderleft element (Left Border).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="borderleft" type="CT_Border"/>
 * ```
 */
export const createW10BorderLeft = (options?: VmlWordBorderOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.type ? { type: { key: "type", value: options.type } } : {}),
      ...(options?.width !== undefined ? { width: { key: "width", value: options.width } } : {}),
      ...(options?.shadow ? { shadow: { key: "shadow", value: options.shadow } } : {}),
    },
    name: "w10:borderleft",
  });

/**
 * Creates a w10:borderright element (Right Border).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="borderright" type="CT_Border"/>
 * ```
 */
export const createW10BorderRight = (options?: VmlWordBorderOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.type ? { type: { key: "type", value: options.type } } : {}),
      ...(options?.width !== undefined ? { width: { key: "width", value: options.width } } : {}),
      ...(options?.shadow ? { shadow: { key: "shadow", value: options.shadow } } : {}),
    },
    name: "w10:borderright",
  });

/**
 * Creates a w10:borderbottom element (Bottom Border).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="borderbottom" type="CT_Border"/>
 * ```
 */
export const createW10BorderBottom = (options?: VmlWordBorderOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.type ? { type: { key: "type", value: options.type } } : {}),
      ...(options?.width !== undefined ? { width: { key: "width", value: options.width } } : {}),
      ...(options?.shadow ? { shadow: { key: "shadow", value: options.shadow } } : {}),
    },
    name: "w10:borderbottom",
  });

// ── Wrap Element ──

export interface VmlWordWrapOptions {
  readonly type?: string;
  readonly side?: string;
  readonly anchorx?: string;
  readonly anchory?: string;
}

/**
 * Creates a w10:wrap element (Text Wrapping).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="wrap" type="CT_Wrap"/>
 * ```
 */
export const createW10Wrap = (options?: VmlWordWrapOptions): XmlComponent =>
  new BuilderElement({
    attributes: {
      ...(options?.type ? { type: { key: "type", value: options.type } } : {}),
      ...(options?.side ? { side: { key: "side", value: options.side } } : {}),
      ...(options?.anchorx ? { anchorx: { key: "anchorx", value: options.anchorx } } : {}),
      ...(options?.anchory ? { anchory: { key: "anchory", value: options.anchory } } : {}),
    },
    name: "w10:wrap",
  });

// ── Anchor Lock ──

/**
 * Creates a w10:anchorlock element (Anchor Lock).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="anchorlock" type="CT_AnchorLock"/>
 * ```
 */
export const createW10AnchorLock = (): XmlComponent =>
  new BuilderElement({ name: "w10:anchorlock" });
