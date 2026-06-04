import type { FileChild } from "@file/file-child";
import type { VmlShapeStyle } from "@file/textbox/shape/shape";
/**
 * Object element for WordprocessingML documents — w:object.
 *
 * Embeds an OLE object in a paragraph using VML shape + objectEmbed/objectLink.
 *
 * Reference: OOXML transitional, wml.xsd, CT_Object / CT_ObjectEmbed / CT_ObjectLink
 *
 * @module
 */
import { BuilderElement, XmlComponent } from "@file/xml-components";

// ── Options ──

export interface ObjectEmbedOptions {
  /** Relationship ID for the embedded object */
  readonly rId: string;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  readonly progId?: string;
  /** Draw aspect ("content" | "icon") */
  readonly drawAspect?: string;
  /** Shape ID */
  readonly shapeId?: string;
  /** Field codes */
  readonly fieldCodes?: string;
}

export interface ObjectLinkOptions extends ObjectEmbedOptions {
  /** Update mode ("always" | "onCall") */
  readonly updateMode: string;
  /** Locked field */
  readonly lockedField?: boolean;
}

export interface ObjectElementOptions {
  /** VML shape style */
  readonly style?: VmlShapeStyle;
  /** Shape ID */
  readonly shapeId?: string;
  /** Original width in twips */
  readonly dxaOrig?: number;
  /** Original height in twips */
  readonly dyaOrig?: number;
  /** Embedded OLE object */
  readonly embed?: ObjectEmbedOptions;
  /** Linked OLE object */
  readonly link?: ObjectLinkOptions;
}

// ── Style formatting ──

const STYLE_MAP: Record<string, string> = {
  width: "width",
  height: "height",
  position: "position",
  left: "left",
  top: "top",
};

function formatStyle(style?: VmlShapeStyle): string | undefined {
  if (!style) return undefined;
  return Object.entries(style)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `${STYLE_MAP[k] ?? k}:${v}`)
    .join(";");
}

// ── Components ──

class ObjectEmbed extends XmlComponent {
  public constructor(options: ObjectEmbedOptions) {
    super("w:objectEmbed");
    const attrs: Record<string, string> = { "r:id": options.rId };
    if (options.progId) attrs["w:progId"] = options.progId;
    if (options.drawAspect) attrs["w:drawAspect"] = options.drawAspect;
    if (options.shapeId) attrs["w:shapeId"] = options.shapeId;
    if (options.fieldCodes) attrs["w:fieldCodes"] = options.fieldCodes;
    this.root.push({ _attr: attrs });
  }
}

class ObjectLink extends XmlComponent {
  public constructor(options: ObjectLinkOptions) {
    super("w:objectLink");
    const attrs: Record<string, string> = {
      "r:id": options.rId,
      "w:updateMode": options.updateMode,
    };
    if (options.progId) attrs["w:progId"] = options.progId;
    if (options.drawAspect) attrs["w:drawAspect"] = options.drawAspect;
    if (options.shapeId) attrs["w:shapeId"] = options.shapeId;
    if (options.fieldCodes) attrs["w:fieldCodes"] = options.fieldCodes;
    if (options.lockedField) attrs["w:lockedField"] = "1";
    this.root.push({ _attr: attrs });
  }
}

/** Creates a w:object element with VML shape and optional embed/link */
export class ObjectElement extends XmlComponent implements FileChild {
  public readonly fileChild = Symbol();

  public constructor(options: ObjectElementOptions) {
    super("w:object");

    // Attributes
    const attrs: Record<string, number> = {};
    if (options.dxaOrig !== undefined) attrs["w:dxaOrig"] = options.dxaOrig;
    if (options.dyaOrig !== undefined) attrs["w:dyaOrig"] = options.dyaOrig;
    if (Object.keys(attrs).length > 0) {
      this.root.push({ _attr: attrs });
    }

    // VML shape (required by the XSD — any namespace urn:schemas-microsoft-com:vml)
    const shapeAttrs: Record<string, string> = {
      id: options.shapeId ?? `_x0000_i1025`,
      type: "#_x0000_t75",
    };
    const styleStr = formatStyle(options.style);
    if (styleStr) shapeAttrs.style = styleStr;
    this.root.push(
      new BuilderElement({
        name: "v:shape",
        attributes: Object.fromEntries(
          Object.entries(shapeAttrs).map(([key, value]) => [key, { key, value }]),
        ),
        children: [],
      }),
    );

    // Office element (o:OLEObject)
    this.root.push(
      new BuilderElement({
        name: "o:OLEObject",
        attributes: {
          Type: { key: "Type", value: options.link ? "Link" : "Embed" },
          ProgID: {
            key: "ProgID",
            value: options.embed?.progId ?? options.link?.progId,
          },
          ShapeID: {
            key: "ShapeID",
            value: options.shapeId ?? `_x0000_i1025`,
          },
          DrawAspect: {
            key: "DrawAspect",
            value: options.embed?.drawAspect ?? options.link?.drawAspect ?? "Content",
          },
          ObjectID: {
            key: "ObjectID",
            value: `_${generateId()}`,
          },
        },
        children: [],
      }),
    );

    // Embed or Link
    if (options.embed) {
      this.root.push(new ObjectEmbed(options.embed));
    } else if (options.link) {
      this.root.push(new ObjectLink(options.link));
    }
  }
}

function generateId(): string {
  return Math.random().toString(16).slice(2, 18).toUpperCase().padEnd(16, "0");
}
