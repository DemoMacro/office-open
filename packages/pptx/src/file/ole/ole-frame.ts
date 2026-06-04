/**
 * OLE Object frame for PresentationML — p:graphicFrame with p:oleObj.
 *
 * Embeds an OLE object (e.g., Excel worksheet, Word document) on a slide.
 *
 * Reference: OOXML transitional, pml.xsd, CT_OleObject
 *
 * @module
 */
import { Transform2D } from "@file/drawingml/transform-2d";
import { BuilderElement, buildAttrObject, type Context, XmlComponent } from "@file/xml-components";
import { emuPosition } from "@util/position";

let nextOleId = 2048;

// ── Options ──

export interface OleEmbedOptions {
  /** Relationship ID for the embedded OLE data */
  readonly rId: string;
  /** Last update in document (ISO 8601) */
  readonly lastEdited?: string;
}

export interface OleLinkOptions {
  /** Relationship ID for the linked OLE data */
  readonly rId: string;
  /** Last update in document (ISO 8601) */
  readonly updateLastEdited?: string;
  /** Automatic or manual update */
  readonly autoUpdate?: boolean;
}

export interface OleOptions {
  /** Position and size */
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  readonly progId?: string;
  /** Shape ID */
  readonly spid?: string;
  /** Object name */
  readonly name?: string;
  /** Show as icon */
  readonly showAsIcon?: boolean;
  /** Image width (EMU) for icon/preview */
  readonly imgW?: number;
  /** Image height (EMU) for icon/preview */
  readonly imgH?: number;
  /** Embed mode (provides rId for embedded OLE) */
  readonly embed?: OleEmbedOptions;
  /** Link mode (provides rId for linked OLE) */
  readonly link?: OleLinkOptions;
  /** Relationship ID for the preview/icon image */
  readonly imgRId?: string;
}

// ── Component ──

/**
 * p:graphicFrame — Slide-level graphic frame wrapping an OLE object.
 *
 * Produces the AlternateContent wrapper needed by PowerPoint to render OLE objects.
 */
export class OleFrame extends XmlComponent {
  private readonly options: OleOptions;

  public constructor(options: OleOptions) {
    super("p:graphicFrame");
    this.options = options;

    const id = nextOleId++;
    const name = options.name ?? `Object ${id}`;

    this.root.push(new OleGraphicFrameNonVisual(id, name));
    this.root.push(
      new Transform2D(
        {
          ...emuPosition(options),
        },
        "p",
      ),
    );
    this.root.push(new OleGraphic(options));
  }

  public override toXml(context: Context): string {
    // Register OLE relationship in context if needed
    const file = context.fileData as { oleObjects?: OleData[] };
    if (file?.oleObjects) {
      file.oleObjects.push({
        key: `ole_${this.options.spid ?? "1"}`,
        rId: this.options.embed?.rId ?? this.options.link?.rId ?? "",
        progId: this.options.progId,
      });
    }
    return super.toXml(context);
  }
}

export interface OleData {
  readonly key: string;
  readonly rId: string;
  readonly progId?: string;
}

// ── Sub-components ──

class OleGraphicFrameNonVisual extends XmlComponent {
  public constructor(id: number, name: string) {
    super("p:nvGraphicFramePr");
    this.root.push(
      new BuilderElement({
        name: "p:cNvPr",
        attributes: {
          id: { key: "id", value: id },
          name: { key: "name", value: name },
        },
      }),
    );
    this.root.push(
      new BuilderElement({
        name: "p:cNvGraphicFramePr",
        children: [
          new BuilderElement({
            name: "a:graphicFrameLocks",
            attributes: { noGrp: { key: "noGrp", value: 1 } },
          }),
        ],
      }),
    );
    this.root.push(new BuilderElement({ name: "p:nvPr" }));
  }
}

class OleGraphic extends XmlComponent {
  public constructor(options: OleOptions) {
    super("a:graphic");
    this.root.push(new OleGraphicData(options));
  }
}

class OleGraphicData extends XmlComponent {
  public constructor(options: OleOptions) {
    super("a:graphicData");
    this.root.push(
      buildAttrObject({
        uri: "http://schemas.openxmlformats.org/presentationml/2006/ole",
      }),
    );
    this.root.push(new OleObjectElement(options));
  }
}

class OleObjectElement extends XmlComponent {
  public constructor(options: OleOptions) {
    super("p:oleObj");
    const attrs: Record<string, string | number | boolean> = {
      name: options.name ?? "OLE Object",
    };
    if (options.spid) attrs.spid = options.spid;
    if (options.showAsIcon) attrs.showAsIcon = 1;
    if (options.imgW) attrs.imgW = options.imgW;
    if (options.imgH) attrs.imgH = options.imgH;
    if (options.progId) attrs.progId = options.progId;
    const rId = options.embed?.rId ?? options.link?.rId;
    if (rId) attrs["r:id"] = rId;
    this.root.push(buildAttrObject(attrs));

    // embed or link child
    if (options.embed) {
      this.root.push(
        new BuilderElement({
          name: "p:embed",
        }),
      );
    } else if (options.link) {
      this.root.push(
        new BuilderElement({
          name: "p:link",
          attributes: options.link.autoUpdate
            ? { updateAutomatic: { key: "updateAutomatic", value: 1 } }
            : undefined,
        }),
      );
    }

    // Optional pic for icon/preview
    if (options.imgRId) {
      this.root.push(new OlePic(options));
    }
  }
}

class OlePic extends XmlComponent {
  public constructor(options: OleOptions) {
    super("p:pic");
    // Minimal pic for OLE preview
    this.root.push(
      new BuilderElement({
        name: "p:nvPicPr",
        children: [
          new BuilderElement({
            name: "p:cNvPr",
            attributes: {
              id: { key: "id", value: 2 },
              name: { key: "name", value: options.name ?? "OLE Preview" },
            },
          }),
          new BuilderElement({ name: "p:cNvPicPr" }),
          new BuilderElement({ name: "p:nvPr" }),
        ],
      }),
    );
  }
}
