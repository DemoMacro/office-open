/**
 * Frameset and Frame XML generators for webSettings.xml.
 *
 * Framesets define a frames-based layout within a Word document (web layout).
 *
 * Reference: OOXML transitional, wml.xsd, CT_Frameset / CT_Frame
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

// ── Options ──

export interface FrameOptions {
  /** Frame size (e.g., "50%") */
  readonly size?: string;
  /** Frame name */
  readonly name?: string;
  /** Frame title */
  readonly title?: string;
  /** Source file link rId */
  readonly sourceRId?: string;
  /** Margin width in pixels */
  readonly marginWidth?: number;
  /** Margin height in pixels */
  readonly marginHeight?: number;
  /** Scrollbar mode */
  readonly scrollbar?: "on" | "off" | "auto";
  /** No resize allowed */
  readonly noResizeAllowed?: boolean;
  /** Linked to file */
  readonly linkedToFile?: boolean;
}

export interface FramesetSplitbarOptions {
  /** Splitbar width in twips */
  readonly width?: number;
  /** Splitbar color (hex, e.g., "auto") */
  readonly color?: string;
  /** No border */
  readonly noBorder?: boolean;
  /** Flat borders */
  readonly flatBorders?: boolean;
}

export interface FramesetOptions {
  /** Frameset size (e.g., "100%") */
  readonly size?: string;
  /** Splitbar configuration */
  readonly splitbar?: FramesetSplitbarOptions;
  /** Layout direction */
  readonly layout?: "rows" | "cols";
  /** Frameset title */
  readonly title?: string;
  /** Child framesets and frames (in order) */
  readonly children?: readonly (FramesetOptions | FrameOptions)[];
}

// ── Components ──

class FrameScrollbar extends XmlComponent {
  public constructor(val: "on" | "off" | "auto") {
    super("w:scrollbar");
    this.root.push({ _attr: { "w:val": val } });
  }
}

class FrameElement extends XmlComponent {
  public constructor(options: FrameOptions) {
    super("w:frame");
    if (options.size !== undefined) {
      this.root.push({ _attr: {} });
      this.root.push(new StringElement("w:sz", options.size));
    }
    if (options.name !== undefined) {
      this.root.push(new StringElement("w:name", options.name));
    }
    if (options.title !== undefined) {
      this.root.push(new StringElement("w:title", options.title));
    }
    if (options.sourceRId !== undefined) {
      this.root.push(new RelElement("w:sourceFileName", options.sourceRId));
    }
    if (options.marginWidth !== undefined) {
      this.root.push(new PixelsElement("w:marW", options.marginWidth));
    }
    if (options.marginHeight !== undefined) {
      this.root.push(new PixelsElement("w:marH", options.marginHeight));
    }
    if (options.scrollbar !== undefined) {
      this.root.push(new FrameScrollbar(options.scrollbar));
    }
    if (options.noResizeAllowed) {
      this.root.push(new OnOffElement("w:noResizeAllowed"));
    }
    if (options.linkedToFile) {
      this.root.push(new OnOffElement("w:linkedToFile"));
    }
  }
}

class FramesetSplitbar extends XmlComponent {
  public constructor(options: FramesetSplitbarOptions) {
    super("w:framesetSplitbar");
    if (options.width !== undefined) {
      this.root.push(new TwipsElement("w:w", options.width));
    }
    if (options.color !== undefined) {
      this.root.push(new StringElement("w:color", options.color));
    }
    if (options.noBorder) {
      this.root.push(new OnOffElement("w:noBorder"));
    }
    if (options.flatBorders) {
      this.root.push(new OnOffElement("w:flatBorders"));
    }
  }
}

class FrameLayout extends XmlComponent {
  public constructor(val: "rows" | "cols") {
    super("w:frameLayout");
    this.root.push({ _attr: { "w:val": val } });
  }
}

/** Helper to build a Frameset (can be nested) */
export class Frameset extends XmlComponent {
  public constructor(options: FramesetOptions) {
    super("w:frameset");
    if (options.size !== undefined) {
      this.root.push(new StringElement("w:sz", options.size));
    }
    if (options.splitbar) {
      this.root.push(new FramesetSplitbar(options.splitbar));
    }
    if (options.layout !== undefined) {
      this.root.push(new FrameLayout(options.layout));
    }
    if (options.title !== undefined) {
      this.root.push(new StringElement("w:title", options.title));
    }
    if (options.children) {
      for (const child of options.children) {
        if ("children" in child || "layout" in child || "splitbar" in child) {
          this.root.push(new Frameset(child as FramesetOptions));
        } else {
          this.root.push(new FrameElement(child as FrameOptions));
        }
      }
    }
  }
}

// ── Internal helpers ──

class StringElement extends XmlComponent {
  public constructor(tag: string, val: string) {
    super(tag);
    this.root.push({ _attr: { "w:val": val } });
  }
}

class RelElement extends XmlComponent {
  public constructor(tag: string, rId: string) {
    super(tag);
    this.root.push({ _attr: { "r:id": rId } });
  }
}

class PixelsElement extends XmlComponent {
  public constructor(tag: string, val: number) {
    super(tag);
    this.root.push({ _attr: { "w:val": val } });
  }
}

class TwipsElement extends XmlComponent {
  public constructor(tag: string, val: number) {
    super(tag);
    this.root.push({ _attr: { "w:w": val, "w:type": "dxa" } });
  }
}

class OnOffElement extends XmlComponent {
  public constructor(tag: string) {
    super(tag);
  }
}
