import type { EffectsOptions } from "@file/drawingml/effects";
import { createPptxEffectList } from "@file/drawingml/effects";
import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

export interface BackgroundOptions {
  readonly fill?: FillOptions;
  readonly effects?: EffectsOptions;
  readonly shadeToTitle?: boolean;
  readonly blackWhiteMode?:
    | "clr"
    | "gray"
    | "ltGray"
    | "invGray"
    | "gmGray"
    | "bw"
    | "auto"
    | "black"
    | "white";
}

/**
 * p:bg — Slide background.
 */
export class Background extends XmlComponent {
  private readonly blackWhiteMode?: string;

  public constructor(options: BackgroundOptions = {}) {
    super("p:bg");
    this.blackWhiteMode = options.blackWhiteMode;
    this.root.push(new BackgroundProperties(options));
  }

  public override toXml(context: Context): string {
    if (this.blackWhiteMode) {
      this.root.unshift({ _attr: { "p:bwMode": this.blackWhiteMode } });
      const result = super.toXml(context);
      this.root.shift();
      return result;
    }
    return super.toXml(context);
  }
}

class BackgroundProperties extends XmlComponent {
  private readonly shadeToTitle: boolean;

  public constructor(options: BackgroundOptions) {
    super("p:bgPr");
    this.shadeToTitle = !!options.shadeToTitle;
    this.root.push(
      options.fill !== undefined ? buildFill(options.fill) : buildFill({ type: "none" }),
    );
    if (options.effects) {
      const el = createPptxEffectList(options.effects);
      if (el) this.root.push(el);
    }
  }

  public override toXml(context: Context): string {
    if (this.shadeToTitle) {
      this.root.unshift({ _attr: { shadeToTitle: 1 } });
      const result = super.toXml(context);
      this.root.shift();
      return result;
    }
    return super.toXml(context);
  }
}
