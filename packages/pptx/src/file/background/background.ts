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
}

/**
 * p:bg — Slide background.
 */
export class Background extends XmlComponent {
  public constructor(options: BackgroundOptions = {}) {
    super("p:bg");
    this.root.push(new BackgroundProperties(options));
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
