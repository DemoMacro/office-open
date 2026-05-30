import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml, xml } from "@office-open/xml";

import type { RunPropertiesOptions } from "./run-properties";
import { RunProperties, buildRunProperties } from "./run-properties";

export interface RunOptions extends RunPropertiesOptions {
  readonly text?: string;
}

/**
 * a:r — A run of text with properties.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class TextRun extends XmlComponent {
  private readonly options: RunOptions;

  public constructor(options: RunOptions | string = {}) {
    super("a:r");
    this.options = typeof options === "string" ? { text: options } : options;
  }

  /**
   * Fast path: simple properties (no hyperlink/fill/shadow/outline) skip
   * RunProperties side effects and serialize directly.
   * Complex path uses RunProperties.toXml() + direct text serialization.
   */
  public override toXml(context: Context): string {
    const opts = this.options;
    const hasRPr = RunProperties.hasProperties(opts);

    // Simple path: no side-effect-requiring properties
    if (!hasRPr || (!opts.hyperlink && !opts.fill && !opts.shadow && !opts.outline)) {
      let body = "";
      if (hasRPr) {
        const rPrObj = buildRunProperties(opts);
        if (rPrObj) body += xml(rPrObj);
      }
      if (opts.text) {
        body += `<a:t>${escapeXml(opts.text)}</a:t>`;
      }
      return body.length === 0 ? "<a:r/>" : `<a:r>${body}</a:r>`;
    }

    // Complex path: use RunProperties.toXml() for side effects (hyperlink registration etc.)
    let body = new RunProperties(opts).toXml(context);
    if (opts.text) {
      body += `<a:t>${escapeXml(opts.text)}</a:t>`;
    }
    return body ? `<a:r>${body}</a:r>` : "<a:r/>";
  }
}
