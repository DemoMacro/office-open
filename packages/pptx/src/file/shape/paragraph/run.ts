import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

import { RunProperties } from "./run-properties";
import type { RunPropertiesOptions } from "./run-properties";

export interface RunOptions extends RunPropertiesOptions {
  readonly text?: string;
}

/**
 * a:r — A run of text with properties.
 */
export class TextRun extends XmlComponent {
  private readonly options: RunOptions;

  public constructor(options: RunOptions | string = {}) {
    super("a:r");
    this.options = typeof options === "string" ? { text: options } : options;
  }

  public override toXml(context: Context): string {
    const body = new RunProperties(this.options).toXml(context);
    if (this.options.text) {
      return `<a:r>${body}<a:t>${escapeXml(this.options.text)}</a:t></a:r>`;
    }
    return body ? `<a:r>${body}</a:r>` : "<a:r/>";
  }
}
