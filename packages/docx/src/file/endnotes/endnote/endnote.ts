import type { FileChild } from "@file/file-child";
import { XmlComponent } from "@file/xml-components";

import { EndnoteRefRun } from "./run/endnote-ref-run";

export const EndnoteType = {
  CONTINUATION_SEPARATOR: "continuationSeparator",

  SEPARATOR: "separator",
} as const;

export interface EndnoteOptions {
  readonly id: number;
  readonly type?: (typeof EndnoteType)[keyof typeof EndnoteType];
  readonly children: readonly FileChild[];
}

export class Endnote extends XmlComponent {
  public constructor(options: EndnoteOptions) {
    super("w:endnote");
    const attr: Record<string, string | number> = { "w:id": options.id };
    if (options.type !== undefined) {
      attr["w:type"] = options.type;
    }
    this.root.push({ _attr: attr });

    for (let i = 0; i < options.children.length; i++) {
      const child = options.children[i];

      if (i === 0 && "addRunToFront" in child) {
        (child as any).addRunToFront(new EndnoteRefRun());
      }

      this.root.push(child);
    }
  }
}
