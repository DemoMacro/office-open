import { describe, expect, it } from "vite-plus/test";

import { compileDocument } from "../compiler";
import type { DocumentOptions } from "./core-properties";

describe("picture drawing wrappers", () => {
  it("preserves pic:spPr bwMode in nested drawing wrappers", () => {
    const picture = {
      type: "png" as const,
      data: new Uint8Array([1, 2, 3]),
      transformation: { width: 9525, height: 9525 },
      blackWhiteMode: "gray" as const,
    };
    const options: DocumentOptions = {
      sections: [
        {
          headers: { default: [{ paragraph: { children: [{ picture }] } }] },
          footers: { default: [{ paragraph: { children: [{ picture }] } }] },
          children: [
            {
              sdt: {
                properties: { richText: true },
                children: [{ paragraph: { children: [{ picture }] } }],
              },
            },
            { paragraph: { children: [{ picture }] } },
          ],
        },
      ],
    };

    const files = compileDocument(options);
    const xml = (name: string) => new TextDecoder().decode(files[name] as Uint8Array);

    for (const [name, count] of [
      ["word/document.xml", 2],
      ["word/header1.xml", 1],
      ["word/footer1.xml", 1],
    ] as const) {
      expect(xml(name).match(/bwMode="gray"/g)).toHaveLength(count);
    }
  });
});
