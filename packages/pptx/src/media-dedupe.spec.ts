import { unzipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentationSync } from "./generate";

const image = new Uint8Array([1, 2, 3, 4]);

describe("PPTX media deduplication", () => {
  it("reuses one package entry for repeated fresh pictures with identical bytes", () => {
    const output = generatePresentationSync(
      {
        slides: [
          {
            children: [0, 1].map((id) => ({
              picture: {
                id,
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                data: image,
                type: "png",
              },
            })),
          },
        ],
      },
      { type: "uint8array" },
    );

    const media = Object.keys(unzipSync(output)).filter((path) => path.startsWith("ppt/media/"));
    expect(media).toEqual(["ppt/media/image1.png"]);
  });
});
