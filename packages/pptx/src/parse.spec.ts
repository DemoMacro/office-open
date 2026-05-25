import { Packer } from "@export/packer/packer";
import { describe, expect, it } from "vite-plus/test";

import { File as Presentation } from "./file/file";
import type { IMasterDefinition, ISlideOptions } from "./file/file";
import { parsePresentation } from "./parse";

describe("parsePresentation", () => {
  it("returns IParsedPresentation with slides", async () => {
    const pres = new Presentation({
      slides: [
        {
          children: [
            { shape: { x: 100, y: 100, width: 600, height: 400, text: "Hello", fill: "4472C4" } },
          ],
        },
      ],
    });
    const buffer = await Packer.toBuffer(pres);
    const result = parsePresentation(new Uint8Array(buffer));

    expect(result.slides).to.exist;
    expect(result.slides.length).to.equal(1);
    expect(result.masters).to.be.undefined;
  });

  it("parses single master file with undefined masters", async () => {
    const pres = new Presentation({
      slides: [
        { children: [{ shape: { x: 0, y: 0, width: 200, height: 100, text: "A" } }] },
        { children: [{ shape: { x: 0, y: 0, width: 200, height: 100, text: "B" } }] },
      ],
    });
    const buffer = await Packer.toBuffer(pres);
    const result = parsePresentation(new Uint8Array(buffer));

    expect(result.slides.length).to.equal(2);
    expect(result.masters).to.be.undefined;
  });

  it("parses multi-master file", async () => {
    const masters: IMasterDefinition[] = [
      {
        name: "light",
        theme: { name: "Light", colors: { dark1: "333333" } },
      },
      {
        name: "dark",
        theme: { name: "Dark", colors: { dark1: "FFFFFF" } },
      },
    ];

    const slides: ISlideOptions[] = [
      {
        master: "light",
        layout: "blank",
        children: [{ shape: { x: 100, y: 100, width: 600, height: 400, text: "Light slide" } }],
      },
      {
        master: "dark",
        layout: "blank",
        children: [
          {
            shape: {
              x: 100,
              y: 100,
              width: 600,
              height: 400,
              text: "Dark slide",
              paragraphs: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: "Dark slide", fill: "FFFFFF" }],
                },
              ],
            },
          },
        ],
      },
    ];

    const pres = new Presentation({ title: "Multi-master", masters, slides });
    const buffer = await Packer.toBuffer(pres);
    const result = parsePresentation(new Uint8Array(buffer));

    expect(result.slides.length).to.equal(2);
    expect(result.masters).to.exist;
    expect(result.masters!.length).to.equal(2);
    // Master name is derived from theme name
    expect(result.slides[0].master).to.equal("Light");
    expect(result.slides[1].master).to.equal("Dark");
    expect(result.masters![0].name).to.equal("Light");
    expect(result.masters![1].name).to.equal("Dark");
    expect(result.masters![0].theme?.name).to.equal("Light");
    expect(result.masters![1].theme?.name).to.equal("Dark");
  });

  it("round-trips multi-master structure", async () => {
    const masters: IMasterDefinition[] = [
      { name: "m1", theme: { name: "Theme One" } },
      { name: "m2", theme: { name: "Theme Two" } },
    ];
    const slides: ISlideOptions[] = [
      {
        master: "m1",
        children: [{ shape: { x: 50, y: 50, width: 400, height: 300, text: "S1" } }],
      },
      {
        master: "m2",
        children: [{ shape: { x: 50, y: 50, width: 400, height: 300, text: "S2" } }],
      },
    ];

    const pres = new Presentation({ masters, slides });
    const buffer = await Packer.toBuffer(pres);

    // First parse
    const parsed1 = parsePresentation(new Uint8Array(buffer));
    expect(parsed1.masters!.length).to.equal(2);

    // Re-generate from parsed data
    const pres2 = new Presentation({
      masters: parsed1.masters as IMasterDefinition[],
      slides: parsed1.slides as ISlideOptions[],
    });
    const buffer2 = await Packer.toBuffer(pres2);

    // Second parse
    const parsed2 = parsePresentation(new Uint8Array(buffer2));
    expect(parsed2.slides.length).to.equal(2);
    expect(parsed2.masters!.length).to.equal(2);
    // Master name derived from theme name, not original master name
    expect(parsed2.slides[0].master).to.equal("Theme One");
    expect(parsed2.slides[1].master).to.equal("Theme Two");
  });
});
