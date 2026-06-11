import { describe, expect, it } from "vite-plus/test";

import { generatePresentation } from "./generate";
import { parsePresentation } from "./parse";
import type { MasterDefinition, PresentationOptions, SlideOptions } from "./shared/file";

describe("parsePresentation", () => {
  it("returns PresentationOptions with slides", async () => {
    const options: PresentationOptions = {
      slides: [
        {
          children: [
            {
              shape: {
                x: 100,
                y: 100,
                width: 600,
                height: 400,
                textBody: { text: "Hello" },
                fill: "4472C4",
              },
            },
          ],
        },
      ],
    };
    const buffer = await generatePresentation(options);
    const result = parsePresentation(buffer);

    expect(result.slides).to.exist;
    expect(result.slides!.length).to.equal(1);
    expect(result.masters).to.be.undefined;
  });

  it("parses single master file with undefined masters", async () => {
    const options: PresentationOptions = {
      slides: [
        { children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "A" } } }] },
        { children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "B" } } }] },
      ],
    };
    const buffer = await generatePresentation(options);
    const result = parsePresentation(buffer);

    expect(result.slides!.length).to.equal(2);
    expect(result.masters).to.be.undefined;
  });

  it("parses core properties and size", async () => {
    const options: PresentationOptions = {
      title: "Test Title",
      creator: "Test Creator",
      slides: [
        { children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "A" } } }] },
      ],
    };
    const buffer = await generatePresentation(options);
    const result = parsePresentation(buffer);

    expect(result.title).to.equal("Test Title");
    expect(result.creator).to.equal("Test Creator");
    expect(result.size).to.equal("16:9");
  });

  it("parses multi-master file", async () => {
    const masters: MasterDefinition[] = [
      {
        name: "light",
        theme: { name: "Light", colors: { dark1: "333333" } },
      },
      {
        name: "dark",
        theme: { name: "Dark", colors: { dark1: "FFFFFF" } },
      },
    ];

    const slides: SlideOptions[] = [
      {
        master: "light",
        layout: "blank",
        children: [
          { shape: { x: 100, y: 100, width: 600, height: 400, textBody: { text: "Light slide" } } },
        ],
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
              textBody: {
                text: "Dark slide",
                children: [
                  {
                    properties: { bullet: { type: "none" } },
                    children: [{ text: "Dark slide", fill: "FFFFFF" }],
                  },
                ],
              },
            },
          },
        ],
      },
    ];

    const buffer = await generatePresentation({ title: "Multi-master", masters, slides });
    const result = parsePresentation(buffer);

    expect(result.slides!.length).to.equal(2);
    expect(result.masters).to.exist;
    expect(result.masters!.length).to.equal(2);
    // Master name is derived from theme name
    expect(result.slides![0].master).to.equal("Light");
    expect(result.slides![1].master).to.equal("Dark");
    expect(result.masters![0].name).to.equal("Light");
    expect(result.masters![1].name).to.equal("Dark");
    expect(result.masters![0].theme?.name).to.equal("Light");
    expect(result.masters![1].theme?.name).to.equal("Dark");
  });

  it("round-trips multi-master structure", async () => {
    const masters: MasterDefinition[] = [
      { name: "m1", theme: { name: "Theme One" } },
      { name: "m2", theme: { name: "Theme Two" } },
    ];
    const slides: SlideOptions[] = [
      {
        master: "m1",
        children: [{ shape: { x: 50, y: 50, width: 400, height: 300, textBody: { text: "S1" } } }],
      },
      {
        master: "m2",
        children: [{ shape: { x: 50, y: 50, width: 400, height: 300, textBody: { text: "S2" } } }],
      },
    ];

    const buffer = await generatePresentation({ masters, slides });

    // First parse
    const parsed1 = parsePresentation(buffer);
    expect(parsed1.masters!.length).to.equal(2);

    // Re-generate from parsed data
    const buffer2 = await generatePresentation(parsed1);

    // Second parse
    const parsed2 = parsePresentation(buffer2);
    expect(parsed2.slides!.length).to.equal(2);
    expect(parsed2.masters!.length).to.equal(2);
    // Master name derived from theme name, not original master name
    expect(parsed2.slides![0].master).to.equal("Theme One");
    expect(parsed2.slides![1].master).to.equal("Theme Two");
  });
});
