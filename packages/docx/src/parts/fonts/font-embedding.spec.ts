import { unzipSync, zipSync } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../../generate";
import { parseDocument } from "../../parse";

const decoder = new TextDecoder();

/** 64 bytes so the font obfuscation (first 32 bytes XORed) has a payload. */
const FONT_DATA = new Uint8Array(64).fill(0x42);

const options = (name: string) => ({
  sections: [{ children: [{ paragraph: { children: ["Embedded font"] } }] }],
  fonts: [
    {
      name,
      embedRegular: {
        data: FONT_DATA,
        fontKey: "00000000-0000-0000-0000-000000000000",
      },
    },
  ],
});

describe("embedded font part names", () => {
  it("escapes the relationship target and packs the part under the same path", async () => {
    const zip = unzipSync(await generateDocument(options("My Font")));
    expect(zip["word/fonts/My%20Font.odttf"]).toBeDefined();
    expect(zip["word/fonts/My Font.odttf"]).toBeUndefined();
    expect(decoder.decode(zip["word/_rels/fontTable.xml.rels"])).toContain(
      'Target="fonts/My%20Font.odttf"',
    );
  });

  it("round-trips a non-ASCII font name through parse and re-generation", async () => {
    const parsed = await parseDocument(await generateDocument(options("Café Font")));
    expect(parsed.fonts?.[0]?.embedRegular?.odttfPath).toBe("word/fonts/Café Font.odttf");

    const zip = unzipSync(await generateDocument(parsed));
    expect(zip["word/fonts/Caf%C3%A9%20Font.odttf"]).toBeDefined();
    expect(decoder.decode(zip["word/_rels/fontTable.xml.rels"])).toContain(
      'Target="fonts/Caf%C3%A9%20Font.odttf"',
    );
  });

  it("resolves a font part stored under the decoded name", async () => {
    const files = unzipSync(await generateDocument(options("My Font")));
    // A source package whose ZIP entry carries the raw name while the
    // relationship Target is percent-encoded (Office's own layout).
    files["word/fonts/My Font.odttf"] = files["word/fonts/My%20Font.odttf"]!;
    delete files["word/fonts/My%20Font.odttf"];

    const parsed = await parseDocument(zipSync(files));
    expect(parsed.fonts?.[0]?.embedRegular?.odttfPath).toBe("word/fonts/My Font.odttf");
    expect(parsed.fonts?.[0]?.embedRegular?.data).toBeDefined();
  });

  it("preserves an embedded-font reference without inventing a relationship part", async () => {
    const files = unzipSync(await generateDocument(options("My Font")));
    delete files["word/_rels/fontTable.xml.rels"];
    delete files["word/fonts/My%20Font.odttf"];

    const parsed = await parseDocument(zipSync(files));
    expect(parsed.fonts?.[0]?.embedRegular?.embedRid).toBe("rId1");

    const output = unzipSync(await generateDocument(parsed));
    expect(output["word/_rels/fontTable.xml.rels"]).toBeUndefined();
    expect(decoder.decode(output["word/fontTable.xml"])).toContain('r:id="rId1"');
  });

  it("preserves an embedded-font relationship when its binary part is missing", async () => {
    const files = unzipSync(await generateDocument(options("My Font")));
    delete files["word/fonts/My%20Font.odttf"];

    const parsed = await parseDocument(zipSync(files));
    expect(parsed.fonts?.[0]?.embedRegular).toMatchObject({
      embedRid: "rId1",
      odttfPath: "word/fonts/My Font.odttf",
    });
    expect(parsed.fonts?.[0]?.embedRegular?.data).toBeUndefined();

    const output = unzipSync(await generateDocument(parsed));
    expect(output["word/fonts/My%20Font.odttf"]).toBeUndefined();
    expect(decoder.decode(output["word/_rels/fontTable.xml.rels"])).toContain(
      'Target="fonts/My%20Font.odttf"',
    );
    expect(decoder.decode(output["word/fontTable.xml"])).toContain('r:id="rId1"');
  });
});
