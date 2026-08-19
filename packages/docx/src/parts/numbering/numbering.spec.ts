import { parse as parseXml } from "@office-open/xml";
import { AlignmentType } from "@parts/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraphProperties } from "../../body";
import type { DocxReadContext, DocxWriteContext } from "../../context";
import { LevelFormat, LevelSuffix } from "./level";
import { Numbering, parseNumberingDefinitions } from "./numbering";

const writeCtx = {
  file: {
    media: {
      addMedia: (_d: unknown, _t: string, _b: unknown, fileName?: string) => ({
        fileName: fileName ?? "image1.png",
      }),
    },
  },
} as unknown as DocxWriteContext;

describe("Numbering", () => {
  describe("#constructor", () => {
    it("creates a default numbering with one abstract and one concrete instance", () => {
      const numbering = new Numbering({
        abstractNumberings: [],
      });

      const xml = numbering.serialize(writeCtx);

      expect(xml).to.contain("<w:numbering");
      expect(xml).to.contain("<w:abstractNum");
      expect(xml).to.contain("<w:num ");

      // Should contain abstractNumId attribute
      expect(xml).to.contain('w:abstractNumId="');
    });

    describe("#createConcreteNumberingInstance", () => {
      it("should create a concrete numbering instance", () => {
        const numbering = new Numbering({
          abstractNumberings: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        expect(numbering.concreteNumbering).to.have.length(0);

        numbering.createConcreteNumberingInstance("test-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(1);
      });

      it("should not create a concrete numbering instance if reference is invalid", () => {
        const numbering = new Numbering({
          abstractNumberings: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        expect(numbering.concreteNumbering).to.have.length(0);

        numbering.createConcreteNumberingInstance("invalid-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(0);
      });

      it("should not create a concrete numbering instance if one already exists", () => {
        const numbering = new Numbering({
          abstractNumberings: [
            {
              levels: [
                {
                  level: 0,
                },
              ],
              reference: "test-reference",
            },
          ],
        });

        expect(numbering.concreteNumbering).to.have.length(0);

        numbering.createConcreteNumberingInstance("test-reference", 0);
        numbering.createConcreteNumberingInstance("test-reference", 0);

        expect(numbering.concreteNumbering).to.have.length(1);
      });
    });
    describe("#referenceConfigMap", () => {
      it("should store level configs into referenceConfigMap", () => {
        const numbering = new Numbering({
          abstractNumberings: [
            {
              levels: [
                {
                  level: 0,
                  start: 10,
                },
              ],
              reference: "test-reference",
            },
          ],
        });
        numbering.createConcreteNumberingInstance("test-reference", 0);
        const referenceConfig = numbering.referenceConfig[0];
        if (!referenceConfig) throw new Error("referenceConfig not parsed");
        const zeroLevelConfig = referenceConfig[0];
        if (!zeroLevelConfig) throw new Error("zero level config not parsed");
        expect(zeroLevelConfig.start).to.be.equal(10);
      });
    });
  });
});

describe("parseNumberingDefinitions (round-trip)", () => {
  // Numbering level pPr carries indent/spacing, never numPr, so an empty read
  // context suffices for parseParagraphProperties.
  const ctx = {} as unknown as DocxReadContext;

  it("reads back every level field the serializer writes", () => {
    const numbering = new Numbering({
      abstractNumberings: [
        {
          reference: "decimal-list",
          levels: [
            {
              level: 0,
              start: 5,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
              suffix: LevelSuffix.TAB,
              levelRestart: 0,
              templateCode: "0409000F",
              isLegalNumberingStyle: true,
              legacy: { space: 0, indent: 0 },
              run: { font: "Arial", bold: true },
              paragraph: { indent: { left: 720, hanging: 360 } },
            },
          ],
        },
      ],
    });
    // Numbering only auto-creates the abstract definition from config; the
    // concrete w:num instance (what parse iterates) needs an explicit instance.
    numbering.createConcreteNumberingInstance("decimal-list", 0);

    const xml = numbering.serialize(writeCtx);
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    expect(opts).toBeDefined();
    // serialize() always emits the built-in default-bullet-numbering (bullet
    // levels), so after round-trip the parsed config holds it alongside the
    // decimal config. The parsed reference is derived from numId (list_N), not
    // the original config name, so locate the decimal entry by its level format.
    const decimalConfig = opts!.abstractNumberings.find(
      (c) => c.levels[0]?.format === LevelFormat.DECIMAL,
    );
    expect(decimalConfig).toBeDefined();
    const lvl = decimalConfig!.levels[0];
    if (!lvl) throw new Error("decimal level not parsed");
    expect(lvl.start).toBe(5);
    expect(lvl.format).toBe(LevelFormat.DECIMAL);
    expect(lvl.text).toBe("%1.");
    expect(lvl.alignment).toBe(AlignmentType.LEFT);
    expect(lvl.suffix).toBe(LevelSuffix.TAB);
    expect(lvl.levelRestart).toBe(0);
    expect(lvl.templateCode).toBe("0409000F");
    expect(lvl.isLegalNumberingStyle).toBe(true);
    expect(lvl.legacy).toEqual({ enabled: true, space: 0, indent: 0 });
    expect(lvl.run?.bold).toBe(true);
    // font:"Arial" serializes as w:rFonts ascii+hAnsi (Word convention: hAnsi
    // defaults to the ascii font), so it round-trips as a multi-field object —
    // assert the ascii facet survives.
    const runFont = lvl.run?.font as { ascii?: string } | string | undefined;
    expect(typeof runFont === "string" ? runFont : runFont?.ascii).toBe("Arial");
    expect(lvl.paragraph?.indent).toEqual({ left: 720, hanging: 360 });
  });

  it("round-trips numPicBullets (pict) and numIdMacAtCleanup", () => {
    const numbering = new Numbering({
      abstractNumberings: [],
      numPicBullets: [
        { numPicBulletId: 3, pict: { children: [{ shape: { id: "_x0000_i1025" } }] } },
      ],
      numIdMacAtCleanup: 9,
    });
    const xml = numbering.serialize(writeCtx);
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    expect(opts?.numPicBullets).toHaveLength(1);
    expect(opts?.numPicBullets?.[0]?.numPicBulletId).toBe(3);
    expect(opts?.numPicBullets?.[0]?.pict?.children?.[0]).toMatchObject({
      shape: { id: "_x0000_i1025" },
    });
    expect(opts?.numPicBullets?.[0]?.drawing).toBeUndefined();
    expect(opts?.numIdMacAtCleanup).toBe(9);
  });

  it("round-trips numPicBullet with drawing choice", () => {
    const numbering = new Numbering({
      abstractNumberings: [],
      numPicBullets: [{ numPicBulletId: 5, drawing: "<w:drawing><wp:inline/></w:drawing>" }],
    });
    const xml = numbering.serialize(writeCtx);
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    expect(opts?.numPicBullets?.[0]?.numPicBulletId).toBe(5);
    expect(opts?.numPicBullets?.[0]?.drawing).toContain("<w:drawing");
    expect(opts?.numPicBullets?.[0]?.pict).toBeUndefined();
  });

  it("round-trips abstractNum name, styleLink, numStyleLink", () => {
    const numbering = new Numbering({
      abstractNumberings: [
        {
          reference: "linked-list",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1." }],
          properties: { name: "My List", styleLink: "ListStyle", numStyleLink: "NumStyle" },
        },
      ],
    });
    numbering.createConcreteNumberingInstance("linked-list", 0);
    const xml = numbering.serialize(writeCtx);
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    const linked = opts?.abstractNumberings.find((c) => c.properties?.name === "My List");
    expect(linked).toBeDefined();
    expect(linked?.properties?.styleLink).toBe("ListStyle");
    expect(linked?.properties?.numStyleLink).toBe("NumStyle");
  });

  it("keeps orphan abstractNums (no w:num reference) as instance-less definitions", () => {
    // Word keeps abstractNums whose concrete num was deleted; styles may still
    // reference them. Dropping them loses the definition wholesale on round-trip.
    const xml =
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:multiLevelType w:val="hybridMultilevel"/><w:name w:val="Used List"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      "</w:abstractNum>" +
      '<w:abstractNum w:abstractNumId="7">' +
      '<w:tmpl w:val="0409001D"/><w:name w:val="Orphan List"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/></w:lvl>' +
      "</w:abstractNum>" +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      "</w:numbering>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    expect(opts?.abstractNumberings.find((c) => c.reference === "list_1")).toBeDefined();
    const orphan = opts?.abstractNumberings.find((c) => c.reference === "abstract_7");
    expect(orphan).toBeDefined();
    expect(orphan?.instanceCount ?? 0).toBe(0);
    expect(orphan?.properties?.name).toBe("Orphan List");
    expect(orphan?.properties?.tmpl).toBe("0409001D");
    expect(orphan?.levels[0]?.start).toBe(5);

    // Re-emitted: both definitions survive, only the referenced one gets a w:num.
    const out = new Numbering({ abstractNumberings: opts!.abstractNumberings }).serialize(writeCtx);
    expect(out).toContain('w:name w:val="Orphan List"');
    expect((out.match(/<w:num /g) ?? []).length).toBe(1);
  });

  it("keeps a nested w:lvl lvlOverride on the instance, abstract untouched", () => {
    // CT_NumLvl choice: startOverride | lvl. The nested-lvl form redefines the
    // level wholesale (Word's "restart and redefine" shape) — parse keeps it
    // on the instance config so the abstract definition and the override both
    // round-trip; merging into the abstract levels would re-emit the override
    // content inside w:abstractNum and drop the w:lvlOverride wrapper.
    const xml =
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      "</w:abstractNum>" +
      '<w:num w:numId="1">' +
      '<w:abstractNumId w:val="0"/>' +
      '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">' +
      '<w:start w:val="3"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/>' +
      "</w:lvl></w:lvlOverride>" +
      "</w:num>" +
      "</w:numbering>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    const first = opts?.abstractNumberings[0];
    expect(first).toBeDefined();
    expect(first!.levels[0]?.start).toBe(1);
    expect(first!.levels[0]?.format).toBe("decimal");
    expect(first!.overrideLevels?.[0]?.num).toBe(0);
    expect(first!.overrideLevels?.[0]?.level?.start).toBe(3);
    expect(first!.overrideLevels?.[0]?.level?.format).toBe("lowerLetter");
    expect(first!.overrideLevels?.[0]?.level?.text).toBe("(%1)");

    // Re-emitted: the override wrapper and its nested level survive verbatim.
    const out = new Numbering({ abstractNumberings: opts!.abstractNumberings }).serialize(writeCtx);
    expect(out).toContain(
      '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/>',
    );
  });

  it("keeps per-instance lvlOverride/startOverride separate from the abstract level", () => {
    // abstract level 0 starts at 1; the concrete num re-pins it to 3 via
    // lvlOverride. The override stays on the instance config so the w:num
    // re-emits it instead of silently reverting the list's restart numbering.
    const xml =
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:multiLevelType w:val="hybridMultilevel"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      "</w:abstractNum>" +
      '<w:num w:numId="1">' +
      '<w:abstractNumId w:val="0"/>' +
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>' +
      "</w:num>" +
      "</w:numbering>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    const first = opts?.abstractNumberings[0];
    expect(first).toBeDefined();
    expect(first!.levels[0]?.start).toBe(1);
    expect(first!.overrideLevels?.[0]).toEqual({ num: 0, start: 3 });

    const out = new Numbering({ abstractNumberings: opts!.abstractNumberings }).serialize(writeCtx);
    expect(out).toContain('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>');
  });
});

describe("numbering instance fidelity", () => {
  it("emits w:num for instances the body never references (instanceCount)", async ({ expect }) => {
    const { Numbering } = await import("./numbering");
    const ctx = { media: { array: [] } } as never;
    const n = new Numbering({
      abstractNumberings: [
        {
          reference: "list_9",
          levels: [{ level: 0, format: "bullet", text: "", alignment: "left" }],
          instanceCount: 1,
        },
      ],
    });
    const xml = n.serialize(ctx);
    expect((xml.match(/<w:num /g) ?? []).length).toBe(1);
    expect(xml).toContain("<w:abstractNumId");
  });

  it("does not inject the default bullet numbering when only numPicBullets round-trip", async ({
    expect,
  }) => {
    const { Numbering } = await import("./numbering");
    const ctx = { media: { array: [] } } as never;
    const n = new Numbering({
      abstractNumberings: [],
      numPicBullets: [{ numPicBulletId: 1 }],
    });
    const xml = n.serialize(ctx);
    expect(xml).not.toContain("w:abstractNum ");
    expect((xml.match(/<w:num /g) ?? []).length).toBe(0);
  });

  it("round-trips an empty lvlText value", async ({ expect }) => {
    const { Numbering } = await import("./numbering");
    const n = new Numbering({
      abstractNumberings: [
        {
          reference: "list_1",
          levels: [{ level: 0, format: "decimal", text: "", alignment: "left" }],
        },
      ],
    });
    n.createConcreteNumberingInstance("list_1", 0);
    const xml = n.serialize(writeCtx as never);
    expect(xml).toContain('<w:lvlText w:val=""/>');
    // And parse keeps the empty string (not dropped as falsy)
    const doc = parseXml(xml.replace(/^<\?xml[^>]*\?>/, ""));
    const opts = parseNumberingDefinitions(
      doc.elements![0]!,
      parseParagraphProperties,
      {} as unknown as DocxReadContext,
    );
    expect(opts!.abstractNumberings[0]!.levels[0]!.text).toBe("");
  });
});
