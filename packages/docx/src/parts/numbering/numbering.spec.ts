import { parse as parseXml } from "@office-open/xml";
import { AlignmentType } from "@parts/paragraph";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraphProperties } from "../../body";
import type { DocxReadContext } from "../../context";
import { LevelFormat, LevelSuffix } from "./level";
import { Numbering, parseNumberingDefinitions } from "./numbering";

describe("Numbering", () => {
  describe("#constructor", () => {
    it("creates a default numbering with one abstract and one concrete instance", () => {
      const numbering = new Numbering({
        config: [],
      });

      const xml = numbering.serialize();

      expect(xml).to.contain("<w:numbering");
      expect(xml).to.contain("<w:abstractNum");
      expect(xml).to.contain("<w:num ");

      // Should contain abstractNumId attribute
      expect(xml).to.contain('w:abstractNumId="');
    });

    describe("#createConcreteNumberingInstance", () => {
      it("should create a concrete numbering instance", () => {
        const numbering = new Numbering({
          config: [
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
          config: [
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
          config: [
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
          config: [
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
        const zeroLevelConfig = referenceConfig[0];
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
      config: [
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
              lvlRestart: 0,
              templateCode: "0409000F",
              isLegalNumberingStyle: true,
              legacy: { space: 0, indent: 0 },
              style: {
                run: { font: "Arial", bold: true },
                paragraph: { indent: { left: 720, hanging: 360 } },
              },
            },
          ],
        },
      ],
    });
    // Numbering only auto-creates the abstract definition from config; the
    // concrete w:num instance (what parse iterates) needs an explicit instance.
    numbering.createConcreteNumberingInstance("decimal-list", 0);

    const xml = numbering.serialize();
    const el = parseXml(xml).elements![0];
    const opts = parseNumberingDefinitions(el, parseParagraphProperties, ctx);

    expect(opts).toBeDefined();
    // serialize() always emits the built-in default-bullet-numbering (bullet
    // levels), so after round-trip the parsed config holds it alongside the
    // decimal config. The parsed reference is derived from numId (list_N), not
    // the original config name, so locate the decimal entry by its level format.
    const decimalConfig = opts!.config.find((c) => c.levels[0]?.format === LevelFormat.DECIMAL);
    expect(decimalConfig).toBeDefined();
    const lvl = decimalConfig!.levels[0];
    expect(lvl.start).toBe(5);
    expect(lvl.format).toBe(LevelFormat.DECIMAL);
    expect(lvl.text).toBe("%1.");
    expect(lvl.alignment).toBe(AlignmentType.LEFT);
    expect(lvl.suffix).toBe(LevelSuffix.TAB);
    expect(lvl.lvlRestart).toBe(0);
    expect(lvl.templateCode).toBe("0409000F");
    expect(lvl.isLegalNumberingStyle).toBe(true);
    expect(lvl.legacy).toEqual({ space: 0, indent: 0 });
    expect(lvl.style?.run?.bold).toBe(true);
    // font:"Arial" serializes as w:rFonts ascii+hAnsi (Word convention: hAnsi
    // defaults to the ascii font), so it round-trips as a multi-field object —
    // assert the ascii facet survives.
    const runFont = lvl.style?.run?.font as { ascii?: string } | string | undefined;
    expect(typeof runFont === "string" ? runFont : runFont?.ascii).toBe("Arial");
    expect(lvl.style?.paragraph?.indent).toEqual({ left: 720, hanging: 360 });
  });
});
