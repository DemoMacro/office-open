import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { stringifyRunInline } from "../../inline";
import { stringifyRunProperties } from "../stringify";
import type { RunPropertiesOptions } from "./properties";
import type { RunOptions } from "./run";
import { breakXml } from "./run";
import { parseRun, parseRunProperties, parsedRunToOptions } from "./run-parse";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W16SE_NS = 'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"';

function roundTrip(opts: RunPropertiesOptions): RunPropertiesOptions {
  const rPr = stringifyRunProperties(opts)!;
  const doc = parseXml(`<w:r ${W_NS}>${rPr}</w:r>`);
  const r = doc.elements?.[0];
  if (!r) throw new Error("parsed document has no root element");
  const rPrEl = r.elements?.[0];
  if (!rPrEl) throw new Error("run has no rPr element");
  return parseRunProperties(rPrEl);
}

describe("parseRunProperties round-trip", () => {
  it("preserves duplicate rFonts elements in order", () => {
    const sourceXml =
      `<w:rPr ${W_NS}>` +
      '<w:rFonts w:ascii="Arial"/>' +
      '<w:rFonts w:eastAsia="SimSun"/>' +
      '<w:rFonts w:hAnsi="Consolas"/>' +
      "</w:rPr>";
    const sourceElement = parseXml(sourceXml).elements?.[0];
    if (!sourceElement) throw new Error("parsed document has no root element");

    const first = parseRunProperties(sourceElement);
    expect(first.font).toEqual({ ascii: "Arial" });
    expect(first.fontDuplicates).toEqual([{ eastAsia: "SimSun" }, { hAnsi: "Consolas" }]);

    const generatedXml = stringifyRunProperties(first)!;
    expect(generatedXml.match(/<w:rFonts\b/g)).toHaveLength(3);
    expect(generatedXml).toContain('<w:rFonts w:ascii="Arial"/><w:rFonts w:eastAsia="SimSun"/>');

    const generatedElement = parseXml(`<w:r ${W_NS}>${generatedXml}</w:r>`).elements?.[0]
      ?.elements?.[0];
    if (!generatedElement) throw new Error("generated run has no rPr element");
    const second = parseRunProperties(generatedElement);
    expect(second).toEqual(first);
    expect(stringifyRunProperties(second)).toBe(generatedXml);
  });

  it("round-trips color with themeColor/themeTint/themeShade", () => {
    const result = roundTrip({
      color: { val: "FF0000", themeColor: "text1", themeTint: "99", themeShade: "BF" },
    });
    expect(result.color).toEqual({
      val: "FF0000",
      themeColor: "text1",
      themeTint: "99",
      themeShade: "BF",
    });
  });

  it("preserves source order when text precedes a break", () => {
    const source = `<w:r ${W_NS}><w:t>after</w:t><w:br/></w:r>`;
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed run has no root element");
    const opts = parsedRunToOptions(parseRun(element, {} as never)) as RunOptions;
    expect(opts.children).toEqual(["after", { break: 1 }]);
    expect(stringifyRunInline(opts, {} as never)).toBe("<w:r><w:t>after</w:t><w:br/></w:r>");
  });

  it("preserves source order when run properties follow content", () => {
    const source = `<w:r ${W_NS}><w:t>after</w:t><w:rPr><w:b/></w:rPr></w:r>`;
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed run has no root element");
    const opts = parsedRunToOptions(parseRun(element, {} as never)) as RunOptions;
    expect(opts.childOrder).toEqual(["content", "runProperties"]);
    expect(stringifyRunInline(opts, {} as never)).toBe(
      "<w:r><w:t>after</w:t><w:rPr><w:b/></w:rPr></w:r>",
    );
  });

  it("parses fallback run properties when a compatibility choice is empty", () => {
    const doc = parseXml(
      `<w:rPr ${W_NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
        `<mc:AlternateContent><mc:Choice Requires="w16se"/><mc:Fallback>` +
        `<w:rFonts w:ascii="Segoe UI Emoji" w:eastAsia="Segoe UI Emoji" ` +
        `w:hAnsi="Segoe UI Emoji" w:cs="Segoe UI Emoji"/></mc:Fallback>` +
        `</mc:AlternateContent></w:rPr>`,
    );
    const rPrEl = doc.elements?.[0];
    if (!rPrEl) throw new Error("parsed document has no rPr element");

    expect(parseRunProperties(rPrEl).font).toEqual({
      ascii: "Segoe UI Emoji",
      complexScript: "Segoe UI Emoji",
      eastAsia: "Segoe UI Emoji",
      hAnsi: "Segoe UI Emoji",
    });
  });

  it("round-trips color as plain string when no theme attributes", () => {
    const result = roundTrip({ color: "FF0000" });
    expect(result.color).toBe("FF0000");
  });

  it("round-trips underline with themeColor/themeTint/themeShade", () => {
    const result = roundTrip({
      underline: {
        type: "single",
        color: "FF0000",
        themeColor: "text1",
        themeTint: "99",
        themeShade: "BF",
      },
    });
    expect(result.underline).toEqual({
      type: "single",
      color: "FF0000",
      themeColor: "text1",
      themeTint: "99",
      themeShade: "BF",
    });
  });

  it("omits underline val when the source omits it", () => {
    const xml = stringifyRunProperties({ underline: { color: "FF0000" } })!;
    expect(xml).toContain('<w:u w:color="FF0000"/>');
    expect(xml).not.toContain("w:val");
    expect(roundTrip({ underline: {} })).toEqual({ underline: {} });
  });

  it("round-trips eastAsianLayout", () => {
    const result = roundTrip({
      eastAsianLayout: {
        id: 1,
        combine: true,
        combineBrackets: "round",
        vertical: true,
        verticalCompress: false,
      },
    });
    expect(result.eastAsianLayout).toEqual({
      id: 1,
      combine: true,
      combineBrackets: "round",
      vertical: true,
      verticalCompress: false,
    });
  });

  it("round-trips contentPartRId", () => {
    const result = roundTrip({ contentPartRId: "rId42" });
    expect(result.contentPartRId).toBe("rId42");
  });

  it("round-trips rPrChange revision with inner rPr", () => {
    const result = roundTrip({
      bold: true,
      revision: { id: 1, author: "A", date: "2024-01-01T00:00:00Z", italic: true },
    });
    expect(result.revision).toBeDefined();
    const rev = result.revision as unknown as Record<string, unknown>;
    expect(rev.id).toBe(1);
    expect(rev.author).toBe("A");
    expect(rev.date).toBe("2024-01-01T00:00:00Z");
    expect(rev.italic).toBe(true);
  });

  it("normalizes characterSpacing UniversalMeasure (mm) to twips", () => {
    const result = roundTrip({ characterSpacing: "0.5mm" });
    expect(result.characterSpacing).toBe(28);
  });

  // b/bCs, i/iCs, sz/szCs are independent toggle/measure properties (Latin vs
  // complex script, per ISO/IEC 29500). stringify must not auto-pair them
  // (bold → b+bCs would inflate), so round-trip stays field-faithful:
  // a source <w:b/> round-trips as <w:b/>, not <w:b/><w:bCs/>.
  it("does not auto-pair bCs when only bold is set", () => {
    const rPr = stringifyRunProperties({ bold: true })!;
    expect(rPr).toContain("<w:b");
    expect(rPr).not.toContain("bCs");
  });

  it("emits bCs only when boldComplexScript is explicitly set", () => {
    const rPr = stringifyRunProperties({ bold: true, boldComplexScript: true })!;
    expect(rPr).toContain("<w:b");
    expect(rPr).toContain("<w:bCs");
  });

  it("round-trips bold without inflating boldComplexScript", () => {
    const result = roundTrip({ bold: true });
    expect(result.bold).toBe(true);
    expect(result.boldComplexScript).toBeUndefined();
  });

  it("preserves explicit false-valued toggle spellings", () => {
    const source =
      `<w:rPr ${W_NS}>` +
      '<w:b w:val="false"/><w:bCs w:val="0"/><w:i w:val="off"/><w:iCs w:val="f"/>' +
      "</w:rPr>";
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed rPr has no root element");
    const parsed = parseRunProperties(element);
    expect(parsed).toMatchObject({
      bold: false,
      boldRaw: "false",
      boldComplexScript: false,
      boldComplexScriptRaw: "0",
      italic: false,
      italicRaw: "off",
      italicComplexScript: false,
      italicComplexScriptRaw: "f",
    });
    const xml = stringifyRunProperties(parsed)!;
    expect(xml).toContain('<w:b w:val="false"/>');
    expect(xml).toContain('<w:bCs w:val="0"/>');
    expect(xml).toContain('<w:i w:val="off"/>');
    expect(xml).toContain('<w:iCs w:val="f"/>');
  });

  it("does not auto-pair iCs when only italic is set", () => {
    const rPr = stringifyRunProperties({ italic: true })!;
    expect(rPr).toContain("<w:i");
    expect(rPr).not.toContain("iCs");
  });

  it("does not auto-pair szCs when only size is set", () => {
    const rPr = stringifyRunProperties({ size: 24 })!;
    expect(rPr).toContain("<w:sz ");
    expect(rPr).not.toContain("szCs");
  });

  it("emits szCs when sizeComplexScript is an explicit number", () => {
    const rPr = stringifyRunProperties({ size: 24, sizeComplexScript: 20 })!;
    expect(rPr).toContain("<w:sz ");
    expect(rPr).toContain("<w:szCs");
  });

  it("round-trips all three verticalAlign values", () => {
    for (const val of ["baseline", "subscript", "superscript"] as const) {
      const rPr = stringifyRunProperties({ verticalAlign: val })!;
      expect(rPr).toContain(`<w:vertAlign w:val="${val}"/>`);
      expect(roundTrip({ verticalAlign: val }).verticalAlign).toBe(val);
    }
  });
});

describe("parseRun rsid attributes", () => {
  it("reads w:rsidR/w:rsidRPr/w:rsidDel (hex verbatim, leading zeros kept)", () => {
    const doc = parseXml(
      `<w:r ${W_NS} w:rsidR="00992297" w:rsidRPr="00112233" w:rsidDel="AABBCCDD"><w:t>hi</w:t></w:r>`,
    );
    // parseRun does not use its context for rsid reads.
    const rEl = doc.elements?.[0];
    if (!rEl) throw new Error("parsed document has no root element");
    const parsed = parseRun(rEl, {} as never);
    expect(parsed.additionRsid).toBe("00992297");
    expect(parsed.runPropertiesRsid).toBe("00112233");
    expect(parsed.deletionRsid).toBe("AABBCCDD");
  });
});

describe("empty run preservation (CT_Run allows an empty w:r)", () => {
  it("keeps a content-less run as {} instead of dropping it", () => {
    const doc = parseXml(`<w:r ${W_NS}/>`);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    expect(parsedRunToOptions(parseRun(el, {} as never))).toEqual({});
  });

  it("keeps a run whose bare text node falls outside the content model", () => {
    const doc = parseXml(`<w:r ${W_NS}>stray</w:r>`);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    expect(parsedRunToOptions(parseRun(el, {} as never))).toEqual({});
  });
});

describe("parseRun text space preservation", () => {
  it("keeps the source marker on multiple w:t segments", () => {
    const doc = parseXml(`<w:r ${W_NS}><w:t>a</w:t><w:t xml:space="preserve">b</w:t></w:r>`);
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(el, {} as never));
    expect(opts).toEqual({
      children: ["a", { text: "b", preserveSpace: true }],
      preserveSpace: true,
    });
    const xml = stringifyRunInline(opts as never, {} as never);
    expect(xml).toContain("<w:t>a</w:t>");
    expect(xml).toContain('<w:t xml:space="preserve">b</w:t>');
  });
});

describe("parseRun break clear (CT_Br/@w:clear)", () => {
  it("parses and stringifies textWrapping break types", () => {
    const doc = parseXml(`<w:r ${W_NS}><w:br w:type="textWrapping"/></w:r>`);
    const brEl = doc.elements?.[0];
    if (!brEl) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(brEl, {} as never));
    expect(opts).toEqual({ break: { count: 1, type: "textWrapping" } });
    expect(breakXml((opts as RunOptions).break)).toBe('<w:br w:type="textWrapping"/>');
  });

  it("preserves a break type and clear together", () => {
    expect(breakXml({ count: 1, type: "column", clear: "all" })).toBe(
      '<w:br w:type="column" w:clear="all"/>',
    );
  });

  it("parses w:br/@w:clear into a structured break", () => {
    const doc = parseXml(`<w:r ${W_NS}><w:br w:clear="all"/></w:r>`);
    const brEl = doc.elements?.[0];
    if (!brEl) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(brEl, {} as never)) as {
      break: { count: number; clear: string };
    };
    expect(opts.break).toEqual({ count: 1, clear: "all" });
  });

  it("parses a plain line break as a count", () => {
    const doc = parseXml(`<w:r ${W_NS}><w:br/></w:r>`);
    const plainBrEl = doc.elements?.[0];
    if (!plainBrEl) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(plainBrEl, {} as never)) as {
      break: number;
    };
    expect(opts.break).toBe(1);
  });

  it("stringifies a structured break with clear (breakXml)", () => {
    expect(breakXml({ count: 2, clear: "left" })).toBe(
      '<w:br w:clear="left"/><w:br w:clear="left"/>',
    );
    expect(breakXml(3)).toBe("<w:br/><w:br/><w:br/>");
  });
});

describe("parseRun customMarkFollows (CT_FtnEdnRef)", () => {
  it("parses w:footnoteReference/@w:customMarkFollows", () => {
    const doc = parseXml(
      `<w:r ${W_NS}><w:footnoteReference w:id="3" w:customMarkFollows="true"/></w:r>`,
    );
    const fnEl = doc.elements?.[0];
    if (!fnEl) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(fnEl, {} as never)) as {
      footnoteReference: { id: number; customMarkFollows: boolean };
    };
    expect(opts.footnoteReference).toEqual({ id: 3, customMarkFollows: true });
  });

  it("parses a plain footnoteReference as a numeric id", () => {
    const doc = parseXml(`<w:r ${W_NS}><w:footnoteReference w:id="5"/></w:r>`);
    const plainFnEl = doc.elements?.[0];
    if (!plainFnEl) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(plainFnEl, {} as never)) as {
      footnoteReference: number;
    };
    expect(opts.footnoteReference).toBe(5);
  });
});

describe("Office 2016 symbol parse", () => {
  it("preserves the w16se symbol vocabulary used by Word", () => {
    const doc = parseXml(
      `<w:r ${W_NS} ${W16SE_NS}><w16se:sym w:font="Webdings" w:char="F04E"/></w:r>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    expect(parsedRunToOptions(parseRun(el, {} as never))).toEqual({
      symbolRun: { char: "F04E", symbolFont: "Webdings", kind: "office2016" },
    });
  });

  it("accepts the schema symEx spelling and extension-qualified attributes", () => {
    const doc = parseXml(
      `<w:r ${W_NS} ${W16SE_NS}><w16se:symEx w16se:font="Webdings" w16se:char="F04E"/></w:r>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    expect(parsedRunToOptions(parseRun(el, {} as never))).toEqual({
      symbolRun: { char: "F04E", symbolFont: "Webdings", kind: "office2016" },
    });
  });

  it("unwraps an Office 2016 symbol from markup compatibility choice", () => {
    const doc = parseXml(
      `<w:r ${W_NS} ${W16SE_NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
        `<mc:AlternateContent><mc:Choice Requires="w16se">` +
        `<w16se:symEx w16se:font="Webdings" w16se:char="F04E"/>` +
        `</mc:Choice><mc:Fallback><w:t>icon</w:t></mc:Fallback></mc:AlternateContent></w:r>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    expect(parsedRunToOptions(parseRun(el, {} as never))).toEqual({
      symbolRun: { char: "F04E", symbolFont: "Webdings", kind: "office2016" },
    });
  });

  it("unwraps a strict OLE object from markup compatibility choice", () => {
    const doc = parseXml(
      `<w:r ${W_NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
        `xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
        `<mc:AlternateContent><mc:Choice Requires="v">` +
        `<w:object w:dxaOrig="100" w:dyaOrig="50">` +
        `<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:100pt;height:50pt" o:ole=""/>` +
        `</w:object></mc:Choice>` +
        `<mc:Fallback><w:object><w:drawing/></w:object></mc:Fallback></mc:AlternateContent></w:r>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");

    expect(parsedRunToOptions(parseRun(el, {} as never))).toMatchObject({
      object: {
        dxaOrig: 100,
        dyaOrig: 50,
        shapeId: "_x0000_i1025",
        mcChoiceRequires: "v",
        mcFallback: "<mc:Fallback><w:object><w:drawing/></w:object></mc:Fallback>",
      },
    });
  });
});

describe("parsedRunToOptions mixed block children", () => {
  it("keeps symbols and picts ordered in children[] (no flattening drop)", () => {
    const doc = parseXml(
      `<w:r ${W_NS}><w:t>a</w:t><w:sym w:char="F04A" w:font="Webdings"/><w:t>b</w:t>` +
        `<w:pict><v:rect id="_x0000_s1027"/></w:pict></w:r>`,
    );
    const el = doc.elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parsedRunToOptions(parseRun(el, {} as never)) as {
      children: Record<string, unknown>[];
    };
    // All four children survive in order — neither the text simplification
    // nor the exclusive pict/symbol flattening may drop siblings.
    expect(opts.children).toHaveLength(4);
    expect(opts.children[1]).toMatchObject({
      symbolRun: { char: "F04A", symbolFont: "Webdings" },
    });
    expect(opts.children[3]).toMatchObject({
      pict: { children: [{ rect: { id: "_x0000_s1027" } }] },
    });
  });
});
