import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { customizationsDesc } from "./descriptor";

const writeCtx = {} as WriteContext;
const readCtx = {} as ReadContext;

describe("customizationsDesc", () => {
  it("round-trips toolbar data, mask, and fciBasedOn", () => {
    const source =
      '<wne:tcg xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">' +
      '<wne:toolbars><wne:toolbarData r:id="rId1"/>' +
      '<wne:acdManifest><wne:acdEntry wne:acdName="acd0"/></wne:acdManifest></wne:toolbars>' +
      '<wne:keymaps><wne:keymap wne:mask="1" wne:kcmPrimary="0170"/></wne:keymaps>' +
      '<wne:acds><wne:acd wne:acdName="acd0" wne:fciBasedOn="Symbol"/></wne:acds></wne:tcg>';
    const options = customizationsDesc.parse(parseXml(source).elements![0]!, readCtx);
    expect(options).toEqual({
      toolbarDataRelationshipId: "rId1",
      toolbarActionNames: ["acd0"],
      keyMaps: [{ mask: true, primaryKeyCode: "0170" }],
      actionDefinitions: [{ name: "acd0", basedOn: "Symbol" }],
    });
    const xml = customizationsDesc.stringify(options, writeCtx)!;
    expect(xml).toContain('xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"');
    expect(xml).toContain('<wne:toolbarData r:id="rId1"/>');
    expect(xml).toContain('<wne:keymap wne:mask="1" wne:kcmPrimary="0170"/>');
    expect(xml).toContain('wne:fciBasedOn="Symbol"');
  });

  it("round-trips key maps, toolbar actions, and command definitions", () => {
    const options = {
      keyMaps: [
        { primaryKeyCode: "0249", actionName: "acd14" },
        { secondaryKeyCode: "0031", characterCode: "000000A7" },
        { primaryKeyCode: "0032", macroName: "PROJECT.NEWMACROS.MACRO" },
      ],
      toolbarActionNames: ["acd0", "acd1"],
      actionDefinitions: [
        { name: "acd0", argumentValue: "AQAAAAAA", indexBasedOn: "0065" },
        { name: "acd1" },
      ],
    };
    const xml = customizationsDesc.stringify(options, writeCtx)!;
    expect(xml).toContain('<wne:keymap wne:kcmPrimary="0249">');
    expect(xml).toContain('<wne:keymap wne:kcmSecondary="0031">');
    expect(xml).toContain('<wne:macro wne:macroName="PROJECT.NEWMACROS.MACRO"/>');
    expect(customizationsDesc.parse(parseXml(xml).elements![0]!, readCtx)).toEqual(options);
  });
});
