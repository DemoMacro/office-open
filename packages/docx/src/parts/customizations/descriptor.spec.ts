import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { customizationsDesc } from "./descriptor";

const writeCtx = {} as WriteContext;
const readCtx = {} as ReadContext;

describe("customizationsDesc", () => {
  it("round-trips key maps, toolbar actions, and command definitions", () => {
    const options = {
      keyMaps: [
        { primaryKeyCode: "0249", actionName: "acd14" },
        { secondaryKeyCode: "0031", characterCode: "000000A7" },
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
    expect(customizationsDesc.parse(parseXml(xml).elements![0]!, readCtx)).toEqual(options);
  });
});
