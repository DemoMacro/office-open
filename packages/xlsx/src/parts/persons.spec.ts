import { unzipSync, zipSync } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { parseWorkbookSync } from "../parse";
import { personsDesc, type PersonsOptions } from "./persons";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

const options: PersonsOptions = {
  persons: [
    {
      displayName: "Reviewer",
      userId: "user@example.com",
      providerId: "None",
      id: "{00000000-0000-0000-0000-000000000001}",
    },
  ],
};

describe("personsDesc", () => {
  it("preserves an empty person list", () => {
    const xml = personsDesc.stringify({ persons: [] }, writeCtx);
    const root = parseXml(xml!, { nativeTypeAttributes: true }).elements?.[0];
    expect(root).toBeDefined();
    expect(personsDesc.parse(root!, readCtx)).toEqual({ persons: [] });
  });

  it("rejects unknown person-list children", () => {
    const root = parseXml(
      '<tc:personList xmlns:tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><tc:unknown/></tc:personList>',
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => personsDesc.parse(root!, readCtx)).toThrowError(/unexpected child of personList/);
  });

  it("rejects unknown person attributes and missing displayName", () => {
    const ns = 'xmlns:tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"';
    const unknownAttr = parseXml(
      `<tc:personList ${ns}><tc:person displayName="A" extra="1"/></tc:personList>`,
      {
        nativeTypeAttributes: true,
      },
    ).elements?.[0];
    expect(() => personsDesc.parse(unknownAttr!, readCtx)).toThrowError(/unexpected attribute/);
    const missingName = parseXml(`<tc:personList ${ns}><tc:person/></tc:personList>`, {
      nativeTypeAttributes: true,
    }).elements?.[0];
    expect(() => personsDesc.parse(missingName!, readCtx)).toThrowError(/missing required value/);
  });

  it("round-trips persons", () => {
    const xml = personsDesc.stringify(options, writeCtx);
    const root = parseXml(xml!, { nativeTypeAttributes: true }).elements?.[0];
    expect(root).toBeDefined();
    expect(personsDesc.parse(root!, readCtx)).toEqual(options);
  });

  it("emits the typed part, workbook relationship, and content type", async () => {
    const output = (await generateWorkbook(
      { worksheets: [{ name: "Data" }], persons: options },
      { type: "uint8array" },
    )) as Uint8Array;
    const files = unzipSync(output);
    const xml = new TextDecoder().decode(files["xl/persons/person.xml"]!);
    const rels = new TextDecoder().decode(files["xl/_rels/workbook.xml.rels"]!);
    const types = new TextDecoder().decode(files["[Content_Types].xml"]!);

    expect(xml).toContain("tc:personList");
    expect(rels).toContain("2017/10/relationships/person");
    expect(types).toContain("application/vnd.ms-excel.person+xml");

    const parsed = parseWorkbookSync(zipSync(files));
    expect(parsed.persons).toEqual(options);
    expect(parsed.personsPath).toBe("xl/persons/person.xml");
    expect(parsed.rawParts).toBeUndefined();
  });
});
