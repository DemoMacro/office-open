import type { OoxmlPackageFormat } from "@office-open/core";
import { OOXML_PACKAGE_FORMATS, unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generate, parseInput, type GenerateOptionsMap, type GenerateType } from "./generate";

const OOXML_TYPES = Object.keys(OOXML_PACKAGE_FORMATS) as OoxmlPackageFormat[];

const PACKAGE_FIXTURES = {
  wordprocessing: { sections: [{ children: [{ paragraph: { children: ["Variant"] } }] }] },
  presentation: {
    slides: [
      {
        children: [
          {
            shape: {
              x: 100,
              y: 100,
              width: 600,
              height: 400,
              textBody: { text: "Variant" },
            },
          },
        ],
      },
    ],
  },
  spreadsheet: { worksheets: [{ rows: [{ cells: [{ value: "Variant" }] }] }] },
} as const;

const findOverride = (contentTypes: string, partPath: string): string => {
  const match = contentTypes.match(
    new RegExp(
      `<Override PartName="/${partPath.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}" ContentType="([^"]+)"`,
    ),
  );
  return match?.[1] ?? "";
};

describe("parseInput", () => {
  it("should parse a JSON string", async () => {
    const result = await parseInput('{"sections": []}');
    expect(result).toEqual({ sections: [] });
  });

  it("should parse a JSON string with nested objects", async () => {
    const input = JSON.stringify({ worksheets: [{ rows: [{ cells: [{ value: 1 }] }] }] });
    const result = await parseInput(input);
    expect(result).toEqual({ worksheets: [{ rows: [{ cells: [{ value: 1 }] }] }] });
  });

  it("should parse a JSON string starting with [", async () => {
    const input = JSON.stringify([{ foo: "bar" }]);
    const result = await parseInput(input);
    expect(result).toEqual([{ foo: "bar" }]);
  });

  it("should throw on invalid JSON string", async () => {
    await expect(parseInput("not json")).rejects.toThrow();
  });

  it("should trim whitespace before parsing", async () => {
    const result = await parseInput('  {"sections": []}  ');
    expect(result).toEqual({ sections: [] });
  });

  it("should handle empty object", async () => {
    const result = await parseInput("{}");
    expect(result).toEqual({});
  });

  it("should handle empty array", async () => {
    const result = await parseInput("[]");
    expect(result).toEqual([]);
  });
});

describe("generate OOXML package variants", () => {
  it.for(OOXML_TYPES)("emits the %s package and main-part content type", async (type) => {
    const format = OOXML_PACKAGE_FORMATS[type];
    const bytes = (await generate({
      type,
      options: PACKAGE_FIXTURES[format.family] as GenerateOptionsMap[GenerateType],
      outputType: "uint8array",
    })) as Uint8Array;
    const contentTypes = new TextDecoder().decode(unzipSync(bytes)["[Content_Types].xml"]);

    expect(findOverride(contentTypes, format.mainPartPath)).toEqual(format.mainContentType);
  });
});

describe("generate RTF", () => {
  it("writes RTF through the aggregate API", async () => {
    const bytes = (await generate({
      type: "rtf",
      options: { sections: [{ children: [{ paragraph: { children: ["Aggregate"] } }] }] },
      outputType: "uint8array",
    })) as Uint8Array;
    expect(new TextDecoder().decode(bytes)).toContain("Aggregate");
  });
});
