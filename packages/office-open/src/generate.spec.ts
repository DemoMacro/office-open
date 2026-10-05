import { OOXML_PACKAGE_FORMATS, type OoxmlPackageFormat, unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateOffice, type OfficeGenerateFormat, type OfficeOptionsFor } from "./generate";

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
      `<Override PartName="/${partPath.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&")}" ContentType="([^"]+)"`,
    ),
  );
  return match?.[1] ?? "";
};

describe("generateOffice OOXML package variants", () => {
  it.for(OOXML_TYPES)("emits the %s package and main-part content type", async (type) => {
    const format = OOXML_PACKAGE_FORMATS[type];
    const bytes = (await generateOffice(
      type as OfficeGenerateFormat,
      PACKAGE_FIXTURES[format.family] as OfficeOptionsFor[OfficeGenerateFormat],
      "uint8array",
    )) as Uint8Array;
    const contentTypes = new TextDecoder().decode(unzipSync(bytes)["[Content_Types].xml"]);

    expect(findOverride(contentTypes, format.mainPartPath)).toEqual(format.mainContentType);
  });
});

describe("generateOffice RTF", () => {
  it("writes RTF through the aggregate API", async () => {
    const bytes = (await generateOffice(
      "rtf",
      { sections: [{ children: [{ paragraph: { children: ["Aggregate"] } }] }] },
      "uint8array",
    )) as Uint8Array;
    expect(new TextDecoder().decode(bytes)).toContain("Aggregate");
  });
});
