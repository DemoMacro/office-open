import { readFileSync } from "node:fs";
import { readdir } from "node:fs/promises";
import { join } from "node:path";

import { describe, expect, it } from "vite-plus/test";

import packageJson from "../package.json";

const FORBIDDEN_NAMES = [
  "ChartChartOptions",
  "Element",
  "ElementCompact",
  "OdfXmlNode",
  "OdpOptions",
  "OdsOptions",
  "OdtOptions",
  "odfExtensions",
] as const;

const RUNTIME_EXPORTS: Record<string, readonly string[]> = {
  ".": [
    "detectOfficeFile",
    "detectOfficeFormat",
    "generate",
    "generateDocumentFile",
    "generateOfficeDocument",
    "generatePresentationFile",
    "generateToFile",
    "generateWorkbookFile",
    "parseDocumentFile",
    "parseInput",
    "parseOfficeDocument",
    "parsePresentationFile",
    "parseWorkbookFile",
  ],
  "./ai": [
    "docxTool",
    "formatToolError",
    "odpTool",
    "odsTool",
    "odtTool",
    "officeOpenTools",
    "pptxTool",
    "schemaLookupTool",
    "xlsxTool",
  ],
  "./convert": [
    "csvToWorkbook",
    "fromDrawingParagraph",
    "toDocxConnector",
    "toDocxGroup",
    "toDocxPicture",
    "toDocxShape",
    "toDocxSmartArt",
    "toDocxTable",
    "toDrawingParagraph",
    "toPptxConnector",
    "toPptxGroup",
    "toPptxPicture",
    "toPptxShape",
    "toPptxSmartArt",
    "toPptxTable",
    "toXlsxConnector",
    "toXlsxGroup",
    "toXlsxPicture",
    "toXlsxShape",
    "toXlsxTable",
    "tsvToWorkbook",
    "workbookToCsv",
    "workbookToTsv",
  ],
  "./core": [
    "convertEmuToInches",
    "convertEmuToPixels",
    "convertEmuToPoints",
    "convertEmuToTwip",
    "convertInchesToEmu",
    "convertInchesToTwip",
    "convertMillimetersToTwip",
    "convertOutput",
    "convertPixelsToEmu",
    "convertPointsToEmu",
    "convertPointsToTwip",
    "convertToEmu",
    "convertToInch",
    "convertToPt",
    "convertToTwip",
    "convertTwipToEmu",
    "convertUniversalMeasureToEmu",
    "convertUniversalMeasureToInch",
    "convertUniversalMeasureToPt",
    "convertUniversalMeasureToTwip",
  ],
  "./detect": ["detectOfficeFile"],
  "./doc": ["DocParseError", "parseDocument"],
  "./docx": [
    "generateDocument",
    "generateDocumentStream",
    "generateDocumentSync",
    "parseDocument",
    "parseDocumentSync",
    "patchDocument",
  ],
  "./formats": ["detectOfficeFormat"],
  "./generate": ["generate", "generateOfficeDocument", "generateToFile", "parseInput"],
  "./odf": [
    "generateOdp",
    "generateOds",
    "generateOdt",
    "OdpParseError",
    "OdsParseError",
    "OdtParseError",
    "parseOdp",
    "parseOds",
    "parseOdt",
  ],
  "./parse": ["parseOfficeDocument"],
  "./ppt": ["PptParseError", "parsePresentation"],
  "./pptx": [
    "generatePresentation",
    "generatePresentationStream",
    "generatePresentationSync",
    "parsePresentation",
    "parsePresentationSync",
    "patchPresentation",
  ],
  "./presentation": ["generatePresentationFile", "parsePresentationFile"],
  "./rtf": ["parseRtf", "RtfParseError"],
  "./schemas": [
    "assertKnownDefinitions",
    "docxSchema",
    "entryNames",
    "getSkeletonSchema",
    "pptxSchema",
    "renderSliceTypeText",
    "SCHEMAS",
    "SCHEMA_ENTRIES",
    "SCHEMA_ROOTS",
    "sliceDocumentSchema",
    "sliceSchema",
    "UnknownDefinitionError",
    "validateDocumentInput",
    "xlsxSchema",
  ],
  "./workbook": ["generateWorkbookFile", "parseWorkbookFile"],
  "./xls": ["parseWorkbook", "XlsParseError"],
  "./xml": [
    "allChildren",
    "attr",
    "attrBool",
    "attrMeasure",
    "attrNum",
    "attrs",
    "attrsRaw",
    "childCount",
    "children",
    "childText",
    "collectText",
    "colorAttr",
    "element",
    "escapeXml",
    "findChild",
    "findDeep",
    "findFirst",
    "hasChild",
    "isNonEmpty",
    "nativeTypeValue",
    "OOXML_XML_DECLARATION",
    "parse",
    "parseAttributes",
    "selfCloseElement",
    "stringify",
    "stringifyElement",
    "textOf",
    "unescapeXml",
  ],
  "./xlsx": [
    "generateWorkbook",
    "generateWorkbookStream",
    "generateWorkbookSync",
    "parseWorkbook",
    "parseWorkbookSync",
    "parseXlsx",
    "patchWorkbook",
  ],
};

const SOURCE_PATHS: Record<string, string> = {
  ".": "src/index.ts",
  "./docx": "src/docx.ts",
  "./pptx": "src/pptx.ts",
  "./xlsx": "src/xlsx.ts",
  "./document": "src/document.ts",
  "./workbook": "src/workbook.ts",
  "./presentation": "src/presentation.ts",
  "./doc": "src/doc.ts",
  "./xls": "src/xls.ts",
  "./ppt": "src/ppt.ts",
  "./rtf": "src/rtf.ts",
  "./odf": "src/odf.ts",
  "./formats": "src/formats.ts",
  "./detect": "src/detect.ts",
  "./parse": "src/parse.ts",
  "./core": "src/core.ts",
  "./xml": "src/xml.ts",
  "./ai": "src/ai/index.ts",
  "./generate": "src/generate.ts",
  "./schemas": "src/schemas/index.ts",
  "./convert": "src/convert/index.ts",
};

const DIST_PATHS: Record<string, string> = {
  ".": "dist/index.d.mts",
  "./ai": "dist/ai/index.d.mts",
  "./convert": "dist/convert/index.d.mts",
  "./core": "dist/core.d.mts",
  "./detect": "dist/detect.d.mts",
  "./doc": "dist/doc.d.mts",
  "./docx": "dist/docx.d.mts",
  "./formats": "dist/formats.d.mts",
  "./generate": "dist/generate.d.mts",
  "./odf": "dist/odf.d.mts",
  "./parse": "dist/parse.d.mts",
  "./ppt": "dist/ppt.d.mts",
  "./pptx": "dist/pptx.d.mts",
  "./presentation": "dist/presentation.d.mts",
  "./rtf": "dist/rtf.d.mts",
  "./schemas": "dist/schemas/index.d.mts",
  "./workbook": "dist/workbook.d.mts",
  "./xls": "dist/xls.d.mts",
  "./xml": "dist/xml.d.mts",
  "./xlsx": "dist/xlsx.d.mts",
};

function sourceTypeNames(path: string): Set<string> {
  const names = new Set<string>();
  const source = readFileSync(path, "utf8");
  for (const match of source.matchAll(/export type \{([\s\S]*?)\}/g)) {
    for (const token of match[1]!.split(",")) {
      const name = token.trim();
      if (name) names.add(name);
    }
  }
  for (const match of source.matchAll(/export (?:declare )?(?:interface|type) ([A-Za-z0-9_]+)/g)) {
    names.add(match[1]!);
  }
  return names;
}

function declaredNames(path: string): { names: Set<string>; typeNames: Set<string> } {
  const text = readFileSync(path, "utf8");
  const names = new Set<string>();
  const typeNames = new Set<string>();
  for (const match of text.matchAll(/export (type )?\{([\s\S]*?)\};/g)) {
    for (const token of match[2]!.split(",")) {
      const name = token
        .trim()
        .replace(/^type /, "")
        .replace(/ as .*$/, "");
      if (!name) continue;
      (match[1] ? typeNames : names).add(name);
    }
  }
  return { names, typeNames };
}

describe("public aggregate API", () => {
  it("freezes package.json export paths and conditions", () => {
    expect(Object.keys(packageJson.exports)).toEqual(Object.keys(SOURCE_PATHS));
    for (const entry of Object.values(packageJson.exports)) {
      expect(Object.keys(entry)).toEqual(["types", "import"]);
      expect(entry.types).toMatch(/^\.\/dist\/.*\.d\.mts$/);
      expect(entry.import).toMatch(/^\.\/dist\/.*\.mjs$/);
    }
  });

  it("uses explicit named exports in every aggregate entry", async () => {
    const files = await readdir("src", { recursive: true });
    const offenders = files
      .filter((file) => file.endsWith(".ts") && !file.endsWith(".spec.ts"))
      .map((file) => ({ file, source: readFileSync(join("src", file), "utf8") }))
      .filter(({ source }) => /export (?:type )?\*/.test(source))
      .map(({ file }) => file);
    expect(offenders).toEqual([]);
  });

  it("freezes runtime exports for every package subpath", async () => {
    for (const [entry, expected] of Object.entries(RUNTIME_EXPORTS)) {
      const packageRoot = new URL("../", import.meta.url);
      const modulePath = new URL(SOURCE_PATHS[entry]!, packageRoot).href;
      const module = await import(/* @vite-ignore */ modulePath);
      expect(Object.keys(module).sort(), entry).toEqual([...expected].sort());
      for (const forbidden of FORBIDDEN_NAMES) {
        expect(Object.keys(module), `${entry}:${forbidden}`).not.toContain(forbidden);
      }
    }
  });

  it("freezes declared type exports and never leaks forbidden names", () => {
    for (const [entry, path] of Object.entries(DIST_PATHS)) {
      const declared = declaredNames(path);
      const declaredAll = new Set([...declared.names, ...declared.typeNames]);
      const expected = new Set([
        ...sourceTypeNames(SOURCE_PATHS[entry]!),
        ...RUNTIME_EXPORTS[entry]!,
      ]);
      expect(declaredAll, entry).toEqual(expected);
      const forbiddenNames = entry === "." ? FORBIDDEN_NAMES : [];
      for (const forbidden of forbiddenNames) {
        expect(declared.names.has(forbidden), `${entry}:${forbidden}`).toBe(false);
        expect(declared.typeNames.has(forbidden), `${entry}:${forbidden}`).toBe(false);
      }
    }
    const rootDeclaration = readFileSync(DIST_PATHS["."]!, "utf8");
    for (const forbidden of FORBIDDEN_NAMES) {
      expect(rootDeclaration).not.toMatch(new RegExp(`\\b${forbidden}\\b`));
    }
  });
});
