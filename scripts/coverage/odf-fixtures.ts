import {
  drawingVocabularyXml,
  generateChartDocument,
  generateDatabaseDocument,
  generateOcf,
  officeFormsXml,
  manifestOptionsXml,
  readOcf,
  parseChartBody,
  parseChartDocument,
  parseDatabaseDocument,
  parseDrawingVocabulary,
  parseOfficeForms,
  parseManifestOptions,
} from "@office-open/odf";
import { generatePresentation, parsePresentation } from "@office-open/odp";
import { generateWorkbook, parseWorkbook } from "@office-open/ods";
import { generateDocument, parseDocument } from "@office-open/odt";
import { parse, type Element } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

function assert(condition: unknown, message: string): void {
  if (!condition) throw new Error(message);
}

function assertEqual(actual: unknown, expected: unknown, label: string): void {
  const left = JSON.stringify(actual);
  const right = JSON.stringify(expected);
  assert(left === right, `${label}: expected ${right}, received ${left}`);
}

function assertThrows(run: () => unknown, expected: string): void {
  let error: unknown;
  try {
    run();
  } catch (cause) {
    error = cause;
  }
  assert(error instanceof Error, `expected an error containing ${expected}`);
  assert(error.message.includes(expected), `expected ${expected}, received ${error.message}`);
}

function element(xml: string): Element {
  const result = parse(xml, { ignoreDeclaration: true }).elements?.find(
    (candidate): candidate is Element => candidate.type === "element",
  );
  assert(result, "fixture XML has no root element");
  return result;
}

function unzipText(files: Record<string, Uint8Array>, path: string): string {
  return strFromU8(files[path]!);
}

export interface OdfCoverageFixtureResult {
  semanticRoundTrip: true;
  rawFallbackElements?: readonly string[];
}

export const ODF_COVERAGE_FIXTURES = {
  "manifest-package-entry"(): OdfCoverageFixtureResult {
    const source = {
      version: "1.3",
      entries: [
        {
          fullPath: "/",
          mediaType: "application/vnd.oasis.opendocument.text",
        },
        {
          fullPath: "content.xml",
          mediaType: "text/xml",
        },
      ],
    } as const;
    const parsed = manifestRoundTrip(source);
    assertEqual(parsed.version, source.version, "manifest version");
    assertEqual(parsed.entries, source.entries, "manifest entries");
    return { semanticRoundTrip: true };
  },
  "drawing-vocabulary-desc"(): OdfCoverageFixtureResult {
    const source = element("<svg:desc>Drawing</svg:desc>");
    const parsed = parseDrawingVocabulary(source, "content.xml", "/svg:desc");
    assert(!("contentXml" in parsed), "expected the drawing vocabulary model");
    assertEqual(parsed.name, "svg:desc", "drawing vocabulary element");
    assertEqual(parsed.text, "Drawing", "drawing vocabulary text");
    assert(drawingVocabularyXml(parsed).includes("Drawing"), "drawing vocabulary XML");
    return { semanticRoundTrip: true };
  },
  "odt-paragraph"(): OdfCoverageFixtureResult {
    const source = {
      title: "ODT ownership",
      creator: "office-open",
      created: "2026-01-02T03:04:05Z",
      styles: {
        paragraphStyles: [
          {
            id: "Custom",
            name: "Custom",
            paragraph: { alignment: "center" },
            run: { bold: true },
          },
        ],
      },
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ text: "Bold", bold: true }, { text: " ODT" }],
              },
            },
            {
              table: {
                rows: [{ cells: [{ children: [{ paragraph: "A1" }] }] }],
              },
            },
          ],
        },
      ],
    } as const;
    const generated = generateDocument(source);
    const files = unzipSync(generated);
    const odtContent = unzipText(files, "content.xml");
    assert(
      odtContent.includes('office:version="1.3"'),
      `ODT content root: ${odtContent.slice(0, 220)}`,
    );
    assert(unzipText(files, "styles.xml").includes("<style:style"), "ODT styles");
    assert(unzipText(files, "meta.xml").includes("<dc:title>ODT ownership</dc:title>"), "ODT meta");
    assert(odtContent.includes('fo:font-weight="bold"'), "ODT text style");
    const parsed = parseDocument(generateDocument(JSON.parse(JSON.stringify(source))));
    assertEqual(collectText(parsed.sections), ["Bold", " ODT", "A1"], "ODT semantic round trip");
    return { semanticRoundTrip: true };
  },
  "ods-cell"(): OdfCoverageFixtureResult {
    const source = {
      title: "ODS ownership",
      creator: "office-open",
      created: "2026-01-02T03:04:05Z",
      styles: {
        paragraphStyles: [
          {
            id: "Custom",
            name: "Custom",
            paragraph: { alignment: "center" },
            run: { bold: true },
          },
        ],
      },
      worksheets: [
        {
          name: "Data",
          columns: [{ min: 1, max: 1, width: 120 }],
          rows: [{ height: 18, cells: [{ value: 1.5, formula: "A1*2" }] }],
        },
      ],
    } as const;
    const generated = generateWorkbook(source);
    const files = unzipSync(generated);
    const odsContent = unzipText(files, "content.xml");
    assert(
      odsContent.includes('table:formula="of:=A1*2"'),
      `ODS formula: ${odsContent.slice(odsContent.indexOf("<table:table-cell"), odsContent.indexOf("<table:table-cell") + 300)}`,
    );
    assert(unzipText(files, "styles.xml").includes("<office:styles/>"), "ODS styles part");
    assert(odsContent.includes('style:row-height="18pt"'), "ODS row style");
    assert(unzipText(files, "meta.xml").includes("<dc:title>ODS ownership</dc:title>"), "ODS meta");
    const parsed = parseWorkbook(generated);
    const cell = parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0];
    assertEqual(cell?.value, 1.5, "ODS cell value");
    assertEqual(cell?.formula, "A1*2", "ODS formula round trip");
    return { semanticRoundTrip: true };
  },
  "odp-notes"(): OdfCoverageFixtureResult {
    const source = {
      title: "ODP ownership",
      creator: "office-open",
      created: "2026-01-02T03:04:05Z",
      styles: {
        paragraphStyles: [
          {
            id: "Custom",
            name: "Custom",
            paragraph: { alignment: "center" },
            run: { bold: true },
          },
        ],
      },
      slides: [
        {
          name: "Page",
          children: [
            {
              shape: {
                name: "Title",
                x: 914400,
                y: 914400,
                width: 3657600,
                height: 914400,
                textBody: { paragraphs: [{ text: "ODP" }] },
              },
            },
          ],
          notes: "Notes",
        },
      ],
    } as const;
    const generated = generatePresentation(source);
    const files = unzipSync(generated);
    const odpContent = unzipText(files, "content.xml");
    assert(
      odpContent.includes("<presentation:notes>"),
      `ODP notes part: ${odpContent.slice(odpContent.indexOf("<presentation:notes"), odpContent.indexOf("<presentation:notes") + 240)}`,
    );
    assert(unzipText(files, "styles.xml").includes("<office:styles/>"), "ODP styles part");
    assert(
      unzipText(files, "styles.xml").includes('style:page-layout-name="PM1"'),
      "ODP master-page style",
    );
    assert(unzipText(files, "meta.xml").includes("<dc:title>ODP ownership</dc:title>"), "ODP meta");
    const parsed = parsePresentation(generated);
    assertEqual(parsed.slides?.[0]?.notes, source.slides[0].notes, "ODP semantic round trip");
    assertEqual(
      collectText(parsed.slides?.[0]?.children),
      ["ODP"],
      "ODP shape semantic round trip",
    );
    return { semanticRoundTrip: true };
  },
  "chart-embedded-object"(): OdfCoverageFixtureResult {
    const chart = {
      name: "Embedded chart",
      type: "column",
      series: [{ name: "Sales", values: [1, 2] }],
      col: 1,
      row: 1,
      anchorType: "absolute",
      absoluteX: 360000,
      absoluteY: 720000,
      extentCx: 2160000,
      extentCy: 1440000,
    } as const;
    const generated = generateWorkbook({
      worksheets: [{ name: "Data", charts: [chart] }],
    });
    const parsed = parseWorkbook(generated);
    const parsedChart = parsed.worksheets?.[0]?.charts?.[0];
    assertEqual(parsedChart?.name, chart.name, "embedded chart name");
    assertEqual(
      parsedChart?.series,
      [{ ...chart.series[0], dataPoints: [] }],
      "embedded chart series",
    );
    const chartFiles = Object.keys(unzipSync(generated)).filter((path) =>
      path.endsWith("/content.xml"),
    );
    assert(chartFiles.length > 0, `chart subdocument part: ${chartFiles.join(", ")}`);
    return { semanticRoundTrip: true };
  },
  "database-formula-embedded-object"(): OdfCoverageFixtureResult {
    const databaseSource = {
      dataSource: {
        name: "db:data-source",
        children: [
          {
            name: "db:connection-data",
            children: [{ name: "db:database-description", text: "Library catalog" }],
          },
        ],
      },
    } as const;
    const parsedDatabase = parseDatabaseDocument(generateDatabaseDocument(databaseSource));
    assertEqual(
      parsedDatabase.dataSource.children?.[0]?.children?.[0]?.text,
      databaseSource.dataSource.children[0].children[0].text,
      "database semantic round trip",
    );
    const chart = {
      name: "Formula object",
      type: "bar",
      series: [{ name: "Computed", values: [9] }],
      col: 1,
      row: 1,
      anchorType: "absolute",
      absoluteX: 360000,
      absoluteY: 720000,
      extentCx: 1800000,
      extentCy: 1080000,
    } as const;
    const workbook = generateWorkbook({
      worksheets: [
        {
          name: "Data",
          rows: [{ cells: [{ value: 9, formula: "SUM(B1:B1)" }] }],
          charts: [chart],
        },
      ],
    });
    const parsedWorkbook = parseWorkbook(workbook);
    assertEqual(
      parsedWorkbook.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.formula,
      "SUM(B1:B1)",
      "formula semantic round trip",
    );
    assertEqual(
      parsedWorkbook.worksheets?.[0]?.charts?.[0]?.series,
      [{ ...chart.series[0], dataPoints: [] }],
      "embedded object semantic round trip",
    );
    return { semanticRoundTrip: true };
  },
  "chart-column-series"(): OdfCoverageFixtureResult {
    const source = {
      title: "Chart",
      chart: { type: "column", series: [{ name: "Sales", values: [1, 2] }] },
    } as const;
    const parsed = parseChartDocument(generateChartDocument(source));
    assertEqual(parsed.title, source.title, "chart title");
    assertEqual(parsed.chart.series[0]?.name, source.chart.series[0].name, "chart series name");
    assertEqual(parsed.chart.series[0]?.values, source.chart.series[0].values, "chart values");
    return { semanticRoundTrip: true };
  },
  "database-connection"(): OdfCoverageFixtureResult {
    const source = {
      title: "Database",
      dataSource: {
        name: "db:data-source",
        children: [
          {
            name: "db:connection-data",
            children: [{ name: "db:database-description", text: "Library catalog" }],
          },
        ],
      },
    } as const;
    const parsed = parseDatabaseDocument(generateDatabaseDocument(source));
    assertEqual(parsed.dataSource.name, source.dataSource.name, "database root");
    assertEqual(
      parsed.dataSource.children?.[0]?.children?.[0]?.name,
      source.dataSource.children[0].children[0].name,
      "database child",
    );
    assertEqual(
      parsed.dataSource.children?.[0]?.children?.[0]?.text,
      source.dataSource.children[0].children[0].text,
      "database description",
    );
    return { semanticRoundTrip: true };
  },
  "form-grid-xforms"(): OdfCoverageFixtureResult {
    const source = {
      forms: [
        {
          name: "Library",
          controls: [
            {
              kind: "form:grid",
              id: "grid",
              columns: [
                {
                  attributes: { "form:name": "title" },
                  controls: [{ kind: "form:text", id: "title" }],
                },
              ],
            },
          ],
        },
      ],
    } as const;
    const parsed = parseOfficeForms(element(officeFormsXml(source)));
    const control = parsed.forms[0]?.controls[0];
    assertEqual(parsed.forms[0]?.name, source.forms[0].name, "form name");
    assertEqual(control?.kind, source.forms[0].controls[0].kind, "form control kind");
    assertEqual(
      control?.columns?.[0]?.controls[0]?.kind,
      source.forms[0].controls[0].columns[0].controls[0].kind,
      "nested form control",
    );
    return { semanticRoundTrip: true };
  },
  "encrypted-manifest"(): OdfCoverageFixtureResult {
    const source = {
      version: "1.3",
      entries: [
        {
          fullPath: "content.xml",
          mediaType: "text/xml",
          encryptionData: {
            checksumType: "SHA1/1K",
            checksum: "aa==",
            algorithm: { name: "Blowfish CFB", initialisationVector: "bb==" },
            startKeyGeneration: { name: "SHA256", keySize: 32 },
            keyDerivation: {
              name: "PBKDF2",
              salt: "cc==",
              iterationCount: 1024,
              keySize: 16,
            },
          },
        },
      ],
      encryptedKeys: [{ algorithm: "RSA", keyId: "MQ==", keyPacket: "Ag==", cipherValue: "Aw==" }],
    } as const;
    const parsed = manifestRoundTrip(source);
    assertEqual(parsed.entries, source.entries, "encrypted manifest entries");
    assertEqual(parsed.encryptedKeys, source.encryptedKeys, "encrypted manifest keys");
    return { semanticRoundTrip: true };
  },
  "linked-graphics"(): OdfCoverageFixtureResult {
    const image = {
      type: "image",
      href: "https://example.com/logo.png",
      sourceUrl: "https://example.com/logo.png",
      width: 914400,
      height: 914400,
    } as const;
    const generated = generateWorkbook({
      worksheets: [{ rows: [{ cells: [{ value: "Logo", graphics: [image] }] }] }],
    });
    assertEqual(
      Object.keys(readOcf(generated, "application/vnd.oasis.opendocument.spreadsheet").binaries),
      [],
      "linked image has no media part",
    );
    const parsed = parseWorkbook(generated);
    const graphic = parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.graphics?.[0];
    assertEqual(graphic?.sourceUrl, image.sourceUrl, "linked image source");
    assertEqual(graphic?.data, undefined, "linked image has no bytes");
    return { semanticRoundTrip: true };
  },
  "percent-length-overlay"(): OdfCoverageFixtureResult {
    const overlays = [
      {
        name: "ro99",
        family: "table-row",
        properties: [
          {
            name: "style:table-row-properties",
            attributes: { "style:row-height": 20, "style:use-optimal-row-height": false },
          },
        ],
      },
      {
        name: "P1",
        family: "paragraph",
        properties: [{ name: "style:paragraph-properties", attributes: { "fo:line-height": 150 } }],
      },
    ] as const;
    const generated = generateWorkbook({
      worksheets: [{ rows: [{ cells: [{ value: "Overlay" }] }] }],
      styleOverlays: overlays,
    });
    const content = unzipSync(generated)["content.xml"]!;
    const contentXml = strFromU8(content);
    assert(contentXml.includes('style:row-height="20pt"'), "length overlay scale");
    assert(contentXml.includes('fo:line-height="150%"'), "percent overlay scale");
    const referenced = zipSync({
      ...unzipSync(generated),
      "content.xml": strToU8(
        contentXml.replace("<table:table-row>", '<table:table-row table:style-name="ro99">'),
      ),
    });
    const parsed = parseWorkbook(referenced);
    assertEqual(parsed.worksheets?.[0]?.rows?.[0]?.height, 20, "raw length fallback semantics");
    assertEqual(parsed.styleOverlays, [overlays[1]], "raw percent overlay semantic round trip");
    return {
      semanticRoundTrip: true,
      rawFallbackElements: ["style:table-row-properties", "style:paragraph-properties"],
    };
  },
} as const;

export const ODF_NEGATIVE_FIXTURES = {
  "manifest rejects foreign element"() {
    assertThrows(
      () =>
        parseManifestOptions(
          manifestOptionsXml({
            version: "1.3",
            entries: [{ fullPath: "content.xml", mediaType: "text/xml" }],
          }).replace("</manifest:manifest>", "<office:unknown/></manifest:manifest>"),
        ),
      "Unknown manifest element",
    );
  },
  "drawing vocabulary rejects foreign element"() {
    assertThrows(
      () => parseDrawingVocabulary(element("<draw:unknown/>"), "content.xml", "/draw:unknown"),
      "no canonical drawing vocabulary mapping",
    );
  },
  "ODT rejects foreign body element"() {
    assertThrows(
      () =>
        parseDocument(
          replaceOdfXml(
            generateDocument({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
            "</office:text>",
            "<office:unknown/></office:text>",
          ),
        ),
      "no canonical",
    );
  },
  "ODS rejects foreign body element"() {
    assertThrows(
      () =>
        parseWorkbook(
          replaceOdfXml(
            generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: 1.5 }] }] }] }),
            "</office:spreadsheet>",
            "<office:unknown/></office:spreadsheet>",
          ),
        ),
      "no canonical",
    );
  },
  "ODP rejects foreign slide element"() {
    assertThrows(
      () =>
        parsePresentation(
          replaceOdfXml(
            generatePresentation({ slides: [{ notes: "ODP" }] }),
            "</draw:page>",
            "<draw:unknown/></draw:page>",
          ),
        ),
      "no canonical",
    );
  },
  "chart rejects foreign element"() {
    assertThrows(
      () => parseChartBody(chartWithForeignElement()),
      "no canonical ChartSpaceOptions mapping",
    );
  },
  "database rejects foreign element"() {
    assertThrows(
      () =>
        parseDatabaseDocument(
          generateDatabaseDocument({
            dataSource: {
              name: "db:data-source",
              children: [{ name: "db:unknown" }],
            },
          }),
        ),
      "no canonical database model mapping",
    );
  },
  "forms reject foreign element"() {
    assertThrows(
      () =>
        parseOfficeForms(
          element("<office:forms><form:form><form:unknown/></form:form></office:forms>"),
        ),
      "no canonical",
    );
  },
} as const;

function manifestRoundTrip(source: unknown): unknown {
  return parseManifestOptions(manifestOptionsXml(source));
}

function replaceOdfXml(data: Uint8Array, search: string, replacement: string): Uint8Array {
  const files = unzipSync(data);
  const content = strFromU8(files["content.xml"]!).replace(search, replacement);
  return zipSync({ ...files, "content.xml": strToU8(content) });
}

function chartWithForeignElement(): string {
  const chart = 'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"';
  const office = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"';
  return `<?xml version="1.0"?><office:document-content ${office} ${chart}><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area><chart:foo/></chart:plot-area></chart:chart></office:chart></office:body></office:document-content>`;
}

function collectText(value: unknown, result: string[] = []): string[] {
  if (Array.isArray(value)) {
    for (const child of value) collectText(child, result);
  } else if (typeof value === "object" && value !== null) {
    for (const [key, child] of Object.entries(value)) {
      if (key === "text" && typeof child === "string") result.push(child);
      else collectText(child, result);
    }
  }
  return result.filter(Boolean);
}

export function runOdfCoverageFixture(
  testId: keyof typeof ODF_COVERAGE_FIXTURES,
): OdfCoverageFixtureResult {
  return ODF_COVERAGE_FIXTURES[testId]();
}

export function runOdfNegativeFixture(testId: keyof typeof ODF_NEGATIVE_FIXTURES): void {
  ODF_NEGATIVE_FIXTURES[testId]();
}
