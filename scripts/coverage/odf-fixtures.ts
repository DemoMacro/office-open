import {
  drawingVocabularyXml,
  generateChartDocument,
  generateDatabaseDocument,
  generateOcf,
  officeFormsXml,
  manifestOptionsXml,
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

export const ODF_COVERAGE_FIXTURES = {
  "manifest-package-entry"() {
    const source = {
      root: {
        name: "manifest:manifest",
        attributes: { "manifest:version": "1.3" },
        children: [
          {
            name: "manifest:file-entry",
            attributes: {
              "manifest:full-path": "/",
              "manifest:media-type": "application/vnd.oasis.opendocument.text",
            },
          },
        ],
      },
    };
    const parsed = JSON.parse(JSON.stringify(manifestRoundTrip(source)));
    assertEqual(parsed.root.name, source.root.name, "manifest root");
    assertEqual(parsed.root.attributes, source.root.attributes, "manifest attributes");
    assertEqual(parsed.root.children[0].name, source.root.children[0].name, "manifest child");
    assertEqual(
      parsed.root.children[0].attributes,
      source.root.children[0].attributes,
      "manifest child attributes",
    );
  },
  "drawing-vocabulary-desc"() {
    const source = element("<svg:desc>Drawing</svg:desc>");
    const parsed = parseDrawingVocabulary(source, "content.xml", "/svg:desc");
    assert(!("contentXml" in parsed), "expected the drawing vocabulary model");
    assertEqual(parsed.name, "svg:desc", "drawing vocabulary element");
    assertEqual(parsed.text, "Drawing", "drawing vocabulary text");
    assert(drawingVocabularyXml(parsed).includes("Drawing"), "drawing vocabulary XML");
  },
  "odt-paragraph"() {
    const parsed = parseDocument(
      generateDocument({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
    );
    assertEqual(collectText(parsed.sections), ["ODT"], "ODT paragraph");
  },
  "ods-cell"() {
    const parsed = parseWorkbook(
      generateWorkbook({ worksheets: [{ rows: [{ cells: [{ value: 1.5 }] }] }] }),
    );
    assertEqual(parsed.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value, 1.5, "ODS cell value");
  },
  "odp-notes"() {
    const parsed = parsePresentation(generatePresentation({ slides: [{ notes: "ODP" }] }));
    assertEqual(parsed.slides?.[0]?.notes, "ODP", "ODP notes");
  },
  "chart-column-series"() {
    const source = {
      title: "Chart",
      chart: { type: "column", series: [{ name: "Sales", values: [1, 2] }] },
    } as const;
    const parsed = parseChartDocument(generateChartDocument(source));
    assertEqual(parsed.title, source.title, "chart title");
    assertEqual(parsed.chart.series[0]?.name, source.chart.series[0].name, "chart series name");
    assertEqual(parsed.chart.series[0]?.values, source.chart.series[0].values, "chart values");
  },
  "database-connection"() {
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
  },
  "form-grid-xforms"() {
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
  },
} as const;

export const ODF_NEGATIVE_FIXTURES = {
  "manifest rejects foreign element"() {
    assertThrows(
      () =>
        manifestRoundTrip({
          root: {
            name: "manifest:manifest",
            children: [{ name: "office:unknown" }],
          },
        }),
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

export function runOdfCoverageFixture(testId: keyof typeof ODF_COVERAGE_FIXTURES): void {
  ODF_COVERAGE_FIXTURES[testId]();
}

export function runOdfNegativeFixture(testId: keyof typeof ODF_NEGATIVE_FIXTURES): void {
  ODF_NEGATIVE_FIXTURES[testId]();
}
