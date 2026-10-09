import type { ChartSpaceOptions } from "@office-open/core";
import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import {
  type ChartDocumentOptions,
  type DatabaseDocumentOptions,
  generateChartDocument,
  generateDatabaseDocument,
  officeFormsXml,
  OdfSchemaError,
  parseOfficeForms,
  parseChartBody,
  parseChartDocument,
  parseDatabaseDocument,
} from "./index";

describe("ODF schema codecs", () => {
  it("round-trips chart subdocuments through the shared chart model", () => {
    const chart: ChartDocumentOptions = {
      title: "Chart",
      chart: { type: "column", series: [{ name: "Sales", values: [1, 2] }] },
    };
    const parsed = parseChartDocument(generateChartDocument(chart));
    expect(parsed.title).toBe(chart.title);
    expect(parsed.chart.type).toBe("column");
    expect(parsed.chart.series[0]).toMatchObject(chart.chart.series[0]!);
  });

  it("round-trips exact bar direction and local category/value caches", () => {
    const chart: ChartSpaceOptions = {
      type: "bar",
      categories: ["North", "South"],
      series: [{ name: "Sales", values: [10, 20] }],
    } as const;
    const parsed = parseChartDocument(generateChartDocument({ chart })).chart;
    expect(parsed.type).toBe("bar");
    expect(parsed.categories).toEqual(chart.categories);
    expect(parsed.series[0]).toMatchObject({ name: "Sales", values: [10, 20] });
  });

  it("renders sparse category cache points in the local table", () => {
    const parsed = parseChartDocument(
      generateChartDocument({
        chart: {
          type: "column",
          categories: ["First", { index: 1, text: "Second" }],
          series: [{ name: "Sales", values: [1, 2] }],
        },
      }),
    ).chart;
    expect(parsed.categories).toEqual(["First", "Second"]);
  });

  it("round-trips scatter X/Y values and bubble sizes", () => {
    const scatter = {
      type: "scatter",
      series: [{ name: "Points", xValues: [1, 2], yValues: [3, 5] }],
    } as const;
    expect(parseChartDocument(generateChartDocument({ chart: scatter })).chart).toMatchObject(
      scatter,
    );

    const bubble = {
      type: "bubble",
      series: [{ name: "Bubbles", xValues: [1, 2], yValues: [3, 5], bubbleSize: [4, 6] }],
    } as const;
    expect(parseChartDocument(generateChartDocument({ chart: bubble })).chart).toMatchObject(
      bubble,
    );
  });

  it("round-trips chart titles, legends, axes, series, points, and surfaces", () => {
    const chart: ChartSpaceOptions = {
      type: "column",
      title: { text: "Sales", shapeProperties: { fill: { type: "solid", color: "FFF7E6" } } },
      shapeProperties: { fill: { type: "solid", color: "F3F3F3" } },
      plotAreaShapeProperties: {
        fill: { type: "solid", color: "FFFFFF" },
        outline: { type: "solidFill", color: "CCCCCC", width: 9525 },
      },
      showLegend: true,
      legendPosition: "top",
      legendOverlay: true,
      legendShapeProperties: { outline: { type: "solidFill", color: "999999" } },
      axes: [
        {
          kind: "category",
          title: "Region",
          tickLabelPosition: "none",
          majorGridlines: true,
          shapeProperties: { fill: { type: "solid", color: "EEEEEE" } },
        },
        { kind: "value", title: "Amount", delete: true, minorGridlines: true },
      ],
      sideWall: { shapeProperties: { fill: { type: "solid", color: "D9E2F3" } } },
      floor: { shapeProperties: { fill: { type: "solid", color: "E2F0D9" } } },
      series: [
        {
          name: "Sales",
          values: [1, 2],
          shapeProperties: { fill: { type: "solid", color: "4472C4" } },
          dataPoints: [
            { index: 0, shapeProperties: { fill: { type: "solid", color: "FF0000" } } },
            { index: 1, shapeProperties: { fill: { type: "solid", color: "FF0000" } } },
          ],
        },
      ],
    };
    expect(parseChartDocument(generateChartDocument({ chart })).chart).toMatchObject(chart);
  });

  it("preserves external chart formulas while retaining cached model values", () => {
    const chart: ChartSpaceOptions = {
      type: "column",
      categoryFormula: "Sheet1!$A$2:$A$3",
      series: [
        {
          name: undefined,
          nameFormula: "Sheet1!$B$1",
          valueFormula: "Sheet1!$B$2:$B$3",
          values: [],
        },
      ],
    };
    expect(parseChartDocument(generateChartDocument({ chart })).chart).toMatchObject(chart);
  });

  it("reports unsupported chart elements with structured schema errors", () => {
    const body = `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area><chart:axis chart:dimension="x"/><chart:axis chart:dimension="y"/><chart:foo/></chart:plot-area></chart:chart></office:chart></office:body></office:document-content>`;
    let error: unknown;
    try {
      parseChartBody(body);
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OdfSchemaError);
    expect(error).toMatchObject({
      part: "content.xml",
      path: "/office:document-content/office:body/office:chart/chart:chart[chart:plot-area]/chart:foo",
      name: "chart:foo",
      reason: "element has no canonical ChartSpaceOptions mapping",
    });
  });

  it("round-trips database subdocuments as a constrained database model", () => {
    const database = {
      title: "Database",
      dataSource: {
        name: "db:data-source",
        children: [
          {
            name: "db:connection-data",
            children: [{ name: "db:database-description", text: "Library catalog" }],
          },
          {
            name: "db:forms",
            forms: {
              forms: [
                {
                  name: "Search",
                  controls: [
                    {
                      kind: "form:combobox",
                      id: "search-box",
                      name: "query",
                      properties: [{ name: "ReadOnly", value: { type: "boolean", value: true } }],
                      events: [
                        { eventName: "change", language: "Basic", macroName: "Search.Change" },
                      ],
                      options: [{ attributes: { "form:value": "books" }, label: "Books" }],
                      xformsModel: {
                        attributes: { id: "search-model" },
                        children: [
                          { name: "xforms:instance", attributes: { id: "data" }, text: "books" },
                        ],
                      },
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    };
    const parsed = parseDatabaseDocument(
      generateDatabaseDocument(database as DatabaseDocumentOptions),
    );
    expect(parsed.dataSource).toMatchObject(database.dataSource);
    expect(parsed.dataSource.children?.[1]?.forms?.forms[0]?.controls[0]?.kind).toBe(
      "form:combobox",
    );
  });

  it("round-trips form properties, events, options, columns, and XForms", () => {
    const forms = {
      forms: [
        {
          name: "Library",
          controls: [
            {
              kind: "form:grid" as const,
              id: "grid",
              columns: [
                {
                  attributes: { "form:name": "title" },
                  controls: [
                    {
                      kind: "form:text" as const,
                      id: "title",
                      attributes: { "form:disabled": true },
                    },
                  ],
                  properties: [
                    { name: "DataField", value: { type: "string" as const, value: "title" } },
                  ],
                  events: [{ eventName: "click", language: "Basic", macroName: "Grid.Click" }],
                },
              ],
            },
          ],
          properties: [
            { name: "DataSource", value: { type: "string" as const, value: "Library" } },
          ],
          events: [{ eventName: "load", language: "Basic", macroName: "Form.Load" }],
        },
      ],
    };
    expect(
      parseOfficeForms(parse(officeFormsXml(forms), { ignoreDeclaration: true }).elements?.[0]),
    ).toMatchObject(forms);
  });

  it("rejects unknown form elements and attributes with structured diagnostics", () => {
    const namespaces =
      'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"';
    const unknownElement = parse(
      `<office:forms ${namespaces}><form:form><form:unknown form:id="x"/></form:form></office:forms>`,
      { ignoreDeclaration: true },
    ).elements?.[0];
    expect(() => parseOfficeForms(unknownElement)).toThrow(OdfSchemaError);
    const unknownAttribute = parse(
      `<office:forms ${namespaces}><form:form form:unknown="x"/></office:forms>`,
      { ignoreDeclaration: true },
    ).elements?.[0];
    expect(() => parseOfficeForms(unknownAttribute)).toThrow(OdfSchemaError);
  });
});
