import type {
  ChartSpaceOptions,
  CorePropertiesOptions,
  FormContainerOptions,
  FormControlOptions,
  OfficeFormsOptions,
} from "@office-open/core";
import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import {
  type ChartDocumentOptions,
  type DatabaseDocumentOptions,
  generateChartDocument,
  generateDatabaseDocument,
  officeFormsXml,
  OdfSchemaError,
  metaXml,
  parseOfficeForms,
  parseMeta,
  parseChartBody,
  parseChartDocument,
  parseDatabaseDocument,
} from "./index";

describe("ODF schema codecs", () => {
  it("round-trips the complete ODF metadata surface and rejects unknown metadata", () => {
    const files = {
      "meta.xml": metaXml({
        creator: "Original Author",
        lastModifiedBy: "Latest Author",
        created: "2026-01-01T00:00:00Z",
        modified: "2026-01-02T00:00:00Z",
      }),
    };
    expect(parseMeta(files)).toMatchObject({
      creator: "Original Author",
      lastModifiedBy: "Latest Author",
    });
    const metadata: CorePropertiesOptions = {
      application: "Application/1.0",
      printedBy: "Printer",
      editingDuration: { hours: 1, minutes: 2, seconds: 3.5 },
      template: {
        href: "https://example.invalid/template.ots",
        actuate: "onRequest",
        title: "Template",
        modified: "2026-01-03T00:00:00Z",
      },
      autoReload: { href: "https://example.invalid/reload", show: "replace", delay: { days: 1 } },
      hyperlinkBehavior: { targetFrameName: "_blank", show: "new" },
      statistics: { pageCount: 2, wordCount: 12, nonWhitespaceCharacterCount: 70 },
      userDefinedProperties: [
        { name: "Number", type: "float", value: 1.5 },
        { name: "Date", type: "date", value: "2026-01-01T00:00:00Z" },
        { name: "Duration", type: "time", value: { minutes: 3 } },
        { name: "Flag", type: "boolean", value: true },
        { name: "Text", type: "string", value: "Value" },
        { name: "Plain", type: "text", value: "Plain" },
      ],
    };
    expect(parseMeta({ "meta.xml": metaXml(metadata) })).toMatchObject(metadata);
    expect(
      parseMeta({
        "meta.xml":
          '<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><meta:user-defined meta:name=""/></office:meta></office:document-meta>',
      }),
    ).toEqual({ userDefinedProperties: [{ name: "", type: "text", value: "" }] });
    expect(() =>
      parseMeta({
        "meta.xml":
          '<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><meta:unknown>Value</meta:unknown></office:meta></office:document-meta>',
      }),
    ).toThrow(OdfSchemaError);
  });

  it("maps legacy Dublin Core initial creators to canonical creators", () => {
    expect(
      parseMeta({
        "meta.xml":
          '<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:initial-creator>Original Author</dc:initial-creator></office:meta></office:document-meta>',
      }),
    ).toEqual({ creator: "Original Author" });
  });

  it("round-trips chart subdocuments through the shared chart model", async () => {
    const chart: ChartDocumentOptions = {
      title: "Chart",
      chart: { type: "column", series: [{ name: "Sales", values: [1, 2] }] },
    };
    const parsed = await parseChartDocument(generateChartDocument(chart));
    expect(parsed.title).toBe(chart.title);
    expect(parsed.chart.type).toBe("column");
    expect(parsed.chart.series[0]).toMatchObject(chart.chart.series[0]!);
  });

  it("round-trips exact bar direction and local category/value caches", async () => {
    const chart: ChartSpaceOptions = {
      type: "bar",
      categories: ["North", "South"],
      series: [{ name: "Sales", values: [10, 20] }],
    } as const;
    const parsed = (await parseChartDocument(generateChartDocument({ chart }))).chart;
    expect(parsed.type).toBe("bar");
    expect(parsed.categories).toEqual(chart.categories);
    expect(parsed.series[0]).toMatchObject({ name: "Sales", values: [10, 20] });
  });

  it("renders sparse category cache points in the local table", async () => {
    const parsed = await parseChartDocument(
      generateChartDocument({
        chart: {
          type: "column",
          categories: ["First", { index: 1, text: "Second" }],
          series: [{ name: "Sales", values: [1, 2] }],
        },
      }),
    );
    expect(parsed.chart.categories).toEqual(["First", "Second"]);
  });

  it("round-trips scatter X/Y values and bubble sizes", async () => {
    const scatter = {
      type: "scatter",
      series: [{ name: "Points", xValues: [1, 2], yValues: [3, 5] }],
    } as const;
    expect(
      (await parseChartDocument(generateChartDocument({ chart: scatter }))).chart,
    ).toMatchObject(scatter);

    const bubble = {
      type: "bubble",
      series: [{ name: "Bubbles", xValues: [1, 2], yValues: [3, 5], bubbleSize: [4, 6] }],
    } as const;
    expect(
      (await parseChartDocument(generateChartDocument({ chart: bubble }))).chart,
    ).toMatchObject(bubble);
  });

  it("round-trips chart titles, legends, axes, series, points, and surfaces", async () => {
    const chart: ChartSpaceOptions = {
      type: "column",
      title: { text: "Sales", shapeProperties: { fill: { type: "solid", color: "FFF7E6" } } },
      shapeProperties: { fill: { type: "solid", color: "F3F3F3" } },
      plotAreaShapeProperties: {
        fill: { type: "solid", color: "FFFFFF" },
        outline: { type: "solidFill", color: "CCCCCC", width: 9525 },
      },
      coordinateRegion: { x: 380000, y: 140000, width: 4400000, height: 3100000 },
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
    expect((await parseChartDocument(generateChartDocument({ chart }))).chart).toMatchObject(chart);
  });

  it("preserves external chart formulas while retaining cached model values", async () => {
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
    expect((await parseChartDocument(generateChartDocument({ chart }))).chart).toMatchObject(chart);
  });

  it("round-trips ODF coordinate regions and date axes", async () => {
    const chart: ChartSpaceOptions = {
      type: "column",
      categories: ["2024", "2025"],
      coordinateRegion: { x: 380000, y: 140000, width: 4400000, height: 3100000 },
      axes: [{ kind: "date", title: "Date" }],
      series: [{ name: "Sales", values: [1, 2] }],
    };
    const parsed = (await parseChartDocument(generateChartDocument({ chart }))).chart;
    expect(parsed.coordinateRegion).toEqual(chart.coordinateRegion);
    expect(parsed.axes?.[0]).toMatchObject({ kind: "date", title: "Date" });
  });

  it("round-trips ODF subtitles and footers", async () => {
    const chart: ChartSpaceOptions = {
      type: "column",
      subtitle: "Subtitle",
      footer: "Footer",
      series: [{ name: "Sales", values: [1, 2] }],
    };
    const parsed = (await parseChartDocument(generateChartDocument({ chart }))).chart;
    expect(parsed.subtitle).toEqual(chart.subtitle);
    expect(parsed.footer).toEqual(chart.footer);
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
      elementName: "chart:foo",
      reason: "element has no canonical ChartSpaceOptions mapping",
    });
  });

  it("round-trips database forms as component references", async () => {
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
            children: [
              {
                name: "db:component",
                attributes: {
                  "db:name": "Search",
                  "xlink:type": "simple",
                  "xlink:href": "Forms/Search.odt",
                },
              },
            ],
          },
        ],
      },
    };
    const parsed = await parseDatabaseDocument(
      generateDatabaseDocument(database as DatabaseDocumentOptions),
    );
    expect(parsed.dataSource).toMatchObject(database.dataSource);
  });

  it("round-trips every form control and shared form structure", () => {
    const controls: FormControlOptions[] = [
      {
        control: "text",
        id: "text",
        value: "",
        currentValue: "",
        maxLength: 0,
        disabled: false,
        tabIndex: 0,
      },
      { control: "textArea", id: "textArea", paragraphs: ["Line"] },
      {
        control: "formattedText",
        id: "formattedText",
        minValue: "1",
        maxValue: "9",
        validation: true,
      },
      {
        control: "number",
        id: "number",
        value: 1.5,
        currentValue: 2.5,
        minValue: 0,
        maxValue: 10,
        spinButton: true,
      },
      {
        control: "date",
        id: "date",
        value: "2026-01-31",
        currentValue: "2026-02-01",
        minValue: "2025-01-01",
        maxValue: "2027-01-01",
      },
      {
        control: "time",
        id: "time",
        value: "01:02:03",
        currentValue: "04:05:06",
        minValue: "00:00:00",
        maxValue: "23:59:59",
      },
      {
        control: "comboBox",
        id: "comboBox",
        items: [{ label: "Books" }],
        dropdown: true,
        autoComplete: true,
        listSourceType: "valueList",
      },
      {
        control: "listBox",
        id: "listBox",
        options: [{ label: "Books", value: "books", currentSelected: true, selected: false }],
        multiple: true,
        listLinkageType: "selectionIndices",
      },
      {
        control: "checkbox",
        id: "checkbox",
        currentState: "unknown",
        triState: true,
        state: "unknown",
      },
      { control: "password", id: "password", echoChar: "*" },
      { control: "file", id: "file", currentValue: "a.odt" },
      { control: "fixedText", id: "fixedText", label: "Fixed", multiLine: true },
      {
        control: "button",
        id: "button",
        buttonType: "push",
        defaultButton: false,
        toggle: true,
        focusOnClick: true,
      },
      { control: "image", id: "image", buttonType: "url", href: "logo.png" },
      { control: "radio", id: "radio", currentSelected: true, selected: false },
      { control: "frame", id: "frame", label: "Frame" },
      { control: "imageFrame", id: "imageFrame", imageData: "logo.png", readonly: true },
      { control: "hidden", id: "hidden", value: "hidden" },
      {
        control: "grid",
        id: "grid",
        columns: [{ name: "title", controls: [{ control: "text", id: "title", xmlId: "title" }] }],
      },
      {
        control: "valueRange",
        id: "valueRange",
        minValue: 0,
        maxValue: 100,
        stepSize: 1,
        pageStepSize: 10,
        orientation: "horizontal",
      },
      { control: "genericControl", id: "genericControl" },
    ];
    const nested: FormContainerOptions = {
      name: "Nested",
      controls: [{ control: "hidden", id: "nestedHidden", value: "Nested" }],
    };
    const forms: OfficeFormsOptions = {
      automaticFocus: false,
      designMode: false,
      containers: [
        {
          name: "Library",
          href: "Forms/Search.odt",
          commandType: "table",
          tabCycle: "records",
          properties: [
            { name: "Float", value: { type: "float", value: 1.5 } },
            { name: "Percentage", value: { type: "percentage", value: 50 } },
            { name: "Currency", value: { type: "currency", value: 2, currency: "USD" } },
            { name: "Date", value: { type: "date", value: "2026-01-01T00:00:00Z" } },
            { name: "Time", value: { type: "time", value: "PT1H" } },
            { name: "Boolean", value: { type: "boolean", value: false } },
            { name: "String", value: { type: "string", value: "" } },
            { name: "Void", value: { type: "void" } },
            { name: "FloatList", value: { type: "float", values: [0, 1.5] } },
            { name: "StringList", value: { type: "string", values: ["", "A"] } },
            { name: "BooleanList", value: { type: "boolean", values: [false, true] } },
          ],
          events: [
            { target: "macro", eventName: "click", language: "Basic", macroName: "Form.Click" },
            {
              target: "url",
              eventName: "load",
              language: "URL",
              href: "script.odt",
              actuate: "onRequest",
            },
          ],
          controls,
          forms: [nested],
          connectionResource: { href: "database.odb" },
        },
      ],
      models: [
        {
          attributes: { id: "search-model" },
          children: [
            {
              name: "xforms:instance",
              attributes: { id: "data" },
              children: ["books", { name: "xforms:bind", attributes: { nodeset: "/data" } }],
            },
          ],
        },
      ],
    };
    const parsed = parseOfficeForms(
      parse(officeFormsXml(forms), { ignoreDeclaration: true }).elements?.[0],
    );
    const withIds = (source: FormControlOptions[]) =>
      source.map((control) => ({ ...control, xmlId: control.id }));
    expect(parsed).toEqual({
      ...forms,
      containers: [
        {
          ...forms.containers[0],
          controls: withIds(controls),
          forms: [{ ...nested, controls: withIds(nested.controls) }],
        },
      ],
    });
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

  it("preserves both control ids and derives missing xml ids deterministically", () => {
    const namespaces =
      'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace"';
    const source = `<office:forms ${namespaces}><form:form><form:text xml:id="xml-id" form:id="form-id" form:name=""/><form:text form:id="derived"/></form:form></office:forms>`;
    const parsed = parseOfficeForms(parse(source, { ignoreDeclaration: true }).elements?.[0]);
    expect(parsed.containers[0]?.controls).toEqual([
      { control: "text", xmlId: "xml-id", id: "form-id", name: "" },
      { control: "text", xmlId: "derived", id: "derived" },
    ]);
  });

  it("round-trips explicit false root form settings without containers", () => {
    const forms: OfficeFormsOptions = {
      automaticFocus: false,
      designMode: false,
      containers: [],
    };
    const parsed = parseOfficeForms(
      parse(officeFormsXml(forms), { ignoreDeclaration: true }).elements?.[0],
    );
    expect(parsed).toEqual(forms);
  });

  it("rejects malformed form values, ids, event placement, and grid controls structurally", () => {
    const namespaces =
      'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"';
    const cases: Array<[string, string, string, string]> = [
      [
        `<office:forms ${namespaces} form:automatic-focus="yes"/>`,
        "form:automatic-focus",
        "form:automatic-focus",
        "invalid boolean value: yes",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:date form:id="date" form:value="2026-02-30"/></form:form></office:forms>`,
        "form:value",
        "form:value",
        "invalid date value: 2026-02-30",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:time form:id="time" form:value="24:00:01"/></form:form></office:forms>`,
        "form:value",
        "form:value",
        "invalid time value: 24:00:01",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:listbox form:id="list" form:list-source-type="unknown"/></form:form></office:forms>`,
        "form:list-source-type",
        "form:list-source-type",
        "invalid listSourceType value: unknown",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:text form:id="1bad"/></form:form></office:forms>`,
        "form:id",
        "form:id",
        "invalid NCName",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:text form:id="same"/><form:text form:id="same"/></form:form></office:forms>`,
        "same",
        "same",
        "duplicate form control id",
      ],
      [
        `<office:forms ${namespaces}><form:form><script:event-listener script:event-name="click" script:language="Basic" script:macro-name="Click"/></form:form></office:forms>`,
        "script:event-listener",
        "script:event-listener",
        "child has no canonical form mapping",
      ],
      [
        `<office:forms ${namespaces}><form:form><form:grid form:id="grid"><form:column><form:button form:id="button"/></form:column></form:grid></form:form></office:forms>`,
        "form:button",
        "form:button",
        "control is not allowed in a form grid column",
      ],
    ];
    for (const [source, expectedPathSuffix, expectedName, expectedReason] of cases) {
      let error: unknown;
      try {
        parseOfficeForms(parse(source, { ignoreDeclaration: true }).elements?.[0]);
      } catch (cause) {
        error = cause;
      }
      expect(error).toBeInstanceOf(OdfSchemaError);
      expect(error).toMatchObject({
        part: "content.xml",
        elementName: expectedName,
        reason: expectedReason,
      });
      expect((error as OdfSchemaError).path.endsWith(expectedPathSuffix)).toBe(true);
    }
  });
});
