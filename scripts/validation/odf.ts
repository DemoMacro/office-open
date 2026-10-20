/**
 * Validate generated ODF documents against the official ODF 1.3 Relax NG
 * schemas (odf-schemas/). For each sample the script validates the original
 * output and its parse → generate round trip, checking every XML part:
 *   - META-INF/manifest.xml            → manifest schema
 *   - content.xml, styles.xml, meta.xml → main schema
 *
 * Usage: npx tsx scripts/validation/odf.ts
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";
import { RelaxNGValidator, XmlDocument } from "libxml2-wasm";

import type { FormControlOptions, OfficeFormsOptions } from "../../packages/core/src";
import type { DocumentOptions } from "../../packages/docx/src";
import { generateChartDocument, parseChartDocument } from "../../packages/odf/src";
import { generatePresentation, parsePresentation } from "../../packages/odp/src";
import { generateWorkbook, parseWorkbook } from "../../packages/ods/src";
import { generateDocument, parseDocument } from "../../packages/odt/src";
import type { PresentationOptions } from "../../packages/pptx/src";
import type { WorkbookOptions } from "../../packages/xlsx/src";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const MAIN_SCHEMA = path.resolve(__dirname, "../../odf-schemas/OpenDocument-v1.3-schema.rng");
const MANIFEST_SCHEMA = path.resolve(
  __dirname,
  "../../odf-schemas/OpenDocument-v1.3-manifest-schema.rng",
);

const FORM_CONTROLS: FormControlOptions[] = [
  { control: "text", id: "text", xmlId: "text", value: "", currentValue: "", maxLength: 0 },
  { control: "textArea", id: "textArea", xmlId: "textArea", paragraphs: ["Line"] },
  {
    control: "formattedText",
    id: "formattedText",
    xmlId: "formattedText",
    minValue: "0",
    maxValue: "9",
    validation: false,
  },
  {
    control: "number",
    id: "number",
    xmlId: "number",
    value: 0,
    currentValue: 1,
    minValue: 0,
    maxValue: 9,
  },
  { control: "date", id: "date", xmlId: "date", value: "2026-01-31", currentValue: "2026-02-01" },
  { control: "time", id: "time", xmlId: "time", value: "01:02:03" },
  { control: "comboBox", id: "comboBox", xmlId: "comboBox", items: [{ label: "" }] },
  {
    control: "listBox",
    id: "listBox",
    xmlId: "listBox",
    options: [{ label: "", currentSelected: false, selected: false, value: "" }],
  },
  {
    control: "checkbox",
    id: "checkbox",
    xmlId: "checkbox",
    currentState: "unknown",
    state: "unknown",
    triState: true,
  },
  { control: "password", id: "password", xmlId: "password", echoChar: "*" },
  { control: "file", id: "file", xmlId: "file", currentValue: "a.odt" },
  { control: "fixedText", id: "fixedText", xmlId: "fixedText", label: "", multiLine: true },
  {
    control: "button",
    id: "button",
    xmlId: "button",
    buttonType: "push",
    defaultButton: false,
    toggle: false,
    focusOnClick: false,
  },
  { control: "image", id: "image", xmlId: "image", buttonType: "url", href: "logo.png" },
  { control: "radio", id: "radio", xmlId: "radio", currentSelected: false, selected: false },
  { control: "frame", id: "frame", xmlId: "frame", label: "Frame" },
  { control: "imageFrame", id: "imageFrame", xmlId: "imageFrame", imageData: "logo.png" },
  { control: "hidden", id: "hidden", xmlId: "hidden", value: "" },
  {
    control: "grid",
    id: "grid",
    xmlId: "grid",
    columns: [{ name: "title", controls: [{ control: "text", id: "title", xmlId: "title" }] }],
  },
  {
    control: "valueRange",
    id: "valueRange",
    xmlId: "valueRange",
    minValue: 0,
    maxValue: 100,
    stepSize: 1,
    pageStepSize: 10,
  },
  { control: "genericControl", id: "genericControl", xmlId: "genericControl" },
];

const ODF_FORMS: OfficeFormsOptions = {
  automaticFocus: false,
  designMode: false,
  containers: [{ name: "Library", href: "Forms/Search.odt", controls: FORM_CONTROLS }],
};

const ODT_DOCUMENT: DocumentOptions = {
  title: "ODF schema check",
  creator: "office-open",
  created: "2026-01-02T03:04:05Z",
  forms: ODF_FORMS,
  sections: [
    {
      children: [
        { paragraph: "Heading paragraph" },
        {
          paragraph: {
            alignment: "center",
            children: [
              { text: "Bold", bold: true },
              { text: " and " },
              { text: "italic", italic: true },
            ],
          },
        },
        {
          table: {
            rows: [
              { cells: [{ children: [{ paragraph: "A1" }] }, { children: [{ paragraph: "B1" }] }] },
              { cells: [{ children: [{ paragraph: "A2" }] }, { children: [{ paragraph: "B2" }] }] },
            ],
          },
        },
      ],
    },
  ],
};

const ODS_WORKBOOK: WorkbookOptions = {
  title: "ODF schema check",
  forms: ODF_FORMS,
  worksheets: [
    {
      name: "Data",
      columns: [{ min: 1, max: 2, width: 120 }],
      rows: [
        { height: 18, cells: [{ value: "Text" }, { value: true }] },
        { cells: [{ value: 4.5 }, { value: 9, formula: "SUM(A1:B1)" }] },
      ],
    },
    { name: "Empty" },
  ],
};

const ODP_PRESENTATION: PresentationOptions = {
  size: { width: 9144000, height: 6858000 },
  slides: [
    {
      forms: ODF_FORMS,
      children: [
        {
          shape: {
            name: "Title",
            x: 914400,
            y: 914400,
            width: 3657600,
            height: 914400,
            textBody: {
              paragraphs: [
                { text: "Slide text" },
                { children: [{ text: "Bold", bold: true, size: 20 }] },
              ],
            },
          },
        },
      ],
    },
  ],
};

const ODF_CHART = {
  title: "ODF chart schema check",
  chart: {
    type: "column",
    series: [{ name: "Sales", values: [1, 2] }],
    shapeProperties: { styleName: "chart-style" },
    styleResources: {
      styles: [
        {
          name: "chart-style",
          attributes: { "style:family": "chart" },
          properties: [
            {
              name: "style:graphic-properties",
              attributes: {
                "draw:fill": "gradient",
                "draw:fill-gradient-name": "gradient",
                "draw:stroke": "dash",
                "draw:stroke-dash": "dash",
                "svg:stroke-width": "0.1cm",
              },
            },
          ],
        },
      ],
      gradients: [
        {
          name: "gradient",
          attributes: {
            "draw:style": "linear",
            "draw:start-color": "#ffffff",
            "draw:end-color": "#000000",
          },
        },
      ],
      hatches: [{ name: "hatch", attributes: { "draw:style": "single" } }],
      fillImages: [
        {
          name: "image",
          attributes: { "xlink:href": "media/image.png", "svg:width": "1cm", "svg:height": "1cm" },
          data: new Uint8Array([137, 80, 78, 71]),
        },
      ],
      strokeDashes: [
        {
          name: "dash",
          attributes: {
            "draw:style": "rect",
            "draw:dots1": "1",
            "draw:dots1-length": "1cm",
            "draw:distance": "0.05cm",
          },
        },
      ],
    },
  },
} as const;

const SAMPLES: Array<[string, Uint8Array]> = [
  ["odt", await generateDocument(ODT_DOCUMENT)],
  ["odt-empty", await generateDocument({ sections: [] })],
  ["ods", await generateWorkbook(ODS_WORKBOOK)],
  ["ods-empty", await generateWorkbook({})],
  ["odp", await generatePresentation(ODP_PRESENTATION)],
  ["odp-empty", await generatePresentation({})],
  ["odf-chart", await generateChartDocument(ODF_CHART)],
];

const roundTrips: Array<[string, Uint8Array]> = [
  ["odt", await generateDocument(await parseDocument(SAMPLES[0]![1]))],
  ["ods", await generateWorkbook(await parseWorkbook(SAMPLES[2]![1]))],
  ["odp", await generatePresentation(await parsePresentation(SAMPLES[4]![1]))],
  ["odf-chart", await generateChartDocument(await parseChartDocument(SAMPLES[6]![1]))],
];

const mainValidator = RelaxNGValidator.fromDoc(
  XmlDocument.fromBuffer(new Uint8Array(readFileSync(MAIN_SCHEMA))),
);
const manifestValidator = RelaxNGValidator.fromDoc(
  XmlDocument.fromBuffer(new Uint8Array(readFileSync(MANIFEST_SCHEMA))),
);

let pass = 0;
let fail = 0;
const failures: string[] = [];

for (const [label, data] of [
  ...SAMPLES,
  ...roundTrips.map(([name, data]) => [`${name} (round-trip)`, data] as [string, Uint8Array]),
]) {
  for (const [entry, bytes] of Object.entries(unzipSync(data))) {
    if (!entry.endsWith(".xml")) continue;
    const validator = entry === "META-INF/manifest.xml" ? manifestValidator : mainValidator;
    try {
      const document = XmlDocument.fromBuffer(bytes);
      validator.validate(document);
      document.dispose();
      pass += 1;
      console.log(`  PASS  ${label} ${entry}`);
    } catch (error) {
      fail += 1;
      const message = String(error).split("\n")[0];
      failures.push(`${label} ${entry}: ${message}`);
      console.error(`  FAIL  ${label} ${entry}: ${message}`);
    }
  }
}

mainValidator.dispose();
manifestValidator.dispose();

console.log(`\n${pass} passed, ${fail} failed`);
if (fail > 0) process.exitCode = 1;
