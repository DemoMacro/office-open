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

import type { DocumentOptions } from "../../packages/docx/src";
import { generateOdp, parseOdp } from "../../packages/odp/src";
import { generateOds, parseOds } from "../../packages/ods/src";
import { generateOdt, parseOdt } from "../../packages/odt/src";
import type { PresentationOptions } from "../../packages/pptx/src";
import type { WorkbookOptions } from "../../packages/xlsx/src";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const MAIN_SCHEMA = path.resolve(__dirname, "../../odf-schemas/OpenDocument-v1.3-schema.rng");
const MANIFEST_SCHEMA = path.resolve(
  __dirname,
  "../../odf-schemas/OpenDocument-v1.3-manifest-schema.rng",
);

const ODT_DOCUMENT: DocumentOptions = {
  title: "ODF schema check",
  creator: "office-open",
  created: "2026-01-02T03:04:05Z",
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

const SAMPLES: Array<[string, Uint8Array]> = [
  ["odt", generateOdt(ODT_DOCUMENT)],
  ["odt-empty", generateOdt({ sections: [] })],
  ["ods", generateOds(ODS_WORKBOOK)],
  ["ods-empty", generateOds({})],
  ["odp", generateOdp(ODP_PRESENTATION)],
  ["odp-empty", generateOdp({})],
];

const roundTrips: Array<[string, Uint8Array]> = [
  ["odt", generateOdt(parseOdt(SAMPLES[0]![1]))],
  ["ods", generateOds(parseOds(SAMPLES[2]![1]))],
  ["odp", generateOdp(parseOdp(SAMPLES[4]![1]))],
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
