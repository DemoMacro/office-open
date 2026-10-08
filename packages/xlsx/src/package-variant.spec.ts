import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook, generateWorkbookStream, generateWorkbookSync } from "./generate";
import type { WorkbookOptions } from "./parts/file";

const workbook: WorkbookOptions = {
  worksheets: [{ name: "Sheet1", rows: [{ cells: [{ reference: "A1", value: "Variant" }] }] }],
};

const VARIANTS = ["standard", "macro", "template", "macroTemplate"] as const;

type Variant = (typeof VARIANTS)[number];

const ENCRYPTED_WORKBOOK: WorkbookOptions = {
  encrypted: {
    data: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
  },
};

const CONTENT_TYPES: Record<Variant, string> = {
  standard: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  macro: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  template: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  macroTemplate: "application/vnd.ms-excel.template.macroEnabled.main+xml",
};

const LEGACY_TEMPLATE_TYPE = "application/vnd.ms-excel.template.main+xml";

const BLOB_TYPES: Record<Variant, string> = {
  standard: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  macro: "application/vnd.ms-excel.sheet.macroenabled.12",
  template: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
  macroTemplate: "application/vnd.ms-excel.template.macroenabled.12",
};

const decodeContentTypes = (buffer: Uint8Array): string =>
  new TextDecoder().decode(unzipSync(buffer)["[Content_Types].xml"]!);

const collectStream = async (options: WorkbookOptions, variant: Variant): Promise<Uint8Array> => {
  const chunks: Uint8Array[] = [];
  const stream = generateWorkbookStream(options, { packageVariant: variant });
  await stream.pipeTo(
    new WritableStream<Uint8Array>({
      write: (chunk) => {
        chunks.push(chunk);
      },
    }),
  );
  return new Uint8Array(chunks.flatMap((chunk) => Array.from(chunk)));
};

describe("workbook package variants", () => {
  it.each(VARIANTS)("declares %s workbook content", async (variant) => {
    const buffer = (await generateWorkbook(workbook, {
      packageVariant: variant,
      type: "uint8array",
    })) as Uint8Array;

    expect(decodeContentTypes(buffer)).toContain(
      `<Override PartName="/xl/workbook.xml" ContentType="${CONTENT_TYPES[variant]}"/>`,
    );
  });

  it.each(VARIANTS)("sets %s Blob MIME type", async (variant) => {
    const blob = await generateWorkbook(workbook, { packageVariant: variant, type: "blob" });
    expect(blob.type).toBe(BLOB_TYPES[variant]);
  });

  it.each(VARIANTS)("sets %s encrypted passthrough Blob MIME type", async (variant) => {
    const blob = generateWorkbookSync(ENCRYPTED_WORKBOOK, {
      packageVariant: variant,
      type: "blob",
    });
    expect(blob.type).toBe(BLOB_TYPES[variant]);
  });

  it.each(VARIANTS)("streams %s workbook content", async (variant) => {
    const buffer = await collectStream(workbook, variant);

    expect(decodeContentTypes(buffer)).toContain(
      `<Override PartName="/xl/workbook.xml" ContentType="${CONTENT_TYPES[variant]}"/>`,
    );
  });
});

describe("workbook legacy content types", () => {
  it("preserves the legacy template main content type on round-trip", async () => {
    const buffer = (await generateWorkbook(
      {
        ...workbook,
        contentTypes: {
          defaults: [],
          overrides: [{ partName: "/xl/workbook.xml", contentType: LEGACY_TEMPLATE_TYPE }],
        },
      },
      { packageVariant: "template", type: "uint8array" },
    )) as Uint8Array;

    expect(decodeContentTypes(buffer)).toContain(
      `<Override PartName="/xl/workbook.xml" ContentType="${LEGACY_TEMPLATE_TYPE}"/>`,
    );
  });
});
