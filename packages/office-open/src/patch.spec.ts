import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument, patchDocument } from "./document";
import { parseDocument } from "./document";
import { generatePresentation, patchPresentation } from "./presentation";
import { generateWorkbook, patchWorkbook } from "./workbook";

describe("canonical patch dispatch", () => {
  it("patches RTF and regenerates canonical content", async () => {
    const patched = await patchDocument({
      format: "rtf",
      data: "{\\rtf1\\ansi Original}",
      commands: [
        {
          type: "replace",
          at: { field: "sections", route: [{ index: 0 }], text: { literal: "Original" } },
          value: "Patched RTF",
        },
      ],
    });
    const parsed = await parseDocument(new TextDecoder().decode(patched as Uint8Array));
    expect(parsed.sections?.[0]?.children?.[0]).toEqual({
      paragraph: { children: [{ text: "Patched RTF" }] },
    });
  });

  it("patches ODT and preserves source-only package members", async () => {
    const source = await generateDocument(
      "odt",
      {
        sections: [{ children: [{ paragraph: "ODT" }] }],
        packageMembers: [
          { path: "audit.bin", mediaType: "application/octet-stream", data: new Uint8Array([7]) },
        ],
      },
      "uint8array",
    );
    const patched = (await patchDocument({
      format: "odt",
      data: source,
      commands: [
        {
          type: "replace",
          at: { field: "sections", route: [{ index: 0 }], text: { literal: "ODT" } },
          value: "Patched ODT",
        },
      ],
    })) as Uint8Array;
    expect(unzipSync(patched)["audit.bin"]).toEqual(new Uint8Array([7]));
  });

  it("patches ODS canonical cells", async () => {
    const source = await generateWorkbook(
      "ods",
      { worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] },
      "uint8array",
    );
    const patched = (await patchWorkbook({
      format: "ods",
      data: source,
      commands: [
        {
          type: "replace",
          at: { field: "worksheets", route: [{ index: 0 }], text: { literal: "ODS" } },
          value: "Patched ODS",
        },
      ],
    })) as Uint8Array;
    expect(new TextDecoder().decode(unzipSync(patched)["content.xml"])).toContain("Patched ODS");
  });

  it("patches ODP canonical text", async () => {
    const source = await generatePresentation(
      "odp",
      {
        slides: [
          {
            children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "ODP" } } }],
          },
        ],
      },
      "uint8array",
    );
    const patched = (await patchPresentation({
      format: "odp",
      data: source,
      commands: [
        {
          type: "replace",
          at: { field: "slides", route: [{ index: 0 }], text: { literal: "ODP" } },
          value: "Patched ODP",
        },
      ],
    })) as Uint8Array;
    expect(new TextDecoder().decode(unzipSync(patched)["content.xml"])).toContain("Patched ODP");
  });
});
