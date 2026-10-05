import { describe, expect, it } from "vite-plus/test";

import { generateDocument, parseDocument } from "./document";
import { detectOffice } from "./formats";
import { generateOffice } from "./generate";
import * as officeOpen from "./index";
import { generatePresentation, parsePresentation } from "./presentation";
import { SCHEMAS } from "./schemas";
import { generateWorkbook, parseWorkbook } from "./workbook";

const documentOptions = { sections: [{ children: [{ paragraph: "Family" }] }] };
const workbookOptions = { worksheets: [{ rows: [{ cells: [{ value: "Family" }] }] }] };
const presentationOptions = {
  slides: [
    {
      children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "Family" } } }],
    },
  ],
};

describe("family APIs", () => {
  it("generates and parses each OpenDocument family", async () => {
    const document = await parseDocument(
      await generateDocument("odt", documentOptions, "uint8array"),
    );
    const workbook = await parseWorkbook(
      await generateWorkbook("ods", workbookOptions, "uint8array"),
    );
    const presentation = await parsePresentation(
      await generatePresentation("odp", presentationOptions, "uint8array"),
    );

    expect(document.sections).toHaveLength(1);
    expect(workbook.worksheets).toHaveLength(1);
    expect(presentation.slides).toHaveLength(1);
  });

  it("dispatches canonical formats through the root generator", async () => {
    const bytes = (await generateOffice("odt", documentOptions, "uint8array")) as Uint8Array;
    expect(detectOffice(bytes).format).toBe("odt");
  });

  it("detects text input through the root detector", () => {
    expect(detectOffice("{\\rtf1\\ansi Family}").family).toBe("document");
    expect(() => detectOffice(new TextEncoder().encode("not office"))).toThrow("Unable to detect");
  });

  it("exposes only the aggregate public runtime API", () => {
    expect(Object.keys(officeOpen).sort()).toEqual([
      "detectOffice",
      "generateDocument",
      "generateOffice",
      "generatePresentation",
      "generateWorkbook",
      "parseDocument",
      "parseOffice",
      "parsePresentation",
      "parseWorkbook",
      "patchDocument",
      "patchPresentation",
      "patchWorkbook",
    ]);
  });

  it("keeps canonical ODF inputs on the three OOXML schemas", () => {
    expect(Object.keys(SCHEMAS).sort()).toEqual(["docx", "pptx", "xlsx"]);
  });
});
