import { unzipSync, zipSync } from "@office-open/core";
import type { WorkbookOptions } from "@parts/file";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { parseWorkbookSync } from "../parse";

const MACRO_SHEET_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
  ` xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
  ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
  ` mc:Ignorable="x14ac xr xr2 xr3 xr6"` +
  ` xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"` +
  ` xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"` +
  ` xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"` +
  ` xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"` +
  ` xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"` +
  ` xr6:uid="{00000000-0001-0000-0000-000000000000}">` +
  `<dimension ref="A1:A2"/>` +
  `<sheetData>` +
  `<row r="1"><c r="A1" t="b"><f>ALERT("sample")</f><v>0</v></c></row>` +
  `<row r="2"><c r="A2" t="b"><f>HALT()</f><v>1</v></c></row>` +
  `</sheetData>` +
  `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
  `</xm:macrosheet>`;

async function buildMacroWorkbook(): Promise<Uint8Array> {
  const source = (await generateWorkbook(
    { worksheets: [{ name: "Sheet1" }] },
    { type: "uint8array" },
  )) as Uint8Array;
  const archive = unzipSync(source);
  const encoder = new TextEncoder();
  archive["xl/macrosheets/sheet1.xml"] = encoder.encode(MACRO_SHEET_XML);
  const workbook = new TextDecoder().decode(archive["xl/workbook.xml"]!);
  archive["xl/workbook.xml"] = encoder.encode(
    workbook.replace(/<\/sheets>/, `<sheet name="Macro1" sheetId="2" r:id="rIdMacro"/></sheets>`),
  );
  const rels = new TextDecoder().decode(archive["xl/_rels/workbook.xml.rels"]!);
  archive["xl/_rels/workbook.xml.rels"] = encoder.encode(
    rels.replace(
      /<\/Relationships>/,
      `<Relationship Id="rIdMacro" ` +
        `Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" ` +
        `Target="macrosheets/sheet1.xml"/></Relationships>`,
    ),
  );
  const types = new TextDecoder().decode(archive["[Content_Types].xml"]!);
  archive["[Content_Types].xml"] = encoder.encode(
    types.replace(
      /<\/Types>/,
      `<Override PartName="/xl/macrosheets/sheet1.xml" ` +
        `ContentType="application/vnd.ms-excel.macrosheet+xml"/></Types>`,
    ),
  );
  return zipSync(archive);
}

describe("macrosheets", () => {
  it("parses Excel 4.0 macro sheets with sheet identity and XLM formulas", async () => {
    const opts = parseWorkbookSync(await buildMacroWorkbook());
    const macro = opts.macrosheets?.[0];
    expect(macro).toBeDefined();
    expect(macro?.name).toBe("Macro1");
    expect(macro?.sheetId).toBe(2);
    expect(macro?.sourcePath).toBe("xl/macrosheets/sheet1.xml");
    expect(macro?.uid).toBe("{00000000-0001-0000-0000-000000000000}");
    expect(macro?.uidPrefix).toBe("xr6");
    expect(macro?.rows?.[0]?.cells?.[0]?.formula).toMatchObject({ formula: `ALERT("sample")` });
    expect(macro?.rows?.[1]?.cells?.[0]?.formula).toMatchObject({ formula: "HALT()" });
  });

  it("round-trips macro sheet content through generate and re-parse", async () => {
    const opts: WorkbookOptions = parseWorkbookSync(await buildMacroWorkbook());
    const output = unzipSync((await generateWorkbook(opts, { type: "uint8array" })) as Uint8Array);
    const macroXml = new TextDecoder().decode(output["xl/macrosheets/sheet1.xml"]!);
    expect(macroXml).toContain("<xm:macrosheet");
    expect(macroXml).toContain(`xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"`);
    expect(macroXml).toContain(`xr6:uid="{00000000-0001-0000-0000-000000000000}"`);
    expect(macroXml).toContain(`<f>ALERT(&quot;sample&quot;)</f>`);
    expect(macroXml).toContain("<f>HALT()</f>");
    const rels = new TextDecoder().decode(output["xl/_rels/workbook.xml.rels"]!);
    expect(rels).toContain("relationships/xlMacrosheet");
    expect(rels).toContain('Target="macrosheets/sheet1.xml"');
    const types = new TextDecoder().decode(output["[Content_Types].xml"]!);
    expect(types).toContain("application/vnd.ms-excel.macrosheet+xml");
    const reparsed = parseWorkbookSync(zipSync(output));
    const macro = reparsed.macrosheets?.[0];
    expect(macro?.name).toBe("Macro1");
    expect(macro?.uid).toBe("{00000000-0001-0000-0000-000000000000}");
    expect(macro?.uidPrefix).toBe("xr6");
    expect(macro?.rows?.[0]?.cells?.[0]?.formula).toMatchObject({ formula: `ALERT("sample")` });
  });
});
