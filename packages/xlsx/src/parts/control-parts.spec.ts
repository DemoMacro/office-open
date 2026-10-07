import { unzipSync, zipSync } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "../generate";
import { parseWorkbookSync } from "../parse";
import { activeXControlDesc } from "./active-x-control";
import { controlPropertiesDesc } from "./control-properties";
import { buildWorksheetXml, worksheetDesc } from "./worksheet";
import type { WorksheetOptions } from "./worksheet";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CONTROL_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const CTRL_PROPS_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const ACTIVE_X_BINARY_REL =
  "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";

const SHEET_XML =
  XML_DECL +
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/>` +
  `<oleObjects/><controls>` +
  `<control shapeId="1" r:id="rId1" name="Choice"><controlPr locked="0" defaultSize="0" ` +
  `print="0" disabled="1" recalcAlways="1" uiObject="1" autoFill="0" autoLine="0" autoPict="0" ` +
  `macro="Module1.Run" altText="Choose item" linkedCell="Sheet1!$A$1" ` +
  `listFillRange="Sheet1!$B$1:$B$4" cf="pict" r:id="rId3">` +
  `<anchor moveWithCells="1"><from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>` +
  `<to><xdr:col>2</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>3</xdr:row>` +
  `<xdr:rowOff>20</xdr:rowOff></to></anchor></controlPr></control>` +
  `<control shapeId="2" r:id="rId2" name="Command1"><controlPr>` +
  `<anchor><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>` +
  `<xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>3</xdr:col><xdr:colOff>1</xdr:colOff>` +
  `<xdr:row>4</xdr:row><xdr:rowOff>1</xdr:rowOff></to></anchor></controlPr></control>` +
  `</controls></worksheet>`;

const FORM_CONTROL_XML =
  XML_DECL +
  `<x14:formControlPr xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ` +
  `objectType="Drop" checked="Mixed" dropStyle="Combo" min="1" max="20" inc="2" page="4" val="7" ` +
  `textAlign="Center" vertical="Top" lockText="1" noThreeD="1" noThreeD2="1">` +
  `<fmlaLink>Sheet1!$A$1</fmlaLink><fmlaRange>Sheet1!$B$1:$B$4</fmlaRange>` +
  `<fmlaMacro>[0]!Module1.Run</fmlaMacro><item>Alpha</item><item>Beta</item>` +
  `</x14:formControlPr>`;

const ACTIVE_X_XML =
  XML_DECL +
  `<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" ` +
  `ax:classid="{8BD21D10-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistPropertyBag" ` +
  `ax:license="test-license" r:id="rId1">` +
  `<ax:ocxPr ax:name="_ExtentX" ax:value="1561"/>` +
  `<ax:ocxPr ax:name="_ExtentY" ax:value="582"/></ax:ocx>`;

async function fixtureWorkbook(): Promise<Uint8Array> {
  const source = (await generateWorkbook(
    { worksheets: [{ name: "Controls" }] },
    { type: "uint8array" },
  )) as Uint8Array;
  const archive = unzipSync(source);
  archive["xl/worksheets/sheet1.xml"] = new TextEncoder().encode(SHEET_XML);
  archive["xl/worksheets/_rels/sheet1.xml.rels"] = new TextEncoder().encode(
    XML_DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${CTRL_PROPS_REL}" Target="../ctrlProps/ctrlProp1.xml"/>` +
      `<Relationship Id="rId2" Type="${CONTROL_REL}" Target="../activeX/activeX1.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/icon.png"/>` +
      `</Relationships>`,
  );
  archive["xl/ctrlProps/ctrlProp1.xml"] = new TextEncoder().encode(FORM_CONTROL_XML);
  archive["xl/activeX/activeX1.xml"] = new TextEncoder().encode(ACTIVE_X_XML);
  archive["xl/activeX/_rels/activeX1.xml.rels"] = new TextEncoder().encode(
    XML_DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${ACTIVE_X_BINARY_REL}" Target="activeX1.bin"/>` +
      `</Relationships>`,
  );
  archive["xl/activeX/activeX1.bin"] = new Uint8Array([1, 2, 3]);
  archive["xl/media/icon.png"] = new Uint8Array([4, 5, 6]);
  return zipSync(archive);
}

describe("worksheet control parts", () => {
  it("allocates fresh control parts globally across worksheets", async () => {
    const output = unzipSync(
      (await generateWorkbook(
        {
          worksheets: [
            {
              name: "Form",
              controls: [
                {
                  kind: "form",
                  shapeId: 1,
                  formControlProperties: { objectType: "Checkbox" },
                },
              ],
            },
            {
              name: "ActiveX",
              controls: [
                {
                  kind: "activeX",
                  shapeId: 2,
                  activeXControl: { classId: "{TEST-GLOBAL}" },
                },
              ],
            },
          ],
        },
        { type: "uint8array" },
      )) as Uint8Array,
    );
    expect(output["xl/ctrlProps/ctrlProp1.xml"]).toBeDefined();
    expect(output["xl/activeX/activeX1.xml"]).toBeDefined();
    expect(new TextDecoder().decode(output["xl/worksheets/_rels/sheet1.xml.rels"]!)).toContain(
      'Target="../ctrlProps/ctrlProp1.xml"',
    );
    expect(new TextDecoder().decode(output["xl/worksheets/_rels/sheet2.xml.rels"]!)).toContain(
      'Target="../activeX/activeX1.xml"',
    );
    expect(new TextDecoder().decode(output["[Content_Types].xml"]!)).toContain(
      'PartName="/xl/ctrlProps/ctrlProp1.xml"',
    );
    expect(new TextDecoder().decode(output["[Content_Types].xml"]!)).toContain(
      'PartName="/xl/activeX/activeX1.xml"',
    );
  });

  it("round-trips full CT_ControlPr, x14:formControlPr, and ActiveX metadata", async () => {
    const parsed = parseWorkbookSync(await fixtureWorkbook());
    const controls = parsed.worksheets?.[0]?.controls;
    expect(controls).toHaveLength(2);
    expect(controls?.[0]).toMatchObject({
      kind: "form",
      shapeId: 1,
      rId: "rId1",
      name: "Choice",
      properties: {
        locked: false,
        defaultSize: false,
        print: false,
        disabled: true,
        recalcAlways: true,
        uiObject: true,
        autoFill: false,
        autoLine: false,
        autoPict: false,
        macro: "Module1.Run",
        altText: "Choose item",
        linkedCell: "Sheet1!$A$1",
        listFillRange: "Sheet1!$B$1:$B$4",
        cf: "pict",
        iconRid: "rId3",
      },
      formControlProperties: {
        objectType: "Drop",
        checked: "Mixed",
        dropStyle: "Combo",
        fmlaLink: "Sheet1!$A$1",
        fmlaRange: "Sheet1!$B$1:$B$4",
        fmlaMacro: "[0]!Module1.Run",
        min: 1,
        max: 20,
        increment: 2,
        page: 4,
        val: 7,
        textAlign: "Center",
        vertical: "Top",
        lockText: true,
        noThreeD: true,
        noThreeD2: true,
        items: ["Alpha", "Beta"],
      },
    });
    expect(controls?.[1]).toMatchObject({
      kind: "activeX",
      shapeId: 2,
      rId: "rId2",
      name: "Command1",
      activeXControl: {
        classId: "{8BD21D10-EC42-11CE-9E0D-00AA006002F3}",
        persistence: "persistPropertyBag",
        license: "test-license",
        properties: [
          { name: "_ExtentX", value: "1561" },
          { name: "_ExtentY", value: "582" },
        ],
        relationshipId: "rId1",
        binaryPath: "xl/activeX/activeX1.bin",
      },
    });

    const output = unzipSync(
      (await generateWorkbook(parsed, { type: "uint8array" })) as Uint8Array,
    );
    expect(new TextDecoder().decode(output["xl/ctrlProps/ctrlProp1.xml"]!)).toContain(
      `<fmlaRange>Sheet1!$B$1:$B$4</fmlaRange>`,
    );
    expect(new TextDecoder().decode(output["xl/activeX/activeX1.xml"]!)).toContain(
      `ax:classid="{8BD21D10-EC42-11CE-9E0D-00AA006002F3}"`,
    );
    expect(new TextDecoder().decode(output["xl/activeX/activeX1.xml"]!)).toContain('r:id="rId1"');
    expect(new TextDecoder().decode(output["xl/activeX/_rels/activeX1.xml.rels"]!)).toContain(
      `Target="activeX1.bin"`,
    );
    expect(output["xl/activeX/activeX1.bin"]).toEqual(new Uint8Array([1, 2, 3]));
  });

  it("parses and re-emits a whole controls container wrapped by mc:AlternateContent", () => {
    const readContext = {
      resolveRelationship: (rid: string) =>
        rid === "rId1" ? "xl/ctrlProps/ctrlProp1.xml" : undefined,
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const xml =
      XML_DECL +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/>` +
      `<mc:AlternateContent><mc:Choice Requires="x14" ` +
      `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
      `<controls><control shapeId="1" r:id="rId1"><controlPr><anchor><from>` +
      `<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>` +
      `<xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>` +
      `<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control>` +
      `</controls></mc:Choice></mc:AlternateContent></worksheet>`;
    const el = parseXml(xml, { nativeTypeAttributes: true }).elements?.[0];
    if (!el) throw new Error("fixture has no root");
    const parsed = worksheetDesc.parse(el, readContext) as WorksheetOptions;
    expect(parsed.controlsAlternateContent).toBe(true);
    expect(parsed.controls?.[0]).toMatchObject({ kind: "form", shapeId: 1 });
    const output = buildWorksheetXml(parsed, {});
    expect(output).toContain("<mc:AlternateContent xmlns:mc=");
    expect(output).toContain('<controls><control shapeId="1"');
  });

  it("describes form and ActiveX part payloads without exposing binary content", () => {
    const writeContext = {} as WriteContext;
    const form = controlPropertiesDesc.stringify(
      {
        objectType: "Radio",
        checked: "Checked",
        increment: 3,
        items: ["One"],
      },
      writeContext,
    );
    expect(form).toContain('objectType="Radio"');
    expect(form).toContain('inc="3"');
    expect(form).toContain("<item>One</item>");
    const activeX = activeXControlDesc.stringify(
      {
        classId: "{TEST}",
        persistence: "persistStreamInit",
        license: "",
        properties: [{ name: "_Version", value: "393216" }],
      },
      writeContext,
    );
    if (!activeX) throw new Error("ActiveX stringify returned no XML");
    expect(activeX).toBe(
      `<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" ` +
        `ax:classid="{TEST}" ax:persistence="persistStreamInit" ax:license="">` +
        `<ax:ocxPr ax:name="_Version" ax:value="393216"/></ax:ocx>`,
    );
    const activeXElement = parseXml(activeX).elements?.[0];
    if (!activeXElement) throw new Error("ActiveX fixture has no root");
    expect(
      activeXControlDesc.parse(activeXElement, {
        resolveRelationship: () => undefined,
        getPart: () => undefined,
        getRaw: () => undefined,
      }),
    ).toEqual({
      classId: "{TEST}",
      persistence: "persistStreamInit",
      license: "",
      properties: [{ name: "_Version", value: "393216" }],
    });
  });

  it("rejects unsupported form-control children with a structured error", () => {
    const xml =
      XML_DECL +
      `<x14:formControlPr xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ` +
      `objectType="Button"><unknown/></x14:formControlPr>`;
    const el = parseXml(xml).elements![0]!;
    expect(() =>
      controlPropertiesDesc.parse(el, {
        resolveRelationship: () => undefined,
        getPart: () => undefined,
        getRaw: () => undefined,
      }),
    ).toThrow(/xl\/ctrlProps\/formControlPr: unknown: unsupported element/);
  });
});
