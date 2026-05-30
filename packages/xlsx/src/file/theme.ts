/**
 * Default theme for XLSX files — matches Microsoft Office's output structure.
 * Produces xl/theme/theme1.xml that Excel accepts without repair warnings.
 *
 * XSD reference: CT_OfficeStyleSheet in dml-main.xsd
 * - themeElements (required): clrScheme, fontScheme, fmtScheme
 * - objectDefaults (optional)
 * - extraClrSchemeLst (optional)
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

export class DefaultTheme extends BaseXmlComponent {
  public constructor() {
    super("a:theme");
  }

  public override prepForXml(_context: Context): IXmlableObject {
    return {
      "a:theme": [
        {
          _attr: {
            "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            name: "Office Theme",
          },
        },
        // themeElements (CT_BaseStyles) — all 3 children required
        {
          "a:themeElements": [
            // ── Color scheme (CT_ColorScheme) — 12 colors required ──
            {
              "a:clrScheme": [
                { _attr: { name: "Office" } },
                { "a:dk1": [{ "a:sysClr": { _attr: { val: "windowText", lastClr: "000000" } } }] },
                { "a:lt1": [{ "a:sysClr": { _attr: { val: "window", lastClr: "FFFFFF" } } }] },
                { "a:dk2": [{ "a:srgbClr": { _attr: { val: "44546A" } } }] },
                { "a:lt2": [{ "a:srgbClr": { _attr: { val: "E7E6E6" } } }] },
                { "a:accent1": [{ "a:srgbClr": { _attr: { val: "5B9BD5" } } }] },
                { "a:accent2": [{ "a:srgbClr": { _attr: { val: "ED7D31" } } }] },
                { "a:accent3": [{ "a:srgbClr": { _attr: { val: "A5A5A5" } } }] },
                { "a:accent4": [{ "a:srgbClr": { _attr: { val: "FFC000" } } }] },
                { "a:accent5": [{ "a:srgbClr": { _attr: { val: "4472C4" } } }] },
                { "a:accent6": [{ "a:srgbClr": { _attr: { val: "70AD47" } } }] },
                { "a:hlink": [{ "a:srgbClr": { _attr: { val: "0563C1" } } }] },
                { "a:folHlink": [{ "a:srgbClr": { _attr: { val: "954F72" } } }] },
              ],
            },
            // ── Font scheme (CT_FontScheme) — majorFont + minorFont required ──
            // CT_FontCollection requires: latin, ea, cs (all minOccurs=1)
            {
              "a:fontScheme": [
                { _attr: { name: "Office" } },
                {
                  "a:majorFont": [
                    {
                      "a:latin": {
                        _attr: { typeface: "Calibri Light", panose: "020F0302020204030204" },
                      },
                    },
                    { "a:ea": { _attr: { typeface: "" } } },
                    { "a:cs": { _attr: { typeface: "" } } },
                  ],
                },
                {
                  "a:minorFont": [
                    {
                      "a:latin": { _attr: { typeface: "Calibri", panose: "020F0502020204030204" } },
                    },
                    { "a:ea": { _attr: { typeface: "" } } },
                    { "a:cs": { _attr: { typeface: "" } } },
                  ],
                },
              ],
            },
            // ── Format scheme (CT_StyleMatrix) — 4 style lists, each with 3+ entries ──
            {
              "a:fmtScheme": [
                { _attr: { name: "Office" } },
                // fillStyleLst — 3 fills
                {
                  "a:fillStyleLst": [
                    { "a:solidFill": [{ "a:schemeClr": [{ _attr: { val: "phClr" } }] }] },
                    gradFill2(
                      { pos: "0", lumMod: "110000", satMod: "105000", tint: "67000" },
                      { pos: "50000", lumMod: "105000", satMod: "103000", tint: "73000" },
                      { pos: "100000", lumMod: "105000", satMod: "109000", tint: "81000" },
                    ),
                    gradFill2(
                      { pos: "0", satMod: "103000", lumMod: "102000", tint: "94000" },
                      { pos: "50000", satMod: "110000", lumMod: "100000", shade: "100000" },
                      { pos: "100000", lumMod: "99000", satMod: "120000", shade: "78000" },
                    ),
                  ],
                },
                // lnStyleLst — 3 lines
                {
                  "a:lnStyleLst": [lineProps("6350"), lineProps("12700"), lineProps("19050")],
                },
                // effectStyleLst — 3 effects
                {
                  "a:effectStyleLst": [
                    { "a:effectStyle": [{ "a:effectLst": [] }] },
                    { "a:effectStyle": [{ "a:effectLst": [] }] },
                    {
                      "a:effectStyle": [
                        {
                          "a:effectLst": [
                            {
                              "a:outerShdw": [
                                {
                                  _attr: {
                                    blurRad: "57150",
                                    dist: "19050",
                                    dir: "5400000",
                                    algn: "ctr",
                                    rotWithShape: "0",
                                  },
                                },
                                {
                                  "a:srgbClr": [
                                    { _attr: { val: "000000" } },
                                    { "a:alpha": { _attr: { val: "63000" } } },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
                // bgFillStyleLst — 3 fills
                {
                  "a:bgFillStyleLst": [
                    { "a:solidFill": [{ "a:schemeClr": [{ _attr: { val: "phClr" } }] }] },
                    {
                      "a:solidFill": [
                        {
                          "a:schemeClr": [
                            { _attr: { val: "phClr" } },
                            { "a:tint": { _attr: { val: "95000" } } },
                            { "a:satMod": { _attr: { val: "170000" } } },
                          ],
                        },
                      ],
                    },
                    gradFill2(
                      {
                        pos: "0",
                        tint: "93000",
                        satMod: "150000",
                        shade: "98000",
                        lumMod: "102000",
                      },
                      {
                        pos: "50000",
                        tint: "98000",
                        satMod: "130000",
                        shade: "90000",
                        lumMod: "103000",
                      },
                      { pos: "100000", shade: "63000", satMod: "120000" },
                    ),
                  ],
                },
              ],
            },
          ],
        },
        { "a:objectDefaults": [] },
        { "a:extraClrSchemeLst": [] },
      ],
    };
  }
}

// ── Helper functions (file-scoped) ──

interface GsEntry {
  readonly pos: string;
  readonly lumMod?: string;
  readonly satMod?: string;
  readonly tint?: string;
  readonly shade?: string;
}

/** Build a gradFill with 3 gradient stops */
function gradFill2(a: GsEntry, b: GsEntry, c: GsEntry): IXmlableObject {
  return {
    "a:gradFill": [
      { _attr: { rotWithShape: "1" } },
      {
        "a:gsLst": [gsEntry(a), gsEntry(b), gsEntry(c)],
      },
      { "a:lin": { _attr: { ang: "5400000", scaled: "0" } } },
    ],
  };
}

/** Build a single gradient stop */
function gsEntry(e: GsEntry): IXmlableObject {
  const children: IXmlableObject[] = [{ _attr: { val: "phClr" } }];
  if (e.lumMod) children.push({ "a:lumMod": { _attr: { val: e.lumMod } } });
  if (e.satMod) children.push({ "a:satMod": { _attr: { val: e.satMod } } });
  if (e.tint) children.push({ "a:tint": { _attr: { val: e.tint } } });
  if (e.shade) children.push({ "a:shade": { _attr: { val: e.shade } } });
  return { "a:gs": [{ _attr: { pos: e.pos } }, { "a:schemeClr": children }] };
}

/** Build a line properties entry */
function lineProps(w: string): IXmlableObject {
  return {
    "a:ln": [
      { _attr: { w, cap: "flat", cmpd: "sng", algn: "ctr" } },
      { "a:solidFill": [{ "a:schemeClr": [{ _attr: { val: "phClr" } }] }] },
      { "a:prstDash": { _attr: { val: "solid" } } },
      { "a:miter": { _attr: { lim: "800000" } } },
    ],
  };
}
