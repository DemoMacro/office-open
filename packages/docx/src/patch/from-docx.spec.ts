import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

vi.mock("fflate", () => ({
  strFromU8: vi.fn().mockImplementation((data: Uint8Array) => new TextDecoder().decode(data)),
  unzipSync: vi.fn(),
  zip: vi
    .fn()
    .mockImplementation(
      (_data: unknown, _opts: unknown, cb: (err: null, data: Uint8Array) => void) =>
        cb(null, new Uint8Array(0)),
    ),
}));

import { unzipSync } from "@office-open/core";

import { patchDocument } from "./from-docx";

const MOCK_XML = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
    <w:body>
        <w:p w14:paraId="2499FE9F" w14:textId="0A3D130F" w:rsidR="00B51233"
            w:rsidRDefault="007B52ED" w:rsidP="007B52ED">
            <w:pPr>
                <w:pStyle w:val="Title" />
            </w:pPr>
            <w:r>
                <w:t>Hello World</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="6410D9A0" w14:textId="7579AB49" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED" />
        <w:p w14:paraId="57ACF964" w14:textId="315D7A05" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED">
            <w:r>
                <w:t>Hello {{name}},</w:t>
            </w:r>
            <w:r w:rsidR="008126CB">
                <w:t xml:space="preserve"> how are you?</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="38C7DF4A" w14:textId="66CDEC9A" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED" />
        <w:p w14:paraId="04FABE2B" w14:textId="3DACA001" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED">
            <w:r>
                <w:t>{{paragraph_replace}}</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="7AD7975D" w14:textId="77777777" w:rsidR="00EF161F"
            w:rsidRDefault="007B52ED" />
        <w:p w14:paraId="3BD6D75A" w14:textId="19AE3121" w:rsidR="00EF161F"
            w:rsidRDefault="007B52ED">
            <w:r>
                <w:t>{{table}}</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="76023962" w14:textId="4E606AB9" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED" />
        <w:tbl>
            <w:tblPr>
                <w:tblStyle w:val="TableGrid" />
                <w:tblW w:w="0" w:type="auto" />
                <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"
                    w:lastColumn="0" w:noHBand="0" w:noVBand="1" />
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="3003" />
                <w:gridCol w:w="3003" />
                <w:gridCol w:w="3004" />
            </w:tblGrid>
            <w:tr w:rsidR="00EF161F" w14:paraId="1DEC5955" w14:textId="77777777" w:rsidTr="00EF161F">
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="54DA5587" w14:textId="625BAC60" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F">
                        <w:r>
                            <w:t>{{table_heading_1}}</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="57100910" w14:textId="71FD5616" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3004" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="1D388FAB" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
            </w:tr>
            <w:tr w:rsidR="00EF161F" w14:paraId="0F53D2DC" w14:textId="77777777" w:rsidTr="00EF161F">
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="0F2BCCED" w14:textId="3C3B6706" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F">
                        <w:r>
                            <w:t>Item: {{item_1}}</w:t>
                        </w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="1E6158AC" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3004" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="17937748" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
            </w:tr>
            <w:tr w:rsidR="00EF161F" w14:paraId="781DAC1A" w14:textId="77777777" w:rsidTr="00EF161F">
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="1DCD0343" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3003" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="5D02E3CD" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
                <w:tc>
                    <w:tcPr>
                        <w:tcW w:w="3004" w:type="dxa" />
                    </w:tcPr>
                    <w:p w14:paraId="52EA0DBB" w14:textId="77777777" w:rsidR="00EF161F"
                        w:rsidRDefault="00EF161F" />
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:p w14:paraId="47CD1FBC" w14:textId="23474CBC" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED" />
        <w:p w14:paraId="0ACCEE90" w14:textId="67907499" w:rsidR="00EF161F"
            w:rsidRDefault="0077578F">
            <w:r>
                <w:t>{{image_test}}</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="23FA9862" w14:textId="77777777" w:rsidR="0077578F"
            w:rsidRDefault="0077578F" />
        <w:p w14:paraId="01578F2F" w14:textId="3BDC6C85" w:rsidR="007B52ED"
            w:rsidRDefault="007B52ED">
            <w:r>
                <w:t>Thank you</w:t>
            </w:r>
        </w:p>
        <w:sectPr w:rsidR="007B52ED" w:rsidSect="0072043F">
            <w:headerReference w:type="default" r:id="rId6" />
            <w:footerReference w:type="default" r:id="rId7" />
            <w:pgSz w:w="11900" w:h="16840" />
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"
                w:footer="708" w:gutter="0" />
            <w:cols w:space="708" />
            <w:docGrid w:linePitch="360" />
        </w:sectPr>
    </w:body>
</w:document>
`;

/**
 * Creates a mock Unzipped object that mimics fflate's unzipSync output.
 * Uses Object.fromEntries to avoid @typescript-eslint/naming-convention
 * violations on OOXML file paths like "word/document.xml".
 */
const createMockUnzipped = (
  entries: [string, string | Uint8Array][],
): Record<string, Uint8Array<ArrayBuffer>> =>
  Object.fromEntries(
    entries.map(([key, value]): [string, Uint8Array<ArrayBuffer>] => [
      key,
      typeof value === "string"
        ? new TextEncoder().encode(value)
        : (value as Uint8Array<ArrayBuffer>),
    ]),
  );

describe("from-docx", () => {
  describe("patchDocument", () => {
    describe("document.xml and [Content_Types].xml", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should patch the document", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {
            image_test: {
              children: [
                {
                  image: {
                    data: Buffer.from(""),
                    transformation: { height: 100, width: 100 },
                    type: "png",
                  },
                },
              ],
              type: "paragraph",
            },
            item_1: {
              children: [
                "#657",
                {
                  hyperlink: {
                    link: "https://www.bbc.co.uk/news",
                    children: ["BBC News Link"],
                  },
                },
              ],
              type: "paragraph",
            },
            name: {
              children: ["Sir. ", "John Doe", "(The Conqueror)"],
              type: "paragraph",
            },
            paragraph_replace: {
              children: [
                {
                  paragraph: {
                    children: [
                      "This is a ",
                      {
                        hyperlink: {
                          link: "https://www.google.co.uk",
                          children: ["Google Link"],
                        },
                      },
                      {
                        image: {
                          data: Buffer.from(""),
                          transformation: { height: 100, width: 100 },
                          type: "png",
                        },
                      },
                    ],
                  },
                },
              ],
              type: "document",
            },
          },
        });
        expect(output).to.not.be.undefined;
      });

      it("should patch the document", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {},
        });
        expect(output).to.not.be.undefined;
      });

      it("should skip UTF-16 types", async () => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", Buffer.from([0xff, 0xfe])],
          ]),
        );
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {},
        });
        expect(output).to.not.be.undefined;
      });

      it("should patch the document with custom delimiters", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {
            image_test: {
              children: [
                {
                  image: {
                    data: Buffer.from(""),
                    transformation: { height: 100, width: 100 },
                    type: "png",
                  },
                },
              ],
              type: "paragraph",
            },
            item_1: {
              children: [
                "#657",
                {
                  hyperlink: {
                    link: "https://www.bbc.co.uk/news",
                    children: ["BBC News Link"],
                  },
                },
              ],
              type: "paragraph",
            },
            name: {
              children: ["Sir. ", "John Doe", "(The Conqueror)"],
              type: "paragraph",
            },
            paragraph_replace: {
              children: [
                {
                  paragraph: {
                    children: [
                      "This is a ",
                      {
                        hyperlink: {
                          link: "https://www.google.co.uk",
                          children: ["Google Link"],
                        },
                      },
                      {
                        image: {
                          data: Buffer.from(""),
                          transformation: { height: 100, width: 100 },
                          type: "png",
                        },
                      },
                    ],
                  },
                },
              ],
              type: "document",
            },
          },
          placeholderDelimiters: { end: "}}", start: "{{" },
        });
        expect(output).to.not.be.undefined;
      });

      it("should patch the document with no patches", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {},
        });
        expect(output).to.not.be.undefined;
      });

      it("throws error with empty delimiters", async () => {
        await expect(() =>
          patchDocument({
            data: Buffer.from(""),
            outputType: "uint8array",
            placeholders: {},
            placeholderDelimiters: { end: "", start: "" },
          }),
        ).rejects.toThrow();
      });

      it("throws error with whitespace-only delimiters", async () => {
        await expect(() =>
          patchDocument({
            data: Buffer.from(""),
            outputType: "uint8array",
            placeholders: {},
            placeholderDelimiters: { end: " ", start: " " },
          }),
        ).rejects.toThrowError();
      });
    });

    describe("document.xml and [Content_Types].xml with relationships", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            [
              "word/_rels/document.xml.rels",
              `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
            ],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should use the relationships file rather than create one", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {
            image_test: {
              children: [
                {
                  image: {
                    data: Buffer.from(""),
                    transformation: { height: 100, width: 100 },
                    type: "png",
                  },
                },
                {
                  hyperlink: {
                    link: "https://www.google.co.uk",
                    children: ["Google Link"],
                  },
                },
              ],
              type: "paragraph",
            },
          },
        });
        expect(output).to.not.be.undefined;
      });
    });

    describe("document.xml and [Content_Types].xml without relationships file", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should create a relationships file for hyperlink only patches", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {
            image_test: {
              children: [
                {
                  hyperlink: {
                    link: "https://www.google.co.uk",
                    children: ["Google Link"],
                  },
                },
              ],
              type: "paragraph",
            },
          },
        });
        expect(output).to.not.be.undefined;
      });
    });

    describe("document.xml without attributes on w:document", () => {
      const MOCK_XML_NO_ATTRS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document>
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello {{name}}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>`;

      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML_NO_ATTRS],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should patch a document whose w:document element has no attributes", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          placeholders: {
            name: {
              children: ["World"],
              type: "paragraph",
            },
          },
        });
        expect(output).to.not.be.undefined;
      });
    });

    describe("document.xml", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([["word/document.xml", MOCK_XML]]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should throw an error if the content types is not found", () =>
        expect(
          patchDocument({
            data: Buffer.from(""),
            outputType: "uint8array",
            placeholders: {
              image_test: {
                children: [
                  {
                    image: {
                      data: Buffer.from(""),
                      transformation: { height: 100, width: 100 },
                      type: "png",
                    },
                  },
                ],
                type: "paragraph",
              },
            },
          }),
        ).rejects.toThrowError());
    });

    describe("Images", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["word/document.bmp", ""],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should throw an error if the content types is not found", () =>
        expect(
          patchDocument({
            data: Buffer.from(""),
            outputType: "uint8array",
            placeholders: {
              image_test: {
                children: [
                  {
                    image: {
                      data: Buffer.from(""),
                      transformation: { height: 100, width: 100 },
                      type: "png",
                    },
                  },
                ],
                type: "paragraph",
              },
            },
          }),
        ).rejects.toThrowError());
    });

    describe("output types", () => {
      beforeEach(() => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      it("should export to nodebuffer", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "nodebuffer",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(output instanceof Uint8Array).toBe(true);
      });

      it("should export to blob", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "blob",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(output).toBeInstanceOf(Blob);
      });

      it("should export to arraybuffer", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "arraybuffer",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(output).toBeInstanceOf(ArrayBuffer);
      });

      it("should export to base64", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "base64",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(typeof output).toBe("string");
      });

      it("should export to string", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "string",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(typeof output).toBe("string");
      });

      it("should export to binarystring", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "binarystring",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(typeof output).toBe("string");
      });

      it("should export to array", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "array",
          placeholders: {
            name: { children: ["World"], type: "paragraph" },
          },
        });
        expect(Array.isArray(output)).toBe(true);
      });
    });

    describe("append", () => {
      // Real unzip (bypass the fflate mock) to inspect the actually-zipped output.
      // zipAndConvert uses nativeZipAsync under Node, so spying on the mocked
      // fflate `zip` captures nothing; we re-inflate the returned buffer instead.
      let realUnzip: typeof import("fflate").unzipSync;

      beforeEach(async () => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`],
          ]),
        );
        realUnzip = (await vi.importActual<typeof import("fflate")>("fflate")).unzipSync;
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      const decodeDoc = async (output: unknown): Promise<string> => {
        const unzipped = realUnzip(output as Uint8Array);
        const entry = unzipped["word/document.xml"];
        expect(entry, "word/document.xml should be zipped").toBeDefined();
        return new TextDecoder().decode(entry);
      };

      it("should splice appended children before the trailing sectPr", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          append: [
            { paragraph: { children: ["APPENDED_MARKER"] } },
            { paragraph: { children: ["SECOND_APPENDED"] } },
          ],
        });

        const xml = await decodeDoc(output);
        expect(xml).toContain("APPENDED_MARKER");
        expect(xml).toContain("SECOND_APPENDED");
        // Appended content must precede the body's final section properties.
        const markerPos = xml.indexOf("APPENDED_MARKER");
        const sectPrPos = xml.indexOf("<w:sectPr");
        expect(markerPos).toBeGreaterThan(-1);
        expect(sectPrPos).toBeGreaterThan(-1);
        expect(markerPos).toBeLessThan(sectPrPos);
      });

      it("should append when no placeholders are given", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          append: [{ paragraph: { children: ["ONLY_APPEND"] } }],
        });

        expect(await decodeDoc(output)).toContain("ONLY_APPEND");
      });

      it("should serialize appended tables via the compile-path stringifier", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          append: [
            {
              table: {
                columnWidths: [2000, 2000],
                rows: [
                  {
                    cells: [
                      { children: [{ paragraph: "APPENDED_CELL_A" }] },
                      { children: [{ paragraph: "APPENDED_CELL_B" }] },
                    ],
                  },
                ],
              },
            },
          ],
        });

        const xml = await decodeDoc(output);
        expect(xml).toContain("<w:tbl>");
        expect(xml).toContain("APPENDED_CELL_A");
        expect(xml).toContain("APPENDED_CELL_B");
      });
    });

    describe("comments", () => {
      let realUnzip: typeof import("fflate").unzipSync;
      const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`;

      beforeEach(async () => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            ["[Content_Types].xml", CT],
          ]),
        );
        realUnzip = (await vi.importActual<typeof import("fflate")>("fflate")).unzipSync;
      });

      afterEach(() => {
        vi.restoreAllMocks();
      });

      const decodeEntry = async (output: unknown, path: string): Promise<string> => {
        const unzipped = realUnzip(output as Uint8Array);
        const entry = unzipped[path];
        expect(entry, `${path} should be zipped`).toBeDefined();
        return new TextDecoder().decode(entry);
      };

      it("wraps the Nth paragraph + creates comments.xml with rels/content types", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          comments: {
            paragraphs: {
              0: [{ author: "Alice", children: [{ children: ["First note"] }] }],
            },
          },
        });

        const xml = await decodeEntry(output, "word/document.xml");
        expect(xml).toContain('<w:commentRangeStart w:id="0"');
        expect(xml).toContain('<w:commentRangeEnd w:id="0"');
        expect(xml).toContain('<w:commentReference w:id="0"');

        const commentsXml = await decodeEntry(output, "word/comments.xml");
        expect(commentsXml).toContain("Alice");
        expect(commentsXml).toContain("First note");

        expect(await decodeEntry(output, "word/_rels/document.xml.rels")).toContain("comments.xml");
        expect(await decodeEntry(output, "[Content_Types].xml")).toContain("/word/comments.xml");
      });

      it("wraps the placeholder run before substitution", async () => {
        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          comments: {
            placeholders: {
              name: [{ author: "Bob", children: [{ children: ["On the name"] }] }],
            },
          },
        });

        const xml = await decodeEntry(output, "word/document.xml");
        expect(xml).toContain('<w:commentRangeStart w:id="0"');
        expect(xml).toContain('<w:commentReference w:id="0"');
      });

      it("continues ids + merges entries with an existing comments part", async () => {
        vi.mocked(unzipSync).mockImplementation(() =>
          createMockUnzipped([
            ["word/document.xml", MOCK_XML],
            [
              "word/comments.xml",
              `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="5" w:author="Old" w:date="2024-01-01T00:00:00Z"><w:p/></w:comment></w:comments>`,
            ],
            ["[Content_Types].xml", CT],
          ]),
        );

        const output = await patchDocument({
          data: Buffer.from(""),
          outputType: "uint8array",
          comments: {
            paragraphs: {
              0: [{ author: "Alice", children: [{ children: ["New note"] }] }],
            },
          },
        });

        // New comment id continues at 6 (after existing max 5).
        const xml = await decodeEntry(output, "word/document.xml");
        expect(xml).toContain('<w:commentRangeStart w:id="6"');

        const commentsXml = await decodeEntry(output, "word/comments.xml");
        expect(commentsXml).toContain("Old"); // existing preserved
        expect(commentsXml).toContain("Alice"); // new merged

        // No duplicate content-type override.
        const types = await decodeEntry(output, "[Content_Types].xml");
        expect((types.match(/\/word\/comments\.xml/g) ?? []).length).toBe(1);
      });

      it("throws when commenting a non-existent paragraph index", async () => {
        await expect(
          patchDocument({
            data: Buffer.from(""),
            outputType: "uint8array",
            comments: {
              paragraphs: {
                999: [{ author: "X", children: [{ children: ["y"] }] }],
              },
            },
          }),
        ).rejects.toThrow(/index 999/);
      });
    });
  });
});
