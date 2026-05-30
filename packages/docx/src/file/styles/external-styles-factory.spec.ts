import { beforeEach, describe, expect, it } from "vite-plus/test";

import { ExternalStylesFactory } from "./external-styles-factory";

describe("External styles factory", () => {
  let externalStyles: string;

  beforeEach(() => {
    externalStyles = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:mc="first" xmlns:r="second">
            <w:docDefaults>
            <w:rPrDefault>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:eastAsiaTheme="minorHAnsi" w:hAnsi="Arial" w:cstheme="minorHAnsi"/>
                    <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
                </w:rPr>
            </w:rPrDefault>
            <w:pPrDefault>
                <w:pPr>
                    <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
                </w:pPr>
            </w:pPrDefault>
            </w:docDefaults>

            <w:latentStyles w:defLockedState="1" w:defUIPriority="99">
            </w:latentStyles>

            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                <w:name w:val="Normal"/>
                <w:qFormat/>
            </w:style>

            <w:style w:type="paragraph" w:styleId="Heading1">
                <w:name w:val="heading 1"/>
                <w:basedOn w:val="Normal"/>
                <w:pPr>
                    <w:keepNext/>
                    <w:keepLines/>

                    <w:pBdr>
                        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
                  </w:pBdr>
                </w:pPr>
            </w:style>
        </w:styles>`;
  });

  describe("#parse", () => {
    it("should parse w:styles attributes", () => {
      const importedStyle = new ExternalStylesFactory().newInstance(externalStyles);

      expect(
        (importedStyle.initialStyles as { _attr: Record<string, string> })._attr,
      ).to.deep.equal({
        "xmlns:mc": "first",
        "xmlns:r": "second",
      });
    });

    it("should parse other child elements of w:styles", () => {
      const importedStyle = new ExternalStylesFactory().newInstance(externalStyles);
      const importedStyles = importedStyle.importedStyles!;

      expect(JSON.parse(JSON.stringify(importedStyles[0]))).to.deep.equal({
        root: [
          {
            root: [
              {
                root: [
                  {
                    root: [
                      {
                        _attr: {
                          "w:ascii": "Arial",
                          "w:cstheme": "minorHAnsi",
                          "w:eastAsiaTheme": "minorHAnsi",
                          "w:hAnsi": "Arial",
                        },
                      },
                    ],
                    rootKey: "w:rFonts",
                  },
                  {
                    root: [
                      {
                        _attr: {
                          "w:bidi": "ar-SA",
                          "w:eastAsia": "en-US",
                          "w:val": "en-US",
                        },
                      },
                    ],
                    rootKey: "w:lang",
                  },
                ],
                rootKey: "w:rPr",
              },
            ],
            rootKey: "w:rPrDefault",
          },
          {
            root: [
              {
                root: [
                  {
                    root: [
                      {
                        _attr: {
                          "w:after": "160",
                          "w:line": "259",
                          "w:lineRule": "auto",
                        },
                      },
                    ],
                    rootKey: "w:spacing",
                  },
                ],
                rootKey: "w:pPr",
              },
            ],
            rootKey: "w:pPrDefault",
          },
        ],
        rootKey: "w:docDefaults",
      });
      expect(JSON.parse(JSON.stringify(importedStyles[1]))).to.deep.equal({
        root: [
          {
            _attr: {
              "w:defLockedState": "1",
              "w:defUIPriority": "99",
            },
          },
        ],
        rootKey: "w:latentStyles",
      });
    });

    it("should return empty styles when style element isn't found", () => {
      const result = new ExternalStylesFactory().newInstance(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><foo/>`,
      );
      expect(result.importedStyles).to.deep.equal([]);

      const emptyResult = new ExternalStylesFactory().newInstance(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      );
      expect(emptyResult.importedStyles).to.deep.equal([]);
    });

    it("should handle w:styles with no child elements", () => {
      const emptyStyles = `
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:styles xmlns:mc="first" xmlns:r="second">
            </w:styles>`;
      const importedStyle = new ExternalStylesFactory().newInstance(emptyStyles);

      expect(importedStyle.importedStyles).to.deep.equal([]);
    });

    it("should parse styles elements", () => {
      const importedStyle = new ExternalStylesFactory().newInstance(externalStyles);
      const importedStyles = importedStyle.importedStyles!;

      expect(importedStyles.length).to.equal(4);
      expect(JSON.parse(JSON.stringify(importedStyles[2]))).to.deep.equal({
        root: [
          {
            _attr: {
              "w:default": "1",
              "w:styleId": "Normal",
              "w:type": "paragraph",
            },
          },
          {
            root: [
              {
                _attr: {
                  "w:val": "Normal",
                },
              },
            ],
            rootKey: "w:name",
          },
          {
            root: [],
            rootKey: "w:qFormat",
          },
        ],
        rootKey: "w:style",
      });

      expect(JSON.parse(JSON.stringify(importedStyles[3]))).to.deep.equal({
        root: [
          {
            _attr: {
              "w:styleId": "Heading1",
              "w:type": "paragraph",
            },
          },
          {
            root: [
              {
                _attr: {
                  "w:val": "heading 1",
                },
              },
            ],
            rootKey: "w:name",
          },
          {
            root: [
              {
                _attr: {
                  "w:val": "Normal",
                },
              },
            ],
            rootKey: "w:basedOn",
          },
          {
            root: [
              {
                root: [],
                rootKey: "w:keepNext",
              },
              {
                root: [],
                rootKey: "w:keepLines",
              },
              {
                root: [
                  {
                    root: [
                      {
                        _attr: {
                          "w:color": "auto",
                          "w:space": "1",
                          "w:sz": "4",
                          "w:val": "single",
                        },
                      },
                    ],
                    rootKey: "w:bottom",
                  },
                ],
                rootKey: "w:pBdr",
              },
            ],
            rootKey: "w:pPr",
          },
        ],
        rootKey: "w:style",
      });
    });
  });
});
