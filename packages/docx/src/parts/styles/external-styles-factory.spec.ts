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

      expect(importedStyle.initialAttributes).to.deep.equal({
        "xmlns:mc": "first",
        "xmlns:r": "second",
      });
    });

    it("should parse other child elements of w:styles", () => {
      const importedStyle = new ExternalStylesFactory().newInstance(externalStyles);
      const importedStyles = importedStyle.importedStyles!;

      // docDefaults
      expect(importedStyles[0]._raw).to.include("w:docDefaults");
      expect(importedStyles[0]._raw).to.include("w:rFonts");
      expect(importedStyles[0]._raw).to.include("w:spacing");

      // latentStyles
      expect(importedStyles[1]._raw).to.include("w:latentStyles");
      expect(importedStyles[1]._raw).to.include("w:defLockedState");
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

      // Normal style
      expect(importedStyles[2]._raw).to.include('w:styleId="Normal"');
      expect(importedStyles[2]._raw).to.include('<w:name w:val="Normal"/>');
      expect(importedStyles[2]._raw).to.include("<w:qFormat/>");

      // Heading1 style
      expect(importedStyles[3]._raw).to.include('w:styleId="Heading1"');
      expect(importedStyles[3]._raw).to.include('<w:name w:val="heading 1"/>');
      expect(importedStyles[3]._raw).to.include('<w:basedOn w:val="Normal"/>');
      expect(importedStyles[3]._raw).to.include("<w:keepNext/>");
      expect(importedStyles[3]._raw).to.include("<w:keepLines/>");
    });
  });
});
