import { describe, expect, it } from "vite-plus/test";

import { File } from "./file";

describe("File", () => {
  describe("#comments", () => {
    it("should create comments", () => {
      const doc = new File({
        comments: {
          children: [],
        },
        sections: [],
      });

      expect(doc.comments).to.not.be.undefined;
    });
  });

  describe("#numbering", () => {
    it("should create", () => {
      const doc = new File({
        numbering: { config: [] },
        sections: [],
      });

      expect(doc.numbering).to.not.be.undefined;
    });
  });

  describe("#getters", () => {
    it("should have defined getters", () => {
      const doc = new File({
        sections: [],
      });

      expect(doc.coreProperties).to.not.be.undefined;
      expect(doc.media).to.not.be.undefined;
      expect(doc.fileRelationships).to.not.be.undefined;
      expect(doc.headers).to.not.be.undefined;
      expect(doc.footers).to.not.be.undefined;
      expect(doc.contentTypes).to.not.be.undefined;
      expect(doc.customProperties).to.not.be.undefined;
      expect(doc.appProperties).to.not.be.undefined;
      expect(doc.footNotes).to.not.be.undefined;
      expect(doc.settings).to.not.be.undefined;
      expect(doc.comments).to.not.be.undefined;
    });
  });

  describe("#externalStyles", () => {
    it("should work with external styles", () => {
      const doc = new File({
        externalStyles: `
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
                    </w:styles>`,
        sections: [],
      });

      expect(doc.styles).to.not.be.undefined;
    });

    it("should merge external styles with default styles when both are provided", () => {
      const doc = new File({
        externalStyles: `
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <w:styles xmlns:mc="first" xmlns:r="second">
                        <w:style w:type="paragraph" w:styleId="Heading1">
                            <w:name w:val="heading 1"/>
                        </w:style>
                    </w:styles>`,
        sections: [],
        styles: {
          default: {
            heading1: {
              run: {
                size: 28,
              },
            },
          },
        },
      });

      expect(doc.styles).to.not.be.undefined;
    });
  });

  describe("#features", () => {
    it("should work with updateFields", () => {
      const doc = new File({
        features: {
          updateFields: true,
        },
        sections: [],
      });

      expect(doc.styles).to.not.be.undefined;
    });

    it("should work with trackRevisions", () => {
      const doc = new File({
        features: {
          trackRevisions: true,
        },
        sections: [],
      });

      expect(doc.styles).to.not.be.undefined;
    });
  });

  describe("#hyphenation", () => {
    it("should work with autoHyphenation", () => {
      const doc = new File({
        hyphenation: {
          autoHyphenation: true,
        },
        sections: [],
      });

      expect(doc.styles).to.not.be.undefined;
    });
  });
});
