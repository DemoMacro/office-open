import { unzipSync, zipSync } from "@office-open/core";
import { buildThemeXml, type ThemeOptions } from "@office-open/core/theme";
import { describe, expect, it } from "vite-plus/test";

import { generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";

const THEME_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const documentXml =
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>';
const documentRelsXml =
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId1" Type="${THEME_RELATIONSHIP_TYPE}" Target="theme/theme1.xml"/></Relationships>`;

const themeOptions: ThemeOptions = {
  name: "Custom Theme",
  colorScheme: {
    name: "Custom Colors",
    accent1: "123456",
    accent2: "234567",
    accent3: "345678",
    accent4: "456789",
    accent5: "56789A",
    accent6: "6789AB",
    dark2: "111111",
    light2: "222222",
    hyperlink: "789ABC",
    followedHyperlink: "89ABCD",
  },
  fontScheme: {
    name: "Custom Fonts",
    majorFont: {
      latin: { typeface: "Custom Serif" },
      eastAsian: { typeface: "Major EA" },
      complexScript: { typeface: "Major CS" },
    },
    minorFont: {
      latin: { typeface: "Custom Sans" },
      eastAsian: { typeface: "Minor EA" },
      complexScript: { typeface: "Minor CS" },
    },
  },
};

const themePackage = () => {
  const themeXml = '<?xml version="1.0"?>' + buildThemeXml(themeOptions);
  return zipSync({
    "word/document.xml": new TextEncoder().encode(documentXml),
    "word/_rels/document.xml.rels": new TextEncoder().encode(documentRelsXml),
    "word/theme/theme1.xml": new TextEncoder().encode(themeXml),
  });
};

describe("document theme part", () => {
  it("parses a themed package into the canonical theme model", () => {
    const parsed = parseDocumentSync(themePackage());
    expect(parsed.theme?.name).toBe("Custom Theme");
    expect(parsed.theme?.colorScheme?.accent1).toBe("123456");
    expect(parsed.theme?.fontScheme?.majorFont?.latin?.typeface).toBe("Custom Serif");
    expect(parsed.theme?.fontScheme?.minorFont?.latin?.typeface).toBe("Custom Sans");
  });

  it("regenerates the theme part instead of passing it through", () => {
    const zip = unzipSync(
      generateDocumentSync(parseDocumentSync(themePackage()), { type: "uint8array" }),
    );
    const theme = new TextDecoder().decode(zip["word/theme/theme1.xml"]!);
    expect(theme).toContain('name="Custom Theme"');
    expect(theme).toContain('<a:accent1><a:srgbClr val="123456"/></a:accent1>');
    expect(theme).toContain('typeface="Custom Sans"');
    expect(Object.keys(zip)).not.toContain("word/_rels/theme1.xml.rels");
  });

  it("round-trips the canonical model without raw theme storage", () => {
    const options = parseDocumentSync(themePackage());
    expect(options.rawParts?.some((part) => part.path.startsWith("word/theme/")) ?? false).toBe(
      false,
    );
    const regenerated = parseDocumentSync(generateDocumentSync(options, { type: "uint8array" }));
    expect(regenerated.theme).toEqual(options.theme);
  });

  it("emits the default theme when no canonical theme is present", () => {
    const zip = unzipSync(
      generateDocumentSync({ sections: [{ children: [] }] }, { type: "uint8array" }),
    );
    const theme = new TextDecoder().decode(zip["word/theme/theme1.xml"]!);
    expect(theme).toContain('name="Office Theme"');
    expect(Object.keys(zip)).not.toContain("word/_rels/theme1.xml.rels");
  });

  it("registers theme media relationships only in the theme part", () => {
    const zip = unzipSync(
      generateDocumentSync(
        {
          sections: [],
          theme: {
            name: "Media Theme",
            objectDefaults: {
              shapeDefault: {
                shapeProperties: {
                  fill: {
                    type: "blip",
                    data: new Uint8Array([1, 2, 3]),
                    imageType: "png",
                    fileName: "theme-media.png",
                  },
                },
              },
            },
          },
        },
        { type: "uint8array" },
      ),
    );
    const rels = new TextDecoder().decode(zip["word/theme/_rels/theme1.xml.rels"]!);
    expect(rels).toContain('Target="../media/theme-media.png"');
    expect(zip["word/media/theme-media.png"]).toEqual(new Uint8Array([1, 2, 3]));
  });
});
