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
const themeMediaRelsXml =
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId1" Type="${THEME_RELATIONSHIP_TYPE.replace("/theme", "/image")}" Target="../media/theme-media.png"/></Relationships>`;
const mediaBlipThemeXml =
  '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Media Theme">' +
  '<a:themeElements><a:fmtScheme name="Office">' +
  '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>" +
  "<a:effectStyle><a:effectLst/></a:effectStyle>" +
  "<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:blipFill><a:blip r:embed="rId1"/></a:blipFill></a:bgFillStyleLst>' +
  "</a:fmtScheme></a:themeElements></a:theme>";

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

  it("resolves format-scheme blip fills against nested theme relationships", () => {
    const source = zipSync({
      "word/document.xml": new TextEncoder().encode(documentXml),
      "word/_rels/document.xml.rels": new TextEncoder().encode(documentRelsXml),
      "word/theme/theme1.xml": new TextEncoder().encode(mediaBlipThemeXml),
      "word/theme/_rels/theme1.xml.rels": new TextEncoder().encode(themeMediaRelsXml),
      "word/media/theme-media.png": new Uint8Array([1, 2, 3]),
    });

    const parsed = parseDocumentSync(source);
    const fill = parsed.theme?.formatScheme?.backgroundFillStyles.at(-1);
    expect(fill).toMatchObject({ type: "blip", fileName: "theme-media.png" });

    const output = unzipSync(generateDocumentSync(parsed, { type: "uint8array" }));
    expect(output["word/media/theme-media.png"]).toEqual(new Uint8Array([1, 2, 3]));
    expect(output["word/theme/_rels/theme1.xml.rels"]).toBeDefined();
    const regenerated = parseDocumentSync(zipSync(output));
    expect(regenerated.theme?.formatScheme?.backgroundFillStyles.at(-1)).toEqual(fill);
  });
});
