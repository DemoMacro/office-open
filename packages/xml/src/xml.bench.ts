import { parse as txmlParse, stringify as txmlStringify } from "txml";
import { bench, describe } from "vite-plus/test";
import xmlPkg from "xml";
import { xml2js as xml2jsOriginal, js2xml as js2xmlOriginal } from "xml-js";

import { parse } from "../src/parse";
import { stringify } from "../src/stringify";

const XML_STRING = '<w:p w:val="1"><w:r><w:t>Bold</w:t></w:r><w:r><w:t>Normal</w:t></w:r></w:p>';
const COMPLEX_XML =
  '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title Text</w:t></w:r></w:p><w:p><w:r><w:t>Body paragraph with special chars: &amp; &lt; &gt;</w:t></w:r></w:p></w:body></w:document>';

// Bun.XML (bun ≥1.4, native) — only present under the bun runtime; the whole
// block self-skips under node so the bench stays runnable in both.
interface BunXmlApi {
  parse(xml: string): unknown;
  stringify(value: unknown): string;
}
const bunXml = (globalThis as { Bun?: { XML?: BunXmlApi } }).Bun?.XML;
// Bun.XML consumes its own nested-object format ("@attr" prefixed keys,
// repeated children as arrays) — prebuilt like the xml (npm) inputs.
const BUN_SIMPLE = {
  "w:p": {
    "@w:val": "1",
    "w:r": [{ "w:t": "Bold" }, { "w:t": "Normal" }],
  },
};
const BUN_COMPLEX = {
  "w:document": {
    "@xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w:body": {
      "w:p": [
        { "w:pPr": { "w:pStyle": { "@w:val": "Title" } }, "w:r": { "w:t": "Title Text" } },
        { "w:r": { "w:t": "Body paragraph with special chars: & < >" } },
      ],
    },
  },
};

// Parameters type of the xml (npm) serializer, so object literals are checked.
type XmlPkgInput = Parameters<typeof xmlPkg>[0];

// xml (npm) consumes its own object format ({ tag: [...] }), not an Element tree.
// Built once outside the bench so only the serialize call itself is measured.
const XML_PKG_SIMPLE: XmlPkgInput = {
  "w:p": [
    { _attr: { "w:val": "1" } },
    { "w:r": [{ "w:t": "Bold" }] },
    { "w:r": [{ "w:t": "Normal" }] },
  ],
};
const XML_PKG_COMPLEX: XmlPkgInput = {
  "w:document": [
    { _attr: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" } },
    {
      "w:body": [
        {
          "w:p": [
            { "w:pPr": [{ "w:pStyle": [{ _attr: { "w:val": "Title" } }] }] },
            { "w:r": [{ "w:t": "Title Text" }] },
          ],
        },
        {
          "w:p": [{ "w:r": [{ "w:t": "Body paragraph with special chars: & < >" }] }],
        },
      ],
    },
  ],
};

describe("Benchmark: parse — ours vs original xml-js", () => {
  bench("parse() simple XML", () => {
    parse(XML_STRING, { compact: false });
  });

  bench("xml2js (original) simple XML", () => {
    xml2jsOriginal(XML_STRING, { compact: false });
  });

  bench("parse() complex OOXML", () => {
    parse(COMPLEX_XML, { compact: false });
  });

  bench("xml2js (original) complex OOXML", () => {
    xml2jsOriginal(COMPLEX_XML, { compact: false });
  });

  bench("parse() with captureSpaces", () => {
    parse(COMPLEX_XML, { compact: false, captureSpacesBetweenElements: true });
  });

  bench("xml2js (original) with captureSpaces", () => {
    xml2jsOriginal(COMPLEX_XML, { compact: false, captureSpacesBetweenElements: true });
  });

  bench("txml simple XML", () => {
    txmlParse(XML_STRING);
  });

  bench("txml complex OOXML", () => {
    txmlParse(COMPLEX_XML);
  });

  if (bunXml) {
    bench("Bun.XML.parse simple XML", () => {
      bunXml.parse(XML_STRING);
    });

    bench("Bun.XML.parse complex OOXML", () => {
      bunXml.parse(COMPLEX_XML);
    });
  }
});

describe("Benchmark: stringify — ours vs xml-js & xml", () => {
  const parsedSimple = parse(XML_STRING, { compact: false });
  const parsedComplex = parse(COMPLEX_XML, { compact: false });
  // txml consumes its own node format ({tagName, attributes, children}).
  const txmlParsedSimple = txmlParse(XML_STRING);
  const txmlParsedComplex = txmlParse(COMPLEX_XML);

  bench("stringify() simple element", () => {
    stringify(parsedSimple);
  });

  bench("js2xml (original) simple element", () => {
    js2xmlOriginal(parsedSimple);
  });

  bench("xml (npm) simple element", () => {
    xmlPkg(XML_PKG_SIMPLE);
  });

  bench("stringify() complex OOXML", () => {
    stringify(parsedComplex);
  });

  bench("js2xml (original) complex OOXML", () => {
    js2xmlOriginal(parsedComplex);
  });

  bench("xml (npm) complex OOXML", () => {
    xmlPkg(XML_PKG_COMPLEX);
  });

  bench("txml stringify() simple element", () => {
    txmlStringify(txmlParsedSimple);
  });

  bench("txml stringify() complex OOXML", () => {
    txmlStringify(txmlParsedComplex);
  });

  // txml skips entity encoding by default (encodeEntities: false), which is
  // not output-equivalent to ours — this is the fair comparison.
  bench("txml stringify() simple element (entities)", () => {
    txmlStringify(txmlParsedSimple, { encodeEntities: true });
  });

  bench("txml stringify() complex OOXML (entities)", () => {
    txmlStringify(txmlParsedComplex, { encodeEntities: true });
  });

  if (bunXml) {
    bench("Bun.XML.stringify simple element", () => {
      bunXml.stringify(BUN_SIMPLE);
    });

    bench("Bun.XML.stringify complex OOXML", () => {
      bunXml.stringify(BUN_COMPLEX);
    });
  }
});
