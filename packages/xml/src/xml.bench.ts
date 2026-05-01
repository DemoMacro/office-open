import { bench, describe } from "vitest";
import xmlOriginal from "xml";
import { xml2js as xml2jsOriginal, js2xml as js2xmlOriginal } from "xml-js";

import { toElement } from "../src/convert";
import { xml2js } from "../src/parse";
import { xml } from "../src/serialize";
import { js2xml } from "../src/stringify";

// Simple OOXML-like fixture for benchmarking
const SIMPLE_ELEMENT = { "w:t": "Hello World" };
const NESTED_ELEMENT = {
    "w:p": [
        { _attr: { "w:val": "1" } },
        { "w:r": [{ "w:t": "Bold text content" }] },
        { "w:r": [{ "w:t": "Normal text content" }] },
    ],
};
const XML_STRING = '<w:p w:val="1"><w:r><w:t>Bold</w:t></w:r><w:r><w:t>Normal</w:t></w:r></w:p>';
const COMPLEX_XML =
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Title Text</w:t></w:r></w:p><w:p><w:r><w:t>Body paragraph with special chars: &amp; &lt; &gt;</w:t></w:r></w:p></w:body></w:document>';

describe("Benchmark: serialize — ours vs original xml", () => {
    bench("xml() simple element", () => {
        xml(SIMPLE_ELEMENT);
    });

    bench("xml (original) simple element", () => {
        xmlOriginal(SIMPLE_ELEMENT);
    });

    bench("xml() nested element", () => {
        xml(NESTED_ELEMENT);
    });

    bench("xml (original) nested element", () => {
        xmlOriginal(NESTED_ELEMENT);
    });

    bench("xml() nested with declaration", () => {
        xml(NESTED_ELEMENT, { declaration: { encoding: "UTF-8", standalone: "yes" } });
    });

    bench("xml (original) nested with declaration", () => {
        xmlOriginal(NESTED_ELEMENT, { declaration: { encoding: "UTF-8", standalone: "yes" } });
    });
});

describe("Benchmark: parse — ours vs original xml-js", () => {
    bench("xml2js() simple XML", () => {
        xml2js(XML_STRING, { compact: false });
    });

    bench("xml2js (original) simple XML", () => {
        xml2jsOriginal(XML_STRING, { compact: false });
    });

    bench("xml2js() complex OOXML", () => {
        xml2js(COMPLEX_XML, { compact: false });
    });

    bench("xml2js (original) complex OOXML", () => {
        xml2jsOriginal(COMPLEX_XML, { compact: false });
    });

    bench("xml2js() with captureSpaces", () => {
        xml2js(COMPLEX_XML, { compact: false, captureSpacesBetweenElements: true });
    });

    bench("xml2js (original) with captureSpaces", () => {
        xml2jsOriginal(COMPLEX_XML, { compact: false, captureSpacesBetweenElements: true });
    });
});

describe("Benchmark: stringify — ours vs original xml-js", () => {
    const parsedSimple = xml2js(XML_STRING, { compact: false });
    const parsedComplex = xml2js(COMPLEX_XML, { compact: false });

    bench("js2xml() simple element", () => {
        js2xml(parsedSimple);
    });

    bench("js2xml (original) simple element", () => {
        js2xmlOriginal(parsedSimple);
    });

    bench("js2xml() complex OOXML", () => {
        js2xml(parsedComplex);
    });

    bench("js2xml (original) complex OOXML", () => {
        js2xmlOriginal(parsedComplex);
    });
});

describe("Benchmark: convert — toElement vs bridge", () => {
    bench("toElement() simple", () => {
        toElement(SIMPLE_ELEMENT);
    });

    bench("toElement() nested", () => {
        toElement(NESTED_ELEMENT);
    });

    bench("xml() + xml2js() bridge (simple)", () => {
        xml2js(xml(SIMPLE_ELEMENT), { compact: false });
    });

    bench("xml() + xml2js() bridge (nested)", () => {
        xml2js(xml(NESTED_ELEMENT), { compact: false });
    });
});
