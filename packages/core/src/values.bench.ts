import { describe, bench } from "vite-plus/test";

import { convertMillimetersToTwip, convertInchesToTwip } from "./converters";
import { uniqueId, uniqueUuid, hashedId, uniqueNumericIdCreator } from "./id-generators";
import { decimalNumber, hexColorValue, hpsMeasureValue, universalMeasureValue } from "./values";
import type { IContext } from "./xml-components/base";
import { OnOffElement, BuilderElement, StringContainer } from "./xml-components/elements";

const ctx: IContext = { stack: [] };

describe("values validators", () => {
    bench("decimalNumber", () => {
        decimalNumber(42.7);
    });

    bench("hexColorValue (6-char hex)", () => {
        hexColorValue("FF0000");
    });

    bench("hexColorValue (# prefix)", () => {
        hexColorValue("#00FF00");
    });

    bench("universalMeasureValue", () => {
        universalMeasureValue("10.5mm");
    });

    bench("hpsMeasureValue (number)", () => {
        hpsMeasureValue(24);
    });

    bench("hpsMeasureValue (string)", () => {
        hpsMeasureValue("12pt");
    });
});

describe("converters", () => {
    bench("convertMillimetersToTwip", () => {
        convertMillimetersToTwip(25.4);
    });

    bench("convertInchesToTwip", () => {
        convertInchesToTwip(1);
    });
});

describe("id generators", () => {
    bench("uniqueId", () => {
        uniqueId();
    });

    bench("uniqueUuid", () => {
        uniqueUuid();
    });

    bench("hashedId", () => {
        hashedId("benchmark-test-input");
    });

    bench("uniqueNumericIdCreator", () => {
        const gen = uniqueNumericIdCreator();
        gen();
    });
});

describe("xml component prepForXml", () => {
    bench("OnOffElement (true)", () => {
        new OnOffElement("w:b", true).prepForXml(ctx);
    });

    bench("OnOffElement (false)", () => {
        new OnOffElement("w:b", false).prepForXml(ctx);
    });

    bench("BuilderElement with attributes", () => {
        new BuilderElement({
            name: "w:pPr",
            attributes: { style: { key: "w:val", value: "Heading1" } },
        }).prepForXml(ctx);
    });

    bench("BuilderElement with children", () => {
        new BuilderElement({
            name: "w:p",
            children: [new StringContainer("w:t", "Hello World")],
        }).prepForXml(ctx);
    });
});
