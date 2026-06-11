import { element } from "@office-open/xml";
import { describe, bench } from "vite-plus/test";

import { uniqueId, uniqueUuid, hashedId, uniqueNumericIdCreator } from "./util/generators";
import {
  decimalNumber,
  hexColorValue,
  hpsMeasureValue,
  universalMeasureValue,
} from "./util/values";

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

  const numericId = uniqueNumericIdCreator();
  bench("uniqueNumericId", () => {
    numericId();
  });
});

describe("xml element builder", () => {
  bench("element self-closing", () => {
    element("w:b");
  });

  bench("element with attributes", () => {
    element("w:pPr", { "w:val": "Heading1" });
  });

  bench("element with children", () => {
    element("w:p", undefined, ["Hello World"]);
  });
});
