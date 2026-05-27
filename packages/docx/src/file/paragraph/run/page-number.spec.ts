import { describe, expect, it } from "vite-plus/test";

import {
  buildCurrentSection,
  buildNumberOfPages,
  buildNumberOfPagesSection,
  buildPage,
} from "./page-number";

describe("buildPage", () => {
  it("should build PAGE instrText", () => {
    expect(buildPage()).toEqual({
      "w:instrText": [{ _attr: { "xml:space": "preserve" } }, "PAGE"],
    });
  });
});

describe("buildNumberOfPages", () => {
  it("should build NUMPAGES instrText", () => {
    expect(buildNumberOfPages()).toEqual({
      "w:instrText": [{ _attr: { "xml:space": "preserve" } }, "NUMPAGES"],
    });
  });
});

describe("buildNumberOfPagesSection", () => {
  it("should build SECTIONPAGES instrText", () => {
    expect(buildNumberOfPagesSection()).toEqual({
      "w:instrText": [{ _attr: { "xml:space": "preserve" } }, "SECTIONPAGES"],
    });
  });
});

describe("buildCurrentSection", () => {
  it("should build SECTION instrText", () => {
    expect(buildCurrentSection()).toEqual({
      "w:instrText": [{ _attr: { "xml:space": "preserve" } }, "SECTION"],
    });
  });
});
