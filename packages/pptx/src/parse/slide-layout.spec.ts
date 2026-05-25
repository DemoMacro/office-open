import { xml2js } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseSlideLayoutType } from "./slide-layout";

const NS = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

function layoutEl(name: string) {
  return xml2js(`<p:sldLayout ${NS}><p:cSld name="${name}"><p:spTree/></p:cSld></p:sldLayout>`)
    .elements![0];
}

function layoutElNoName() {
  return xml2js(`<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`).elements![0];
}

describe("parseSlideLayoutType", () => {
  it("maps Title Slide to title", () => {
    expect(parseSlideLayoutType(layoutEl("Title Slide"))).to.equal("title");
  });

  it("maps Blank to blank", () => {
    expect(parseSlideLayoutType(layoutEl("Blank"))).to.equal("blank");
  });

  it("maps Title and Content to obj", () => {
    expect(parseSlideLayoutType(layoutEl("Title and Content"))).to.equal("obj");
  });

  it("maps Section Header to secHead", () => {
    expect(parseSlideLayoutType(layoutEl("Section Header"))).to.equal("secHead");
  });

  it("returns custom name for unknown layout", () => {
    expect(parseSlideLayoutType(layoutEl("My Custom Layout"))).to.equal("My Custom Layout");
  });

  it("returns blank when no name attribute", () => {
    expect(parseSlideLayoutType(layoutElNoName())).to.equal("blank");
  });
});
