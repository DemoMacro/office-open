import { describe, it, expect } from "vite-plus/test";

import { validateOpcConsistency, type OpcIssue } from "./opc-consistency";
import { DOCX_PARTS, PPTX_PARTS } from "./part-registry";

const NS_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

const CT = {
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
  doc: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  core: "application/vnd.openxmlformats-package.core-properties+xml",
  app: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  custom: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
} as const;

const REL = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  app: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  fontTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
} as const;

const def = (ext: string, ct: string) => `<Default Extension="${ext}" ContentType="${ct}"/>`;
const ov = (partName: string, ct: string) =>
  `<Override PartName="${partName}" ContentType="${ct}"/>`;
const rel = (id: string, type: string, target: string, targetMode?: string) =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"${targetMode ? ` TargetMode="${targetMode}"` : ""}/>`;

/** Minimal but fully consistent DOCX package — the green baseline for mutations. */
function contentTypesXml(extra = ""): string {
  return (
    `<Types xmlns="${NS_TYPES}">` +
    def("rels", CT.rels) +
    def("xml", CT.xml) +
    ov("/word/document.xml", CT.doc) +
    ov("/word/styles.xml", CT.styles) +
    ov("/word/numbering.xml", CT.numbering) +
    ov("/word/footnotes.xml", CT.footnotes) +
    ov("/word/endnotes.xml", CT.endnotes) +
    ov("/word/settings.xml", CT.settings) +
    ov("/word/fontTable.xml", CT.fontTable) +
    ov("/docProps/core.xml", CT.core) +
    ov("/docProps/app.xml", CT.app) +
    ov("/docProps/custom.xml", CT.custom) +
    extra +
    `</Types>`
  );
}

function rootRels(): string {
  return (
    `<Relationships xmlns="${NS_RELS}">` +
    rel("rId1", REL.officeDocument, "word/document.xml") +
    rel("rId2", REL.core, "docProps/core.xml") +
    rel("rId3", REL.app, "docProps/app.xml") +
    `</Relationships>`
  );
}

function documentRels(overrides: Record<string, string> = {}): string {
  const map: Record<string, string> = {
    rId1: rel("rId1", REL.styles, "styles.xml"),
    rId2: rel("rId2", REL.numbering, "numbering.xml"),
    rId3: rel("rId3", REL.footnotes, "footnotes.xml"),
    rId4: rel("rId4", REL.endnotes, "endnotes.xml"),
    rId5: rel("rId5", REL.settings, "settings.xml"),
    rId6: rel("rId6", REL.fontTable, "fontTable.xml"),
    ...overrides,
  };
  return `<Relationships xmlns="${NS_RELS}">${Object.values(map).join("")}</Relationships>`;
}

function baseDocxEntries(): Map<string, string> {
  return new Map<string, string>([
    ["[Content_Types].xml", contentTypesXml()],
    ["_rels/.rels", rootRels()],
    ["word/_rels/document.xml.rels", documentRels()],
    ["word/document.xml", "<w:document/>"],
    ["word/styles.xml", "<w:styles/>"],
    ["word/numbering.xml", "<w:numbering/>"],
    ["word/footnotes.xml", "<w:footnotes/>"],
    ["word/endnotes.xml", "<w:endnotes/>"],
    ["word/settings.xml", "<w:settings/>"],
    ["word/fontTable.xml", "<w:fonts/>"],
    ["docProps/core.xml", "<cp:coreProperties/>"],
    ["docProps/app.xml", "<Properties/>"],
    ["docProps/custom.xml", "<Properties/>"],
  ]);
}

const codes = (issues: readonly OpcIssue[]): string[] => issues.map((i) => i.code);

describe("validateOpcConsistency", () => {
  it("returns no issues for a consistent package", () => {
    expect(validateOpcConsistency(baseDocxEntries(), DOCX_PARTS)).toEqual([]);
  });

  it("O2 — flags a missing always part", () => {
    const entries = baseDocxEntries();
    entries.delete("word/document.xml");
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O2");
    expect(issues.some((i) => i.code === "O2" && i.part === "word/document.xml")).toBe(true);
  });

  it("O3 — flags a dangling internal relationship target", () => {
    const entries = baseDocxEntries();
    entries.set(
      "word/_rels/document.xml.rels",
      documentRels({ rId1: rel("rId1", REL.styles, "missing.xml") }),
    );
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O3");
    expect(issues.some((i) => i.code === "O3" && i.part === "word/_rels/document.xml.rels")).toBe(
      true,
    );
  });

  it("O3 — ignores external relationship targets", () => {
    const entries = baseDocxEntries();
    entries.set(
      "_rels/.rels",
      `<Relationships xmlns="${NS_RELS}">` +
        rel("rId1", REL.officeDocument, "word/document.xml") +
        rel("rId2", REL.core, "docProps/core.xml") +
        rel("rId3", REL.app, "docProps/app.xml") +
        rel("rId4", REL.hyperlink, "https://example.com", "External") +
        `</Relationships>`,
    );
    expect(codes(validateOpcConsistency(entries, DOCX_PARTS))).not.toContain("O3");
  });

  it("O5 — flags a part with no Override and no Default for its extension", () => {
    const entries = baseDocxEntries();
    entries.set("word/data.bin", "binary");
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O5");
    expect(issues.some((i) => i.code === "O5" && i.part === "word/data.bin")).toBe(true);
  });

  it("O6 — flags a stale Override pointing at an absent part", () => {
    const entries = baseDocxEntries();
    entries.set(
      "[Content_Types].xml",
      contentTypesXml(
        ov(
          "/word/comments.xml",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        ),
      ),
    );
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O6");
    expect(issues.some((i) => i.code === "O6" && i.part === "/word/comments.xml")).toBe(true);
  });

  it("O7 — flags duplicate relationship Ids", () => {
    const entries = baseDocxEntries();
    entries.set(
      "word/_rels/document.xml.rels",
      documentRels({ rId2: rel("rId1", REL.numbering, "numbering.xml") }),
    );
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O7");
    expect(issues.some((i) => i.code === "O7" && i.part === "word/_rels/document.xml.rels")).toBe(
      true,
    );
  });

  it("O1 — flags an undeclared orphan part as warn", () => {
    const entries = baseDocxEntries();
    entries.set("word/mystery.xml", "<x/>");
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).toContain("O1");
    const o1 = issues.find((i) => i.code === "O1");
    expect(o1?.severity).toBe("warn");
  });

  it("O1 — suppresses whitelisted media paths", () => {
    const entries = baseDocxEntries();
    entries.set("word/media/image1.png", "<png-bytes/>");
    const issues = validateOpcConsistency(entries, DOCX_PARTS);
    expect(codes(issues)).not.toContain("O1");
  });

  it("resolves ../ segments in relationship targets", () => {
    // Real-world PPTX: slide1.xml.rels references its layout via ../slideLayouts/.
    const entries = new Map<string, string>([
      [
        "[Content_Types].xml",
        `<Types xmlns="${NS_TYPES}">${def("rels", CT.rels)}${def("xml", CT.xml)}${ov("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")}${ov("/ppt/slides/slide1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")}</Types>`,
      ],
      [
        "_rels/.rels",
        `<Relationships xmlns="${NS_RELS}">${rel("rId1", REL.officeDocument, "ppt/presentation.xml")}</Relationships>`,
      ],
      [
        "ppt/_rels/presentation.xml.rels",
        `<Relationships xmlns="${NS_RELS}">${rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", "slides/slide1.xml")}</Relationships>`,
      ],
      [
        "ppt/slides/_rels/slide1.xml.rels",
        `<Relationships xmlns="${NS_RELS}">${rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml")}</Relationships>`,
      ],
      ["ppt/presentation.xml", "<p:presentation/>"],
      ["ppt/slides/slide1.xml", "<p:slide/>"],
      ["ppt/slideLayouts/slideLayout1.xml", "<p:sldLayout/>"],
    ]);
    const issues = validateOpcConsistency(entries, PPTX_PARTS);
    expect(issues.filter((i) => i.code === "O3")).toHaveLength(0);
  });
});
