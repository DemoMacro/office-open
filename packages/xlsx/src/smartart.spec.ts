import { unzipSync } from "@office-open/core/opc";
import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";
import type { WorkbookOptions } from "@parts/file";
import type { WorksheetSmartArtOptions } from "@parts/worksheet";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook } from "./generate";
import { parseWorkbookSync } from "./parse";

const customLayout: LayoutDefinitionOptions = {
  uniqueId: "urn:office-open:test/layout/default",
  layoutNode: { children: [{ algorithm: { type: "snake" } }] },
};
const customStyle: StyleDefinitionOptions = {
  uniqueId: "urn:office-open:test/style/default",
  styleLabels: [{ name: "node0", textProperties: {} }],
};
const customColors: ColorDefinitionOptions = {
  uniqueId: "urn:office-open:test/colors/accent1_2",
  styleLabels: [
    {
      name: "node0",
      fillColorList: { meth: "cycle", colors: [{ value: "accent1" }] },
    },
  ],
};

function smartArtOptions(): WorksheetSmartArtOptions {
  return {
    col: 1,
    row: 1,
    name: "Test SmartArt",
    nodes: [{ text: "Root", children: [{ text: "Child" }] }, { text: "Leaf" }],
    layout: customLayout,
    style: customStyle,
    color: customColors,
  };
}

async function generate(options: WorkbookOptions): Promise<Uint8Array> {
  return (await generateWorkbook(options, {
    type: "uint8array",
    reproducible: { date: "2026-01-01T00:00:00.000Z" },
  })) as Uint8Array;
}

async function roundTrip(options: WorkbookOptions): Promise<WorkbookOptions> {
  return parseWorkbookSync(await generate(options));
}

describe("XLSX SmartArt", () => {
  it("authors canonical diagram parts, relationships, and content types", async () => {
    const archive = unzipSync(await generate({ worksheets: [{ smartArts: [smartArtOptions()] }] }));
    const paths = [
      "xl/diagrams/data1.xml",
      "xl/diagrams/layout1.xml",
      "xl/diagrams/quickStyle1.xml",
      "xl/diagrams/colors1.xml",
    ];
    const parts = Object.fromEntries(
      paths.map((path) => [path, new TextDecoder().decode(archive[path]!)]),
    );

    expect(parts["xl/diagrams/data1.xml"]).toContain(">Root<");
    expect(parts["xl/diagrams/data1.xml"]).toContain(">Child<");
    expect(parts["xl/diagrams/data1.xml"]).toContain("layout/default");
    expect(parts["xl/diagrams/layout1.xml"]).toContain(
      'uniqueId="urn:office-open:test/layout/default"',
    );
    expect(parts["xl/diagrams/quickStyle1.xml"]).toContain(
      'uniqueId="urn:office-open:test/style/default"',
    );
    expect(parts["xl/diagrams/colors1.xml"]).toContain(
      'uniqueId="urn:office-open:test/colors/accent1_2"',
    );

    const rels = new TextDecoder().decode(archive["xl/drawings/_rels/drawing1.xml.rels"]!);
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"',
    );
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"',
    );
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"',
    );
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"',
    );

    const contentTypes = new TextDecoder().decode(archive["[Content_Types].xml"]!);
    expect(contentTypes).toContain(
      'PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"',
    );
    expect(contentTypes).toContain(
      'PartName="/xl/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"',
    );
    expect(contentTypes).toContain(
      'PartName="/xl/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"',
    );
    expect(contentTypes).toContain(
      'PartName="/xl/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"',
    );
  });

  it("keeps custom definition parse → generate → parse stable", async () => {
    const first = await roundTrip({ worksheets: [{ smartArts: [smartArtOptions()] }] });
    const second = await roundTrip(first);
    const firstSmartArt = first.worksheets?.[0]?.smartArts?.[0];
    const secondSmartArt = second.worksheets?.[0]?.smartArts?.[0];

    expect(secondSmartArt).toEqual(firstSmartArt);
    expect(firstSmartArt).toMatchObject({
      nodes: [{ text: "Root", children: [{ text: "Child" }] }, { text: "Leaf" }],
      dataPath: "xl/diagrams/data1.xml",
      layoutPath: "xl/diagrams/layout1.xml",
      quickStylePath: "xl/diagrams/quickStyle1.xml",
      colorsPath: "xl/diagrams/colors1.xml",
      layout: { uniqueId: "urn:office-open:test/layout/default" },
      style: { uniqueId: "urn:office-open:test/style/default" },
      color: { uniqueId: "urn:office-open:test/colors/accent1_2" },
    });
    expect(firstSmartArt?.dataModel?.points).toBeDefined();
  });

  it("excludes modeled diagram parts from rawParts and rebuilds canonical output", async () => {
    const parsed = await roundTrip({ worksheets: [{ smartArts: [smartArtOptions()] }] });
    const rawPaths = parsed.rawParts?.map((raw) => raw.path) ?? [];
    expect(rawPaths).not.toContain("xl/diagrams/data1.xml");
    expect(rawPaths).not.toContain("xl/diagrams/layout1.xml");
    expect(rawPaths).not.toContain("xl/diagrams/quickStyle1.xml");
    expect(rawPaths).not.toContain("xl/diagrams/colors1.xml");

    const smartArt = parsed.worksheets?.[0]?.smartArts?.[0];
    if (!smartArt) throw new Error("expected a parsed SmartArt");
    smartArt.nodes.push({ text: "Added" });
    delete smartArt.dataModel;
    const archive = unzipSync(await generate(parsed));
    expect(new TextDecoder().decode(archive["xl/diagrams/data1.xml"]!)).toContain(">Added<");
  });
});
