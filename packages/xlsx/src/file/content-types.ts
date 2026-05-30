/**
 * Content Types module for XLSX packages.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

const XLSX_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSX_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const XLSX_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const XLSX_SHARED_STRINGS =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const XLSX_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const XLSX_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

type EntryType = "Default" | "Override";

interface ContentEntry {
  readonly type: EntryType;
  readonly contentType: string;
  readonly key: string;
}

const STATIC_ENTRIES: readonly ContentEntry[] = [
  {
    type: "Default",
    contentType: "application/vnd.openxmlformats-package.relationships+xml",
    key: "rels",
  },
  { type: "Default", contentType: "application/xml", key: "xml" },
  { type: "Default", contentType: "image/png", key: "png" },
  { type: "Default", contentType: "image/jpeg", key: "jpeg" },
  { type: "Override", contentType: XLSX_MAIN, key: "/xl/workbook.xml" },
  {
    type: "Override",
    contentType: "application/vnd.openxmlformats-package.core-properties+xml",
    key: "/docProps/core.xml",
  },
  {
    type: "Override",
    contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    key: "/docProps/app.xml",
  },
];

const STATIC_CHILDREN: IXmlableObject[] = [
  { _attr: { xmlns: "http://schemas.openxmlformats.org/package/2006/content-types" } },
  ...STATIC_ENTRIES.map((e) => {
    if (e.type === "Default") {
      return { Default: { _attr: { ContentType: e.contentType, Extension: e.key } } };
    }
    return { Override: { _attr: { ContentType: e.contentType, PartName: e.key } } };
  }),
];

// Pre-compiled static XML fragment (module-level constant)
const STATIC_XML = STATIC_ENTRIES.map((e) =>
  e.type === "Default"
    ? `<Default ContentType="${e.contentType}" Extension="${e.key}"/>`
    : `<Override ContentType="${e.contentType}" PartName="${e.key}"/>`,
).join("");

export class ContentTypes extends BaseXmlComponent {
  private readonly dynamicEntries: ContentEntry[] = [];

  public constructor() {
    super("Types");
  }

  public addWorksheet(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_WORKSHEET,
      key: `/xl/worksheets/sheet${index}.xml`,
    });
  }

  public addStyles(): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_STYLES,
      key: "/xl/styles.xml",
    });
  }

  public addSharedStrings(): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_SHARED_STRINGS,
      key: "/xl/sharedStrings.xml",
    });
  }

  public addTheme(index: number = 1): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_THEME,
      key: `/xl/theme/theme${index}.xml`,
    });
  }

  public addChart(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_CHART,
      key: `/xl/charts/chart${index}.xml`,
    });
  }

  public addDrawing(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
      key: `/xl/drawings/drawing${index}.xml`,
    });
  }

  public override prepForXml(_context: Context): IXmlableObject {
    const children = [...STATIC_CHILDREN];
    for (const e of this.dynamicEntries) {
      if (e.type === "Default") {
        children.push({
          Default: { _attr: { ContentType: e.contentType, Extension: e.key } },
        });
      } else {
        children.push({
          Override: { _attr: { ContentType: e.contentType, PartName: e.key } },
        });
      }
    }
    return { Types: children };
  }

  public override toXml(_context: Context): string {
    const p: string[] = [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      STATIC_XML,
    ];
    for (const e of this.dynamicEntries) {
      if (e.type === "Default") {
        p.push(`<Default ContentType="${e.contentType}" Extension="${e.key}"/>`);
      } else {
        p.push(`<Override ContentType="${e.contentType}" PartName="${e.key}"/>`);
      }
    }
    p.push("</Types>");
    return p.join("");
  }
}
