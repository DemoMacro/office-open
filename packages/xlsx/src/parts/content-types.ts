/**
 * Content Types module for XLSX packages.
 *
 * @module
 */

const XLSX_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const XLSX_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const XLSX_CHARTSHEET =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const XLSX_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const XLSX_SHARED_STRINGS =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const XLSX_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const XLSX_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

type EntryType = "Default" | "Override";

interface ContentEntry {
  type: EntryType;
  contentType: string;
  key: string;
}

const STATIC_ENTRIES: ContentEntry[] = [
  {
    type: "Default",
    contentType: "application/vnd.openxmlformats-package.relationships+xml",
    key: "rels",
  },
  { type: "Default", contentType: "application/xml", key: "xml" },
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

// Pre-compiled static XML fragment (module-level constant)
const STATIC_XML = STATIC_ENTRIES.map((e) =>
  e.type === "Default"
    ? `<Default ContentType="${e.contentType}" Extension="${e.key}"/>`
    : `<Override ContentType="${e.contentType}" PartName="${e.key}"/>`,
).join("");

export class ContentTypes {
  private dynamicEntries: ContentEntry[] = [];

  public addWorksheet(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_WORKSHEET,
      key: `/xl/worksheets/sheet${index}.xml`,
    });
  }

  public addChartsheet(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: XLSX_CHARTSHEET,
      key: `/xl/chartsheets/sheet${index}.xml`,
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

  public addComments(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
      key: `/xl/comments${index}.xml`,
    });
  }

  public addVmlDrawing(): void {
    if (this.dynamicEntries.some((e) => e.type === "Default" && e.key === "vml")) return;
    this.dynamicEntries.push({
      type: "Default",
      contentType: "application/vnd.openxmlformats-officedocument.vmlDrawing",
      key: "vml",
    });
  }

  public addImageType(extension: "png" | "jpeg"): void {
    const contentType = extension === "png" ? "image/png" : "image/jpeg";
    if (this.dynamicEntries.some((e) => e.type === "Default" && e.key === extension)) return;
    this.dynamicEntries.push({ type: "Default", contentType, key: extension });
  }

  public addPivotTable(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
      key: `/xl/pivotTables/pivotTable${index}.xml`,
    });
  }

  public addPivotCacheDefinition(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
      key: `/xl/pivotCache/pivotCacheDefinition${index}.xml`,
    });
  }

  public addPivotCacheRecords(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
      key: `/xl/pivotCache/pivotCacheRecords${index}.xml`,
    });
  }

  public addTable(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
      key: `/xl/tables/table${index}.xml`,
    });
  }

  public addExternalLink(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
      key: `/xl/externalLinks/externalLink${index}.xml`,
    });
  }

  public addCalcChain(): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
      key: "/xl/calcChain.xml",
    });
  }

  public addDialogsheet(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
      key: `/xl/dialogsheets/sheet${index}.xml`,
    });
  }

  public addRevisionHeaders(): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml",
      key: "/xl/revisionHeaders.xml",
    });
  }

  public addRevisionLog(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml",
      key: `/xl/revisions/revision${index}.xml`,
    });
  }

  public addQueryTable(index: number): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
      key: `/xl/queryTables/queryTable${index}.xml`,
    });
  }

  public addMetadata(): void {
    this.dynamicEntries.push({
      type: "Override",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
      key: "/xl/metadata.xml",
    });
  }

  public serialize(): string {
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
