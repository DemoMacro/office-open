/**
 * XLSX Compiler — compiles a File object into a Zippable structure.
 *
 * @module
 */
import { Formatter } from "@export/formatter";
import { CalcChain } from "@file/calc-chain";
import { Chartsheet } from "@file/chartsheet";
import { Comments } from "@file/comments";
import { Drawing, type ImageOptions, type ChartAnchorOptions } from "@file/drawing/drawing";
import { ExternalLinkXml } from "@file/external-link";
import type { File } from "@file/file";
import {
  PivotCacheDefinitionXml,
  PivotCacheRecordsXml,
  PivotTableXml,
  aggregate,
} from "@file/pivot";
import type { PivotSourceData, PivotTableOptions } from "@file/pivot";
import { TableXml } from "@file/table";
import { VmlNotes } from "@file/vml-notes";
import { WorkbookXml, type TablePartReference } from "@file/workbook";
import type { BaseXmlComponent, Context } from "@file/xml-components";
import {
  ChartSpace,
  Relationships,
  TargetModeType,
  compileMapping,
  type XmlifyedFile,
  type Zippable,
} from "@office-open/core";

export class Compiler {
  private readonly formatter = new Formatter();

  public compile(
    file: File,
    overrides: readonly XmlifyedFile[] = [],
    mediaLevel: number = 0,
  ): Zippable {
    const context: Context = { fileData: file, stack: [] };
    const f = this.formatter;

    const mapping: Record<string, { data: string; path: string }> = {};

    const fmt = (component: BaseXmlComponent) => f.formatToXml(component, context);

    // Core properties
    mapping["Properties"] = {
      data: fmt(file.coreProperties),
      path: "docProps/core.xml",
    };

    // App properties
    mapping["AppProperties"] = {
      data: fmt(file.appProperties),
      path: "docProps/app.xml",
    };

    // File-level relationships (_rels/.rels)
    mapping["FileRelationships"] = {
      data: fmt(file.fileRelationships),
      path: "_rels/.rels",
    };

    // Worksheets — formatted BEFORE SharedStrings so strings are registered
    const worksheets = file.worksheets;
    let globalMediaIdx = 0;
    let globalChartIdx = 0;
    let globalPivotIdx = 0;
    let globalPivotCacheIdx = 0;
    let globalTableIdx = 0;
    const pivotCacheDataMap = new Map<string, { cacheId: number; cacheIdx: number }>();
    const calcChain = new CalcChain();
    const allTableParts: TablePartReference[] = [];

    for (let i = 0; i < worksheets.length; i++) {
      const ws = worksheets[i];
      const imgOpts = ws.imageOptions;
      const chartOpts = ws.charts;
      const hlOpts = ws.hyperlinkOptions;
      const sheetName = file.worksheetConfigs[i]?.name ?? `Sheet${i + 1}`;

      // Worksheet uses toXml() fast path (zero-allocation string concat)
      let sheetXml = fmt(ws);

      // Collect formula cells for calcChain
      const sheetIdx = i + 1;
      const wsRows = ws.worksheetRows;
      for (let ri = 0; ri < wsRows.length; ri++) {
        const rowOpts = wsRows[ri];
        const rowNumber = rowOpts.rowNumber ?? ri + 1;
        if (!rowOpts.cells) continue;
        for (let ci = 0; ci < rowOpts.cells.length; ci++) {
          const cell = rowOpts.cells[ci];
          if (!cell.formula) continue;
          const ref = cell.reference ?? columnToLetter(ci + 1) + rowNumber;
          calcChain.addCell({
            r: ref,
            i: sheetIdx,
            a: cell.formula.type === "array",
          });
        }
      }

      const hasMedia = imgOpts.length > 0 || chartOpts.length > 0;
      const hasExternalHyperlinks = hlOpts.some((h) => h.target.type === "external");
      const commentOpts = ws.commentOptions;
      const hasComments = commentOpts.length > 0;
      const pivotOpts = ws.pivotTables;
      const hasPivots = pivotOpts.length > 0;
      const tableOpts = ws.tables;
      const hasTables = tableOpts.length > 0;

      // Worksheet-level relationships
      let wsRels: Relationships | undefined;
      let nextRid = 0;

      if (hasMedia || hasExternalHyperlinks || hasComments || hasPivots || hasTables) {
        wsRels = new Relationships();
      }

      if (hasExternalHyperlinks) {
        for (const hl of hlOpts) {
          if (hl.target.type !== "external") continue;
          const rid = ++nextRid;
          wsRels!.addRelationship(
            rid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            hl.target.url,
            "External",
          );
        }
      }

      if (hasMedia) {
        const drawingImages: ImageOptions[] = [];
        const drawingCharts: ChartAnchorOptions[] = [];
        const drawingRels = new Relationships();
        let rid = 1;

        // Process images
        for (const img of imgOpts) {
          const mediaKey = `image_${globalMediaIdx}`;
          const ext = img.type === "jpeg" || img.type === "jpg" ? "jpeg" : "png";
          file.media.addImage(mediaKey, {
            fileName: `image${globalMediaIdx + 1}.${ext}`,
            type: ext,
            data: img.data,
            width: 0,
            height: 0,
          });

          drawingRels.addRelationship(
            rid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `../media/image${globalMediaIdx + 1}.${ext}`,
          );

          drawingImages.push({
            col: img.col,
            row: img.row,
            rId: `rId${rid}`,
          });
          rid++;
          globalMediaIdx++;
        }

        // Process charts
        for (const chart of chartOpts) {
          const chartKey = `chart_${globalChartIdx}`;
          file.charts.addChart(chartKey, {
            key: chartKey,
            chartSpace: new ChartSpace(chart),
          });

          drawingRels.addRelationship(
            rid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            `../charts/chart${globalChartIdx + 1}.xml`,
          );

          drawingCharts.push({
            col: chart.col,
            row: chart.row,
            rId: `rId${rid}`,
          });
          rid++;
          globalChartIdx++;
        }

        // Generate drawing XML
        const drawing = new Drawing(drawingImages, drawingCharts);
        const drawingIdx = i + 1;
        mapping[`Drawing${i}`] = {
          data: fmt(drawing),
          path: `xl/drawings/drawing${drawingIdx}.xml`,
        };

        // Drawing relationships
        mapping[`DrawingRels${i}`] = {
          data: fmt(drawingRels),
          path: `xl/drawings/_rels/drawing${drawingIdx}.xml.rels`,
        };

        // Insert drawing reference before the closing </worksheet> tag.
        const drawingRid = ++nextRid;
        const closingTag = "</worksheet>";
        sheetXml =
          sheetXml.slice(0, -closingTag.length) + `<drawing r:id="rId${drawingRid}"/>` + closingTag;

        // Add drawing relationship to worksheet rels
        wsRels!.addRelationship(
          drawingRid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          `../drawings/drawing${drawingIdx}.xml`,
        );

        file.contentTypes.addDrawing(drawingIdx);
      }

      // Comments
      if (hasComments) {
        const commentsIdx = i + 1;

        // Comments XML
        const commentsXml = new Comments(commentOpts);
        mapping[`Comments${i}`] = {
          data: fmt(commentsXml),
          path: `xl/comments${commentsIdx}.xml`,
        };

        // VML drawing (required by Excel for legacy comment rendering)
        const vmlNotes = new VmlNotes(commentOpts);
        mapping[`VmlDrawing${i}`] = {
          data: vmlNotes.toXml(),
          path: `xl/drawings/vmlDrawing${commentsIdx}.vml`,
        };

        // Worksheet rels: comments → comments XML, legacyDrawing → VML file
        const commentsRid = ++nextRid;
        wsRels!.addRelationship(
          commentsRid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          `../comments${commentsIdx}.xml`,
        );

        const vmlRid = ++nextRid;
        wsRels!.addRelationship(
          vmlRid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
          `../drawings/vmlDrawing${commentsIdx}.vml`,
        );

        // Insert legacyDrawing reference before closing </worksheet>
        const closingTag = "</worksheet>";
        sheetXml =
          sheetXml.slice(0, -closingTag.length) +
          `<legacyDrawing r:id="rId${vmlRid}"/>` +
          closingTag;

        // Register content types
        file.contentTypes.addComments(commentsIdx);
        file.contentTypes.addVmlDrawing();
      }

      // Pivot tables
      if (hasPivots) {
        for (const pt of pivotOpts) {
          globalPivotIdx++;
          const pivotIdx = globalPivotIdx;

          // Extract source data from source sheet
          const sourceSheet = pt.sourceSheet ?? sheetName;
          const sourceWsIdx = file.worksheetConfigs.findIndex(
            (ws) => (ws.name ?? `Sheet${file.worksheetConfigs.indexOf(ws) + 1}`) === sourceSheet,
          );
          if (sourceWsIdx === -1) continue;

          const sourceWs = worksheets[sourceWsIdx];
          const sourceData = extractPivotSourceData(sourceWs.worksheetRows, pt.source);

          // Deduplicate pivot caches by source reference
          const cacheKey = `${sourceSheet}:${pt.source}`;
          let cacheId: number;
          let cacheIdx: number;
          const existing = pivotCacheDataMap.get(cacheKey);
          if (existing) {
            cacheId = existing.cacheId;
            cacheIdx = existing.cacheIdx;
          } else {
            globalPivotCacheIdx++;
            cacheIdx = globalPivotCacheIdx;
            cacheId = cacheIdx - 1;
            pivotCacheDataMap.set(cacheKey, { cacheId, cacheIdx });

            // Generate pivotCacheDefinition
            const cacheDefRels = new Relationships();
            cacheDefRels.addRelationship(
              1,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
              "pivotCacheRecords1.xml",
            );

            const cacheDef = new PivotCacheDefinitionXml(
              cacheIdx,
              pt.source.split(":")[0] ? pt.source : "A1",
              sourceSheet,
              sourceData,
              "rId1",
            );

            mapping[`PivotCacheDef${cacheIdx}`] = {
              data: fmt(cacheDef),
              path: `xl/pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
            };
            mapping[`PivotCacheDefRels${cacheIdx}`] = {
              data: fmt(cacheDefRels),
              path: `xl/pivotCache/_rels/pivotCacheDefinition${cacheIdx}.xml.rels`,
            };

            // Generate pivotCacheRecords
            const cacheRecords = new PivotCacheRecordsXml(sourceData);
            mapping[`PivotCacheRecords${cacheIdx}`] = {
              data: fmt(cacheRecords),
              path: `xl/pivotCache/pivotCacheRecords${cacheIdx}.xml`,
            };

            // Content types
            file.contentTypes.addPivotCacheDefinition(cacheIdx);
            file.contentTypes.addPivotCacheRecords(cacheIdx);

            // Register in workbook
            const wbPivotRid = file.workbookRelationships.relationshipCount + 1;
            file.workbookRelationships.addRelationship(
              wbPivotRid,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
              `pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
            );
            file.registerPivotCache(cacheId, `rId${wbPivotRid}`);
          }

          // Generate pivotTable
          const pivotTable = new PivotTableXml(pt, sourceData, cacheId);
          mapping[`PivotTable${pivotIdx}`] = {
            data: fmt(pivotTable),
            path: `xl/pivotTables/pivotTable${pivotIdx}.xml`,
          };

          // pivotTable rels → cacheDefinition
          const ptRels = new Relationships();
          ptRels.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
            `../pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
          );
          mapping[`PivotTableRels${pivotIdx}`] = {
            data: fmt(ptRels),
            path: `xl/pivotTables/_rels/pivotTable${pivotIdx}.xml.rels`,
          };

          // Worksheet rels → pivotTable
          const ptRid = ++nextRid;
          wsRels!.addRelationship(
            ptRid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
            `../pivotTables/pivotTable${pivotIdx}.xml`,
          );

          file.contentTypes.addPivotTable(pivotIdx);
        }
      }

      // Tables (list objects)
      const wsTableParts: TablePartReference[] = [];
      if (hasTables) {
        for (const tbl of tableOpts) {
          globalTableIdx++;
          const tableIdx = globalTableIdx;

          // Generate table XML
          const tableXml = new TableXml({
            ...tbl,
            id: tbl.id ?? tableIdx,
          });
          mapping[`Table${tableIdx}`] = {
            data: fmt(tableXml),
            path: `xl/tables/table${tableIdx}.xml`,
          };

          // Worksheet rels → table
          const tblRid = ++nextRid;
          wsRels!.addRelationship(
            tblRid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            `../tables/table${tableIdx}.xml`,
          );

          wsTableParts.push({ rId: `rId${tblRid}` });
          allTableParts.push({ rId: `rId${tblRid}` });

          file.contentTypes.addTable(tableIdx);
        }
      }

      // Pre-render pivot table data into sheetData
      if (hasPivots) {
        const rendered = renderPivotSheetData(
          pivotOpts,
          worksheets,
          file.sharedStrings,
          file.worksheetConfigs,
          sheetName,
        );
        if (rendered.sheetData.length > 0) {
          // Replace empty <sheetData/> or <sheetData></sheetData> with rendered data
          sheetXml = sheetXml.replace(/<sheetData\/>|<sheetData><\/sheetData>/, rendered.sheetData);
          // Inject <dimension> before <sheetViews> (XSD sequence order: dimension before sheetViews)
          if (!sheetXml.includes("<dimension")) {
            sheetXml = sheetXml.replace(
              "<sheetViews",
              `<dimension ref="${rendered.dimensionRef}"/><sheetViews`,
            );
          }
        }
      }

      // Insert tableParts before closing </worksheet> tag
      if (wsTableParts.length > 0) {
        const tablePartsXml = WorkbookXml.buildTablePartsXml(wsTableParts);
        const closingTag = "</worksheet>";
        sheetXml = sheetXml.slice(0, -closingTag.length) + tablePartsXml + closingTag;
      }

      // Write worksheet rels if needed
      if (wsRels) {
        mapping[`WorksheetRels${i}`] = {
          data: fmt(wsRels),
          path: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        };
      }

      mapping[`Worksheet${i}`] = {
        data: sheetXml,
        path: `xl/worksheets/sheet${i + 1}.xml`,
      };
    }

    // Chartsheets — chart-only sheets
    const chartsheetConfigs = file.chartsheetConfigs;
    for (let i = 0; i < chartsheetConfigs.length; i++) {
      const csOpts = chartsheetConfigs[i];
      const chartsheet = new Chartsheet(csOpts);

      // Register chart in the charts collection
      const chartDef = csOpts.chart;
      const csChartGlobalIdx = file.charts.array.length;
      const csChartKey = `cs_chart_${csChartGlobalIdx}`;
      file.charts.addChart(csChartKey, {
        key: csChartKey,
        chartSpace: new ChartSpace({
          type: chartDef.type as
            | "column"
            | "bar"
            | "line"
            | "pie"
            | "doughnut"
            | "area"
            | "scatter"
            | "bubble",
          title: chartDef.title,
          categories: chartDef.categories,
          series: chartDef.series,
        }),
      });

      // Chartsheet relationships: drawing (required)
      const csRels = new Relationships();
      const csDrawingIdx = i + 1;
      csRels.addRelationship(
        1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        `../drawings/drawing${csDrawingIdx}.xml`,
      );
      chartsheet.setDrawingRId("rId1");

      // Drawing rels: chart reference
      const csDrawingRels = new Relationships();
      csDrawingRels.addRelationship(
        1,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        `../charts/chart${csChartGlobalIdx + 1}.xml`,
      );

      // Minimal drawing XML with chart anchor (XSD: absoluteAnchor → pos, ext, graphicFrame, clientData)
      const csDrawingXml =
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<xdr:absoluteAnchor>` +
        `<xdr:pos x="0" y="0"/>` +
        `<xdr:ext cx="9308969" cy="6096000"/>` +
        `<xdr:graphicFrame>` +
        `<xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="9308969" cy="6096000"/></xdr:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/></a:graphicData></a:graphic>` +
        `</xdr:graphicFrame>` +
        `<xdr:clientData/>` +
        `</xdr:absoluteAnchor>` +
        `</xdr:wsDr>`;

      mapping[`ChartsheetDrawing${i}`] = {
        data: csDrawingXml,
        path: `xl/drawings/drawing${csDrawingIdx}.xml`,
      };
      mapping[`ChartsheetDrawingRels${i}`] = {
        data: fmt(csDrawingRels),
        path: `xl/drawings/_rels/drawing${csDrawingIdx}.xml.rels`,
      };
      file.contentTypes.addDrawing(csDrawingIdx);

      mapping[`ChartsheetRels${i}`] = {
        data: fmt(csRels),
        path: `xl/chartsheets/_rels/sheet${i + 1}.xml.rels`,
      };
      mapping[`Chartsheet${i}`] = {
        data: chartsheet.toXml(context),
        path: `xl/chartsheets/sheet${i + 1}.xml`,
      };
      file.contentTypes.addChartsheet(i + 1);
    }

    // Workbook XML
    let workbookXml = fmt(file.workbookXml);

    // External links — generate XML files and inject externalReferences into workbook
    const extLinks = file.externalLinks;
    if (extLinks.length > 0) {
      const extRefs: { rId: string }[] = [];
      for (let ei = 0; ei < extLinks.length; ei++) {
        const elIdx = ei + 1;
        const elRid = file.workbookRelationships.relationshipCount + 1;
        file.workbookRelationships.addRelationship(
          elRid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
          `externalLinks/externalLink${elIdx}.xml`,
        );

        // Create the rels file for this external link (points to the external workbook)
        const elOpts = extLinks[ei];
        let bookRId: string | undefined;
        if (elOpts.externalBook?.target) {
          const elRels = new Relationships();
          elRels.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
            elOpts.externalBook.target,
            TargetModeType.EXTERNAL,
          );
          bookRId = "rId1";
          mapping[`ExternalLinkRels${elIdx}`] = {
            data: fmt(elRels),
            path: `xl/externalLinks/_rels/externalLink${elIdx}.xml.rels`,
          };
        }

        // Generate the external link XML
        const elXml = new ExternalLinkXml({ ...elOpts, bookRId });
        mapping[`ExternalLink${elIdx}`] = {
          data: fmt(elXml),
          path: `xl/externalLinks/externalLink${elIdx}.xml`,
        };

        extRefs.push({ rId: `rId${elRid}` });
        file.contentTypes.addExternalLink(elIdx);
      }

      // Inject externalReferences into workbook XML (replace the placeholder)
      const extRefsXml = WorkbookXml.buildExternalReferencesXml(extRefs);
      workbookXml = workbookXml.replace("<!--EXTERNAL_REFS-->", extRefsXml);
    } else {
      // Remove the placeholder if no external links
      workbookXml = workbookXml.replace("<!--EXTERNAL_REFS-->", "");
    }

    mapping["Workbook"] = {
      data: workbookXml,
      path: "xl/workbook.xml",
    };

    // Workbook relationships
    mapping["WorkbookRelationships"] = {
      data: fmt(file.workbookRelationships),
      path: "xl/_rels/workbook.xml.rels",
    };

    // Shared Strings — AFTER worksheets so all strings are collected
    const sharedStrings = file.sharedStrings;
    if (sharedStrings.count > 0) {
      mapping["SharedStrings"] = {
        data: fmt(sharedStrings),
        path: "xl/sharedStrings.xml",
      };
    }

    // Register predefined DXFs before styles XML is generated
    for (const dxf of file.dxfEntries) {
      file.registerDxf(dxf);
    }

    // Styles
    mapping["Styles"] = {
      data: fmt(file.styles),
      path: "xl/styles.xml",
    };

    // Theme
    mapping["Theme"] = {
      data: fmt(file.theme),
      path: "xl/theme/theme1.xml",
    };

    // Charts — AFTER worksheets so charts are registered
    for (let i = 0; i < file.charts.array.length; i++) {
      const chartData = file.charts.array[i];
      mapping[`Chart${i}`] = {
        data: fmt(chartData.chartSpace),
        path: `xl/charts/chart${i + 1}.xml`,
      };
      file.contentTypes.addChart(i + 1);
    }

    // Calculation chain — auto-generated from formula cells
    if (calcChain.count > 0) {
      mapping["CalcChain"] = {
        data: calcChain.toXml(context),
        path: "xl/calcChain.xml",
      };
      file.contentTypes.addCalcChain();
      const calcChainRid = file.workbookRelationships.relationshipCount + 1;
      file.workbookRelationships.addRelationship(
        calcChainRid,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        "calcChain.xml",
      );
    }

    // Register image content types
    const imageExts = new Set<string>();
    for (const img of file.media.array) {
      const ext = img.fileName.endsWith(".png") ? "png" : "jpeg";
      if (!imageExts.has(ext)) {
        imageExts.add(ext);
        file.contentTypes.addImageType(ext as "png" | "jpeg");
      }
    }

    // Content Types — must be last
    mapping["ContentTypes"] = {
      data: fmt(file.contentTypes),
      path: "[Content_Types].xml",
    };

    // Convert mapping to Zippable
    const mediaFiles: Array<{ data: Uint8Array; path: string }> = [];
    for (const img of file.media.array) {
      mediaFiles.push({ data: img.data, path: `xl/media/${img.fileName}` });
    }

    return compileMapping(mapping, overrides, mediaFiles, mediaLevel);
  }
}

/**
 * Extract source data from worksheet rows for pivot cache generation.
 */
function extractPivotSourceData(
  rows: readonly import("@file/worksheet").RowOptions[],
  sourceRef: string,
): PivotSourceData {
  // Parse source range like "A1:D11"
  const parts = sourceRef.split(":");
  const startMatch = parts[0]?.match(/^([A-Z]+)(\d+)$/);
  const endMatch = parts[1]?.match(/^([A-Z]+)(\d+)$/);
  if (!startMatch) {
    return { fieldNames: [], records: [] };
  }

  const startRow = parseInt(startMatch[2], 10) - 1;
  const endRow = endMatch ? parseInt(endMatch[2], 10) - 1 : startRow;
  const startCol = colLetterToIndex(startMatch[1]);
  const endCol = endMatch ? colLetterToIndex(endMatch[1]) : startCol;
  const colCount = endCol - startCol + 1;

  // First row is headers
  const headerRow = rows[startRow];
  const fieldNames: string[] = [];
  if (headerRow?.cells) {
    for (let c = startCol; c <= endCol && c < headerRow.cells.length; c++) {
      fieldNames.push(String(headerRow.cells[c]?.value ?? `Col${c}`));
    }
  }

  // Remaining rows are data
  const records: (string | number)[][] = [];
  for (let r = startRow + 1; r <= endRow; r++) {
    const row = rows[r];
    if (!row?.cells) continue;
    const record: (string | number)[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const val = row.cells[c]?.value;
      if (typeof val === "number") {
        record.push(val);
      } else if (val instanceof Date) {
        record.push(val.getTime());
      } else {
        record.push(String(val ?? ""));
      }
    }
    if (record.length === colCount) {
      records.push(record);
    }
  }

  return { fieldNames, records };
}

function colLetterToIndex(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Pre-render pivot table output as sheetData cells.
 * Excel requires this so the pivot table is visible on open without manual refresh.
 */
function renderPivotSheetData(
  pivotOpts: readonly PivotTableOptions[],
  worksheets: readonly import("@file/worksheet").Worksheet[],
  sharedStrings: import("@file/shared-strings").SharedStrings,
  worksheetConfigs: readonly import("@file/worksheet").WorksheetOptions[],
  currentSheetName: string,
): { sheetData: string; dimensionRef: string } {
  // Collect cells by row number to merge cells from different pivot tables on the same row
  const rowCells = new Map<number, string[]>();
  let maxRow = 0;
  let maxCol = 0;
  let minRow = Infinity;
  let minCol = Infinity;

  for (const pt of pivotOpts) {
    const location = pt.location ?? "A3";
    const locMatch = location.match(/^([A-Z]+)(\d+)$/);
    if (!locMatch) continue;

    const startCol = colLetterToIndex(locMatch[1]);
    const startRow = parseInt(locMatch[2], 10);
    const rowFieldNames = pt.rows;
    const dataFields = pt.data;

    const sourceSheetName = pt.sourceSheet ?? currentSheetName;
    const sourceWsIdx = worksheetConfigs.findIndex(
      (ws) => (ws.name ?? `Sheet${worksheetConfigs.indexOf(ws) + 1}`) === sourceSheetName,
    );
    if (sourceWsIdx === -1) continue;

    const sourceRows = worksheets[sourceWsIdx]?.worksheetRows ?? [];
    const sourceData = extractPivotSourceData(sourceRows, pt.source);
    if (sourceData.fieldNames.length === 0) continue;

    const fields = sourceData.fieldNames;
    const rowFieldIndices = rowFieldNames.map((n) => fields.indexOf(n));
    const dataFieldIndices = dataFields.map((df) => fields.indexOf(df.field));

    if (rowFieldIndices.some((idx) => idx === -1)) continue;
    if (dataFieldIndices.some((idx) => idx === -1)) continue;

    // Group records by row field values
    const groupMap = new Map<string, { keys: (string | number)[]; values: number[][] }>();
    for (const record of sourceData.records) {
      const groupKey = rowFieldIndices.map((fi) => String(record[fi])).join("|");
      let group = groupMap.get(groupKey);
      if (!group) {
        group = {
          keys: rowFieldIndices.map((fi) => record[fi]),
          values: dataFieldIndices.map(() => []),
        };
        groupMap.set(groupKey, group);
      }
      for (let di = 0; di < dataFieldIndices.length; di++) {
        const val = record[dataFieldIndices[di]];
        if (typeof val === "number") {
          group.values[di].push(val);
        }
      }
    }

    const endCol = startCol + rowFieldNames.length + dataFields.length - 1;
    minCol = Math.min(minCol, startCol);
    maxCol = Math.max(maxCol, endCol);

    // Helper to add cells to a row
    const addCells = (rowIdx: number, cells: string[]) => {
      let arr = rowCells.get(rowIdx);
      if (!arr) {
        arr = [];
        rowCells.set(rowIdx, arr);
      }
      arr.push(...cells);
      minRow = Math.min(minRow, rowIdx);
      maxRow = Math.max(maxRow, rowIdx);
    };

    // Header row
    const headerCells: string[] = [];
    for (const rfName of rowFieldNames) {
      const cellRef = colIndexToLetterCompiler(startCol + headerCells.length) + startRow;
      const strIdx = sharedStrings.register(rfName);
      headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
    }
    for (const df of dataFields) {
      const cellRef = colIndexToLetterCompiler(startCol + headerCells.length) + startRow;
      const subtotal = df.summarize ?? "sum";
      const dfName = df.name ?? `${subtotal === "sum" ? "Sum" : subtotal} of ${df.field}`;
      const strIdx = sharedStrings.register(dfName);
      headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
    }
    addCells(startRow, headerCells);

    // Data rows
    let currentRow = startRow + 1;
    for (const [, group] of groupMap) {
      const cells: string[] = [];
      for (let ri = 0; ri < group.keys.length; ri++) {
        const cellRef = colIndexToLetterCompiler(startCol + ri) + currentRow;
        const strIdx = sharedStrings.register(String(group.keys[ri]));
        cells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
      }
      for (let di = 0; di < dataFields.length; di++) {
        const colOffset = rowFieldNames.length + di;
        const cellRef = colIndexToLetterCompiler(startCol + colOffset) + currentRow;
        const subtotal = dataFields[di].summarize ?? "sum";
        const result = aggregate(group.values[di], subtotal);
        cells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
      }
      addCells(currentRow, cells);
      currentRow++;
    }

    // Grand total row
    const gtCells: string[] = [];
    const gtStrIdx = sharedStrings.register("Grand Total");
    gtCells.push(
      `<c r="${colIndexToLetterCompiler(startCol)}${currentRow}" t="s"><v>${gtStrIdx}</v></c>`,
    );
    for (let di = 0; di < dataFields.length; di++) {
      const colOffset = rowFieldNames.length + di;
      const cellRef = colIndexToLetterCompiler(startCol + colOffset) + currentRow;
      const subtotal = dataFields[di].summarize ?? "sum";
      const allValues = sourceData.records
        .map((r) => r[dataFieldIndices[di]])
        .filter((v): v is number => typeof v === "number");
      const result = aggregate(allValues, subtotal);
      gtCells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
    }
    addCells(currentRow, gtCells);
  }

  if (rowCells.size === 0) return { sheetData: "", dimensionRef: "" };

  // Build sheetData — rows must be sorted by row number
  const parts: string[] = ["<sheetData>"];
  const sortedRows = [...rowCells.entries()].sort((a, b) => a[0] - b[0]);
  for (const [rowIdx, cells] of sortedRows) {
    parts.push(`<row r="${rowIdx}" x14ac:dyDescent="0.25">`);
    parts.push(...cells);
    parts.push("</row>");
  }
  parts.push("</sheetData>");

  const dimStartCol = colIndexToLetterCompiler(minCol === Infinity ? 0 : minCol);
  const dimStartRow = minRow === Infinity ? 1 : minRow;
  const dimensionRef = `${dimStartCol}${dimStartRow}:${colIndexToLetterCompiler(maxCol)}${maxRow}`;
  return { sheetData: parts.join(""), dimensionRef };
}

function colIndexToLetterCompiler(col: number): string {
  let result = "";
  let n = col + 1; // 0-indexed to 1-indexed
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function columnToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}
