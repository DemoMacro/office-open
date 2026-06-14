/**
 * XLSX Compiler — compiles WorkbookOptions into a Zippable structure.
 *
 * Accepts pure JSON WorkbookOptions — no intermediate File class needed.
 * Uses XlsxWriteContext for shared state (strings, styles, media, charts).
 *
 * @module
 */

import {
  APP_PROPS_XML,
  Relationships,
  TargetModeType,
  buildCorePropertiesXmlString,
  compileMapping,
  type XmlifyedFile,
  type Zippable,
} from "@office-open/core";
import { chartSpaceDesc } from "@office-open/core/chart";
import type { CalcCell } from "@parts/calc-chain";
import { calcChainDesc } from "@parts/calc-chain";
import { chartsheetDesc } from "@parts/chartsheet";
import { commentsDesc, vmlNotesDesc } from "@parts/comments";
import type { ImageOptions, ChartAnchorOptions } from "@parts/drawing";
import { drawingDesc } from "@parts/drawing";
import { externalLinkDesc } from "@parts/external-link";
import type { WorkbookOptions } from "@parts/file";
import { aggregate, collectUniqueValues } from "@parts/pivot";
import type { PivotSourceData, PivotTableOptions } from "@parts/pivot";
import { pivotCacheDefDesc, pivotCacheRecordsDesc } from "@parts/pivot-cache";
import { pivotTableDesc } from "@parts/pivot-table";
import { sharedStringsDesc } from "@parts/shared-strings";
import type { SharedStrings } from "@parts/shared-strings";
import { stylesDesc } from "@parts/styles";
import { tableDesc } from "@parts/table";
import { createThemeXml } from "@parts/theme";
import type { TablePartReference, SheetDefinition } from "@parts/workbook";
import { workbookDesc, buildTablePartsXml, buildExternalReferencesXml } from "@parts/workbook";
import { buildWorksheetXml, type WorksheetContext } from "@parts/worksheet";
import type { RowOptions, WorksheetOptions } from "@parts/worksheet";

import { XlsxWriteContext } from "./context";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Compile workbook options into a Zippable structure.
 */
export function compileWorkbook(
  options: WorkbookOptions,
  overrides: XmlifyedFile[] = [],
  mediaLevel: number = 0,
): Zippable {
  const ctx = new XlsxWriteContext();
  const mapping: Record<string, { data: string; path: string }> = {};

  const worksheetConfigs = options.worksheets ?? [];
  const chartsheetConfigs = options.chartsheets ?? [];

  // Core properties
  mapping["Properties"] = {
    data: XML_DECL + buildCorePropertiesXmlString(options),
    path: "docProps/core.xml",
  };

  // App properties
  mapping["AppProperties"] = {
    data: XML_DECL + APP_PROPS_XML,
    path: "docProps/app.xml",
  };

  // File-level relationships (_rels/.rels)
  const fileRels = buildFileRelationships();
  mapping["FileRelationships"] = {
    data: XML_DECL + fileRels.serialize(),
    path: "_rels/.rels",
  };

  // Register base content types
  for (let i = 0; i < worksheetConfigs.length; i++) {
    ctx.contentTypes.addWorksheet(i + 1);
  }
  ctx.contentTypes.addStyles();
  ctx.contentTypes.addSharedStrings();
  ctx.contentTypes.addTheme();

  // Register predefined DXFs before worksheets use styles
  for (const dxf of options.dxfs ?? []) {
    ctx.registerDxf(dxf);
  }

  // Re-apply parsed style sections so a declarative round-trip preserves
  // them (colors, table/cell styles, extensions) alongside DXFs.
  if (options.colors) ctx.styles.setColors(options.colors);
  if (options.tableStyles) ctx.styles.setTableStyles(options.tableStyles);
  if (options.cellStyles) ctx.styles.setCustomCellStyles(options.cellStyles);
  if (options.styleExtensions) ctx.styles.setExtensions(options.styleExtensions);

  // Build workbook relationships
  buildWorkbookRelationships(ctx.workbookRels, worksheetConfigs.length, chartsheetConfigs.length);

  // Build sheet definitions for workbook XML
  const sheets: SheetDefinition[] = [];
  let sheetId = 1;
  let rId = 1;
  for (const ws of worksheetConfigs) {
    sheets.push({
      name: ws.name ?? `Sheet${sheetId}`,
      sheetId: sheetId++,
      rId: `rId${rId++}`,
    });
  }
  for (const cs of chartsheetConfigs) {
    sheets.push({
      name: cs.name ?? `Chart${sheetId}`,
      sheetId: sheetId++,
      rId: `rId${rId++}`,
    });
  }

  // Worksheets — formatted BEFORE SharedStrings so strings are registered
  let globalMediaIdx = 0;
  let globalChartIdx = 0;
  let globalPivotIdx = 0;
  let globalPivotCacheIdx = 0;
  let globalTableIdx = 0;
  const pivotCacheDataMap = new Map<string, { cacheId: number; cacheIdx: number }>();
  const calcCells: CalcCell[] = [];
  const allTableParts: TablePartReference[] = [];

  const wsContext: WorksheetContext = { sharedStrings: ctx.sharedStrings, styles: ctx.styles };

  for (let i = 0; i < worksheetConfigs.length; i++) {
    const wsOpts = worksheetConfigs[i];
    const imgOpts = wsOpts.images ?? [];
    const chartOpts = wsOpts.charts ?? [];
    const hlOpts = wsOpts.hyperlinks ?? [];
    const sheetName = wsOpts.name ?? `Sheet${i + 1}`;

    // Worksheet uses buildWorksheetXml fast path (zero-allocation string concat)
    let sheetXml = buildWorksheetXml(wsOpts, wsContext);

    // Collect formula cells for calcChain
    const sheetIdx = i + 1;
    const wsRows = wsOpts.rows ?? [];
    for (let ri = 0; ri < wsRows.length; ri++) {
      const rowOpts = wsRows[ri];
      const rowNumber = rowOpts.rowNumber ?? ri + 1;
      if (!rowOpts.cells) continue;
      for (let ci = 0; ci < rowOpts.cells.length; ci++) {
        const cell = rowOpts.cells[ci];
        if (!cell.formula) continue;
        const ref = cell.reference ?? columnToLetter(ci + 1) + rowNumber;
        calcCells.push({
          reference: ref,
          sheetIndex: sheetIdx,
          array: cell.formula.type === "array",
        });
      }
    }

    const hasMedia = imgOpts.length > 0 || chartOpts.length > 0;
    const hasExternalHyperlinks = hlOpts.some((h) => h.target.type === "external");
    const commentOpts = wsOpts.comments ?? [];
    const hasComments = commentOpts.length > 0;
    const pivotOpts = wsOpts.pivotTables ?? [];
    const hasPivots = pivotOpts.length > 0;
    const tableOpts = wsOpts.tables ?? [];
    const hasTables = tableOpts.length > 0;
    const bgImg = wsOpts.backgroundImage;

    // Worksheet-level relationships
    let wsRels: Relationships | undefined;
    let nextRid = 0;

    if (hasMedia || hasExternalHyperlinks || hasComments || hasPivots || hasTables || bgImg) {
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
        ctx.media.addImage(mediaKey, {
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
        ctx.charts.addChart(chartKey, {
          key: chartKey,
          chartSpaceXml: chartSpaceDesc.stringify(chart, ctx) ?? "",
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

      // Generate drawing XML (via descriptor)
      const drawingXml = drawingDesc.stringify(
        { images: drawingImages, charts: drawingCharts },
        ctx,
      );
      const drawingIdx = i + 1;
      mapping[`Drawing${i}`] = {
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${drawingXml}`,
        path: `xl/drawings/drawing${drawingIdx}.xml`,
      };

      // Drawing relationships
      mapping[`DrawingRels${i}`] = {
        data: XML_DECL + drawingRels.serialize(),
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

      ctx.contentTypes.addDrawing(drawingIdx);
    }

    // Comments
    if (hasComments) {
      const commentsIdx = i + 1;

      // Comments XML (via descriptor)
      const commentsXml = commentsDesc.stringify({ comments: commentOpts }, ctx);
      mapping[`Comments${i}`] = {
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${commentsXml}`,
        path: `xl/comments${commentsIdx}.xml`,
      };

      // VML drawing (via descriptor)
      const vmlXml = vmlNotesDesc.stringify({ comments: commentOpts }, ctx);
      mapping[`VmlDrawing${i}`] = {
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${vmlXml}`,
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
        sheetXml.slice(0, -closingTag.length) + `<legacyDrawing r:id="rId${vmlRid}"/>` + closingTag;

      // Register content types
      ctx.contentTypes.addComments(commentsIdx);
      ctx.contentTypes.addVmlDrawing();
    }

    // Background picture
    if (bgImg) {
      const ext = bgImg.type === "jpg" ? "jpeg" : bgImg.type;
      const mediaKey = `bg_${i}`;
      const mediaIdx = globalMediaIdx + 1;
      ctx.media.addImage(mediaKey, {
        fileName: `image${mediaIdx}.${ext}`,
        type: ext,
        data: bgImg.data,
        width: 0,
        height: 0,
      });
      globalMediaIdx++;
      const bgRid = ++nextRid;
      wsRels!.addRelationship(
        bgRid,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        `../media/image${mediaIdx}.${ext}`,
      );
      sheetXml = sheetXml.replace("<!--BACKGROUND_PICTURE-->", `<picture r:id="rId${bgRid}"/>`);
    }

    // Pivot tables
    if (hasPivots) {
      for (const pt of pivotOpts) {
        globalPivotIdx++;
        const pivotIdx = globalPivotIdx;

        // Extract source data from source sheet
        const sourceSheet = pt.sourceSheet ?? sheetName;
        const sourceWsIdx = worksheetConfigs.findIndex(
          (ws) => (ws.name ?? `Sheet${worksheetConfigs.indexOf(ws) + 1}`) === sourceSheet,
        );
        if (sourceWsIdx === -1) continue;

        const sourceRows = worksheetConfigs[sourceWsIdx].rows ?? [];
        const sourceData = extractPivotSourceData(sourceRows, pt.source);

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
          cacheId = cacheIdx;
          pivotCacheDataMap.set(cacheKey, { cacheId, cacheIdx });

          // Generate pivotCacheDefinition
          const cacheDefRels = new Relationships();
          cacheDefRels.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
            "pivotCacheRecords1.xml",
          );

          const cacheDefXml =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            pivotCacheDefDesc.stringify(
              {
                sourceRef: pt.source.split(":")[0] ? pt.source : "A1",
                sourceSheet,
                sourceData,
                recordsRid: "rId1",
              },
              ctx,
            );

          mapping[`PivotCacheDef${cacheIdx}`] = {
            data: cacheDefXml,
            path: `xl/pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
          };
          mapping[`PivotCacheDefRels${cacheIdx}`] = {
            data: XML_DECL + cacheDefRels.serialize(),
            path: `xl/pivotCache/_rels/pivotCacheDefinition${cacheIdx}.xml.rels`,
          };

          // Generate pivotCacheRecords
          const cacheRecordsXml =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            pivotCacheRecordsDesc.stringify({ sourceData }, ctx);
          mapping[`PivotCacheRecords${cacheIdx}`] = {
            data: cacheRecordsXml,
            path: `xl/pivotCache/pivotCacheRecords${cacheIdx}.xml`,
          };

          // Content types
          ctx.contentTypes.addPivotCacheDefinition(cacheIdx);
          ctx.contentTypes.addPivotCacheRecords(cacheIdx);

          // Register in workbook
          const wbPivotRid = ctx.workbookRels.relationshipCount + 1;
          ctx.workbookRels.addRelationship(
            wbPivotRid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
            `pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
          );
          ctx.pivotCacheRefs.push({ cacheId, rId: `rId${wbPivotRid}` });
        }

        // Generate pivotTable
        const pivotTableXml =
          XML_DECL + pivotTableDesc.stringify({ options: pt, sourceData, cacheId }, ctx);
        mapping[`PivotTable${pivotIdx}`] = {
          data: pivotTableXml,
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
          data: XML_DECL + ptRels.serialize(),
          path: `xl/pivotTables/_rels/pivotTable${pivotIdx}.xml.rels`,
        };

        // Worksheet rels → pivotTable
        const ptRid = ++nextRid;
        wsRels!.addRelationship(
          ptRid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
          `../pivotTables/pivotTable${pivotIdx}.xml`,
        );

        ctx.contentTypes.addPivotTable(pivotIdx);
      }
    }

    // Tables (list objects)
    const wsTableParts: TablePartReference[] = [];
    if (hasTables) {
      for (const tbl of tableOpts) {
        globalTableIdx++;
        const tableIdx = globalTableIdx;

        // Generate table XML
        const tableXmlStr =
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          tableDesc.stringify({ ...tbl, id: tbl.id ?? tableIdx }, ctx);
        mapping[`Table${tableIdx}`] = {
          data: tableXmlStr,
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

        ctx.contentTypes.addTable(tableIdx);
      }
    }

    // Pre-render pivot table data into sheetData
    if (hasPivots) {
      const rendered = renderPivotSheetData(
        pivotOpts,
        worksheetConfigs,
        ctx.sharedStrings,
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
      const tablePartsXml = buildTablePartsXml(wsTableParts);
      const closingTag = "</worksheet>";
      sheetXml = sheetXml.slice(0, -closingTag.length) + tablePartsXml + closingTag;
    }

    // Write worksheet rels if needed
    if (wsRels) {
      mapping[`WorksheetRels${i}`] = {
        data: XML_DECL + wsRels.serialize(),
        path: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
      };
    }

    mapping[`Worksheet${i}`] = {
      data: sheetXml,
      path: `xl/worksheets/sheet${i + 1}.xml`,
    };
  }

  // Chartsheets — chart-only sheets
  for (let i = 0; i < chartsheetConfigs.length; i++) {
    const csOpts = chartsheetConfigs[i];

    // Register chart in the charts collection
    const chartDef = csOpts.chart;
    const csChartGlobalIdx = ctx.charts.array.length;
    const csChartKey = `cs_chart_${csChartGlobalIdx}`;
    ctx.charts.addChart(csChartKey, {
      key: csChartKey,
      chartSpaceXml:
        chartSpaceDesc.stringify(
          {
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
          },
          ctx,
        ) ?? "",
    });

    // Chartsheet relationships: drawing (required)
    const csRels = new Relationships();
    const csDrawingIdx = i + 1;
    csRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
      `../drawings/drawing${csDrawingIdx}.xml`,
    );

    // Drawing rels: chart reference
    const csDrawingRels = new Relationships();
    csDrawingRels.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      `../charts/chart${csChartGlobalIdx + 1}.xml`,
    );

    // Minimal drawing XML with chart anchor
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
      data: XML_DECL + csDrawingRels.serialize(),
      path: `xl/drawings/_rels/drawing${csDrawingIdx}.xml.rels`,
    };
    ctx.contentTypes.addDrawing(csDrawingIdx);

    mapping[`ChartsheetRels${i}`] = {
      data: XML_DECL + csRels.serialize(),
      path: `xl/chartsheets/_rels/sheet${i + 1}.xml.rels`,
    };
    mapping[`Chartsheet${i}`] = {
      data: XML_DECL + chartsheetDesc.stringify({ ...csOpts, drawingRId: "rId1" }, ctx),
      path: `xl/chartsheets/sheet${i + 1}.xml`,
    };
    ctx.contentTypes.addChartsheet(i + 1);
  }

  // Workbook XML (via descriptor)
  let wbXml =
    workbookDesc.stringify(
      {
        sheets,
        pivotCaches: ctx.pivotCacheRefs,
        protection: options.workbookProtection,
        customViews: options.customWorkbookViews,
        fileRecoveryPr: options.fileRecoveryPr,
        functionGroups: options.functionGroups,
        webPublishing: options.webPublishing,
        fileSharing: options.fileSharing,
        volTypes: options.volTypes,
        webPublishObjects: options.webPublishObjects,
      },
      ctx,
    ) ?? "";

  // External links — generate XML files and inject externalReferences into workbook
  const extLinks = options.externalLinks ?? [];
  if (extLinks.length > 0) {
    const extRefs: { rId: string }[] = [];
    for (let ei = 0; ei < extLinks.length; ei++) {
      const elIdx = ei + 1;
      const elRid = ctx.workbookRels.relationshipCount + 1;
      ctx.workbookRels.addRelationship(
        elRid,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
        `externalLinks/externalLink${elIdx}.xml`,
      );

      // Create the rels file for this external link
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
          data: XML_DECL + elRels.serialize(),
          path: `xl/externalLinks/_rels/externalLink${elIdx}.xml.rels`,
        };
      }

      // Generate the external link XML
      mapping[`ExternalLink${elIdx}`] = {
        data: XML_DECL + externalLinkDesc.stringify({ ...elOpts, bookRId }, ctx),
        path: `xl/externalLinks/externalLink${elIdx}.xml`,
      };

      extRefs.push({ rId: `rId${elRid}` });
      ctx.contentTypes.addExternalLink(elIdx);
    }

    // Inject externalReferences into workbook XML
    const extRefsXml = buildExternalReferencesXml(extRefs);
    wbXml = wbXml.replace("<!--EXTERNAL_REFS-->", extRefsXml);
  } else {
    wbXml = wbXml.replace("<!--EXTERNAL_REFS-->", "");
  }

  mapping["Workbook"] = {
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${wbXml}`,
    path: "xl/workbook.xml",
  };

  // Workbook relationships
  mapping["WorkbookRelationships"] = {
    data: XML_DECL + ctx.workbookRels.serialize(),
    path: "xl/_rels/workbook.xml.rels",
  };

  // Shared Strings — AFTER worksheets so all strings are collected
  if (ctx.sharedStrings.count > 0) {
    const ssXml = sharedStringsDesc.stringify(ctx.sharedStrings.toDescriptorOptions(), ctx);
    mapping["SharedStrings"] = {
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${ssXml}`,
      path: "xl/sharedStrings.xml",
    };
  }

  // Styles (via descriptor — delegates to Styles.toXml internally)
  const stylesXml = stylesDesc.stringify({ styles: ctx.styles }, ctx);
  mapping["Styles"] = {
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${stylesXml}`,
    path: "xl/styles.xml",
  };

  // Theme
  mapping["Theme"] = {
    data: XML_DECL + createThemeXml(),
    path: "xl/theme/theme1.xml",
  };

  // Charts — AFTER worksheets so charts are registered
  for (let i = 0; i < ctx.charts.array.length; i++) {
    const chartData = ctx.charts.array[i];
    mapping[`Chart${i}`] = {
      data: XML_DECL + chartData.chartSpaceXml,
      path: `xl/charts/chart${i + 1}.xml`,
    };
    ctx.contentTypes.addChart(i + 1);
  }

  // Calculation chain — auto-generated from formula cells
  if (calcCells.length > 0) {
    mapping["CalcChain"] = {
      data: calcChainDesc.stringify({ cells: calcCells }, ctx) ?? "",
      path: "xl/calcChain.xml",
    };
    ctx.contentTypes.addCalcChain();
    const calcChainRid = ctx.workbookRels.relationshipCount + 1;
    ctx.workbookRels.addRelationship(
      calcChainRid,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
      "calcChain.xml",
    );
  }

  // Register image content types
  const imageExts = new Set<string>();
  for (const img of ctx.media.array) {
    const ext = img.fileName.endsWith(".png") ? "png" : "jpeg";
    if (!imageExts.has(ext)) {
      imageExts.add(ext);
      ctx.contentTypes.addImageType(ext as "png" | "jpeg");
    }
  }

  // Content Types — must be last
  mapping["ContentTypes"] = {
    data: XML_DECL + ctx.contentTypes.serialize(),
    path: "[Content_Types].xml",
  };

  // Convert mapping to Zippable
  const mediaFiles: Array<{ data: Uint8Array; path: string }> = [];
  for (const img of ctx.media.array) {
    mediaFiles.push({ data: img.data, path: `xl/media/${img.fileName}` });
  }

  return compileMapping(mapping, overrides, mediaFiles, mediaLevel);
}

// ── Pure helper functions ──

function buildFileRelationships(): Relationships {
  const rels = new Relationships();
  rels.addRelationship(
    1,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "xl/workbook.xml",
  );
  rels.addRelationship(
    2,
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "docProps/core.xml",
  );
  rels.addRelationship(
    3,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "docProps/app.xml",
  );
  return rels;
}

function buildWorkbookRelationships(rels: Relationships, wsCount: number, csCount: number): void {
  let rid = 1;
  for (let i = 0; i < wsCount; i++) {
    rels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
      `worksheets/sheet${i + 1}.xml`,
    );
  }
  for (let i = 0; i < csCount; i++) {
    rels.addRelationship(
      rid++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
      `chartsheets/sheet${i + 1}.xml`,
    );
  }
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "styles.xml",
  );
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "theme/theme1.xml",
  );
  rels.addRelationship(
    rid++,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    "sharedStrings.xml",
  );
}

function extractPivotSourceData(rows: RowOptions[], sourceRef: string): PivotSourceData {
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
      const hv = headerRow.cells[c]?.value;
      fieldNames.push(
        typeof hv === "string"
          ? hv
          : typeof hv === "number" || typeof hv === "boolean"
            ? String(hv)
            : `Col${c}`,
      );
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
        record.push(typeof val === "string" ? val : typeof val === "boolean" ? String(val) : "");
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

function renderPivotSheetData(
  pivotOpts: PivotTableOptions[],
  worksheetConfigs: WorksheetOptions[],
  sharedStrings: SharedStrings,
  currentSheetName: string,
): { sheetData: string; dimensionRef: string } {
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

    const sourceRows = worksheetConfigs[sourceWsIdx]?.rows ?? [];
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
          keys: rowFieldIndices.map((fi) => {
            const v = record[fi];
            return typeof v === "string" || typeof v === "number" ? v : String(v ?? "");
          }),
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

    // Column field info for cross-tab layout
    const colFieldNames = pt.columns ?? [];
    const colFieldIndices = colFieldNames.map((n) => fields.indexOf(n));

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

    if (colFieldIndices.length > 0 && !colFieldIndices.some((idx) => idx === -1)) {
      // --- Cross-tab layout (with column fields) ---
      // Unique column values for the first column field
      const colUniqueVals = collectUniqueValues(sourceData.records, colFieldIndices[0]).map((v) =>
        typeof v === "string" || typeof v === "number" ? String(v) : String(v ?? ""),
      );

      // Build cross-tab map: rowKey → colKey → aggregated values per data field
      const crossTabMap = new Map<
        string,
        { rowKeys: (string | number)[]; colData: Map<string, number[][]>; rowTotals: number[][] }
      >();
      for (const record of sourceData.records) {
        const rowKey = rowFieldIndices.map((fi) => String(record[fi])).join("|");
        const colKey = colFieldIndices.map((fi) => String(record[fi])).join("|");
        let entry = crossTabMap.get(rowKey);
        if (!entry) {
          entry = {
            rowKeys: rowFieldIndices.map((fi) => {
              const v = record[fi];
              return typeof v === "string" || typeof v === "number" ? v : String(v ?? "");
            }),
            colData: new Map(),
            rowTotals: dataFieldIndices.map(() => []),
          };
          crossTabMap.set(rowKey, entry);
        }
        let colValues = entry.colData.get(colKey);
        if (!colValues) {
          colValues = dataFieldIndices.map(() => []);
          entry.colData.set(colKey, colValues);
        }
        for (let di = 0; di < dataFieldIndices.length; di++) {
          const val = record[dataFieldIndices[di]];
          if (typeof val === "number") {
            colValues[di].push(val);
            entry.rowTotals[di].push(val);
          }
        }
      }

      // Column count: row fields + column unique values + 1 (grand total column)
      const numColVals = colUniqueVals.length;
      const totalCols = rowFieldNames.length + numColVals + 1;
      const endCol = startCol + totalCols - 1;
      minCol = Math.min(minCol, startCol);
      maxCol = Math.max(maxCol, endCol);

      // Header row: [rowFieldName(s), ...colUniqueVals, dataFieldName or "Grand Total"]
      const headerCells: string[] = [];
      for (const rfName of rowFieldNames) {
        const cellRef = colIndexToLetter(startCol + headerCells.length) + startRow;
        const strIdx = sharedStrings.register(rfName);
        headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
      }
      for (const cv of colUniqueVals) {
        const cellRef = colIndexToLetter(startCol + headerCells.length) + startRow;
        const strIdx = sharedStrings.register(cv);
        headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
      }
      // Last header cell: data field name (e.g., "Total Revenue")
      {
        const cellRef = colIndexToLetter(startCol + headerCells.length) + startRow;
        const df0 = dataFields[0];
        const subtotal = df0.summarize ?? "sum";
        const dfName = df0.name ?? `${subtotal === "sum" ? "Sum" : subtotal} of ${df0.field}`;
        const strIdx = sharedStrings.register(dfName);
        headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
      }
      addCells(startRow, headerCells);

      // Data rows
      let currentRow = startRow + 1;
      for (const [, entry] of crossTabMap) {
        const cells: string[] = [];
        // Row label
        for (let ri = 0; ri < entry.rowKeys.length; ri++) {
          const cellRef = colIndexToLetter(startCol + ri) + currentRow;
          const strIdx = sharedStrings.register(String(entry.rowKeys[ri]));
          cells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
        }
        // Column values for each unique column value
        for (let ci = 0; ci < numColVals; ci++) {
          const colKey = colUniqueVals[ci];
          const colValues = entry.colData.get(colKey);
          const colOffset = rowFieldNames.length + ci;
          const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
          const subtotal = dataFields[0].summarize ?? "sum";
          const result = colValues ? aggregate(colValues[0], subtotal) : 0;
          cells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
        }
        // Row total (last column)
        {
          const colOffset = rowFieldNames.length + numColVals;
          const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
          const subtotal = dataFields[0].summarize ?? "sum";
          const result = aggregate(entry.rowTotals[0], subtotal);
          cells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
        }
        addCells(currentRow, cells);
        currentRow++;
      }

      // Grand total row
      const gtCells: string[] = [];
      const gtStrIdx = sharedStrings.register("Grand Total");
      gtCells.push(
        `<c r="${colIndexToLetter(startCol)}${currentRow}" t="s"><v>${gtStrIdx}</v></c>`,
      );
      for (let ci = 0; ci < numColVals; ci++) {
        const colKey = colUniqueVals[ci];
        const subtotal = dataFields[0].summarize ?? "sum";
        const colAllValues: number[] = [];
        for (const record of sourceData.records) {
          const recColKey = colFieldIndices.map((fi) => String(record[fi])).join("|");
          if (recColKey === colKey) {
            const val = record[dataFieldIndices[0]];
            if (typeof val === "number") colAllValues.push(val);
          }
        }
        const colOffset = rowFieldNames.length + ci;
        const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
        const result = aggregate(colAllValues, subtotal);
        gtCells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
      }
      // Grand total (bottom-right)
      {
        const subtotal = dataFields[0].summarize ?? "sum";
        const allValues = sourceData.records
          .map((r) => r[dataFieldIndices[0]])
          .filter((v): v is number => typeof v === "number");
        const colOffset = rowFieldNames.length + numColVals;
        const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
        const result = aggregate(allValues, subtotal);
        gtCells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
      }
      addCells(currentRow, gtCells);
    } else {
      // --- Simple layout (no column fields) ---
      const endCol = startCol + rowFieldNames.length + dataFields.length - 1;
      minCol = Math.min(minCol, startCol);
      maxCol = Math.max(maxCol, endCol);

      // Header row
      const headerCells: string[] = [];
      for (const rfName of rowFieldNames) {
        const cellRef = colIndexToLetter(startCol + headerCells.length) + startRow;
        const strIdx = sharedStrings.register(rfName);
        headerCells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
      }
      for (const df of dataFields) {
        const cellRef = colIndexToLetter(startCol + headerCells.length) + startRow;
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
          const cellRef = colIndexToLetter(startCol + ri) + currentRow;
          const strIdx = sharedStrings.register(String(group.keys[ri]));
          cells.push(`<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
        }
        for (let di = 0; di < dataFields.length; di++) {
          const colOffset = rowFieldNames.length + di;
          const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
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
        `<c r="${colIndexToLetter(startCol)}${currentRow}" t="s"><v>${gtStrIdx}</v></c>`,
      );
      for (let di = 0; di < dataFields.length; di++) {
        const colOffset = rowFieldNames.length + di;
        const cellRef = colIndexToLetter(startCol + colOffset) + currentRow;
        const subtotal = dataFields[di].summarize ?? "sum";
        const allValues = sourceData.records
          .map((r) => r[dataFieldIndices[di]])
          .filter((v): v is number => typeof v === "number");
        const result = aggregate(allValues, subtotal);
        gtCells.push(`<c r="${cellRef}"><v>${result}</v></c>`);
      }
      addCells(currentRow, gtCells);
    }
  }

  if (rowCells.size === 0) return { sheetData: "", dimensionRef: "" };

  // Build sheetData
  const parts: string[] = ["<sheetData>"];
  const sortedRows = [...rowCells.entries()].sort((a, b) => a[0] - b[0]);
  for (const [rowIdx, cells] of sortedRows) {
    parts.push(`<row r="${rowIdx}" x14ac:dyDescent="0.25">`);
    parts.push(...cells);
    parts.push("</row>");
  }
  parts.push("</sheetData>");

  const dimStartCol = colIndexToLetter(minCol === Infinity ? 0 : minCol);
  const dimStartRow = minRow === Infinity ? 1 : minRow;
  const dimensionRef = `${dimStartCol}${dimStartRow}:${colIndexToLetter(maxCol)}${maxRow}`;
  return { sheetData: parts.join(""), dimensionRef };
}

function colIndexToLetter(col: number): string {
  let result = "";
  let n = col + 1;
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
