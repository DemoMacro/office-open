/**
 * XLSX Compiler — compiles a File object into a Zippable structure.
 *
 * @module
 */
import { Formatter } from "@export/formatter";
import { Drawing, type ImageOptions, type ChartAnchorOptions } from "@file/drawing/drawing";
import type { File } from "@file/file";
import type { BaseXmlComponent, Context } from "@file/xml-components";
import {
  ChartSpace,
  Relationships,
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

    // Workbook
    mapping["Workbook"] = {
      data: fmt(file.workbookXml),
      path: "xl/workbook.xml",
    };

    // Workbook relationships
    mapping["WorkbookRelationships"] = {
      data: fmt(file.workbookRelationships),
      path: "xl/_rels/workbook.xml.rels",
    };

    // Worksheets — formatted BEFORE SharedStrings so strings are registered
    const worksheets = file.worksheets;
    let globalMediaIdx = 0;
    let globalChartIdx = 0;

    for (let i = 0; i < worksheets.length; i++) {
      const ws = worksheets[i];
      const imgOpts = ws.imageOptions;
      const chartOpts = ws.charts;

      // Worksheet uses toXml() fast path (zero-allocation string concat)
      let sheetXml = fmt(ws);

      const hasMedia = imgOpts.length > 0 || chartOpts.length > 0;

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
        // Use slice instead of replace to avoid O(n) full-string scan.
        const closingTag = "</worksheet>";
        sheetXml =
          sheetXml.slice(0, -closingTag.length) + `<drawing r:id="rId${rid}"/>` + closingTag;

        // Worksheet needs its own rels for drawing reference
        const wsRels = new Relationships();
        wsRels.addRelationship(
          rid,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
          `../drawings/drawing${drawingIdx}.xml`,
        );
        mapping[`WorksheetRels${i}`] = {
          data: fmt(wsRels),
          path: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        };

        file.contentTypes.addDrawing(drawingIdx);
      }

      mapping[`Worksheet${i}`] = {
        data: sheetXml,
        path: `xl/worksheets/sheet${i + 1}.xml`,
      };
    }

    // Shared Strings — AFTER worksheets so all strings are collected
    const sharedStrings = file.sharedStrings;
    if (sharedStrings.count > 0) {
      mapping["SharedStrings"] = {
        data: fmt(sharedStrings),
        path: "xl/sharedStrings.xml",
      };
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
