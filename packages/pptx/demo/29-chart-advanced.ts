import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Exercises the chart descriptor's advanced surface so the generated chart parts
// traverse dml-chart.xsd: full axes, series decorations, 3D view/layout, line
// extras, pie/doughnut/bubble/surface/ofPie scalars, multi-level categories, and
// chartSpace-level containers (color map override, protection, print settings).

const options: PresentationOptions = {
  title: "Phase 4 Demo - Chart Advanced Coverage",
  creator: "Demo",
  slides: [
    // Slide 1: full axis configuration (scaling, units, gridlines, ticks, title).
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Axes — scaling, units, gridlines, ticks, title" },
            fill: "4472C4",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "column",
            title: "Quarterly Sales",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "Product A", values: [100, 200, 300, 400] },
              { name: "Product B", values: [150, 180, 250, 350] },
            ],
            axes: [
              {
                kind: "category",
                id: 111111111,
                crossAxisId: 222222222,
                position: "bottom",
                title: "Quarter",
                auto: true,
                labelOffset: 100,
                majorTickMark: "out",
                minorTickMark: "none",
                tickLabelPosition: "nextTo",
                numberFormat: "General",
                crosses: "zero",
              },
              {
                kind: "value",
                id: 222222222,
                crossAxisId: 111111111,
                position: "left",
                title: "Revenue ($)",
                scaling: { min: 0, max: 500, orientation: "ascending" },
                majorUnit: 100,
                minorUnit: 20,
                crossBetween: "between",
                majorGridlines: true,
                majorTickMark: "out",
                tickLabelPosition: "nextTo",
                numberFormat: "$#,##0",
                displayUnits: { builtInUnit: "thousands", label: true },
                crosses: "zero",
              },
            ],
            pivotFormats: [{ index: 0, marker: { symbol: "square", size: 5 } }],
            legendPosition: "top",
            legendEntries: [{ index: 1, delete: true }],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 2: series decorations (marker, smooth, data points, trendline, error bars, labels).
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: {
              text: "Series — marker, smooth, data points, trendline, error bars, labels",
            },
            fill: "70AD47",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "line",
            title: "Monthly Revenue",
            categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            series: [
              {
                name: "2024",
                values: [10, 15, 13, 20, 25, 30],
                smooth: true,
                marker: { symbol: "circle", size: 7 },
                invertIfNegative: true,
                trendlines: [{ type: "linear", name: "Trend", dispEq: true, dispRSqr: true }],
                errorBars: { direction: "y", barType: "both", valueType: "fixedValue", value: 2 },
                dataLabels: { showVal: true, position: "top" },
                dataPoints: [{ index: 2, marker: { symbol: "diamond", size: 9 } }],
              },
              {
                name: "2025",
                values: [12, 18, 16, 22, 28, 35],
                marker: { symbol: "square", size: 6 },
              },
            ],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 3: 3D view, walls/floor, manual layout, gap depth, bar shape.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "3D view, walls, floor, manual layout, bar shape" },
            fill: "5B9BD5",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "column",
            threeD: true,
            title: "Quarterly Sales (3D)",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "Product A", values: [100, 200, 300, 400], shape: "cylinder" },
              { name: "Product B", values: [150, 180, 250, 350], shape: "cylinder" },
            ],
            view3D: { rotX: 30, rotY: 20, depthPercent: 150, perspective: 30, rAngAx: true },
            floor: { thickness: "10%" },
            sideWall: { thickness: "5%" },
            backWall: { thickness: "5%" },
            gapDepth: 150,
            plotAreaLayout: {
              layoutTarget: "inner",
              xMode: "edge",
              yMode: "edge",
              wMode: "edge",
              hMode: "edge",
              x: 0.1,
              y: 0.1,
              w: 0.8,
              h: 0.8,
            },
            showLegend: true,
          },
        },
      ],
    },
    // Slide 4: line extras (up/down bars, hi-low, drop lines) + data table on a second chart.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Up/down bars, hi-low lines, drop lines, data table" },
            fill: "FFC000",
          },
        },
        {
          chart: {
            x: "0.8cm",
            y: "1.9cm",
            width: "9cm",
            height: "8cm",
            type: "line",
            title: "Up/Down + Hi-Low + Drop",
            categories: ["Jan", "Feb", "Mar", "Apr", "May"],
            series: [
              { name: "High", values: [30, 35, 32, 40, 38] },
              { name: "Low", values: [10, 15, 13, 20, 18] },
            ],
            upDownBars: true,
            upDownBarsGapWidth: 150,
            highLowLines: true,
            dropLines: true,
            showLegend: true,
          },
        },
        {
          chart: {
            x: "10cm",
            y: "1.9cm",
            width: "8cm",
            height: "8cm",
            type: "column",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "Product A", values: [100, 200, 300, 400] },
              { name: "Product B", values: [150, 180, 250, 350] },
            ],
            gapWidth: 150,
            overlap: -20,
            dataTable: {
              showHorizontalBorder: true,
              showVerticalBorder: true,
              showOutline: true,
              showLegendKeys: true,
            },
          },
        },
      ],
    },
    // Slide 5: pie first slice + explosion, doughnut hole size.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Pie first slice + explosion, doughnut hole size" },
            fill: "ED7D31",
          },
        },
        {
          chart: {
            x: "0.8cm",
            y: "1.9cm",
            width: "8cm",
            height: "8cm",
            type: "pie",
            title: "Market Share",
            categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
            series: [
              {
                name: "Browser",
                values: [65, 18, 3, 5, 9],
                dataPoints: [
                  { index: 0, explosion: 25 },
                  { index: 1, explosion: 10 },
                ],
              },
            ],
            firstSliceAngle: 45,
          },
        },
        {
          chart: {
            x: "9.8cm",
            y: "1.9cm",
            width: "8cm",
            height: "8cm",
            type: "doughnut",
            title: "Revenue by Region",
            categories: ["North", "South", "East", "West"],
            series: [{ name: "Revenue", values: [35, 25, 22, 18] }],
            holeSize: 60,
          },
        },
      ],
    },
    // Slide 6: bubble scale, size represents, negative bubbles.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Bubble scale, size represents, negative bubbles" },
            fill: "7030A0",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "bubble",
            title: "Sales vs Profit",
            bubbleScale: 80,
            sizeRepresents: "area",
            showNegativeBubbles: true,
            series: [
              {
                name: "Product A",
                xValues: [10, 20, 30, 40, 50],
                yValues: [5, 15, 10, 25, 30],
                bubbleSize: [100, -50, 150, 300, 250],
              },
            ],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 7: surface wireframe.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Surface wireframe" },
            fill: "5B9BD5",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "surface",
            title: "Temperature Surface",
            wireframe: true,
            categories: ["Jan", "Feb", "Mar", "Apr", "May"],
            series: [
              { name: "North", values: [5, 8, 15, 20, 25] },
              { name: "South", values: [20, 22, 25, 30, 35] },
            ],
            bandFormats: [{ index: 0 }, { index: 1 }, { index: 2 }],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 8: bar-of-pie (ofPie) split type, position, second pie size.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Bar-of-pie: split type, position, second pie size" },
            fill: "ED7D31",
          },
        },
        {
          chart: {
            x: "2cm",
            y: "1.9cm",
            width: "14cm",
            height: "8.8cm",
            type: "ofPie",
            title: "Revenue Concentration",
            ofPieType: "bar",
            splitType: "position",
            splitPosition: 2,
            secondPieSize: 75,
            categories: ["A", "B", "C", "D", "E", "F", "G", "H"],
            series: [{ name: "Revenue", values: [40, 25, 15, 8, 5, 3, 2, 2] }],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 9: multi-level categories.
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: { text: "Multi-level categories" },
            fill: "70AD47",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.8cm",
            type: "column",
            title: "Hierarchical Categories",
            multiLevelCategories: [
              ["H1", "H1", "H2", "H2"],
              ["A", "B", "C", "D"],
            ],
            series: [{ name: "Series 1", values: [10, 20, 30, 40] }],
            showLegend: true,
          },
        },
      ],
    },
    // Slide 10: chartSpace-level containers (blanks behavior, color map override, protection, print settings).
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "16cm",
            height: "1.1cm",
            textBody: {
              text: "ChartSpace: blanks, color map override, protection, print settings",
            },
            fill: "4472C4",
          },
        },
        {
          chart: {
            x: "1.3cm",
            y: "1.9cm",
            width: "16cm",
            height: "8.5cm",
            type: "column",
            title: "Full ChartSpace Options",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [{ name: "Product A", values: [100, 200, 300, 400] }],
            plotVisOnly: false,
            displayBlanksAs: "span",
            showDataLabelsOverMax: true,
            colorMappingOverride: {
              background1: "dark1",
              text1: "light1",
            },
            protection: {
              chartObject: true,
              data: true,
              formatting: false,
              selection: true,
              userInterface: false,
            },
            printSettings: {
              headerFooter: {
                oddHeader: "Chart Report",
                oddFooter: "Page 1",
                differentFirst: true,
                alignWithMargins: true,
              },
              pageMargins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 },
              pageSetup: {
                orientation: "landscape",
                paperSize: 9,
                horizontalDpi: 300,
                verticalDpi: 300,
                copies: 1,
              },
            },
            showLegend: true,
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/29-chart-advanced.pptx", buffer);
