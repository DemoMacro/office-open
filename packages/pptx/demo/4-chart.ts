import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Phase 4 Demo - Charts",
  creator: "Demo",
  slides: [
    // Column chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Column Chart" },
            fill: "4472C4",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "column",
            title: "Quarterly Sales",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "Product A", values: [100, 200, 300, 400] },
              { name: "Product B", values: [150, 180, 250, 350] },
            ],
            showLegend: true,
            style: 2,
          },
        },
      ],
    },
    // Pie chart
    {
      background: { fill: "F5F5F5" },
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Pie Chart" },
            fill: "ED7D31",
          },
        },
        {
          chart: {
            x: 100,
            y: 120,
            width: 500,
            height: 350,
            type: "pie",
            title: "Market Share",
            categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
            series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
          },
        },
      ],
    },
    // Line chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Line Chart" },
            fill: "70AD47",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "line",
            title: "Monthly Revenue",
            categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            series: [
              { name: "2024", values: [10, 15, 13, 20, 25, 30] },
              { name: "2025", values: [12, 18, 16, 22, 28, 35] },
            ],
            style: 2,
          },
        },
      ],
    },
    // Scatter chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Scatter Chart" },
            fill: "5B9BD5",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "scatter",
            title: "Height vs Weight",
            categories: ["Person A", "Person B", "Person C", "Person D"],
            series: [{ name: "Measurements", values: [160, 175, 155, 170] }],
          },
        },
      ],
    },
    // Area chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Area Chart" },
            fill: "FFC000",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "area",
            title: "Revenue Over Time",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "2024", values: [100, 200, 300, 400] },
              { name: "2025", values: [120, 220, 350, 500] },
            ],
          },
        },
      ],
    },
    // Bubble chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Bubble Chart" },
            fill: "7030A0",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "bubble",
            title: "Sales vs Profit (size = Revenue)",
            series: [
              {
                name: "Product A",
                xValues: [10, 20, 30, 40, 50],
                yValues: [5, 15, 10, 25, 30],
                bubbleSize: [100, 200, 150, 300, 250],
              },
              {
                name: "Product B",
                xValues: [15, 25, 35, 45, 55],
                yValues: [8, 12, 18, 22, 28],
                bubbleSize: [120, 180, 220, 280, 200],
              },
            ],
            showLegend: true,
            style: 2,
          },
        },
      ],
    },
    // Doughnut chart
    {
      background: { fill: "F5F5F5" },
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Doughnut Chart" },
            fill: "ED7D31",
          },
        },
        {
          chart: {
            x: 100,
            y: 120,
            width: 500,
            height: 350,
            type: "doughnut",
            title: "Revenue by Region",
            categories: ["North", "South", "East", "West"],
            series: [{ name: "Revenue", values: [35, 25, 22, 18] }],
            showLegend: true,
          },
        },
      ],
    },
    // Radar chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Radar Chart" },
            fill: "70AD47",
          },
        },
        {
          chart: {
            x: 100,
            y: 120,
            width: 500,
            height: 350,
            type: "radar",
            title: "Team Skills Comparison",
            categories: ["Speed", "Strength", "Agility", "Endurance", "Strategy"],
            series: [
              { name: "Team A", values: [85, 70, 90, 60, 80] },
              { name: "Team B", values: [70, 90, 65, 85, 75] },
            ],
            showLegend: true,
          },
        },
      ],
    },
    // Stock chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Stock Chart" },
            fill: "FF0000",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "stock",
            title: "ACME Corp — Daily Stock",
            categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
            series: [
              { name: "High", values: [105, 108, 112, 110, 115] },
              { name: "Low", values: [98, 103, 106, 105, 110] },
              { name: "Close", values: [103, 106, 110, 108, 113] },
            ],
            showLegend: true,
          },
        },
      ],
    },
    // Surface chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "Surface Chart" },
            fill: "5B9BD5",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "surface",
            title: "Temperature Surface",
            categories: ["Jan", "Feb", "Mar", "Apr", "May"],
            series: [
              { name: "North", values: [5, 8, 15, 20, 25] },
              { name: "South", values: [20, 22, 25, 30, 35] },
            ],
            showLegend: true,
          },
        },
      ],
    },
    // 3D Column chart
    {
      background: { fill: "F5F5F5" },
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "3D Column Chart" },
            fill: "4472C4",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "column",
            threeD: true,
            title: "Quarterly Sales (3D)",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            series: [
              { name: "Product A", values: [100, 200, 300, 400] },
              { name: "Product B", values: [150, 180, 250, 350] },
            ],
            showLegend: true,
            style: 2,
          },
        },
      ],
    },
    // 3D Pie chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "3D Pie Chart" },
            fill: "ED7D31",
          },
        },
        {
          chart: {
            x: 100,
            y: 120,
            width: 500,
            height: 350,
            type: "pie",
            threeD: true,
            title: "Market Share (3D)",
            categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
            series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
          },
        },
      ],
    },
    // 3D Line chart
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 60,
            textBody: { text: "3D Line Chart" },
            fill: "70AD47",
          },
        },
        {
          chart: {
            x: 50,
            y: 120,
            width: 600,
            height: 350,
            type: "line",
            threeD: true,
            title: "Monthly Revenue (3D)",
            categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            series: [
              { name: "2024", values: [10, 15, 13, 20, 25, 30] },
              { name: "2025", values: [12, 18, 16, 22, 28, 35] },
            ],
            style: 2,
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
