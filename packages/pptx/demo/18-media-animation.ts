import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const videoData = new Uint8Array(fs.readFileSync(path.join(__dirname, "assets/test-video.mp4")));
const posterData = new Uint8Array(fs.readFileSync(path.join(__dirname, "assets/test-poster.png")));

const options: PresentationOptions = {
  slides: [
    // Slide 1: Video with play animation
    {
      children: [
        {
          video: {
            id: 100,
            x: 50,
            y: 100,
            width: 400,
            height: 300,
            data: videoData,
            type: "mp4",
            poster: posterData,
            posterType: "png",
          },
        },
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 50,
            textBody: { text: "Video Auto-Play" },
            fill: "4472C4",
          },
        },
      ],
      animations: [
        {
          shapeId: 100,
          options: {
            mediaType: "playVideo",
            trigger: "withPrevious",
          },
        },
      ],
    },

    // Slide 2: Property animation (opacity)
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 50,
            textBody: { text: "Opacity Property Animation" },
            fill: "70AD47",
          },
        },
        {
          shape: {
            id: 2,
            x: 150,
            y: 150,
            width: 200,
            height: 100,
            textBody: { text: "Fade Me" },
            fill: "FFC000",
          },
        },
      ],
      animations: [
        {
          shapeId: 2,
          options: {
            attributeName: "style.opacity",
            calcMode: "lin",
            valueType: "num",
            from: "1",
            to: "0.3",
            duration: 1000,
          },
        },
      ],
    },

    // Slide 3: Text-level animation with charRange
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 50,
            textBody: { text: "Text-Level Animation" },
            fill: "5B9BD5",
          },
        },
        {
          shape: {
            id: 2,
            textBody: {
              children: [
                {
                  children: [{ text: "Hello " }],
                },
                {
                  children: [{ text: "World" }],
                },
              ],
            },
            x: 100,
            y: 150,
            width: 300,
            height: 100,
            fill: "FFF2CC",
          },
        },
      ],
      animations: [
        {
          shapeId: 2,
          options: {
            type: "fade",
            charRange: [0, 5],
            duration: 500,
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
