import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const videoPath = path.join(__dirname, "assets/test-video.mp4");
const videoData = new Uint8Array(fs.readFileSync(videoPath));
const posterPath = path.join(__dirname, "assets/test-poster.png");
const posterData = new Uint8Array(fs.readFileSync(posterPath));

const options: PresentationOptions = {
  title: "Video Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Video Embedding Demo",
                      size: 32,
                      bold: true,
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          video: {
            x: "1.3cm",
            y: "2.6cm",
            width: "12.7cm",
            height: "7.1cm",
            data: videoData,
            type: "mp4",
            name: "Big Buck Bunny",
            poster: posterData,
            posterType: "png",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "10.3cm",
            width: "13.2cm",
            height: "1.1cm",
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Video: Big Buck Bunny (360p, 10s, ~1MB MP4)",
                      size: 14,
                      fill: "666666",
                    },
                  ],
                },
              ],
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
