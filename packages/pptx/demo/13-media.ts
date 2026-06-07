import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

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
            x: 50,
            y: 30,
            width: 500,
            height: 50,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Video Embedding Demo",
                      fontSize: 32,
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
            x: 50,
            y: 100,
            width: 480,
            height: 270,
            data: videoData,
            type: "mp4",
            name: "Big Buck Bunny",
            poster: posterData,
            posterType: "png",
          },
        },
        {
          shape: {
            x: 50,
            y: 390,
            width: 500,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Video: Big Buck Bunny (360p, 10s, ~1MB MP4)",
                      fontSize: 14,
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

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
