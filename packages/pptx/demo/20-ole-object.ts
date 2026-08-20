// OLE Object on a slide — the binary is registered as
// ppt/embeddings/oleObject1.bin and wired through the slide relationship.

import { mkdirSync, writeFileSync } from "node:fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Minimal well-formed OLE compound file (CFB v3, 512-byte sectors): header,
// one FAT sector, one directory sector with an empty root storage. Real OLE
// payloads (e.g. an actual Excel sheet) would replace this.
function makeOleStub(): Uint8Array {
  const SECTOR = 512;
  const data = new Uint8Array(SECTOR * 3); // header + FAT sector + directory sector
  const view = new DataView(data.buffer);
  const ENDOFCHAIN = 0xfffffffe;
  const FATSECT = 0xfffffffd;
  const FREESECT = 0xffffffff;

  // Header
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  for (const [i, b] of sig.entries()) data[i] = b;
  view.setUint16(24, 0x3e, true); // minor version
  view.setUint16(26, 3, true); // major version
  view.setUint16(28, 0xfffe, true); // byte order (little-endian)
  view.setUint16(30, 9, true); // sector shift = 512 bytes
  view.setUint16(32, 6, true); // mini sector shift = 64 bytes
  view.setUint32(44, 1, true); // number of FAT sectors
  view.setUint32(48, 1, true); // first directory sector (after the FAT sector)
  view.setUint32(60, ENDOFCHAIN, true); // first mini FAT sector
  view.setUint32(64, 0, true); // number of mini FAT sectors
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(72, 0, true); // number of DIFAT sectors
  view.setUint32(76, 0, true); // DIFAT[0] = FAT sector 0
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, FREESECT, true); // unused DIFAT slots

  // FAT sector (sector 0): FAT[0] marks itself, FAT[1] ends the directory chain
  view.setUint32(SECTOR, FATSECT, true);
  view.setUint32(SECTOR + 4, ENDOFCHAIN, true);
  for (let i = 2; i < SECTOR / 4; i++) view.setUint32(SECTOR + i * 4, FREESECT, true);

  // Directory sector (sector 1): root entry only, no children, no mini stream
  const dir = SECTOR * 2;
  const name = "Root Entry";
  for (let i = 0; i < name.length; i++) view.setUint16(dir + i * 2, name.charCodeAt(i), true);
  view.setUint32(dir + 64, (name.length + 1) * 2, true); // name length incl. terminator
  data[dir + 68] = 5; // object type = root storage
  data[dir + 69] = 1; // color = black
  view.setUint32(dir + 72, FREESECT, true); // left sibling
  view.setUint32(dir + 76, FREESECT, true); // right sibling
  view.setUint32(dir + 80, FREESECT, true); // child
  view.setUint32(dir + 120, ENDOFCHAIN, true); // start sector
  view.setUint32(dir + 124, 0, true); // stream size

  return data;
}

const options: PresentationOptions = {
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "15.9cm",
            height: "1.1cm",
            textBody: { text: "OLE Object Demo" },
            fill: "4472C4",
          },
        },
        {
          ole: {
            x: "2.6cm",
            y: "2.6cm",
            width: "10.6cm",
            height: "7.9cm",
            progId: "Excel.Sheet.12",
            name: "Embedded Worksheet",
            showAsIcon: false,
            embed: {
              data: makeOleStub(),
            },
            // Office refuses to open the presentation when an oleObj has no
            // picture, so the preview icon is part of the object itself.
            iconImage: {
              data: "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFklEQVR4nGNwKTpCEmIY1TCqYfhqAAB07noQI9onJgAAAABJRU5ErkJggg==",
              type: "png",
            },
          },
        },
      ],
    },
    {
      children: [
        {
          ole: {
            x: "2.6cm",
            y: "2.6cm",
            width: "10.6cm",
            height: "7.9cm",
            progId: "Excel.Sheet.12",
            name: "Linked Worksheet",
            // Linked OLE — no bytes in the package; the source URL becomes an
            // External oleObject relationship of this slide (p:oleObj + p:link).
            link: {
              url: "https://example.com/sales.xlsx",
              autoUpdate: true,
            },
            iconImage: {
              data: "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAFklEQVR4nGNwKTpCEmIY1TCqYfhqAAB07noQI9onJgAAAABJRU5ErkJggg==",
              type: "png",
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/20-ole-object.pptx", buffer);
