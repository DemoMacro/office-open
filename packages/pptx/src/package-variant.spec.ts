import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generatePresentation, generatePresentationSync } from "./generate";
import type { PresentationOptions } from "./shared/file";

const CFB_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

const PACKAGE_VARIANTS = [
  {
    packageVariant: "standard",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    mainContentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  },
  {
    packageVariant: "macro",
    mimeType: "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
    mainContentType: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  },
  {
    packageVariant: "template",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.template",
    mainContentType:
      "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
  },
  {
    packageVariant: "macroTemplate",
    mimeType: "application/vnd.ms-powerpoint.template.macroEnabled.12",
    mainContentType: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
  },
] as const;

const options = (): PresentationOptions => ({
  slides: [
    {
      children: [{ shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "A" } } }],
    },
  ],
  contentTypes: {
    defaults: [],
    overrides: [{ partName: "/ppt/presentation.xml", contentType: "stale/main" }],
  },
});

const contentTypes = (buffer: Uint8Array): string => {
  const entry = unzipSync(buffer)["[Content_Types].xml"];
  if (!entry) throw new Error("missing zip entry: [Content_Types].xml");
  return new TextDecoder().decode(entry);
};

describe("PPTX package variants", () => {
  it.each(PACKAGE_VARIANTS)(
    "declares $packageVariant content type and Blob MIME",
    ({ packageVariant, mimeType, mainContentType }) => {
      const buffer = generatePresentationSync(options(), {
        type: "uint8array",
        packageVariant,
      });
      expect(contentTypes(buffer)).toContain(
        `<Override PartName="/ppt/presentation.xml" ContentType="${mainContentType}"/>`,
      );

      const blob = generatePresentationSync(options(), { type: "blob", packageVariant });
      expect(blob.type).toBe(mimeType.toLowerCase());
    },
  );

  it.each(PACKAGE_VARIANTS)(
    "uses $packageVariant MIME for encrypted passthrough",
    async ({ packageVariant, mimeType }) => {
      const blob = await generatePresentation(
        { encrypted: { data: CFB_BYTES } },
        { type: "blob", packageVariant },
      );
      expect(blob.type).toBe(mimeType.toLowerCase());
    },
  );
});
