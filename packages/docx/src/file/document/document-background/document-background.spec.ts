import type { Context } from "@file/xml-components";
import * as convenienceFunctions from "@util/convenience-functions";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { DocumentBackground } from "./document-background";

describe("DocumentBackground", () => {
  beforeEach(() => {
    vi.spyOn(convenienceFunctions, "hashedId").mockReturnValue("abc123");
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe("#constructor()", () => {
    it("should create a DocumentBackground with no options", () => {
      const documentBackground = new DocumentBackground({});
      const xml = documentBackground.toXml({
        stack: [],
        file: null!,
        viewWrapper: null!,
      } as unknown as Context);
      expect(xml).to.equal("<w:background/>");
    });

    it("should create a DocumentBackground with color", () => {
      const documentBackground = new DocumentBackground({ color: "ffff00" });
      const xml = documentBackground.toXml({
        stack: [],
        file: null!,
        viewWrapper: null!,
      } as unknown as Context);
      expect(xml).to.contain('w:color="ffff00"');
    });

    it("should create a DocumentBackground with all theme options", () => {
      const documentBackground = new DocumentBackground({
        color: "ffff00",
        themeColor: "test",
        themeShade: "0A",
        themeTint: "0B",
      });
      const xml = documentBackground.toXml({
        stack: [],
        file: null!,
        viewWrapper: null!,
      } as unknown as Context);
      expect(xml).to.contain('w:color="ffff00"');
      expect(xml).to.contain('w:themeColor="test"');
      expect(xml).to.contain('w:themeShade="0A"');
      expect(xml).to.contain('w:themeTint="0B"');
    });
  });

  describe("#toXml() with image", () => {
    it("should register image with Media when image option is provided", () => {
      const addImage = vi.fn();
      const mockContext = {
        file: { media: { addImage } },
        viewWrapper: { relationships: { addRelationship: vi.fn() } },
        stack: [],
      } as unknown as Context;

      const documentBackground = new DocumentBackground({
        image: {
          data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
          type: "png",
        },
      });

      documentBackground.toXml(mockContext);

      expect(addImage).toHaveBeenCalledWith(
        "abc123.png",
        expect.objectContaining({
          type: "png",
          fileName: "abc123.png",
        }),
      );
    });

    it("should render v:background with v:fill", () => {
      const addImage = vi.fn();
      const mockContext = {
        file: { media: { addImage } },
        viewWrapper: { relationships: { addRelationship: vi.fn() } },
        stack: [],
      } as unknown as Context;

      const documentBackground = new DocumentBackground({
        image: {
          data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
          type: "png",
        },
      });

      const xml = documentBackground.toXml(mockContext);

      expect(xml).to.contain("v:background");
      expect(xml).to.contain("v:fill");
    });

    it("should not add v:background when no image option is provided", () => {
      const documentBackground = new DocumentBackground({ color: "FF0000" });
      const xml = documentBackground.toXml({
        stack: [],
        file: null!,
        viewWrapper: null!,
      } as unknown as Context);

      expect(xml).to.contain('w:color="FF0000"');
      expect(xml).not.to.contain("v:background");
      expect(xml).not.to.contain("v:fill");
    });
  });
});
