import { Formatter } from "@export/formatter";
import type { IContext } from "@file/xml-components";
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
            const tree = new Formatter().format(documentBackground);
            expect(tree).to.deep.equal({
                "w:background": {
                    _attr: {},
                },
            });
        });

        it("should create a DocumentBackground with color", () => {
            const documentBackground = new DocumentBackground({ color: "ffff00" });
            const tree = new Formatter().format(documentBackground);
            expect(tree).to.deep.equal({
                "w:background": {
                    _attr: {
                        "w:color": "ffff00",
                    },
                },
            });
        });

        it("should create a DocumentBackground with all theme options", () => {
            const documentBackground = new DocumentBackground({
                color: "ffff00",
                themeColor: "test",
                themeShade: "0A",
                themeTint: "0B",
            });
            const tree = new Formatter().format(documentBackground);
            expect(tree).to.deep.equal({
                "w:background": {
                    _attr: {
                        "w:color": "ffff00",
                        "w:themeColor": "test",
                        "w:themeShade": "0A",
                        "w:themeTint": "0B",
                    },
                },
            });
        });
    });

    describe("#prepForXml() with image", () => {
        it("should register image with Media when image option is provided", () => {
            const addImage = vi.fn();
            const mockContext = {
                file: { Media: { addImage } },
                viewWrapper: { Relationships: { addRelationship: vi.fn() } },
                stack: [],
            } as unknown as IContext;

            const documentBackground = new DocumentBackground({
                image: {
                    data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
                    type: "png",
                },
            });

            documentBackground.prepForXml(mockContext);

            expect(addImage).toHaveBeenCalledWith(
                "abc123.png",
                expect.objectContaining({
                    type: "png",
                    fileName: "abc123.png",
                }),
            );
        });

        it("should render v:background with v:fill after prepForXml", () => {
            const addImage = vi.fn();
            const mockContext = {
                file: { Media: { addImage } },
                viewWrapper: { Relationships: { addRelationship: vi.fn() } },
                stack: [],
            } as unknown as IContext;

            const documentBackground = new DocumentBackground({
                image: {
                    data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
                    type: "png",
                },
            });

            const tree = new Formatter().format(documentBackground, mockContext);

            expect(tree).to.have.property("w:background");
            const bg = tree["w:background"];
            expect(bg).to.be.an("array");

            // Should contain v:background with v:fill
            const vBg = bg.find((el: Record<string, unknown>) => el["v:background"]);
            expect(vBg).to.exist;
            const vBgContent = vBg["v:background"];
            expect(vBgContent).to.be.an("array");
            const vFill = vBgContent.find((el: Record<string, unknown>) => el["v:fill"]);
            expect(vFill).to.exist;
        });

        it("should not add v:background when no image option is provided", () => {
            const documentBackground = new DocumentBackground({ color: "FF0000" });
            const tree = new Formatter().format(documentBackground);

            expect(tree).to.deep.equal({
                "w:background": {
                    _attr: {
                        "w:color": "FF0000",
                    },
                },
            });
        });
    });
});
