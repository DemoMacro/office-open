import { Formatter } from "@export/formatter";
import { HorizontalPositionAlign, VerticalPositionAlign } from "@file/shared";
import { describe, expect, it } from "vite-plus/test";

import { FrameAnchorType, createFrameProperties } from "./frame-properties";

describe("createFrameProperties", () => {
    it("should create", () => {
        const currentFrameProperties = createFrameProperties({
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
            height: 1000,
            position: {
                x: 1000,
                y: 3000,
            },
            type: "absolute",
            width: 4000,
        });

        const tree = new Formatter().format(currentFrameProperties);
        expect(tree).to.deep.equal({
            "w:framePr": {
                _attr: {
                    "w:h": 1000,
                    "w:hAnchor": "margin",
                    "w:vAnchor": "margin",
                    "w:w": 4000,
                    "w:x": 1000,
                    "w:y": 3000,
                },
            },
        });
    });

    it("should create with the space attribute", () => {
        const currentFrameProperties = createFrameProperties({
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
            height: 1000,
            position: {
                x: 1000,
                y: 3000,
            },
            space: {
                horizontal: 100,
                vertical: 200,
            },
            type: "absolute",
            width: 4000,
        });

        const tree = new Formatter().format(currentFrameProperties);
        expect(tree).to.deep.equal({
            "w:framePr": {
                _attr: {
                    "w:h": 1000,
                    "w:hAnchor": "margin",
                    "w:hSpace": 100,
                    "w:vAnchor": "margin",
                    "w:vSpace": 200,
                    "w:w": 4000,
                    "w:x": 1000,
                    "w:y": 3000,
                },
            },
        });
    });

    it("should create without x and y", () => {
        const currentFrameProperties = createFrameProperties({
            alignment: {
                x: HorizontalPositionAlign.CENTER,
                y: VerticalPositionAlign.TOP,
            },
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
            height: 1000,
            space: {
                horizontal: 100,
                vertical: 200,
            },
            type: "alignment",
            width: 4000,
        });

        const tree = new Formatter().format(currentFrameProperties);
        expect(tree).to.deep.equal({
            "w:framePr": {
                _attr: {
                    "w:h": 1000,
                    "w:hAnchor": "margin",
                    "w:hSpace": 100,
                    "w:vAnchor": "margin",
                    "w:vSpace": 200,
                    "w:w": 4000,
                    "w:xAlign": "center",
                    "w:yAlign": "top",
                },
            },
        });
    });

    it("should create without alignments", () => {
        const currentFrameProperties = createFrameProperties({
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
            height: 1000,
            position: {
                x: 1000,
                y: 3000,
            },
            space: {
                horizontal: 100,
                vertical: 200,
            },
            type: "absolute",
            width: 4000,
        });

        const tree = new Formatter().format(currentFrameProperties);
        expect(tree).to.deep.equal({
            "w:framePr": {
                _attr: {
                    "w:h": 1000,
                    "w:hAnchor": "margin",
                    "w:hSpace": 100,
                    "w:vAnchor": "margin",
                    "w:vSpace": 200,
                    "w:w": 4000,
                    "w:x": 1000,
                    "w:y": 3000,
                },
            },
        });
    });
});
