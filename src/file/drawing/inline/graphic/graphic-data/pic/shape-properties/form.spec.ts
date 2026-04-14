import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { Form } from "./form/form";

describe("Form", () => {
    describe("#constructor()", () => {
        it("should create", () => {
            const tree = new Formatter().format(
                new Form({
                    emus: {
                        x: 100,
                        y: 100,
                    },
                    pixels: {
                        x: 100,
                        y: 100,
                    },
                }),
            );

            expect(tree).to.deep.equal({
                "a:xfrm": [
                    {
                        _attr: {},
                    },
                    {
                        "a:off": {
                            _attr: {
                                x: 0,
                                y: 0,
                            },
                        },
                    },
                    {
                        "a:ext": {
                            _attr: {
                                cx: 100,
                                cy: 100,
                            },
                        },
                    },
                ],
            });
        });

        it("should create with flip", () => {
            const tree = new Formatter().format(
                new Form({
                    emus: {
                        x: 100,
                        y: 100,
                    },
                    flip: {
                        horizontal: true,
                        vertical: true,
                    },
                    pixels: {
                        x: 100,
                        y: 100,
                    },
                }),
            );

            expect(tree).to.deep.equal({
                "a:xfrm": [
                    {
                        _attr: {
                            flipH: true,
                            flipV: true,
                        },
                    },
                    {
                        "a:off": {
                            _attr: {
                                x: 0,
                                y: 0,
                            },
                        },
                    },
                    {
                        "a:ext": {
                            _attr: {
                                cx: 100,
                                cy: 100,
                            },
                        },
                    },
                ],
            });
        });
    });
});
