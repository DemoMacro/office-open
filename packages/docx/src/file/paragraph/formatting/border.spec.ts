import { Formatter } from "@export/formatter";
import { BorderStyle } from "@file/border";
import { beforeEach, describe, expect, it } from "vite-plus/test";

import { Border, ThematicBreak } from "./border";

describe("Border", () => {
    describe("#constructor", () => {
        it("should create", () => {
            const border = new Border({
                between: {
                    color: "FF0000",
                    size: 10,
                    space: 9,
                    style: BorderStyle.WAVE,
                },
                bottom: {
                    color: "FF0000",
                    size: 4,
                    space: 3,
                    style: BorderStyle.WAVE,
                },
                left: {
                    color: "FF0000",
                    size: 6,
                    space: 5,
                    style: BorderStyle.WAVE,
                },
                right: {
                    color: "FF0000",
                    size: 8,
                    space: 7,
                    style: BorderStyle.WAVE,
                },
                top: {
                    color: "FF0000",
                    size: 2,
                    space: 1,
                    style: BorderStyle.WAVE,
                },
            });

            const tree = new Formatter().format(border);

            expect(tree).to.deep.equal({
                "w:pBdr": [
                    {
                        "w:top": {
                            _attr: {
                                "w:color": "FF0000",
                                "w:space": 1,
                                "w:sz": 2,
                                "w:val": "wave",
                            },
                        },
                    },
                    {
                        "w:left": {
                            _attr: {
                                "w:color": "FF0000",
                                "w:space": 5,
                                "w:sz": 6,
                                "w:val": "wave",
                            },
                        },
                    },
                    {
                        "w:bottom": {
                            _attr: {
                                "w:color": "FF0000",
                                "w:space": 3,
                                "w:sz": 4,
                                "w:val": "wave",
                            },
                        },
                    },
                    {
                        "w:right": {
                            _attr: {
                                "w:color": "FF0000",
                                "w:space": 7,
                                "w:sz": 8,
                                "w:val": "wave",
                            },
                        },
                    },
                    {
                        "w:between": {
                            _attr: {
                                "w:color": "FF0000",
                                "w:space": 9,
                                "w:sz": 10,
                                "w:val": "wave",
                            },
                        },
                    },
                ],
            });
        });

        it("should not add empty borders element if there are no borders defined", () => {
            const tb = new Border({});
            expect(() => new Formatter().format(tb)).to.throw();
        });
    });
});

describe("ThematicBreak", () => {
    let thematicBreak: ThematicBreak;

    beforeEach(() => {
        thematicBreak = new ThematicBreak();
    });

    describe("#constructor()", () => {
        it("should create a Thematic Break with correct border properties", () => {
            const tree = new Formatter().format(thematicBreak);
            expect(tree).to.deep.equal({
                "w:pBdr": [
                    {
                        "w:bottom": {
                            _attr: {
                                "w:color": "auto",
                                "w:space": 1,
                                "w:sz": 6,
                                "w:val": "single",
                            },
                        },
                    },
                ],
            });
        });
    });
});
