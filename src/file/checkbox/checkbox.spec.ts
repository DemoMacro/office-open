import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { CheckBox } from "./checkbox";

describe("CheckBox", () => {
    describe("#constructor()", () => {
        it("should create a CheckBox with proper root and default values (no alias, no custom state)", () => {
            const checkBox = new CheckBox();

            const tree = new Formatter().format(checkBox);

            expect(tree).to.deep.equal({
                "w:sdt": [
                    {
                        "w:sdtPr": [
                            {
                                "w14:checkbox": [
                                    {
                                        "w14:checked": {
                                            _attr: {
                                                "w14:val": "0",
                                            },
                                        },
                                    },
                                    {
                                        "w14:checkedState": {
                                            _attr: {
                                                "w14:font": "MS Gothic",
                                                "w14:val": "2612",
                                            },
                                        },
                                    },
                                    {
                                        "w14:uncheckedState": {
                                            _attr: {
                                                "w14:font": "MS Gothic",
                                                "w14:val": "2610",
                                            },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "w:sdtContent": [
                            {
                                "w:r": [
                                    {
                                        "w:sym": {
                                            _attr: {
                                                "w:char": "2610",
                                                "w:font": "MS Gothic",
                                            },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });

        it.each([
            ["2713", "Segoe UI Symbol", "2713", "Segoe UI Symbol"],
            [undefined, undefined, "2612", "MS Gothic"],
        ])(
            "should create a CheckBox with proper root and custom values",
            (inputChar, inputFont, actualChar, actualFont) => {
                const checkBox = new CheckBox({
                    alias: "Custom Checkbox",
                    checked: true,
                    checkedState: {
                        font: inputFont,
                        value: inputChar,
                    },
                    uncheckedState: {
                        font: "Segoe UI Symbol",
                        value: "2705",
                    },
                });

                const tree = new Formatter().format(checkBox);

                expect(tree).to.deep.equal({
                    "w:sdt": [
                        {
                            "w:sdtPr": [
                                {
                                    "w:alias": {
                                        _attr: {
                                            "w:val": "Custom Checkbox",
                                        },
                                    },
                                },
                                {
                                    "w14:checkbox": [
                                        {
                                            "w14:checked": {
                                                _attr: {
                                                    "w14:val": "1",
                                                },
                                            },
                                        },
                                        {
                                            "w14:checkedState": {
                                                _attr: {
                                                    "w14:font": actualFont,
                                                    "w14:val": actualChar,
                                                },
                                            },
                                        },
                                        {
                                            "w14:uncheckedState": {
                                                _attr: {
                                                    "w14:font": "Segoe UI Symbol",
                                                    "w14:val": "2705",
                                                },
                                            },
                                        },
                                    ],
                                },
                            ],
                        },
                        {
                            "w:sdtContent": [
                                {
                                    "w:r": [
                                        {
                                            "w:sym": {
                                                _attr: {
                                                    "w:char": actualChar,
                                                    "w:font": actualFont,
                                                },
                                            },
                                        },
                                    ],
                                },
                            ],
                        },
                    ],
                });
            },
        );

        it("should create a CheckBox with proper root, custom state, and no alias", () => {
            const checkBox = new CheckBox({
                checked: false,
                checkedState: {
                    font: "Segoe UI Symbol",
                    value: "2713",
                },
                uncheckedState: {
                    font: "Segoe UI Symbol",
                    value: "2705",
                },
            });

            const tree = new Formatter().format(checkBox);

            expect(tree).to.deep.equal({
                "w:sdt": [
                    {
                        "w:sdtPr": [
                            {
                                "w14:checkbox": [
                                    {
                                        "w14:checked": {
                                            _attr: {
                                                "w14:val": "0",
                                            },
                                        },
                                    },
                                    {
                                        "w14:checkedState": {
                                            _attr: {
                                                "w14:font": "Segoe UI Symbol",
                                                "w14:val": "2713",
                                            },
                                        },
                                    },
                                    {
                                        "w14:uncheckedState": {
                                            _attr: {
                                                "w14:font": "Segoe UI Symbol",
                                                "w14:val": "2705",
                                            },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "w:sdtContent": [
                            {
                                "w:r": [
                                    {
                                        "w:sym": {
                                            _attr: {
                                                "w:char": "2705",
                                                "w:font": "Segoe UI Symbol",
                                            },
                                        },
                                    },
                                ],
                            },
                        ],
                    },
                ],
            });
        });
    });
});
