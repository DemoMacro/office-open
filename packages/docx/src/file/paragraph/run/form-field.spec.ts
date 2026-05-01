import { Formatter } from "@export/formatter";
import { describe, expect, it } from "vite-plus/test";

import { createBegin } from "./field";
import { createFormFieldData, FormFieldTextType } from "./form-field";

describe("createFormFieldData", () => {
    describe("checkbox", () => {
        it("should create a basic checkbox", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    checkBox: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [{ "w:checkBox": {} }],
            });
        });

        it("should create a checkbox with checked state", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    checkBox: { checked: true, default: true },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:checkBox": [{ "w:default": {} }, { "w:checked": {} }],
                    },
                ],
            });
        });

        it("should create a checkbox with sizeAuto", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    checkBox: { sizeAuto: true },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:checkBox": [{ "w:sizeAuto": {} }],
                    },
                ],
            });
        });

        it("should create a checkbox with explicit size", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    checkBox: { size: 20 },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:checkBox": [{ "w:size": { _attr: { "w:val": 20 } } }],
                    },
                ],
            });
        });
    });

    describe("dropdown list", () => {
        it("should create a dropdown with entries", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    dropDownList: { entries: ["Option A", "Option B", "Option C"] },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:ddList": [
                            { "w:listEntry": { _attr: { "w:val": "Option A" } } },
                            { "w:listEntry": { _attr: { "w:val": "Option B" } } },
                            { "w:listEntry": { _attr: { "w:val": "Option C" } } },
                        ],
                    },
                ],
            });
        });

        it("should create a dropdown with result and default", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    dropDownList: { entries: ["Yes", "No"], result: 0, default: 1 },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:ddList": [
                            { "w:result": { _attr: { "w:val": 0 } } },
                            { "w:default": { _attr: { "w:val": 1 } } },
                            { "w:listEntry": { _attr: { "w:val": "Yes" } } },
                            { "w:listEntry": { _attr: { "w:val": "No" } } },
                        ],
                    },
                ],
            });
        });
    });

    describe("text input", () => {
        it("should create a basic text input", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    textInput: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [{ "w:textInput": {} }],
            });
        });

        it("should create a text input with all options", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    textInput: {
                        type: FormFieldTextType.REGULAR,
                        default: "Enter text",
                        maxLength: 100,
                        format: "UPPERCASE",
                    },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:textInput": [
                            { "w:type": { _attr: { "w:val": "regular" } } },
                            { "w:default": { _attr: { "w:val": "Enter text" } } },
                            { "w:maxLength": { _attr: { "w:val": 100 } } },
                            { "w:format": { _attr: { "w:val": "UPPERCASE" } } },
                        ],
                    },
                ],
            });
        });

        it("should create a number text input", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    textInput: { type: FormFieldTextType.NUMBER },
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    {
                        "w:textInput": [{ "w:type": { _attr: { "w:val": "number" } } }],
                    },
                ],
            });
        });
    });

    describe("common options", () => {
        it("should include name and label", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    name: "MyField",
                    label: 42,
                    checkBox: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    { "w:name": { _attr: { "w:val": "MyField" } } },
                    { "w:label": { _attr: { "w:val": 42 } } },
                    { "w:checkBox": {} },
                ],
            });
        });

        it("should include enabled and calcOnExit", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    enabled: false,
                    calcOnExit: true,
                    textInput: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    { "w:enabled": { _attr: { "w:val": false } } },
                    { "w:calcOnExit": {} },
                    { "w:textInput": {} },
                ],
            });
        });

        it("should include tabIndex, entryMacro, and exitMacro", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    tabIndex: 5,
                    entryMacro: "OnEntry",
                    exitMacro: "OnExit",
                    textInput: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    { "w:tabIndex": { _attr: { "w:val": 5 } } },
                    { "w:entryMacro": { _attr: { "w:val": "OnEntry" } } },
                    { "w:exitMacro": { _attr: { "w:val": "OnExit" } } },
                    { "w:textInput": {} },
                ],
            });
        });

        it("should include helpText and statusText", () => {
            const tree = new Formatter().format(
                createFormFieldData({
                    helpText: { type: "text", value: "Press F1 for help" },
                    statusText: { type: "autoText", value: "StatusBarText" },
                    textInput: {},
                }),
            );
            expect(tree).to.deep.equal({
                "w:ffData": [
                    { "w:helpText": { _attr: { "w:type": "text", "w:val": "Press F1 for help" } } },
                    {
                        "w:statusText": {
                            _attr: { "w:type": "autoText", "w:val": "StatusBarText" },
                        },
                    },
                    { "w:textInput": {} },
                ],
            });
        });
    });
});

describe("createBegin with form field", () => {
    it("should create a begin field char without form field", () => {
        const tree = new Formatter().format(createBegin());
        expect(tree).to.deep.equal({
            "w:fldChar": { _attr: { "w:fldCharType": "begin" } },
        });
    });

    it("should create a begin field char with dirty flag", () => {
        const tree = new Formatter().format(createBegin(true));
        expect(tree).to.deep.equal({
            "w:fldChar": { _attr: { "w:fldCharType": "begin", "w:dirty": true } },
        });
    });

    it("should embed checkbox form field in begin field char", () => {
        const tree = new Formatter().format(
            createBegin(false, {
                name: "Check1",
                checkBox: { checked: true, sizeAuto: true },
            }),
        );
        expect(tree).to.deep.equal({
            "w:fldChar": [
                { _attr: { "w:fldCharType": "begin", "w:dirty": false } },
                {
                    "w:ffData": [
                        { "w:name": { _attr: { "w:val": "Check1" } } },
                        {
                            "w:checkBox": [{ "w:sizeAuto": {} }, { "w:checked": {} }],
                        },
                    ],
                },
            ],
        });
    });

    it("should embed dropdown form field in begin field char", () => {
        const tree = new Formatter().format(
            createBegin(false, {
                name: "DropDown1",
                dropDownList: { entries: ["Yes", "No", "Maybe"], result: 0 },
            }),
        );
        expect(tree).to.deep.equal({
            "w:fldChar": [
                { _attr: { "w:fldCharType": "begin", "w:dirty": false } },
                {
                    "w:ffData": [
                        { "w:name": { _attr: { "w:val": "DropDown1" } } },
                        {
                            "w:ddList": [
                                { "w:result": { _attr: { "w:val": 0 } } },
                                { "w:listEntry": { _attr: { "w:val": "Yes" } } },
                                { "w:listEntry": { _attr: { "w:val": "No" } } },
                                { "w:listEntry": { _attr: { "w:val": "Maybe" } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });

    it("should embed text input form field in begin field char", () => {
        const tree = new Formatter().format(
            createBegin(false, {
                name: "Text1",
                textInput: { type: FormFieldTextType.REGULAR, default: "Enter name" },
            }),
        );
        expect(tree).to.deep.equal({
            "w:fldChar": [
                { _attr: { "w:fldCharType": "begin", "w:dirty": false } },
                {
                    "w:ffData": [
                        { "w:name": { _attr: { "w:val": "Text1" } } },
                        {
                            "w:textInput": [
                                { "w:type": { _attr: { "w:val": "regular" } } },
                                { "w:default": { _attr: { "w:val": "Enter name" } } },
                            ],
                        },
                    ],
                },
            ],
        });
    });
});
