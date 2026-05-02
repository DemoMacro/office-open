import { describe, expect, it } from "vite-plus/test";

import type { IContext } from "./base";
import { XmlComponent, IgnoreIfEmptyXmlComponent, EMPTY_OBJECT } from "./component";

const emptyContext: IContext = { stack: [] };

describe("XmlComponent", () => {
    it("should create an element with rootKey", () => {
        const el = new (class extends XmlComponent {
            public constructor() {
                super("w:p");
            }
        })();
        expect(el.prepForXml(emptyContext)).toEqual({ "w:p": {} });
    });

    it("should nest child components", () => {
        class TestEl extends XmlComponent {
            public constructor() {
                super("w:p");
                this.root.push(
                    new (class extends XmlComponent {
                        public constructor() {
                            super("w:r");
                        }
                    })(),
                );
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toEqual({
            "w:p": [{ "w:r": {} }],
        });
    });

    it("should include text content", () => {
        class TestEl extends XmlComponent {
            public constructor() {
                super("w:t");
                this.root.push("Hello World");
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toEqual({ "w:t": ["Hello World"] });
    });
});

describe("IgnoreIfEmptyXmlComponent", () => {
    it("should include element when it has content", () => {
        class TestEl extends IgnoreIfEmptyXmlComponent {
            public constructor() {
                super("w:p");
                this.root.push("content");
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toEqual({ "w:p": ["content"] });
    });

    it("should exclude element when empty", () => {
        class TestEl extends IgnoreIfEmptyXmlComponent {
            public constructor() {
                super("w:p");
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toBeUndefined();
    });

    it("should include element when includeIfEmpty is true", () => {
        class TestEl extends IgnoreIfEmptyXmlComponent {
            public constructor() {
                super("w:p", true);
            }
        }

        const el = new TestEl();
        expect(el.prepForXml(emptyContext)).toEqual({ "w:p": {} });
    });
});

describe("EMPTY_OBJECT", () => {
    it("should be a frozen empty object", () => {
        expect(Object.isSealed(EMPTY_OBJECT)).toBe(true);
        expect(Object.keys(EMPTY_OBJECT)).toHaveLength(0);
    });
});
