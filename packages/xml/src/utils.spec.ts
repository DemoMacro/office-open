import { describe, expect, it } from "vite-plus/test";

import type { Element } from "./types";
import {
    findChild,
    children,
    allChildren,
    childText,
    textOf,
    collectText,
    attr,
    attrNum,
    attrBool,
    hasChild,
    findDeep,
    childCount,
} from "./utils";

describe("findChild", () => {
    it("should find first direct child by name", () => {
        const parent: Element = {
            elements: [
                { name: "a", text: "1" },
                { name: "b", text: "2" },
                { name: "a", text: "3" },
            ],
        };
        const result = findChild(parent, "a");
        expect(result).toBeDefined();
        expect(textOf(result)).toBe("1");
    });

    it("should return undefined when not found", () => {
        const parent: Element = { elements: [{ name: "a" }] };
        expect(findChild(parent, "b")).toBeUndefined();
    });

    it("should handle undefined parent", () => {
        expect(findChild(undefined, "a")).toBeUndefined();
    });
});

describe("children", () => {
    it("should return all matching children", () => {
        const parent: Element = {
            elements: [
                { name: "a", text: "1" },
                { name: "b", text: "2" },
                { name: "a", text: "3" },
            ],
        };
        const result = children(parent, "a");
        expect(result).toHaveLength(2);
    });

    it("should return empty array for no matches", () => {
        const parent: Element = { elements: [{ name: "a" }] };
        expect(children(parent, "b")).toEqual([]);
    });

    it("should handle undefined parent", () => {
        expect(children(undefined, "a")).toEqual([]);
    });
});

describe("allChildren", () => {
    it("should return all direct children", () => {
        const parent: Element = {
            elements: [{ name: "a" }, { name: "b" }],
        };
        expect(allChildren(parent)).toHaveLength(2);
    });

    it("should handle undefined parent", () => {
        expect(allChildren(undefined)).toEqual([]);
    });
});

describe("childText", () => {
    it("should return text of first matching child", () => {
        const parent: Element = {
            elements: [{ name: "a", text: "hello" }],
        };
        expect(childText(parent, "a")).toBe("hello");
    });

    it("should return empty string when not found", () => {
        expect(childText({ elements: [] }, "a")).toBe("");
    });
});

describe("textOf", () => {
    it("should return .text property", () => {
        expect(textOf({ text: "hello" })).toBe("hello");
    });

    it("should join child element text", () => {
        const el: Element = {
            elements: [{ text: "hello" }, { text: " " }, { text: "world" }],
        };
        expect(textOf(el)).toBe("hello world");
    });

    it("should return empty string for undefined", () => {
        expect(textOf(undefined)).toBe("");
    });
});

describe("collectText", () => {
    it("should collect text recursively", () => {
        const el: Element = {
            elements: [{ text: "a" }, { elements: [{ text: "b" }, { text: "c" }] }],
        };
        expect(collectText(el)).toBe("abc");
    });

    it("should return empty string for undefined", () => {
        expect(collectText(undefined)).toBe("");
    });
});

describe("attr", () => {
    it("should return attribute value as string", () => {
        const el: Element = { attributes: { foo: "bar" } };
        expect(attr(el, "foo")).toBe("bar");
    });

    it("should return undefined for missing attribute", () => {
        expect(attr({ attributes: {} }, "foo")).toBeUndefined();
    });

    it("should handle undefined element", () => {
        expect(attr(undefined, "foo")).toBeUndefined();
    });

    it("should convert number attribute to string", () => {
        const el: Element = { attributes: { count: 42 } };
        expect(attr(el, "count")).toBe("42");
    });
});

describe("attrNum", () => {
    it("should return number value", () => {
        const el: Element = { attributes: { count: "42" } };
        expect(attrNum(el, "count")).toBe(42);
    });

    it("should return undefined for non-numeric", () => {
        const el: Element = { attributes: { count: "abc" } };
        expect(attrNum(el, "count")).toBeUndefined();
    });

    it("should handle undefined element", () => {
        expect(attrNum(undefined, "count")).toBeUndefined();
    });
});

describe("attrBool", () => {
    it("should return true for 'true'/'1'", () => {
        const el: Element = { attributes: { flag: "true" } };
        expect(attrBool(el, "flag")).toBe(true);
    });

    it("should return false for 'false'/'0'", () => {
        const el: Element = { attributes: { flag: "false" } };
        expect(attrBool(el, "flag")).toBe(false);
    });

    it("should handle string 'true' attribute", () => {
        const el: Element = { attributes: { flag: "true" } };
        expect(attrBool(el, "flag")).toBe(true);
    });
});

describe("hasChild", () => {
    it("should return true when child exists", () => {
        const parent: Element = { elements: [{ name: "a" }] };
        expect(hasChild(parent, "a")).toBe(true);
    });

    it("should return false when child does not exist", () => {
        const parent: Element = { elements: [{ name: "a" }] };
        expect(hasChild(parent, "b")).toBe(false);
    });
});

describe("findDeep", () => {
    it("should find deeply nested elements", () => {
        const parent: Element = {
            elements: [
                {
                    elements: [{ name: "target", text: "found" }],
                },
            ],
        };
        const result = findDeep(parent, "target");
        expect(result).toHaveLength(1);
        expect(textOf(result[0])).toBe("found");
    });

    it("should find multiple matches at different depths", () => {
        const parent: Element = {
            elements: [
                { name: "target", text: "1" },
                { elements: [{ name: "target", text: "2" }] },
            ],
        };
        expect(findDeep(parent, "target")).toHaveLength(2);
    });

    it("should handle undefined parent", () => {
        expect(findDeep(undefined, "a")).toEqual([]);
    });
});

describe("childCount", () => {
    it("should count direct children", () => {
        const parent: Element = { elements: [{ name: "a" }, { name: "b" }] };
        expect(childCount(parent)).toBe(2);
    });

    it("should handle undefined parent", () => {
        expect(childCount(undefined)).toBe(0);
    });
});
