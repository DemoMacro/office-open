import { describe, expect, it } from "vite-plus/test";

import { decodeUriPath, encodeUriPath } from "./uri-path";

describe("encodeUriPath", () => {
  it("escapes spaces and non-ASCII per segment, keeping separators", () => {
    expect(encodeUriPath("word/fonts/My Font.odttf")).toBe("word/fonts/My%20Font.odttf");
    expect(encodeUriPath("word/fonts/Café Font.odttf")).toBe("word/fonts/Caf%C3%A9%20Font.odttf");
  });

  it("round-trips an already-escaped path through decode", () => {
    const escaped = "word/fonts/My%20Font.odttf";
    expect(encodeUriPath(decodeUriPath(escaped))).toBe(escaped);
  });
});

describe("decodeUriPath", () => {
  it("decodes each segment, keeping separators", () => {
    expect(decodeUriPath("word/fonts/My%20Font.odttf")).toBe("word/fonts/My Font.odttf");
    expect(decodeUriPath("word/fonts/Caf%C3%A9%20Font.odttf")).toBe("word/fonts/Café Font.odttf");
  });

  it("returns the input unchanged when a segment is malformed", () => {
    expect(decodeUriPath("word/fonts/100%.odttf")).toBe("word/fonts/100%.odttf");
  });

  it("is the identity on unescaped paths", () => {
    expect(decodeUriPath("word/fonts/Arial.odttf")).toBe("word/fonts/Arial.odttf");
  });
});
