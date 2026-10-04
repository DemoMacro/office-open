import { expect, it } from "vitest";

import { tokenizeRtf } from "../tokenizer";
import { parseRtfTokens } from "./document";

it("maps metadata and body tokens to a canonical document", () => {
  const source = "{\\rtf1\\ansi{\\info{\\title Boundary}}\\par Body}";
  const document = parseRtfTokens(tokenizeRtf(source), source);

  expect(document.title).toBe("Boundary");
  expect(document.sections?.[0]?.children).toEqual([
    { paragraph: { children: [{ text: "Body" }] } },
  ]);
});
