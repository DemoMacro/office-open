import { describe, expect, it } from "vite-plus/test";

import { preserveChartDecimalAttributes } from "./sheet-drawing";

describe("preserveChartDecimalAttributes", () => {
  it("keeps descriptor-emitted chart lexemes verbatim", () => {
    expect(preserveChartDecimalAttributes('<c:x val="1.367845117845118E-2"/>')).toBe(
      '<c:x val="1.367845117845118E-2"/>',
    );
    expect(preserveChartDecimalAttributes('<c:w val="0.55000000000000004"/>')).toBe(
      '<c:w val="0.55000000000000004"/>',
    );
  });
});
