import { describe, expect, it } from "vite-plus/test";

import { officeOpenTools } from "./index";

describe("officeOpenTools", () => {
  it("should export three tools with correct keys", () => {
    const keys = Object.keys(officeOpenTools);
    expect(keys).toContain("generate-docx");
    expect(keys).toContain("generate-pptx");
    expect(keys).toContain("generate-xlsx");
    expect(keys).toHaveLength(3);
  });

  it("each tool should have a function execute", () => {
    for (const tool of Object.values(officeOpenTools)) {
      expect(typeof tool.execute).toBe("function");
    }
  });
});
