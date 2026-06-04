import { describe, expect, it } from "vite-plus/test";

import { createCellMargin, createTableCellMargin } from "./table-cell-margin";

describe("TableCellMargin", () => {
  describe("#createTableCellMargin", () => {
    it("should return undefined if there are no margins specified", () => {
      const cellMargin = createTableCellMargin({});
      expect(cellMargin).to.be.undefined;
    });
  });

  describe("#createCellMargin", () => {
    it("should return undefined if there are no margins specified", () => {
      const cellMargin = createCellMargin({});
      expect(cellMargin).to.be.undefined;
    });
  });
});
