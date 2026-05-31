import { describe, expect, it } from "vite-plus/test";

import { Paragraph } from "../../paragraph";
import { TableCell } from "../table-cell";
import { TableRow } from "./table-row";

describe("TableRow", () => {
  describe("#rootIndexToColumnIndex", () => {
    it("should get the correct virtual column index by root index", () => {
      const tableRow = new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
        ],
      });

      expect(tableRow.rootIndexToColumnIndex(1)).to.equal(0);
      expect(tableRow.rootIndexToColumnIndex(2)).to.equal(3);
      expect(tableRow.rootIndexToColumnIndex(3)).to.equal(4);
      expect(tableRow.rootIndexToColumnIndex(4)).to.equal(5);

      expect(() => tableRow.rootIndexToColumnIndex(0)).to.throw(
        `cell 'rootIndex' should between 1 to 4`,
      );
      expect(() => tableRow.rootIndexToColumnIndex(5)).to.throw(
        `cell 'rootIndex' should between 1 to 4`,
      );
    });
  });

  describe("#columnIndexToRootIndex", () => {
    it("should get the correct root index by virtual column index", () => {
      const tableRow = new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
        ],
      });

      expect(tableRow.columnIndexToRootIndex(0)).to.equal(1);
      expect(tableRow.columnIndexToRootIndex(1)).to.equal(1);
      expect(tableRow.columnIndexToRootIndex(2)).to.equal(1);

      expect(tableRow.columnIndexToRootIndex(3)).to.equal(2);
      expect(tableRow.columnIndexToRootIndex(4)).to.equal(3);

      expect(tableRow.columnIndexToRootIndex(5)).to.equal(4);
      expect(tableRow.columnIndexToRootIndex(6)).to.equal(4);
      expect(tableRow.columnIndexToRootIndex(7)).to.equal(4);

      expect(() => tableRow.columnIndexToRootIndex(-1)).to.throw(
        `cell 'columnIndex' should not less than zero`,
      );
      expect(() => tableRow.columnIndexToRootIndex(8)).to.throw(
        `cell 'columnIndex' should not great than 7`,
      );
    });

    it("should allow end new cell index", () => {
      const tableRow = new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
          }),
          new TableCell({
            children: [new Paragraph("test")],
            columnSpan: 3,
          }),
        ],
      });

      expect(tableRow.columnIndexToRootIndex(8, true)).to.equal(5);
      // For column 10, just place the new cell at the end of row
      expect(tableRow.columnIndexToRootIndex(10, true)).to.equal(5);
    });
  });
});
