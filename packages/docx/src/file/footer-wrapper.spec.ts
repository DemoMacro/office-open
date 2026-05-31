import { describe, expect, it, vi } from "vite-plus/test";

import { FooterWrapper } from "./footer-wrapper";
import { Media } from "./media";
import { Paragraph } from "./paragraph";
import { Table, TableCell, TableRow } from "./table";

describe("FooterWrapper", () => {
  describe("#add", () => {
    it("should call the underlying footer's addParagraph", () => {
      const file = new FooterWrapper(new Media(), 1);
      const spy = vi.spyOn(file.view, "add");
      file.add(new Paragraph({}));

      expect(spy).toBeCalled();
    });

    it("should call the underlying footer's addParagraph", () => {
      const file = new FooterWrapper(new Media(), 1);
      const spy = vi.spyOn(file.view, "add");
      file.add(
        new Table({
          rows: [
            new TableRow({
              cells: [
                new TableCell({
                  children: [new Paragraph("hello")],
                }),
              ],
            }),
          ],
        }),
      );

      expect(spy).toBeCalled();
    });
  });

  describe("#Media", () => {
    it("should get Media", () => {
      const media = new Media();
      const file = new FooterWrapper(media, 1);
      expect(file.media).to.equal(media);
    });
  });
});
