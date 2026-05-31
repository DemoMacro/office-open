import { describe, expect, it, vi } from "vite-plus/test";

import { HeaderWrapper } from "./header-wrapper";
import { Media } from "./media";
import { Paragraph } from "./paragraph";
import { Table, TableCell, TableRow } from "./table";

describe("HeaderWrapper", () => {
  describe("#add", () => {
    it("should call the underlying header's addChildElement for Paragraph", () => {
      const wrapper = new HeaderWrapper(new Media(), 1);
      const spy = vi.spyOn(wrapper.view, "add");
      wrapper.add(new Paragraph({}));

      expect(spy).toBeCalled();
    });

    it("should call the underlying header's addChildElement for Table", () => {
      const wrapper = new HeaderWrapper(new Media(), 1);
      const spy = vi.spyOn(wrapper.view, "add");
      wrapper.add(
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
      const file = new HeaderWrapper(media, 1);
      expect(file.media).to.equal(media);
    });
  });
});
