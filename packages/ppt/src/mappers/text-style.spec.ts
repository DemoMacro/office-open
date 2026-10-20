import { describe, expect, it } from "vite-plus/test";

import { RecordType } from "../records";
import { parseTextStyle } from "./text-style";

describe("legacy PowerPoint text styles", () => {
  it("keeps a zero-character collection before a later collection", () => {
    const data = new Uint8Array(32);
    const view = new DataView(data.buffer);
    view.setUint16(0, RecordType.styleTextPropAtom, true);
    view.setUint32(4, 24, true);
    view.setUint16(8, 0, true);
    view.setUint16(10, 0, true);
    view.setUint32(12, 0, true);
    view.setUint16(16, 2, true);
    view.setUint16(18, 0, true);
    view.setUint32(20, 0, true);
    view.setUint16(24, 0, true);
    view.setUint16(26, 0, true);
    view.setUint32(28, 0, true);

    const styles = parseTextStyle(
      view,
      {
        type: RecordType.styleTextPropAtom,
        offset: 0,
        version: 0,
        instance: 0,
        length: 24,
        end: 32,
        children: [],
      },
      2,
      new Map(),
    );
    expect(styles.paragraphs.map((collection) => collection.characters)).toEqual([0, 2]);
    expect(styles.characters.map((collection) => collection.characters)).toEqual([2]);
  });
});
