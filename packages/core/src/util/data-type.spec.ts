import { describe, expect, it } from "vite-plus/test";

import { toUint8ArrayAsync } from "./data-type";

describe("async binary input", () => {
  it("rejects a readable stream that exceeds the materialization limit", async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([1, 2]));
        controller.enqueue(new Uint8Array([3]));
        controller.close();
      },
    });

    await expect(toUint8ArrayAsync(stream, { maxBytes: 2 })).rejects.toThrow(
      "Binary input exceeds configured size limit",
    );
  });

  it("rejects a blob before materializing it", async () => {
    await expect(
      toUint8ArrayAsync(new Blob([new Uint8Array([1, 2, 3])]), { maxBytes: 2 }),
    ).rejects.toThrow("Binary input exceeds configured size limit");
  });
});
