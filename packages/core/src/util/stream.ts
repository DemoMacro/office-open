/**
 * Wraps an in-memory `Uint8Array` as a single-chunk `ReadableStream`.
 * Useful for format packages that generate synchronously but expose a
 * streaming API for interface parity with async OOXML writers.
 */
export function streamUint8Array(data: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
}
