import { describe, expect, it } from "vite-plus/test";

import { buildCorePropertiesXmlString } from "../opc/core";
import { createDataModel } from "../smartart/tree-to-model";
import { nextVmlShapeId } from "../vector/shapes";
import { uniqueId, uniqueUuid } from "./generators";
import { activeReproducibleScope, withReproducibleGeneration } from "./reproducible";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

describe("withReproducibleGeneration", () => {
  it("returns the fn result and removes the scope afterwards", () => {
    const value = withReproducibleGeneration({}, () => {
      expect(activeReproducibleScope()).toBeDefined();
      return 42;
    });
    expect(value).toBe(42);
    expect(activeReproducibleScope()).toBeUndefined();
  });

  it("defaults the date to the Unix epoch", () => {
    withReproducibleGeneration({}, () => {
      expect(activeReproducibleScope()?.date).toBe("1970-01-01T00:00:00.000Z");
    });
  });

  it("removes the scope when a sync fn throws", () => {
    expect(() =>
      withReproducibleGeneration({}, () => {
        throw new Error("boom");
      }),
    ).toThrow("boom");
    expect(activeReproducibleScope()).toBeUndefined();
  });

  it("removes the scope when an async fn rejects", async () => {
    await expect(
      withReproducibleGeneration({}, async () => {
        throw new Error("boom");
      }),
    ).rejects.toThrow("boom");
    expect(activeReproducibleScope()).toBeUndefined();
  });

  it("nests, the innermost scope winning", () => {
    withReproducibleGeneration({ date: "outer" }, () => {
      expect(activeReproducibleScope()?.date).toBe("outer");
      withReproducibleGeneration({ date: "inner" }, () => {
        expect(activeReproducibleScope()?.date).toBe("inner");
      });
      expect(activeReproducibleScope()?.date).toBe("outer");
    });
    expect(activeReproducibleScope()).toBeUndefined();
  });

  it("keeps the outer scope installed while an inner async fn settles", async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const outer = withReproducibleGeneration({ date: "outer" }, async () => {
      await gate;
      return activeReproducibleScope()?.date;
    });
    const inner = withReproducibleGeneration({ date: "inner" }, async () => {
      return activeReproducibleScope()?.date;
    });
    expect(await inner).toBe("inner");
    expect(activeReproducibleScope()?.date).toBe("outer");
    release();
    expect(await outer).toBe("outer");
    expect(activeReproducibleScope()).toBeUndefined();
  });

  it("gives concurrent scopes independent counters", async () => {
    const solo = withReproducibleGeneration({ date: "solo" }, () => [uniqueId(), uniqueUuid()]);
    const generate = (date: string) =>
      withReproducibleGeneration({ date }, async () => {
        // The compile-like id reads are synchronous; only the pack awaits.
        const ids = [uniqueId(), uniqueUuid()];
        await new Promise((resolve) => setTimeout(resolve, 1));
        return ids;
      });
    const [first, second] = await Promise.all([generate("a"), generate("b")]);
    expect(first).toEqual(solo);
    expect(second).toEqual(solo);
    expect(activeReproducibleScope()).toBeUndefined();
  });
});

describe("reproducible scope readers", () => {
  it("makes uniqueId deterministic and 21 lowercase alphanumerics", () => {
    const ids = withReproducibleGeneration({}, () => [uniqueId(), uniqueId()]);
    expect(ids[0]).toMatch(/^[a-z0-9]{21}$/);
    expect(ids[0]).not.toBe(ids[1]);
    expect(withReproducibleGeneration({}, () => uniqueId())).toBe(ids[0]);
  });

  it("makes uniqueUuid a deterministic v4-shaped UUID", () => {
    const uuid = withReproducibleGeneration({}, () => uniqueUuid());
    expect(uuid).toMatch(UUID_RE);
    expect(withReproducibleGeneration({}, () => uniqueUuid())).toBe(uuid);
  });

  it("keeps uniqueId and uniqueUuid random outside a scope", () => {
    expect(uniqueId()).not.toBe(uniqueId());
    expect(uniqueUuid()).toMatch(UUID_RE);
  });

  it("feeds the scope date to core-properties defaults", () => {
    const xml = withReproducibleGeneration({ date: "2020-02-02T00:00:00.000Z" }, () =>
      buildCorePropertiesXmlString({}),
    );
    expect(xml.match(/2020-02-02T00:00:00\.000Z/g)).toHaveLength(2);
  });

  it("allocates VML shape ids from the scope counter", () => {
    expect(withReproducibleGeneration({}, () => [nextVmlShapeId(), nextVmlShapeId()])).toEqual([
      "_x0000_s1025",
      "_x0000_s1026",
    ]);
  });

  it("makes SmartArt GUIDs deterministic", () => {
    const first = withReproducibleGeneration({}, () => createDataModel([{ text: "a" }]));
    const second = withReproducibleGeneration({}, () => createDataModel([{ text: "a" }]));
    expect(first).toBe(second);
    expect(first).toMatch(/modelId="\{[0-9A-F-]+\}"/);
  });
});
