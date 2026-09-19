import { describe, expect, it } from "vite-plus/test";

import { buildCorePropertiesXmlString } from "../opc/core";
import { createDataModel } from "../smartart/tree-to-model";
import { nextVmlShapeId } from "../vector/shapes";
import { uniqueId, uniqueUuid } from "./generators";
import { createReproducibleScope } from "./reproducible";

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

describe("createReproducibleScope", () => {
  it("defaults the date to the Unix epoch", () => {
    expect(createReproducibleScope().date).toBe("1970-01-01T00:00:00.000Z");
  });

  it("uses the caller-supplied date", () => {
    expect(createReproducibleScope({ date: "2020-02-02T00:00:00.000Z" }).date).toBe(
      "2020-02-02T00:00:00.000Z",
    );
  });

  it("yields independent counters per scope", () => {
    const a = createReproducibleScope();
    const b = createReproducibleScope();
    // Interleaved reads stay in lockstep — a's consumption never advances b.
    expect([a.nextId(), a.nextUuid(), a.nextVmlShapeId(), a.nextDrawingId()]).toEqual([
      b.nextId(),
      b.nextUuid(),
      b.nextVmlShapeId(),
      b.nextDrawingId(),
    ]);
  });

  it("makes nextId deterministic 21-char lowercase alphanumerics", () => {
    const scope = createReproducibleScope();
    const [first, second] = [scope.nextId(), scope.nextId()];
    expect(first).toMatch(/^[a-z0-9]{21}$/);
    expect(first).not.toBe(second);
    expect(createReproducibleScope().nextId()).toBe(first);
  });

  it("makes nextUuid a deterministic v4-shaped UUID", () => {
    const scope = createReproducibleScope();
    const uuid = scope.nextUuid();
    expect(uuid).toMatch(UUID_RE);
    expect(createReproducibleScope().nextUuid()).toBe(uuid);
  });

  it("keeps uniqueId and uniqueUuid random outside a scope", () => {
    expect(uniqueId()).not.toBe(uniqueId());
    expect(uniqueUuid()).toMatch(UUID_RE);
  });

  it("feeds the scope date to core-properties defaults", () => {
    const xml = buildCorePropertiesXmlString(
      {},
      createReproducibleScope({ date: "2020-02-02T00:00:00.000Z" }),
    );
    expect(xml.match(/2020-02-02T00:00:00\.000Z/g)).toHaveLength(2);
  });

  it("allocates VML shape ids from the scope counter", () => {
    expect([
      nextVmlShapeId(createReproducibleScope()),
      nextVmlShapeId(createReproducibleScope()),
    ]).toEqual(["_x0000_s1025", "_x0000_s1025"]);
  });

  it("keeps the VML shape id on the process-global counter when no scope is passed", () => {
    expect(nextVmlShapeId()).not.toBe(nextVmlShapeId());
  });

  it("makes SmartArt GUIDs deterministic", () => {
    const first = createDataModel(
      [{ text: "a" }],
      "default",
      "simple1",
      "accent1_2",
      createReproducibleScope(),
    );
    const second = createDataModel(
      [{ text: "a" }],
      "default",
      "simple1",
      "accent1_2",
      createReproducibleScope(),
    );
    expect(first).toBe(second);
    expect(first).toMatch(/modelId="\{[0-9A-F-]+\}"/);
  });
});
