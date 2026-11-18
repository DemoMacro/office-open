import { lengthToEmu } from "../runtime/xml";

/** Simple ODF group transform values that map to the canonical group model. */
export interface OdfGroupTransform {
  x?: number;
  y?: number;
  rotation?: number;
}

/** Parses ODF translate/rotate transforms that have a canonical projection. */
export function parseOdfGroupTransform(value: string | undefined): OdfGroupTransform | null {
  if (value === undefined) return {};
  const expressions = [...value.matchAll(/(translate|rotate)\s*\(([^)]*)\)/gu)];
  const remainder = value.replace(/(translate|rotate)\s*\(([^)]*)\)/gu, "").trim();
  if (expressions.length === 0 || remainder.length > 0) return null;
  const result: OdfGroupTransform = {};
  for (const expression of expressions) {
    const values = (expression[2] ?? "").trim().split(/[ ,]+/u).filter(Boolean);
    if (expression[1] === "translate") {
      if (values.length !== 1 && values.length !== 2) return null;
      const x = lengthToEmu(values[0]);
      const y = lengthToEmu(values[1] ?? "0cm");
      if (x === undefined || y === undefined) return null;
      result.x = (result.x ?? 0) + x;
      result.y = (result.y ?? 0) + y;
      continue;
    }
    if (values.length !== 1) return null;
    const rotation = Number((values[0] ?? "").replace(/deg$/u, ""));
    if (!Number.isFinite(rotation)) return null;
    result.rotation = (result.rotation ?? 0) + rotation;
  }
  return result;
}

/** Serializes the canonical group rotation as an ODF transform. */
export function odfGroupTransform(rotation: number | undefined): string | undefined {
  return rotation === undefined ? undefined : `rotate(${rotation}deg)`;
}
