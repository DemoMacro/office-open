import type { ConcreteNumbering } from "@file/numbering";

export function replaceNumberingPlaceholders(
  xml: string,
  concreteNumberings: readonly ConcreteNumbering[],
): string {
  let result = xml;

  for (const { reference, instance, numId } of concreteNumberings) {
    result = result.replace(new RegExp(`{${reference}-${instance}}`, "g"), numId.toString());
  }

  return result;
}
