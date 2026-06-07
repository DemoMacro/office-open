import type { ConcreteNumbering } from "@file/numbering";

export function replaceNumberingPlaceholders(
  xml: string,
  concreteNumberings: ConcreteNumbering[],
): string {
  let result = xml;

  for (const { reference, instance, numId } of concreteNumberings) {
    result = result.replaceAll(`{${reference}-${instance}}`, numId.toString());
  }

  return result;
}
