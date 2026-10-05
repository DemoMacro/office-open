import { assertAllowed, reject } from "./context";
import { rtfText } from "./escape";

export function writeComplexField(field: { instruction: string; result?: string }): string {
  return [
    "{\\field",
    "{\\*\\fldinst ",
    rtfText(field.instruction),
    "}",
    field.result === undefined ? "" : `{\\fldrslt ${rtfText(field.result)}}`,
    "}",
  ].join("");
}

export function writePageReference(
  reference: { bookmarkId: string; hyperlink?: boolean },
  path: string,
): string {
  if (reference.hyperlink === false) {
    reject(
      "field",
      `${path}.pageReference`,
      "hyperlink",
      "RTF PAGE fields are hyperlinked by projection",
    );
  }
  return writeComplexField({ instruction: `PAGEREF ${reference.bookmarkId}` });
}

export function isFieldChild(child: object): boolean {
  return "complexField" in child || "pageReference" in child;
}

export function writeFieldChild(child: object, path: string): string | undefined {
  if ("complexField" in child) {
    assertAllowed(child.complexField, ["instruction", "result"], "field", `${path}.complexField`);
    const field = child.complexField;
    if (typeof field !== "object" || field === null) {
      reject("field", `${path}.complexField`, "instruction", "field options are required");
    }
    return writeComplexField(field as { instruction: string; result?: string });
  }
  if ("pageReference" in child) {
    assertAllowed(
      child.pageReference,
      ["bookmarkId", "hyperlink"],
      "field",
      `${path}.pageReference`,
    );
    const reference = child.pageReference;
    if (typeof reference !== "object" || reference === null) {
      reject("field", `${path}.pageReference`, "bookmarkId", "bookmark reference is required");
    }
    return writePageReference(reference as { bookmarkId: string; hyperlink?: boolean }, path);
  }
  return undefined;
}
