export function rtfText(value: string): string {
  let result = "";
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    const code = character.codePointAt(0) ?? 0;
    if (character === "{" || character === "}" || character === "\\") {
      result += `\\${character}`;
    } else if (code === 0x0d) {
      result += "\\line ";
    } else if (code === 0x09) {
      result += "\\tab ";
    } else if (code < 0x80) {
      result += character;
    } else {
      result += `\\u${code > 0x7fff ? code - 0x10000 : code}?`;
    }
  }
  return result;
}

export function control(word: string, parameter?: number): string {
  return `\\${word}${parameter === undefined ? "" : String(parameter)} `;
}
