export function rtfText(value: string): string {
  const parts: string[] = [];
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!;
    const code = character.charCodeAt(0);
    if (character === "{" || character === "}" || character === "\\") {
      parts.push(`\\${character}`);
    } else if (code === 0x0d) {
      parts.push("\\line ");
    } else if (code === 0x09) {
      parts.push("\\tab ");
    } else if (character === " ") {
      parts.push("\\'20");
    } else if (code < 0x80) {
      parts.push(character);
    } else {
      parts.push(`\\u${code > 0x7fff ? code - 0x10000 : code}?`);
    }
  }
  return parts.join("");
}

export function control(word: string, parameter?: number): string {
  return `\\${word}${parameter === undefined ? "" : String(parameter)} `;
}
