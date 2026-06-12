/**
 * Format tool execution errors into LLM-friendly messages.
 *
 * When a tool's `execute` function throws, the AI SDK captures the error as a
 * `tool-error` content part and sends it back to the LLM for self-correction.
 * This module rewrites internal runtime errors into messages the LLM can
 * understand and act on.
 *
 * @module
 */

export function formatToolError(type: string, error: unknown): string {
  const msg = error instanceof Error ? error.message : String(error);

  // ── DOCX: paragraph child errors ──

  if (msg.includes("Unsupported paragraph child type:")) {
    const keys = msg.split(": ").pop() ?? "";
    return (
      `Invalid paragraph child "${keys}". ` +
      `Paragraph children must use a wrapper key: ` +
      `{ paragraph: { children: [{ text: "...", bold?: true }] } }, ` +
      `{ table: { rows: [...] } }, { image: { ... } }, ` +
      `{ toc: { ... } }, { textbox: { ... } }, ` +
      `{ pageBreak: true }, { columnBreak: true }, etc. ` +
      `Do not use raw property names like { bold: ... } as paragraph children.`
    );
  }

  if (msg.includes("Unsupported run child type:")) {
    const keys = msg.split(": ").pop() ?? "";
    return (
      `Invalid run child "${keys}". ` +
      `Run children must be text objects: { text: "...", bold?: true, italic?: true, size?: number, color?: "RRGGBB", ... }. ` +
      `The "text" key is required. Plain strings are also accepted as children. ` +
      `Do not use bare property objects like { bold: true } without a "text" key.`
    );
  }

  if (msg.includes("Unknown section child type")) {
    return (
      `Unknown section child. ` +
      `Section children must use a wrapper key: ` +
      `{ paragraph: { ... } }, { table: { ... } }, { image: { ... } }, ` +
      `{ toc: { ... } }, { textbox: { ... } }, { pageBreak: true }, ` +
      `{ sdt: { ... } }, { altChunk: { ... } }, etc.`
    );
  }

  // ── General: iterable errors ──

  if (msg.includes("not iterable")) {
    const field = type === "docx" ? "sections" : type === "pptx" ? "slides" : "worksheets";
    return `"${field}" must be an array. Received a non-iterable value.`;
  }

  // ── Fallback ──

  return `${type.toUpperCase()} generation failed: ${msg}`;
}
