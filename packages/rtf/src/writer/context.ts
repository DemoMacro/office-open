import type { DocumentOptions, RunOptions } from "@office-open/docx";

import { RtfGenerateError } from "../errors";

export interface RtfGenerateContext {
  colors: string[];
  fonts: string[];
}

export function createContext(): RtfGenerateContext {
  return { colors: [], fonts: [] };
}

export function reject(part: string, path: string, name: string, reason: string): never {
  throw new RtfGenerateError(`${path}.${name}: ${reason}`, { part, path, name, reason });
}

export function assertAllowed(
  value: unknown,
  allowed: readonly string[],
  part: string,
  path: string,
): void {
  if (typeof value !== "object" || value === null) return;
  for (const name of Object.keys(value)) {
    if (!allowed.includes(name)) {
      reject(part, path, name, "RTF cannot represent this field without loss");
    }
  }
}

export function registerFont(context: RtfGenerateContext, font: string): number {
  const existing = context.fonts.indexOf(font);
  if (existing >= 0) return existing;
  context.fonts.push(font);
  return context.fonts.length - 1;
}

export function fontIndex(
  context: RtfGenerateContext,
  run: RunOptions,
  path: string,
): number | undefined {
  if (run.font === undefined) return undefined;
  if (typeof run.font !== "string") {
    reject("run", `${path}.font`, "font", "RTF writer supports concrete font names only");
  }
  return registerFont(context, run.font);
}

export function colorIndex(
  context: RtfGenerateContext,
  run: RunOptions,
  path: string,
): number | undefined {
  if (run.color === undefined) return undefined;
  if (typeof run.color !== "string") {
    reject("run", `${path}.color`, "color", "RTF writer supports hex colors only");
  }
  return registerColor(context, run.color);
}

export function registerColor(context: RtfGenerateContext, color: string): number {
  const existing = context.colors.indexOf(color);
  if (existing >= 0) return existing + 1;
  context.colors.push(color);
  return context.colors.length;
}

export type RtfMetadata = Pick<
  DocumentOptions,
  "title" | "subject" | "creator" | "keywords" | "description" | "category"
>;

export function supportedMetadata(options: DocumentOptions): RtfMetadata {
  return {
    ...(options.title !== undefined ? { title: options.title } : {}),
    ...(options.subject !== undefined ? { subject: options.subject } : {}),
    ...(options.creator !== undefined ? { creator: options.creator } : {}),
    ...(options.keywords !== undefined ? { keywords: options.keywords } : {}),
    ...(options.description !== undefined ? { description: options.description } : {}),
    ...(options.category !== undefined ? { category: options.category } : {}),
  };
}
