import { readXmlFromZip } from "@office-open/core";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { SectionJson } from "./types";

export interface ParsedLevel {
    level: number;
    format?: string;
    text?: string;
    alignment?: string;
    start?: number;
    suffix?: string;
    isLegalNumberingStyle?: boolean;
    style?: {
        paragraph?: { indent?: { left?: number; hanging?: number } };
        run?: Record<string, unknown>;
    };
}

export interface ParsedAbstractNum {
    id: string;
    levels: ParsedLevel[];
}

export interface ParsedNum {
    numId: string;
    abstractNumId: string;
    levelOverrides?: Array<{ ilvl: number; startOverride?: number }>;
}

export interface ParsedNumberingConfig {
    reference: string;
    levels: ParsedLevel[];
}

export function parseNumbering(zip: Map<string, Uint8Array>): {
    abstractNums: ParsedAbstractNum[];
    nums: ParsedNum[];
} {
    const xml = readXmlFromZip(zip, "word/numbering.xml");
    if (!xml) return { abstractNums: [], nums: [] };

    const abstractNums: ParsedAbstractNum[] = [];
    const nums: ParsedNum[] = [];

    for (const child of xml.elements ?? []) {
        if (child.name === "w:abstractNum") {
            const id = attr(child, "w:abstractNumId");
            if (id !== undefined) {
                abstractNums.push({ id, levels: parseLevels(child) });
            }
        } else if (child.name === "w:num") {
            const numId = attr(child, "w:numId");
            const abstractNumIdEl = findChild(child, "w:abstractNumId");
            const abstractNumId = abstractNumIdEl ? attr(abstractNumIdEl, "w:val") : undefined;
            if (numId !== undefined && abstractNumId !== undefined) {
                nums.push({ numId, abstractNumId, levelOverrides: parseLevelOverrides(child) });
            }
        }
    }

    return { abstractNums, nums };
}

function parseLevels(abstractNum: Element): ParsedLevel[] {
    const levels: ParsedLevel[] = [];
    for (const child of abstractNum.elements ?? []) {
        if (child.name === "w:lvl") {
            const level = parseLevel(child);
            if (level) levels.push(level);
        }
    }
    return levels;
}

function parseLevel(lvl: Element): ParsedLevel | undefined {
    const ilvl = attrNum(lvl, "w:ilvl") ?? 0;
    const result: ParsedLevel = { level: ilvl };

    const numFmt = findChild(lvl, "w:numFmt");
    if (numFmt) {
        const val = attr(numFmt, "w:val");
        if (val) result.format = val;
    }

    const lvlText = findChild(lvl, "w:lvlText");
    if (lvlText) {
        const val = attr(lvlText, "w:val");
        if (val) result.text = val;
    }

    const lvlJc = findChild(lvl, "w:lvlJc");
    if (lvlJc) {
        const val = attr(lvlJc, "w:val");
        if (val) result.alignment = val;
    }

    const start = findChild(lvl, "w:start");
    if (start) {
        const val = attrNum(start, "w:val");
        if (val !== undefined) result.start = val;
    }

    const suff = findChild(lvl, "w:suff");
    if (suff) {
        const val = attr(suff, "w:val");
        if (val) result.suffix = val;
    }

    const isLgl = findChild(lvl, "w:isLgl");
    if (isLgl) {
        const val = attr(isLgl, "w:val");
        if (val === "1" || val === "true") {
            result.isLegalNumberingStyle = true;
        }
    }

    const pPr = findChild(lvl, "w:pPr");
    if (pPr) {
        const ind = findChild(pPr, "w:ind");
        if (ind) {
            const left = attrNum(ind, "w:left");
            const hanging = attrNum(ind, "w:hanging");
            if (left !== undefined || hanging !== undefined) {
                if (!result.style) result.style = {};
                result.style.paragraph = {};
                if (left !== undefined)
                    result.style.paragraph.indent = { ...result.style.paragraph.indent, left };
                if (hanging !== undefined)
                    result.style.paragraph.indent = { ...result.style.paragraph.indent, hanging };
            }
        }
    }

    return result;
}

function parseLevelOverrides(
    num: Element,
): Array<{ ilvl: number; startOverride?: number }> | undefined {
    const overrides: Array<{ ilvl: number; startOverride?: number }> = [];
    for (const child of num.elements ?? []) {
        if (child.name === "w:lvlOverride") {
            const ilvl = attrNum(child, "w:ilvl");
            const startOverride = findChild(child, "w:startOverride");
            const start = startOverride ? attrNum(startOverride, "w:val") : undefined;
            if (ilvl !== undefined) {
                overrides.push({ ilvl, ...(start !== undefined && { startOverride: start }) });
            }
        }
    }
    return overrides.length > 0 ? overrides : undefined;
}

/**
 * Build numbering config from parsed data and remap paragraph references.
 * Only includes numIds that are actually used by paragraphs.
 */
export function buildNumberingConfig(
    data: { abstractNums: ParsedAbstractNum[]; nums: ParsedNum[] },
    sections: SectionJson[],
): ParsedNumberingConfig[] {
    if (data.abstractNums.length === 0) return [];

    const abstractMap = new Map(data.abstractNums.map((a) => [a.id, a]));
    const numToAbstract = new Map(data.nums.map((n) => [n.numId, n.abstractNumId]));

    // Collect used numIds
    const usedNumIds = new Set<string>();
    for (const section of sections) {
        collectUsedNumIds(section.children, usedNumIds);
    }

    if (usedNumIds.size === 0) return [];

    const config: ParsedNumberingConfig[] = [];
    const numIdToReference = new Map<string, string>();

    for (const numId of usedNumIds) {
        const abstractId = numToAbstract.get(numId);
        if (!abstractId) continue;
        const abstract = abstractMap.get(abstractId);
        if (!abstract) continue;

        const reference = `num-${numId}`;
        numIdToReference.set(numId, reference);

        // Deep clone levels to avoid mutating parsed data
        const levels = abstract.levels.map((l) => ({
            ...l,
            style: l.style
                ? {
                      ...l.style,
                      paragraph: l.style.paragraph
                          ? { ...l.style.paragraph, indent: { ...l.style.paragraph.indent } }
                          : undefined,
                  }
                : undefined,
        }));

        // Apply lvlOverride start values
        const num = data.nums.find((n) => n.numId === numId);
        if (num?.levelOverrides) {
            for (const override of num.levelOverrides) {
                const level = levels.find((l) => l.level === override.ilvl);
                if (level && override.startOverride !== undefined) {
                    level.start = override.startOverride;
                }
            }
        }

        config.push({ reference, levels });
    }

    // Remap paragraph references
    for (const section of sections) {
        remapNumberingReferences(section.children, numIdToReference);
    }

    return config;
}

function collectUsedNumIds(children: SectionJson["children"], used: Set<string>): void {
    for (const child of children) {
        const c = child as Record<string, unknown>;
        if (c.$type === "paragraph" && c.numbering) {
            const n = c.numbering as { reference: string };
            if (n.reference) used.add(n.reference);
        }
    }
}

function remapNumberingReferences(
    children: SectionJson["children"],
    numIdToReference: Map<string, string>,
): void {
    for (const child of children) {
        const c = child as Record<string, unknown>;
        if (c.$type === "paragraph" && c.numbering) {
            const n = c.numbering as { reference: string };
            const newRef = numIdToReference.get(n.reference);
            if (newRef) n.reference = newRef;
        }
    }
}
