/**
 * Glossary compile phase: the glossary document part, its companion parts
 * (numbering, styles, settings, webSettings, fontTable), the glossary's own
 * relationships, and the companion relationships.
 *
 * The glossary body stringifies against a dedicated view wrapper so its
 * hyperlinks and drawings register on the glossary's rels instead of the
 * main document rels; media placeholders re-emit source rIds on round-trip.
 *
 * @module
 */

import {
  RELATIONSHIP_TYPES,
  Relationships,
  findAndReplaceImagePlaceholders,
  optionalRelsPart,
  partPathToRelsPath,
  replaceNumberingPlaceholders,
  type RelationshipType,
  type XmlifyedFile,
} from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { documentNamespaceDialect } from "@parts/document/document-attributes";
import { FontWrapper } from "@parts/fonts/font-wrapper";
import type { NoteChild, NotesData } from "@parts/notes/shared";
import { DefaultStylesFactory, stringifyDocDefaults } from "@parts/styles/factory";
import type { StylesOptions } from "@parts/styles/styles";

import type { PartCtxFactory } from "../compiler";
import type { BodyContext, DocxWriteContext } from "../context";
import {
  endnotesDesc,
  fontTableDesc,
  footnotesDesc,
  glossaryDesc,
  settingsDesc,
  webSettingsDesc,
} from "../parts";
import { Numbering } from "../parts/numbering";
import { Styles } from "../parts/styles/styles";
import { compileHeaderFooterPart } from "./headerfooter";
import {
  XML_DECL,
  registerPartMedia,
  relativePartTarget,
  resolvePartCharts,
  resolvePartMedia,
} from "./shared";

/** Default companion paths relative to word/ when the round-trip name is absent. */
const DEFAULT_PART_NAMES = {
  numbering: "glossary/numbering.xml",
  settings: "glossary/settings.xml",
  styles: "glossary/styles.xml",
  webSettings: "glossary/webSettings.xml",
  fontTable: "glossary/fontTable.xml",
  footnotes: "glossary/footnotes.xml",
  endnotes: "glossary/endnotes.xml",
} as const;

/** {fileName → source rId} for glossary media rels, so the body XML's
 *  re-emitted `r:embed` keeps pointing at the same relationship. */
function glossarySourceMediaRids(ctx: DocxWriteContext, partPath: string): Map<string, string> {
  const sourceRids = new Map<string, string>();
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== partPath || !rel.target.startsWith("../media/")) continue;
    sourceRids.set(rel.target.slice("../media/".length), rel.rId);
  }
  return sourceRids;
}

/** Source relationships owned by one part, preserved verbatim for a rebuilt
 *  part whose rels the compiler re-emits. */
function sourceOwnedRels(ctx: DocxWriteContext, sourcePath: string): Relationships {
  const rels = new Relationships();
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source === sourcePath) rels.claimSourceRel(rel);
  }
  return rels;
}

/** Assemble the glossary styles part the same way the main styles part is
 *  assembled in the write context: a round-trip origin re-emits the parsed
 *  docDefaults/latentStyles verbatim, fresh authoring merges factory builtins
 *  with user styles. */
function buildGlossaryStyles(glossaryStyles: StylesOptions): Styles {
  const factory = new DefaultStylesFactory().newInstance(
    glossaryStyles.roundTripped ? {} : glossaryStyles.default,
  );
  const mergeById = <T extends { id: string }>(
    factoryStyles: T[] | undefined,
    userStyles: T[] | undefined,
  ): T[] => {
    const base = factoryStyles ?? [];
    if (!userStyles || userStyles.length === 0) return base;
    const userIds = new Set(userStyles.map((style) => style.id));
    return [...base.filter((style) => !userIds.has(style.id)), ...userStyles];
  };
  if (!glossaryStyles.roundTripped) {
    return new Styles({
      importedStyles: factory.importedStyles,
      initialAttributes: glossaryStyles.initialAttributes ?? factory.initialAttributes,
      paragraphStyles: mergeById(factory.paragraphStyles, glossaryStyles.paragraphStyles),
      characterStyles: mergeById(factory.characterStyles, glossaryStyles.characterStyles),
      tableStyles: mergeById(factory.tableStyles, glossaryStyles.tableStyles),
      numberingStyles: mergeById(factory.numberingStyles, glossaryStyles.numberingStyles),
    });
  }
  const docDefaults =
    glossaryStyles.default?.document !== undefined
      ? stringifyDocDefaults(glossaryStyles.default.document, false)
      : (glossaryStyles.docDefaultsXml ?? factory.importedStyles?.[0] ?? "");
  const latentStyles = glossaryStyles.latentStylesXml ?? "";
  return new Styles({
    importedStyles: [docDefaults, latentStyles],
    initialAttributes: glossaryStyles.initialAttributes ?? factory.initialAttributes,
    paragraphStyles: glossaryStyles.paragraphStyles,
    characterStyles: glossaryStyles.characterStyles,
    tableStyles: glossaryStyles.tableStyles,
    numberingStyles: glossaryStyles.numberingStyles,
    styleOrder: glossaryStyles.styleOrder,
  });
}

export interface GlossaryCompileResult {
  entries: {
    Glossary?: XmlifyedFile;
    GlossaryRelationships?: XmlifyedFile;
    GlossaryNumbering?: XmlifyedFile;
    GlossaryNumberingRelationships?: XmlifyedFile;
    GlossaryStyles?: XmlifyedFile;
    GlossarySettings?: XmlifyedFile;
    GlossarySettingsRelationships?: XmlifyedFile;
    GlossaryWebSettings?: XmlifyedFile;
    GlossaryFontTable?: XmlifyedFile;
    GlossaryFontTableRelationships?: XmlifyedFile;
    GlossaryFootnotes?: XmlifyedFile;
    GlossaryFootnotesRelationships?: XmlifyedFile;
    GlossaryEndnotes?: XmlifyedFile;
    GlossaryEndnotesRelationships?: XmlifyedFile;
    GlossaryHeaders?: XmlifyedFile[];
    GlossaryHeaderRelationships?: XmlifyedFile[];
    GlossaryFooters?: XmlifyedFile[];
    GlossaryFooterRelationships?: XmlifyedFile[];
  };
  /** Font wrapper whose embedded .odttf binaries the compiler packs. */
  fontTable?: FontWrapper;
}

export function compileGlossaryParts(
  ctx: DocxWriteContext,
  mkCtx: PartCtxFactory,
): GlossaryCompileResult {
  const glossary = ctx.glossaryOptions;
  if (!glossary) return { entries: {} };
  const glossaryPartPath = `word/${glossary.partName ?? "glossary/document.xml"}`;
  const entries: GlossaryCompileResult["entries"] = {};
  ctx.glossaryFontTable = undefined;
  ctx.glossaryHeaderParts = [];
  ctx.glossaryFooterParts = [];

  // Companion font table: pre-assign an odttf path that cannot collide with
  // the main font table's fresh slots, then wrap so stringify sees the
  // allocated fontKeys and the compiler packs the same binaries.
  let glossaryFontTable: FontWrapper | undefined;
  if (glossary.fonts?.length) {
    const taken = new Set(
      ctx.fontTable.fontOptionsWithKey.map(
        (font) => font.odttfPath ?? `word/fonts/${font.name.split(".")[0]}.odttf`,
      ),
    );
    for (const font of glossary.fonts) {
      if (font.data === undefined || font.odttfPath) continue;
      const fallback = `word/fonts/${font.name.split(".")[0]}.odttf`;
      if (taken.has(fallback)) font.odttfPath = `word/fonts/glossary/${font.name}.odttf`;
    }
    glossaryFontTable = new FontWrapper(glossary.fonts, ctx.reproducible, "glossary");
  }
  const fontTablePartName = glossaryFontTable
    ? (glossary.fontTablePartName ?? DEFAULT_PART_NAMES.fontTable)
    : glossary.fontTablePartName;

  const glossaryNumbering = glossary.numbering
    ? new Numbering(glossary.numbering, false)
    : undefined;
  const numberingPartName = glossaryNumbering
    ? (glossary.numberingPartName ?? DEFAULT_PART_NAMES.numbering)
    : glossary.numberingPartName;
  const settingsPartName = glossary.settings
    ? (glossary.settingsPartName ?? DEFAULT_PART_NAMES.settings)
    : glossary.settingsPartName;
  const stylesPartName = glossary.styles
    ? (glossary.stylesPartName ?? DEFAULT_PART_NAMES.styles)
    : glossary.stylesPartName;
  const webSettingsPartName = glossary.webSettings
    ? (glossary.webSettingsPartName ?? DEFAULT_PART_NAMES.webSettings)
    : glossary.webSettingsPartName;

  const glossaryRels = new Relationships();
  const previousNumbering = ctx.numbering;
  const mainNumbering = previousNumbering;
  if (glossaryNumbering) ctx.numbering = glossaryNumbering;
  try {
    const glossaryCtx = mkCtx({ relationships: glossaryRels, partName: glossaryPartPath });
    const glossaryXml = glossaryDesc.stringify(glossary, glossaryCtx) ?? "";
    const sourceMediaRids = glossarySourceMediaRids(ctx, glossaryPartPath);
    const resolvedGlossary = findAndReplaceImagePlaceholders(
      glossaryXml,
      ctx.media.array,
      1,
      "rId",
      sourceMediaRids,
    );
    for (const ref of resolvedGlossary.referenced) {
      const override = sourceMediaRids.get(ref.fileName);
      if (override !== undefined) {
        glossaryRels.addRelationship(
          override,
          RELATIONSHIP_TYPES.image,
          `../media/${ref.fileName}`,
        );
      } else {
        glossaryRels.add(RELATIONSHIP_TYPES.image, `../media/${ref.fileName}`);
      }
    }

    // Companion parts: keep the source rel id/target on round-trip, allocate
    // fresh relative targets otherwise.
    const addCompanionRel = (type: RelationshipType, partName: string): void => {
      const source = (ctx._options.passthroughRelationships ?? []).find(
        (rel) => rel.source === glossaryPartPath && rel.relationshipType === type,
      );
      if (source) {
        glossaryRels.claimSourceRel(source);
        return;
      }
      glossaryRels.add(type, relativePartTarget(glossaryPartPath, `word/${partName}`));
    };
    if (numberingPartName) addCompanionRel(RELATIONSHIP_TYPES.numbering, numberingPartName);
    if (settingsPartName) addCompanionRel(RELATIONSHIP_TYPES.settings, settingsPartName);
    if (stylesPartName) addCompanionRel(RELATIONSHIP_TYPES.styles, stylesPartName);
    if (webSettingsPartName) addCompanionRel(RELATIONSHIP_TYPES.webSettings, webSettingsPartName);
    if (fontTablePartName) addCompanionRel(RELATIONSHIP_TYPES.fontTable, fontTablePartName);
    if (glossary.footnotes || glossary.footnoteSeparators) {
      addCompanionRel(
        RELATIONSHIP_TYPES.footnotes,
        glossary.footnotesPartName ?? DEFAULT_PART_NAMES.footnotes,
      );
    }
    if (glossary.endnotes || glossary.endnoteSeparators) {
      addCompanionRel(
        RELATIONSHIP_TYPES.endnotes,
        glossary.endnotesPartName ?? DEFAULT_PART_NAMES.endnotes,
      );
    }

    entries.Glossary = {
      data:
        XML_DECL +
        (glossaryNumbering
          ? replaceNumberingPlaceholders(resolvedGlossary.xml, glossaryNumbering.concreteNumbering)
          : resolvedGlossary.xml),
      path: glossaryPartPath,
    };
    entries.GlossaryRelationships = optionalRelsPart(
      glossaryRels,
      XML_DECL,
      partPathToRelsPath(glossaryPartPath),
    );

    if (glossaryNumbering && numberingPartName) {
      const numberingPath = `word/${numberingPartName}`;
      const numberingXml = glossaryNumbering.serialize(ctx);
      const numberingMedia = findAndReplaceImagePlaceholders(numberingXml, ctx.media.array, 1);
      const numberingRels = new Relationships();
      for (const [i, ref] of numberingMedia.referenced.entries()) {
        numberingRels.addRelationship(
          1 + i,
          RELATIONSHIP_TYPES.image,
          relativePartTarget(numberingPath, `word/media/${ref.fileName}`),
        );
      }
      entries.GlossaryNumbering = { data: XML_DECL + numberingMedia.xml, path: numberingPath };
      entries.GlossaryNumberingRelationships = optionalRelsPart(
        numberingRels,
        XML_DECL,
        partPathToRelsPath(numberingPath),
      );
    }
    if (glossary.styles && stylesPartName) {
      const stylesXml = buildGlossaryStyles(glossary.styles).serialize(
        documentNamespaceDialect(ctx),
      );
      // Glossary styles reference the MAIN numbering part (numPr numId lives
      // in the main document's numbering space), so remap against the main
      // concrete numbering even while the glossary numbering is installed.
      entries.GlossaryStyles = {
        data: XML_DECL + replaceNumberingPlaceholders(stylesXml, mainNumbering.concreteNumbering),
        path: `word/${stylesPartName}`,
      };
    }
    if (glossary.settings && settingsPartName) {
      const settingsPath = `word/${settingsPartName}`;
      entries.GlossarySettings = {
        data: XML_DECL + (settingsDesc.stringify(glossary.settings, ctx) ?? ""),
        path: settingsPath,
      };
      const settingsRels = sourceOwnedRels(ctx, settingsPath);
      entries.GlossarySettingsRelationships = optionalRelsPart(
        settingsRels,
        XML_DECL,
        partPathToRelsPath(settingsPath),
      );
    }
    if (glossary.webSettings && webSettingsPartName) {
      entries.GlossaryWebSettings = {
        data: XML_DECL + (webSettingsDesc.stringify(glossary.webSettings, ctx) ?? ""),
        path: `word/${webSettingsPartName}`,
      };
    }
    if (glossaryFontTable && fontTablePartName) {
      const fontTablePath = `word/${fontTablePartName}`;
      entries.GlossaryFontTable = {
        data:
          XML_DECL +
          (fontTableDesc.stringify({ fonts: glossaryFontTable.fontOptionsWithKey }, ctx) ?? ""),
        path: fontTablePath,
      };
      entries.GlossaryFontTableRelationships = optionalRelsPart(
        glossaryFontTable.relationships,
        XML_DECL,
        partPathToRelsPath(fontTablePath),
      );
    }

    // Glossary notes parts: the same wiring shape as the main notes parts,
    // scoped to the glossary paths and the glossary numbering space.
    const compileGlossaryNotesPart = (
      desc: CustomDescriptor<NotesData, BodyContext>,
      partPath: string,
      data: NotesData,
      partKey: "GlossaryFootnotes" | "GlossaryEndnotes",
      relsKey: "GlossaryFootnotesRelationships" | "GlossaryEndnotesRelationships",
    ): void => {
      const rels = new Relationships();
      const notesCtx = mkCtx({ relationships: rels, partName: partPath });
      const xmlData = XML_DECL + (desc.stringify(data, notesCtx) ?? "");
      const relCount = rels.nextRelationshipId;
      const resolved = resolvePartMedia(xmlData, ctx, relCount);
      registerPartMedia(rels, ctx, resolved);
      const resolvedXml = resolvePartCharts(
        resolved.xml,
        ctx,
        rels,
        resolved.embeddingOffset + resolved.embeddingRefs.length,
        partPath,
      );
      // Orphaned externals (rel entries no note content references) have no
      // model field; claim them so the rebuilt rels keeps the source entries.
      for (const rel of ctx._options.passthroughRelationships ?? []) {
        if (rel.source === partPath) rels.claimSourceRel(rel);
      }
      entries[partKey] = {
        data: replaceNumberingPlaceholders(resolvedXml, ctx.numbering.concreteNumbering),
        path: partPath,
      };
      if (rels.relationshipCount > 0) {
        entries[relsKey] = optionalRelsPart(
          rels,
          XML_DECL,
          partPathToRelsPath(partPath),
        ) as XmlifyedFile;
      }
    };
    if (glossary.footnotes || glossary.footnoteSeparators) {
      const footnotes = glossary.footnotes ?? [];
      const notes = new Map<number, NoteChild[]>();
      let nextId = 1;
      for (const note of footnotes) {
        const id = note.id ?? nextId;
        nextId = Math.max(nextId, id) + 1;
        notes.set(id, note.children);
      }
      compileGlossaryNotesPart(
        footnotesDesc,
        `word/${glossary.footnotesPartName ?? DEFAULT_PART_NAMES.footnotes}`,
        {
          notes,
          separator: glossary.footnoteSeparators?.separator ?? undefined,
          continuationSeparator: glossary.footnoteSeparators?.continuationSeparator ?? undefined,
          continuationNotice: glossary.footnoteSeparators?.continuationNotice,
        },
        "GlossaryFootnotes",
        "GlossaryFootnotesRelationships",
      );
    }
    if (glossary.endnotes || glossary.endnoteSeparators) {
      const endnotes = glossary.endnotes ?? [];
      const notes = new Map<number, NoteChild[]>();
      let nextId = 1;
      for (const note of endnotes) {
        const id = note.id ?? nextId;
        nextId = Math.max(nextId, id) + 1;
        notes.set(id, note.children);
      }
      compileGlossaryNotesPart(
        endnotesDesc,
        `word/${glossary.endnotesPartName ?? DEFAULT_PART_NAMES.endnotes}`,
        {
          notes,
          separator: glossary.endnoteSeparators?.separator ?? undefined,
          continuationSeparator: glossary.endnoteSeparators?.continuationSeparator ?? undefined,
          continuationNotice: glossary.endnoteSeparators?.continuationNotice,
        },
        "GlossaryEndnotes",
        "GlossaryEndnotesRelationships",
      );
    }

    // Headers/footers registered during glossary body stringification.
    ctx.glossaryHeaderParts.forEach((entry, index) => {
      const compiled = compileHeaderFooterPart("header", entry, index, ctx, mkCtx);
      (entries.GlossaryHeaders ??= []).push(compiled.part);
      if (compiled.rels) (entries.GlossaryHeaderRelationships ??= []).push(compiled.rels);
    });
    ctx.glossaryFooterParts.forEach((entry, index) => {
      const compiled = compileHeaderFooterPart("footer", entry, index, ctx, mkCtx);
      (entries.GlossaryFooters ??= []).push(compiled.part);
      if (compiled.rels) (entries.GlossaryFooterRelationships ??= []).push(compiled.rels);
    });
  } finally {
    ctx.numbering = previousNumbering;
  }
  ctx.glossaryFontTable = glossaryFontTable;
  return { entries, fontTable: glossaryFontTable };
}
