import { applyPatchCommands } from "@office-open/core";
import type { OutputByType, OutputType } from "@office-open/core";
import { resolveRelationshipTarget } from "@office-open/core/opc";

import { generatePresentation } from "./generate";
import { parsePresentation } from "./parse";
import type { PresentationPatchOptions } from "./patch-commands";
import type { PresentationOptions, SlideOptions } from "./shared/file";

/**
 * Patch an existing PPTX package through its canonical `PresentationOptions`.
 *
 * Commands run in order against the parsed model and unrelated canonical
 * fields — including round-trip-only package fields — remain intact.
 *
 * @publicApi
 */
export const patchPresentation = async <T extends OutputType = "uint8array">({
  data,
  commands,
  outputType,
  packageVariant,
  placeholderDelimiters,
}: PresentationPatchOptions<T>): Promise<OutputByType[T]> => {
  const options = await parsePresentation(data);
  const originalSlides = [...(options.slides ?? [])];
  const originalSlideReferences = new Map(
    originalSlides.map((slide, index) => [slide, index] as const),
  );
  applyPatchCommands(options, commands, placeholderDelimiters);
  reconcileSlideSections(options, originalSlideReferences);
  reconcileCustomShowSlides(options);
  clearRemovedSlidePartPaths(options);
  reconcileInsertedSlides(options, originalSlides);
  removeDeletedSlideContentTypes(options, originalSlides);
  const output = await generatePresentation(options, {
    type: outputType ?? "uint8array",
    packageVariant,
  });
  return output as OutputByType[T];
};

/** Reindex explicit sections after slide collection removal, insertion, and reordering. */
function reconcileSlideSections(
  options: PresentationOptions,
  originalSlideReferences: ReadonlyMap<SlideOptions, number>,
): void {
  if (!options.sections) return;
  const currentIndices = new Map(
    (options.slides ?? []).map((slide, index) => [slide, index] as const),
  );
  const originalByIndex = new Map(
    [...originalSlideReferences].map(([slide, index]) => [index, slide] as const),
  );
  options.sections = options.sections.map((section) => ({
    ...section,
    slideIndices: section.slideIndices.flatMap((index) => {
      const slide = originalByIndex.get(index);
      const currentIndex = slide === undefined ? undefined : currentIndices.get(slide);
      return currentIndex === undefined ? [] : [currentIndex];
    }),
  }));
}

/** Clear source paths for slide-owned parts whose canonical content was removed. */
function clearRemovedSlidePartPaths(options: PresentationOptions): void {
  for (const slide of options.slides ?? []) {
    if (slide.notes === undefined) delete slide.notesSourcePath;
    if (!slide.comments?.length) delete slide.commentSourcePath;
  }
}

/** Remove custom-show references whose slide no longer survives patch commands. */
function reconcileCustomShowSlides(options: PresentationOptions): void {
  if (!options.customShows?.length) return;
  const sourceRelationships = new Map(
    (options.passthroughRelationships ?? [])
      .filter((relationship) => relationship.source === "ppt/presentation.xml")
      .map((relationship) => [relationship.rId, relationship.target] as const),
  );
  const retainedTargets = new Set(
    (options.slides ?? []).flatMap((slide) =>
      slide.sourcePath ? [slide.sourcePath.toLowerCase()] : [],
    ),
  );
  for (const show of options.customShows) {
    show.slides = show.slides.filter(({ rId }) => {
      const target = sourceRelationships.get(rId);
      const packagePath =
        target === undefined
          ? undefined
          : resolveRelationshipTarget("ppt/presentation.xml", target);
      return packagePath !== undefined && retainedTargets.has(packagePath.toLowerCase());
    });
  }
}

/** Assign package paths and presentation identities to slides inserted by patch commands. */
function reconcileInsertedSlides(
  options: PresentationOptions,
  originalSlides: readonly SlideOptions[],
): void {
  const usedSlidePaths = new Set(
    originalSlides.map((slide, index) =>
      (slide.sourcePath ?? `ppt/slides/slide${index + 1}.xml`).toLowerCase(),
    ),
  );
  const usedNotesPaths = new Set(
    originalSlides.flatMap((slide) =>
      slide.notesSourcePath ? [slide.notesSourcePath.toLowerCase()] : [],
    ),
  );
  const usedCommentPaths = new Set(
    originalSlides.flatMap((slide) =>
      slide.commentSourcePath ? [slide.commentSourcePath.toLowerCase()] : [],
    ),
  );
  const usedSlideIds = new Set<number>();
  for (const slide of [...originalSlides, ...(options.slides ?? [])]) {
    if (slide.slideId !== undefined) usedSlideIds.add(slide.slideId);
  }
  let nextSlideId = Math.max(255, ...usedSlideIds) + 1;

  for (const slide of options.slides ?? []) {
    if (slide.sourcePath === undefined) {
      let part = 1;
      while (usedSlidePaths.has(`ppt/slides/slide${part}.xml`)) part += 1;
      slide.sourcePath = `ppt/slides/slide${part}.xml`;
      usedSlidePaths.add(slide.sourcePath.toLowerCase());
    }
    if (slide.slideId === undefined) {
      while (usedSlideIds.has(nextSlideId)) nextSlideId += 1;
      slide.slideId = nextSlideId;
      usedSlideIds.add(nextSlideId);
      nextSlideId += 1;
    }
    if (slide.notes !== undefined && slide.notesSourcePath === undefined) {
      slide.notesSourcePath = nextPackagePath("ppt/notesSlides/notesSlide", usedNotesPaths);
    }
    if (slide.comments?.length && slide.commentSourcePath === undefined) {
      slide.commentSourcePath = nextPackagePath("ppt/comments/comment", usedCommentPaths);
    }
  }
}

/** Allocate the first unused sequentially named package path. */
function nextPackagePath(prefix: string, used: Set<string>): string {
  let part = 1;
  while (used.has(`${prefix}${part}.xml`.toLowerCase())) part += 1;
  const path = `${prefix}${part}.xml`;
  used.add(path);
  return path;
}

/** Drop content-type declarations for slide-owned parts removed by a patch command. */
function removeDeletedSlideContentTypes(
  options: PresentationOptions,
  originalSlides: readonly SlideOptions[],
): void {
  const retainedSlidePaths = new Set(
    (options.slides ?? []).map((slide) => slide.sourcePath?.toLowerCase()),
  );
  const removedPaths = new Set<string>();
  const retainedPartPaths = new Set(
    (options.slides ?? []).flatMap((slide) => [
      ...(slide.notesSourcePath ? [slide.notesSourcePath.toLowerCase()] : []),
      ...(slide.commentSourcePath ? [slide.commentSourcePath.toLowerCase()] : []),
    ]),
  );
  for (const slide of originalSlides) {
    if (slide.notesSourcePath && !retainedPartPaths.has(slide.notesSourcePath.toLowerCase())) {
      removedPaths.add(slide.notesSourcePath.toLowerCase());
    }
    if (slide.commentSourcePath && !retainedPartPaths.has(slide.commentSourcePath.toLowerCase())) {
      removedPaths.add(slide.commentSourcePath.toLowerCase());
    }
  }
  originalSlides.forEach((slide, index) => {
    const slidePath = (slide.sourcePath ?? `ppt/slides/slide${index + 1}.xml`).toLowerCase();
    if (retainedSlidePaths.has(slidePath)) return;
    removedPaths.add(slidePath);
    if (slide.notesSourcePath) removedPaths.add(slide.notesSourcePath.toLowerCase());
    if (slide.commentSourcePath) removedPaths.add(slide.commentSourcePath.toLowerCase());
  });
  options.contentTypes = options.contentTypes && {
    ...options.contentTypes,
    overrides: options.contentTypes.overrides.filter(
      (override) => !removedPaths.has(override.partName.replace(/^\/+/, "").toLowerCase()),
    ),
  };
}
