import type {
  PatchAddress,
  PatchField,
  PatchNodeAddress,
  PatchNestedValue,
  PatchPosition,
  PatchPropertyPath,
  PatchTextAddress,
  PatchOptions,
} from "@office-open/core";
import type { OoxmlPackageVariant, OutputType } from "@office-open/core";

import type { PresentationOptions } from "./shared/file";

/** Canonical root fields addressable by a PPTX patch command. */
export type PresentationPatchField = PatchField<PresentationOptions>;

type PresentationPatchCollectionField = {
  [Field in PresentationPatchField]: NonNullable<
    PresentationOptions[Field]
  > extends readonly unknown[]
    ? Field
    : never;
}[PresentationPatchField];

/** Canonical item that can be inserted into a PPTX collection field. */
export type PresentationPatchInsertValue = PatchNestedValue<
  PresentationOptions[PresentationPatchCollectionField]
>;

/** Canonical value assignable to an addressed PPTX root or nested node. */
export type PresentationPatchValue =
  | PresentationOptions[PresentationPatchField]
  | PatchNestedValue<PresentationOptions[PresentationPatchField]>;

/** Canonical partial value accepted by a PPTX update command. */
export type PresentationPatchUpdateValue = PresentationPatchValue | PresentationPatchInsertValue;

/** One ordered canonical PPTX edit. */
export type PresentationPatchCommand =
  | {
      type: "insert";
      at: PatchAddress<PresentationPatchField>;
      position?: PatchPosition;
      value: readonly PresentationPatchInsertValue[];
    }
  | { type: "replace"; at: PatchNodeAddress<PresentationPatchField>; value: PresentationPatchValue }
  | {
      type: "replace";
      at: PatchTextAddress<PresentationPatchField>;
      value: string;
    }
  | {
      type: "update";
      at: PatchNodeAddress<PresentationPatchField>;
      set?: PresentationPatchUpdateValue;
      unset?: PatchPropertyPath;
    }
  | { type: "remove"; at: PatchNodeAddress<PresentationPatchField> }
  | { type: "remove"; at: PatchTextAddress<PresentationPatchField> }
  | { type: "move"; at: PatchNodeAddress<PresentationPatchField>; to: PatchPosition };

/**
 * Options for patching an existing PPTX package through its canonical model.
 *
 * @property packageVariant - Output package variant; required when patching a macro/template source into the same variant.
 */
export interface PresentationPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<PresentationPatchCommand> {
  outputType?: T;
  packageVariant?: OoxmlPackageVariant;
}
