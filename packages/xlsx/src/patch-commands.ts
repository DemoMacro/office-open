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

import type { WorkbookOptions } from "./parts/file";

/** Canonical root fields addressable by an XLSX patch command. */
export type WorkbookPatchField = PatchField<WorkbookOptions>;

type WorkbookPatchCollectionField = {
  [Field in WorkbookPatchField]: NonNullable<WorkbookOptions[Field]> extends readonly unknown[]
    ? Field
    : never;
}[WorkbookPatchField];

/** Canonical item that can be inserted into an XLSX collection field. */
export type WorkbookPatchInsertValue = PatchNestedValue<
  WorkbookOptions[WorkbookPatchCollectionField]
>;

/** Canonical value assignable to an addressed XLSX root or nested node. */
export type WorkbookPatchValue =
  | WorkbookOptions[WorkbookPatchField]
  | PatchNestedValue<WorkbookOptions[WorkbookPatchField]>;

/** Canonical partial value accepted by an XLSX update command. */
export type WorkbookPatchUpdateValue = WorkbookPatchValue | WorkbookPatchInsertValue;

/** One ordered canonical XLSX edit. */
export type WorkbookPatchCommand =
  | {
      type: "insert";
      at: PatchAddress<WorkbookPatchField>;
      position?: PatchPosition;
      value: readonly WorkbookPatchInsertValue[];
    }
  | { type: "replace"; at: PatchNodeAddress<WorkbookPatchField>; value: WorkbookPatchValue }
  | {
      type: "replace";
      at: PatchTextAddress<WorkbookPatchField>;
      value: string;
    }
  | {
      type: "update";
      at: PatchNodeAddress<WorkbookPatchField>;
      set?: WorkbookPatchUpdateValue;
      unset?: PatchPropertyPath;
    }
  | { type: "remove"; at: PatchNodeAddress<WorkbookPatchField> }
  | { type: "remove"; at: PatchTextAddress<WorkbookPatchField> }
  | { type: "move"; at: PatchNodeAddress<WorkbookPatchField>; to: PatchPosition };

/**
 * Options for patching an existing XLSX package through its canonical model.
 *
 * @property packageVariant - Output package variant; required when patching a macro/template source into the same variant.
 */
export interface WorkbookPatchOptions<
  T extends OutputType = "uint8array",
> extends PatchOptions<WorkbookPatchCommand> {
  outputType?: T;
  packageVariant?: OoxmlPackageVariant;
}
