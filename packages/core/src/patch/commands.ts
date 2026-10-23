import { PatchError, type PatchDiagnostic } from "./diagnostics";
import type { PlaceholderDelimiters } from "./types";

/** A canonical property name or array index used below a patch field. */
export type PatchPathStep = string | number;

/** A property path used by `update.unset`. */
export type PatchPropertyPath = readonly PatchPathStep[];

/** One canonical-tree hop from the selected root field. */
export type PatchRouteStep =
  | { property: string }
  | { collection: string }
  | { index: number }
  | { id: string | number }
  | { name: string }
  | { reference: string }
  | { address: string };

/** Text selected inside the canonical subtree reached by the address. */
export type PatchTextSelection =
  | { literal: string; all?: boolean }
  | { placeholder: string; all?: boolean };

/** Canonical root field plus optional structured route to a nested node. */
export interface PatchAddress<TField extends string = string> {
  /** Root canonical field, for example `sections`, `slides`, or `worksheets`. */
  field: TField;
  /** Structured canonical route below the field; XML paths are not valid. */
  route?: readonly PatchRouteStep[];
  /** Select text inside the addressed subtree instead of the node itself. */
  text?: PatchTextSelection;
}

/** Address a canonical node rather than a text selection. */
export interface PatchNodeAddress<TField extends string = string> {
  /** Root canonical field, for example `sections`, `slides`, or `worksheets`. */
  field: TField;
  /** Structured canonical route below the field; XML paths are not valid. */
  route?: readonly PatchRouteStep[];
  /** Must remain absent when addressing a canonical node. */
  text?: undefined;
}

/** Address a text selection inside a canonical subtree. */
export interface PatchTextAddress<TField extends string = string> {
  /** Root canonical field, for example `sections`, `slides`, or `worksheets`. */
  field: TField;
  /** Structured canonical route below the field; XML paths are not valid. */
  route?: readonly PatchRouteStep[];
  /** Select text inside the addressed subtree. */
  text: PatchTextSelection;
}

/** Insertion or move destination inside an array. */
export type PatchPosition =
  | "start"
  | "end"
  | number
  | { before: PatchRouteStep }
  | { after: PatchRouteStep };

/** Unified patch actions shared by document, presentation, and workbook APIs. */
export type PatchAction = "insert" | "replace" | "update" | "remove" | "move";

/** Runtime command view consumed by `applyPatchCommands`. */
export type RuntimePatchCommand =
  | {
      type: "insert";
      at: PatchAddress;
      position?: PatchPosition;
      value: readonly unknown[];
    }
  | { type: "replace"; at: PatchAddress; value: unknown }
  | {
      type: "update";
      at: PatchAddress;
      set?: unknown;
      unset?: PatchPropertyPath;
    }
  | { type: "remove"; at: PatchAddress }
  | { type: "move"; at: PatchAddress; to: PatchPosition };

type StringKey<TKey> = TKey extends string ? TKey : never;

/** A canonical field that accepts patch commands. */
export type PatchField<TOptions extends object> = StringKey<keyof TOptions>;

/** An item of a canonical collection field. */
export type PatchCollectionItem<TOptions extends object, TField extends PatchField<TOptions>> =
  NonNullable<TOptions[TField]> extends readonly (infer Item)[] ? Item : never;

/** Every canonical collection item accepted by insert commands. */
export type PatchInsertValue<TOptions extends object> = {
  [Field in PatchField<TOptions>]: PatchCollectionItem<TOptions, Field>;
}[PatchField<TOptions>];

/** Canonical partial values accepted by update commands and their collection items. */
export type PatchUpdateValue<TOptions extends object> =
  | PartialDeep<TOptions>
  | {
      [Field in PatchField<TOptions>]: PartialDeep<PatchCollectionItem<TOptions, Field>>;
    }[PatchField<TOptions>];

type PatchNestedDepth = { 0: never; 1: 0; 2: 1; 3: 2; 4: 3; 5: 4; 6: 5 };

/** Any canonical object or nested collection item reachable within six routes. */
export type PatchNestedValue<T, Depth extends keyof PatchNestedDepth = 6> = [Depth] extends [never]
  ? never
  : T extends readonly (infer Item)[]
    ? PatchNestedValue<Item, PatchNestedDepth[Depth]>
    : T extends object
      ? T | PatchNestedValue<T[keyof T], PatchNestedDepth[Depth]>
      : never;

/** Generic insert command parameterized by a canonical options model. */
export interface PatchInsertCommand<
  TOptions extends object,
  TField extends PatchField<TOptions> = PatchField<TOptions>,
> {
  type: "insert";
  at: PatchAddress<TField>;
  position?: PatchPosition;
  value: readonly PatchInsertValue<TOptions>[];
}

/** Generic replace command parameterized by a canonical options model. */
export interface PatchReplaceCommand<
  TOptions extends object,
  TField extends PatchField<TOptions> = PatchField<TOptions>,
> {
  type: "replace";
  at: PatchNodeAddress<TField>;
  value: TOptions[TField];
}

/** Generic text-selection replace command. */
export interface PatchTextReplaceCommand<TField extends string> {
  type: "replace";
  at: PatchTextAddress<TField>;
  value: string;
}

/** Generic text-selection removal command. */
export interface PatchTextRemoveCommand<TField extends string> {
  type: "remove";
  at: PatchTextAddress<TField>;
}

/** Generic partial-update command for a format's canonical model. */
export interface PatchUpdateCommand<
  TOptions extends object,
  TField extends PatchField<TOptions> = PatchField<TOptions>,
> {
  type: "update";
  at: PatchNodeAddress<TField>;
  set?: PatchUpdateValue<TOptions>;
  unset?: PatchPropertyPath;
}

/** Generic remove command for a format's canonical model. */
export interface PatchRemoveCommand<TField extends string> {
  type: "remove";
  at: PatchNodeAddress<TField>;
}

/** Generic move command for a format's canonical model. */
export interface PatchMoveCommand<TField extends string> {
  type: "move";
  at: PatchNodeAddress<TField>;
  to: PatchPosition;
}

/** Union of the five canonical patch commands. */
export type PatchCommand<TOptions extends object> =
  | PatchInsertCommand<TOptions>
  | PatchReplaceCommand<TOptions>
  | PatchTextReplaceCommand<PatchField<TOptions>>
  | PatchUpdateCommand<TOptions>
  | PatchRemoveCommand<PatchField<TOptions>>
  | PatchTextRemoveCommand<PatchField<TOptions>>
  | PatchMoveCommand<PatchField<TOptions>>;

type PartialDeepValue<T> = T extends Date
  ? T
  : T extends readonly (infer Item)[]
    ? readonly PartialDeepValue<Item>[]
    : T extends object
      ? { [Key in keyof T]?: PartialDeepValue<T[Key]> }
      : T;

/** Recursive partial update payload for plain serializable option trees. */
export type PartialDeep<T> = T extends readonly (infer Item)[]
  ? readonly PartialDeepValue<Item>[]
  : PartialDeepValue<T>;

interface ModelTarget {
  readonly parent: object;
  readonly key: string | number;
  readonly value: unknown;
  readonly path: string;
}

/**
 * Apply ordered commands to a canonical options object. Commands see the
 * result of every preceding command; output is serialized only after success.
 */
export function applyPatchCommands<TOptions extends object>(
  options: TOptions,
  commands: readonly RuntimePatchCommand[],
  placeholderDelimiters: PlaceholderDelimiters = { start: "{{", end: "}}" },
): void {
  commands.forEach((command, index) => {
    if (command.type === "update") seedMissingRoot(options, command);
    switch (command.type) {
      case "insert":
        insertCommand(options, command, index, placeholderDelimiters);
        return;
      case "replace":
        replaceCommand(options, command, index, placeholderDelimiters);
        return;
      case "update":
        updateCommand(options, command, index, placeholderDelimiters);
        return;
      case "remove":
        removeCommand(options, command, index, placeholderDelimiters);
        return;
      case "move":
        moveCommand(options, command, index, placeholderDelimiters);
        return;
      default:
        throw patchError(
          command as RuntimePatchCommand,
          index,
          addressPath((command as RuntimePatchCommand).at),
          "invalid-operation",
          `Unknown patch command type: ${String((command as RuntimePatchCommand).type)}`,
        );
    }
  });
}

function insertCommand(
  options: object,
  command: Extract<RuntimePatchCommand, { type: "insert" }>,
  index: number,
  _delimiters: PlaceholderDelimiters,
): void {
  if (command.value.length === 0) {
    throw patchError(
      command,
      index,
      addressPath(command.at),
      "invalid-operation",
      "Insert value must contain at least one item.",
    );
  }
  const target = resolveTarget(options, command.at, index, command.type, true);
  if (!Array.isArray(target.parent) && !Array.isArray(target.value)) {
    throw patchError(
      command,
      index,
      target.path,
      "invalid-target",
      "Insert target must address a canonical array.",
    );
  }

  if (Array.isArray(target.parent)) {
    const insertionIndex = positionIndex(
      target.parent,
      command.position,
      index,
      command,
      target.path,
    );
    target.parent.splice(insertionIndex, 0, ...command.value);
    return;
  }
  if (!Array.isArray(target.value)) return;
  const insertionIndex = positionIndex(target.value, command.position, index, command, target.path);
  target.value.splice(insertionIndex, 0, ...command.value);
}

function replaceCommand(
  options: object,
  command: Extract<RuntimePatchCommand, { type: "replace" }>,
  index: number,
  delimiters: PlaceholderDelimiters,
): void {
  const target = resolveTarget(options, command.at, index, command.type, true, true);
  if (command.at.text) {
    if (typeof command.value !== "string") {
      throw patchError(
        command,
        index,
        target.path,
        "invalid-operation",
        "Text replace value must be a string.",
      );
    }
    replaceSelectedText(
      target.value,
      command.at.text,
      command.value,
      index,
      command,
      target.path,
      delimiters,
    );
    return;
  }
  setTargetValue(target, command.value);
}

function updateCommand(
  options: object,
  command: Extract<RuntimePatchCommand, { type: "update" }>,
  index: number,
  delimiters: PlaceholderDelimiters,
): void {
  const target = resolveTarget(options, command.at, index, command.type, true, true);
  if (command.at.text) {
    if (typeof command.set !== "string") {
      throw patchError(
        command,
        index,
        target.path,
        "invalid-operation",
        "Text update value must be a string.",
      );
    }
    replaceSelectedText(
      target.value,
      command.at.text,
      command.set,
      index,
      command,
      target.path,
      delimiters,
    );
    return;
  }
  if (command.set !== undefined) {
    if (isRecord(target.value) && isRecord(command.set)) mergeValue(target.value, command.set);
    else setTargetValue(target, command.set);
  }
  if (command.unset) {
    unsetProperties(target.value, command.unset, command, index, target.path);
  }
}

function removeCommand(
  options: object,
  command: Extract<RuntimePatchCommand, { type: "remove" }>,
  index: number,
  delimiters: PlaceholderDelimiters,
): void {
  const target = resolveTarget(options, command.at, index, command.type);
  if (command.at.text) {
    replaceSelectedText(target.value, command.at.text, "", index, command, target.path, delimiters);
    return;
  }
  if (Array.isArray(target.parent) && typeof target.key === "number") {
    target.parent.splice(target.key, 1);
    return;
  }
  Reflect.deleteProperty(target.parent, target.key);
}

function moveCommand(
  options: object,
  command: Extract<RuntimePatchCommand, { type: "move" }>,
  index: number,
  _delimiters: PlaceholderDelimiters,
): void {
  const target = resolveTarget(options, command.at, index, command.type);
  if (!Array.isArray(target.parent)) {
    throw patchError(
      command,
      index,
      target.path,
      "invalid-target",
      "Move target must address an item in a canonical array.",
    );
  }

  if (!Array.isArray(target.parent) || typeof target.key !== "number") {
    throw patchError(
      command,
      index,
      target.path,
      "invalid-target",
      "Move address must select one array item.",
    );
  }
  const targetArray = target.parent;
  const currentIndex = target.key;

  const requestedIndex = positionIndex(targetArray, command.to, index, command, target.path);
  const [item] = targetArray.splice(currentIndex, 1);
  const destinationIndex = currentIndex < requestedIndex ? requestedIndex - 1 : requestedIndex;
  targetArray.splice(destinationIndex, 0, item);
}

function resolveTarget(
  options: object,
  address: PatchAddress,
  commandIndex: number,
  operation: string,
  allowMissingTarget = false,
  allowMissingRoot = false,
): ModelTarget {
  if (!Reflect.has(options, address.field)) {
    if (allowMissingRoot && (address.route ?? []).length === 0) {
      Reflect.set(options, address.field, undefined);
    } else if (allowMissingTarget && (address.route ?? []).length === 0) {
      Reflect.set(options, address.field, []);
    } else {
      throw modelError(
        operation,
        commandIndex,
        address.field,
        "missing-target",
        `Canonical field "${address.field}" does not exist.`,
      );
    }
  }
  if (allowMissingTarget && !Reflect.has(options, address.field)) {
    Reflect.set(options, address.field, undefined);
  }

  let parent: object = options;
  let key: string | number = address.field;
  let value: unknown = Reflect.get(options, address.field);
  let path = address.field;

  const route = address.route ?? [];
  for (const [stepIndex, step] of route.entries()) {
    const allowMissing = allowMissingTarget && stepIndex === route.length - 1;
    const child = selectStep(value, step, commandIndex, operation, path, stepIndex, allowMissing);
    path = `${path}.${stepLabel(step)}`;
    parent = containerFor(value);
    key = child.key;
    value = child.value;
  }

  return { parent, key, value, path };
}

function selectStep(
  value: unknown,
  step: PatchRouteStep,
  commandIndex: number,
  operation: string,
  path: string,
  stepIndex: number,
  allowMissing = false,
): { key: string | number; value: unknown } {
  const stepPath = `${path}[${stepIndex}]`;
  if ("property" in step) {
    if (!isRecord(value) || !Reflect.has(value, step.property)) {
      throw patchError(
        operation,
        commandIndex,
        stepPath,
        "missing-target",
        `Canonical property "${step.property}" does not exist.`,
      );
    }
    return { key: step.property, value: Reflect.get(value, step.property) };
  }
  if ("collection" in step) {
    if (!isRecord(value)) {
      throw patchError(
        operation,
        commandIndex,
        stepPath,
        "invalid-target",
        "Collection route requires a canonical object.",
      );
    }
    if (!Reflect.has(value, step.collection)) {
      if (!allowMissing) {
        throw patchError(
          operation,
          commandIndex,
          stepPath,
          "missing-target",
          `Canonical collection "${step.collection}" does not exist.`,
        );
      }
      Reflect.set(value, step.collection, []);
    }
    const collection = Reflect.get(value, step.collection);
    if (!Array.isArray(collection)) {
      throw patchError(
        operation,
        commandIndex,
        stepPath,
        "invalid-target",
        `Canonical collection "${step.collection}" is not an array.`,
      );
    }
    return { key: step.collection, value: collection };
  }
  if ("index" in step) {
    if (
      !Array.isArray(value) ||
      !Number.isInteger(step.index) ||
      step.index < 0 ||
      step.index >= value.length
    ) {
      throw patchError(
        operation,
        commandIndex,
        stepPath,
        "missing-target",
        "Index route did not match a canonical item.",
      );
    }
    return { key: step.index, value: value[step.index] };
  }
  if (!Array.isArray(value)) {
    throw patchError(
      operation,
      commandIndex,
      stepPath,
      "invalid-target",
      "Selector route requires a canonical array.",
    );
  }

  const matches = value
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => matchesStep(item, step));
  if (matches.length === 0) {
    const selector =
      "name" in step
        ? `name "${step.name}"`
        : "id" in step
          ? `id "${step.id}"`
          : "reference" in step
            ? `reference "${step.reference}"`
            : "address" in step
              ? `address "${step.address}"`
              : "selector";
    throw patchError(
      operation,
      commandIndex,
      stepPath,
      "missing-target",
      `Route selector ${selector} did not match a canonical item.`,
    );
  }
  if (matches.length > 1) {
    throw patchError(
      operation,
      commandIndex,
      stepPath,
      "invalid-target",
      "Route selector matched multiple canonical items.",
    );
  }
  const match = matches[0]!;
  return { key: match.index, value: match.item };
}

function matchesStep(item: unknown, step: PatchRouteStep): boolean {
  if (!isRecord(item)) return false;
  if ("property" in step) return false;
  if ("index" in step || "collection" in step) return false;
  if ("id" in step) return "id" in item && String(item.id) === String(step.id);
  if ("name" in step) return item.name === step.name;
  if ("reference" in step) return item.reference === step.reference;
  return item.address === step.address;
}

function containerFor(value: unknown): object {
  return isRecord(value) || Array.isArray(value) ? value : {};
}

function positionIndex(
  array: readonly unknown[],
  position: PatchPosition = "end",
  commandIndex: number,
  command: RuntimePatchCommand,
  path: string,
): number {
  if (position === "start") return 0;
  if (position === "end") return array.length;
  if (typeof position === "number") {
    if (!Number.isInteger(position) || position < 0 || position > array.length) {
      throw patchError(
        command,
        commandIndex,
        path,
        "invalid-target",
        "Array position must be an in-range integer.",
      );
    }
    return position;
  }

  const route = "before" in position ? position.before : position.after;
  const matches = array
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => matchesStep(item, route));
  if (matches.length !== 1) {
    throw patchError(
      command,
      commandIndex,
      path,
      "missing-target",
      "Position selector must match exactly one canonical item.",
    );
  }
  const matchIndex = matches[0]!.index;
  return "before" in position ? matchIndex : matchIndex + 1;
}

function replaceSelectedText(
  target: unknown,
  selection: PatchTextSelection,
  replacement: string,
  commandIndex: number,
  command: RuntimePatchCommand,
  path: string,
  delimiters: PlaceholderDelimiters,
): void {
  const needle =
    "placeholder" in selection
      ? `${delimiters.start}${selection.placeholder}${delimiters.end}`
      : selection.literal;
  const replaceAll = selection.all === true;
  let replaced = false;

  const hasReplaced = (): boolean => replaced;
  const markReplaced = (): void => {
    replaced = true;
  };
  walkText(target, needle, replacement, markReplaced, hasReplaced, replaceAll);

  if (!replaced) {
    throw patchError(
      command,
      commandIndex,
      path,
      "missing-target",
      "Text selection did not match canonical content.",
    );
  }
}

function walkText(
  value: unknown,
  needle: string,
  replacement: string,
  markReplaced: () => void,
  hasReplaced: () => boolean,
  replaceAll: boolean,
): void {
  if (!replaceAll && hasReplaced()) return;
  if (Array.isArray(value)) {
    replaceAcrossItems(value, needle, replacement, markReplaced, hasReplaced, replaceAll);
    for (const [index, item] of value.entries()) {
      if (!replaceAll && hasReplaced()) return;
      if (typeof item === "string")
        value[index] = replaceSingleText(item, needle, replacement, markReplaced, replaceAll);
      else walkText(item, needle, replacement, markReplaced, hasReplaced, replaceAll);
    }
    return;
  }
  if (typeof value === "string") {
    replaceSingleText(value, needle, replacement, markReplaced, replaceAll);
    return;
  }
  if (!isRecord(value)) return;
  for (const [property, item] of Object.entries(value)) {
    if (!replaceAll && hasReplaced()) return;
    if (typeof item === "string")
      Reflect.set(
        value,
        property,
        replaceSingleText(item, needle, replacement, markReplaced, replaceAll),
      );
    else walkText(item, needle, replacement, markReplaced, hasReplaced, replaceAll);
  }
}

function replaceAcrossItems(
  items: unknown[],
  needle: string,
  replacement: string,
  markReplaced: () => void,
  hasReplaced: () => boolean,
  replaceAll: boolean,
): void {
  const leaves: PatchTextLeaf[] = [];
  for (const [index, item] of items.entries()) collectTextLeaves(item, items, index, leaves);
  const joined = leaves.map((leaf) => leaf.value).join("");
  const starts: number[] = [];
  let offset = 0;
  for (const leaf of leaves) {
    starts.push(offset);
    offset += leaf.value.length;
  }
  let cursor = 0;
  while (cursor <= offset - needle.length && (replaceAll || !hasReplaced())) {
    const start = joined.indexOf(needle, cursor);
    if (start < 0) return;
    const end = start + needle.length;
    const firstIndex = leaves.findIndex(
      (leaf, index) => starts[index]! <= start && start < starts[index]! + leaf.value.length,
    );
    const lastIndex = leaves.findIndex(
      (leaf, index) => starts[index]! < end && end <= starts[index]! + leaf.value.length,
    );
    if (firstIndex < 0 || lastIndex < 0) {
      cursor = start + 1;
      continue;
    }
    const first = leaves[firstIndex]!;
    const last = leaves[lastIndex]!;
    const before = first.value.slice(0, start - starts[firstIndex]!);
    const after = last.value.slice(end - starts[lastIndex]!);
    const updated = `${before}${replacement}${after}`;
    writeTextLeaf(first, updated);
    for (let index = firstIndex + 1; index <= lastIndex; index += 1)
      writeTextLeaf(leaves[index]!, "");
    markReplaced();
    if (!replaceAll) return;
    cursor = end;
  }
}

function replaceSingleText(
  input: string,
  needle: string,
  replacement: string,
  markReplaced: () => void,
  replaceAll: boolean,
): string {
  if (!input.includes(needle)) return input;
  markReplaced();
  return replaceAll ? input.replaceAll(needle, replacement) : input.replace(needle, replacement);
}

interface PatchTextLeaf {
  readonly array: unknown[];
  readonly index: number;
  readonly property?: string;
  value: string;
}

function collectTextLeaves(
  value: unknown,
  array: unknown[],
  index: number,
  leaves: PatchTextLeaf[],
): void {
  if (typeof value === "string") {
    leaves.push({ array, index, value });
    return;
  }
  if (!isRecord(value)) return;
  if (typeof value.text === "string") {
    leaves.push({ array, index, property: "text", value: value.text });
    return;
  }
  if (Array.isArray(value.children)) {
    for (const [childIndex, child] of value.children.entries())
      collectTextLeaves(child, value.children, childIndex, leaves);
  }
}

function writeTextLeaf(leaf: PatchTextLeaf, value: string): void {
  leaf.value = value;
  if (leaf.property === undefined) leaf.array[leaf.index] = value;
  else {
    const container = leaf.array[leaf.index];
    if (isRecord(container)) Reflect.set(container, leaf.property, value);
  }
}

function mergeValue(target: unknown, patch: unknown): void {
  if (Array.isArray(patch) || !isRecord(patch) || !isRecord(target)) {
    throw new Error("mergeValue requires object targets");
  }
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    const current = Reflect.get(target, key);
    if (isRecord(current) && isRecord(value)) mergeValue(current, value);
    else Reflect.set(target, key, value);
  }
}

function seedMissingRoot<TOptions extends object>(
  options: TOptions,
  command: Extract<RuntimePatchCommand, { type: "update" }>,
): void {
  if (command.at.route?.length || command.set === undefined) return;
  if (Reflect.has(options, command.at.field)) return;

  const seed = Array.isArray(command.set) ? [] : isRecord(command.set) ? {} : undefined;
  Reflect.set(options, command.at.field, seed);
}

function unsetProperties(
  target: unknown,
  paths: PatchPropertyPath,
  command: RuntimePatchCommand,
  commandIndex: number,
  path: string,
): void {
  if (paths.length === 0) {
    throw patchError(
      command,
      commandIndex,
      path,
      "invalid-operation",
      "Unset paths cannot be empty.",
    );
  }
  let parent: unknown = target;
  for (const step of paths.slice(0, -1)) {
    if (!isRecord(parent) && !Array.isArray(parent)) {
      throw patchError(
        command,
        commandIndex,
        path,
        "missing-target",
        "Unset parent is not a canonical object.",
      );
    }
    parent = modelChild(parent, step);
  }
  const last = paths.at(-1)!;
  if (!isRecord(parent) && !Array.isArray(parent)) {
    throw patchError(
      command,
      commandIndex,
      path,
      "missing-target",
      "Unset parent is not a canonical object.",
    );
  }
  if (!Reflect.has(parent, last)) {
    throw patchError(
      command,
      commandIndex,
      path,
      "missing-target",
      "Unset property does not exist.",
    );
  }
  Reflect.deleteProperty(parent, last);
}

function modelChild(parent: object, key: PatchPathStep): unknown {
  if (Array.isArray(parent) && typeof key === "number") return parent[key];
  return Reflect.get(parent, key);
}

function setTargetValue(target: ModelTarget, value: unknown): void {
  if (Array.isArray(target.parent) && typeof target.key === "number") {
    target.parent[target.key] = value;
    return;
  }
  Reflect.set(target.parent, target.key, value);
}

function stepLabel(step: PatchRouteStep): string {
  if ("property" in step) return step.property;
  if ("collection" in step) return step.collection;
  if ("index" in step) return String(step.index);
  if ("id" in step) return `id:${step.id}`;
  if ("name" in step) return `name:${step.name}`;
  if ("reference" in step) return `reference:${step.reference}`;
  return `address:${step.address}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function addressPath(address: PatchAddress): string {
  return address.route?.length
    ? `${address.field}.${address.route.map(stepLabel).join(".")}`
    : address.field;
}

function modelError(
  operation: string,
  index: number,
  path: string,
  reason: PatchDiagnostic["reason"],
  message: string,
): PatchError {
  return new PatchError({
    message,
    operation,
    part: "[canonical-model]",
    path: `commands[${index}].at.${path}`,
    reason,
  });
}

function patchError(
  command: RuntimePatchCommand | string,
  index: number,
  path: string,
  reason: PatchDiagnostic["reason"],
  message: string,
): PatchError {
  return new PatchError({
    message,
    operation: typeof command === "string" ? command : command.type,
    part: "[canonical-model]",
    path: `commands[${index}].at.${path}`,
    reason,
  });
}
