/** One keyboard mapping (CT_Keymap). */
export interface KeyMapOptions {
  /** Keyboard mask (`wne:mask`; `true` serializes as `1`). */
  mask?: boolean;
  /** Primary key-code sequence, e.g. `"0249"` (`wne:kcmPrimary`). */
  primaryKeyCode?: string;
  /** Secondary key-code sequence (`wne:kcmSecondary`). */
  secondaryKeyCode?: string;
  /** Bound built-in command (`wne:acd/@wne:acdName`). */
  actionName?: string;
  /** Bound character code (`wne:wch/@wne:val`). */
  characterCode?: string;
  /** Bound macro name (`wne:macro/@wne:macroName`). */
  macroName?: string;
}

/** One toolbar action definition (CT_Acd). */
export interface ActionDefinitionOptions {
  /** Stable action identifier (`wne:acdName`). */
  name: string;
  /** Built-in command this definition is based on (`wne:fciBasedOn`). */
  basedOn?: string;
  /** Base64 command argument (`wne:argValue`). */
  argumentValue?: string;
  /** Hex index selector (`wne:fciIndexBasedOn`). */
  indexBasedOn?: string;
}

/** Word keyboard and toolbar customizations (`wne:tcg`). */
export interface WordCustomizationsOptions {
  /** Keyboard mappings from `wne:keymaps`. */
  keyMaps?: KeyMapOptions[];
  /** Toolbar-visible action names from `wne:acdManifest`. */
  toolbarActionNames?: string[];
  /** Command definitions from `wne:acds`. */
  actionDefinitions?: ActionDefinitionOptions[];
  /** Attached toolbar binary relationship (`wne:toolbarData/@r:id`). */
  toolbarDataRelationshipId?: string;
}
