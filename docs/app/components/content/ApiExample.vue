<script lang="ts">
import type { VNode } from "vue";
</script>

<script setup lang="ts">
import { Tabs, Button } from "@bysages/vue";
import { computed, ref, onBeforeUpdate } from "vue";

// Code-block language → icon, consumed both here and by the client bundle
// scanner (nuxt.config collectIconNames matches these literals in this file).
const CODE_ICONS: Record<string, string> = {
  docx: "i-vscode-icons-file-type-word",
  pptx: "i-vscode-icons-file-type-powerpoint",
  xlsx: "i-vscode-icons-file-type-excel",
  json: "i-vscode-icons-file-type-json",
  jsonc: "i-vscode-icons-file-type-json",
  html: "i-vscode-icons-file-type-html",
  htm: "i-vscode-icons-file-type-html",
  xml: "i-vscode-icons-file-type-xml",
  markdown: "i-vscode-icons-file-type-markdown",
  md: "i-vscode-icons-file-type-markdown",
  mdc: "i-vscode-icons-file-type-markdown",
  css: "i-vscode-icons-file-type-css",
  scss: "i-vscode-icons-file-type-css",
  less: "i-vscode-icons-file-type-css",
  yaml: "i-vscode-icons-file-type-yaml",
  yml: "i-vscode-icons-file-type-yaml",
  toml: "i-vscode-icons-file-type-toml",
  js: "i-vscode-icons-file-type-js",
  javascript: "i-vscode-icons-file-type-js",
  mjs: "i-vscode-icons-file-type-js",
  cjs: "i-vscode-icons-file-type-js",
  jsx: "i-vscode-icons-file-type-js",
  ts: "i-vscode-icons-file-type-typescript",
  typescript: "i-vscode-icons-file-type-typescript",
  tsx: "i-vscode-icons-file-type-typescript",
  vue: "i-vscode-icons-file-type-vue",
  bash: "i-vscode-icons-file-type-shell",
  sh: "i-vscode-icons-file-type-shell",
  shell: "i-vscode-icons-file-type-shell",
  zsh: "i-vscode-icons-file-type-shell",
  batch: "i-vscode-icons-file-type-shell",
  pnpm: "i-simple-icons-pnpm",
  npm: "i-simple-icons-npm",
  yarn: "i-simple-icons-yarn",
  bun: "i-simple-icons-bun",
  python: "i-vscode-icons-file-type-python",
  py: "i-vscode-icons-file-type-python",
  go: "i-vscode-icons-file-type-go",
  rust: "i-vscode-icons-file-type-rust",
  rs: "i-vscode-icons-file-type-rust",
  ruby: "i-vscode-icons-file-type-ruby",
  rb: "i-vscode-icons-file-type-ruby",
  java: "i-vscode-icons-file-type-java",
  cs: "i-vscode-icons-file-type-csharp",
  csharp: "i-vscode-icons-file-type-csharp",
  kotlin: "i-vscode-icons-file-type-kotlin",
  kt: "i-vscode-icons-file-type-kotlin",
  php: "i-vscode-icons-file-type-php",
  sql: "i-vscode-icons-file-type-sql",
  docker: "i-vscode-icons-file-type-docker",
  dockerfile: "i-vscode-icons-file-type-docker",
  text: "i-vscode-icons-file-type-text",
  plaintext: "i-vscode-icons-file-type-text",
};

const props = withDefaults(
  defineProps<{
    type?: "docx" | "pptx" | "xlsx";
    defaultValue?: string;
  }>(),
  { defaultValue: "0" },
);

const slots = defineSlots<{ default(props?: {}): VNode[] }>();
const model = defineModel<string>();
const exporting = ref(false);
const rerenderCount = ref(0);

const items = computed(() => {
  rerenderCount.value;
  return slots.default?.()?.flatMap(transformSlot).filter(Boolean) || [];
});

function transformSlot(slot: any, index: number): any {
  if (typeof slot.type === "symbol") {
    return slot.children?.map(transformSlot);
  }
  const label: string = slot.props?.filename || slot.props?.language || `${index}`;
  const dot = label.lastIndexOf(".");
  // Labels come in any case ("JSON", "DOCX"); the table's keys are lower.
  return {
    label,
    icon:
      slot.props?.icon ??
      CODE_ICONS[(dot === -1 ? label : label.slice(dot + 1)).toLowerCase()] ??
      CODE_ICONS[label.toLowerCase()],
    component: slot,
    code: slot.props?.code || "",
  };
}

onBeforeUpdate(() => rerenderCount.value++);

function parseExportable(code: string) {
  try {
    const data = JSON.parse(code);
    if (data.sections) return { data, type: "docx" as const };
    if (data.slides) return { data, type: "pptx" as const };
    if (data.worksheets) return { data, type: "xlsx" as const };
  } catch {}
  return undefined;
}

const activeItem = computed(() => {
  const active = model.value ?? props.defaultValue ?? "0";
  return items.value[Number(active)];
});

const showExport = computed(() => !!activeItem.value && !!parseExportable(activeItem.value.code));

async function handleExport() {
  if (!activeItem.value) return;
  const parsed = parseExportable(activeItem.value.code);
  if (!parsed) return;
  exporting.value = true;
  try {
    const { data, type } = parsed;
    const isDocx = props.type === "docx" || type === "docx";
    let blob: Blob;
    let filename: string;

    if (isDocx) {
      const { generateDocument } = await import("office-open/docx");
      const options = data.sections ? data : { sections: Array.isArray(data) ? data : [data] };
      blob = await generateDocument(options, { type: "blob" });
      filename = "document.docx";
    } else if (parsed.type === "xlsx" || props.type === "xlsx") {
      const { generateWorkbook } = await import("office-open/xlsx");
      const options = data.worksheets ? data : { worksheets: Array.isArray(data) ? data : [data] };
      blob = await generateWorkbook(options, { type: "blob" });
      filename = "workbook.xlsx";
    } else {
      const { generatePresentation } = await import("office-open/pptx");
      const options = data.slides ? data : { slides: Array.isArray(data) ? data : [data] };
      blob = await generatePresentation(options, { type: "blob" });
      filename = "presentation.pptx";
    }

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e) {
    console.error("Export failed:", e);
  } finally {
    exporting.value = false;
  }
}
</script>

<template>
  <Tabs.Root v-model="model" :default-value="defaultValue" class="bs-docs-code-group api-example">
    <Tabs.List>
      <Tabs.Trigger v-for="(item, index) of items" :key="index" :value="String(index)">
        <Icon v-if="item.icon" :name="item.icon" class="api-example-icon" />
        <span class="api-example-label">{{ item.label }}</span>
      </Tabs.Trigger>
      <Tabs.Indicator />
    </Tabs.List>

    <Tabs.Content v-for="(item, index) of items" :key="index" :value="String(index)">
      <component :is="item.component" tabindex="-1" />
    </Tabs.Content>

    <!-- Copy rides the pane's own pre bar (the site css re-shows it inside
         code groups); download pins the tab strip's free end. -->
    <div v-if="showExport" class="api-example-actions">
      <Button
        variant="ghost"
        size="sm"
        square
        :disabled="exporting"
        aria-label="Download the generated document"
        @click="handleExport"
      >
        <Icon name="i-lucide-download" />
      </Button>
    </div>
  </Tabs.Root>
</template>

<style scoped>
.api-example {
  position: relative;
}

.api-example-icon {
  width: 1em;
  height: 1em;
  flex-shrink: 0;
}

.api-example-label {
  max-width: 14rem;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* Download rides the tab strip's free right end, always visible — the
   strip is a pure tab row otherwise and the right end stays empty. */
.api-example-actions {
  position: absolute;
  inset-block-start: 0;
  inset-inline-end: var(--bs-space-2);
  z-index: 1;
  display: flex;
  align-items: center;
  height: var(--bs-control-height-sm);
}
</style>
