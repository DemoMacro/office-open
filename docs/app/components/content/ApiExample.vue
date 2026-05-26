<script lang="ts">
import type { VNode } from "vue";
</script>

<script setup lang="ts">
import { TabsRoot, TabsList, TabsIndicator, TabsTrigger, TabsContent } from "reka-ui";
import { computed, ref, onBeforeUpdate } from "vue";

const props = withDefaults(
  defineProps<{
    type?: "docx" | "pptx";
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
  return {
    label: slot.props?.filename || slot.props?.label || `${index}`,
    icon: slot.props?.icon,
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
      const { Document, Packer } = await import("@office-open/docx");
      const options = data.sections ? data : { sections: Array.isArray(data) ? data : [data] };
      const doc = new Document(options);
      blob = await Packer.toBlob(doc);
      filename = "document.docx";
    } else {
      const { Presentation, Packer } = await import("@office-open/pptx");
      const options = data.slides ? data : { slides: Array.isArray(data) ? data : [data] };
      const pres = new Presentation(options);
      blob = await Packer.toBlob(pres);
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
  <TabsRoot
    v-model="model"
    :default-value="defaultValue"
    :unmount-on-hide="false"
    class="group relative my-5 *:not-first:static! *:not-first:my-0!"
  >
    <TabsList
      class="border-muted bg-default relative flex items-center gap-1 overflow-x-auto rounded-t-md border border-b-0 p-2"
    >
      <TabsIndicator
        class="bg-elevated absolute inset-y-2 left-0 w-(--reka-tabs-indicator-size) translate-x-(--reka-tabs-indicator-position) rounded-md shadow-xs transition-[translate,width] duration-200"
      />

      <TabsTrigger
        v-for="(item, index) of items"
        :key="index"
        :value="String(index)"
        class="text-default data-[state=active]:text-highlighted hover:bg-elevated/50 focus-visible:ring-primary relative inline-flex items-center gap-1.5 rounded-md px-2 py-1.5 text-sm transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-inset disabled:cursor-not-allowed disabled:opacity-75"
      >
        <ProseCodeIcon :icon="item.icon" :filename="item.label" class="size-4 shrink-0" />
        <span class="truncate">{{ item.label }}</span>
      </TabsTrigger>
    </TabsList>

    <TabsContent v-for="(item, index) of items" :key="index" :value="String(index)" as-child>
      <component :is="item.component" hide-header tabindex="-1" />
    </TabsContent>

    <UButton
      v-if="showExport"
      :disabled="exporting"
      :loading="exporting"
      icon="i-lucide-download"
      size="sm"
      color="neutral"
      variant="outline"
      style="position: absolute !important"
      class="ring-accented hover:bg-elevated top-[11px] right-[44px] z-10 p-1.5 ring transition ring-inset lg:opacity-0 lg:group-hover:opacity-100"
      @click="handleExport"
    />
  </TabsRoot>
</template>
