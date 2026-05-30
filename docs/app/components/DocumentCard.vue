<script setup lang="ts">
interface Props {
  filename: string;
  base64: string;
  mimeType: string;
  size: number;
}

const props = defineProps<Props>();

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getIcon(mimeType: string): string {
  if (mimeType.includes("wordprocessingml")) return "i-lucide-file-text";
  if (mimeType.includes("presentationml")) return "i-lucide-presentation";
  if (mimeType.includes("spreadsheetml")) return "i-lucide-table";
  return "i-lucide-file";
}

function download() {
  const binaryString = atob(props.base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  const blob = new Blob([bytes], { type: props.mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = props.filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
</script>

<template>
  <div class="flex items-center gap-3 p-2">
    <div class="bg-primary/10 flex size-10 shrink-0 items-center justify-center rounded-lg">
      <UIcon :name="getIcon(mimeType)" class="text-primary size-5" />
    </div>
    <div class="min-w-0 flex-1">
      <p class="text-highlighted truncate text-sm font-medium">
        {{ filename }}
      </p>
      <p class="text-muted text-xs">
        {{ formatSize(size) }}
      </p>
    </div>
    <UButton icon="i-lucide-download" color="primary" variant="soft" size="xs" @click="download" />
  </div>
</template>
