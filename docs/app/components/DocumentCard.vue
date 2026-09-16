<script setup lang="ts">
import { Button } from "@bysages/vue";

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
  <div class="document-card">
    <div class="document-card-icon">
      <Icon :name="getIcon(mimeType)" />
    </div>
    <div class="document-card-meta">
      <p class="document-card-name">{{ filename }}</p>
      <p class="document-card-size">{{ formatSize(size) }}</p>
    </div>
    <Button variant="ghost" size="sm" aria-label="Download" @click="download">
      <Icon name="i-lucide-download" />
    </Button>
  </div>
</template>

<style scoped>
.document-card {
  display: flex;
  align-items: center;
  gap: var(--bs-space-3);
  padding: var(--bs-space-2);
}

.document-card-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  width: var(--bs-space-10);
  height: var(--bs-space-10);
  border: 1px solid var(--bs-color-border);
  border-radius: var(--bs-radius-sm);
  background: var(--bs-color-surface-2);
  font-size: var(--bs-font-size-lg);
}

.document-card-meta {
  flex: 1;
  min-width: 0;
}

.document-card-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: var(--bs-font-size-sm);
  font-weight: 500;
  color: var(--bs-color-text-primary);
}

.document-card-size {
  font-size: var(--bs-font-size-xs);
  color: var(--bs-color-text-secondary);
}
</style>
