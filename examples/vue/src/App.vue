<script setup lang="ts">
import { onBeforeUnmount, ref, shallowRef } from "vue";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

const previewHost = ref<HTMLDivElement | null>(null);
const input = ref<HTMLInputElement | null>(null);
const previewer = shallowRef<FilePreviewer | null>(null);
const fileName = ref("");
const busy = ref(false);
const error = ref("");
const dragging = ref(false);

function errorText(reason: unknown): string {
  return reason instanceof Error ? reason.message : "The selected file could not be previewed.";
}

async function previewFile(file: File): Promise<void> {
  const container = previewHost.value;
  if (!container) return;
  previewer.value?.destroy();
  previewer.value = null;
  container.replaceChildren();
  fileName.value = file.name;
  error.value = "";
  busy.value = true;
  const instance = new FilePreviewer({
    container,
    file: { name: file.name, source: file, mimeType: file.type },
  });
  previewer.value = instance;
  try {
    await instance.open();
  } catch (reason: unknown) {
    if (previewer.value === instance) error.value = errorText(reason);
  } finally {
    if (previewer.value === instance) busy.value = false;
  }
}

function onFileChange(event: Event): void {
  const target = event.currentTarget;
  if (!(target instanceof HTMLInputElement)) return;
  const file = target.files?.[0];
  if (file) void previewFile(file);
  target.value = "";
}

function onDrop(event: DragEvent): void {
  event.preventDefault();
  dragging.value = false;
  const file = event.dataTransfer?.files[0];
  if (file) void previewFile(file);
}

function onDragLeave(event: DragEvent): void {
  const current = event.currentTarget;
  if (!(event.relatedTarget instanceof Node) || !(current instanceof Node) || !current.contains(event.relatedTarget)) dragging.value = false;
}

function clearFile(): void {
  previewer.value?.destroy();
  previewer.value = null;
  previewHost.value?.replaceChildren();
  fileName.value = "";
  error.value = "";
  busy.value = false;
  if (input.value) input.value.value = "";
}

onBeforeUnmount(() => previewer.value?.destroy());
</script>

<template>
  <main class="example-shell">
    <header class="example-header">
      <a class="back-link" href="../">← All framework examples</a>
      <p class="eyebrow">Framework integration</p>
      <h1 class="typography-eyebrow-super">Vue file preview</h1>
      <p class="intro">Preview supported images and Office Open XML files in your Vue app. Files are read locally in your browser and are never uploaded by this demo.</p>
    </header>

    <section
      class="upload-card"
      :class="{ 'is-dragging': dragging }"
      @dragenter.prevent="dragging = true"
      @dragover.prevent
      @dragleave="onDragLeave"
      @drop="onDrop"
    >
      <div class="upload-copy">
        <strong>{{ fileName || "Choose a file to preview" }}</strong>
        <span>{{ busy ? "Preparing local preview…" : "Drop a file here, or browse your device." }}</span>
      </div>
      <div class="upload-actions">
        <label class="button button-primary">
          Choose file
          <input ref="input" type="file" accept="image/*,.svg,.docx,.xlsx,.pptx" @change="onFileChange" />
        </label>
        <button v-if="fileName" class="button button-secondary" type="button" @click="clearFile">Clear</button>
      </div>
    </section>

    <p v-if="error" class="error-message" role="alert">{{ error }}</p>
    <section class="preview-card" aria-label="File preview">
      <p v-if="busy" class="loading-message" role="status">Loading preview…</p>
      <div v-if="!fileName" class="empty-state">
        <span class="file-mark" aria-hidden="true">＋</span>
        <strong>Your preview will appear here</strong>
        <span>Try PNG, JPEG, GIF, WebP, SVG, DOCX, XLSX or PPTX.</span>
      </div>
      <div ref="previewHost" class="preview-host" />
    </section>
    <footer class="example-footer">
      <span>The preview toolbar includes zoom and original-file download controls.</span>
      <a href="https://www.npmjs.com/package/@dukeofwest/universal-file-preview">npm package</a>
    </footer>
  </main>
</template>
