<script setup lang="ts">
// Site override of the theme's PageSection: adds a `features` slot that
// mirrors how the previous (Docus) landing composed — inside the copy
// pane ahead of the links on horizontal sections, full width as a grid
// on vertical ones.
defineProps<{
  orientation?: "vertical" | "horizontal";
  reverse?: boolean;
}>();
</script>

<template>
  <section
    class="bs-docs-section"
    :class="{
      'bs-docs-section--horizontal': orientation === 'horizontal',
      'bs-docs-section--reverse': reverse,
    }"
  >
    <div class="bs-docs-section-main">
      <header v-if="$slots.title || $slots.description" class="bs-docs-section-head">
        <h2 v-if="$slots.title">
          <slot name="title" />
        </h2>
        <p v-if="$slots.description">
          <slot name="description" />
        </p>
        <div
          v-if="orientation === 'horizontal' && $slots.features"
          class="bs-docs-section-features"
        >
          <slot name="features" />
        </div>
        <div v-if="$slots.links" class="bs-docs-section-links">
          <slot name="links" />
        </div>
      </header>
      <!-- MDC has no way back to the default slot once a named one opens,
           so prose-level content (code blocks, groups) travels in `#body`.
           Empty slots also render differently on the server (a
           double-comment fragment) and the client (one comment) — guard
           every slot so the element itself is absent when the slot is. -->
      <div v-if="$slots.body" class="bs-docs-section-body">
        <slot name="body" />
      </div>
    </div>
    <!-- Cards ride below the main lane, full width. -->
    <div v-if="$slots.cards" class="bs-docs-cards">
      <slot name="cards" />
    </div>
  </section>
</template>

<style>
/* The copy-pane feature list: a single quiet column between the
   description and the links. */
.bs-docs-section-features {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: var(--bs-space-4);
  margin-block-start: var(--bs-space-6);
}
</style>
