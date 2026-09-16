<template>
  <!-- A compact feature item for section feature lanes: the theme's
       card is too heavy a vessel for one-line capability claims, which
       read as a quiet icon-beside-the-text list instead. -->
  <div class="bs-docs-page-feature">
    <div v-if="$slots.icon" class="bs-docs-page-feature-icon">
      <slot name="icon" />
    </div>
    <div class="bs-docs-page-feature-body">
      <div v-if="$slots.title" class="bs-docs-page-feature-title">
        <slot name="title" />
      </div>
      <p v-if="$slots.description" class="bs-docs-page-feature-description">
        <slot name="description" />
      </p>
    </div>
  </div>
</template>

<style scoped>
.bs-docs-page-feature {
  display: flex;
  align-items: flex-start;
  gap: var(--bs-gap-sm);
}
.bs-docs-page-feature-icon {
  display: inline-flex;
  flex-shrink: 0;
  font-size: 1.25rem;
}

.bs-docs-page-feature-icon svg {
  inline-size: 1em;
  block-size: 1em;
}

.bs-docs-page-feature-title {
  font-size: var(--bs-font-size-sm);
  font-weight: var(--bs-font-weight-medium);
  line-height: var(--bs-line-height-snug);
}

.bs-docs-page-feature-description {
  margin: var(--bs-space-1) 0 0;
  font-size: var(--bs-font-size-sm);
  color: var(--bs-color-text-secondary);
}
</style>
