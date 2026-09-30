import type { NuxtApp } from "nuxt/app";

import ApiExample from "~/components/content/ApiExample.vue";

export default defineNuxtPlugin((nuxtApp: NuxtApp): void => {
  nuxtApp.vueApp.component("ApiExample", ApiExample);
});
