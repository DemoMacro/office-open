import type { NuxtApp } from "nuxt/app";

import ApiExample from "~/components/content/ApiExample.vue";

export default defineNuxtPlugin((nuxtApp: NuxtApp) => {
  nuxtApp.vueApp.component("ApiExample", ApiExample);
});
