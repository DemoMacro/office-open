import ApiExample from "~/components/content/ApiExample.vue";

export default defineNuxtPlugin((nuxtApp) => {
  nuxtApp.vueApp.component("ApiExample", ApiExample);
});
