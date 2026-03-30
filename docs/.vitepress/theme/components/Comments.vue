<template>
  <div id="giscus-container" class="comments-wrapper"></div>
</template>

<script setup lang="ts">
import { onMounted } from 'vue'

onMounted(() => {
  // Prevent duplicate instances
  document.querySelectorAll('.giscus').forEach(el => el.remove())

  const script = document.createElement('script')
  script.src = 'https://giscus.app/client.js'
  script.setAttribute('data-repo', 'jjlong150/excel-to-graphviz-comments')
  script.setAttribute('data-repo-id', 'R_kgDORypoLg')
  script.setAttribute('data-category', 'Blog Comments')
  script.setAttribute('data-category-id', 'DIC_kwDORypoLs4C5a_C')
  script.setAttribute('data-mapping', 'pathname')
  script.setAttribute('data-strict', '0')
  script.setAttribute('data-reactions-enabled', '1')
  script.setAttribute('data-emit-metadata', '0')
  script.setAttribute('data-input-position', 'bottom')
  script.setAttribute('data-theme', 'preferred_color_scheme')
  script.setAttribute('data-lang', 'en')
  script.setAttribute('crossorigin', 'anonymous')
  script.async = true

  const container = document.getElementById('giscus-container')
  if (container) {
    container.innerHTML = ''
    container.appendChild(script)
  }
})
</script>

<style scoped>
.comments-wrapper {
  margin-top: 2rem;
  padding-top: 1.5rem;
  border-top: 1px solid var(--vp-c-divider);
}
</style>