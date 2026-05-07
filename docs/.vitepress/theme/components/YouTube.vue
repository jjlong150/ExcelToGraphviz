<template>
  <div class="vp-youtube">
    <iframe
      :src="computedSrc"
      title="YouTube video player"
      frameborder="0"
      allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
      allowfullscreen
    ></iframe>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  id: { type: String, required: true },
  start: { type: [String, Number], default: null },
  end: { type: [String, Number], default: null }
})

const computedSrc = computed(() => {
  const base = `https://www.youtube.com/embed/${props.id}`
  const params = new URLSearchParams()

  if (props.start) params.set('start', props.start)
  if (props.end) params.set('end', props.end)

  const query = params.toString()
  return query ? `${base}?${query}` : base
})
</script>

<style scoped>
.vp-youtube {
  position: relative;
  width: 100%;
  padding-bottom: 56.25%; /* 16:9 */
  height: 0;
  overflow: hidden;
  border-radius: 8px;
}

.vp-youtube iframe {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
}
</style>
