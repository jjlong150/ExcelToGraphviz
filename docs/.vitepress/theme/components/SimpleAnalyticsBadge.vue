<template>
  <!--
    Simple Analytics badge. The free plan requires it to be visible in the site footer.
    It is placed in the layout-bottom slot because the default theme footer is hidden on
    pages that have a sidebar. The markup is rendered at build time so that it is in the HTML.
    Embed code from https://www.simpleanalytics.com/badges
  -->
  <div class="sa-badge">
    <a
      href="https://www.simpleanalytics.com?utm_source=exceltographviz.com&utm_content=badge"
      referrerpolicy="origin"
      target="_blank"
    >
      <img
        src="https://simpleanalyticsbadges.com/exceltographviz.com?background=edf3f4&text=677779&logo=ff4f64&radius=3&counter=false"
        width="210"
        height="50"
        alt="Simple Analytics"
        loading="lazy"
        referrerpolicy="no-referrer"
        crossorigin="anonymous"
      />
    </a>
  </div>
</template>

<style scoped>
.sa-badge {
  display: flex;
  justify-content: center;
  padding: 16px 24px 32px;
  background-color: var(--vp-c-bg);
}

.sa-badge a {
  display: inline-block;
  line-height: 0;
}

@media print {
  .sa-badge {
    display: none;
  }
}
</style>
