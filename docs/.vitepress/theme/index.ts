// .vitepress/theme/index.ts
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { withBlogTheme } from 'vitepress-plugin-blog'
import { h } from 'vue'

import Comments from './components/Comments.vue'
import HomeExtra from './components/HomeExtra.vue'
import YouTube from './components/YouTube.vue'

// Import the blog plugin styles
import 'vitepress-plugin-blog/style.css'

// Import our own styles (make sure this is after the blog styles so it can override them if needed)
import './style.css'

const blogTheme = withBlogTheme(DefaultTheme)

export default {
  extends: blogTheme,   // Use 'extends' – this is the most compatible way with the blog plugin

  Layout() {
    return h(blogTheme.Layout!, null, {
      // Only inject extra content on the home page (after the 6 features)
      'home-features-after': () => h(HomeExtra)
    })
  },

  enhanceApp(ctx) {
    // Let the blog theme do its thing first
    blogTheme.enhanceApp?.(ctx)

    // Register your custom component
    ctx.app.component('Comments', Comments)
    ctx.app.component('HomeExtra', HomeExtra)
    ctx.app.component('YouTube', YouTube)
  }
} satisfies Theme