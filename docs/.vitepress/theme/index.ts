// .vitepress/theme/index.ts
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { withBlogTheme } from 'vitepress-plugin-blog'
import Comments from './components/Comments.vue'
import 'vitepress-plugin-blog/style.css'

const blogTheme = withBlogTheme(DefaultTheme)

export default {
  ...blogTheme,
  enhanceApp(ctx) {
    // Call the blog theme's enhanceApp first
    blogTheme.enhanceApp?.(ctx)

    // Then register our Comments component
    ctx.app.component('Comments', Comments)
  }
} satisfies Theme