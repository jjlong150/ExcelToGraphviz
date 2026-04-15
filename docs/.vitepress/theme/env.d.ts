// .vitepress/theme/env.d.ts
/// <reference types="vite/client" />

// Allow side-effect CSS imports (no default export)
declare module '*.css' {
  const content: string
  export default content
}

// Specifically for the blog plugin (in case the generic one doesn't catch it)
declare module 'vitepress-plugin-blog/style.css' {
  const content: string
  export default content
}