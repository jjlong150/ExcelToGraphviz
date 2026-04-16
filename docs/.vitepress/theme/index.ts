// .vitepress/theme/index.ts
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { withBlogTheme } from 'vitepress-plugin-blog'
import { h } from 'vue'

import Comments from './components/Comments.vue'

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
      'home-features-after': () => h('div', { class: 'home-extra-content' }, [
        h('div', {
        innerHTML: `
            <h2>Turn Spreadsheets into Beautiful Graphviz Diagrams</h2>
            <p>The <strong>Relationship Visualizer</strong> spreadsheet transforms your Excel tables into clear, professional <strong>Graphviz diagrams</strong> in seconds. Say goodbye to manual drawing tools — simply enter your data as rows (e.g., "A is related to B"), and watch graphs appear automatically.</p>
            <h3>Why Users Love It</h3>
            <p>Whether you're mapping data flows, org charts, timelines, ERDs, or circuits, this tool makes complex relationships instantly understandable.</p>
            <table class="features-table">
              <tr><td><strong>Draws as you type</strong></td><td>Live Graphviz rendering as data changes</td></tr>
              <tr><td><strong>Powerful styling</strong></td><td>Colors, shapes, fonts, arrows, reusable styles</td></tr>
              <tr><td><strong>Advanced features</strong></td><td>SQL queries, SVG animation, DOT preview, JSON exchange</td></tr>
              <tr><td><strong>Cross-platform</strong></td><td>Works on Windows and macOS</td></tr>
              <tr><td><strong>Sleek UI</strong></td><td>Custom Excel ribbon tabs across all worksheets</td></tr>
              <tr><td><strong>Multilingual</strong></td><td>English · French · German · Italian · Polish</td></tr>
              <tr><td><strong>Absolutely Free</strong></td><td>Free to use · No license required · Donations appreciated</td></tr>
              <tr><td><strong>Open Source</strong></td><td>MIT License</td></tr>
              <tr><td><strong>Rich Code Documentaion</strong></td><td><a href="https://deepwiki.com/jjlong150/ExcelToGraphviz" target="_blank" rel="noopener">AI-powered DeepWiki </a></td></tr>
              <tr><td><strong>Show Your Support</strong></td><td><a href="https://www.buymeacoffee.com/exceltographviz" target="_blank" rel="noopener">Buy Me a Coffee!</a></td></tr>              
              <tr><td><strong>Award Winning</strong><br></td><td>SourceForge Community Choice Award</td></tr>
            </table> 
            <center>      
              <img src="/sourceforge-community-choice.png" alt="SourceForge Community Choice Award" width="90" style="margin-top: 8px;">
            </center>
          `
        })
      ])
    })
  },

  enhanceApp(ctx) {
    // Let the blog theme do its thing first
    blogTheme.enhanceApp?.(ctx)

    // Register your custom component
    ctx.app.component('Comments', Comments)
  }
} satisfies Theme