import { defineConfig } from 'vitepress'
import { blogPlugin } from 'vitepress-plugin-blog/plugin'

// https://vitepress.dev/reference/site-config
export default defineConfig({
  // VitePress plugins are configured under the `vite` property
  // Here we add the blog plugin to enable blogging features in our VitePress site
  // The blog plugin will automatically generate blog pages based on the content in the `blog` directory
  // For more information on how to use the blog plugin, see the documentation:
  // https://vitepress.dev/reference/plugin-blog
  // https://github.com/humanbydefinition/vitepress-plugin-blog

  vite: {
    plugins: [
      blogPlugin({
        postsDir: 'blog/posts',
      })
    ],
    ssr: {
      noExternal: ['vitepress-plugin-blog']
    },
    build: {
      chunkSizeWarningLimit: 1000
    }
  },
  
  head: [
    // Favicon (browser tab icon)
    ['link', { rel: 'icon', href: '/favicon.ico' }],

    // Simple Analytics
    // This script is added to the head of the HTML document to enable Simple Analytics tracking on the site.
    [
      'script',
      {
        async: '',
        src: 'https://scripts.simpleanalyticscdn.com/latest.js'
      }
    ]
  ],
  base: '/',
  lang: 'en-US',
  title: "Excel to Graphviz",
  description: "Excel to Graphviz Relationship Visualizer",
  themeConfig: {
    // https://vitepress.dev/reference/default-theme-config
    nav: [
      { text: 'Blog', link: '/blog/' },
      { text: "Resources",
            items: [
              { text: 'About', link: '/about/' },
              { text: 'License', link: '/license/' },
              { text: 'Privacy', link: '/privacy/' },
              { text: 'Credits', link: '/acknowledge/' },
              { text: 'Changelog', link: '/changelog/' }
            ]
      },
    ],
    
    logo: "/logo.png",

    sidebar: [
      {
        text: 'Getting Started',
        items: [
          { text: 'Overview', link: '/overview/' },
          { text: 'Workbook', link: '/workbook/' },
          { text: 'Launchpad', link: '/launchpad/' },
        ]
      },
      {
        text: 'Graphs',
        items: [
          {
            text: 'Creating Graphs', 
            link: '/create/',
            items: [
              { text: 'Terminology', link: '/terminology/' },
              { text: 'New Workbook', link: '/prepare/' },
              { text: 'Data Worksheet', link: '/dataworksheet/' },
              { text: 'Graphviz Tab', link: '/graphviztab/' },
              { text: 'Core Concepts', link: '/coreconcepts/' }
            ]
          },
          {
            text: 'Styling and Views', 
            link: '/addstyle/',
            items: [
              { text: 'Style Designer', link: '/designer/' },
              { text: 'Style Gallery', link: '/styles/' },
              { text: 'Create Views', link: '/views/' }
            ]
          },
          { text: 'Publishing Graphs', link: '/publish/' },
          { text: 'SVG Post-Processing', link: '/svg/' },
          { text: 'Advanced Graphviz Topics', link: '/advanced/' }
            ]
      },
      {
        text: 'Graphviz',
        items: [
          { text: 'View DOT Source Code', link: '/source/' },
          { text: 'DOT Message Console', link: '/console/' }
        ],
      },
      {
        text: 'Data Manipulation',
        items: [
          { text: 'Using SQL', link: '/sql/' },
          { text: 'SQL to Graph Example', link: '/sql/queries/' },
          { text: 'SQL Extensions', 
            link: '/sql/extensions/',
            items: [
              { text: 'Directives', link: '/sql/directives/' },
              { text: 'Clustering', link: '/sql/clustering/' },
              { text: 'Count Substitution', link: '/sql/counts/' },
              { text: 'Label Splitting', link: '/sql/labelsplit/' },
              { text: 'Chaining Nodes', link: '/sql/chaining/' },
              { text: 'Creating Subgraphs', link: '/sql/subgraphs/' },
              { text: 'Tree Traversal', link: '/sql/recursion/'},
              { text: 'Iteration', link: '/sql/iterate/' },
              { text: 'Enumeration', link: '/sql/enumerate/' } ,
              { text: 'Concatenation', link: '/sql/concatenation/' }
                  ]
          },
          { text: 'Examples',
            items: [
              { text: 'Organization Charts', link: '/sql/orgcharts/' },
              { text: 'Timelines and Roadmaps', link: '/sql/timeline/' }
                    ]
          },
          { text: 'SQL Syntax', link: '/sql/syntax/' },
        ]
      },
      {
        text: 'Data Exchange',
        items: [
          { text: 'Using JSON Files', link: '/exchange/',
            items: [
              { text: 'Export', link: '/exchange/export/' },
              { text: 'Import', link: '/exchange/import/' }
            ]
          },
        ]
      },
      {
        text: 'Setup',
        items: [
          { text: 'Download', link: '/download/' },
          {
            text: 'Install', 
            link: '/install/',
            items: [
              { text: 'Windows Instructions', link: '/install-win/' },
              { text: 'macOS Instructions', link: '/install-mac/' }
            ]
          }
        ]
      },
      {
        text: 'Maintenance',
        items: [
          { text: 'Diagnostics', link: '/diagnostics/' },
          { text: 'Lists', link: '/lists/' },
          { text: 'Settings', link: '/settings/' },
          { text: 'Information', link: '/info/' }
        ],
      }
    ],

    search: {
      provider: 'local'
    },

    footer: {
      message: 'Released under the MIT License.',
      copyright: 'Copyright © 2015-present Jeffrey J. Long.'
    },

    editLink: {
      pattern: 'https://buymeacoffee.com/exceltographviz',
      text: 'Like this tool? Buy me a coffee! ☕'
    },

    socialLinks: [
      { icon: 'github', 
        link: 'https://github.com/jjlong150/ExcelToGraphviz',
        ariaLabel: 'GitHub Repository' },
      { icon: 'linkedin', 
        link: 'https://www.linkedin.com/in/jeffreyjlong/',
        ariaLabel: 'LinkedIn' },
      { icon: 'x', 
        link: 'https://x.com/exceltographviz',
        ariaLabel: 'X/Twitter' },
      { 
        icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24" ><title>Buy Me a Coffee</title><path d="m12,8.07h0s0-.01,0-.01c-.01,0-.03.02-.04.03h.02Z"/><path d="m11.98,20.58h.01s.02-.03.03-.04c-.02,0-.03.02-.05.04Z"/><path d="m12.01,20.7h0s-.03-.02-.05-.02c.01,0,.03.01.04.02Z"/><path d="m14.26,20.04c.23-.21.37-.51.4-.82l.78-8.29c-.35-.12-.7-.2-1.1-.2-.69,0-1.24.24-1.88.51-.72.31-1.53.66-2.59.66-.44,0-.88-.06-1.31-.18l.73,7.5c.03.31.17.61.4.82.23.21.53.33.85.33,0,0,1.04.05,1.38.05.37,0,1.49-.05,1.49-.05.31,0,.62-.12.85-.33Z"/><path d="m12.01,8.05s-.01-.01-.02-.02h.01s0,.02,0,.02Z"/><path d="m18.74,6.79c-.1-.5-.32-.97-.83-1.15-.16-.06-.35-.08-.48-.2-.13-.12-.16-.3-.19-.48-.05-.32-.1-.63-.16-.94-.05-.27-.09-.57-.21-.82-.16-.33-.5-.53-.83-.66-.17-.06-.34-.12-.52-.16-.83-.22-1.71-.3-2.56-.35-1.03-.06-2.06-.04-3.08.05-.76.07-1.57.15-2.29.42-.27.1-.54.21-.74.42-.25.25-.33.64-.15.95.13.22.35.38.58.48.3.13.61.24.94.3.9.2,1.82.28,2.74.31,1.01.04,2.03,0,3.04-.1.25-.03.5-.06.75-.1.29-.04.48-.43.39-.7-.1-.32-.38-.44-.7-.39-.05,0-.09.01-.14.02h-.03c-.11.02-.21.03-.32.04-.22.02-.44.04-.66.06-.49.03-.99.05-1.49.05-.49,0-.97-.01-1.46-.05-.22-.01-.44-.03-.66-.06-.1-.01-.2-.02-.3-.03h-.1s-.02-.02-.02-.02h-.1c-.2-.04-.4-.08-.6-.12-.02,0-.04-.02-.05-.03-.01-.02-.02-.04-.02-.06s0-.04.02-.06c.01-.02.03-.03.05-.03h0c.17-.04.35-.07.52-.1.06,0,.12-.02.18-.03h0c.11,0,.22-.03.33-.04.95-.1,1.9-.13,2.85-.1.46.01.92.04,1.39.09.1.01.2.02.3.03.04,0,.08,0,.11.01h.08c.22.04.44.08.66.13.33.07.75.09.89.45.05.11.07.24.09.36l.03.15s0,0,0,0c.08.36.15.72.23,1.08,0,.03,0,.05,0,.08s-.02.05-.03.07-.04.04-.06.06c-.02.01-.05.02-.08.03h-.05s-.05.01-.05.01c-.15.02-.3.04-.44.05-.29.03-.58.06-.87.09-.58.05-1.16.08-1.74.09-.3,0-.59.01-.89.01-1.18,0-2.36-.07-3.53-.21-.12-.01-.24-.03-.36-.04-.02,0-.11-.01-.13-.02-.08-.01-.16-.02-.24-.04-.27-.04-.54-.09-.81-.13-.33-.05-.64-.03-.94.13-.24.13-.44.34-.56.58-.13.26-.17.55-.22.83-.06.28-.15.59-.11.88.07.63.51,1.14,1.14,1.25.59.11,1.19.19,1.79.27,2.35.29,4.73.32,7.09.1.19-.02.38-.04.58-.06.06,0,.12,0,.18.02s.11.05.15.09c.04.04.08.09.1.15.02.06.03.12.02.18l-.06.58c-.12,1.17-.24,2.35-.36,3.52-.13,1.23-.25,2.46-.38,3.7-.04.35-.07.69-.11,1.04-.03.34-.04.69-.1,1.03-.1.53-.46.86-.99.98-.48.11-.97.17-1.46.17-.55,0-1.09-.02-1.64-.02-.58,0-1.3-.05-1.75-.48-.4-.38-.45-.98-.5-1.49-.07-.68-.14-1.37-.21-2.05l-.4-3.8-.26-2.46s0-.08-.01-.12c-.03-.29-.24-.58-.57-.57-.28.01-.6.25-.57.57l.19,1.82.39,3.77c.11,1.07.22,2.14.33,3.21.02.21.04.41.06.62.12,1.12.98,1.72,2.04,1.89.62.1,1.25.12,1.88.13.81.01,1.62.04,2.41-.1,1.17-.22,2.05-1,2.18-2.21.04-.35.07-.7.11-1.05.12-1.16.24-2.32.36-3.48l.39-3.79.18-1.74c0-.09.05-.17.1-.23s.14-.11.22-.12c.34-.07.66-.18.89-.43.38-.41.46-.94.32-1.47l-.11-.56Zm-12.48,1.18h0s.01,0,.02.02c-.01-.01-.02-.02-.02-.02Zm.1.09h0s0,0,0,0c0,0,0,0,0,0Zm11.26-.08h0c-.12.11-.3.17-.48.19-2.01.3-4.05.45-6.09.38-1.46-.05-2.9-.21-4.34-.42-.14-.02-.29-.05-.39-.15-.06-.06-.08-.14-.1-.23,0,0,0,0,0,.01,0,0,0,0,0-.01-.03-.2.02-.43.05-.6.04-.22.13-.51.39-.54.4-.05.87.12,1.27.18.48.07.96.13,1.45.18,2.07.19,4.17.16,6.23-.12.38-.05.75-.11,1.12-.18.33-.06.7-.17.9.17.14.23.16.55.13.81,0,.12-.06.22-.14.3Z"/><path d="m6.22,7.77s0,0,0,0c0-.02.01-.05,0-.05,0,0,0,.03,0,.05Z"/></svg>' },
        link: 'https://buymeacoffee.com/exceltographviz',
        ariaLabel: 'Buy Me a Coffee' }
    ]
  },

  transformPageData(pageData) {
    // Hide sidebar-based Prev/Next links on blog posts only
    const isBlogPost = 
      pageData.frontmatter?.blogPost === true ||
      pageData.relativePath?.startsWith('blog/posts/')

    if (isBlogPost) {
      pageData.frontmatter.prev = false
      pageData.frontmatter.next = false
      // Optional: also hide the right "On this page" outline
      // pageData.frontmatter.aside = false
    }

    return pageData
  }
})
