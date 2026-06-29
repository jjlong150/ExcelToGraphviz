import { defineConfig } from 'vitepress'
import { blogPlugin } from 'vitepress-plugin-blog/plugin'
import sidebar from './sidebar.mts'
import { joinURL, withoutTrailingSlash } from 'ufo'

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
    ['link', { rel: 'icon',             href: '/favicon.ico',          sizes: '32x32',   type: 'image/x-icon' }],
    ['link', { rel: 'icon',             href: '/favicon-32x32.png',    sizes: '32x32',   type: 'image/png' }],
    ['link', { rel: 'apple-touch-icon', href: '/apple-touch-icon.png', sizes: '180x180', type: 'image/png' }],
    ['link', { rel: 'icon',             href: '/favicon-192x192.png',  sizes: '192x192', type: 'image/png' }],

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
  description: "Convert Excel data into professional Graphviz relationship diagrams. Free Relationship Visualizer tool.",

  sitemap: {
    hostname: 'https://exceltographviz.com'
  },

  lastUpdated: true,

  themeConfig: {
    // https://vitepress.dev/reference/default-theme-config
    nav: [
      { text: 'Blog', link: '/blog/' },
      { text: "Resources",
            items: [
              { text: 'About', link: '/about/' },
              { text: 'License', link: '/license/' },
              { text: 'Privacy', link: '/privacy/' },
              { text: 'Security', link: '/security/' },
              { text: 'Credits', link: '/acknowledge/' },
              { text: 'Changelog', link: '/changelog/' }
            ]
      },
    ],
    
    externalLinkIcon: true,
    
    logo: "/logo.png",

    sidebar,

    search: {
      provider: 'local'
    },

    footer: {
      message: 'Released under the MIT License.',
      copyright: 'Copyright © 2015-present Jeffrey J. Long.'
    },

    editLink: {
      pattern: 'https://buymeacoffee.com/exceltographviz',
      text: 'Find this tool helpful? Consider supporting its development.'
    },

    socialLinks: [
      { icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24" ><title>GitHub Repository</title><path d="M12 0c-6.626 0-12 5.373-12 12 0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23.957-.266 1.983-.399 3.003-.404 1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.921.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576 4.765-1.589 8.199-6.086 8.199-11.386 0-6.627-5.373-12-12-12z"></path></svg>' },
        link: 'https://github.com/jjlong150/ExcelToGraphviz',
        ariaLabel: 'GitHub Repository' },
      { icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24" ><title>LinkedIn</title><path fill="currentColor" d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037c-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85c3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.06 2.06 0 0 1-2.063-2.065a2.064 2.064 0 1 1 2.063 2.065m1.782 13.019H3.555V9h3.564zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0z"/></svg>' },
        link: 'https://www.linkedin.com/in/jeffreyjlong/',
        ariaLabel: 'LinkedIn' },
      { icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24" ><title>X/Twitter</title><path fill="currentColor" d="M14.234 10.162L22.977 0h-2.072l-7.591 8.824L7.251 0H.258l9.168 13.343L.258 24H2.33l8.016-9.318L16.749 24h6.993zm-2.837 3.299l-.929-1.329L3.076 1.56h3.182l5.965 8.532l.929 1.329l7.754 11.09h-3.182z"/></svg>' },
        link: 'https://x.com/exceltographviz',
        ariaLabel: 'X/Twitter' },
      { 
        icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24" ><title>Buy Me a Coffee</title><path d="m12,8.07h0s0-.01,0-.01c-.01,0-.03.02-.04.03h.02Z"/><path d="m11.98,20.58h.01s.02-.03.03-.04c-.02,0-.03.02-.05.04Z"/><path d="m12.01,20.7h0s-.03-.02-.05-.02c.01,0,.03.01.04.02Z"/><path d="m14.26,20.04c.23-.21.37-.51.4-.82l.78-8.29c-.35-.12-.7-.2-1.1-.2-.69,0-1.24.24-1.88.51-.72.31-1.53.66-2.59.66-.44,0-.88-.06-1.31-.18l.73,7.5c.03.31.17.61.4.82.23.21.53.33.85.33,0,0,1.04.05,1.38.05.37,0,1.49-.05,1.49-.05.31,0,.62-.12.85-.33Z"/><path d="m12.01,8.05s-.01-.01-.02-.02h.01s0,.02,0,.02Z"/><path d="m18.74,6.79c-.1-.5-.32-.97-.83-1.15-.16-.06-.35-.08-.48-.2-.13-.12-.16-.3-.19-.48-.05-.32-.1-.63-.16-.94-.05-.27-.09-.57-.21-.82-.16-.33-.5-.53-.83-.66-.17-.06-.34-.12-.52-.16-.83-.22-1.71-.3-2.56-.35-1.03-.06-2.06-.04-3.08.05-.76.07-1.57.15-2.29.42-.27.1-.54.21-.74.42-.25.25-.33.64-.15.95.13.22.35.38.58.48.3.13.61.24.94.3.9.2,1.82.28,2.74.31,1.01.04,2.03,0,3.04-.1.25-.03.5-.06.75-.1.29-.04.48-.43.39-.7-.1-.32-.38-.44-.7-.39-.05,0-.09.01-.14.02h-.03c-.11.02-.21.03-.32.04-.22.02-.44.04-.66.06-.49.03-.99.05-1.49.05-.49,0-.97-.01-1.46-.05-.22-.01-.44-.03-.66-.06-.1-.01-.2-.02-.3-.03h-.1s-.02-.02-.02-.02h-.1c-.2-.04-.4-.08-.6-.12-.02,0-.04-.02-.05-.03-.01-.02-.02-.04-.02-.06s0-.04.02-.06c.01-.02.03-.03.05-.03h0c.17-.04.35-.07.52-.1.06,0,.12-.02.18-.03h0c.11,0,.22-.03.33-.04.95-.1,1.9-.13,2.85-.1.46.01.92.04,1.39.09.1.01.2.02.3.03.04,0,.08,0,.11.01h.08c.22.04.44.08.66.13.33.07.75.09.89.45.05.11.07.24.09.36l.03.15s0,0,0,0c.08.36.15.72.23,1.08,0,.03,0,.05,0,.08s-.02.05-.03.07-.04.04-.06.06c-.02.01-.05.02-.08.03h-.05s-.05.01-.05.01c-.15.02-.3.04-.44.05-.29.03-.58.06-.87.09-.58.05-1.16.08-1.74.09-.3,0-.59.01-.89.01-1.18,0-2.36-.07-3.53-.21-.12-.01-.24-.03-.36-.04-.02,0-.11-.01-.13-.02-.08-.01-.16-.02-.24-.04-.27-.04-.54-.09-.81-.13-.33-.05-.64-.03-.94.13-.24.13-.44.34-.56.58-.13.26-.17.55-.22.83-.06.28-.15.59-.11.88.07.63.51,1.14,1.14,1.25.59.11,1.19.19,1.79.27,2.35.29,4.73.32,7.09.1.19-.02.38-.04.58-.06.06,0,.12,0,.18.02s.11.05.15.09c.04.04.08.09.1.15.02.06.03.12.02.18l-.06.58c-.12,1.17-.24,2.35-.36,3.52-.13,1.23-.25,2.46-.38,3.7-.04.35-.07.69-.11,1.04-.03.34-.04.69-.1,1.03-.1.53-.46.86-.99.98-.48.11-.97.17-1.46.17-.55,0-1.09-.02-1.64-.02-.58,0-1.3-.05-1.75-.48-.4-.38-.45-.98-.5-1.49-.07-.68-.14-1.37-.21-2.05l-.4-3.8-.26-2.46s0-.08-.01-.12c-.03-.29-.24-.58-.57-.57-.28.01-.6.25-.57.57l.19,1.82.39,3.77c.11,1.07.22,2.14.33,3.21.02.21.04.41.06.62.12,1.12.98,1.72,2.04,1.89.62.1,1.25.12,1.88.13.81.01,1.62.04,2.41-.1,1.17-.22,2.05-1,2.18-2.21.04-.35.07-.7.11-1.05.12-1.16.24-2.32.36-3.48l.39-3.79.18-1.74c0-.09.05-.17.1-.23s.14-.11.22-.12c.34-.07.66-.18.89-.43.38-.41.46-.94.32-1.47l-.11-.56Zm-12.48,1.18h0s.01,0,.02.02c-.01-.01-.02-.02-.02-.02Zm.1.09h0s0,0,0,0c0,0,0,0,0,0Zm11.26-.08h0c-.12.11-.3.17-.48.19-2.01.3-4.05.45-6.09.38-1.46-.05-2.9-.21-4.34-.42-.14-.02-.29-.05-.39-.15-.06-.06-.08-.14-.1-.23,0,0,0,0,0,.01,0,0,0,0,0-.01-.03-.2.02-.43.05-.6.04-.22.13-.51.39-.54.4-.05.87.12,1.27.18.48.07.96.13,1.45.18,2.07.19,4.17.16,6.23-.12.38-.05.75-.11,1.12-.18.33-.06.7-.17.9.17.14.23.16.55.13.81,0,.12-.06.22-.14.3Z"/><path d="m6.22,7.77s0,0,0,0c0-.02.01-.05,0-.05,0,0,0,.03,0,.05Z"/></svg>' },
        link: 'https://buymeacoffee.com/exceltographviz',
        ariaLabel: 'Buy Me a Coffee' },
      { 
        icon: { svg: '<svg  xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 44 50" ><title>DeepWiki</title><path fill="currentColor" d="M1.117,20.553l5.351,3.089c0.192,0.111,0.406,0.165,0.621,0.165c0.214,0,0.429-0.057,0.621-0.165l5.351-3.089 c0,0,0.015-0.012,0.022-0.017c0.081-0.049,0.158-0.108,0.227-0.175c0.01-0.01,0.02-0.022,0.03-0.032 c0.059-0.064,0.113-0.133,0.158-0.207c0.007-0.012,0.017-0.022,0.022-0.035c0.047-0.081,0.081-0.167,0.108-0.259 c0.005-0.02,0.01-0.039,0.015-0.059c0.022-0.094,0.039-0.19,0.039-0.291v-3.089c0-1.192,0.643-2.303,1.675-2.9s2.316-0.596,3.35,0 l2.675,1.545c0.086,0.049,0.177,0.084,0.271,0.111c0.02,0.005,0.04,0.012,0.059,0.017c0.091,0.022,0.185,0.035,0.278,0.037 c0.005,0,0.01,0,0.012,0c0.01,0,0.02-0.005,0.029-0.005c0.086,0,0.173-0.012,0.256-0.035c0.015-0.003,0.03-0.005,0.044-0.01 c0.091-0.025,0.18-0.062,0.264-0.108c0.007-0.005,0.017-0.005,0.025-0.01l5.351-3.089c0.384-0.222,0.621-0.631,0.621-1.074V4.69 c0-0.443-0.236-0.852-0.621-1.074l-5.356-3.087c-0.384-0.222-0.855-0.222-1.239,0l-5.351,3.089c0,0-0.015,0.012-0.022,0.017 c-0.081,0.049-0.158,0.108-0.227,0.175c-0.01,0.01-0.02,0.022-0.03,0.032c-0.059,0.064-0.113,0.133-0.158,0.207 c-0.007,0.012-0.017,0.022-0.022,0.034c-0.047,0.081-0.081,0.168-0.108,0.259c-0.005,0.02-0.01,0.039-0.015,0.059 c-0.022,0.094-0.039,0.19-0.039,0.291v3.089c0,1.192-0.643,2.303-1.675,2.902c-1.032,0.596-2.316,0.596-3.35,0L7.705,9.139 C7.618,9.09,7.527,9.055,7.434,9.028c-0.02-0.005-0.039-0.012-0.059-0.017C7.283,8.989,7.19,8.977,7.096,8.974 c-0.015,0-0.027,0-0.042,0c-0.089,0-0.175,0.012-0.259,0.034c-0.015,0.002-0.027,0.005-0.042,0.01 C6.663,9.043,6.574,9.08,6.49,9.127c-0.007,0.005-0.017,0.005-0.025,0.01l-5.348,3.092c-0.384,0.222-0.621,0.631-0.621,1.074v6.178 c0,0.444,0.236,0.852,0.621,1.074V20.553z"></path> <path fill="currentColor" d="M30.262,22.097c1.032-0.596,2.316-0.596,3.35,0l2.675,1.545c0.086,0.049,0.177,0.084,0.271,0.111 c0.02,0.005,0.039,0.012,0.059,0.017c0.091,0.022,0.185,0.034,0.278,0.037c0.005,0,0.01,0,0.012,0c0.01,0,0.02-0.003,0.029-0.005 c0.086,0,0.173-0.012,0.256-0.034c0.015-0.003,0.03-0.005,0.044-0.01c0.091-0.025,0.177-0.062,0.264-0.108 c0.007-0.005,0.017-0.005,0.027-0.01l5.351-3.089c0.384-0.222,0.621-0.631,0.621-1.074v-6.179c0-0.443-0.237-0.852-0.621-1.074 L37.53,9.134c-0.384-0.222-0.855-0.222-1.239,0l-5.351,3.089c0,0-0.015,0.012-0.022,0.017c-0.081,0.049-0.158,0.108-0.227,0.175 c-0.01,0.01-0.02,0.022-0.029,0.032c-0.059,0.064-0.113,0.133-0.158,0.207c-0.007,0.012-0.017,0.022-0.022,0.035 c-0.047,0.081-0.081,0.168-0.108,0.259c-0.005,0.02-0.01,0.039-0.015,0.059c-0.022,0.094-0.039,0.19-0.039,0.291v3.089 c0,1.192-0.643,2.303-1.675,2.902c-1.032,0.596-2.316,0.596-3.35,0l-2.675-1.545c-0.086-0.049-0.177-0.084-0.271-0.111 c-0.02-0.005-0.039-0.012-0.059-0.017c-0.091-0.022-0.185-0.035-0.278-0.037c-0.015,0-0.027,0-0.042,0 c-0.089,0-0.175,0.012-0.259,0.035c-0.015,0.003-0.027,0.005-0.042,0.01c-0.091,0.025-0.18,0.062-0.264,0.108 c-0.007,0.005-0.017,0.005-0.025,0.01l-5.351,3.089c-0.384,0.222-0.621,0.631-0.621,1.074v6.179c0,0.443,0.236,0.852,0.621,1.074 l5.351,3.089c0,0,0.017,0.005,0.025,0.01c0.084,0.047,0.173,0.084,0.264,0.108c0.015,0.005,0.03,0.005,0.044,0.01 c0.084,0.02,0.17,0.032,0.256,0.035c0.01,0,0.02,0.005,0.03,0.005c0.005,0,0.01,0,0.012,0c0.094,0,0.185-0.015,0.278-0.037 c0.02-0.005,0.039-0.01,0.059-0.017c0.094-0.027,0.185-0.062,0.271-0.111l2.675-1.545c1.032-0.596,2.316-0.596,3.35,0 c1.032,0.596,1.675,1.707,1.675,2.9v3.089c0,0.101,0.015,0.197,0.039,0.291c0.005,0.02,0.01,0.039,0.015,0.059 c0.027,0.091,0.061,0.177,0.108,0.259c0.007,0.012,0.015,0.022,0.022,0.034c0.044,0.074,0.099,0.143,0.158,0.207 c0.01,0.01,0.02,0.022,0.029,0.032c0.067,0.066,0.143,0.123,0.227,0.175c0.007,0.005,0.012,0.012,0.022,0.017l5.351,3.089 c0.192,0.111,0.407,0.165,0.621,0.165c0.214,0,0.429-0.057,0.621-0.165l5.351-3.089c0.384-0.222,0.621-0.631,0.621-1.074v-6.179 c0-0.443-0.236-0.852-0.621-1.074l-5.351-3.089c0,0-0.017-0.005-0.025-0.01c-0.084-0.047-0.173-0.084-0.264-0.108 c-0.015-0.005-0.027-0.005-0.042-0.01c-0.086-0.02-0.172-0.032-0.261-0.035c-0.012,0-0.027,0-0.039,0 c-0.094,0-0.187,0.015-0.278,0.037c-0.02,0.005-0.037,0.01-0.057,0.017c-0.094,0.027-0.185,0.062-0.271,0.111l-2.675,1.545 c-1.032,0.596-2.316,0.596-3.348,0c-1.032-0.596-1.675-1.707-1.675-2.902c0-1.195,0.643-2.303,1.675-2.9H30.262z"></path> <path fill="currentColor" d="M27.967,38.054l-5.351-3.089c0,0-0.017-0.005-0.025-0.01c-0.084-0.047-0.172-0.084-0.264-0.108 c-0.015-0.005-0.03-0.005-0.044-0.01c-0.086-0.02-0.172-0.032-0.259-0.035c-0.015,0-0.027,0-0.042,0 c-0.094,0-0.187,0.015-0.278,0.037c-0.02,0.005-0.037,0.01-0.057,0.017c-0.094,0.027-0.185,0.062-0.271,0.111l-2.675,1.545 c-1.032,0.596-2.316,0.596-3.348,0c-1.032-0.596-1.675-1.707-1.675-2.902V30.52c0-0.101-0.015-0.197-0.039-0.291 c-0.005-0.02-0.01-0.039-0.015-0.059c-0.027-0.091-0.062-0.177-0.108-0.259c-0.007-0.012-0.015-0.022-0.022-0.035 c-0.044-0.074-0.099-0.143-0.158-0.207c-0.01-0.01-0.02-0.022-0.03-0.032c-0.066-0.066-0.143-0.123-0.227-0.175 c-0.007-0.005-0.012-0.012-0.022-0.017l-5.351-3.089c-0.384-0.222-0.855-0.222-1.239,0l-5.351,3.089 c-0.384,0.222-0.621,0.631-0.621,1.074v6.179c0,0.443,0.236,0.852,0.621,1.074l5.351,3.089c0,0,0.017,0.007,0.025,0.01 c0.084,0.047,0.17,0.084,0.261,0.108c0.015,0.005,0.03,0.007,0.044,0.01c0.084,0.02,0.17,0.032,0.256,0.035 c0.01,0,0.02,0.005,0.032,0.005c0.005,0,0.01,0,0.015,0c0.094,0,0.185-0.015,0.276-0.037c0.02-0.005,0.039-0.01,0.059-0.017 c0.094-0.027,0.185-0.062,0.271-0.111l2.675-1.545c1.032-0.596,2.316-0.596,3.35,0c1.032,0.596,1.675,1.707,1.675,2.9v3.089 c0,0.101,0.015,0.197,0.039,0.291c0.005,0.02,0.01,0.039,0.015,0.059c0.027,0.091,0.062,0.177,0.108,0.259 c0.007,0.012,0.015,0.022,0.022,0.035c0.044,0.074,0.099,0.143,0.158,0.207c0.01,0.01,0.02,0.022,0.03,0.032 c0.067,0.067,0.143,0.123,0.227,0.175c0.007,0.005,0.012,0.012,0.022,0.017l5.351,3.089c0.192,0.111,0.406,0.165,0.621,0.165 s0.429-0.057,0.621-0.165l5.351-3.089c0.384-0.222,0.621-0.631,0.621-1.074V39.13c0-0.443-0.236-0.852-0.621-1.074L27.967,38.054z"></path></svg>' },
        link: 'https://deepwiki.com/jjlong150/ExcelToGraphviz',
        ariaLabel: 'DeepWiki' }
    ]
  },

  transformPageData(pageData) {
    
    const siteUrl = 'https://exceltographviz.com'

    // Modern clean canonical URL (no trailing slash)
    let path = pageData.relativePath
      .replace(/\.md$/, '')
      .replace(/index$/, '')
      .replace(/\/$/, '')

    const canonicalUrl = path ? `${siteUrl}/${path}` : siteUrl

    // === Blog logic ===
    const isBlogPost = 
      pageData.frontmatter?.blogPost === true ||
      pageData.relativePath?.startsWith('blog/posts/')

    if (isBlogPost) {
      pageData.frontmatter.prev = false
      pageData.frontmatter.next = false
      // pageData.frontmatter.aside = false // uncomment if you want to hide outline too
    }

    // === SEO Meta Tags ===
    pageData.frontmatter.head ??= []

    // === Allow per-page overrides for OG/Twitter titles ===
    const ogTitle = pageData.frontmatter?.ogTitle || pageData.title
    const twitterTitle = pageData.frontmatter?.twitterTitle || ogTitle

    pageData.frontmatter.head.push(
      // Google Search Console Verification Tag
      ['meta', { 
        name: 'google-site-verification', 
        content: 'Nk5wPIfa_duB_rD_ceHGXUhbTQhLn-aDcK8SpbhMiIg' 
      }],

      // Bing Webmaster Tools Verification Tag
      ['meta', {
        name: 'msvalidate.01',
        content: '170417B93AA71CAAE347C7AD019A4460'
      }],

      // Yandex Tools Verification Tag
      ['meta', {
        name: 'yandex-verification',
        content: '1f2054ab09c0647a'
      }],

      // Canonical URL
      ['link', { rel: 'canonical', href: canonicalUrl }],

      // Open Graph
      ['meta', { property: 'og:title', content: ogTitle }],
      ['meta', { property: 'og:description', content: pageData.description || '' }],
      ['meta', { property: 'og:url', content: canonicalUrl }],
      ['meta', { property: 'og:type', content: isBlogPost ? 'article' : 'website' }],
      ['meta', { property: 'og:site_name', content: 'Excel to Graphviz' }],
      ['meta', { property: 'og:locale', content: 'en_US' }], 
           
      // Twitter / X Cards
      ['meta', { name: 'twitter:card', content: 'summary_large_image' }],
      ['meta', { name: 'twitter:title', content: twitterTitle }],
      ['meta', { name: 'twitter:description', content: pageData.description || '' }],
      ['meta', { name: 'twitter:site', content: '@exceltographviz' }]
    )

    // === Allow per-page overrides for social images ===
    const customOgImage = pageData.frontmatter?.ogImage
    const customTwitterImage = pageData.frontmatter?.twitterImage

    // === Social Image Handling ===
    const defaultImage = pageData.frontmatter.layout === 'home'
      ? 'https://exceltographviz.com/social-hero.png'
      : 'https://exceltographviz.com/social-default.png'

    const ogImage = customOgImage || defaultImage
    const twitterImage = customTwitterImage || ogImage

    pageData.frontmatter.head.push(
      ['meta', { property: 'og:image', content: ogImage }],
      ['meta', { name: 'twitter:image', content: twitterImage }],
      ['meta', { property: 'og:image:width', content: '1200' }],
      ['meta', { property: 'og:image:height', content: '630' }],
      ['meta', { property: 'og:image:alt', content: 'Excel to Graphviz Relationship Visualizer' }]
    )
  
    // === JSON-LD Structured Data ===
    const isHome = pageData.frontmatter.layout === 'home'

    const jsonLd = {
      "@context": "https://schema.org",
      "@type": isHome
        ? "WebSite"
        : isBlogPost
          ? "Article"
          : "WebPage",
      "headline": pageData.title,
      "description": pageData.description || "",
      "url": canonicalUrl,
      ...(isBlogPost && {
        "datePublished": pageData.frontmatter.date || undefined,
        "dateModified": pageData.frontmatter.lastUpdated || undefined
      }),
      "author": {
        "@type": "Person",
        "name": "Jeffrey Long",
        "url": "https://exceltographviz.com"
      },
      "publisher": {
        "@type": "Organization",
        "name": "Excel to Graphviz",
        "url": "https://exceltographviz.com",
        "logo": {
          "@type": "ImageObject",
          "url": "https://exceltographviz.com/logo.png"
        }
      }
    }

    pageData.frontmatter.head.push([
      'script',
      { type: 'application/ld+json' },
      JSON.stringify(jsonLd)
    ])

    // === SoftwareApplication JSON-LD (only on download page) ===
    if (pageData.relativePath === 'download/index.md') {
      const softwareJsonLd = {
        "@context": "https://schema.org",
        "@type": "SoftwareApplication",
        "name": "Excel to Graphviz Relationship Visualizer",
        "operatingSystem": "Windows, macOS",
        "applicationCategory": "UtilityApplication",
        "description": "A VBA-powered Excel tool that converts spreadsheet relationships into Graphviz diagrams.",
        "softwareVersion": "10.5.0",
        "downloadUrl": "https://exceltographviz.com/download/",
        "offers": {
          "@type": "Offer",
          "price": "0",
          "priceCurrency": "USD"
        },
        "author": {
          "@type": "Person",
          "name": "Jeffrey Long",
          "url": "https://exceltographviz.com"
        },
        "publisher": {
          "@type": "Organization",
          "name": "Excel to Graphviz",
          "url": "https://exceltographviz.com"
        }
      }

      pageData.frontmatter.head.push([
        'script',
        { type: 'application/ld+json' },
        JSON.stringify(softwareJsonLd)
      ])
    }

    if (pageData.relativePath === 'install-win/index.md') {
      const howToJsonLd = {
        "@context": "https://schema.org",
        "@type": "HowTo",
        "name": "Install Excel to Graphviz Relationship Visualizer on Windows",
        "description": "Step-by-step instructions for installing Graphviz, configuring command-line tools, downloading the Relationship Visualizer assets, unblocking the spreadsheet, and enabling macros in Excel.",
        "totalTime": "PT15M",
        "tool": [
          { "@type": "HowToTool", "name": "Microsoft Excel" },
          { "@type": "HowToTool", "name": "Graphviz" }
        ],
        "step": [
          {
            "@type": "HowToStep",
            "name": "Download and install Graphviz",
            "text": "Download the 32-bit or 64-bit Graphviz EXE installer and ensure the Graphviz bin directory is added to the PATH."
          },
          {
            "@type": "HowToStep",
            "name": "Open Command Prompt as Administrator",
            "text": "Run Command Prompt using 'Run as Administrator', then execute 'dot -c' to register plugins and 'dot -V' to confirm the installation."
          },
          {
            "@type": "HowToStep",
            "name": "Download the Relationship Visualizer assets",
            "text": "Download RelationshipVisualizer.zip from SourceForge and optionally validate SHA1 or MD5 checksums."
          },
          {
            "@type": "HowToStep",
            "name": "Extract the files",
            "text": "Extract all files from the ZIP archive to a local directory."
          },
          {
            "@type": "HowToStep",
            "name": "Unblock the spreadsheet file",
            "text": "Right-click Relationship Visualizer.xlsm, open Properties, and check the Unblock box before clicking OK."
          },
          {
            "@type": "HowToStep",
            "name": "Enable macros and open Excel",
            "text": "Enable VBA macros in Excel’s Trust Center, then open Relationship Visualizer.xlsm and allow macros when prompted."
          }
        ]
      }

      pageData.frontmatter.head.push([
        'script',
        { type: 'application/ld+json' },
        JSON.stringify(howToJsonLd)
      ])
    }

    if (pageData.relativePath === 'install-mac/index.md') {
      const howToJsonLd = {
        "@context": "https://schema.org",
        "@type": "HowTo",
        "name": "Install Excel to Graphviz Relationship Visualizer on macOS",
        "description": "Step-by-step instructions for installing Graphviz using Homebrew, configuring plugins, preparing the AppleScript file, and enabling macros in Excel on macOS.",
        "totalTime": "PT15M",
        "tool": [
          { "@type": "HowToTool", "name": "Microsoft Excel" },
          { "@type": "HowToTool", "name": "Graphviz" },
          { "@type": "HowToTool", "name": "Terminal" }
        ],
        "step": [
          {
            "@type": "HowToStep",
            "name": "Install Graphviz using Homebrew",
            "text": "Run the command 'brew install graphviz' to install Graphviz on macOS."
          },
          {
            "@type": "HowToStep",
            "name": "Configure Graphviz plugins",
            "text": "Open a Terminal window and run 'sudo dot -c' to register Graphviz plugins. Enter your administrator password when prompted."
          },
          {
            "@type": "HowToStep",
            "name": "Confirm Graphviz installation",
            "text": "Run 'dot -V' to confirm Graphviz is installed and responding with a version number."
          },
          {
            "@type": "HowToStep",
            "name": "Download RelationshipVisualizer.zip",
            "text": "Download the RelationshipVisualizer.zip file from SourceForge, which contains the spreadsheet, AppleScript file, documentation, and samples."
          },
          {
            "@type": "HowToStep",
            "name": "Unzip the downloaded file",
            "text": "Extract the contents of RelationshipVisualizer.zip to any local directory."
          },
          {
            "@type": "HowToStep",
            "name": "Determine the path to the dot command",
            "text": "Run 'which dot' in Terminal to determine the installation path of the Graphviz dot command."
          },
          {
            "@type": "HowToStep",
            "name": "Edit ExcelToGraphviz.applescript if needed",
            "text": "If the dot command is not located at /usr/local/bin/dot, edit ExcelToGraphviz.applescript and update the path on line 2 to match the output of 'which dot'."
          },
          {
            "@type": "HowToStep",
            "name": "Copy the AppleScript file to the Excel sandbox folder",
            "text": "Copy ExcelToGraphviz.applescript to ~/Library/Application Scripts/com.microsoft.Excel to comply with macOS sandboxing rules."
          },
          {
            "@type": "HowToStep",
            "name": "Open Relationship Visualizer.xlsm in Excel",
            "text": "Double-click Relationship Visualizer.xlsm to open it in Excel."
          },
          {
            "@type": "HowToStep",
            "name": "Enable macros in Excel",
            "text": "Grant permission for macros to run when prompted, as the tool requires VBA macros to function."
          },
          {
            "@type": "HowToStep",
            "name": "Save the spreadsheet as a template",
            "text": "Use 'File -> Save as Template...' to save the workbook as an Excel Macro-Enabled Template (.xltm) for future use."
          }
        ]
      }

      pageData.frontmatter.head.push([
        'script',
        { type: 'application/ld+json' },
        JSON.stringify(howToJsonLd)
      ])
    }

    return pageData
  }
})

