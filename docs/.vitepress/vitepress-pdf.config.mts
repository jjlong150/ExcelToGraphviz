import { defineUserConfig } from 'vitepress-export-pdf'
import { routeOrder } from './sidebar-to-routes.mts'

const headerTemplate = `<div style="margin-top: -0.4cm; height: 70%; width: 100%; display: flex; justify-content: center; align-items: center; color: lightgray; border-bottom: solid lightgray 1px; font-size: 10px;">
  <span class="title"></span>
</div>`

const footerTemplate = `<div style="margin-bottom: -0.4cm; height: 70%; width: 100%; display: flex; justify-content: flex-start; align-items: center; color: lightgray; border-top: solid lightgray 1px; font-size: 10px;">
  <span style="margin-left: 15px;" class="url"></span>
</div>`

export default defineUserConfig({
  pdfOutlines: true,

  outlineContainerSelector: '.vp-doc',

  routePatterns: ['/**',
                  '!/blog/**',
                  '!/changelog/**',
                  '!/tutorial/**'],

  outFile: 'Relationship Visualizer.pdf',

  outDir: '../dist/pdf',

  puppeteerLaunchOptions: {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  },

  pdfOptions: {
    format: 'Letter',
    printBackground: true,
    displayHeaderFooter: true,
    outline: true,
    preferCSSPageSize: true,
    headerTemplate,
    footerTemplate,
    margin: {
      bottom: 60,
      left: 40,
      right: 40,
      top: 60,
    },
    timeout: 120000
 },

  urlOrigin: 'https://exceltographviz.com/',

  sorter: (pageA, pageB) => {
    const aIndex = routeOrder.indexOf(pageA.path)
    const bIndex = routeOrder.indexOf(pageB.path)

    if (aIndex === -1 && bIndex === -1) {
      return pageA.path.localeCompare(pageB.path)
    }

    if (aIndex === -1) return 1
    if (bIndex === -1) return -1

    return aIndex - bIndex
  }

})
