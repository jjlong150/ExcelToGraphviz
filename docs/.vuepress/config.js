const { description } = require("../../package");

module.exports = {
  /**
   * Ref：https://v1.vuepress.vuejs.org/config/#title
   */
  title: "Excel to Graphviz",
  description: "Excel to Graphviz Relationship Visualizer",
  base: "/",
  /**
   * Ref：https://v1.vuepress.vuejs.org/config/#description
   */
  description: description,

  /**
   * Extra tags to be injected to the page HTML `<head>`
   *
   * ref：https://v1.vuepress.vuejs.org/config/#head
   */
  head: [
    ["meta", { name: "theme-color", content: "#3eaf7c" }],
    ["meta", { name: "apple-mobile-web-app-capable", content: "yes" }],
    [
      "meta",
      { name: "apple-mobile-web-app-status-bar-style", content: "black" },
    ],
  ],

  /**
   * Display line numbers whenever code is dissplayed
   *
   * https://v1.vuepress.vuejs.org/guide/markdown.html#line-numbers
   */
  markdown: {
    lineNumbers: true,
  },

  /**
   * Theme configuration, here is the default theme configuration for VuePress.
   *
   * ref：https://v1.vuepress.vuejs.org/theme/default-theme-config.html
   */
  themeConfig: {
    repo: "",
    editLinks: false,
    docsDir: "",
    editLinkText: "",
    lastUpdated: false,
    nav: [
      {
        text: "Download",
        items: [
          { text: "- Graphviz", link: "https://graphviz.org/download/" },
          {
            text: "- Relationship Visualizer", link: "https://sourceforge.net/projects/relationship-visualizer/",
          },
        ],
      },      
      {
        text: "Install",
        items: [
          { text: "- Microsoft Windows", link: "/install-win/" },
          {
            text: "- macOS", link: "/install-mac/",
          },
        ],
      },
      {
        text: "Worksheets",
        items: [
          { text: "Data",
            items: [
              { text: "- data", link: "/create/#the-data-worksheet" },
              { text: "- graph", link: "/create/#graph-to-worksheet" },
            ],
          },
          { text: "Style",
            items: [
              { text: "- style designer", link: "/tutorial/#the-style-designer-ribbon-tab" },
              { text: "- styles", link: "/styles/" },      
            ],
          },
          { text: "Data Exchange",
            items: [
              { text: "- sql", link: "/sql/" },
              { text: "- json", link: "/exchange/" },      
            ],
          }, 
          { text: "Post-processing",
            items: [
              { text: "- svg", link: "/svg/" },
            ],
          }, 
          { text: "Graphviz dot",
            items: [
              { text: "- source", link: "/source/" },
              { text: "- console", link: "/console/" },      
              { text: "- attributes", link: "/workbook/#help-attributes-worksheet" },     
              { text: "- colors", link: "/workbook/#help-colors-worksheet" },     
              { text: "- shapes", link: "/workbook/#help-shapes-worksheet" },     
            ],
          }, 
          { text: "Maintenance",
            items: [
              { text: "- diagnostics", link: "/diagnostics/" },
              { text: "- lists", link: "/lists/" },      
              { text: "- settings", link: "/settings/" },      
            ],
          }, 
          { text: "Information",
            items: [
              { text: "- info", link: "/info/" },  
            ],
          }, 
        ],
      },
      { text: "About...", 
        items: [
          { text: "- About Excel to Graphviz", link: "/about/" },
          { text: "- Acknowledgements", link: "/acknowledge/" },
          { text: "- Change Log", link: "/changelog/" },
          { text: "- Terminology", link: "/terminology/" },
        ]
      },
      { text: "Donate ☕", link: "https://buymeacoffee.com/exceltographviz"},
    ],
    docsRepo: "https://github.com/jjlong150/ExcelToGraphviz",
    sidebar: "auto",
    logo: "/logo.png",
    // default value is true. Set it to false to hide next page links on all pages
    nextLinks: true,
    // default value is true. Set it to false to hide prev page links on all pages
    prevLinks: true,
  },
  head: [["link", { rel: "icon", href: "/favicon.ico" }]],
  plugins: [
    ["@vuepress/active-header-links"],
    ["vuepress-plugin-simple-analytics"],
    [
      "vuepress-plugin-container",
      {
        type: "feature",
        before: (info) => `<div class="feature"><h2>${info}</h2><p>`,
        after: "</p></div>",
      },
    ],
    [
      "container",
      {
        type: "quote",
      },
    ],
    [
      "vuepress-plugin-container",
      {
        type: "feature_box",
        before: '<div class="features">',
        after: "</div>",
      },
    ],
  ],
};
