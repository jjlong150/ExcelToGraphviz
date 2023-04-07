const { description } = require("../../package");

module.exports = {
  /**
   * Ref：https://v1.vuepress.vuejs.org/config/#title
   */
  title: "Excel to Graphviz",
  description: "Excel to Graphviz Relationship Visualizer",
  base: "/ExcelToGraphviz/",
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
      { text: "Overview", link: "/overview/" },
      { text: "Installallation", link: "/install-win/" },
      { text: "Terminology", link: "/terminology/" },
      { text: "Notices", link: "/thirdparty/" },
    ],
    docsRepo: "https://github.com/jjlong150/ExcelToGraphviz",
    sidebar: "auto",
    logo: "/hero.png",
    // default value is true. Set it to false to hide next page links on all pages
    nextLinks: true,
    // default value is true. Set it to false to hide prev page links on all pages
    prevLinks: true,
  },
  head: [["link", { rel: "icon", href: "/favicon.ico" }]],
  plugins: [
    ["@vuepress/active-header-links"],
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
