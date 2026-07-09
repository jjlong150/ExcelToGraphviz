// docs/.vitepress/sidebar.mts

// The full documentation sidebar. This is the default/fallback sidebar and
// is shown on every route that isn't explicitly overridden below (this
// includes /blog/** — left as-is intentionally, see note at bottom of file).
const docsSidebar = [
  {
    text: 'Overview',
    link: '/overview/',
    items: [
      {
        items: [
          { text: 'Workbook', link: '/workbook/' },
          { text: 'Launchpad', link: '/launchpad/' },
        ]
      },
      {
        text: 'Creating Graphs',
        link: '/create/',
        items: [
          { text: 'New Workbook', link: '/prepare/' },
          { text: 'Data Worksheet', link: '/dataworksheet/' },
          { text: 'Graphviz Tab', link: '/graphviztab/' },
          { text: 'Core Concepts', link: '/coreconcepts/' }
        ]
      },
      {
        text: 'Adding Style',
        link: '/addstyle/',
        items: [
          {
            text: 'Style Designer',
            link: '/designer/',
            collapsed: true,
            items: [
              { text: 'Color', link: '/designer/color/' },
              { text: 'Labels', link: '/designer/labels/' },
              { text: 'Shapes', link: '/designer/shapes/' },
              { text: 'Dimensions', link: '/designer/dimensions/' },
              { text: 'Borders', link: '/designer/borders/' },
              { text: 'Fills', link: '/designer/fills/' },
              { text: 'Images', link: '/designer/images/' },
              { text: 'Edges', link: '/designer/edges/' },
              { text: 'Head & Tail', link: '/designer/head-tail/' },
              { text: 'Clusters', link: '/designer/clusters/' }
            ]
          },
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
    text: 'Setup',
    items: [
      { text: 'Download', link: '/download/' },
      {
        text: 'Install',
        link: '/install/',
        collapsed: true,
        items: [
          { text: 'Windows Instructions', link: '/install-win/' },
          { text: 'macOS Instructions', link: '/install-mac/' }
        ]
      }
    ]
  },
  {
    text: 'SQL Tools <span style="font-size:0.75em; padding:2px 6px; background:#eee; border-radius:4px; color:#555;">Windows‑only</span>',
    link: '/sql/',
    items: [
      { text: 'SQL to Graph Example', link: '/sql/queries/' },
      {
        text: 'SQL Extensions',
        link: '/sql/extensions/',
        collapsed: true,
        items: [
          { text: 'Directives', link: '/sql/directives/' },
          {
            text: 'Label & Text Helpers',
            items: [
              { text: 'Substitute Counts', link: '/sql/counts/' },
              { text: 'Split Long Labels', link: '/sql/labelsplit/' },
              { text: 'Enumerate Values', link: '/sql/enumerate/' },
              { text: 'Concatenate Values', link: '/sql/concatenation/' }
            ]
          },
          {
            text: 'Structure & Grouping',
            items: [
              { text: 'Cluster Nodes', link: '/sql/clustering/' },
              { text: 'Align Nodes', link: '/sql/subgraphs/' }
            ]
          },
          {
            text: 'Iteration & Hierarchy',
            items: [
              { text: 'Chain Nodes', link: '/sql/chaining/' },
              { text: 'Traverse Trees', link: '/sql/recursion/' },
              { text: 'Iterate SQL Results', link: '/sql/iterate/' }
            ]
          }
        ]
      },
      {
        text: 'Advanced SQL Examples',
        collapsed: true,
        items: [
          { text: 'Organization Charts', link: '/sql/orgcharts/' },
          { text: 'Timelines and Roadmaps', link: '/sql/timeline/' }
        ]
      },
      { text: 'SQL Syntax Reference', link: '/sql/syntax/' },
    ]
  },
  {
    text: 'Data Exchange',
    link: '/exchange/',
    items: [
      { text: 'Export JSON', link: '/exchange/export/' },
      { text: 'Import JSON', link: '/exchange/import/' }
    ]
  },
  {
    text: 'Graphviz',
    items: [
      { text: 'DOT Source Code', link: '/source/' },
      { text: 'DOT Message Console', link: '/console/' }
    ],
  },
  {
    text: 'Maintenance',
    items: [
      { text: 'Diagnostics', link: '/diagnostics/' },
      { text: 'Lists', link: '/lists/' },
      { text: 'Settings', link: '/settings/' }
    ],
  },
  {
    text: 'References',
    items: [
      { text: 'Terminology', link: '/terminology/' },
      { text: 'Information', link: '/info/' }
    ]
  }
]

// Path-keyed sidebar. VitePress matches the most specific key that prefixes
// the current route, falling back to '/' for anything not listed.
//
// Standalone resource pages (About/License/Privacy/Security/Credits/Changelog)
// get an empty sidebar — they're single pages with no children, and showing
// the entire 40+ item docs tree next to them was pure noise.
//
// NOTE: /blog/** is deliberately NOT given its own key here. It falls through
// to the '/' default below, which is exactly its current (pre-existing)
// behavior. The blog plugin's own routing/behavior is known to be brittle
// right now — leave it untouched until that's addressed separately.
export default {
  '/about/': [],
  '/pricing/': [],
  '/license/': [],
  '/privacy/': [],
  '/security/': [],
  '/acknowledge/': [],
  '/changelog/': [],
  '/': docsSidebar
}