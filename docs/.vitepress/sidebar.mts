// docs/.vitepress/sidebar.mts

export default [
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
          { text: 'Terminology', link: '/terminology/' },
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
    text: 'Data Manipulation',
    items: [
      { text: 'Using SQL', link: '/sql/' },
      { text: 'SQL to Graph Example', link: '/sql/queries/' },
      {
        text: 'SQL Extensions', 
        link: '/sql/extensions/',
        items: [
          { text: 'Directives', link: '/sql/directives/' },
          { text: 'Cluster Nodes', link: '/sql/clustering/' },
          { text: 'Substitute Counts', link: '/sql/counts/' },
          { text: 'Split Long Labels', link: '/sql/labelsplit/' },
          { text: 'Chain Nodes', link: '/sql/chaining/' },
          { text: 'Align Nodes', link: '/sql/subgraphs/' },
          { text: 'Traverse Trees', link: '/sql/recursion/'},
          { text: 'Iterate SQL Results', link: '/sql/iterate/' },
          { text: 'Enumerate Values', link: '/sql/enumerate/' },
          { text: 'Concatenate Values', link: '/sql/concatenation/' }
        ]
      },
      {
        text: 'Advanced SQL Examples',
        items: [
          { text: 'Organization Charts', link: '/sql/orgcharts/' },
          { text: 'Timelines and Roadmaps', link: '/sql/timeline/' }
        ]
      },
      { text: 'SQL Syntax Reference', link: '/sql/syntax/' },
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
    text: 'Maintenance',
    items: [
      { text: 'Diagnostics', link: '/diagnostics/' },
      { text: 'Lists', link: '/lists/' },
      { text: 'Settings', link: '/settings/' },
      { text: 'Information', link: '/info/' }
    ],
  },
  {
    text: 'Exchange Data',
    items: [
      {
        text: 'Using JSON Files', link: '/exchange/',
        items: [
          { text: 'Export', link: '/exchange/export/' },
          { text: 'Import', link: '/exchange/import/' }
        ]
      }
    ]
  }
]
