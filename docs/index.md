---
# https://vitepress.dev/reference/default-theme-home-page
layout: home
markdownStyles: true

hero:
  name: "Excel to Graphviz"
  text: "Relationship Visualizer"
  tagline: Turn Excel data into professional Graphviz diagrams
  image:
    src: hero.png
    alt: Excel to Graphviz
  actions:
    - theme: brand
      text: Overview
      link: /overview/
    - theme: alt
      text: Download
      link: /download/
    - theme: alt
      text: Install
      link: /install/

features:
  - icon:
      dark: /share-2.svg
      light: /share-2.svg
    title: Visualize Graphs Using Excel
    details: Create Graphviz graphs directly from your Excel data
    link: /create/
  - icon:
      dark: /palette.svg
      light: /palette.svg
    title: Apply Style
    details: Fast, expressive graph styling with a built‑in designer and gallery
    link: /designer/
  - icon:
      dark: /image-down.svg
      light: /image-down.svg
    title: Publish Graphs
    details: Save graphs as image, PDF, or SVG files<br/>Add animation to SVG files
    link: /publish/
  - icon:
      dark: /database-search.svg
      light: /database-search.svg
    title: Manipulate Data Using SQL
    details: Use SQL to retrieve data from Excel and Access (Windows-only)
    link: /sql/
  - icon:
      dark: /monitor-cog.svg
      light: /monitor-cog.svg
    title: View Graphviz Source
    details: Inspect and export the DOT source that defines the graph
    link: /source/
  - icon:
      dark: /file-braces.svg
      light: /file-braces.svg
    title: Exchange Data Using JSON
    details: Exchange workbook data in a JSON format suitable for version control
    link: /exchange/
---
