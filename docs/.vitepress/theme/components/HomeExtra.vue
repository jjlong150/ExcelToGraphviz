<template>
  <div class="home-extra-content">
    <div v-html="html"></div>
  </div>
</template>

<script setup>
const html = `
  <h2>Turn Spreadsheets into Beautiful Graphviz Diagrams</h2>
  <p>The <strong>Relationship Visualizer</strong> spreadsheet transforms your Excel tables into clear, professional <strong>Graphviz diagrams</strong> in seconds. Say goodbye to manual drawing tools — simply enter your data as rows (e.g., "A is related to B"), and watch graphs appear automatically.</p>

  <h3>Why Users Love It</h3>
  <p>Whether you're mapping data flows, organization charts, process workflows, timelines, or dependency diagrams, the Relationship Visualizer turns complex relationships into clear, expressive diagrams.</p>

  <table class="features-table">
    <tr><td><strong>Draws as you type</strong></td><td>Live Graphviz rendering as data changes</td></tr>
    <tr><td><strong>Powerful styling</strong></td><td>Colors, shapes, fonts, arrows, and reusable CSS-like styles</td></tr>
    <tr><td><strong>Advanced features</strong></td><td>SQL queries, SVG animation, DOT preview, and JSON exchange</td></tr>

    <tr><td><strong>Cross-platform</strong></td><td>Works on Windows and macOS</td></tr>
    <tr><td><strong>Sleek UI</strong></td><td>Custom Excel ribbon tabs across all worksheets</td></tr>
    <tr><td><strong>Multilingual</strong></td><td>English · French · German · Italian · Polish</td></tr>

    <tr><td><strong>Absolutely Free</strong></td><td>Free to use · No license required · Donations appreciated</td></tr>
    <tr><td><strong>Open Source</strong></td><td>MIT License</td></tr>
    <tr><td><strong>Rich Code Documentation</strong></td><td><a href="https://deepwiki.com/jjlong150/ExcelToGraphviz" target="_blank" rel="noopener">AI-powered DeepWiki</a></td></tr>
    <tr><td><strong>Award Winning</strong></td><td>SourceForge Community Choice Award</td></tr>

    <tr><td><strong>Show Your Support</strong></td><td><a href="https://www.buymeacoffee.com/exceltographviz" target="_blank" rel="noopener">Buy Me a Coffee</a></td></tr>
  </table>

  <center>
    <img src="sourceforge-community-choice.png" alt="SourceForge Community Choice Award" width="90" style="margin-top: 8px;">
  </center>
`
</script>

<style scoped>
/* Your existing styles still apply */
</style>
