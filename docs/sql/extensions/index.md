# SQL Extensions

SQL extensions provide small, declarative utilities that simplify common graph‑building tasks in the **Relationship Visualizer**.  

Each extension is activated by passing specific values as SQL parameters, enabling expressive diagrams with minimal query logic.

Use the cards below to explore each extension.

<div class="advanced-grid">
  <a class="advanced-card" href="/sql/directives/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M4.887 20h11.868c.893 0 1.664 -.665 1.847 -1.592l2.358 -12c.212 -1.081 -.442 -2.14 -1.462 -2.366a1.784 1.784 0 0 0 -.385 -.042h-11.868c-.893 0 -1.664 .665 -1.847 1.592l-2.358 12c-.212 1.081 .442 2.14 1.462 2.366c.127 .028 .256 .042 .385 .042" />
        <path d="M9 8l4 4l-6 4" />
        <path d="M12 16h3" />
      </svg>
    </span>
    Directives<br><br>Lightweight commands that enable optional behaviors within the SQL pipeline.
  </a>
  <a class="advanced-card" href="/sql/clustering/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M8 8h8v8h-8l0 -8" />
        <path d="M4 6a2 2 0 0 1 2 -2h12a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-12a2 2 0 0 1 -2 -2l0 -12" />
      </svg>
    </span>
    Group Nodes into Nested Clusters<br><br>Group related rows into clusters or subclusters to visually organize sections of your graph.
  </a>
  <a class="advanced-card" href="/sql/labelsplit/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M3 19v-14a2 2 0 0 1 2 -2h14a2 2 0 0 1 2 2v14a2 2 0 0 1 -2 2h-14a2 2 0 0 1 -2 -2" />
        <path d="M11 17h2" />
        <path d="M9 12h6" />
        <path d="M10 7h4" />
      </svg>
    </span>
    Split Long Label Text<br><br>Split long text labels into multiple lines for improved readability.
  </a>
  <a class="advanced-card" href="/sql/subgraphs/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M4 6a2 2 0 0 1 2 -2h2a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-2a2 2 0 0 1 -2 -2l0 -12" />
        <path d="M14 6a2 2 0 0 1 2 -2h2a2 2 0 0 1 2 2v6a2 2 0 0 1 -2 2h-2a2 2 0 0 1 -2 -2l0 -6" />
      </svg>
    </span>
    Align Nodes on the Same Level<br><br>Wrap selected nodes into ranked subgraphs to control layout, alignment, and visual grouping.
  </a>
  <a class="advanced-card" href="/sql/chaining/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M15 6.5a2.5 2.5 0 1 0 5 0a2.5 2.5 0 1 0 -5 0" />
        <path d="M4 17.5a2.5 2.5 0 1 0 5 0a2.5 2.5 0 1 0 -5 0" />
        <path d="M8.5 15.5l7 -7" />
      </svg>
    </span>
    Generate Edges to Chain Nodes<br><br>Generate edges between sequential nodes to create simple chains—useful for timelines or ordered flows. 
  </a>
  <a class="advanced-card" href="/sql/recursion/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M6 20a2 2 0 1 0 -4 0a2 2 0 0 0 4 0" />
        <path d="M16 4a2 2 0 1 0 -4 0a2 2 0 0 0 4 0" />
        <path d="M16 20a2 2 0 1 0 -4 0a2 2 0 0 0 4 0" />
        <path d="M11 12a2 2 0 1 0 -4 0a2 2 0 0 0 4 0" />
        <path d="M21 12a2 2 0 1 0 -4 0a2 2 0 0 0 4 0" />
        <path d="M5.058 18.306l2.88 -4.606" />
        <path d="M10.061 10.303l2.877 -4.604" />
        <path d="M10.065 13.705l2.876 4.6" />
        <path d="M15.063 5.7l2.881 4.61" />
      </svg>
    </span>
    Traverse Trees Recursively<br><br>Use recursive SQL to walk hierarchical data and produce parent‑child structures such as organization charts.
  </a>
  <a class="advanced-card" href="/sql/iterate/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M8.5 16a5.5 5.5 0 1 0 -5.5 -5.5v.5" />
        <path d="M3 16h18" />
        <path d="M18 13l3 3l-3 3" />
      </svg>
    </span>
    Iterate SQL Results<br><br>Iterate over SQL query results to execute a follow‑up query using the initial results.
  </a>
  <a class="advanced-card" href="/sql/enumerate/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M3 10l2 -2v8" />
        <path d="M9 8h3a1 1 0 0 1 1 1v2a1 1 0 0 1 -1 1h-2a1 1 0 0 0 -1 1v2a1 1 0 0 0 1 1h3" />
        <path d="M17 8h2.5a1.5 1.5 0 0 1 1.5 1.5v1a1.5 1.5 0 0 1 -1.5 1.5h-1.5h1.5a1.5 1.5 0 0 1 1.5 1.5v1a1.5 1.5 0 0 1 -1.5 1.5h-2.5" />
      </svg>
    </span>
    Enumerate Values<br><br>Assign incremental numbers to rows for ordering, labeling, or sequence‑based logic.
  </a>
  <a class="advanced-card" href="/sql/counts/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M4 12v-3a3 3 0 0 1 3 -3h13m-3 -3l3 3l-3 3" />
        <path d="M20 12v3a3 3 0 0 1 -3 3h-13m3 3l-3 -3l3 -3" />
        <path d="M11 11l1 -1v4" />
      </svg>
    </span>
    Substitute Counts<br><br>Automatically substitute cluster, subcluster, and row counts into labels and sorting attributes.
  </a>
  <a class="advanced-card" href="/sql/concatenation/">
    <span class="icon">
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path stroke="none" d="M0 0h24v24H0z" fill="none" />
        <path d="M12 4l-8 4l8 4l8 -4l-8 -4" />
        <path d="M4 12l8 4l8 -4" />
        <path d="M4 16l8 4l8 -4" />
      </svg>
    </span>
    Concatenate Values<br><br>Combine multiple fields or computed values into a single label, often used together with iteration.
  </a>
</div>

<style>
.advanced-card,
.advanced-card:visited,
.advanced-card:hover,
.advanced-card:active {
  text-decoration: none !important;
}

.advanced-card {
  transition: transform 0.15s ease, box-shadow 0.15s ease;
}

.advanced-card:hover {
  transform: translateY(-3px);
  box-shadow: 0 4px 12px rgba(0,0,0,0.12);
}
</style>
