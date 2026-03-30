# SQL Extensions

SQL extensions provide small, declarative utilities that simplify common graph‑building tasks in the **Relationship Visualizer**.  

Each extension is activated by passing specific values as SQL parameters, enabling expressive diagrams with minimal query logic.

## Extension Index

| Extension | Description |
| --------- | ----------- |
| [Directives](/sql/directives/) | Lightweight commands enabling optional behaviors within the SQL pipeline. |
| [Clustering](/sql/clustering/) | Group related rows into clusters or subclusters to visually organize sections of your graph. |
| [Count Substitution](/sql/counts/) | Automatically substitute cluster, subcluster, and row counts into labels and `sortv` attributes. Useful for sorting data. |
| [Splitting Labels](/sql/labelsplit/) | Split long text labels into multiple lines for improved readability. |
| [Chaining Nodes](/sql/chaining/) | Generate edges between sequential nodes, creating simple chains. Useful for timelines, or ordered flows. |
| [Creating Subgraphs](/sql/subgraphs/) | Wrap selected nodes into ranked subgraphs to control layout, alignment, and visual grouping. |
| [Tree Traversal](/sql/recursion/) | Use recursive SQL to walk hierarchical data and produce parent‑child structures such as organization charts. |
| [Iteration](/sql/iterate/) | Iterate over SQL query results to execute a follow-up query using the initial results. |
| [Enumeration](/sql/enumerate/) | Assign incremental numbers to rows for ordering, labeling, or sequence‑based logic. |
| [Concatenation](/sql/concatenation/) | Works in conjunction with iteration to build labels by concatenating multiple fields or computed values of a query. |
