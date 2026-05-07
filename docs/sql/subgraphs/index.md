# Align Nodes on the Same Level

Ranking is useful whenever you want certain nodes to align visually in a consistent row or column. By placing nodes into a shared subgraph with a defined rank, you can control the layout and make structural relationships easier to understand. Common scenarios include:

- **Grouping Peers:** Aligning team members, departments, or sibling categories on the same horizontal level to emphasize equal standing.
- **Highlighting Stages:** Showing phases of a project or lifecycle (e.g., Planning, Execution, Review) on a single rank for clarity.
- **Organizing Layers:** Keeping all “input” nodes at the top, “processing” nodes in the middle, and “output” nodes at the bottom.
- **Comparing Alternatives:** Placing multiple options or branches side‑by‑side so users can visually compare them.
- **Clarifying Hierarchies:** Ensuring that children of the same parent appear on the same rank, preventing uneven or zig‑zag layouts.
- **Creating Swimlanes:** Using ranks to form horizontal lanes that separate categories, roles, or functional areas.
- **Stabilizing Layouts:** When Graphviz’s automatic layout shifts nodes unpredictably, explicit ranks help lock the structure into a predictable shape.

These scenarios benefit from the `CREATE RANK` extension because it gives you precise control over how nodes align, making your graphs cleaner, more readable, and more intentional.

Assume you have an Excel workbook containing a worksheet named `Alphabet` with a column heading `letter` and four rows of data: A, B, C, and D. You want all of these nodes to appear on the same rank.

The **CREATE RANK** SQL extension produces subgraphs whose nodes share the same rank.

The SQL is specifed as follows:

```sql
SELECT [letter] AS [Item], 
      TRUE      AS [CREATE RANK], 
      'same'    AS [RANK] 
FROM [Alphabet$]
```

The SQL above results in one row added to the `data` worksheet with: 
- Item = `>`
- Label = `{rank="same"; "A"; "B"; "C"; "D";}`

The values you can pair with `RANK` are `same`, `min`, `source`, `max`, and `sink`. The Graphviz `rankdir` attribute appears to treat these values as case‑sensitive, so be sure to specify them in **lowercase**.

