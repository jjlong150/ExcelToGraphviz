---
prev: /sql/iterate/
next: /sql/timeline/
---

# Enumeration Queries

## Introduction

*Relationship Visualizer* Version 10.0 added support for **enumerated SQL queries**, enabling a query to loop through a numeric range and generate rows for each value in that range. Enumeration is especially useful when you need to fill in missing values, create evenly spaced steps, or build a complete sequence that does not exist in the underlying data.

::: tip Iterate vs. Enumerate?

Within Relationship Visualizer, a distinction is made between **iteration** and **enumeration**:

- **Iterate** refers to stepping through each item in a result set one by one — without tracking position or count. Each item becomes input for a follow-up SQL query.

- **Enumerate** refers to iteration **with explicit tracking of position, index, or count**, using a counter that increments from a starting value to a stopping value. Enumeration is used when the logic depends on the numeric sequence itself. For example, “generate rows from 1 to 12,” “step through values by 5,” or “fill in missing years.”

Enumeration is not about iterating over data — it’s about generating **new** data based on a numeric range.
:::

## Problem Scenario

In the [Iteration Queries](../iterate/README.md) example we built a timeline of Unix shell introductions. The dataset contains a `Year` column, but the years are sparse:
- Some years have shells.
- Many years do not.

The following query creates edges from one year to the next **only for years that appear in the data**:

``` sql
SELECT DISTINCT 
       ([Year] & '_id') AS [ITEM],
       TRUE AS [CREATE EDGES],
       [Year]
FROM [Shells$]
WHERE [Year] IS NOT NULL
ORDER BY [Year] ASC
```

If we create edges directly from the data, the timeline will “jump” over missing years, compressing the visualization and making the evolution appear faster than it actually was.

## Fixing the Gaps with Enumeration

Relationship Visualizer supports **enumerated SQL queries** through four special fields: `ENUMERATE`, `START AT`, `STOP AT`, and `STEP BY`. When these appear in the `SELECT` list, the engine switches into enumeration mode and generates a numeric sequence independent of any table data.

Enumeration is the correct tool when you need a **complete, gap‑free range** such as years, months, sequence numbers, or evenly spaced steps. Instead of relying on whatever values happen to exist in the dataset, enumeration produces a deterministic backbone that fills in all missing points.

## Basic Enumeration Syntax

The simplest enumerated query defines:

- `TRUE AS [ENUMERATE]`  - Indicates that enumeration is desires/active  
- `... AS [START AT]` - The starting value  
- `... AS [STOP AT]` - The **inclusive** stopping value  
- `... AS [STEP BY]` - The increment 
- How each generated step should appear in the output. 
  
  Within any text field, the placeholder `{step}` is replaced with the current numeric value. This allows you to generate IDs, labels, keys, or synthetic attributes without referencing any underlying table.

The result is one row per step, covering the entire numeric range you specify.

## Examples

### Example 1: Create the nodes from year to year using a hard-coded range

Enumeration does not depend on any table — it **creates** data. You can then join, merge, or visualize this generated sequence alongside your real dataset.

``` sql
SELECT TRUE AS [ENUMERATE], 
       1965 AS [START AT], 
       1995 AS [STOP AT], 
       1 AS [STEP BY], 
       '{step}' AS [Item], 
       '{step}' AS [Label],
       'Year'   AS [Style Name]
```

This SQL creates the nodes, and applies a style named **Year** from the `styles` worksheet to the node which is created.

### Example 2: Create the edges from year to year using a hard‑coded range

This example shows how to generate a complete sequence of years using a fixed numeric range. By enabling both `CREATE EDGES` and `ENUMERATE`, the engine produces one row per year and automatically creates edges between consecutive values.

Because the range is hard‑coded, the query does not depend on any underlying table. It simply emits a synthetic timeline from the starting year up to (but not including) the stopping year. Each iteration replaces `{step}` with the current year, allowing you to generate IDs, labels, or edge endpoints directly from the loop counter.

This pattern is useful when you want a predictable, fully populated timeline regardless of what data exists in the workbook.

``` sql
SELECT TRUE AS [CREATE EDGES], 
       TRUE AS [ENUMERATE], 
       1965 AS [START AT], 
       1995 AS [STOP AT], 
       1 AS [STEP BY], 
       '{step}' AS [Item]
```

### Example 3: Create nodes for the minimum year through the maximum year

This example shows how to generate a complete sequence of years based on the actual data in the workbook. Instead of hard‑coding the range, the query calculates the minimum and maximum year values directly from the table and uses them as the bounds for enumeration.

By doing this, the generated sequence automatically adapts to whatever data is present. If new shells are added with earlier or later years, the enumerated range expands accordingly. The result is a flexible, data‑driven timeline that always covers the full span of years represented in the dataset.

Each iteration replaces `{step}` with the current year, producing one node per year. These nodes can be used as standalone timeline markers or combined with edges to create a continuous, gap‑free visualization of the entire period.

``` sql
SELECT TRUE AS [ENUMERATE],
       MIN(CLng([Year])) AS [START AT],
       MAX(CLng([Year])) AS [STOP AT],
       1 AS [STEP BY],
       '{step}' AS [Item],
       '{step}' AS [Label]FROM [shells$]
WHERE IsNumeric([Year])
```

### Example 4: Create edges for the minimum year through the maximum year

This example combines data‑driven bounds with automatic edge creation. By enabling both `ENUMERATE` and `CREATE EDGES`, the engine generates a continuous sequence of years based on the minimum and maximum values found in the dataset, and then creates edges between each consecutive pair.

Because the range is derived from the actual data, the resulting timeline always expands or contracts to match the earliest and latest years present in the table. This makes the query fully adaptive: adding new rows with earlier or later years automatically updates the generated edges.

Each iteration substitutes `{step}` with the current year, producing a clean, gap‑free chain of edges that spans the entire period represented in the dataset. This is the most flexible way to build a complete chronological backbone for visualization.

``` sql
SELECT TRUE AS [CREATE EDGES],
       TRUE AS [ENUMERATE],
       MIN(CLng([Year])) AS [START AT],
       MAX(CLng([Year])) AS [STOP AT],
       1 AS [STEP BY],
       '{step}' AS [Item]
FROM [shells$]
WHERE IsNumeric([Year])
```

## Try it Yourself

This example is included in the samples in the Relationship Visualizer zip file in the directory `17 - Using SQL - Enumeration`.

## Summary

Enumeration gives Relationship Visualizer a way to generate data that does not exist in any table, making it possible to build complete sequences, fill gaps, and create smooth visual structures such as timelines. By defining a numeric range with `START AT`, `STOP AT`, and `STEP BY`, you gain full control over how many rows are produced and how they are labeled. This makes enumeration ideal for scaffolding, synthetic nodes, and any situation where the visualization depends on a continuous progression of values.

Iteration and enumeration complement each other: iteration processes what the data already contains, while enumeration creates what the data is missing. Used together, they allow you to combine real and synthetic information into a single, coherent model that supports clearer, more expressive visualizations.

---

<center>

Like this tool? [Buy me a coffee! ☕](https://www.buymeacoffee.com/exceltographviz)

</center>


