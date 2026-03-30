# Grouping Data into Clusters and Subclusters

One extension adds the ability to scan SQL statements for specific field names that signal grouping values to be wrapped into a **cluster** or **subcluster** (a cluster within a cluster). The syntax remains pure SQL, but certain field names are reserved and given special meaning.

Clusters and subclusters can specify:

- A mandatory *cluster column*. The presence of this column triggers the additional processing needed to insert cluster braces into the result set.
- A **Label**, which allows you to override the label derived from the cluster column. If no label is provided, the value of the cluster column is used.
- A **Style Name**, which must correspond to a style defined on the `styles` worksheet. This can be:
  - a static string  
  - a composite string using substitution or concatenation  
  - or a column value that matches a style name  
  A common example appears in heat‑map scenarios, where values such as `Critical` and `Standard` map to styles of the same name, each with its own fill color.
- **Attributes**, which allow you to include line‑specific attribute values.
- **Tooltip**, which provides hover text in SVG output.

Although the examples below use uppercase for clarity, the implementation is case‑insensitive. Variants such as `CLUSTER`, `Cluster`, `ClUsTeR`, and `cluster` are all treated identically.

| Column                               | Cluster Field Names        | Subcluster Field Names        |
| ------------------------------------ | --------------------------- | ------------------------------ |
| *cluster column (i.e., group by)*    | `CLUSTER`                   | `SUBCLUSTER`                   |
| **Label**                            | `CLUSTER LABEL`             | `SUBCLUSTER LABEL`             |
| **Style Name**                       | `CLUSTER STYLE NAME`        | `SUBCLUSTER STYLE NAME`        |
| **Attributes**                       | `CLUSTER ATTRIBUTES`        | `SUBCLUSTER ATTRIBUTES`        |
| **Tooltip**                          | `CLUSTER TOOLTIP`           | `SUBCLUSTER TOOLTIP`           |

