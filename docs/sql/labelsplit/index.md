---
title: Split Long Labels for Better Readability
description: Use SQL extensions to split long labels into multiple lines, improving clarity and layout in your diagrams.
---

# Split Long Labels

Label splitting is a powerful SQL extension that allows you to break long label text into multiple readable lines within a node. It gives you control over where lines break and how each line is aligned (left, center, or right), resulting in cleaner, more balanced, and professional-looking diagrams.

### When to Use Label Splitting

- Long titles or descriptions that make nodes too wide
- Improving readability in dense diagrams
- Creating consistent node widths across different amounts of text

### How It Works

Add two columns to your SQL query

| Column Name      | Description                         |
| ---------------- | :---------------------------------- |
| `SPLIT LENGTH`   | Desired maximum characters per line |
| `LINE ENDING`    | Line break character and alignment  |

as follows:

```sql
        '5'  as [SPLIT LENGTH],
        '\n' as [LINE ENDING],
```

In this example, the label will be split into multiple lines at boundaries as close as possible to 12 characters. Splits occur only at spaces, so any word longer than 12 characters will remain unbroken.

Line endings can be any string<sup>[1]</sup>. The most commonly used line endings are:

| Line Ending | Meaning / Usage                           |
| :----------:|------------------------------------------ |
| `\n`        | New line with center alignment            |
| `\r`        | New line with right alignment             |
| `\l`        | New line with left alignment              |
| `\|`        | Pipe delimiter (useful for Record shapes) |
| `<br/>`     | HTML line break (for HTML‑like labels)    |

<sup>[1]</sup> New line `\n` is the default if `SPLIT LENGTH` is specified, but `LINE ENDING` is omitted.

Here is an example:

*Before:*

| This is a very long label that stretches the node |
| :-------------------------------------------: |

*After:*

| This is a very<br/>long label<br/>that stretches<br/>the node |
| :-------------------------: |


