---
blogPost: true
title: "RV-KGF: Relationship Visualizer's Knowledge Graph Format"
description: "Relationship Visualizer can now export your diagram's relationships as clean, AI-ready JSON before any of the meaning gets compiled away into a color or shape."
date: 2026-09-13
author: jjlong150
tags: ['knowledge-graph', 'ai', 'json', 'v11.0']
published: true
readingTime: true
sidebar: false
---

**Relationship Visualizer** has always done one thing: turn an Excel spreadsheet of relationships into a Graphviz diagram. For eleven years, that's meant two outputs, the DOT file and the rendered picture. Now there's a third: **RV-KGF**, a JSON export designed to be read by AI tools, graph databases, and anything else that wants the *facts* in your graph without having to look at a picture of them.

I've published the full specification here: **[github.com/jjlong150/rv-kgf](https://github.com/jjlong150/rv-kgf)**.

## The problem with handing an AI your Graphviz DOT file

Here's the thing about DOT: by the time your data gets to an AI agent, it's already been styled. A row that says *"gets from via REST over HTTPS"* becomes a solid black line. A row that says *"gets from via CGI over HTTP"* becomes a red, dashed one. That's exactly what you want for a picture, but if you hand that DOT file to an AI and ask "which of these connections are insecure?", you're asking it to reverse-engineer meaning from a dash pattern it never should have had to.

The frustrating part is that the meaning was sitting right there in the workbook the whole time, as plain, readable text, one step before it got turned into a color. So the fix isn't clever AI interpretation of a diagram, it's just: **export the data in a structured format, without visual style elements.**

That's the whole idea behind RV-KGF. It states what's semantically *true*, such as labels, relationships, containment, tooltips, and style *names* as references to their meaning. At the same time, it deliberately leaves out what a renderer should *do* such as resolved colors, coordinates, layout choices. 

## Three artifacts, one source of truth

Nothing about your existing workflow changes (see [From Spreadsheet to Knowledge Graph](../posts/knowledge-graph-export.md) for the full walkthrough). From the same worksheet, the Publish action now produces:

- **The rendered image** - the visual representation of the graph for human consumption.
- **The DOT file** - the Graphviz instruction set, if you want to re-render, hand-tune it, or source-control it.
- **The JSON (RV-KGF)** - for AI, analysis, or a future converter to reason over.

There's also a small but genuinely useful addition on the `styles` worksheet: a **description** column. Write a sentence explaining what "Call Center Agent" or "Risky Dependency" actually *means*, once, and it shows up in the exported JSON as documentation for that category. Think of it as the text equivalent of the legend on your diagram, minus the color swatches.

## Why not just use an existing graph format?

I looked. Seriously, I searched and this is what I could find: JGF, the D3/NetworkX node-link shape, Cytoscape.js, GraphSON, JSON-LD/RDF, the Property Graph Exchange Format (PG-JSON), even GQL, the new ISO graph query language standard. 

None of them fit cleanly, for reasons that were specific enough to be worth writing down. The short version: RDF doesn't treat relationships as first-class objects with their own rich, repeatable attributes, which is a hard requirement here; the closer matches (PG-JSON, Cytoscape.js) bring along conventions from a different domain that don't map cleanly onto a spreadsheet-and-styles-worksheet source of data.

::: tip Great Minds Think Alike
One nice surprise: PG-JSON's own spec independently recommends the same "synthesize a stub node for an edge that points at something never declared" rule that this format arrived at on its own. Two people solving the same interoperability problem landing on the same answer, without either one copying the other, feels like decent evidence the rule is actually right and not just a personal preference. 
:::

The full comparison, format by format, is in the [industry-comparison doc](https://github.com/jjlong150/rv-kgf/blob/main/docs/industry-comparison.md) if you want the details.

## Where to go from here

- **Read the spec:** [github.com/jjlong150/rv-kgf](https://github.com/jjlong150/rv-kgf) contains the schema reference, design rationale, the industry comparison, a JSON Schema for validation, and three worked examples.
- **Try it:** grab the [latest workbook](https://exceltographviz.com/download/), click **Knowledge Graph** in the ribbon, and view the resulting file in the browser window which appears.
- **Build something with it:** the schema is public and stable at version 1.0. If you write a tool to consume it, I'd genuinely like to hear about it, and I'm happy to link to it from the spec repo.

As always, this is a **free** tool built and maintained by **one person in his spare time**. If this tool accelerates your AI journey, consider [buying me a coffee](https://www.buymeacoffee.com/exceltographviz). Eleven years and 10,000+ downloads in, we're still only at three. 😔

If a coffee is not in the budget, [leaving a 5-star review on SourceForge](https://sourceforge.net/projects/relationship-visualizer/reviews/new) costs nothing but a minute of your time, and helps more than coffee does. As search shifts from links toward AI-generated answers, reviews like these are becoming one of the few signals that still help people actually find a tool like this one.
