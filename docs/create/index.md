# Creating Graphs with Excel to Graphviz

Building a graph with the **Relationship Visualizer** spreadsheet is a simple, structured process. You start by learning the basic vocabulary, then create a workbook, enter your data, apply style, and finally publish or refine the output. This page introduces the full workflow and links to the focused topics that explain each step in detail.

![](./xls2gvworkflow.png)

## Begin with the core concepts  

Before creating your first graph, it helps to understand the basic terms—nodes, edges, labels, styles, and views, as well as the various Graphviz graph layouts. These concepts shape how your data becomes a diagram.  
- [Terminology](/terminology/)

## Create a New Workbook

Every graph begins with a `Relationship Visualizer.xlsm` workbook. It provides the worksheets, formulas, and macro automation that turn your spreadsheet entries into Graphviz output.  
- [Start Here!](/prepare/)

## Understand the `data` worksheet  

The `data` worksheet is where you define the nodes and edges that make up your graph. Each row represents a relationship, and each column contributes meaning—labels, types, styles, and more. This sheet is the foundation of every diagram you build.  
- [`data` Worksheet](/dataworksheet/)
- [The `Graphviz` Ribbon Tab](../dataworksheet/#the-graphviz-ribbon-tab) 
  
## Watch the graph form  

Once the workbook and `data` sheet are ready, you can begin entering your node and edge data. As you type, the tool generates a live preview of the graph, making it easy to refine structure and relationships as you go. 
- [Type data, See graph](/coreconcepts/)

## Add styling

With the structure in place, you can shape the visual presentation. Styles let you control color, layout, grouping, and emphasis—helping your graph tell a clearer story.  
- [Add style to graph elements](/addstyle/)  
- [Design reusable styles](/designer/)  
- [Save styles in a style gallery](/styles/)  
- [Create style-driven views](/views/)

## Publish your finished graph  

When your graph looks the way you want, you can export it as an image, PDF, or SVG. These formats are ideal for documentation, presentations, and sharing with others.  
- [Publish Graphs](/publish/)

## Enhance your SVG output  

If you choose SVG, you can take advantage of its flexibility to add animation, interactivity, or additional styling using SVG post-processing tools.  
- [Post‑process SVG Files](/svg/)

## Explore advanced techniques

Once you’re comfortable with the basics, you can dive into advanced Graphviz topics, and more powerful features such as SQL‑driven graph generation, JSON exchange, and other configuration capabilities.

- [Advanced Graphviz Capabilities](/advanced/)
- [SQL-driven Graph Generation](../sql/)
- [JSON Data Exchange](../exchange/)
- [View Graphviz DOT Source](../source/)
- [Capture Graphviz CLI Messages](../console/)
- [Configure Spreadsheet Settings](../settings/)