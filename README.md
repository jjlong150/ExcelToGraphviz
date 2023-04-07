# Excel to Grapviz

The **Relationship Visualizer** is a macro-enabled **Microsoft Excel** spreadsheet that facilitates the collection of relationship information. It works in conjunction with **Graphviz**, which is open-source graph visualization software. Graphviz's strength is the ability to generate diagrams programmatically. To fulfill this aim, Graphviz uses a simple but powerful graph description language known as DOT.

The Relationship Visualizer removes much of the burden of understanding the DOT language. It allows you to express relationships through text in Excel rows and columns. Macros in the spreadsheet write the row and column data in DOT format into a text file, and then invoke Graphviz to read the text file and interpret the commands to create the graph. The resulting Graph is then displayed within Excel.

<center>

[Documentation](/docs/README.md)

[MIT License](/LICENSE)

**Copyright (c) 2017-2023 Jeffrey J. Long**

</center>
