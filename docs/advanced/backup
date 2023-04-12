---
prev: /publish/
next: /source/
---

# Advanced Topics

This section describes several miscellaneous features in the Relationship Visualizer that can be used to create more elaborate graphs.

## HTML Labels

Graphviz has a feature where if the value of a label attribute for nodes, edges, clusters, or graphs is given as an HTML string that is delimited by `<`...`>`, the label is interpreted as an HTML description. At their simplest, such labels can describe multiple lines of variously aligned text as provided by ordinary string labels. More generally, the label can specify a table like those provided by HTML, with different graphical attributes at each level.

The features and syntax supported by these labels are modeled on HTML. However, there are many aspects that are relevant to Graphviz labels that are not in HTML and, conversely, HTML allows various constructs which are meaningless in Graphviz. The Graphviz creators generally refer to these labels as `HTML-Like Labels` but the reader is warned that these labels are not HTML.

The grammar which Graphviz will accept is fully described at: [https://www.graphviz.org/doc/info/shapes.html\#html](https://www.graphviz.org/doc/info/shapes.html#html)

A basic HTML label can be constructed as text between `<font>` and `</font>` elements, then wrapped in the `<` and `>` delimiters as described above.

For example, a label can be constructed as:

`<<font>This label has <b>bold</b>, <i>italic\</i>, <u\>underlined</u>, and <s>strikeout</s> text</font>>`

and entered as a label value for an edge. In this example, we will relate 'a' to 'b' as we are interested in seeing how the edge is drawn. The 'data' worksheet appears as:

![](../media/c2c3a642c8f836616aceb34ab8aa8b77.png)

Pressing `Refresh Graph` produces the following graph:

![](../media/b7ac4f8c715385f1c64c2c7af465f0b0.png)

A slightly more complex example is to create a HTML table. In this example, the table contains one row with two cells:
`
<<table>

  <tr>
    <td>Cell 1</td>
    <td>Cell 2</td>
  </tr>
</table>>
`
Using it to represent a node named 'c', the 'data' worksheet appears as:

![](../media/ebaeb3c116a73f90e8e7d94927e0341c.png)

Pressing `Refresh Graph` produces the following graph:

![](../media/b658193b1cab4ed1f5227c6e97c21351.png)

HTML labels can be used for Clusters, Nodes, and Edges. In the example below there are three Items named 'a', 'b', and 'c'. HTML labels have been added for node 'a', and the edges from 'a' to 'b' and from 'b' to 'c'. The nodes and edges are wrapped with a border via a cluster that also has an HTML label.

The 'data' worksheet appears as:

![](../media/f1e23ec1945596fc74211930263c6103.png)

Pressing `Refresh Graph` produces the following graph:

![](../media/05faa32ee021c70c198d616e6a590072.png)

## Keywords

### 'Graph', 'Node', and 'Edge' Keywords

Graphviz has a built-in behavior where if a default attribute is defined using a node, edge, or graph statement, any object of the appropriate type defined afterwards will inherit this attribute value. This holds until the default attribute is set to a new value, from which point the new value is used. Objects defined before a default attribute is set will have an empty string value attached to the attribute once the default attribute definition is made.

The Relationship Visualizer also supports this capability by reserving the values "graph", "node", and "edge" as keywords in the `Item` column of the 'data' worksheet. Appropriate syntax statements are added to the DOT source code to put the styles defined by the values in the `Style Name` and `Attributes` columns into effect (when these columns are enabled on the 'settings' worksheet) when these keywords are detected.

In the following example, nodes have been defined with an Item ID of 'a' through 'h". On row 5 a _node_ statement has been placed with an `Attributes` definition which changes the font color to red (the node statement has conditional formatting which changes the cell background color and changes the font to bold italic to differentiate the keyword from ordinary data). All nodes from that point forward are rendered with a red font. This continues until a second _node_ statement is encountered on row 10 that resets the font color to a null string, which tells Graphviz to resume using the default value.

![](../media/4ef7e7631be4d525a4d28f8871d927b7.png)

Pressing `Refresh Graph` produces the following graph:

![](../media/152f5a9ec0efb8ceccf9c74b8af4be3f.png)

Likewise, this same capability exists for edges using the "edge" keyword. In the example below an edge keyword on row 13 sets the edge color to blue for the first 3 edges. A second edge keyword on row 17 changes the color to red for all the remaining edges.

![](../media/964ac013e7e8cf3ab0f12c7620fee055.png)

Produces the following graph:

![](../media/bd08dde1d48f4b1b255d69dd1336861f.png)

Note that a subgraph receives the attribute settings of its parent graph at the time of its definition. This can be useful; for example, one can assign a font to the root graph and all subgraphs will also use the font. For some attributes, however, this property is undesirable. If one attaches a label to the root graph, it is probably not the desired effect to have the label used by all subgraphs. Rather than listing the graph attribute at the top of the graph, and the resetting the attribute as needed in the subgraphs, one can simply defer the attribute definition in the graph until the appropriate subgraphs have been defined.

## Clusters

### Depicting a Relationship from or to a Cluster

You may have a situation where you are trying to represent a dependency diagram where you have nodes inside a cluster, and you want to be able to make nodes and/or clusters dependent on other nodes and/or clusters. In other words, you want an edge to begin and/or end at the border of a cluster. This goal can be accomplished in the Relationship Visualizer with a little bit of additional Graphviz knowledge.

Let's reproduce the diagram below which can be found at: <http://stackoverflow.com/questions/2012036/graphviz-how-to-connect-subgraphs>

![enter image description here](../media/38bac697f7b7f2cc647c1bd17a12298e.png)

Let's start by putting the 'graph', 'node' and 'edge' keyword features described in the previous section into a tangible example. In rows 4-6 below each keyword is specified along with style formatting for each object. These statements are not required to connect clusters, but are illustrative of the keyword feature which was just described, and will help visually distinguish elements of the graph. For graphs, nodes, and edges set the `Attributes` cells with the following formatting information:

- _Graph_: fontname="Arial" fontsize="12" fontcolor="red"
- _Node_: fontname="Arial"
- _Edge_: fontname="Arial" fontsize="8" decorate="true" color="blue"

The spreadsheet should look as follows:

![](../media/eb187d0cf7d02131ea3919dd26a91384.png)

Next we need we enable the "compound" graph attribute which Graphviz sets to 'false' by default. Setting it to 'true' activates a Graphviz feature allowing edges to connect to clusters. The easiest way to do this is to follow the instructions previously described in [Adding Native Graphviz Directives](#adding-native-graphviz-directives) to add a native Graphviz statement. Place a `>` character in the Item column to indicate this is a native command and enter compound="true" in the Label column as shown below.

![](../media/807541d3f7e6f9bbdab491874be514dc.png)

The compound="true" statement is then added to the body of the main graph.

You can also achieve the same result by checking the graph option "Allow edges between clusters"

![](../media/709e6bacd634dfd9b37923f95423461e.png)

Next, define 2 clusters. To make the example clear we will label them as cluster0 and cluster1. Cluster0 will have 4 edge relationships using the letters a, b, c, and d as node names which will fall inside the cluster border. Likewise, cluster1 will have edge relationships using the letters e, f, and g as node names that fall inside a cluster border.

Up until this point we have always defined the start of a cluster with an open brace '{'. When the Relationship Visualizer sees an open brace, it generates an internal name for the cluster to makes things simpler for the user. A cluster must have an Item name to reference it by if we are to connect edges to it, so using the macro-generated name could be haphazard, as the name will change if the cluster definition moves around within the Excel spreadsheet.

The Relationship Visualizer solves that problem by allowing you to provide a name preceding the open brace (for example, "cluster0") to designate a named subgraph. If the name of the subgraph begins with **cluster**, Graphviz notes the subgraph as a special cluster subgraph. If supported, the layout engine will do the layout so that the nodes belonging to the cluster are drawn together, with the entire drawing of the cluster contained within a bounding rectangle. Note that, for good and bad, cluster subgraphs are not part of the DOT language, but solely a syntactic convention adhered to by certain of the layout engines. If the name does not begin with **cluster,** then Graphviz treats the subgraph as an ordinary subgraph.

The spreadsheet now looks as follows to define the clusters. Note in the illustration that the occurrence of the cluster name in the Item and Label columns. We will remove the label later in this example, but for now it helps make the lesson clearer.

![](../media/c6f1884de56f89f0128e0ae492519795.png)

Once the clusters have been defined, we can specify edges that show relationships between the clusters. Five scenarios are depicted:

1.  A relationship from a node within a cluster to another node within a cluster (row 21)
2.  A relationship from a node within a cluster to a node outside of a cluster (row 22)
3.  A relationship from a node within a cluster to the border of another cluster (row 23)
4.  A relationship from the border of a cluster to a node within a cluster (row 24)
5.  A relationship from the border of a cluster to the border of another cluster (row 25)

To specify the cluster were the edge should originate from you must specify a ltail attribute in the Attributes column which specifies the name the cluster the edge should originate from. For example, ltail="cluster0". Note that the item name in the "Item Name" column must reside within that cluster.

When you want the arrowhead of an edge to stop at the border you must specify a lhead attribute in the Attributes column which specifies the name of the cluster to connect to. For example, lhead="cluster1". Note that the item name in the `Related Item` column must reside within that cluster.

The spreadsheet appears as follows to show the five scenarios described above. Descriptions of the scenarios have been added in the Label column to help explain this example. Later in the example we will hide them.

![](../media/4579e4387c7760f7be58ad342be411fc.png)

At this point all the information necessary to create the graph has been entered. When you press the `Refresh Graph` button the graph appears as follows:

![](../media/a88879ebb44e1b6cd8ccabff8b39d3c5.png)

The graph looks like our desired graph. Notice that each blue edge has a black line attached which underlines the label text, and the edge label describes the relationship depicted. The callout line effect was achieved by us adding the attribute decorate="true" when the edge keyword was defined at the top of the spreadsheet. It was useful here to help see the scenario depicted as Graphviz sometimes places labels in locations which can cause confusion.

Now that we see how the edge statements are graphed, we can hide the labels. An easy way to do that is to go to the 'Graphviz ribbon tab, and in the "Edge Options" remove the check box from **Include Label"**.

![](../media/6d22c48e76c441e7ec2c7aee41d8f1c7.png)

Press the `Refresh Graph` button and the graph now appears as:

![](../media/b934f086f332c54294af8c3aaeb27082.png)

This graph is almost the same as the graph from the Internet we are trying to duplicate. The only thing left to do is to remove the cluster labels that were included to illustrate this example. There is not a settings switch to turn these on/off, so here we must go back into the 'data' worksheet and delete the labels so that the data looks as follows:

![](../media/5817f8428af8229102c6a96cf35745c9.png)

Press the `Refresh Graph` button and the graph now appears as the graph on the left; the graph we are duplicating is on the right:

| ![](../media/946dada8e36d723484e4a444f081507a.png) Generated by Relationship Visualizer | ![enter image description here](../media/38bac697f7b7f2cc647c1bd17a12298e.png) <http://i.stack.imgur.com/Ka0t2.png> |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |

The graph was purposely left with blue edges and a slightly smaller font size to differentiate it enough from the target graph to show it was not the original image but kept similar enough to show the goal was met.

If you are interested in making the final changes to make the graphs truly identical you can edit the style definition for the edge keyword on row 5 to remove the color="blue" attribute. Press the `Refresh Graph` button and the graph generated will be identical to the goal image.

### Aligning Nodes Across Clusters

One of the ways Graphviz frees your time is by letting it choose the optimal way to lay out shapes and edges. Sometimes however you want to control the placement for esthetic reasons. Assume that you have the following graph:

![](../media/3d81e1b9fe4b2df33d8f9c5ab20af698.png)

Created by this spreadsheet:

![](../media/77e5a736be8e77659ba2a5414347bcc0.png)

For esthetic reasons, we would like "router1" and "router2" to be aligned. Two native Graphviz DOT commands need to be added to the spreadsheet to make this occur. The first command is shown on line 4, which adds newrank="true" into the body of the main graph. (It appears that newrank is an undocumented attribute added in Graphviz 2.30 which activate a new ranking algorithm which allows defining rank="same" for nodes which belong to clusters).

![](../media/7c9dadf28c48fdbcca1b1147a6718fa2.png)

You can also achieve the same result by checking the graph option "Rank ignoring clusters"

![](../media/709e6bacd634dfd9b37923f95423461e.png)

The next step is to add the following native Graphviz command after the cluster definitions:

{ rank="same"; "router1"; "router2"; }

as shown on row 13 below:

![](../media/aafa7c6e8fc43807d89fb7bf9fe4a622.png)

Press the `Refresh Graph` button, and the graph contains router1 and router2 aligned as shown below:

![](../media/01afe6b67fec07cfdc4ae0a04c579a41.png)

## Nodes

### Shape="Record"

Visually, a record is a box, with fields represented by alternating rows of horizontal or vertical sub-boxes. The Mrecord shape is identical to a record shape, except that the outermost box has rounded corners. Flipping between horizontal and vertical layouts is done by nesting fields in braces "{...}". The top-level orientation in a record is horizontal. Thus, a record with label "A \| B \| C \| D" will have 4 fields oriented left to right, while "{A \| B \| C \| D}" will have them from top to bottom and "A \| { B \| C } \| D" will have "B" over "C", with "A" to the left and "D" to the right of "B" and "C".

The initial orientation of a record node depends on the rankdir attribute. If this attribute is "Top to Bottom" or "Bottom to Top", corresponding to vertical layouts, the top-level fields in a record are displayed horizontally. If, however, this attribute is "Left to Right" or "Right to Left", corresponding to horizontal layouts, the top-level fields are displayed vertically.

![](../media/f570f7c0a2e3d522188252e493674052.png)

Results in:

![](../media/5290228d36d236fc0245ded55ce08583.png)

### Shape="Record" With Ports Specified

You can specify port identifiers as part of the field values in a record shape. The first string in the field Id assigns a port name to the field and can be combined with the node name to indicate where to attach an edge to the node.) The second string is used as the text for the field; it supports the usual escape sequences \\n, \\l and \\r. Therefore, if we specify the following (with color-coding added to highlight the ports):

![](../media/4d7651843c044f5e0af0d14fd312e4f8.png)

We will generate the following graph:

![](../media/35a68818ae9eb5f64e5a0e2313974f05.png)

If we change the graphing direction from "Top to Bottom" to "Left to Right" on the 'Settings' worksheet and regenerate the graph it will appear as follows:

![](../media/7d6168255c5aec45d705f6c6c19beb63.png)

### Shape="Polygon"

Polygon shapes are unique from other shapes in Graphviz and have extra attributes which control how the polygon is created.

The easiest way to learn about polygon shapes is to use the 'style designer'. The default 'Shape' section of the style designer ribbon (Element Type = Node) consists of a single dropdown list for "Node Shape". If you select 'polygon' as the shape the ribbon will change dynamically to present additional choices as shown below:

![](../media/d066b8b2a73eb688bfed8947d9b58f2b.png)

Selecting 'polygon' changes the ribbon to appear as:

![](../media/17a40245c67a327319ee395c204763cd.png) ![](../media/e10bb1aee76e6a2154ced671853fee78.png)

shape="polygon"

**Regular Polygon** - If true, forces the polygon to be regular, i.e., the vertices of the polygon will lie on a circle whose center is the center of the node.

![](../media/19ddaa59d9427359638faa8dea79e2bf.png) ![](../media/b1f9fa316f292be0fb8d602ae4a77866.png)

shape="polygon" regular="Yes"

**Polygon Skew** - Positive values skew top of polygon to right; negative values skew the top of the polygon to the left.

_Positive Skew_

![](../media/6e7d19776d7acf3824922dd1fa3743cd.png) ![](../media/dc4cfbe9033894c44f26b7011d5ccdc4.png)

shape="polygon" skew="1" regular="No"

_Negative Skew_

![](../media/c4b344b1b54c83a2fe2ebdd5b57ccb77.png) ![](../media/1bc5be00cf87ddb3b4899d29550a465e.png)

shape="polygon" skew="-1" regular="No"

**Polygon Distortion** - Positive values cause top part of the polygon to be larger than bottom; negative values do the opposite.

_Positive Distortion_

![](../media/ad42c3bc93b78627c6ad8c49e2ff1835.png) ![](../media/d3a16d0b5a88e15e39c7af65c7b96df1.png)

shape="polygon" distortion="1" regular="No"

_Negative Distortion_

![](../media/c3a60f67f68f3bcd01954cb9281d8bf8.png) ![](../media/6b48aeeda4f98ec38d07b5c4f4ff5f15.png)

shape="polygon" distortion="-1" regular="No"

**Combining Skew with Distortion**

| ![](../media/131fde8d0c21cbde937f364e790d1251.png) | ![](../media/cf4e5073a7590e1b3aaa95805918337d.png) | ![](../media/3e48a62a8c4e79fcd6db37bef589d1bf.png) |
| -------------------------------------------------- | -------------------------------------------------- | -------------------------------------------------- |
| Skew -1, Distortion 1                              | Skew 0, Distortion 1                               | Skew 1, Distortion 1                               |
| ![](../media/e89d2bd25615a593144342db8bc4cd95.png) | ![](../media/085340ed0d232965f7cf3bfd96545943.png) | ![](../media/b902f48209f1013c10632a421a6028d3.png) |
| Skew -1, Distortion 0                              | Skew 0, Distortion 0                               | Skew 1, Distortion 0                               |
| ![](../media/1af1b897c26ed637e15a9837381b48e3.png) | ![](../media/8b96cb5b58691ead550f6659056cb5e9.png) | ![](../media/1c8348e5974113e435f37510ec0553f8.png) |
| Skew -1, Distortion -1                             | Skew 0, Distortion -1                              | Skew 1, Distortion -1                              |

**Number of Sides** - Number of polygon sides.

![](../media/b0f09aeabc1ad08cd2192a46b6625063.png) ![](../media/3f571210fe7626b7f0ab1375cf89e992.png)

`shape="polygon" sides="8" regular="Yes"`

**Polygon Rotation** - Angle, in degrees, to rotate polygon node shapes. For any number of polygon sides, 0 degrees rotation results in a flat base.

![](../media/56a290ec010794994c801f2c90b3af22.png) ![](../media/894f7d55abd6bcfa11b9031aa33152b8.png)

`shape="polygon" sides="8" regular="Yes" orientation="21"`

### Skewing the Angle of Polygon Nodes

Polygon shapes can be angled by changing the skew attribute. Positive values skew top of polygon to right; negative to left. This feature can be illustrated simply by defining seven nodes and varying the skew attribute by 0.5 for a range of values from -1.5 to 1.5 in the 'Attributes' column.

The Excel data is defined as follows:

![](../media/addae74ce647877cd6ed22f062a130ec.png)

Pressing the `Refresh Graph` button, the graph appears as:

![](../media/62b4b80912af0f6f30521b0325fcc464.png)

This feature works for any number of polygon sides. If the sides attribute on row 3 changes from sides="4" to sides="8", the resulting graph appears as:

![](../media/14ebc5e046312002700e254ce2120513.png)

### Distorting the Length of Polygon Nodes

Polygon shapes can be distorted by changing the distortion attribute. Positive values cause the top part of the polygon to be larger than the bottom; negative values do the opposite. This feature can be illustrated simply by defining seven nodes and varying the distortion attribute by 0.5 for a range of values from -1.5 to 1.5 in the 'Attributes' column.

Define the Excel data as follows:

![](../media/6b350287450088b1a409f098c5f29f00.png)

Pressing the `Refresh Graph` button, the graph appears as:

![](../media/8ee9e73fd2c818b14860c59d2696347e.png)

This feature works for any number of polygon sides. If the sides attribute on row 3 changes from sides="4" to sides="8", the resulting graph appears as:

![](../media/79fc67b5adbde446edd657561ad71ff4.png)

## Edges

### Consolidating Edges Using the 'strict' Option

Some sets of data will cause multiple edges to be drawn between the same nodes. A good example of this is the US state border example from the introduction of this document. Every state that shares a border with another state has two relationships. For example, Michigan shares a border with Ohio, and Ohio shares a border with Michigan.

Plotting the state border relationships as an undirected graph, the following data:

![](../media/7b79f76ae06c0ab578bc127cf78d0de5.png)

Generates the following graph:

![](../media/6fd584d0140a729572678ed24f8f86c3.png)

Graphviz can consolidate these duplicate relationships into one relationship. If Graphviz is told that the graph is strict, then multiple edges are not allowed between the same pairs of nodes.

On the `Graphviz` ribbon tab in the "Edge Options" section check the 'Apply "strict" rules' option to "yes" as shown below:

![](../media/e0b39f071dd7bc76462bb6257b7578e2.png)

Press the `Refresh Graph` button and the graph appears as:

![](../media/584a8390e58ebf90511d84c64845c613.png)

### Changing the Order of Edges

Differences in style may be achieved by altering the edge ordering. If the value of the ordering attribute is "out", then the out edges of a node, that is, edges with the node as its tail node, must appear left-to-right in the same order in which they are defined in the input. If the value of the attribute is "in", then the in edges of a node must appear left-to-right in the same order in which they are defined in the input.

Assume you have several edge relationships defined as follows:

![](../media/8b5aa88eca79eff566644963a91891b6.png)

Pressing the `Refresh Graph` button, the graph appears as:

| ![](../media/d7a4cde94943bc529c961c7e5c894b92.png) ordering="in" | ![](../media/a32db80c779d71851d9319af65db97e3.png) ordering="out" |
| ---------------------------------------------------------------- | ----------------------------------------------------------------- |

### Placing a Label at the Head or Tail of an Edge

The default view of the Relationship Visualizer provides a Label column for labeling an edge. Graphviz also supports placing a label at the tail and/or head of an edge via the taillabel and headlabel attributes respectively. There are label columns on the 'data' worksheet which correspond to these attributes, however you must unhide the columns to use them.

For example, if we have a simple relationship such as:

![](../media/5a90333147e9650d3eff895603a41624.png)

Producing the graph:

![](../media/0ca48224b3c95eec6a58f05567021984.png)

We can click the "Show/Hide Columns" dropdown list on the Graphviz tab to expose the additional label columns

![](../media/a63ebf8eb6aabda9bad35af7520ba959.png) ![](../media/32cf186f27c041c563555ae45b966edd.png)

The data worksheet now appears as:

![](../media/7ba1f502f980e5b33f52410ba1d7ba23.png)

If we place the value "tail" at the tail of an edge, and "head" at the head of an edge, the data in the spreadsheet would look as follows:

![](../media/555dd1d1d4c6dd1d67fe3e4ade03bcd1.png)

Pressing the `Refresh Graph` button creates the following graph:

![](../media/70a7d9f16a32837c53456f9444a50d43.png)

### Drawing an Edge from or to the Center of a Node

An edge is clipped to the boundary of a node shape by default. It is possible to override this behavior such that the edge will begin and/or end in the center of the node instead of the node boundary. The attributes headclip and tailclip control this behavior. When set to true (the default) an edge is clipped to the boundary of a node. When set to false the edge goes to the center of the node, or the center of a port if applicable.

The data below:

![](../media/f4fb8dad441c398549015866ad48a36e.png)

Creates the following graph:

![](../media/90edf33f0960fff213f53a8639277e0e.png)

**Note:** _The nodes have blank labels to make the illustration of edges coming from or going to the center of the node easier to see. Enabling the 'Nodes' graph option 'When the `Label` column is blank…' '…use blank for the node label" on the `Graphviz` ribbon tab is required to achieve this effect._

![](../media/46537ed18a2ab0998b1c613966aa18ca.png)

### Head and Tail Options

The 'style designer' was updated in Version 5.0 to help provide more assistance in defining the head and tail attributes for an edge.

![](../media/192ac17604bd3f022781b2a3f0cc9d30.png)

**Head+Tail Label Font Name, Font Size, and Font Color** - These attributes provide a way to differentiate the text at the end of the edges where they meet the node.

![](../media/c1a002afde1e098c2d4bba08205953d2.png) Appears as: ![](../media/d3146f7f10011c4c1738cdf5094d75f6.png)

labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"

**Label Angle** - labelangle= (along with labeldistance=), determines where the head label and tail label are placed with respect to the head (tail) in polar coordinates. The origin in the coordinate system is the point where the edge touches the node. The ray of 0 degrees goes from the origin back along the edge, parallel to the edge at the origin.

The angle, in degrees, specifies the rotation from the 0 degree ray, with positive angles moving counterclockwise and negative angles moving clockwise.

![](../media/ca5ea6312ec3189666321a1ff2d628fa.png) Appears as: ![](../media/e144fb147a088b5a29a80cf391fc6488.png)

labelangle="90"

labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"

**Label Distance** - Multiplicative scaling factor adjusting the distance that the headlabel (taillabel) is from the head (tail) node. The default distance is 10 points. See labelangle for more details.

![](../media/db3e619448238b58510db873ac3c9fd9.png) Appears as: ![](../media/4096c8a32f9f7f755dd5db3d68f7513d.png)

`labeldistance="3"`

`labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"`

**Label Angle & Label Distance** - This example depicts when lableangle= and labeldistance= attributes are used together.

![](../media/876f4b9d1dbfa27bfbfab33744b07bfb.png) Appears as: ![](../media/726ea98317e2b666d5c3143d436e03a3.png)

`labelangle="90" labeldistance="3"`

`labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"`

**Edge Head Port** - Indicates where on the head node to attach the head of the edge. In the default case, the edge is aimed towards the center of the node, and then clipped at the node boundary.

If a compass point is used, it must have the form "n", "ne", "e", "se", "s", "sw", "w", "nw", "c", "_". This specification modifies the edge placement to aim for the corresponding compass point on the port or, in the second form where no port name is supplied, on the node itself. The compass point "c" specifies the center of the node or port. The compass point "_" specifies that an appropriate side of the port adjacent to the exterior of the node should be used, if such exists. Otherwise, the center is used. If no compass point is used with a port name, the default value is "\_".

![](../media/c95b15302b40b1b8de2dba65904a38ba.png) Appears As: ![](../media/9761139853713845c282cc6de415edaf.png)

`headport="n"`

`labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"`

**Edge Tail Port** - Indicates where on the tail node to attach the tail of the edge.

If a compass point is used, it must have the form "n", "ne", "e", "se", "s", "sw", "w", "nw", "c", "_". This specification modifies the edge placement to aim for the corresponding compass point on the port or, in the second form where no port name is supplied, on the node itself. The compass point "c" specifies the center of the node or port. The compass point "_" specifies that an appropriate side of the port adjacent to the exterior of the node should be used, if such exists. Otherwise, the center is used. If no compass point is used with a port name, the default value is "\_".

![](../media/52950a49b2df611060bed2131269abbe.png) Appears as: ![](../media/9c7559d33c6e09fe77a9644cca7559a8.png)

`headport="n" tailport="s"`

`labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"`

## Graphs

### Rotating Graphs 90 Degrees

Graphs can have their drawing orientation set to landscape by setting the rotate attribute equal to 90. The final output is rotated in the counterclockwise direction. The data below:

![](../media/721fd91d780c23a6aef1ea1024903d4d.png)

Creates the following graph:

![](../media/a49de1f0f625a72a9c90299f8119b4b0.png)

An alternate method is to check the "Rotate 90 counterclockwise" option from the Graph Options on the Graphviz tab.

![](../media/761f2489be7ed0141fdd656808167276.png)
