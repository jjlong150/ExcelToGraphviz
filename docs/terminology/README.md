---
prev: /install/
next: /workbook/
---

# Visualization Terminology

In mathematics and computer science, _graph theory_ is the study of graphs, which are mathematical structures used to model pairwise relations between objects.

The terms described in this chapter are used throughout the rest of this manual to explain how to construct visualizations. These terms have their roots in graph theory and/or the Graphviz tool.

## Graph

The following picture illustrates a "graph".

![](../media/67ab8d51d2a46bbbeb2f1b1f7d4c1fe1.png)

## Node

A "graph" in this context is comprised of "nodes".

![](../media/99d7592123e9a29a0bb3e9e7fe243da9.png)

## Edge

"Edges" are lines that connect nodes.

![](../media/f2ed7375b0b033d3cfb46f617f61e233.png)

## Undirected Graph

A graph may be "undirected", meaning that there is no distinction between the two nodes associated with each edge.

![](../media/728e39092179648bdc2e546a72e0624f.png)

## Directed Graph

A graph may be "directed" meaning that there is an explicit direction from one node to another.

![](../media/f4061834a61e3356bc65cc2c4ba4ab76.png)

## Labels

Nodes can have "labels". Labels can be placed inside the node, and outside the node.

![](../media/1c7ca23f1a71903742522c8f630d0d13.png)

Edges can also have labels. Edge labels can be placed on the edge,

![](../media/a60e967a1b22dc436134cb2eefe91753.png)

at the tail and/or head of the edge,

![](../media/7590a1ea8a2fe92f8eace465df0b924b.png)

Or outside the edge (however in my experience they tend to not always render well)

![](../media/ff1e8eb1c6f7686b89a0cdf50e9748cb.png)

Edge labels are helpful in stating what the relationship between the nodes is. For example, a set of family relationships might look as follows:

![](../media/19e2acd3c9d17a93396f47f4b99ce3f4.png)

## Splines

The way in which edges are routed and drawn are called "splines". Several spline types are available in Graphviz. The spline type and a depiction of each follows:

### curved

Edges are drawn as curved arcs between nodes

![](../media/c0087e3a8e9bd632fd48d7a3b254fffe.png)

### line

Edges are drawn as straight lines between nodes

![](../media/d1e3da5e238bf384bf7ad2882e8f9de1.png)

### none

dges (and edge labels) are not drawn between nodes, but the relationships described by the edges affects the placement of the nodes.

![](../media/91fb7409d38ccb764663dfddafdc82e7.png)

### ortho

Edges are drawn with 90-degree angles in the routes between nodes.

![](../media/7c25a63fb116f430ed6f93af46bf26c1.png)

### polyline

Edges are drawn with straight lines and angular bends in the routes between nodes.

![](../media/368fd5b604a05ffabffb4f940ce37359.png)

### spline

Edges are drawn with straight and free-flowing (curvy) lines in the routes between nodes.

![](../media/9209c6b5cf0b939dee639fa1b9d8239c.png)

## ports

A port name can be combined with the node name to indicate where to attach an edge to the node. Graphviz has built-in port names N, S, E, W, NE, NW, SE, SW, C corresponding to compass points North, South, East, West, North East, North West, South East, South West and Center respectively.

![](../media/567dbee9f9f4ae3519b9d2c8b3f6dce6.png)

Custom ports can also be specified when using HTML labels or "record" as the node shape. This feature is explained later in this manual.

## Clusters / Subgraphs

"Clusters" is a feature to draw nodes and edges in a separate rectangular layout region. Clusters exist as subgraphs of a parent graph internal to Graphviz.

Only the "dot", "fdp", "neato" and "osage" layout engines (described in the next section) draw clusters.

In the example that follows, the rectangles labeled "process \#1" and "process \#2" are clusters (subgraphs) within the overall graph.

![](../media/af1eef86eebcb84379d993d114dca859.png)

## Layout Algorithms

Graphviz contains several programs for drawing graphs. Each program has specializations in how they determine how to layout the nodes and edges. Choosing a layout algorithm to use is sometimes a trial-and-error exercise to find which output looks the best.

A description of the layout engines available (as documented on the Graphviz homepage) are as follows:

### circo

circular layout, after Six and Tollis 99, Kauffman and Wiese 02. This is suitable for certain diagrams of multiple cyclic structures, such as certain telecommunications networks.

![](../media/04743315a97480eaa93d2227aa539137.png)

### dot

"hierarchical" or layered drawings of directed graphs. This is the default tool to use if you want to have some control regarding the direction of how the graph is drawn.

![](../media/97e7d47c51aa9e614545cadc841c0264.png)

### fdp

"spring model'' layouts like those of neato but does this by reducing forces rather than working with energy.

![](../media/f20ce45f4d0658b010af1502e6d1dbfc.png)

### neato

"spring model'' layouts. This is the default tool to use if the graph is not too large (about 100 nodes) and you do not know anything else about it. Neato attempts to minimize a global energy function, which is equivalent to statistical multi-dimensional scaling.

![](../media/c5949574ac970952e42a97f0b2864277.png)

### osage

The _osage_ layout algorithm is for large undirected graphed with multiple subgraphs. It separates the graph into "levels" (clusters) and lays out each level in a rectangle. The rectangles are then packed together. Within each rectangle, the subgraph/cluster is laid out.

![](../media/66804945ad6e5a0df74965c2cac96fdc.png)

### patchwork

The patchwork layout engine draws the graph as a squarified treemap. The clusters on the graph are used to create the tree.

![](../media/5a06a33a66c458bf65bd8430604f1779.png)

### sfdp

Multiscale version of _fdp_ for the layout of large graphs.

![](../media/546176879a8673a33c1cc29038dc53ec.png)

### twopi

Radial layouts, after Graham Wills 97. Nodes are placed on concentric circles depending their distance from a given root node.

![](../media/8dff9fd7bf86eadcdbb3c4c4854085c6.png)
