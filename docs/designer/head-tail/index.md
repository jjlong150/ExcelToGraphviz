---
   title: Head & Tail - Style Designer
   description: Guides you through configuring edge head and tail styling, label placement, port options, and clipping controls.
---

# Head & Tail Options

These controls provide assistance in defining the head and tail attributes for an edge.

| ![Screenshot of the Style Designer showing head and tail configuration controls for edge labels and endpoint styling.](./192ac17604bd3f022781b2a3f0cc9d30.png) |
|-|

## Label Font Color - Label Font Name - Label Font Size

These attributes provide a way to differentiate the text at the end of the edges where they meet the node.

| ![Screenshot of the Style Designer ribbon showing controls for labelfontname, labelfontsize, and labelfontcolor.](./c1a002afde1e098c2d4bba08205953d2.png) |
|-|

Appears as:

![Preview showing an edge with styled head and tail labels using Arial font, size 8, and blue text.](./d3146f7f10011c4c1738cdf5094d75f6.png)

With Format string:

`labelfontname="Arial" labelfontsize="8" labelfontcolor="Blue"`

## Label Angle

`labelangle=` controls the **direction** in which a head label or tail label appears around the point where the edge touches the node.

Imagine standing at the spot where the edge meets the node. Now imagine a line pointing straight back along the edge. That's the starting direction (0 degrees).  
`labelangle` tells Graphviz how far to rotate from that starting direction:

- **Positive numbers** rotate the label **to the left**
- **Negative numbers** rotate the label **to the right**

By changing the angle, you choose which “side” of the node the label appears on.

For example, setting the label angle to 90 degrees:

| ![Screenshot of the Style Designer showing labelangle set to 90 degrees with corresponding font controls.](./ca5ea6312ec3189666321a1ff2d628fa.png) |
|-|

Appears as:

![Preview showing a head or tail label positioned 90 degrees from the edge’s attachment point, styled in Arial 8pt blue text.](./e144fb147a088b5a29a80cf391fc6488.png)

With Format String:

`labelangle=90 labelfontname=Arial labelfontsize=8 labelfontcolor=Blue`

## Label Distance

`labeldistance=` controls **how far away** the head label or tail label appears from the point where the edge touches the node.

Instead of setting the distance directly in points, `labeldistance` acts as a **scaling factor**.  
Graphviz starts from a built‑in base distance (about **10 points**), and your value multiplies that distance:

- A value of **1.0** keeps the default spacing  
- A value of **2.0** places the label about twice as far away  
- A value of **0.5** moves it to about half the default distance  

A **point** is a standard typographic unit: there are **72 points in one inch**, so these changes adjust the label’s distance in small, predictable steps.

In short, `labeldistance` tells Graphviz to move the label **closer or farther** from the node by scaling the default distance.

| ![Screenshot of the Style Designer showing the labeldistance control set to a custom value, along with label font options.](./db3e619448238b58510db873ac3c9fd9.png) |
|-|

For example, `labeldistance=3` appears as:

![Preview showing a head or tail label positioned farther from the node due to labeldistance=3, styled in Arial 8pt blue text.](./4096c8a32f9f7f755dd5db3d68f7513d.png)

With Format String:

`labelangle=90 labeldistance=3 labelfontname=Arial labelfontsize=8 labelfontcolor=Blue`

## Label Angle & Label Distance

Used together, `labelangle` and `labeldistance` let you control both **where** a label appears around the node and **how far out** it sits. `labelangle` chooses the direction (left, right, above, below, or anywhere in between) while `labeldistance` scales the default spacing to move the label closer or farther away. Adjusting both gives you precise, intuitive control over label placement at the point where the edge meets the node.

This example depicts when `labelangle=` and `labeldistance=` attributes are used together.

| ![Screenshot of the Style Designer showing labelangle and labeldistance set together, with font options visible.](./876f4b9d1dbfa27bfbfab33744b07bfb.png) |
|-|

Appears as:

![Preview showing a head or tail label positioned using both labelangle=90 and labeldistance=3, styled in Arial 8pt blue text.](./726ea98317e2b666d5c3143d436e03a3.png)

With Format String:

`labelangle=90 labeldistance=3 labelfontname=Arial labelfontsize=8 labelfontcolor=Blue`

## Head Port

Indicates where on the head node to attach the head of the edge. In the default case, the edge is aimed towards the center of the node, and then clipped at the node boundary.

If a compass point is used, it must be one of the following: `n`, `ne`, `e`, `se`, `s`, `sw`, `w`, `nw`, `c`, or `_`. A compass point adjusts the edge’s attachment point so that it aims for the specified location on the port, or, if no port name is provided, on the node itself. The compass point `c` targets the center of the node or port. The compass point `_` instructs Graphviz to choose the side of the port that lies on the exterior of the node; if no such side exists, the center is used instead. When a port name is supplied without a compass point, the default value is `_`.

![Screenshot of the Style Designer showing the headport control set to a compass point value.](./c95b15302b40b1b8de2dba65904a38ba.png)

Appears As:

![Preview showing an edge attached to the north side of the head node using headport=n, with styled label text.](./9761139853713845c282cc6de415edaf.png)

With Format String:

`labelfontname=Arial labelfontsize=8 labelfontcolor=Blue headport=n`

## Tail Port

Indicates where on the tail node to attach the tail of the edge.

If a compass point is used, it must be one of the following: `n`, `ne`, `e`, `se`, `s`, `sw`, `w`, `nw`, `c`, or `_`. A compass point modifies edge placement so that the edge aims for the specified point on the port, or, if no port name is supplied, on the node itself. The compass point `c` targets the center of the node or port. The compass point `_` indicates that Graphviz should choose the side of the port that lies on the exterior of the node; if no such side exists, the center is used instead. When a port name is provided without a compass point, the default compass point is `_`.

![Screenshot of the Style Designer showing the tailport gallery control.](./52950a49b2df611060bed2131269abbe.png)

Appears as:

![Preview showing an edge attached to the south side of the tail node using tailport=s, with styled label text.](./9c7559d33c6e09fe77a9644cca7559a8.png)

With Format String:

`labelfontname=Arial labelfontsize=8 labelfontcolor=Blue headport=n tailport=s`

## Clipping Behavior

Graphviz uses *clipping* to decide how far an edge (its spline and arrowhead) runs into a node. The attributes `headclip` and `tailclip` control this behavior independently for the head and tail ends of an edge. These settings affect the **edge and arrowhead**, not the label text itself.

## Head Clip

`headclip=` controls how the edge is clipped at the **head** node.

- When `headclip=true` (the default), Graphviz clips the spline at the boundary of the head node. The arrowhead sits at the edge of the node shape, rather than running into the center.
- When `headclip=false`, the edge is not clipped to the node boundary. The spline and arrowhead may extend into the interior of the node, often aiming at its center.

| Buttons | Preview |
| :-----: | :-----: |
| ![Ribbon controls showing headclip=true and tailclip=true selected.](./clip-true-true.png) | ![Preview showing the edge head clipped at the node boundary.](./tailclip-true-headclip-true.png) |
|  | Edge head **is** clipped at the node |
| ![Ribbon controls showing headclip=false and tailclip=true selected.](./clip-false-true.png) | ![Preview showing the edge head extending into the node interior.](./tailclip-true-headclip-false.png) |
|  | Edge head **is not** clipped at the node |

## Tail Clip

`tailclip=` controls how the edge is clipped at the **tail** node.

- When `tailclip=true` (the default), Graphviz clips the spline at the boundary of the tail node. The edge meets the node at its outline.
- When `tailclip=false`, the spline is allowed to extend into the node, so the edge may appear to start from a point inside the node.

| Buttons | Preview |
| :-----: | :-----: |
| ![Ribbon controls showing tailclip=true and headclip=true selected.](./clip-true-true.png) | ![Preview showing the edge tail clipped at the node boundary.](./tailclip-true-headclip-true.png) |
|  | Edge tail **is** clipped at the node |
| ![Ribbon controls showing tailclip=false and headclip=true selected.](./clip-true-false.png) | ![Preview showing the edge tail extending into the node interior.](./tailclip-false-headclip-true.png) |
|  | Edge tail **is not** clipped at the node |
