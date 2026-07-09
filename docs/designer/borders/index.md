---
title: Borders - Style Designer
description: Guides you through defining border styles, colors, pen widths, and peripheries to customize node and cluster outlines.
---

# Borders 

## Border Styles

Up to 3 border styles are selectable and are additive, making it possible to have styles such as bold edge and rounded corners. When you click on any of the **Border Style** drop‑down lists you will be presented with the list of choices along with a sample image of the style.

![Screenshot of the border style gallery showing selectable options such as solid, dashed, rounded, and bold, each with a preview.](./border_styles.png)

In this example `Style 1` is set to `rounded` to give the rectangle rounded corners.

![Screenshot of a rectangular node with rounded corners applied using the rounded border style.](./border_styles_rounded.png)

The **Style Designer** provides an adaptive interface for applying multiple border styles.  
As you make selections, additional style options appear dynamically:

- Once a style is chosen, the **Style 2** drop‑down becomes available.  
- Selecting `dashed` as the second choice results in a rounded, dashed border.  
- A **Style 3** drop‑down then appears, allowing you to continue layering styles.

This adaptive behavior makes it easy to combine multiple visual effects without cluttering the interface.

![Screenshot of a node with both rounded and dashed border styles applied, demonstrating layered border effects.](./border_styles_rounded_dashed.png)

## Border Color

How to choose colors has already been explained.  

The **Border Color** controls allow you to specify the color of a node shape or cluster border.

For example:

![Screenshot of the Style Designer showing the border color picker applied to a rectangular node, with the preview panel reflecting the selected border color.](./border_color.png)

## Border Pen Width

The **penwidth** attribute controls the thickness of lines used to draw node borders and edges.

![Graphviz-rendered comparison of border thicknesses demonstrating different penwidth values.](./border_penwidth_dropdown.png)

- **penwidth=1.0** (default)  
  - Standard line thickness.  
  - Borders and edges are drawn with a single‑pixel width.  

- **penwidth>1.0**  
  - Increases line thickness proportionally.  
  - For example, `penwidth=2.0` doubles the thickness, while `penwidth=3.0` triples it.  
  - Useful for emphasizing certain nodes or edges in a diagram.  

- **penwidth<1.0**  
  - Decreases line thickness.  
  - For example, `penwidth=0.5` produces a thinner line than the default.  
  - Can be used for subtle or secondary connections.

For example:

![Screenshot of the Style Designer showing a node with increased border pen width applied, with the preview panel reflecting the thicker outline.](./border_penwidth.png)

## Border Peripheries

The **peripheries** attribute controls how many borders (or outlines) are drawn around a node shape.

![Screenshot of peripheries dropdown list showing 1 to 5 concentric borders to illustrate the peripheries attribute.](./border_peripheries_dropdown.png)

- **peripheries=1** (default)  
  - A single border is drawn around the shape.  

- **peripheries=2**  
  - Two concentric borders are drawn, giving the node a “double‑outlined” appearance.  

- **peripheries=n**  
  - Any positive integer value `n` draws that many concentric borders.  
  - Useful for visually emphasizing certain nodes or distinguishing categories.  

For example:

![Screenshot of the Style Designer showing a node with 2 peripheries applied, with the preview panel displaying concentric borders.](./border_peripheries.png)

