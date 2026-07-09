---
title: Shapes - Style Designer
description: Defines shape selection and polygon controls for skew, distortion, sides, regular mode, and rotation.
---

# Shapes

Graphviz provides a wide variety of **node shapes** that you can apply in the **Style Designer**.  
Shapes define the overall outline of a node and help visually distinguish different types of elements in your diagram.

Shapes can be used to convey meaning, organize information, or simply improve the readability of your graph. 

For example, rectangles may represent processes, ellipses may represent entities, and diamonds may represent decisions.

## Specifying a shape

Click on the `Shape` drop‑down button. 

| ![Screenshot of the Shape drop‑down button used to open the gallery of Graphviz-supported node shapes.](./shape_button.png) |
| --- |

A gallery of shapes supported by Graphviz is presented showing a sample image of the shape. 

![Screenshot of the Shape gallery displaying all Graphviz-supported node shapes, each shown with a rendered preview.](./shape_gallery.png)

Here we pick one of the rectangle shapes. When you select a shape:

- The name of the chosen shape is displayed in the **Style Designer** ribbon as the caption of the `Shape` button.  
- The shape name is added as an attribute in the **Format String** (e.g., `shape=rect`).  
- A preview image is generated to show how the node will appear when rendered by Graphviz.

![Screenshot of the Style Designer showing the rectangle shape selected, with the preview panel updated to display a rectangular node.](./shape_rect.png)

## Polygon Shapes

Polygon shapes are unique from other shapes in Graphviz and have extra attributes which control how the polygon is created.

If you select `polygon` as the shape the ribbon will change dynamically to present additional choices as shown below:

| ![Screenshot of the Style Designer showing the polygon shape selected, prompting additional polygon‑specific options.](./polygon_choose.png) |
| --- |

Selecting `polygon` changes the ribbon to appear as:

| ![Screenshot of the polygon options panel, showing controls for sides, skew, distortion, rotation, and peripheries.](./polygon_options.png) |
| --- |

## Polygon Skew

Positive values skew top of polygon to right; negative values skew the top of the polygon to the left.

### Positive Skew

| ![Screenshot of a polygon node rendered with positive skew, showing the top edge slanted to the right.](./polygon_skew_positive.png) |
| --- | 

![Graphviz-rendered polygon with skew=1, showing a right‑leaning top edge.](./dc4cfbe9033894c44f26b7011d5ccdc4.png)

`shape="polygon" skew="1"`

### Negative Skew

| ![Screenshot of a polygon node rendered with negative skew, showing the top edge slanted to the left.](./polygon_skew_negative.png) |
| --- | 

![Graphviz-rendered polygon with skew=-1, showing a left‑leaning top edge.](./1bc5be00cf87ddb3b4899d29550a465e.png)

`shape="polygon" skew="-1"`

## Polygon Distortion

Positive values cause top part of the polygon to be larger than bottom; negative values do the opposite.

### Positive Distortion

| ![Screenshot of a polygon node rendered with positive distortion, showing a wider top and narrower bottom.](./polygon_distortion_positive.png) |
| --- | 

![Graphviz-rendered polygon with distortion=1 and regular=No, producing a top‑heavy shape.](./d3a16d0b5a88e15e39c7af65c7b96df1.png)

`shape="polygon" distortion="1" regular="No"`

### Negative Distortion

| ![Screenshot of a polygon node rendered with negative distortion, showing a narrower top and wider bottom.](./polygon_distortion_negative.png) |
| --- | 

![Graphviz-rendered polygon with distortion=-1 and regular=No, producing a bottom‑heavy shape.](./6b48aeeda4f98ec38d07b5c4f4ff5f15.png)

`shape="polygon" distortion="-1" regular="No"`


## Combining Skew with Distortion

| + | skew="-1" | skew="0" | skew="1" |
| :---: | :--: | :--: | :--: |
| **distortion="1"** | ![Graphviz-rendered polygon with distortion=1 and skew=-1, producing a top‑heavy shape leaning left.](./131fde8d0c21cbde937f364e790d1251.png) | ![Graphviz-rendered polygon with distortion=1 and skew=0, producing a symmetrical top‑heavy shape.](./cf4e5073a7590e1b3aaa95805918337d.png) | ![Graphviz-rendered polygon with distortion=1 and skew=1, producing a top‑heavy shape leaning right.](./3e48a62a8c4e79fcd6db37bef589d1bf.png) |
| | | |
| **distortion="0"** | ![Graphviz-rendered polygon with distortion=0 and skew=-1, showing a neutral-height shape leaning left.](./e89d2bd25615a593144342db8bc4cd95.png) | ![Graphviz-rendered polygon with distortion=0 and skew=0, showing a neutral, symmetrical polygon.](./085340ed0d232965f7cf3bfd96545943.png) | ![Graphviz-rendered polygon with distortion=0 and skew=1, showing a neutral-height shape leaning right.](./b902f48209f1013c10632a421a6028d3.png) |
| | | |
| **distortion="-1"** | ![Graphviz-rendered polygon with distortion=-1 and skew=-1, producing a bottom‑heavy shape leaning left.](./1af1b897c26ed637e15a9837381b48e3.png) | ![Graphviz-rendered polygon with distortion=-1 and skew=0, producing a symmetrical bottom‑heavy shape.](./8b96cb5b58691ead550f6659056cb5e9.png) | ![Graphviz-rendered polygon with distortion=-1 and skew=1, producing a bottom‑heavy shape leaning right.](./1c8348e5974113e435f37510ec0553f8.png) |

## Regular Polygon

If true, forces the polygon to be regular, i.e., the vertices of the polygon will lie on a circle whose center is the center of the node.

| ![Screenshot of a regular polygon node, showing evenly spaced vertices positioned on a circular boundary.](./polygon_regular.png) |
| --- |

`shape="polygon" regular="Yes"`

## Polygon Sides

The **sides** attribute controls the number of polygon sides used when drawing a node shape.

- **Default**: A polygon has **4 sides** (a square).  
- **sides < 4**:  
  - If the polygon is **not regular**, Graphviz substitutes an **ellipse**.  
  - If the polygon is **regular**, Graphviz substitutes a **circle**.  
- **sides ≥ 4**:  
  - The node is drawn as a polygon with the specified number of sides.  
  - For example, `sides=5` produces a pentagon, `sides=8` a hexagon, and so on.

When you set **sides**, the chosen value is displayed in the **Style Designer** ribbon, added to the **Format String** (e.g., `sides=6`), and shown in the preview image.

![Screenshot of the polygon sides drop‑down list showing selectable values for the number of polygon sides.](./polygon_sides_choices.png)

### sides=8

| ![Screenshot of the Style Designer showing an 8‑sided polygon selected, with the preview panel displaying an octagonal node.](./polygon_sides_8.png) | 
| --- | 

![Graphviz-rendered polygon with sides=8 and regular=yes, producing a symmetric octagon.](./3f571210fe7626b7f0ab1375cf89e992.png) 

`shape="polygon" sides="8" regular="yes"`

Ellipses/circles can also be skewed and distorted to create unique shapes.

| ![Screenshot of the Style Designer showing sides=1 selected, which produces an ellipse or circle depending on regular mode.](./polygon_sides_1.png) | 
| --- | 

### sides=1, with skew and distortion

![Screenshot of a highly distorted and skewed ellipse created using sides=1, skew=1, and distortion=-1.](./polygon_sides_1_skew.png) 

`shape=polygon sides=1 skew=1 distortion="-1" regular=no`

## Polygon Rotation

The **orientation** attribute controls the rotation angle of a node shape.  
It determines how the shape is drawn relative to its default position.

- **orientation=0** (default)  
  - The shape is drawn in its standard upright position.  

- **orientation=n**  
  - The shape is rotated by *n* degrees, **clockwise**.  
  - For example, `orientation=45` tilts the shape diagonally, while `orientation=90` rotates it a quarter turn.  

- **interaction with regular polygons**  
  - When used with polygon shapes (via the **sides** attribute), orientation rotates the polygon around its center.  
  - For any number of polygon sides, 0 degrees rotation results in a flat base.  
  - This is useful for aligning triangles, diamonds, or other polygons to match the desired layout.

When you set **orientation**, the chosen value is displayed in the **Style Designer** ribbon, added to the **Format String** (e.g., `orientation=90`), and shown in the preview image.

| 5-sided regular polygon with no rotation | 5-sided regular polygon rotated 36 degrees clockwise |
| :--: | :--: |
| ![Graphviz-rendered 5‑sided regular polygon with orientation=0, showing a flat base and upright alignment.](./polygon_rotation_0.png) | ![Graphviz-rendered 5‑sided regular polygon rotated 36 degrees clockwise, showing the shape tilted diagonally.](./polygon_rotation_36.png) |
| | |
| ![Preview panel showing the unrotated 5‑sided polygon as rendered by the Style Designer.](./polygon_rotation_0_preview.png) | ![Preview panel showing the 5‑sided polygon rotated 36 degrees clockwise as rendered by the Style Designer.](./polygon_rotation_36_preview.png) |

