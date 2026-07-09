---
title: Images - Style Designer
description: Guides you through setting image paths, choosing images, and adjusting their scale and position in node styles.
---

# Images

As you develop more advanced relationship graphs you may want to use images to represent the nodes in combination with, or in place of the node shapes. Graphviz supports an `image=` attribute where you can provide a file name of an image to include in a node.

The Relationship Visualizer by default will look for images in the directory where the spreadsheet is saved. 

## Image Storage and Paths

If you wish to store images in other locations, you must either:

- Make a configuration change on the **Settings** worksheet to specify the location(s).  
  - The image path must be defined before you can use the `image=` attribute in a style definition.  
- Include the path to the image directly, in either **Relative** or **Absolute** form.

## Relative vs. Absolute Paths

- **Relative Path**  
  - Specifies the image location relative to the workbook directory.  
  - Example: `images/logo.png`  
  - ✅ Easier portability. If the workbook and images are kept together in a folder, cloning or moving the folder preserves the links automatically.  
  - ✅ Ideal for sharing with others or using across multiple devices.  

- **Absolute Path**  
  - Specifies the full location of the image on your system.  
  - Example: `C:/Users/Jeffrey/Documents/Graphviz/images/logo.png`  
  - ✅ Ensures the image is always found, regardless of where the workbook is located.  
  - ✅ Useful when images are stored in a central repository or shared network drive.  
  - ⚠️ Less portable. Moving the workbook without the same directory structure will break the link.

By choosing the appropriate path type, you can balance **portability** (relative paths) with **certainty of location** (absolute paths).

## Add an image path

Switch to the `settings` worksheet and locate the "Image Path:" setting in the **Graph Options** section. To the right of the cell is a button with three dots […]. If you press that button it will bring up the standard directory selection dialog which you can use to choose the directory where the images are stored. Navigate to the directory and press the "OK" button to transfer the path to the cell.

![Screenshot of the settings worksheet showing the Image Path field and the browse button used to select an image directory.](./6c9b1c72a6c130a6ee8b4410456ac9b9.png)

Your settings should appear like this:

![Screenshot of the settings worksheet after an image directory has been selected, with the Image Path field populated.](./5917de49831274d8adc04972405be847.png)

## Specify an image

Image name is an option on the `style designer` worksheet that is useful when you want to create a common style definition where all nodes of a given style use a common icon. For example, it is possible to depict computers with one image, depict databases with another image, and depict computer programmers with yet another image.

**Step 1** – Define a shape. For this example a rectangle will be used.

![Screenshot of the Style Designer showing a rectangular node before an image is applied.](./image_start.png)

**Step 2** – Look to the far right side of the Ribbon to find the image controls.

| ![Screenshot of the Style Designer ribbon showing the image controls section, including the Choose Image button.](./image_controls.png) |
| :--: |

Press the `Choose Image` button.

Navigate to the directory containing the images and choose an image. A small image is selected in order to demonstrate scaling and placement.

![Screenshot of the file selection dialog showing a list of available images to choose from.](./image_select_a_file.png)

The image by default is placed in the center of the node. For example:

| ![Graphviz preview showing the selected image centered inside a rectangular node.](./image_center.png) |
| :--: |

With the image selected, the Ribbon adapts to display additional options which can be used to scale the image, or position the image within the shape.

| ![Screenshot of the Style Designer ribbon showing additional image scaling and positioning controls after an image is selected.](./image_scal_and_position.png) |
| :--: |

## Scale the Image

Adjust the image by clicking the radio buttons in the **Scale** group. Only one button can be selected. If you make a second selection, your first selection is replaced.

| Scale   | Radio Button | Preview | Description |
| :---:   | :---: | :---: | :--- |
| **height**  | ![Height scaling radio button selected.](./image_scale_height.png) | ![Preview showing the image stretched to fill the node height while width remains unchanged.](./image_scale_height_preview.png) | Stretch image to fill node height; width remains unchanged. |
| | | | |
| **width**   | ![Width scaling radio button selected.](./image_scale_width.png) | ![Preview showing the image stretched to fill the node width while height remains unchanged.](./image_scale_width_preview.png) | Stretch image to fill node width; height remains unchanged. |
| | | | |
| **aspect**  | ![Aspect scaling radio button selected.](./image_scale_aspect.png) | ![Preview showing the image uniformly scaled to fit the node while preserving aspect ratio.](./image_scale_aspect_preview.png) | Uniformly scale image to fit node while preserving aspect ratio. |
| | | | |
| **both**    | ![Both‑dimensions scaling radio button selected.](./image_scale_both.png) | ![Preview showing the image stretched to fill both width and height of the node, potentially distorting aspect ratio.](./image_scale_both_preview.png) | Stretch image to fill both width and height of node; aspect ratio may distort. |
| | | | |
| **natural** | ![Natural scaling radio button selected.](./image_scale_natural.png) | ![Preview showing the image displayed at its natural size, with the node expanding to fit.](./image_scale_natural_preview.png) | Use image’s natural size; node expands to fit (default). |

No scaling (i.e., **natural** scaling) will be used in order to demonstrate how to position images which are smaller than the node.

## Adjust the Image Position

The **Position** group contains nine toggle buttons that work in **radio button fashion**. Selecting any button automatically unselects the previously chosen option.

These buttons correspond to the nine possible locations where images can be placed relative to the cell or shape via the **imagepos** attribute:

| + | Left | Center | Right |
| :--: | :--: | :--: | :--: |
| **Top**    | `tl` | `tc` | `tr` |
| **Middle** | `ml` | `mc` | `mr` |
| **Bottom** | `bl` | `bc` | `br` |

By default, when the **imagepos** attribute is omitted, the image is placed in the middle center.

You can reposition the image within the node by selecting a Position radio button, as shown below.

| Position | Buttons pressed | Preview Image |
| :--: | :--: | :--: |
| Default | ![Position control showing the middle‑center button selected.](./image_position_default.png) | ![Preview showing the image centered within the node.](./image_position_default_preview.png) |
| | | |
| Top Left | ![Position control showing the top‑left button selected.](./image_position_tl.png) | ![Preview showing the image placed at the top‑left corner of the node.](./image_position_tl_preview.png) |
| | | |
| Bottom Center | ![Position control showing the bottom‑center button selected.](./image_position_bc.png) | ![Preview showing the image placed at the bottom‑center of the node.](./image_position_bc_preview.png) |
