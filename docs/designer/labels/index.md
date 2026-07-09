---
title: Labels - Style Designer
description: Guides you through defining label fonts, colors, sizes, emphasis, and alignment to control text appearance.
---

# Labels

You can design styles which format label text using the following controls:

- Color  
- Font  
- Font size  
- Bold  
- Italic  
- Label Location

| ![Screenshot of the label appearance controls in the Style Designer, showing options for font, size, color, bold, italic, and label placement.](./label_appearance.png) |
| -- |

## Label Fonts

The **Style Designer** font drop‑downs present a gallery of fonts that Graphviz can render on your chosen operating system:

- **Windows** → The list is derived from the fonts installed on your PC, filtered to remove fonts known to be incompatible with Graphviz.  
- **macOS** → A static list of fonts is provided from the **lists** worksheet.

When the **Font Name** drop‑down is selected for the first time, Graphviz generates a preview image of the letters **Aa Bb Cc** for each font.  

These preview images are cached for future use. You may notice a slight delay the first time as the cache is built, but subsequent displays occur quickly.

An example **Font Name** gallery on Windows 11 appears as follows:

![Screenshot of the Windows 11 Font Name gallery showing preview tiles for each font, with “Aa Bb Cc” rendered in the corresponding typeface.](./font_gallery.png)

## Selecting a Font

The currently selected font name is highlighted in the gallery.  

When you choose a font (e.g., `Comic Sans MS`):

- The **Font Name** caption on the drop‑down changes to the selected font.  
- An icon of the letter **A** in the font appears in the ribbon to the left of the **Font Name**.  
- The font name is added as an attribute in the **Format String**.  
- A new preview image is generated, showing the associated labels rendered in the chosen font.

For example:

![Screenshot of the Style Designer showing Comic Sans MS selected as the font, with the preview panel updated to display labels in that typeface.](./font_comic_sans.png)

## Label Location

Text can be aligned relative to the borders of a shape or cluster. Alignment is available as follows via the alignment buttons:

| ![Screenshot of label alignment controls showing options for top, middle, bottom, left, center, and right alignment.](./text_alignment.png) |
| --- |

| Position| Node |  Cluster |
| --- | :--: |  :---: |
| Top     |  ✅ |   ✅     |
| Center  |  ✅ |   ✅     |
| Bottom  |  ✅ |   ✅     |
| Left    |     |   ✅     |
| Middle  |     |   ✅     |
| Right   |     |  ✅      |
