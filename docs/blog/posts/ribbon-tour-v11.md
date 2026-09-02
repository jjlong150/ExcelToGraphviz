---
blogPost: true
title: A Tour of the Redesigned Ribbon
description: Version 11.0 splits the ribbon into a new Settings tab, a new Data tab, and a trimmed-down Graphviz tab. Here's a guided tour of what moved where.
date: 2026-09-01
author: jjlong150
tags: ['ribbon', 'how-to', 'v11.0']
published: true
readingTime: true
sidebar: false
---

Adding Knowledge Graph publishing meant the old Graphviz tab simply ran out of room. Rather than cram one more group onto an already-crowded tab, I used the opportunity to reorganize the ribbon around how people actually use it: one tab for visibility settings, one tab for generating and publishing output, and a Graphviz tab scoped down to genuine Graphviz-only options. If you upgrade and can't immediately find a button you used to know by heart, this is your map.

## New: Settings tab

A brand-new Settings tab gives you one central place to show or hide individual worksheets and ribbon tabs, with dedicated groups for Command/Graph Options, Data, Console, Exchange, Extensions, Help URLs, Launchpad, Source, SQL (Windows only), Styles, and SVG.

I redesigned the Settings worksheet to match: the old faux tabbed-folder styling is gone in favor of a simple black-and-white scheme, and each group of rows shows or hides based on which button you press on the ribbon.

<!-- SCREENSHOT: the new Settings tab on the ribbon, with its grouped show/hide buttons -->

## New: Data tab

The old, single Graphviz tab is now two tabs. The new Data tab owns everything related to generating and publishing output:

- **Visualize** — Refresh Graph is now a split button with Automatic Refresh in its dropdown, and the new Knowledge Graph button lives here.
- **Publish** — a split button with an "Open after publishing" option, plus checkboxes to choose exactly which output files get created: Graph, DOT, and Knowledge Graph.
- **File Output** — a new render-engine group lets you pick Cairo, GD, GDI+ (Windows), or Quartz (Mac).
- **Styling** — style-related switches collected in one place, including Apply Styles, Apply Attributes, Add Image Path, Transparent Background, and Rotate 90° CCW.
- **Options** — the Node, Edge, Graph, and new Cluster menus, with new tooltip-inclusion controls.
- **Data Worksheet** — the Show Columns menu is now organized into labeled sections, with a new Show Properties toggle.

<!-- SCREENSHOT: the new Data tab on the ribbon, showing the Publish split button and its checkbox dropdown -->

## Graphviz tab, trimmed down

With generation and publishing moved out, the Graphviz tab is now scoped to genuine Graphviz-only options — layout engine, splines, direction, and output order. Splines controls (Compound, Line, Spline) are back as live, selectable buttons after being disabled for a while due to space constraints, and a new Depth option joins the existing Breadth option for output order. If you're focused purely on Knowledge Graphs, a new toggle on the Launchpad tab lets you hide the Graphviz tab entirely.

## Styles tab: configurable cluster naming

Cluster style names — written onto a cluster's opening and closing brace rows — can now be built from a configurable Naming Pattern instead of a fixed suffix. The old Suffix (Begin)/Suffix (End) fields are renamed Affix fields, and a new Naming Pattern field controls where that value is inserted, using `{name}` and `{affix}` placeholders. The default is `{name} {affix}`, but you can just as easily make the affix a prefix instead.

<!-- SCREENSHOT: the Styles tab showing the Affix fields and Naming Pattern field -->

That's the tour. See the [full changelog](/changelog/) for the complete list of ribbon changes, including the SVG tab's new on/off controls.
