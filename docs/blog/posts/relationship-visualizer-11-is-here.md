---
blogPost: true
title: Relationship Visualizer 11.0 Is Here
description: Relationship Visualizer 11.0 adds Knowledge Graph export, a big-diagram SVG viewer, and a reorganized ribbon. Here's what's new, at a glance.
date: 2026-09-02
author: jjlong150
tags: ['release', 'v11.0', 'announcement']
published: true
readingTime: true
sidebar: false
---

Today I'm shipping the biggest release Relationship Visualizer has ever had. Rather than one headline feature, this one has two: your worksheet data can now be exported as a Knowledge Graph instead of just a diagram, and the diagrams you already generate get a dramatically more capable SVG viewer, with pan, zoom, filtering, and highlighting built right in.

Underneath those two, the ribbon got reorganized from top to bottom, SQL and Workbook Exchange both picked up new tricks, and I cleaned up a handful of rough edges that had been bugging me for a while. Here's the tour, with links if you want to go deeper on any one piece.

## Knowledge Graphs, straight from your spreadsheet

Alongside the diagrams you already generate, you can now export your worksheet data as a Knowledge Graph — a structured JSON document meant to be pasted straight into an AI tool rather than rendered as a picture. No database, no separate modeling step, no code. It runs from the same data and the same label, style, and inheritance rules as your diagrams, so the two outputs never disagree with each other.

I wrote up the full story here: **[From Spreadsheet to Knowledge Graph](/blog/posts/knowledge-graph-export)** — and if you want to see it actually put to use, **[I Let an AI Read My Knowledge Graph](/blog/posts/ai-reads-my-knowledge-graph)**.

## A diagram viewer built for big graphs

If you already read my last post, you know this one: a fixed toolbar, hover labels, and click-to-highlight filtering, built directly into every exported SVG. No separate viewer or plugin required — open the file and it's already there.

Full writeup: **[Exploring diagrams just got easier](/blog/posts/interactive-diagrams)**

## A reorganized ribbon

A new Settings tab centralizes worksheet and tab visibility. A new Data tab takes over generation and publishing, freeing the Graphviz tab to focus purely on layout. Cluster naming now uses a configurable pattern instead of a fixed suffix.

Full writeup: **[A Tour of the Redesigned Ribbon](/blog/posts/ribbon-tour-v11)**

## Upgrading from 10.5

There's a short list of breaking changes — all low-impact — worth a quick read before you upgrade, especially if you rely on automatic `shape=plaintext` for HTML-like labels.

Full writeup: **[Upgrading to v11.0: What to Check](/blog/posts/upgrading-to-v11)**

A couple of smaller notes worth a quick read if you're curious: **[Fresher Style Previews After Upgrading](/blog/posts/fresher-style-previews)**.

---

Grab the update and check the [full changelog](/changelog/) for every detail, big and small. As always, I'd love to hear what you think.
