---
blogPost: true
title: "Upgrading to v11.0: What to Check"
description: A short, practical checklist of v11.0's low-impact breaking changes — what changed, who's affected, and how to adjust if you need to.
date: 2026-09-01
author: jjlong150
tags: ['upgrade-guide', 'v11.0']
published: true
readingTime: true
sidebar: false
---

Version 11.0 is a big release, but the breaking changes in it are small and narrow. I don't expect most of you to notice any of them — this post exists so you can check the list in under a minute and get back to work.

## `shape=plaintext` is no longer automatic

In v10.5 and earlier, any node with an HTML-like label and no other style attributes automatically got `shape=plaintext`. I've intentionally dropped that implicit override in v11.0, in favor of a truer, unopinionated Graphviz experience — Graphviz's own default shape now applies unless a style says otherwise.

**Who's affected:** Anyone with HTML-like node labels who never explicitly set a shape. I expect very few of you to notice.

**What to do:** If a node's appearance changes after upgrading, add `shape=plaintext` explicitly — either as a style attribute or a row-level extra attribute.

## Error reporting moved off the data worksheet

Row-validation errors used to be written into a dedicated "error message" worksheet column, shown and hidden around each generation run. Version 11.0 instead builds a localized message and routes it through the console or a message box, and the worksheet error-message column has been removed entirely.

**Who's affected:** Anyone with automation or muscle memory built around that worksheet column.

**What to do:** Check the Console tab or your message-box preference for validation errors going forward. The `!` indicator in column 1 is still there to flag rows with a detected error.

## Internal cleanup, no visible effect

I tightened ribbon callback functions from public to private scope so they no longer clutter Excel's Assign Macro dialog, and removed leftover linter-suppression comments. Neither change affects how the ribbon or any feature behaves — listed here only for completeness.

## Progress bar code removed

The progress bar shown for long-running operations in older versions was turned off back in v10.3 and left dormant since. I've now removed the unused code.

## One thing that *isn't* breaking: your Exchange files

If you use Workbook Exchange to version-control styles and settings, older export files — including ones using the pre-v11.0 Suffix-based cluster naming — still import correctly. Everything new in v11.0 round-trips automatically going forward.

---

That's the whole list. See the [full changelog](/changelog/) for the complete picture of what's new, and open a GitHub issue if anything here catches you by surprise.
