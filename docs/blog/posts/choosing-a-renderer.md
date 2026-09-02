---
blogPost: true
title: Choosing a Graphviz Renderer
description: Cairo, GD, GDI+, and Quartz all draw your diagram — here's how the four render engines in v11.0 actually differ, and when to pick each one.
date: 2026-09-05  # PLACEHOLDER — set to actual publish date before going live
author: jjlong150
tags: ['rendering', 'tips', 'v11.0']
published: false  # flip to true when ready to launch
readingTime: true
sidebar: false
---

Version 11.0 added a render-engine picker to the Data tab: Cairo, GD, GDI+ (Windows only), and Quartz (Mac only). If you've never had a reason to touch it, here's the short version of when you would.

**Cairo** is Graphviz's modern default and the right choice for almost everyone. It handles anti-aliasing, transparency, and font rendering the most consistently across platforms, and it's what I'd recommend leaving selected unless you have a specific reason to change it.

**GD** is the older, simpler renderer. It's faster on very large diagrams and has fewer font-rendering dependencies, which makes it a reasonable fallback if Cairo is producing garbled text or missing glyphs in your environment — something that occasionally happens with unusual font configurations.

**GDI+** (Windows only) draws through Windows' own graphics stack, which can give you output that matches other Windows-native tools more closely, and it's worth trying if you're embedding diagrams into other Windows applications and want visual consistency.

**Quartz** (Mac only) is the macOS equivalent — Apple's native rendering, for the same reason you'd reach for GDI+ on Windows.

<!-- SCREENSHOT: the File Output group on the Data tab with the four renderer checkboxes visible -->

If you've never changed this setting, you don't need to start now — Cairo is a good default. But if you ever run into odd font spacing, missing characters, or rendering artifacts on a particular machine, this is the first thing I'd try switching.

See the [full changelog](/changelog/) for everything else new in v11.0.
