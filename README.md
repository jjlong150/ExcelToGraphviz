# Excel to Graphviz

**Transform Excel tabular data into professional Graphviz diagrams — directly in a spreadsheet.**

The **Relationship Visualizer** is a powerful, free Excel tool that turns row-based data (org structures, dependencies, networks, processes) into stunning Graphviz visualizations using custom styles, and every major layout engine.

Supports Windows and macOS, with a multilingual tabbed ribbon interface.

[![Latest Release v10.1.0](https://img.shields.io/badge/Latest%20Release-v10.1.0%20(Feb%202026)-brightgreen)](https://github.com/jjlong150/ExcelToGraphviz/releases/latest)
[![Download Now](https://img.shields.io/badge/Download%20Now-Free%20Excel%20Tool-green)](https://sourceforge.net/projects/relationship-visualizer/files/latest/download)
[![MIT License](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)
[![Buy Me a Coffee](https://img.shields.io/badge/Support-Buy%20Me%20a%20Coffee-yellow)](https://www.buymeacoffee.com/exceltographviz)

## Quick Links

- 🌐 **Website & Full Documentation**: [exceltographviz.com](https://exceltographviz.com/) — tutorials, overviews, and guides
- 📥 **Latest Release (v10.1.0 — Feb 2026)**: [GitHub Releases](https://github.com/jjlong150/ExcelToGraphviz/releases/latest) — release notes & assets
- ⬇️ **Download Relationship Visualizer**: [SourceForge ZIP (~81 MB)](https://sourceforge.net/projects/relationship-visualizer/files/latest/download) — free & virus-scanned
- 📜 **Changelog**: [exceltographviz.com/changelog](https://exceltographviz.com/changelog) — version history
- 🔧 **Examples Repository**: [github.com/jjlong150/excel-to-graphviz-examples](https://github.com/jjlong150/excel-to-graphviz-examples) — ready-to-use workbooks & patterns
- ☕ **Support the Project**: [Buy Me a Coffee](https://www.buymeacoffee.com/exceltographviz) — optional thanks for ongoing development!

## Quick Start

Get diagramming in minutes — no coding required!

1. **Download & Open** — Extract **Relationship Visualizer.zip** and open the main spreadsheet (enable macros if prompted).
  
   *Windows tip: If Excel opens in Protected View or macros don't load, right-click the extracted .xlsm file → Properties → General tab → check "Unblock" if present. This is a standard Windows security step for internet-downloaded files. Full details in the [Windows Install Guide](./docs/install-win/README.md).*

2. **Enter Data Manually** (easiest way, works on Windows & macOS):
   - Go to the **data** worksheet.
   - Fill in columns like **Item** (node name), **Related Item** (for edges), **Label**, **Style**, etc. — think of it as listing nodes and connections.
   - Check the **Automatic** box in the ribbon (under Graphviz tab) if it's not already on.
   - As you type or edit cells, the graph renders/refreshes instantly to the right (or in a separate graph worksheet if selected)!

3. **Customize & Render**:
   - Adjust layout engine (dot, neato, etc.), styles, colors via the ribbon.
   - Press **Refresh Graph** (or let Automatic handle it) for updates.
   - View/save the output as SVG/PNG.

4. **Optional: Use SQL for Advanced Data** (Windows only):
   - On the **sql** worksheet, write queries to pull/transform data from other sheets/workbooks.
   - Run them to populate the **data** worksheet automatically.
   - macOS users: Skip this — manual entry works great and covers most needs!

For ready-to-use templates, check the [Examples Repository](https://github.com/jjlong150/excel-to-graphviz-examples) and the `samples` directory in the download zip file.

See platform guides: [Windows Install](./docs/install-win/README.md) | [macOS Install](./docs/install-mac/README.md)

## Example Previews

Here are real diagrams generated directly from Excel data using the Relationship Visualizer:

![Unix Shell Timeline](./docs/media/home_timeline.png)  
*Time-oriented dependency timeline showing the evolution of Unix shells (dot engine)*

![London Underground Map](./docs/media/home_london_underground.png)  
*Styled logical map of the London Underground Metropolitan Line (dot engine)*

![Organizational Chart](./docs/media/home_orgchart.png)  
*Hierarchical org chart built from an Excel table (dot engine)*

![Musician-Band Network](./docs/media/home_neato.png)  
*Force-directed network of musicians and bands with custom styling and embedded images (neato engine)*

All previews rendered live from Excel tables using Graphviz — no external tools needed.

## Key Features

- Generate Graphviz DOT code and diagrams directly from Excel tabular data using an intuitive tabbed Ribbon interface
- Manually enter or edit nodes/edges in the data worksheet — graphs render automatically as you type (cross-platform)
- Create, combine, preview, and save reusable styles as stylesheets for consistent diagram formatting
- View generated DOT source, Graphviz console output, and rendered results as SVG or PNG
- Leverage SQL queries (iteration, concatenation, recursion) to import, transform, and dynamically refresh data — build query pipelines for complex scenarios (Windows only)
- Add animation code to exported SVG files for interactive diagrams
- Multilingual Ribbon interface: English, French, German, Italian, Polish
- Cross-platform support: Full SQL and Clipboard features on Windows; core diagramming and rendering on macOS

## Repository Structure

This GitHub repository contains the source code, documentation, and build artifacts for the **Relationship Visualizer** Excel tool. The ready-to-use workbook is distributed as a ZIP via [SourceForge](https://sourceforge.net/projects/relationship-visualizer/), while here you'll find the extracted VBA source, website content, legacy docs, and more for transparency, contributions, and development.

Here's the main directory overview:

```
.
├── .github/                          # GitHub workflows and FUNDING.yml
├── dist/                             # Distribution-ready ZIP assets
│   ├── Relationship Visualizer/      # Distribution assets
│   │   ├── licenses/                 # Component license files
│   │   └── samples/                  # Sample workbooks
│   └── Relationship Visualizer.zip   # Distribution file published on SourceForge
├── docs/                             # https://exceltographviz.com content files
│   ├── .vuepress/                    # VuePress configuration settings
│   ├── topic(s)/                     # Markdown content structured within subdirectories by topic.
│   └── README.md                     # Home page
├── legacy_docs/                      # Legacy user documentation (.docx, .pdf)
├── src/                              # Source files for the workbook
│   ├── applescript/                  # applescript script for running on macOS
│   ├── excel/                        # Excel workbook matching extracted source
│   ├── vba/                          # VBA code extracted from 'Relationship Visualizer.xlsm'
│   │   ├── Class Modules/            # VBA class files (.cls)
│   │   ├── Forms/                    # VBA form files (.frm)
│   │   ├── Microsoft Excel Objects/  # VBA worksheet class files (.cls)
│   │   └── Modules/                  # VBA macro files (.bas)
│   └── xlsm/                         # Supporting files contained in the workbook
│       ├── _rels/                    # Manages ribbon xml for different versions of Excel
│       ├── customUI/                 # VBA worksheet class files (.cls)
│       │   ├── _rels_/               # Manages logical names for ribbon images
│       │   ├── images/               # Images used in the custom ribbon
│       │   ├── customUI.xml          # Ribbon definition for Excel versions prior to 2010.
│       │   └── customUI14.xml        # Ribbon definition for Excel versions 2010 and later.
│       └── docProps/                 # Excel workbook document properties
├── .gitignore                        # Git ignore rules
├── .npmignore                        # NPM ignore rules
├── LICENSE                           # Project license (MIT)
├── README.md                         # This file
├── package-lock.json                 # Detailed snapshot of NPM project dependencies
└── package.json                      # Manifest for Node.js VuePress project
```

## Installation

The *Relationship Visualizer* spreadsheet operates on both **Microsoft Windows** and **Apple macOS**<sup>1</sup>. 

Installation procedures vary by platform, and you can find detailed, platform-specific instructions by following the links below.

| <center><a href="./docs/install-win/README.md"><img src="./docs/install/winos.png" /></a></center> | <center><a href="./docs/install-mac/README.md"><img src="./docs/install/macos.png"/></a></center> |
| ------------------- | ------------------------------- |
| <center>[Microsoft Windows Installation Instructions](./docs/install-win/README.md)</center> | <center>[Apple macOS Installation Instructions](./docs/install-mac/README.md)</center> |

<small>[1] SQL and Clipboard features are not available on Apple macOS.</small>

## Documentation

**Recommended: Full online documentation**

Visit **[exceltographviz.com](https://exceltographviz.com/)** for comprehensive, up-to-date guides:  
- Installation (Windows & macOS)  
- Creating & styling graphs  
- Using SQL queries  
- Publishing & animating SVGs  
- Advanced Graphviz techniques  

**In this repository**  
- Website source (Markdown files built with VuePress): [`docs/`](./docs/) — powers exceltographviz.com  
- Legacy documentation archive (older .docx & .pdf manuals): [`legacy_docs/`](./legacy_docs/) — historical reference, platform-specific
  
## License

[MIT License](./LICENSE) — free to use, modify, distribute.

## Support

If this saves you time, consider [Buy Me a Coffee](https://www.buymeacoffee.com/exceltographviz) — helps cover hosting and dev costs.

Get started today: 

[![Download Excel to Graphviz](https://a.fsdn.com/con/app/sf-download-button)](https://sourceforge.net/projects/relationship-visualizer/files/latest/download)


