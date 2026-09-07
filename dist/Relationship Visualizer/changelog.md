# Change Log

## Version 11.0.0 - September 1, 2026

Version 11.0 is the biggest release in **Relationship Visualizer**'s history, built around two headline features: exporting your data as an AI-ready Knowledge Graph, and a dramatically more capable SVG diagram viewer with pan, zoom, filtering, and highlighting built right in. Alongside those, the ribbon has been reorganized from top to bottom, SQL and Workbook Exchange both gained new capabilities, and a handful of long-standing rough edges were smoothed out. A short list of breaking changes, all low-impact, is included at the end, in case you want to check them before upgrading.

### Knowledge Graphs: A New Way to Export Your Data

Alongside the diagrams you already generate, Relationship Visualizer can now export your worksheet data as a Knowledge Graph — a structured JSON document designed to be pasted straight into an AI tool, rather than rendered as a picture. It's a second, parallel output format that runs alongside the existing Graphviz diagram pipeline, built from the same worksheet data and following the same rules for labels, styles, and inheritance, so the two outputs never disagree with each other.

Each export includes:

- Helpful metadata: format and version, whether the graph is directed, the source workbook and view name, and an export timestamp.
- A trimmed "styles" section that lists only the styles actually in use, and only for node, edge, and cluster types — keeping the file smaller and easier for an AI tool to work with.
- The same label and tooltip inheritance you already use in your diagrams: set a default at the `node[]`, `edge[]`, or `graph[]` level, or within a cluster, and individual rows only need to specify what's different.
- Automatic fan-out for edges: a single row with a comma-separated list of sources or targets expands into multiple discrete edges.
- Automatic placeholders for nodes referenced before they're formally defined, filled in once the real definition is reached.
- A clear warning, instead of silent data loss, for rows using native DOT passthrough, since that syntax isn't representable in JSON.
- A built-in token estimator, so you can gauge roughly how many LLM tokens your exported JSON will consume before you paste it somewhere.

**Typed, arbitrary properties.** Beyond the built-in label, tooltip, and style fields, every node and edge can now carry its own free-form properties. That is,  arbitrary key/value pairs read straight from your worksheet's new Properties column. Write them the same way you'd write extra Graphviz attributes: space-, comma-, or semicolon-separated `key=value` pairs, with quotes around any value that contains spaces or punctuation. `weight=200 domestic=true opened=2019-03-14` comes out as a real number, a real boolean, and a real date without everything being flattened to text. (Dates are deliberately kept as text rather than true date values: converting them through the JSON export's time-zone handling could otherwise shift a date-only value back a day for anyone west of Greenwich.) Properties follow the same graph → node/edge → row inheritance as labels and tooltips, so you can set a default once at the `node[]` or `edge[]` level and override it only where it needs to differ.

**Under the hood.** Making all of this possible required rebuilding how the add-in resolves conflicting label, tooltip, and style information for a given row. Previously, a label could come from the data worksheet, a shared style, or a row-level override, and the code simply concatenated whichever pieces applied into one Graphviz attribute string. If more than one source supplied the same attribute, Graphviz itself decided the winner by taking whichever occurrence came last. In v11.0, a single, well-defined precedence resolves each field before anything is written out. Row-level overrides win, then style templates, then the worksheet data, producing cleaner Graphviz source and guaranteeing the diagram and the Knowledge Graph agree on every label and tooltip. One visible side effect of this change is called out under Breaking Changes below.

#### Viewing Your Knowledge Graph

Excel has no native way to display JSON, so rather than dumping raw text into a worksheet cell the way Graphviz source is shown today, the Knowledge Graph opens in your default browser instead. The viewer's HTML lives inside the workbook itself, as a hidden worksheet. There's no companion file to lose track of, the workbook stays a single self-contained file, and it works identically on Windows and Mac (including trickier sandboxed folder permissions on macOS) with no ActiveX, UserForms, or Trust Center prompts involved. The viewer's own interface is fully localized, in all six supported languages.

A new split button on the Data tab sends the current view to the viewer. Two mutually exclusive modes control how it behaves: display it in a shared tab that quietly refreshes each time you regenerate the graph (handy if you keep flipping between Excel and the browser) or open a fresh tab every time. If you're using the shared-tab mode and accidentally close the browser tab, a "reopen" option in the same dropdown brings it back without regenerating anything.

### A Diagram Viewer Built for Big Graphs

Large diagrams present a real viewing challenge. Zoom out far enough to see the whole thing, and every label is unreadable. Zoom in far enough to read anything, and you've lost all sense of where you are. Version 11.0 adds a set of enhancements to SVG post-processing aimed squarely at that problem, and it's the second headline feature of this release. Everything gets built directly into the SVG files Relationship Visualizer exports, with no separate viewer, plugin, or install step. Open the SVG in a browser and it's already there.

> **A note on trust.** Postprocessing rules are trusted code, not a style preference. Whoever controls the find/replace rules controls what code runs in every diagram the workbook produces. If a postprocessing configuration is ever shared, exported, or imported between users, treat it with the same scrutiny you'd apply to a macro-enabled workbook because functionally, it is one.
>
> This feature ships turned off. If you've obtained the spreadsheet from anywhere other than the official download site, proceed very cautiously before turning it on.

Turning it on now requires explicit confirmation. A status indicator shows whether postprocessing is currently on or off, and a matching pair of On/Off buttons replaces the old single toggle. What used to be a single click is now a deliberate two-step. Clicking `On` pops this warning before anything happens:

> Post-processing injects JavaScript into exported SVGs. This JavaScript runs automatically whenever the SVG file is opened in a browser, with full access to that page.
>
> Only enable this feature if you trust the find/replace rules being used.
>
> Continue?

Turning it `Off` remains a single click. This same risk is also documented on the [Security](https://exceltographviz.com/security/) page.

#### A Toolbar That Doesn't Get Lost

The core idea is simple: controls that stay where you put them, at a size you can actually use, no matter how far you've zoomed into the diagram itself. Sounds obvious. Getting there was less obvious. SVG has no built-in concept of "stay fixed on screen while everything else zooms," so the toolbar, the zoom-percentage readout, and the small zoom buttons on each cluster all had to be taught to counteract the diagram's own zoom in real time. The payoff is that a 500-node diagram and a 5-node diagram now feel like the same tool, just at different scales.

#### Finding Your Way Around a Huge Diagram

A few things came out of that same foundation:

- **Scroll to zoom, drag to pan**, the way you'd expect from a map or a design tool, centered on wherever your cursor is.
- **A live hover label** built into the toolbar. Sweep your mouse across a dense cluster of tiny nodes and read off their names one by one, without needing to zoom in first just to see what you're looking at.
- **Fit Width / Fit Height / 100%** buttons, plus a zoom-percentage readout, so you always know exactly how zoomed in you are and can snap back to a sane view in one click.
- **Per-cluster zoom buttons** that stay a comfortable, constant size whether the whole diagram is zoomed way out or you're already zoomed halfway in. They are easy to find when you need them, unobtrusive when you don't.

#### Filter and Highlight Without Losing Your Place

Beyond navigation, the toolbar also lets you toggle entire categories of elements on and off. Hide every edge and just look at nodes, say, or isolate one cluster. Click any node to highlight its connections, choosing whether you want to see what feeds *into* it, what flows *out of* it, or both.

### Ribbon Reorganization

Version 11.0 brings a significant ribbon expansion: a new Settings tab for controlling worksheet and tab visibility, generation and publishing tooling split out of the Graphviz tab into a new Data tab, and a round of smaller refinements across several other tabs.

#### New: Settings Tab

A brand-new Settings tab gives you one central place to show or hide individual worksheets and ribbon tabs, with dedicated groups for Command/Graph Options, Data (worksheet and tabs), Console, Exchange, Extensions, Help URLs, Launchpad, Source, SQL (Windows only), Styles, and SVG.

The Settings worksheet itself was redesigned to match: the old faux tabbed-folder styling is gone in favor of a simple black-and-white scheme, settings are grouped to mirror their new ribbon locations, and each group of rows shows or hides based on which button you press on the new Settings tab.

#### New: Data Tab

The old, single Graphviz tab has been split in two. A new Data tab now owns everything related to generating and publishing output, while the Graphviz tab itself is scoped down to Graphviz-specific layout and style options (see below).

**Visualize.** "Refresh Graph" is now a split button, with Automatic Refresh moved into its dropdown. The zoom-level dropdown now lists percentages high to low (150% → 5%) instead of low to high. The new Knowledge Graph button lives here too. See the Knowledge Graphs section above.

**Publish.** "Publish" is now a split button with a new "Open after publishing" option in its dropdown, plus three new checkboxes 1) Graph, 2) DOT, and 3) Knowledge that let you choose exactly which output files get created each time you publish.

**File Output.** A new render-engine group lets you pick which Graphviz renderer to use: Cairo, GD, GDI+ (Windows only), or Quartz (Mac only).

**Styling.** A new group collects style-related switches in one place: Apply Styles and Apply Attributes moved here (now checkboxes rather than toggle buttons), alongside Add Image Path, Transparent Background, and Rotate 90° CCW, which used to live inside a dropdown menu.

**Options.** The Node, Edge, and Graph menus moved here from the old Graphviz tab. Node and Edge each gained tooltip-inclusion controls plus options for what to show when the tooltip cell is blank. In addition their blank/default label submenus were flattened for easier access. "Force xlabel Placement" moved here from the old Graph menu, which itself has been removed along with its "Center Drawing" option (it didn't do anything). Taking its place is an entirely new Cluster menu, with the same label and tooltip controls now available for clusters too.

**Data Worksheet.** The Show Columns menu is now organized into labeled sections (comment, item, label, style, knowledge), with a new Show Properties toggle and the old Show Messages toggle removed. The Delete All Data button icon is now red, to better signal that it's destructive.

#### Graphviz Tab, Trimmed Down

Adding Knowledge Graph publishing meant the original Graphviz tab simply ran out of room. It's now scoped to genuine Graphviz-only options such as layout engine, splines, direction, and the like. A new toggle on the Launchpad tab lets you hide it entirely if you're focused purely on Knowledge Graphs.

A few specific changes:

- **Splines**: the Compound, Line, and Spline edge-routing options are back as live, selectable controls, after being disabled for some time due to ribbon space constraints.
- **Output Order**: a new Depth option joins the existing Breadth option, with updated icons for both.
- Three option groups that behave like radio buttons - splines, direction, and output order - no longer force a default selection when the underlying setting is blank. An unset value now shows no button pressed, rather than visually defaulting to a choice you never made.

#### Launchpad: Hide the Graphviz Tab

A new Graphviz toggle joins the existing Source and Console toggles on the Launchpad tab, letting you show or hide the Graphviz ribbon tab.

#### Styles Tab: Configurable Cluster Naming

Cluster style names, the names written onto a cluster's opening and closing brace rows, can now be built from a configurable Naming Pattern instead of a fixed suffix. The old Suffix (Begin)/Suffix (End) fields are renamed 'subgraph-open' Affix / 'subgraph-close' Affix, and a new Naming Pattern field controls where that value is inserted, using `{name}` and `{affix}` placeholders. The default is `{name} {affix}`, but you can just as easily make the affix a prefix instead of a suffix. All three fields were also widened to fit longer patterns. This applies consistently whether the cluster came from the Style Designer or from a SQL `PUBLISH` command (see SQL Enhancements below).

#### SVG Tab

The old single Postprocess toggle became the clearer on/off pair with status indicators and a risk-acknowledgment prompt described in the diagram viewer section above.

### SQL Enhancements

Two new commands let a SQL query publish straight to a Knowledge Graph instead of only a Graphviz diagram:

- **`PUBLISH AS KNOWLEDGE GRAPH`** publishes the current view as JSON, honoring your minify/indent preferences.
- **`PUBLISH ALL VIEWS AS KNOWLEDGE GRAPH`** does the same for every view, producing one JSON file per view.

View-column detection used by `PUBLISH ALL VIEWS` and its Knowledge Graph counterparts is now more reliable. It correctly finds the true last View column even when there's a gap between columns, instead of undercounting and stopping short.

Cluster style naming follows the same configurable Naming Pattern described under Styles Tab above, whether the cluster came from the Style Designer or a SQL `PUBLISH` command.

The worksheet's old `ErrorMessage` output-column mapping for SQL results (which only ever fired if a query happened to return a field literally named `ErrorMessage`) has been replaced with a `Properties` mapping, feeding directly into the typed Properties feature described under Knowledge Graphs above.

### Workbook Exchange (Import & Export)

Exchange, the lightweight way to share or version-control a workbook's styles, settings, SQL connections, and configuration as an external file, was updated to keep pace with everything above:

- Every new v11.0 setting now round-trips through export and import: Knowledge Graph and publishing options, the render-engine picker, the new Clusters settings, node/edge tooltip settings, the Naming Pattern and Affix fields (older Suffix-based export files are still read correctly), the new style Description column, and the new Properties column.
- One deliberate behavior change: the publish output directory is no longer restored from an imported file. Restoring it used to risk silently pointing the workbook at a folder that doesn't exist, or isn't accessible, on the importing machine, so it's now simply left blank.
- Importing styles now scrolls the Styles sheet into view first, so you're watching the top of the sheet, not wherever it happened to be scrolled to, as preview images regenerate. As each preview image is created and inserted, the view scrolls to keep it centered on screen, giving you clear visual feedback that the import is progressing. Screen updating is also suspended for the duration, which speeds up the import and avoids triggering other worksheet events partway through.
- If you import an older export file into a workbook whose Styles sheet has the new Description column, the importer now recognizes that the View columns have shifted and skips restoring a stale column reference, rather than pointing at the wrong column.
- The Yes/No View-switch columns now get green/red conditional formatting applied automatically after import, instead of importing as unstyled cells.

### Startup Improvements

Two small changes when you open the workbook:

- A temporary folder for HTML/JSON output is now created at startup, right alongside the existing color/font image cache folders, so it's ready before the Knowledge Graph viewer needs it rather than being created on first use.
- A version check keeps the Style Designer's color and font preview galleries in sync with the running workbook. Each cache folder gets a small marker file recording which version last built it; if that marker is missing, the cache is treated as stale and rebuilt automatically. Practically, this means anyone upgrading from v10.5 or earlier will see their Style Designer previews silently refresh the first time they use a color or font control in v11.0, rather than risking a mismatched, stale preview.

### Other Small Refinements

A few smaller changes that don't fit neatly under any one heading above:

- You can now force a Graphviz label to be explicitly blank, rather than omitted, by entering an empty pair of quotes.
- Debug labels (enabled via the Debug switch on the Graphviz tab) now respect the Include Edge Ports setting, and can attach to more attribute types than just node and edge labels.
- Alt text for inserted diagram images is now localizable rather than fixed in English.

### ⚠️ Breaking Changes & Upgrade Notes

The changes below are technically breaking, but real-world impact should be minimal for almost everyone. They're included here for completeness, and so you know what to check if something looks different after upgrading.

- **Automatic `shape=plaintext` for HTML-like node labels has been intentionally removed.** In v10.5 and earlier, any node with an HTML-like label and no other style attributes automatically got `shape=plaintext`. That implicit override has been dropped in v11.0 in favor of a truer, unopinionated Graphviz experience. Graphviz's own default shape now applies unless a style specifies otherwise. Nodes that relied on this implicit behavior will now render with Graphviz's default shape instead of `plaintext`. This is expected to affect very few users; anyone who needs the old appearance back can add `shape=plaintext` explicitly, either as a style attribute or as a row-level extra attribute.

- **Error reporting moved off the data worksheet.** Row-validation errors used to be written into a dedicated "error message" worksheet column, which was shown and hidden around each generation run. Version 11.0 instead builds a localized message and routes it through a message/log channel, where it appears in the console or a message box according to your preference; the worksheet error-message column and its show/hide behavior have been removed entirely. The `!` indicator in column 1 has been retained to flag rows where an error was detected.

- **Internal code cleanup, with no visible effect.** Ribbon callback functions were tightened from public to private scope so they no longer clutter Excel's Assign Macro dialog, and leftover RubberduckVBA lint-suppression comments were removed now that AI-assisted code review has taken over that role. Neither change affects how the ribbon or any feature behaves.

- **Progress bar code removed.** The progress bar shown for long-running operations in older versions of the spreadsheet was turned off back in v10.3 and left dormant. In v11.0, the unused code has finally been removed.

## Version 10.5.0

This release focuses on usability enhancements, performance improvements, and important bug fixes.

**Usability Improvements**
- **Improved Style Designer ribbon galleries** - Increased readability and usability of the color and font galleries. Font name previews now use black text on white background for better contrast, and color swatches have been enlarged and changed to filled circles. [(#24)](https://github.com/jjlong150/ExcelToGraphviz/issues/24)  
  **Note:** You must clear your image cache to see these changes:
  - Go to the **Launchpad** tab -> make the `Diagnostics` worksheet visible.
  - Go to the **Diagnostics** tab -> click **Delete Colors** and **Delete Fonts** buttons.
  - Close and reopen the workbook, then switch to the `Style Designer` worksheet to regenerate the previews.

**Performance Improvements**
- **Improved Style Name dropdown performance** - Significantly optimized dropdown population using bulk array reads and per-rowType caching. Much faster response when switching between rows. [(#25)](https://github.com/jjlong150/ExcelToGraphviz/issues/25)

- **Improved data cell read performance** - Optimized how cell values are read during graph generation. This greatly reduces overhead during automatic rendering and makes editing much more responsive on large sheets. [(#19)](https://github.com/jjlong150/ExcelToGraphviz/issues/19)

- **Improved AutoDraw / rendering performance** - Removed unnecessary loading of SQL and SVG settings during normal worksheet rendering. Combined with the cell read optimizations, this results in faster graph generation. [(#20)](https://github.com/jjlong150/ExcelToGraphviz/issues/20)

**Bug Fixes**
- **Improved error handling in CreateGraphWorksheet** - Refined error handling so that issues during graph rendering (file operations, Graphviz execution, picture insertion, etc.) are no longer silently suppressed. Errors are now properly surfaced for easier diagnosis. [(#18)](https://github.com/jjlong150/ExcelToGraphviz/issues/18)

- **Fixed duplicate DOT string updates** - Optimized `CreateGraphWorksheet` to send the DOT string to the Source viewer only once per render instead of twice. [(#21)](https://github.com/jjlong150/ExcelToGraphviz/issues/21)

- **Fixed default output directory handling** - Corrected logic so that when no output directory is specified, graphs are now properly saved next to the workbook (as originally intended). [(#17)](https://github.com/jjlong150/ExcelToGraphviz/issues/17)

- **Fixed exported VBA file encoding** - Removed extended characters from all source code comments so that exported `.bas` and `.cls` files display correctly on GitHub. Files are now consistently saved in Windows-1252 (ANSI) without encoding issues. [(#23)](https://github.com/jjlong150/ExcelToGraphviz/issues/23)

- **DeepWiki indexing cleaned up** - Fixed mismatched folder names that caused routine files to be indexed. The exclusions were corrected, the guidance note clarified, and the cache refreshed. This removes about two dozen noise files while keeping the existing 27‑page structure intact. [(#22)](https://github.com/jjlong150/ExcelToGraphviz/issues/22)

## Version 10.4.0

This release provides several small but meaningful usability improvements.

**Enhancements**

The Relationship Visualizer received the following improvements:

- **Enhanced placeholder engine for labels** - Node, edge, cluster, and graph label builders now support template‑driven placeholders (`{label}`, `{xlabel}`, `{taillabel}`, `{headlabel}`), allowing `styles` worksheet formats to dynamically expand or fall back to data‑layer values. 

- **Smoother, cleaner AutoDraw updates** - `AutoDraw` has been revamped so graph updates now happen across more events in a single clean pass, without screen flashes or repeated triggers. The result is a more consistent, more polished *live preview* experience while you edit your data.


- **Smarter preview updates in the Styles sheet** - The preview image now updates automatically whenever you modify a style row (node, edge, or cluster), giving you instant visual feedback as you fine‑tune your formatting.

- **Standardized encoding for Graphviz integration** - `ExecuteAndCapture` now uses UTF‑8 for all command input and output, fixing an internal mismatch where the function previously sent UTF‑8 to Graphviz but returned Unicode on stdout and stderr.

- **Comprehensive module documentation for DeepWiki** - All VBA modules have been updated with clear, structured header comments to improve DeepWiki's analysis and generate more accurate, better‑organized documentation throughout the project.
  
## Version 10.3.0

This release continues to expand the **SQL capabilities** of the Relationship Visualizer.

**Enhancements**

The Relationship Visualizer received the following improvements:

- **Added support for “n” levels of clusters** – Previously, automatic clustering was limited to two levels via the `CLUSTER` and `SUBCLUSTER` column names. SQL clustering now supports an open‑ended number of levels by using integer‑suffixed field names. [(#13, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/13)

  For example:
  
  `SELECT [continent] AS [CLUSTER1], [country] AS [CLUSTER2], [state] AS [CLUSTER3], [county] AS [CLUSTER5], [city] AS [CLUSTER6] ...`

  Each level supports its own label, tooltip, style name, and attributes (e.g., `[CLUSTER1 LABEL]`, `[CLUSTER1 STYLE NAME]`, `[CLUSTER1 ATTRIBUTES]`, `[CLUSTER1 TOOLTIP]`).

- **Added `{label}` placeholder in cluster labels** – Introduced a `{label}` substitution token for cluster labels. This update allows HTML‑like formatting strings to be stored in the saved style format, enabling enhancements such as embedding an image next to the label value. When the graph is rendered, the label value from the `data` worksheet replaces the `{label}` placeholder.
  [(#14, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/14) 

- **Revised format‑string parsing** – Updated the logic that interprets saved style formats to correctly recognize HTML‑like syntax and pass it through to the Style Designer when the **Edit Style** button is used.  [(#15, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/15) 

- **Revised refresh‑all‑previews behavior** – Removed the progress‑bar dialog displayed when refreshing all style previews. Percent‑complete progress is now shown in the status bar. [(#16, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/16)

## Version 10.2.0

This small release focuses primarily on expanding the **SQL capabilities** of the Relationship Visualizer.

### Enhancements

The Relationship Visualizer received the following improvements:

- **Added SQL placeholder substitution** – SQL extensions now support `SET PLACEHOLDER name = value` definitions and expand `{name}` tokens in statements before execution, enabling template‑style SQL for parameter‑driven queries.
- **Added filename sanitization** – Introduced a routine that replaces invalid characters in filenames to prevent file write errors.

## Version 10.1.0

This release focuses primarily on expanding the **SQL capabilities** of the Relationship Visualizer.

### Enhancements

The Relationship Visualizer received the following improvements:

- **Concatenation Mode for Iterative SQL Queries** - Combine multiple detail records into a single string per header row (e.g. list of songs under an album, list of countries per continent, etc.).
  
- **Floating Action Buttons** - Context-sensitive floating icons now appear next to selected cells in supported columns. Each button triggers a specific action:
  - *'sql' Worksheet*
    -  `✎` Edit this SQL statement
    -  `▶` Run this SQL statement
    -  `⌕` View query status using a dialog 
  - *'styles' Worksheet*
    - `✎` Edit this style in the Style Designer
    - `↻` Refresh this style's preview image 

  Buttons auto-appear/disappear on selection, respect sheet protection, and support per-button validation (e.g. SQL only shows "Run" when SQL row is active and begins with `SELECT`).

- **Image Zoom Dropdown List** - Restored the Image Zoom dropdown list which was removed in the V9.0 UI refresh. The v9.0 Zoom In (+) and Zoom Out (-) buttons, and the 5% increments from 5%-150% have been retained.

## Version 10.0.0

This release focuses primarily on expanding and strengthening the **SQL capabilities** of the Relationship Visualizer.

### Enhancements

The Relationship Visualizer received the following improvements:

- Added support for using **Microsoft Access** (`.accdb` / `.mdb`) files as SQL data sources. [(#5, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/5) 
- Added SQL **enumeration** support for generating range‑based result sets (e.g. `from x to y by z`). [(#6, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/6)  
- Added SQL **iterative query‑set** execution with dynamic placeholder substitution. It allows a result from one query to be used as a parameter in a second query. [(#7, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/7) 
- Added SQL **error logging** to capture query failures in an external log file. [(#9, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/9)
- Improved usability by abbreviating long file‑system paths in the `SQL` and `Graphviz` ribbon tabs. [(#10, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/10)  
- Implemented a comprehensive set of **ADO SQL hardening changes**, improving reliability, determinism, and fault‑tolerance across the entire execution pipeline. [(#11, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/11)  
- Expanded the SQL log‑to‑file feature to include **environment documentation**, improving diagnostics and reproducibility. [(#12, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/12)

### Defect Fixes

The following defects were corrected:

- Addressed an intermittent SQL execution failure caused by underlying COM/Automation instability in VBA. [(#4, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/4) 
- **Breaking change**: SQL query clustering previously grouped results by **CLUSTER LABEL** instead of **CLUSTER**, and failed when cluster labels were null. [(#8, closed)](https://github.com/jjlong150/ExcelToGraphviz/issues/8) 

  The fix correctly groups by **CLUSTER**, but this correction is **not backward‑compatible** and may alter the structure of existing graphs that relied on the prior (incorrect) behavior. You will have to modify cluster-oriented SQL statements if you use this feature.

## Version 9.1.0

### Font Name Improvements

Revised the logic used to build the font list in the **Style Designer**.

- Expanded the available font choices by removing the earlier prefix/suffix‑based screening logic that filtered out fonts.
- Added an **excluded fonts** list on the `lists` worksheet. Testing shows these fonts fall back to a default font in Graphviz, so they are filtered out of the available font list.
- Cross‑checked the retrieved font list against the excluded‑fonts list to produce a clean set of fonts that Graphviz can reliably render.
- Added logic to **remove duplicates** and **alphabetize** the final font list.

Restructured the Windows and macOS code paths so both platforms now share the same filtering, deduplication, and sorting pipeline.  
- macOS still relies on the font list stored in the `lists` worksheet, but the code is now structured to allow a native macOS font‑enumeration solution to be added cleanly when available.

Updated the **Style Designer** ribbon tab to use a two‑image preview approach.  
- The ribbon continues to display the legacy (backward‑compatible) `A` icon beside the font name.  
- The font gallery now shows a wider 14pt preview image using the characters `Aa Bb Cc` to provide a clearer representation of each font.

## Version 9.0.0

### UI Visual Refresh

Replaced all [built-in Office Ribbon icons](https://spreadsheet1.com/microsoft-office-excel-ribbon-imagemso-icons-gallery.html) with [Google Material icons](https://fonts.google.com/icons).  
- Modernized the overall appearance  
- Improved visual clarity and contrast  
- Ensured consistent iconography across Windows and macOS  

**Breaking** - Several dropdown lists were redesigned as individual buttons that behave like radio buttons.
- Graphviz tab: **Zoom**, **Layout**, **Splines**  
- Style Designer: **Scale**, **Image Position**  
- Simplifies interaction and reduces misclicks  
- Enables dedicated tooltip text for every layout, spline mode, and image-position option  

Overall, the refresh makes the Ribbon cleaner and easier to scan at a glance.

### New Features

**Silent Mode / Message Routing**  
You can now disable message boxes entirely for "run silent" operation, as requested in this [issue](https://github.com/jjlong150/ExcelToGraphviz/issues/3).  
Errors and notifications can be routed to:  
- Message boxes  
- The Excel status bar  
- The Console worksheet  

These options are controlled via enablement buttons on the **Console** tab, giving users full control over how the tool communicates during interactive or automated operations.

**Support for Graphviz 14.1.0 `radius` Attribute**  
Graphviz introduced a new `radius` attribute for rounding corners on orthogonal edges.  
The `Edge` mode of the **Style Designer** tab now includes a Gallery control that visually previews radius values from 0 to 20, making it easy to choose the desired corner roundness.

---

### Improvements

**Breaking** - Image Zoom redesigned
- Replaced fixed zoom levels (25%, 50%, 75%, 100%)  
- New range: **5%-150%** in **5% increments**  
- Controlled via **Zoom In** and **Zoom Out** buttons  
- Provides finer control and a smoother editing experience  

**Windows Ribbon Performance**  
Removed the one-second delay when switching Ribbon tabs after selecting a worksheet.  
- Windows now switches tabs instantly  
- macOS retains the asynchronous delay to avoid race conditions in Excel's event model  

**Accessibility Cleanup**  
Resolved all Excel "Accessibility" warnings to ensure a cleaner, more compliant workbook environment.

**Internal Code Quality Enhancements**  
- Replaced numerous string literals with named constants  
- Improved maintainability and reduced risk of typos  
- Strengthened the error-reporting pipeline with locale-aware timestamps, normalized messages, and safer numeric parsing  

## Version 8.0.1
### Fixes
- A previous update resolving image deletion on low-memory systems inadvertently disabled SVG removal. This regression has now been corrected. 

## Version 8.0.0

### Style Designer

Replaced dropdowns with Ribbon galleries; visual grid-style controls that display selectable items like colors or images for faster, more intuitive style selection.
- Color galleries now display entire color schemes in a compact, high-speed format.
- `fontname` previews are shown as a gallery alongside font names, with increased preview size for easier identification. (Note: existing font images may be deleted via the `diagnostics` tab.)
- Node `shape`, edge `arrowhead`, `arrowtail`, `headport`, and `tailport` choices are now logically grouped for easier selection.

Added RGB Color Picker support
- Enables color selection using the operating system's native RGB dialog on both Windows and macOS.
- On macOS, this requires updating the `ExcelToGraphviz.applescript` file to version 3.
- The picker can be launched independently or preloaded with a color from Graphviz X11, SVG, or Brewer schemes.

Improved preview rendering
- Updated color and font previews to display selections directly within the Ribbon.
- Refreshed icons for X11 and SVG color schemes, as well as polygon shape options.

Performance optimizations
- Faster algorithm (Win OS) to exclude font names which Graphviz converts to 'Arial'.
- Rewrote string comparisons and concatenations for faster execution.
- Removed progress bar when loading large color schemes. New speed makes them no longer needed.
- Images are now pre-cached at workbook open if available; otherwise, they're generated on first use.

Enhanced style saving
- Labels (`label`, `xlabel`, `headlabel`, `taillabel`) can optionally be saved as part of the style format string which is ideal for edge annotations like protocols or cardinality.
- Added support for naming styles when saving to the `styles` worksheet.
- Introduced a prominent `Save` button within the style definition canvas area.

Improved image path handling
- Automatically extracts relative image paths to improve portability.  
  Example: If the workbook is in `c:\users\jeff\data\example1\Relationship Visualizer.xlsm` and the image is in `c:\users\jeff\data\example1\images\network.png`, the saved path will be `images\network.png`.

Fixed image deletion issue on low-memory systems
- Addressed a bug on 32-bit Atom CPUs with 2GB RAM by switching to a more resource-efficient method for deleting preview images.

### Styles

One-click style restoration in Style Designer
- Select a row on the `styles` worksheet containing a `node`, `edge`, or `cluster` format string, then click the `[...]` button to instantly reset `style designer` to match the saved format.

Auto-refresh preview
- Saving a style automatically updates the preview image on the `styles` worksheet.

### SQL

Connection pooling added 
- Implemented in response to a March 2025 Office update that causes ADO connections to take over 12 seconds (previously under 4 milliseconds). See: [Excel ADO connection issue in recent Office 365 update](https://learn.microsoft.com/en-us/answers/questions/5443040/excel-ado-connection-issue-in-recent-office-365-up?forum=msoffice-all&referrer=answers)  
- Workbook connections are now reused across all SQL statements during a `Run SQL Statements` batch run.  
- Users can choose to close connections after batch execution or keep them open until manually closed or workbook exit.  
  Note: Keeping connections open may improve performance but can prevent access to referenced workbooks.

New default data source support 
- Users can now specify a default data directory and Excel workbook.
- Managed via new controls in the SQL tab of the Ribbon.
- Entries in the file name column, and `SET DATA FILE` statements take precedence when resolving conflicts.

SQL editor access
- Added a `[...]` button next to SQL statements to open the SQL edit form.

New SQL extensions for Graphviz automation 
- Assume the you have an Excel workbook containing a worksheet named `Alphabet` with a column heading of `letter` with four rows of data with letters A, B, C, and D in the `letter` column. 
  
  The following SQL creates nodes `A`, `B`, `C`, `D`:

  ```sql
  SELECT [letter] AS [Item] from [Alphabet$]
  ```
- New *CREATE EDGES* syntax automatically generates edges like `A -> B`, `B -> C`, `C -> D`. 

  ```sql
  SELECT [letter] AS [Item], TRUE AS [CREATE EDGES] FROM [Alphabet$]
  ```
- The new *CREATE RANK* syntax produces subgraphs with a shared rank:

  ```sql
  SELECT [letter] AS [Item], TRUE AS [CREATE RANK], 'same' AS [RANK] FROM [Alphabet$]
  ```
  The SQL above results in one row added to the `data` worksheet with: 
  - Item = `>`
  - Label = `{rank="same"; "A"; "B"; "C"; "D";}`

### SVG
- Added `[...]` button to Find and Replace cells to open the SVG editor form.
- Updated animation logic in one of the post-processing options to accept `1` or `0` for toggling inclusion of zoom buttons on clusters.

### Miscellaneous

macOS compatibility
- Fixed Excel version check to compare major and minor versions numerically, resolving a string-compare bug introduced when Excel reached three digits in version 16.100.
- Revised the `ExcelToGraphviz.applescript` script version check to use integers instead of strings
  - An alert triggers if installed script version is below 3.
  - The new RGB color controls are hidden when an outdated script is detected, preventing users from interacting with buttons that would otherwise fail silently.

Optimizations
- Introduced new constants throughout the codebase to improve clarity and maintainability.
- Further optimized string concatenation routines for better performance.

JSON Import/Export
- Updated to support exporting and restoring the new SQL settings.


## Version 7.2.02
### Fixes
- Enforce Graphviz ID naming rules by wrapping ID values in quotes when they do not begin with a letter or underscore.

## Version 7.2.01
### Fixes
- Resolved [Output Directory macro 'Cannot run' #2](https://github.com/jjlong150/ExcelToGraphviz/issues/2) defect.

## Version 7.2.0

### Enhancements

`sql` Worksheet

Added support for recursive SQL queries, enabling the creation of hierarchies such as organization charts and connected data paths.

- A candidate dataset should include a pair of related columns that enable hierarchical traversal. For example:
  - The columns `employee id` and `manager id` in a business organization contain the relationship between an employee, and the manager they report to. 
  - The columns `station` and `next station` in a subway map contain the point-to-point destinations of subway stations on a given subway line. 

- New SQL conventions provide the ability to:
  - *Anchor the Base Case* - where you define the starting point (such as the top-level parent in a hierarchy).
  - *Define the Recursive Member* - where you specify how to recursively connect results to construct subsequent levels.

- Explore the updated sample workbooks for practical examples:
  - `12 - Using SQL - Trees` - Shows how to traverse any node-to-node structure by starting at a node and iterating through either its connected predecessor nodes, or connected successor nodes.
  - `13 - Using SQL - Organization Charts` - Shows how to build a complete organization chart, or extract a branch of an organization chart.

### Fixes
- Resolved a screen flicker issue on the `style designer` worksheet that occurred during the initial creation of color images for ribbon dropdown lists.
- Incorporated enhancements in `ExecuteAndCapture` code based on recommendations from *RubberduckVBA* code inspections.

## Version 7.1.0

Fixed a bug that caused Excel to freeze when Graphviz wrote more than 4096 bytes of message output:

- The *Relationship Visualizer* spreadsheet uses the `ExecuteAndCapture` routine to run Graphviz's `dot` command and capture any `dot` output messages.
- `ExecuteAndCapture` reads messages via an interprocess pipe with a fixed size of 4096 bytes.
- `dot` paused after writing 4096 bytes of messages, waiting for `ExecuteAndCapture` to read and clear the data from the pipe before resuming.
- `ExecuteAndCapture` was paused, waiting for `dot` to finish before reading any data from the pipe.

This resulted in a deadlock. To eliminate the deadlock:

- `ExecuteAndCapture` now monitors the execution of `dot` in real-time, instead of waiting for it to complete.
- `ExecuteAndCapture` periodically checks the pipe for data. Any data found is read and removed from the pipe.
- `dot` is now able to pause and resume as needed until graph generation is complete and all messages are captured.


## Version 7.0.0

### Worksheet Changes

`console` Worksheet
- Added a **new** `console` worksheet which displays the Graphviz `dot` command's error & diagnostic messages.
- Added a **new**, associated `Console` ribbon tab which offers various logging options.
  
`diagnostics` Worksheet
- Added a **new** `diagnostics` worksheet which documents the environment in which the tool is being used.  
- Added a **new**, associated `Diagnostics` ribbon tab containing buttons which will clear the `Style Designer` image caches for fonts and colors when pressed.

`info` Worksheet
- Renamed the `about...` worksheet as `info`.
- Hid the `info` worksheet by default. The `info` worksheet containing all the credits and license information is now toggled via a button on the `Launchpad` ribbon tab.
- Added a **new**, associated `Info` ribbon tab with buttons for **Excel to Graphviz**-related web links (such as Github, SourceForge, Buy me a Coffee, etc.).
  
`source` Worksheet
- Added a **new** modeless pop-up window which displays the `dot` source used to create the graph image. This pop-up works in addition to the source display on the `source` worksheet making it easier to show the image and `dot` source simultaneously when conducting training.
- All Graphviz `dot` source code rendered in Excel is displayed in the `source` worksheet, and also in the new pop-up window. Previously only the source used to render the `data` worksheet was displayed. Now the source generated by the `style designer` is also viewable.
  
`sql` Worksheet
- SQL statements can grow in height beyond Excel's capabilities to display as a worksheet row, especially if multiple SQL statements are `UNION`ed into one large SQL statement. A **new** pop-up form was added which allows editing the contents of a cell containing a SQL statement.  This form allows you to see the entire contents of the cell, scroll through it, and edit it.
- Added `Copy to Clipboard` button which will copy SQL statements to the clipboard.
- Added the pseudo-sql statement `SET DATA FILE` which can be used to specify the file name of the workbook to be queried. This statement makes it easy to query different files when using the filter capability as you don't need to change the file name for every SQL statement.
- Added support to specify `CLUSTER LABEL` and `SUBCLUSTER LABEL` in SQL commands so that you can have cluster headings which are different than the value you are grouping by, or to null out the heading and just have a border.
- Added a utility routine `RunSQLAsExtension` which lets you run the SQL code from the `Extension` tab, so you can keep the `SQL` worksheet hidden.

`settings` Worksheet
- Added **new** `CLUSTER LABEL` and `SUBCLUSTER LABEL` identifier strings to the `SQL` settings.
- Added a **new** settings cell to store the name of the view currently being graphed. This enhancement allows you to write formulas that include the view name in the graph output, especially when using the 'All Views to File' graphing capabilities.
- Removed the timeout setting. See Miscellaneous changes below for reason.
  
`style designer` Worksheet
- Locked the `Labels` group of controls in one consistent location. They no longer shift left or right depending on `Node`, `Edge`, or `Cluster` radio button. `Labels` controls are now consistently on the right of the `Color Scheme` group. Dynamic controls such as `Shape` now begin to the right of `Labels`.
- Moved the `Colors` button to the `Launchpad` ribbon tab.
- Added a static label to show the name of the currently selected color scheme.
- Added a **new** group of controls for defining `packmode` attributes. The controls make it easy to specify the maximum number of components (subclusters or nodes) per row/column, row-major vs. column-major layout, shape alignments, and activation of user-defined sorting.
- Made UI changes to conserve ribbon space for lower resolution monitors
  - Moved the gradient fill color dropdown beneath the primary fill color dropdown. `Gradient Fill` options group now appears only after a second fill color is selected.
  - Image `Scale` and `Position` controls are only displayed if an image file name has been specified.
  - New `packmode` controls are only visible when the layout on the `Graphviz` tab is set to `osage`, and the design mode on the `style designer` tab is set to `cluster`.
- Made adjustments to the font preview image, as portions of the image were getting cut off by Graphviz version 11 and above. 
  - You will need to clear the **Excel to Graphviz** font image cache to see the change if upgrading from a prior version of this spreadsheet.
  - Refer to `Diagnostics` above to learn how to clear image caches.
- Improved the performance of displaying/resetting the `Style Designer` ribbon tab. The  time required on the author's PC to load 700+ images in the dropdown lists and display the tab was reduced from ~13 seconds on v6.1.01 of this spreadsheet to ~4 seconds on this new version. (Note: First time use still requires the 700+ images to be created and placed into a cache, and is not included in this timing). Performance improvements include:
  - Cached the `Gray*` preview images under the names `Gray*` and `Grey*` for the `X11` color scheme, resulting in 100 less images being loaded into memory. The 15% reduction from 656 to 556 cached images provides performance and memory use improvements.
  - Eliminated numerous Windows 11 font names as choices, as they are not recognized by the Graphviz pango font mapper.
  - Made refinements on when to handle events.
- Made improvements to the `dot` source code which generates the preview images. 
  - Cluster previews now show 7 nodes so the effects of `packmode` attibutes are visible.
  - Edge previews now provide language translations of "HEAD" and "TAIL" labels
- Added emoji and symbol fonts to the list of fonts.
  
`styles` Worksheet
- Added a **new** ribbon tab for the `styles` worksheet.
- Added the ability to generate preview images of the style definitions in either singular or bulk fashion.
  
`svg` Worksheet
- Added a **new** pop-up form which allows editing the contents of a cell containing a large replacement string. Replacement strings can grow beyond Excel's capabilities to display as a row. This form allows you to see the entire contents of the cell, and edit it. It also makes splitting post-processing directives across multiple rows unnecessary.
- Provided enhanced JavaScript for smoother SVG animation.
- Provided alternate styling for SVG animation more akin to macOS controls. This version is commented-out by default. You can choose which to use by commenting-out one or the other.
- Added `Copy to Clipboard`, `Graph to File` and `All views to File` buttons on SVG tab. `Copy to Clipboard` is not available on macOS as the clipboard code is Windows OS-specific.

### Ribbon Tab Changes

`Launchpad` Ribbon Tab
- Consolidated all the buttons for showing/hiding worksheets onto a **new** `Launchpad` ribbon tab.
- Provided buttons to show/hide previously hidden worksheets such as the language translations.

`Exchange` tab

- Updated to include the `CLUSTER LABEL`, `SUBCLUSTER LABEL` settings from `settings` worksheet in exports and imports.
- Updated to include `Launchpad` worksheet show\hide settings.

`Graphviz` tab

- Moved the `Show/Hide Worksheets` and `Language` controls to the `Launchpad` ribbon tab, as they are not Graphviz-related.
- Converted `Style` and `Debug` dropdown menu items to check boxes on the main ribbon in the new space freed up by moving the `Show/Hide Worksheets` and `Language` controls to the `Launchpad` tab.
- Added `Include Image Path` as a `Graph` check option so the image path can be omitted from `dot` source when images are not being used.
- Eased restrictions on when automatic drawing can occur so that ribbon tab changes can be observed in either the `data` or `graph` worksheet.

### Miscellaneous Changes

- Windows OS: Replaced the Open Source `ShellAndWait()` function used to run the Graphviz `dot` command with a new Open Source function `ExecuteAndCapture()` which can run the Graphviz `dot` command and return the messages which `dot` writes to the standard output, and standard error message pipes. These messages are then displayed on the new `console` worksheet. A tradeoff of this code replacement is that the timeout capabilities which `ShellAndWait()` provided are not present in `ExecuteAndCapture`.

- Windows OS: Replaced the code used to copy `dot` source code to clipboard with a new implementation which does not rely on an Internet Explorer ActiveX object.

- Eliminated adding quotes to most strings. This change reduces the number of string concatenations, which improves performance slightly. More important, it makes the Graphviz source easier to read, and use in other Graphviz editors which sometimes do not like the quoted strings.

- Created a `Graphviz` class for rendering graphs. It accepts the `dot` code as a string, and handles the writing to a file, executing `dot`, and returning any messages. 
  - Refactored all places in the code base previously writing files and using the `ShellAndWait()` function to convert the code to instantiate `Graphviz` objects and render graphs using this new approach (which greatly simplified the internal coding). 
  - You can now view any `dot` source code processed by the class on the `source` worksheet, such as the `dot` source generated by the `style designer` when creating preview images.
  
- Eliminated many of the named lists from the `lists` worksheet which were only used to determine the position of which dropdown item to select when the workbook was opened. The ribbon now directly generates and returns the dropdown id from the saved value, as opposed to searching the list and returning a relative integer position. This change simplifies future maintenance, and elimination of list searching will improve performance slightly.

- When a worksheet is activated, the ribbon switches to the tab which is most appropriate for the worksheet. This behavior was modified to introduce a one second delay and run in an asynchrous fashion as on macOS the ribbon controls need to complete any refresh before the tab can be switched. The asynchrous delay allows for pending events to process.

- Simplified the test which determines if a label is HTML-like, which makes it easier to specify labels which use simple HTML to do things like bold or itacize a portion of the label. For example `<This text should be <b>bold</b>>` will now work, where previously it would have required the whole label to be wrapped in HTML such as `<<p>This text should be <b>bold</b></p>>` .

- Revised an internal `SplitMultilineText` function to split lines by line breaks (such as \n) into an array, then apply the split logic to each line in the array, and concatenate the results. Previously the line breaks were included as part of the text string and were counted when determining where to make splits. It also lets you retain line breaks for things like splitting the label into multiple lines separated by a blank line.

- Accepted **RubberDuckVBA** suggestions to improve code quality.

## Version 6.1.01

Style Designer enhanced to include `Mrecord` in the list of shapes. 

## Version 6.1.00

Style Designer enhanced to show a progress indicator when loading large dropdown lists. 

## Version 6.0.03

`Style Designer` ribbon tab was enhanced to not load names and images of colors and fonts for hidden dropdown lists. This change significantly reduced the time to load the tab from ~15 seconds to ~8 seconds.

## Version 6.0.02

Swapped the original SVG postprocessing find/replace scripting with a new contribution which adds animated edge highlighting, and cluster zoom in/zoom out capability. 

## Version 6.0.01

`Style Designer` ribbon was enhanced to cache references to color and font images as they get loaded from the file system. 

For Graphviz's default `X11` color scheme, this change eliminates over 4,000 file system accesses, reduces the amount of memory used by the ribbon (6 copies of color images, and 1 copy of font images eliminated), and reduces the initial load time of the tab by a couple of seconds. 

The tab still takes about 15 seconds to load, but 15 is better than 17. Other color schemes load faster as well. 

## Version 6.0.00

New Online Help

- [Excel to Graphviz](https://exceltographviz.com/) now has its own website at https://exceltographviz.com/
    
- Ribbon tabs have been updated with `Help` buttons which take you to the online help for that tab.

Enhanced SQL Capabilities

- SQL queries can now specify spreadsheet columns to group as clusters, or group as subclusters within a cluster. 

	For example, a query such as 

	`SELECT [City] As [Item], [State] As [Cluster], [County] As [Subcluster] from [Geography$] WHERE [Population] > 100000` 

	would draw a cluster for each State. Within the State, clusters will be drawn for each County, and each County cluster will have nodes for each City found with a population greater than 100000. This type of query compliments the `osage` layout engine.

- Pseudo-sql statements have been added which allow you to run several SQL statements in a row, then output a graph. Statements take the form below, where words in uppercase are the new commands, and words in lowercase are graph prefixes:

  - `RESET`
  - `PREVIEW`
  - `PREVIEW AS DIRECTED GRAPH`
  - `PREVIEW AS UNDIRECTED GRAPH`
  - `PUBLISH`
  - `PUBLISH` *example01*
  - `PUBLISH AS DIRECTED GRAPH` *example01-directed-graph*
  - `PUBLISH AS UNDIRECTED GRAPH` *example01-undirected-graph*
  - `PUBLISH ALL VIEWS`
  - `PUBLISH ALL VIEWS` *example01-allviews*
  - `PUBLISH ALL VIEWS AS DIRECTED GRAPH` *example01-all-views-directed*
  - `PUBLISH ALL VIEWS AS UNDIRECTED GRAPH` *example01-all-views-undirected*

- Filtering. The SQL worksheet lets you add values to columns to the right of the SQL which can be filtered on to control which statements get executed. In this way you can define sets of queries and generate a batch of diagrams, or filter to a specific set to generate a single diagram.

- Label splitting. SQL queries can now split values used as labels at an approximate length which can be specified in the query. Logic finds the closest blank between words to use as the split location. You can also specify the line delimiter to center, left, or right align the text.

- Samples. A **new** Enterprise Architecture sample has been contributed which lets you fill in a reference spreadsheet, and it generates a set of EA diagrams which span the relationships in the various worksheets in the workbook. See SQL examples in the `samples` directory for running examples of all the new capabilities.

SVG Post-processing Capability

- Added a new worksheet and ribbon tab for post-processing graphs which are output as SVG files. The post-processing works by performing find/replace actions against the file using values on the `svg` worksheet. 

	The default implementation provided will insert JavaScript code which highlights nodes and edges with thick magenta lines if you click on them. Another use case scenario (not provided) would be to modify image references to provide a base URL.

	The `svg` ribbon tab contains an on/off switch to control if post-processing should be performed. The feature is turned off by default to ensure pure Graphviz SVG files are created OOTB.
 
Default `Graph to File` Directory

- Removed restrictions on the `Graph to File` button which disabled the button if an output directory was not specified. Now the current directory where the workbook resides will be used as the place to write the file if no other directory is specified.

Style Designer Feedback During Loading

- Graphviz's default color scheme is `X11` which has 656 colors. The `Style Designer` ribbon tab has 7 color dropdown lists. This results in 4,592 file accesses to fill the dropdown lists with preview images, which takes about 17 seconds on a fairly recent computer. The spreadsheet appears to be frozen during this period. A change was made to disable events during the loading of a list, but turn them on between lists to provide feedback in the status bar as to which list is being loaded. This small amount of feedback changes every 2 seconds or so and lets you know that the spreadsheet is working, and is not hung.

Sample Workbook Updates

- The workbooks and JSON files in the "samples" directory were updated to version 6.0.00.

## Version 5.8.00


Solid-fill Gradients for Nodes and Clusters 
- A new `Weight` percentage for the primary fill color has been added to the Style Designer which allows for the creation of solid-fill gradients. 
- The `Weight` dropdown list dynamically appears under the `Fill Color` dropdown once a value for `Gradient Fill Color` has been selected. 

Tooltip Support
- A new `Tooltip` column has been added to the 'data' worksheet which will add the `tooltip` attribute to clusters, nodes, and edges if the image format is `SVG`.
- The `Tooltip` column is hidden by default, and can be unhidden by selecting the "Tooltip" entry of the `Columns` list on the Graphviz tab.
- Office 365 displays SVG files in the Excel as a picture, older versions of Excel do not support SVG. To see the tooltip information, you must save the graph to a SVG file and view it in a browser such as Microsoft Edge which has support for SVG files. 
- Export/Import of the Tooltip data column has been added to the Exchange mechanism.
- The fontname in the predefined styles on the 'styles' worksheet were changed from "Calibri" to "Arial" to be more SVG and Mac friendly.

Sample Workbooks
- The workbooks and JSON files in the "samples" directory were updated to version 5.8.00.

File Size
- Redundant images created by the Office RibbonX Editor tool have been removed, which reduced the Excel workbook size by over 1MB.

## Version 5.7.01

Polish language translation originally created using the http://www.deepl.com free translation service have been reviewed and edited by Polish-speaking contributor Arek Czak.

Style Designer color dropdown lists now show the default color Graphviz will use when no color has been chosen as the first value in the list.

## Version 5.7.00
Added Polish language translation. The translation was created using the http://www.deepl.com free translation service (and may have errors). 

## Version 5.6.00

Fixed a small bug where the "outputorder" attribute was being emitted as "orderoutput".

Added `overlap`, `mode`, `smoothing`, `dim`, `dimen` attributes to the Algorithm group on the Graphviz tab. The show/hide display of these options are controlled by the layout algoritm selected. The ribbon will dynamically change to show only attributes which apply to that layout.

Continued to purge performance slowing string operations, helping to improve performance on older machines.

## Version 5.5.01

Made several code performance optimizations to the `getItemColor` and `getItemLabel` callbacks in the Style Designer ribbon menu which greatly reduce the time it takes to LOAD the lists of color choices names and preview images. 

Made changes which reduce the time it takes to CREATE the color preview images which are displayed in the color choice dropdown lists.

## Version 5.5.00

### Multiple Language Support

Introduced a **"Language"** dropdown to the Graphviz tab providing a capability to toggle different languages. 

Initial languages supported:
- English (US)
- English (UK)
- French
- German
- Italian

Items translated (using language translation software, and may have errors) include: 
- The ribbon interface tab names, ribbon groups, ribbon control labels, buttons, super tip help text.
- Worksheet names
- Column headings
- Error and status messages

Items not translated include: 
- The `HELP - Attributes` worksheet
- Style and View names on the `styles` worksheet
- Portions of the `settings` worksheet pertaining to worksheet column layouts
- The user documentation
- The sample worksheets

Hidden worksheets contain the language translations. You can edit these worksheets to make corrections. Simply unhide the appropriate worksheet which starts with `locale_`, correct the text, save, and close the workbook, and reopen the workbook.


### Metric Measurement Support

Added metric units for height and width to the Style Designer. 

A checkbox lets you toggle the lists between inch units (which Graphviz uses) and millimeters (which the Style Designer will convert to inches)

### Font Preview Images

Added font preview images to the Style Designer Font drop-down lists. 

You may notice a delay the first time the Style Designer ribbon is displayed while preview images are created using the list of fonts on your PC. This is a one-time installation activity, and these images remain cached for future use. 

In addition, Windows 11/Office 365 has introduced thousands of fonts and font variations (e.g., bold, italic). Code has been added to prune the font list to just the fonts which Graphviz can use.
