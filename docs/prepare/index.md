---
title: Create a New Workbook for Graph Building
description: Create a new Relationship Visualizer workbook, save it as a macro‑enabled file, resolve Excel security prompts, and ensure the custom Graphviz ribbon loads correctly.
---

# Create a New Workbook

The first action is to launch Microsoft Excel. When Excel starts, it will suggest sample spreadsheets you can create. This will contain the `Relationship Visualizer` template which you saved as a template as part of the installation steps. 

Select this template to create a new workbook.

![Excel start screen showing the Relationship Visualizer template available for creating a new workbook.](./create_new_workbook.png)

## Save the Workbook as a Macro-Enabled Workbook

The workbook will appear as shown below.

![Newly created Relationship Visualizer workbook displaying the default data and Graphviz ribbon tabs.](./new_workbook.png)

Perform a "**FILE -> Save As**" action. Choose a directory where you would like to save the file and change the file name from `Relationship Visualizer1` to something meaningful to you.

The most important step is to set the `Save as type:` dropdown list item as **Excel Macro-Enabled Workbook (*.xlsm)**. You will not be able to run the macros that create the visualizations unless the workbook is _macro-enabled_.

![Excel Save As dialog showing the Save as type field set to “Excel Macro‑Enabled Workbook (*.xlsm)” for enabling VBA macros in Relationship Visualizer.](./save_as.png)

The workbook you just saved may show a **BLOCKED CONTENT** message. If so, click the `Trust Center` button.

![Microsoft Excel security warning banner showing blocked content with a button linking to Trust Center settings.](../media/blocked.png)

The security settings for running macros will be displayed. 

Choose the `Enable VBA macros (not recommended; potentially dangerous code can run)` radio button, and press `OK`.

![Excel Trust Center macro settings dialog showing the option to enable VBA macros for running Relationship Visualizer automation.](../media/trust_center.png)

## Close and Reopen the New Workbook

Assuming that you changed the file name from `Relationship Visualizer1` to something meaningful to you, you should now close the file and reopen it.

When you reopen the workbook the message stating that macros have been blocked will be gone. The spreadsheet will appear as follows, displaying a `data` worksheet and a custom ribbon tab named `Graphviz`.

![Reopened Relationship Visualizer workbook showing the data worksheet and the custom Graphviz ribbon tab after enabling macros.](./reopen_workbook.png)

::: warning WARNING - Ribbon Fails to Update Dynamically After “Save As”
When you use **File → Save As** to change the workbook’s file name, Excel continues to associate the ribbon with the *original* file name. Because of this stale reference, any code that programmatically switches ribbon tabs will stop working.

To work around the issue, you can either manually switch tabs as you move between worksheets, or close and reopen the workbook. Reopening forces Excel to reload the ribbon under the new file name, restoring normal tab‑switching behavior.

This is a known issue in **Microsoft Excel** that affects workbooks using a custom ribbon ([1](https://stackoverflow.com/questions/33673898/macro-button-under-customized-ribbon-tab-tries-to-open-old-excel-file), [2](https://www.mrexcel.com/board/threads/custom-ribbon-macros-point-to-old-workbook.1257482/)). 

:::

::: tip
Any time you save a copy of the spreadsheet using **File → Save As** and change the workbook’s file name, you should close the workbook and reopen it. This forces Excel to reload the custom ribbon under the new file name and restores normal tab‑switching behavior.
:::

