# Security 

## Excel Macro‑Enabled Workbook Considerations

When deciding whether to trust a macro‑enabled Excel workbook (such as an `.xlsm` file) downloaded from the Internet, several security factors should be evaluated to minimize risk. Macro‑enabled workbooks contain executable VBA code, which can pose significant security threats if malicious or poorly vetted.

The sections below outline the key considerations that help you make an informed, cautious decision before enabling macros or running downloaded Excel tools, and how these considerations apply specifically to the `Relationship Visualizer.xlsm` spreadsheet.

## Assess Reputation and Trustworthiness

Verify the **reputation of the repository owner**. Review their profile for activity history, contributions, and community engagement. Well‑established or transparent contributors are generally more trustworthy.

- A public professional profile for the author and repository owner is available on [LinkedIn](https://www.linkedin.com/in/jeffreyjlong/).
  
## Review Repository Activity

Review the project’s [Changelog](/changelog/), which documents the history of its releases and provides insight into development cadence and maintenance quality.

Evaluate the repository’s visibility and engagement. A well‑maintained project with stars, forks, active issues, and community participation is less likely to host malicious or abandoned code.

- The *Excel to Graphviz Relationship Visualizer* has published releases on **SourceForge** continuously since **October 17, 2015**.
- It holds a **five‑star rating** on SourceForge.
- SourceForge has awarded it a **Community Choice** badge, recognizing open‑source projects that have reached the milestone of **10,000 total downloads**.

Check for **code reviews**, comments, or discussions in the repository.

- Contributions from multiple users or public conversations can indicate scrutiny, transparency, and reliability.

Review the project’s [Issues Log](https://github.com/jjlong150/ExcelToGraphviz/issues).

- **Open issues** may highlight concerns raised by the community or known software defects.
- **Closed issues** provide insight into past problems and whether they were resolved promptly and responsibly.

## Inspect the Code

**Review the macro code** before enabling it.

- The VBA macro source code for the *Excel to Graphviz Relationship Visualizer* workbook is published on GitHub at
    
  <https://github.com/jjlong150/ExcelToGraphviz/tree/main/src>  
  
  It can be viewed directly in a web browser before downloading or opening the workbook.

**Open the workbook in a protected environment** and inspect the VBA code (Alt+F11 in Excel). Look for suspicious or unexpected actions such as:

- **Network connections** (e.g., accessing external URLs).  
  - The *Excel to Graphviz Relationship Visualizer* includes hyperlinks to online help resources and Graphviz rendering sites, but does not perform automated network calls.

- **File system modifications** (e.g., creating, deleting, or altering files).  
  - The workbook generates text and image files as part of its normal operation.

- **Execution of external programs or scripts.**  
  - The workbook invokes the Graphviz `dot` command‑line program to convert `.dot` source files into rendered graphs.

If you lack VBA expertise, **consider using online tools or sandbox environments** to analyze the code for malicious behavior.

All macros in the *Excel to Graphviz Relationship Visualizer* undergo static code analysis using [RubberduckVBA](https://rubberduckvba.com/) prior to publication. Rubberduck provides inspections, code metrics, and quality checks that help ensure maintainability and reduce the likelihood of hidden or unsafe behavior.

## Verify Files

One of the most important security steps is knowing **exactly where the file came from**. Avoid downloads from unauthorized mirrors, repackaged distributions, or any file not obtained directly from the official project pages, as these sources may compromise integrity or security.

### Use only official releases

Official releases of the *Excel to Graphviz Relationship Visualizer* workbook are published exclusively on:

- **SourceForge** → <https://sourceforge.net/projects/relationship-visualizer/>
- **GitHub** → <https://github.com/jjlong150/ExcelToGraphviz/releases>

### Protect from viruses

- SourceForge performs automatic scanning before publishing files.
- GitHub does not scan release assets, so user‑side scanning is essential.

| Platform        | Virus Scanning of Uploaded Files | Notes |
|-----------------|----------------------------------|-------|
| **SourceForge** | ✔️ Yes — automatic antivirus scanning on upload | Files are only made available for download **after** passing a successful virus scan. |
| **GitHub**      | ❌ No — GitHub does *not* scan release assets | GitHub security tools (Dependabot, code scanning) apply to **source code**, not binary release files. Users must scan downloaded files themselves. |

Regardless of platform, always scan downloaded files with **reputable antivirus tools**.

### Check file hash values

Verifying published hash values ensures the file you downloaded is **identical** to the one the author released.

| Platform        | Hash Types Provided | Where to Find Them | Notes |
|-----------------|---------------------|---------------------|-------|
| **SourceForge** | **SHA1**, **MD5**   | On the **Files** tab → click the circle **(i)** icon next to each file | Hashes are shown for every downloadable artifact. |
| **GitHub**      | **SHA256**          | In the **Assets** section of each release | Hashes accompany ZIP files containing the workbook and source code. |

::: info Why verify hashes?
Matching the published hash confirms the file has not been altered, corrupted, or tampered with.  
Even a one‑byte change produces a completely different hash.
:::

### Unblock internet files with caution

When you download an executable or ZIP file on Windows, the system may mark it as coming from the internet and block it from running. This is a safety measure: **Windows treats downloaded files as potentially unsafe** until you explicitly confirm you trust them. Macro‑enabled Excel workbooks fall into this category as well.

::: tip How to unblock a file
1. Right‑click the downloaded file.
2. Select **Properties**.
3. In the **General** tab, look for an **Unblock** checkbox near the bottom (it appears only if the file is blocked).
4. Check **Unblock** to remove the restriction.
5. Click **OK** or **Apply**.

If you don’t see the **Unblock** option, the file may already be unblocked, or other security layers (administrator permissions, antivirus, or SmartScreen) may be preventing it from running.
:::

::: warning
Unblocking a ZIP file does **not** unblock the files inside it.  
You may need to unblock the extracted workbook separately.
:::

After extracting the ZIP, you may need to repeat the **Unblock** process for the extracted workbook or any executable files it contains.

::: info Why Windows preserves the “Mark of the Web” inside ZIP archives

Windows attaches a *Mark of the Web* (MOTW) to downloaded files to indicate they originated from an untrusted zone (e.g., the Internet). When a ZIP file carries this mark:

- Windows propagates the MOTW to **every file extracted** from the ZIP.
- Each extracted file is treated as if it were downloaded individually from the Internet.
- Excel uses the MOTW to enforce **stricter macro security**, including:
  - Blocking macros by default,
  - Showing security warnings,
  - Requiring explicit user trust before enabling code.

This behavior is intentional: it prevents malicious ZIP archives from bypassing security simply by packaging harmful files inside a container. As a result, even after unblocking the ZIP itself, you may still need to unblock the extracted workbook before Excel will allow macros to run.
:::

## Excel Security Settings

Ensure Excel’s macro settings are configured to **disable macros by default** with a prompt to enable them (Trust Center → Macro Settings). This prevents macros from running automatically.

Avoid enabling macros unless you fully trust the file and its source. If Excel prompts you to enable macros, proceed thoughtfully.

Excel also supports **Trusted Locations** which are designated folders where macro‑enabled spreadsheets can open without prompting. This lets you work with your own tools and automation safely, *without* loosening security for all files. Only add folders you control (for example, a local development directory), and avoid marking broad system folders or shared network paths as trusted.

::: tip How to configure a Trusted Location
1. Open **File → Options → Trust Center**.  
2. Click **Trust Center Settings…**.  
3. Select **Trusted Locations**.  
4. Click **Add new location…**.  
5. Choose a folder you control (e.g., a dedicated “ExcelTools” directory).  
6. Confirm and close the dialog.

Once added, any macro‑enabled workbook stored in that folder will open without security prompts while all other files remain protected.
:::

## Test in an Isolated Environment

Open the workbook in a **sandboxed or isolated environment**, such as a virtual machine (VM) or a dedicated device not connected to sensitive networks or data. This limits potential damage if the macro is malicious.

Alternatively, use a **cloud-based or disposable environment** (e.g., Windows Sandbox) to test the file.

## Review Purpose and Necessity

Evaluate whether the workbook’s functionality **requires macros**. If the workbook can be used without enabling macros, it’s safer to leave them disabled.

- The *Excel to Graphviz Relationship Visualizer* workbook is fully dependent on enabled macros.
- If you cannot enable macros, the software cannot operate.

## Review Documentation

Check whether the repository provides **clear documentation** explaining the macro’s purpose and functionality. A lack of transparent, well‑structured documentation is a red flag.

- Comprehensive end‑user documentation is published at [https://exceltographviz.com](https://exceltographviz.com).
- The *Excel to Graphviz Relationship Visualizer* source code is fully documented, though the codebase contains thousands of lines of VBA and the comments are primarily written for maintenance and support.
- In addition, the project now includes DeepWiki-generated documentation at [https://deepwiki.com/jjlong150/ExcelToGraphviz](https://deepwiki.com/jjlong150/ExcelToGraphviz), offering structured, cross‑linked explanations of modules, workflows, and internal architecture to help users and developers understand how the system operates.

## Check Community Feedback

Look for **user feedback** in the project’s issue trackers, discussions, and social channels. Reports of suspicious behavior, unexpected macro activity, or security concerns should always be taken seriously.

- **SourceForge** → [Support Tickets](https://sourceforge.net/p/relationship-visualizer/tickets/)
- **GitHub** → [Issues](https://github.com/jjlong150/ExcelToGraphviz/issues)
- **X / Twitter** → [@ExcelToGraphviz](https://x.com/exceltographviz)
- **External sources** → Forums, blogs, and community discussions such as the [Graphviz Forum](https://forum.graphviz.org/) may also provide independent feedback or warnings.

## Is it Being Maintained?

Confirm that the repository is **actively maintained**. Abandoned projects or outdated files may contain unpatched vulnerabilities or rely on deprecated dependencies.

- The *Excel to Graphviz Relationship Visualizer* has published:
  - Downloadable runtime files on [SourceForge](https://sourceforge.net/projects/relationship-visualizer/) continuously since **October 17, 2015**.
  - Exported VBA code and website markdown content on [GitHub](https://github.com/jjlong150/ExcelToGraphviz) since **February 11, 2022**.

- Look for recent commits, releases, or updates that address bugs, compatibility issues, or security concerns. Active issue responses and ongoing documentation updates (including DeepWiki‑generated content) are also strong indicators of healthy maintenance.

## Assess Alternative Options

Consider whether a **non–macro‑enabled version** or an alternative tool can meet your needs without the risks associated with running macros.

When possible, explore trusted, well‑established libraries or tools rather than relying on obscure or unverified downloads from SourceForge, GitHub, or other hosting sites. Established ecosystems often provide safer, audited, and actively maintained solutions.

## Practical Steps

**Download cautiously**: Only download from the project’s official SourceForge or GitHub pages — avoid third‑party mirrors or unverified sources.

**Back up data**: Ensure important files are backed up before opening any macro‑enabled workbook to minimize risk in case of corruption or unexpected behavior.

**Limit permissions**: Run Excel with standard user permissions rather than administrator rights to reduce the potential impact of malicious code.

## Detect Red Flags

Treat any of the following as a warning sign and proceed with extreme caution:

- Lack of documentation or unclear macro purpose.
- Requests for unusual permissions or unexpected external connections.
- Poorly rated, unverified, or inactive repositories.
- Antivirus warnings or negative community reports.

## Conclusion

By combining these practices - verifying the source, reviewing documentation, inspecting code, using secure environments, and staying cautious - you can make an informed decision about trusting a macro‑enabled Excel workbook from SourceForge or GitHub. When in doubt, seek guidance from a cybersecurity professional or avoid enabling macros altogether.
