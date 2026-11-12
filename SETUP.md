# Smart Control Center — Windows Setup Guide

Two ways to run this. Pick one.

- **Embedded** — the Control Center lives inside one specific workbook (e.g. your
  GFD reporting file). Simple, no extra install, tab only shows in that file.
- **Add-in** — a small `.xlam` loads at Excel startup and puts a permanent
  "Smart Control Center" tab on the Ribbon in every workbook. One click deploys
  the dashboard into whichever file you're working in.

Both need this one-time setting, since installing/updating modules is done by
code:
**File > Options > Trust Center > Trust Center Settings > Macro Settings >
check "Trust access to the VBA project object model".**

---

## Option A — Embedded in one workbook

1. Open your workbook, save it as **.xlsm** if it isn't already.
2. Alt+F11 to open the VBA editor. File > Import File, and import these 20
   files (from the `SmartControlCenter` folder): `modUtilities`, `modSettings`,
   `modAuditLog`, `modBackup`, `modDataSources`, `modKPI`, `modCharts`,
   `modSearch`, `modDataQuality`, `modHighlighter`, `modDataFilter`,
   `modDataManager`, `modDataExporter`, `modReportGenerator`,
   `modEmailAutomation`, `modCodeAutomation`, `modJupyter`,
   `modFileAutomation`, `modWorkbookConverter`, `modMain`.
3. Also import `AddIn/modRibbon.bas` — it works standalone too.
4. Double-click **ThisWorkbook** in the Project pane, paste in the contents
   of `ThisWorkbook_AutoOpen.txt`, save.
5. Run `modMain.InitializeSmartControlCenter` once from the VBE (F5) to build
   the sheet immediately, or just close and reopen the file — `Workbook_Open`
   does it for you from now on.
6. Add the Ribbon tab — close the file in Excel first, then use the **easiest
   option**: open the `AddIn` folder, drag your `.xlsm` file onto
   `Install-RibbonTab.bat` and drop it. A black window walks you through it
   and tells you SUCCESS or exactly what went wrong. No PowerShell knowledge
   needed — the .bat handles that for you.
7. Reopen the workbook — you'll see a **Smart Control Center** tab on the Ribbon.

## Option B — Global Add-in with its own tab

1. Open a new blank workbook, save as **.xlam** (File > Save As > Excel
   Add-in) into the default Add-ins folder Excel suggests
   (`%APPDATA%\Microsoft\AddIns`), e.g. `SmartControlCenter.xlam`.
2. Alt+F11, import the same 20 core modules **plus** `AddIn/modInstaller.bas`
   and `AddIn/modRibbon.bas` (22 total).
3. Close and reopen the add-in file once in the VBE to save it in .xlam form
   with the code included (Ctrl+S).
4. Inject the ribbon: close the .xlam, then drag it onto
   `AddIn\Install-RibbonTab.bat` and drop it (same tool as Option A).
5. In Excel: **File > Options > Add-ins > Manage: Excel Add-ins > Go... >
   Browse...**, select `SmartControlCenter.xlam`, check its box, OK.
6. The **Smart Control Center** tab now appears every time Excel starts, in
   every workbook. Open (or switch to) any data workbook and click
   **Install / Update in This Workbook** — that copies the modules in and
   builds the dashboard there. Every other button (Refresh, Data Quality,
   Highlighter, Search, Settings, Backup, Help) auto-installs first if needed,
   then runs against whichever workbook is active.
7. To update the add-in later: edit the modules in the `.xlam`'s own VBE, then
   click **Install / Update in This Workbook** in any file to push the new
   version into it — no reinstalling the add-in itself.

---

## Upgrading from Embedded to a Global Add-in

Already have the tab working inside one `.xlsm` (Option A) and want it to show
in *every* workbook instead? Convert that same file rather than starting over:

1. Open your working `.xlsm`. Alt+F11 → double-click **ThisWorkbook** → replace
   its code with the updated `ThisWorkbook_AutoOpen.txt` (it now skips itself
   when running as an add-in, so nothing tries to build a dashboard inside the
   hidden add-in file).
2. File > Import File → import `AddIn/modInstaller.bas` (you already have
   `modRibbon` from setting up the tab).
3. File > Save As → file type **Excel Add-in (*.xlam)** → Excel will offer its
   default Add-ins folder — save there, e.g. `SmartControlCenter.xlam`.
   The ribbon tab you already injected carries over automatically; if it
   doesn't show when you reopen it, drag the new `.xlam` onto
   `Install-RibbonTab.bat` once.
4. **File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...**,
   select `SmartControlCenter.xlam`, check its box, OK.
5. Restart Excel. The **Smart Control Center** tab now shows for every
   workbook you open. On a new workbook, click **Install / Update in This
   Workbook** to deploy the dashboard into it.

If you keep using the original `.xlsm` too, its own embedded copy still shows
its own tab whenever that file is open — with the add-in also loaded, you may
briefly see two "Smart Control Center" tabs while that one file is active.
Harmless, just a little confusing; ignore it or delete that file's ribbon copy
if it bothers you (open it with `Install-RibbonTab.bat`'s companion — the
Custom UI Editor — and clear the customUI14.xml part).


## Troubleshooting the Ribbon Tab

Run `Install-RibbonTab.bat` (drag your file onto it) and read what it prints —
it now verifies its own work and tells you plainly SUCCESS or FAILED with a
reason. Common causes if it still doesn't show up:

- **The file was open in Excel (or OneDrive was syncing it).** Close Excel
  completely — and pause OneDrive sync if the file lives in a synced folder —
  then run the installer again.
- **"We found a problem with some content"** appears when you open the file
  afterwards. This is normal the first time — click **Yes**/**Repair** once;
  Excel discards only the part it didn't like and keeps your workbook, sheets
  and VBA project untouched. If the ribbon still doesn't appear after that,
  the XML had an error — re-run the installer (it rewrites the ribbon part
  fresh each time, which fixes most cases) and try again.
- **The file was downloaded from the internet / a shared drive.** Windows
  marks it "blocked": right-click the file > Properties > check **Unblock**
  > OK, then reopen.
- **You installed into a .xlam but the tab still isn't there.** Confirm it's
  actually loaded: File > Options > Add-ins > Manage **Excel Add-ins** > Go —
  your add-in must be **checked**, not just listed.
- **PowerShell scripts are blocked entirely on this PC.** The `.bat` already
  runs PowerShell with `-ExecutionPolicy Bypass` for just this one script, so
  it works even if your normal PowerShell policy is Restricted. If it still
  refuses to run, your organization's Group Policy may block PowerShell
  outright — in that case use the free **Custom UI Editor for Office** tool
  instead (search that exact name; it's a small, well-known download): open
  your file with it, paste in the contents of `customUI14.xml`, validate
  (green checkmark), save, reopen in Excel.
- Each run makes a `yourfile.xlsm.bak` copy first — if anything looks wrong,
  restore from that and try again.

## Notes

- Backups (Backup Now / before destructive operations) go to a `Backups`
  subfolder next to the workbook, so the workbook must be saved to disk first.
- If Excel shows a security warning because the file came from the internet,
  right-click it in File Explorer > Properties > **Unblock**, then reopen.
- Everything uses late binding — no extra VBA references to set.

## What each module needs on the PC

| Needs | Modules |
|---|---|
| Nothing extra (Excel only) | Dashboard, Data Manager, Filter, Highlighter, Search, Data Quality, Report Generator, Settings, Audit Log, Workbook Converter |
| Windows Scripting Runtime (`Scripting.FileSystemObject`, `Scripting.Dictionary`) — built into Windows, no install needed | Highlighter, Data Quality, Search, KPI, Report Generator, Jupyter Generator, File Tools |
| Microsoft Outlook installed and signed in | Email Automation |
| ADODB (built into Windows) | Jupyter Generator (for clean UTF-8 file output) |
| Python + Jupyter/JupyterLab/VS Code — only to *open* the generated notebook, not to generate it | Jupyter Generator's output file |
| PDF export uses Excel's own `ExportAsFixedFormat` | Data Exporter, Workbook Converter, Report Generator, Email Automation |
