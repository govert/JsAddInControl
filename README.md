# Testing Excel Hybrid Add-in Sample

This directory contains a sample Excel desktop test environment that combines two different add-in technologies in one solution:

- an Excel JavaScript add-in, responsible for a web-based custom task pane,
- an Excel-DNA add-in, responsible for native ribbon callbacks and UI Automation-based control of the JavaScript add-in UI.

The sample is designed to answer one specific question:

Can an Excel-DNA add-in running inside desktop Excel drive an Office JS custom task pane indirectly, even though the two add-ins are separate and do not share a direct runtime API?

The answer in this sample is "yes, via UI Automation", with the usual caveat that UI Automation is inherently more fragile than a direct supported integration path.

## What this sample demonstrates

When the sample is loaded successfully into Excel, you get:

- a `Testing Test` ribbon group from the Office JS add-in,
- a `Show Smiley Pane` button from that Office JS add-in,
- a custom task pane titled `Testing JS Test Add-in`,
- a large smiling face emoji rendered inside the task pane,
- a `Testing DNA` ribbon tab from the Excel-DNA add-in,
- a `Show JS Pane` button that uses UI Automation to invoke the Office JS ribbon button,
- a `Hide JS Pane` button that uses UI Automation to click the task pane close button,
- a `Dump UIA Tree` button that writes a filtered UI Automation snapshot to a text file and a new workbook.

This makes the sample useful both as:

- a runnable demo,
- a reference for adapting the UI Automation selectors to a real Excel add-in.

## High-level architecture

The two add-ins are independent:

- the Office JS add-in runs as a standard web add-in hosted from `https://localhost:3000`,
- the Excel-DNA add-in runs as a native `.xll` add-in targeting .NET Framework 4.8.

There is no direct call path between them.

Instead, the Excel-DNA add-in controls the Office JS add-in by interacting with Excel's user interface through Windows UI Automation:

1. The Excel-DNA ribbon callback runs in-process in Excel.
2. It locates Excel's main window handle through the Excel COM object model.
3. It creates a UI Automation root from that window.
4. It finds either:
   - the Office JS ribbon button, or
   - the Office JS task pane close button.
5. It invokes that UI element through `InvokePattern`.

This is deliberate. In many real-world hybrid add-in designs, you do not have a supported bridge from Excel-DNA into Office JS task pane lifetime management. UI Automation is one pragmatic fallback.

## Repository layout

### Parent directory

The parent directory contains:

- [README.md](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\README.md): this document.
- [Start-TestingExcelAddIns.ps1](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\Start-TestingExcelAddIns.ps1): convenience launcher intended to start the local web server, build/load the Excel-DNA add-in, and open both in one Excel session.

### Office JS project

The Office JS add-in lives under [OfficeJsAddIn](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn).

Important files:

- [manifest.xml](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\manifest.xml)
- [server.js](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\server.js)
- [taskpane.html](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\src\taskpane.html)
- [commands.html](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\src\commands.html)

### Excel-DNA project

The Excel-DNA add-in lives under [ExcelDnaAddIn](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn).

Important files:

- [ExcelDnaAddIn.csproj](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\ExcelDnaAddIn.csproj)
- [RibbonController.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\RibbonController.cs)
- [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs)

## Understanding the Office JS add-in

The Office JS add-in is a standard task pane add-in for Excel desktop.

Its manifest currently declares:

- display name: `Testing JS Test Add-in`
- provider name: `Local Test Environment`
- source location: `https://localhost:3000/src/taskpane.html`
- ribbon group label: `Testing Test`
- ribbon button label: `Show Smiley Pane`

The key implementation choice is in [manifest.xml](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\manifest.xml):

- the ribbon button uses the built-in `ShowTaskpane` action,
- not a custom shared-runtime function.

That matters because:

- `ShowTaskpane` is reliable in this sample,
- attempts to implement a true one-button show/hide toggle were unstable in this environment,
- the current sample therefore treats the Office JS command as "show only",
- the pane is closed either by Excel's `X` button or by the Excel-DNA add-in's UI Automation logic.

The visual content of the pane is intentionally simple. It exists to make task pane lifetime easy to validate:

- if the pane appears, hosting works,
- if the smiley is visible, the page loaded correctly,
- if the pane content is blocked, the certificate/server path is still wrong.

## Understanding the Excel-DNA add-in

The Excel-DNA add-in contributes a custom ribbon tab called `Testing DNA`.

That ribbon is defined in [RibbonController.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\RibbonController.cs). The main buttons are:

- `Show JS Pane`
- `Hide JS Pane`
- `Dump UIA Tree`

The heavy lifting is in [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs).

That file contains:

- the UI Automation selectors,
- the logic for locating Excel's UIA root,
- the search logic for ribbon controls and task pane controls,
- the UI tree dump logic,
- the threading boundary for deferred main-thread work.

### What `Show JS Pane` does

`Show JS Pane` does not show the pane directly through an API. Instead, it:

1. Gets the Excel main window handle from `ExcelDnaUtil.Application`.
2. Builds a UI Automation root with `AutomationElement.FromHandle`.
3. Selects the Excel `Home` tab.
4. Finds the Office JS button named `Show Smiley Pane`.
5. Invokes that control through UI Automation.
6. Queues a best-effort switch back to the `Testing DNA` tab using `ExcelAsyncUtil.QueueAsMacro`.

The queued follow-up matters because COM and UIA should not be driven from a background worker thread in this sample.

### What `Hide JS Pane` does

`Hide JS Pane` looks for the Office JS task pane chrome and tries to invoke the close button.

The relevant selectors are intentionally near the top of [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs), so they are easy to retune:

- `JsShowRibbonButtonName`
- `TaskPaneCloseButtonName`
- `TaskPaneOptionsButtonName`
- `TaskPaneWindowClassName`
- `TaskPaneCloseButtonClassName`
- `TaskPaneTitleCandidates`
- `TaskPaneAnchorTexts`

This is the main file to edit when adapting the sample to another ribbon layout, title string, Office channel, or localization.

### What `Dump UIA Tree` does

`Dump UIA Tree` is a diagnostic tool for retuning UI Automation selectors.

It currently:

- traverses a filtered set of UI Automation nodes,
- writes a text dump into `%TEMP%`,
- creates a new workbook,
- writes the same filtered node list into a worksheet.

The filtering tries to keep the dump focused on:

- ribbon controls,
- NetUI/Mso task pane chrome,
- controls likely to matter for show/hide automation.

It intentionally avoids most worksheet and grid noise.

## Threading model

This sample should not access the Excel COM object model or UI Automation from worker threads.

Current rule:

- ribbon callbacks run on Excel's main thread,
- deferred follow-up work is returned to Excel's main thread via `ExcelAsyncUtil.QueueAsMacro`,
- the sample does not intentionally access COM/UIA from a background thread.

This is important because Excel COM and its UI surface are timing-sensitive, STA-oriented, and prone to subtle failure when touched from the wrong thread.

If you extend this sample, keep that rule.

## Prerequisites

To run the sample, you need:

- Windows,
- desktop Microsoft Excel,
- Node.js and npm,
- a .NET SDK with .NET Framework 4.8 targeting support,
- PowerShell 7 available as `pwsh`.

## Running the sample

There are two ways to run it:

- the convenience launcher from the parent directory,
- manual startup of each add-in.

### Recommended: use the parent launcher

From the parent directory, run:

```powershell
pwsh -File .\Start-TestingExcelAddIns.ps1
```

This script is intended to:

- install Office JS dependencies if needed,
- export and trust a localhost development certificate,
- register the Office JS manifest for sideloading,
- stop anything already listening on port `3000`,
- start the local Office JS HTTPS server,
- build the Excel-DNA project when needed,
- attach to an existing Excel instance or create one,
- register the Excel-DNA `.xll`,
- open the Office JS sideload workbook in that same Excel instance.

If you want to force an Excel-DNA rebuild:

```powershell
pwsh -File .\Start-TestingExcelAddIns.ps1 -RebuildExcelDna
```

### Manual Office JS startup

From [OfficeJsAddIn](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn):

```powershell
npm install
npm start
```

Then, from another shell in the same directory:

```powershell
npm run sideload
```

The local content should be served from:

```text
https://localhost:3000
```

### Manual Excel-DNA startup

From [ExcelDnaAddIn](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn):

```powershell
dotnet restore
dotnet build
```

The Excel-DNA package now generates the `.dna` files during build output, rather than requiring a root project `.dna` file.

At the time of writing, the build output includes names like:

- `bin\Debug\net48\ExcelDnaAddIn-AddIn.xll`
- `bin\Debug\net48\ExcelDnaAddIn-AddIn64.xll`
- `bin\Debug\net48\publish\ExcelDnaAddIn-AddIn-packed.xll`
- `bin\Debug\net48\publish\ExcelDnaAddIn-AddIn64-packed.xll`

Load the correct `.xll` for your Excel bitness through:

- `File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...`

## Typical test flow

Once both add-ins are loaded into the same Excel instance, test in this order:

1. Click `Testing Test` -> `Show Smiley Pane`.
2. Confirm the Office JS task pane appears.
3. Confirm the pane shows the smiling face content.
4. Close the pane with the task pane `X`.
5. Click `Testing DNA` -> `Show JS Pane`.
6. Confirm the pane appears again, this time because the Excel-DNA add-in invoked the Office JS ribbon control through UI Automation.
7. Click `Testing DNA` -> `Hide JS Pane`.
8. Confirm the Excel-DNA add-in can close the pane through UI Automation.
9. If any selector fails, click `Testing DNA` -> `Dump UIA Tree` and inspect the dump outputs.

## How to adapt this sample

If you are using this sample as a starting point for a real solution, the most likely places you will edit are:

### Office JS changes

Edit [manifest.xml](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\manifest.xml) when you need to change:

- display name,
- ribbon group name,
- ribbon button name,
- task pane source page,
- URLs or assets.

Edit [taskpane.html](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\OfficeJsAddIn\src\taskpane.html) when you need to change:

- the pane content,
- styling,
- branding,
- any client-side UI for the web pane.

### Excel-DNA automation changes

Edit [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs) when you need to change:

- ribbon button names,
- task pane title names,
- UI Automation class names,
- search depth,
- dump filtering,
- matching heuristics.

Edit [RibbonController.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\RibbonController.cs) when you need to change:

- the Excel-DNA tab layout,
- button labels,
- button callbacks,
- success/failure messaging.

## Troubleshooting

### The Office JS pane opens but shows blocked content or certificate errors

Likely cause:

- the localhost HTTPS certificate chain is not trusted correctly.

What to do:

- rerun [Start-TestingExcelAddIns.ps1](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\Start-TestingExcelAddIns.ps1),
- confirm `https://localhost:3000` is reachable,
- confirm Excel is not still holding onto an older broken registration.

### Two Excel windows or processes appear

Likely cause:

- the add-ins were launched independently instead of being attached to the same Excel instance.

What to do:

- close Excel fully,
- start again with [Start-TestingExcelAddIns.ps1](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\Start-TestingExcelAddIns.ps1).

### `Show JS Pane` cannot find the Office JS button

Likely causes:

- the ribbon label changed,
- the UI Automation tree differs on this Excel build,
- the Office JS command is not currently loaded on the expected tab.

What to do:

- run `Dump UIA Tree`,
- inspect the filtered output,
- adjust the constants and matching rules in [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs).

### `Hide JS Pane` cannot find the close button

Likely causes:

- the task pane title string changed,
- the title bar control hierarchy differs,
- the class names differ on this machine or Office build.

What to do:

- leave the pane visible,
- run `Dump UIA Tree`,
- update the task pane selectors in [OfficeJsPaneAutomation.cs](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\ExcelDnaAddIn\OfficeJsPaneAutomation.cs).

### The Office JS pane does not load even though it opens

Likely causes:

- local HTTPS server is not running,
- certificate trust is incomplete,
- Excel opened a stale sideload registration.

What to do:

- verify `https://localhost:3000/src/taskpane.html` responds,
- restart the local server,
- reopen Excel through the launcher.

### The Excel-DNA launcher path does not match the generated `.xll` names

Important note:

- the Excel-DNA build currently generates `*-AddIn.xll` style names,
- but any scripts or tooling that still assume the older `ExcelDnaAddIn64.xll` naming will need to be updated.

If the launcher stops finding the add-in output, fix the expected output path in [Start-TestingExcelAddIns.ps1](C:\Work\ExcelDna\TechnicalSupport\Millennium\JsAddInControl\Start-TestingExcelAddIns.ps1).

## Known limitations

- The Office JS ribbon button is currently "show" only, not a robust toggle.
- UI Automation is inherently brittle compared with a direct supported API.
- Control names, class names, or tree shapes may vary by Excel version, Office channel, localization, DPI, or layout.
- The dump tool is diagnostic, not authoritative.
- The sample is optimized for clarity and testability, not for production-hardening.

## Why this sample is structured this way

This sample deliberately favors explicitness over abstraction.

You can see:

- exactly which ribbon labels are being targeted,
- exactly which task pane controls are being matched,
- exactly where the UI Automation traversal starts,
- exactly where to edit the selectors when adapting the sample.

That makes it a better diagnostic and learning project than a heavily abstracted one.

## Suggested next improvements

If you continue evolving this sample, the most useful next steps are:

- update the launcher to follow the current generated Excel-DNA output names,
- improve `Hide JS Pane` by anchoring more precisely from the title area,
- add supported-pattern details and parent-chain context to the dump output,
- make the UI Automation selector set configurable rather than hard-coded,
- add an integration note describing expected differences across Office builds.
