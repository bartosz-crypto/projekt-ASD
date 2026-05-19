# AsdRcSlab — Project Handoff

## Overview

AutoCAD 2015 plugin (.NET Framework 4.8) for Speedeck Foundations.
Automates RC slab drawing workflow: punching report processing, pile
classification, drawing annotation, title block management, and reference
frames.

**Stack**: AutoCAD 2015 API, WPF (UI), EPPlus 4.5.3.3 (Excel),
Newtonsoft.Json 13. C# 7.x targeting .NET Framework 4.8.

## Loading the plugin

```
NETLOAD <path>\AsdRcSlab.dll
```

On unload before rebuild:
```
NETUNLOAD AsdRcSlab
```

## Commands

### Active in ribbon

Tab "ASD RC SLAB" has 2 panels:

**TITLE BLOCK** (2 buttons):
- `ASD-GAI` — "Copy from GA": opens GA DXF/DWG, copies A1-BL attributes
  (CLIENT/PROJ/APPROVED/TITLE_1 prefix/DRAWING_NUMBER), copies SLAB NOTES
  (AREA, PERIMETER, THICKNESS, VOLUME), maps SLAB THICKNESS to HYSTOOLS
  variant (225→DK90, 300→DK165).
- `ASD-RCN` — "Sheet Numbering": auto-detects TITLE_3 from Model space
  + viewports, auto-builds SCALE from viewport scales, sets DATE to current
  month + year, renames layouts to `<DRAWING_NUMBER suffix>C1`, creates
  reference frames "SEE DRG. ..." on layer ASD-RCN-REFS pointing to all
  other layouts in the document.

**PH CONDITIONS** (4 buttons in 2x2 grid):
- `ASD-PXIE` — "Load Punching": opens Excel Punching Report, picks plot
  (supports ranges like PLOT 4-5), loads pile data into SessionData.Piles.
- `ASD-PAA` — "Assign PH": runs PhAssigner classification on piles, then
  calls DrawingAnnotator to annotate detail circles + update PH templates
  in Model space. Opens PhAssignResultsDialog with editable PhAction
  ComboBox + "Update Drawing" button.
- `ASD-PHR` — "PH Report": opens PhAssignResultsDialog as read-only view
  (PhAction not editable, no Update button).
- `ASD-PHV` — "Validate PH": currently stub.

### Hidden / command-line only

These commands have `[CommandMethod]` attributes but no ribbon buttons.
Reachable from command line:

- `ASD-PROJ`, `ASD-OPEN`, `ASD-SET` — project lifecycle (stubs/partial)
- `ASD-GBOT`, `ASD-GTOP`, `ASD-BMM`, `ASD-LAP` — reinforcement (stubs/partial)
- `ASD-GSETUP` — setup command, no ribbon button (orphaned since handoff #1)
- `ASD-BBSV`, `ASD-PIV`, `ASD-GER`, `ASD-QAP` — QA validators (stubs)
- `ASD-BSX`, `ASD-PDF`, `ASD-CAG`, `ASD-TRX` — exports (stubs)

## Architecture

### Core source files

- `Commands.cs` — main plugin entry, all `[CommandMethod]` definitions,
  GA→RC copying logic (~1700 lines after cleanup).
- `PunchingParser.cs` — Excel parsing, multi-PLOT detection with ranges
  (PLOT N-M).
- `PlotInfo.cs` — plot data model (FirstPlotNumber, LastPlotNumber,
  PileCount, StartRow). `IsRange` property.
- `SessionData.cs` — static singleton with current state (Piles, CurrentPlot,
  PhAssigned flag).
- `PileData.cs` — single pile record (PileId, Util, Reinf, Location,
  PhAction).
- `PhAssigner.cs` — pile classification logic: `phNum = (level-1)*3 + locIdx + 1`
  with validation rules R77, R79, R27.
- `DrawingAnnotator.cs` — Model space annotation: hatches on detail
  circles (LayerPhHatch), PH labels (LayerPhText), template AP-TEXT
  updates, NOT USED crosses (LayerNotUsed = "AP-NOTUSED").
- `RibbonBuilder.cs` — WPF Ribbon UI builder.
- `App.cs` — `IExtensionApplication.Initialize()` → calls RibbonBuilder.

### WPF dialogs

- `PhAssignResultsDialog.xaml(.cs)` — pile results grid, two modes:
  edit (PAA) / read-only (PHR). Shows TOTAL formulas (H12/H16) in bold.
- `PlotPickerDialog.xaml(.cs)` — selects plot from Excel (multi-PLOT
  files including ranges).
- `NewProjectDialog.xaml(.cs)` — new project setup (partial).
- `SettingsDialog.xaml(.cs)` — plugin settings (partial).

## Key conventions

### Layer naming

- `AP-TEXT` — PH detail templates in Model space (MText). Modified by
  AnnotatePhDetailLabels.
- `AP rebar top` (LayerPhText) — labels on pile detail circles. Cleared
  by CleanupPreviousAnnotations before each Annotate.
- `AP-Hatch` (LayerPhHatch) — hatches on pile details. Same cleanup.
- `AP-NOTUSED` (LayerNotUsed) — diagonal crosses on unused PH templates
  (PHs with 0 piles). Same cleanup. Color ACI 1 (red) hardcoded on
  entity (ColorIndex=1) to override any layer color.
- `ASD-RCN-REFS` — reference frames "SEE DRG. ..." on each RC layout.
  Color ACI 7 (white) on layer; lines have explicit ColorIndex=7 to
  override. Cleared per-layout each ASD-RCN run.
- `PCN-Text` — SLAB NOTES in GA paperspace (read from).
- `SD-Text` — SLAB NOTES in RC paperspace (written to).

### Block + attributes

- `A1-BL` — title block. Attributes copied by GaiFieldsToCopy:
  - CLIENT_1, CLIENT_2, CLIENT_3
  - PROJ_1, PROJ_2, PROJ_3
  - APPROVED (added in p53)
  
  Plus separately handled:
  - TITLE_1 (prefix-replace via ReplaceRcTitlePrefix; if GA prefix empty,
    strips prefix from RC)
  - DRAWING_NUMBER (auto-numbered from first GA layout suffix via
    ExtractFirstGaDrawingNumber; e.g. GA100 → RC100, RC101, ...)
  - TITLE_3, SCALE, DATE — set by ASD-RCN, not ASD-GAI

### File naming

- Test Excel: `*-PLOT_*_report_punching.xlsx`
- Test GA: `<project_code>-DR-GA<numbers>__Rev_<rev>__-_PLOT_<n>.dxf`
- Test RC: `test-punching.dxf` (legacy from handoff #1)

## Known issues / tech debt

1. **HATCH renders as solid fill** on first run — needs manual scale
   tweak (e.g. 10 → 11 → 10) to render ANSI31 pattern. Investigated in
   p37-p41, rolled back in p42. Workaround: user manually adjusts after
   ASD-PAA. Root cause not found in AutoCAD 2015 .NET API; likely
   requires Circle→Polyline boundary conversion + LoopType experimentation
   in a future session. Accepted as MVP limitation.

2. **CONCRETE VOLUME loses tail** intentionally (p54) — only the numeric
   value + m/m³ is preserved in RC, any GA tail like "PILE CAP INC." or
   "INC. GROUND BEAM" is stripped. Decision based on user preference
   for deterministic output. To preserve tails, restore the
   `ConcreteVolumeReplaceRx` to use Group 4 capture again.

3. **DISTRIBUTION timing bug** (BUG-01 from handoff #1, HIGH priority, L size)
   — never touched in any of the sessions. Original WIP at commit
   f638ee9 before handoff. Still pending.

4. **Stub commands in Commands.cs** without ribbon buttons:
   ASD-PROJ, ASD-OPEN, ASD-SET, ASD-GBOT, ASD-GTOP, ASD-BMM, ASD-LAP,
   ASD-GSETUP, ASD-BBSV, ASD-PIV, ASD-GER, ASD-QAP, ASD-BSX, ASD-PDF,
   ASD-CAG, ASD-TRX, ASD-PHV. Either implement and re-add to ribbon, or
   remove from Commands.cs if obsolete.

5. **Excel binaries tracked in git** — bin/obj directories included since
   handoff #1, no `.gitignore` yet. Repo size growing unnecessarily.

6. **No README** in repo root.

7. **No `.bundle` package** for one-click deployment via
   `%APPDATA%\Autodesk\ApplicationPlugins\`. Currently requires manual
   NETLOAD on every AutoCAD session unless added to startup acaddoc.lsp.

## Tags / rollback points

7 immutable rollback points:

| Tag | Commit | Sprint summary |
|-----|--------|----------------|
| stabilization-1 | (p08) | PXIE + DrawingAnnotator + hatch scale |
| stabilization-2 | (p22) | ASD-GAI + ASD-RCN + ribbon cleanup |
| stabilization-3 | (p26) | Reference frames SEE DRG |
| stabilization-4 | (p47) | PAA/PHR split + edit + formulas + cross marking |
| stabilization-5 | (p54) | Plot ranges + RC numbering from GA + HYSTOOLS + APPROVED + VOLUME w/o tail |
| stabilization-6 | (p57) | VOLUME m bez superscript + RC empty (not "—") + strip prefix TITLE_1 |
| stabilization-7 | (p60+p61, this tag) | Debug logs cleanup + PL→EN refactor + handoff doc |

To rollback: `git reset --hard stabilization-<N>`.

## Common workflows

### Full RC drawing setup from scratch

1. Open RC template in AutoCAD.
2. `NETLOAD AsdRcSlab.dll`.
3. `ASD-PXIE` → select Excel punching report → pick plot.
4. `ASD-PAA` → classifies piles + annotates Model space + opens dialog.
5. (optional) Edit PhAction in dialog if automatic classification needs
   override. Click "Update Drawing" to re-annotate.
6. `ASD-GAI` → select GA DXF/DWG → copies attributes + SLAB notes.
7. `ASD-RCN` → auto-fills TITLE_3, SCALE, DATE, renames layouts to
   `<RC-NNN>C1`, creates reference frames.
8. Hatch on detail circles renders as solid fill — manually tweak scale
   via Properties to fix (known issue #1).

### Quick re-annotate after PhAction edits

1. `ASD-PAA` → dialog opens.
2. Edit PhAction in DataGrid (ComboBox).
3. "Update Drawing" → cleanup + re-annotation. Stats and TOTAL formulas
   refresh live.

## Development notes

- `NETUNLOAD AsdRcSlab` before every rebuild, AutoCAD locks the DLL.
- Build output: `bin\Build_verify\AsdRcSlab.dll`. Copy to `bin\Publish\`
  manually if AutoCAD blocks direct rebuild to Publish.
- Polish comments in code are intentional (developer-facing, not user).
  All user-facing strings are EN as of p60.
- `claude-code` sessions used through prompt files (this repo's
  `outputs/p*.txt` series, plus the conversational meta-prompting on
  the Claude side that generated them).
