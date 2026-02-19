# BMS Report System — HTML Output Project

## Goal
Add HTML output option to the existing xBase/Clipper report engine so that any report in the BMS system can be exported as an HTML file that looks identical to the current printed/file output.

## Current Report Architecture

### How reports work today

Every report in the system follows the same pattern:

1. **Report PRG** (e.g. `RPQC01V1.PRG`) — defines criteria, sort options, fields, query blocks
2. **TheReport class** (`THEREPO.PRG`) — orchestrates the entire flow via `Exec()` method
3. **PrnFace** (`PRNFACE.PRG`) — UI for user to pick sort order, criteria, and **output destination**
4. **RH2 template** (e.g. `RPQC01V1.RH2`) — binary report layout: headers, body, footers, grouping, subtotals
5. **rpGenReport()** — the report engine reads the RH2, iterates the temp DBF, evaluates field expressions, and calls `rpPageOut()` for each rendered page
6. **rpPageOut()** (`PAGEOUT.PRG`) — receives `aPage[]` (array of ready-formatted text lines) and sends them to the chosen destination

### Current output destinations (set in `PRNFACE.PRG`)

| Destination | Constant | What happens |
|-------------|----------|-------------|
| Printer | `rpPRINTER (1)` | `rpLPrint()` sends lines to LPT port |
| File | `rpPRINTER_FILE (2)` | `fwrite()`/`fwriteline()` to a text file |
| Screen | `rpDISPLAY (3)` | Preview on screen |
| SpreadSheet | — | `Lexport()` + CSV via `COPY TO ... DELIMITED` |
| V7 Export | — | `COPY TO ... SDF` |
| Query | — | Interactive `BroCenter` browse |

### Key insight
`rpPageOut()` in `PAGEOUT.PRG` is the single chokepoint for all printed/file output. The `aPage[]` array it receives contains **fully formatted text lines** — headers, column titles, data rows with `|` separators, `+---+` borders, subtotals, everything. The RH2 engine has already done all the grouping, aggregation, and formatting.

## Data flow

```
Report PRG (RPQC01V1.PRG)
  → TheReport:Exec()                    [THEREPO.PRG]
    → prnFace()                         [PRNFACE.PRG]  — user picks criteria & destination
    → TheReport:CreateCondDb()          — filters D_LINE → temp t_line01.dbf
    → TheReport:IndexTempFile()         — indexes by chosen sort expression
    → TheReport:Print()                 — loads RH2, sets destination
      → rpQuickLoad( cRepFileName )     — loads RH2 template
      → rpDestination( oRP, nDest )     — sets output target
      → rpGenReport( oRP )              — engine iterates data, renders pages
        → rpPageOut( oRP, aPage )       [PAGEOUT.PRG] — OUTPUT HAPPENS HERE
          → rpLPrint (printer) or fwrite (file)
```

## The Plan: Add HTML destination

### Approach
Add `rpHTML (4)` as a new destination type. Intercept at `rpPageOut()` — the same place where Printer and File output happen. Since `aPage[]` already contains fully formatted monospace text, wrapping it in `<pre>` tags produces visually identical output with zero risk of breaking existing reports.

### Files to modify

| File | Change | Lines |
|------|--------|-------|
| `PRNFACE.PRG` | Add "HTML" to `aDevices` array + `DefineDest` handler | ~5 |
| `THEREPO.PRG` | Add `rpHTML` to `GetDestin()` + setup in `Print()` | ~10 |
| `PAGEOUT.PRG` | Add `CASE rpHTML` block that calls `HtmlPageOut()` | ~15 |
| **`HTMLOUT.PRG`** (new) | HTML open/write/close functions | ~50 |

### What NOT to change
- Individual report PRGs (RPQC01V1.PRG etc.) — untouched
- RH2 files — untouched
- Report engine (rpGenReport) — untouched
- Existing destinations (Printer, File, Screen) — untouched

## Source files in this repo

### Report engine core
| File | Purpose |
|------|---------|
| `THEREPO.PRG` | `TheReport` class — Exec(), Print(), SpreadSheet(), V7Export(), Query(), CreateCondDb(), GetDestin() |
| `PAGEOUT.PRG` | `rpPageOut()` — the output dispatcher. **Main interception point for HTML** |
| `PRNFACE.PRG` | UI for criteria selection and destination choice |
| `PRNCRIT.PRG` | Criteria input functions (critBrowse, GetFinishDate, etc.) |
| `REPHF.PRG` | `RepHF` class — report header/footer builder |
| `REPSTUFF.PRG` | Report preprocessing helpers |
| `PRNPRG.PRG` | Printer escape code definitions (HP, Epson, Kyocera) |
| `PRNOUT.PRG` | Print queue management |
| `PRINTERS.PRG` | Printer database management (printers.dbf) |

### Example report
| File | Purpose |
|------|---------|
| `RPQC01V1.PRG` | "Production line report" — sample report PRG showing the standard pattern |

### RH2 templates (binary)
| File | Used when |
|------|----------|
| `RPQC01V1.RH2` | Classic report, most sort options |
| `RPQC01V2.RH2` | Sort by Workstation with form-feed between each |
| `RPQC01V3.RH2` | TopDown sort, grouped by value with subtotals |
| `RPQC01V4.RH2` | New report variant (added 2002), most sort options |
| `RPQC01V5.RH2` | New report, sort by Project+Line |

## Technical notes

- Language: xBase (Clipper/Harbour dialect) with CA-Tools extensions
- Database: DBF/CDX (DBFCDXAX driver)
- Report engine: TVR (The Visual Reporter) — proprietary, functions like rpNew(), rpQuickLoad(), rpGenReport() are from its library
- Encoding: Windows-1255 (Hebrew) for some field values and comments
- All reports output to monospace 120-column format (landscape)
