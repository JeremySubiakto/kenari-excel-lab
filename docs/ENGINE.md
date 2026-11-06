# The formula engine

**Chosen: Formualizer 0.13.0**, MIT or Apache-2.0. Its two files ship with the site in `assets/engine/`, so the app does not depend on a CDN.

## How it was tested

`tests/engine-battery.json` holds 76 formulas on one small sheet, plus 2 spill behaviors. The formulas ran on three engines and on real Microsoft Excel 16.113 for Mac (Microsoft 365). The spill behaviors ran on the engines only. `tests/excel-results.json` is what Excel returned. `scripts/dev/excel-crosscheck.py` reruns the battery in real Excel.

| Engine | License | Compressed download | Matched real Excel, of 76 | XLOOKUP, FILTER, UNIQUE, SORT, LET, SUMIFS, spilling |
| --- | --- | --- | --- | --- |
| **Formualizer 0.13.0** | MIT or Apache-2.0 | about 2.7 MB | 69 | All pass |
| IronCalc 0.9.1 | MIT or Apache-2.0 | about 0.7 MB | 68 | XLOOKUP cannot return several columns, or take an array constant. TEXTJOIN with an array is missing |
| HyperFormula 3.5.0 | GPLv3, or paid | about 0.2 MB | 47 | No LET. SORTBY and XMATCH came back #NAME? as well |

Two misses for every engine are the regional-format difference described below, so they are not engine errors. Counts are close between the top two, so the choice rests on the required functions in the last column.

HyperFormula's VLOOKUP also came back `#NAME?` in this setup. That may be how I loaded it. I did not chase it, because the license and the missing LET already ruled it out.

## What Formualizer gets wrong

Each one is a listed known gap in the test, so a new version that fixes it will fail the test and get noticed.

1. **No LAMBDA, MAP or BYROW.** `=LAMBDA(x,x*2)(21)` returns `#N/IMPL!`, and MAP and BYROW return `#NAME?`. Missions on these need real Excel with upload checking, or another engine.
2. **No GROUPBY.** No engine tested has it. Teach it in real Excel.
3. **Exact number comparison.** `=(0.1+0.2)=0.3` is FALSE here and TRUE in Excel, which compares at about 15 digits. The answer checker rounds, so it is safe. A mission that teaches this quirk must not rely on the engine.
4. **DATE returns a date object**, and `core/engine.js` converts it back to a date number.
5. **Size.** 2.7 MB compressed, 10 MB once unpacked. First load took about 3 seconds with a simulated slow phone (6x slower processor, 8 Mbps). After that the browser cache should make it faster.
6. **Young project.** Version 0.x, with no outside record yet. Pin the version and keep the battery running.

## What real Excel showed about Indonesian settings

The Mac's Excel uses `;` between arguments and `,` as the decimal point. So:

- `=SUMIFS(A;B;C)` is what a learner on Indonesian settings types. The checker accepts both `,` and `;`.
- `=VSTACK(1,2)` on those settings means the single number 1.2.
- **Format codes follow the settings too.** `=TEXT(1234567.5,"#,##0")` returned `1234567,50` and `=TEXT(0.256,"0.0%")` returned `26%` in this Excel. The engine returns `1,234,568` and `25.6%`, which is what Excel on US settings gives. A lesson on TEXT must warn about this.

## What could not be cross-checked

- **Real spill behavior.** AppleScript can only array-enter a formula across a fixed range, so a short spill shows `#N/A` in the spare cells. The values were compared with that in mind. `H2#` and `#SPILL!` were tested in the engines only.
- **Reading an uploaded workbook.** Formualizer has `fromXlsxBytes`, which may serve the Phase 2 upload check. It has not been tried.

## How the cross-check ran

`scripts/dev/excel-crosscheck.py` drives a new unsaved workbook through AppleScript, one short script per test. Writing a file into Excel's sandbox folder hung on a macOS prompt, so nothing is written to disk. The script refuses to run if Excel is already open, so it never touches your work.
