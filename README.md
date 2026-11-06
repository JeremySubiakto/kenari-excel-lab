# Kenari Excel Lab

Learn Excel free, from the first SUM to PivotTables, Power Query, loans and macros.

**Play it: https://jeremysubiakto.github.io/kenari-excel-lab/**

You are the new analyst at Rasa Kenari, a made-up food company, and later at Bank Kenari, the made-up bank that lends to its distributors. Colleagues send requests in Kicau, a work chat. You answer each one on a practice sheet that works like Excel, and Kicau checks your formula on a second sheet with different numbers, so a typed answer or a frozen reference does not pass.

- 87 requests in 10 chapters: sums and averages, IF and grades, lookups, text and dates, dynamic arrays, reports (PivotTables, charts, conditional formatting, data validation), Power Query, loans and what-if, macros, and a final early-warning report.
- Each chapter ends with one request done in real Excel: download the workbook, do the work, upload it back, and Kicau checks the file.
- Works on a phone and a computer. No account. Your name, answers and progress stay in your browser.

Rasa Kenari, Bank Kenari, their products, distributors and people are all invented. See `docs/NAMES.md`.

Made by Jeremy Subiakto. Also by him: [Kenari Collection SQL Lab](https://jeremysubiakto.github.io/kenari-sql-lab/). Something wrong? [Report a problem](https://github.com/JeremySubiakto/kenari-excel-lab/issues/new?template=problem.md).

## Run it on your computer

You need Node 22 or newer. There is nothing to install.

```bash
node scripts/dev/serve.mjs
```

Then open http://localhost:8766/

## Test it

```bash
node --no-warnings scripts/verify.mjs
node --no-warnings scripts/dev/kicau-e2e.mjs
```

The first checks the data, the answer checker and the macro stepper. The second plays every request on a phone and a laptop in a private Chrome, on port 8766.

## What is where

- `index.html`: the app, Kicau
- `assets/js/kicau/missions.js`: every request, lesson, hint and reply
- `assets/js/kicau/app.js`: the chat, the practice sheet, the builders and the Excel download and upload
- `assets/js/core/data.js`: the seeded data for Rasa Kenari and Bank Kenari
- `assets/js/core/check.js`: answer checking
- `assets/js/core/vba.js`: the macro stepper. It covers first macros and is not a full VBA interpreter
- `assets/engine/`: the formula engine, Formualizer 0.13.0 (MIT or Apache-2.0), shipped with the site
- `docs/ENGINE.md`: why this engine, and where it differs from Excel

## License

Code under MIT (`LICENSE`). Lessons under CC BY-NC 4.0 (`LICENSE-LESSONS.md`).
