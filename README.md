# Plain for Windows

**Office without the bloat.** Opens Word, Excel and PowerPoint files, does the fifth of the job that fills most of the day, and never damages the rest.

**[Download Plain-for-Windows-1.4.0-x64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.4.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.4.0-arm64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.4.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime to install, no administrator. Put it anywhere and double-click. Windows will warn you about an unknown publisher: choose **More info**, then **Run anyway**. It is not signed, because a signing certificate costs money this does not make.

The console twin, for scripts: [plain-1.0.0-x64.exe](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.0.0-x64.exe) and [plain-1.0.0-arm64.exe](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.0.0-arm64.exe).

![Plain showing a workbook](docs/screenshots/book.png)

## Why this exists

Every other editor that opens a `.docx` reads it into its own idea of a document and writes that idea back out. Whatever the idea has no room for is quietly gone: the chart, the macro, the tracked changes, the header nobody looked at. You open a file to change three words, you save, and you have handed back something subtly different from what you were sent.

Plain does the opposite. It opens the file as the package of parts it really is, holds every part as the exact bytes it occupies, and writes those same bytes back for every part it did not edit.

So the promise is not "it looks the same". It is this:

> **Open a file, save it without changing anything, and you get the same file back, byte for byte.**

That is a claim a test can check rather than a claim a README can make, and it is checked on every build against real Word, Excel and PowerPoint documents. You can check it yourself:

```
plain roundtrip *.docx *.xlsx *.pptx
```

Change one cell and only the parts holding that cell are rewritten. The status bar counts it on every save: *10 parts read, 4 shown, 6 kept byte for byte*.

## The 80/20 of it

Word, Excel and PowerPoint are enormous, and almost none of that is what almost anyone does. The real day is typing in cells, fixing a total, changing a paragraph, correcting a slide, replacing a name throughout, reading a comment someone left, printing a PDF. That is the fifth of the features that covers the great majority of the work, and it is all Plain has: one row of buttons, no ribbon, no tabs to hunt through, no sign-in, no first-run tour.

The other four fifths are not dropped. They are kept. Plain is not a smaller Office that throws away what it cannot draw, it is a smaller Office that carries what it cannot draw through untouched, so the parts of the file you never use survive the parts you do.

| | Office | Plain |
|---|---|---|
| Download | an installer and about 4 GB on disk | one exe, under 60 MB |
| Install | installer, sign-in, first-run tour | put it anywhere and double-click |
| Account | Microsoft account | none |
| Network | always | a daily version check you can turn off, and nothing else |
| Charts, macros, pivot tables | edits them | keeps them byte for byte |

## What it does

**Spreadsheets.** Cells, text and formulas across every sheet. Sort rows by any column, with Plain refusing whenever a formula would be made to mean something else. Filter to show only the rows you want. Drag column edges to resize, double-click one to fit it to its contents, and keep the header rows on screen while the rest scrolls. What the selection adds up to is in the status bar. Joined cells are drawn as one. Sheets can be added, renamed, moved and taken out, with every formula that named a renamed sheet rewritten to follow it. Rows have a height, cells have colour, alignment and lines round them, and columns can be held on screen as well as rows. Take out rows that say the same thing, split a column, and ask any cell what it reads and what reads it. Insert and delete rows and columns, and every formula in the workbook is rewritten so it still means what it meant. Totals are worked out as you go, covering the common functions including lookups, conditional sums and dates, and refusing anything it does not fully understand rather than guessing. Number formats, bold, fill down, copy and paste as tab separated text.

**Documents.** The body text, with headings and tables shown as headings and tables. Bold, italic, heading levels, bullets and numbered lists. Word count.

**Presentations.** The text on each slide, picked from a rail of slides.

**All three.** **New** makes an empty spreadsheet, document or presentation: pick which, then say where it goes. Find and replace across the whole file. Save a copy. Print. Save as PDF. Comments and tracked changes, read, accepted or turned down one at a time. Document properties, viewable and strippable. The pictures inside the file, shown so you can see what you are sending. Multi-step undo. Unsaved work kept every half minute in case the machine stops.

![Plain showing a document](docs/screenshots/doc.png)

## What it keeps but does not show

Charts, pivot tables, macros, SmartArt, embedded objects, slide layouts, masters, themes, headers, footers, footnotes, and anything else. Every one is in the saved file unchanged.

The panel on the right names them, in as few lines as it honestly can. Several of a kind become one line, so a document with eighteen embedded fonts says *18 embedded fonts, 10.4 MB* rather than eighteen rows you cannot tell apart. Across a set of real business documents that takes the panel from 847 rows to 86.

## Honest limits

- **No page layout.** Matching Word's pagination needs Word's own fonts and line breaking. Plain shows a document as one scrolling column and says so. Printing and PDF give you that view, not Word's pages.
- **It does not evaluate every formula.** It covers arithmetic, comparisons, text, lookups, conditional sums and dates. Anything else shows nothing rather than something wrong, and Excel recalculates the file when it opens it.
- **No fonts, colours or alignment.** Bold, italic, heading level and lists, and no more.
- **It does not draw shapes or place pictures.** They are kept, listed, and can be looked at.
- **Retyping a paragraph flattens mixed formatting inside it.** Plain says so when it happens.
- **English only, left to right.** There is no right-to-left support in the editing surface.
- **It has never been opened in Microsoft Office.** Everything here is checked against LibreOffice, against macOS's PDF engine, and against Plain's own byte-for-byte round trip. Those are good proxies and they have caught real bugs, but Office is stricter in places and nobody has yet confirmed this on a machine that has it.
- **Not a replacement for Office.** It is what to reach for when you need to change three words in a contract without a four gigabyte install.

## How it is checked

| | |
|---|---|
| Checks in the engine | 641 |
| End to end through the console | 36 |
| Driving the real window with real keystrokes | 21 |
| Real business documents round tripped byte for byte | 39 of 39 |
| Real documents edited, saved, and reopened elsewhere | 39 of 39 |
| Formula results agreeing with LibreOffice | 52 of 53 |
| Damaged-file mutations survived | 90,000 |

Everything runs on Windows on ARM and on x64. `PLAIN_CORPUS=<folder>` points the checks at a folder of your own documents and demands a byte-identical round trip on every one.

## Command line

`plain.exe` is the same program as a console app.

```
plain info <file>                what the file is, and what Plain keeps untouched
plain parts <file> [--json]      every part, and whether Plain shows it or preserves it
plain text <file> [--numbered]   the text, as plain text
plain cells <file> [sheet]       every filled cell, as reference<tab>value
plain get <file> <ref>           one cell, or one block by number
plain set <file> <ref> <value>   change one cell or block, then save
plain replace <file> <find> <with> [--case] [--whole] [--dry-run]
plain row <file> insert|delete <n>        put a row in or take one out
plain column <file> insert|delete <ref>   the same for a column
plain new <file.xlsx|.docx|.pptx>         make a new empty file
plain pdf <file> [out.pdf]                write it out as a PDF
plain csv <file> [sheet] [--formatted]    a sheet as comma separated values
plain import <file> <csv> [at]            read a csv into a sheet
plain count <file>                        words, characters, paragraphs
plain images <file> [--save <dir>]        the pictures inside it
plain changes <file> [--accept N|--reject N|--accept-all|--reject-all]
plain comments <file> [--remove N|--remove-all]
plain props <file> [--set Name=value] [--strip]
plain apply <file> --script <s>           many changes in one pass
plain roundtrip <file>...                 prove a save changes nothing
plain selftest
```

Exit codes: 0 fine, 1 something to look at, 2 problem, 64 usage.

## Privacy

No account, no telemetry, no analytics, and nothing about you or your files ever leaves the machine.

The one exception, and it is the only network request Plain makes: once a day it asks GitHub what the latest version is, and if there is a newer one it says so in the status bar. Nothing downloads or installs itself. Turn it off in Reading and settings and Plain reaches the network never. The console twin never checks at all. See [PRIVACY.md](PRIVACY.md).

## Building it yourself

```
dotnet run --project src/Plain.Selftest      # the checks
scripts/publish.sh all                       # the four exes into dist/
```

`Plain.Core` is a dependency-free .NET 9 library and builds anywhere. The window is WPF and needs Windows.

## Help

[docs/Help.html](docs/Help.html) covers what it edits, what it keeps, the keyboard and the honest limits.

Free, MIT, built by Keith Adler. More at [keithadler.github.io](https://keithadler.github.io/).
