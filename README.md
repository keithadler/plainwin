# Plain for Windows

Opens Word, Excel and PowerPoint files. Edits the things you actually change. Never damages the rest.

## Download

**[Plain-for-Windows-1.4.4-x64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.4.4-x64.exe)** — for ordinary Intel and AMD PCs

**[Plain-for-Windows-1.4.4-arm64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.4.4-arm64.exe)** — for Windows on ARM

Then:

1. Put it anywhere. There is no installer.
2. Double-click it.
3. Windows says "unknown publisher". Choose **More info**, then **Run anyway**.

Windows 10 or 11. No runtime to install, no administrator, no account, no sign-in.

It is not signed, because a certificate costs money this does not make. What you can check instead: the executables
are built by [a public job](.github/workflows/release.yml) from a commit you can read, and
[SHA256SUMS.txt](https://github.com/keithadler/plainwin/releases/latest/download/SHA256SUMS.txt) is published beside
them.

```powershell
Get-FileHash "Plain-for-Windows-1.4.4-x64.exe" -Algorithm SHA256
```

![Plain showing a workbook](docs/screenshots/book.png)

## The promise

Every editor that opens a `.docx` reads it into its own idea of a document and writes that idea back out. Whatever
it has no room for is quietly gone: the chart, the macro, the tracked changes, the header nobody looked at. You
change three words and hand back something subtly different.

Plain opens the file as the package of parts it really is, and writes back every part it did not edit as the exact
bytes it found.

> **Open a file, save it without changing anything, and you get the same file back, byte for byte.**

That is not a claim a README can make. It is one a test can check, and it runs on every build. Check it yourself:

```
plain roundtrip *.docx *.xlsx *.pptx
```

Change one cell and only the parts holding that cell are rewritten. The status bar counts it every time you save:
*12 parts read, 6 shown, 6 kept byte for byte.*

## What it does

### Spreadsheets

| | |
|---|---|
| **Cells** | Text, numbers and formulas across every sheet |
| **Formulas** | Worked out as you type, including lookups, conditional sums and dates. Anything it does not fully understand shows nothing rather than something wrong |
| **Rows and columns** | Insert, delete, sort, filter, take out duplicates, split a column. Every formula in the workbook is rewritten so it still means what it meant |
| **Sheets** | Add, rename, move, remove. Renaming rewrites every formula that named the sheet |
| **Look** | Colour, alignment, wrapping, borders, number formats, bold. Column widths and row heights, with autofit |
| **On screen** | Freeze rows and columns, joined cells drawn as one, and what the selection adds up to in the status bar |
| **Understanding it** | Ask any cell what it reads and what reads it. See the choices a cell will accept |

### Documents

Paragraphs, headings and lists. Tables, with rows you can add and remove. Bold and italic. Headers and footers.
Links you can follow, edit or create. Comments and tracked changes, read and settled one at a time. Word count.
Pictures you can insert and describe.

### Presentations

The text on every slide. Speaker notes, read and changed. Slides you can add, remove and reorder. Pictures, shown
and described.

### All three

New blank files. Find and replace across the whole file, or across a whole folder. Print and PDF with page setup.
Save a copy. Undo and redo. Document properties. A list of what a file would carry with it before you send it.
Compare two versions. Unsaved work kept every half minute in case the machine stops.

![Plain showing a document](docs/screenshots/doc.png)

## What it keeps but does not show

Charts, pivot tables, macros, SmartArt, embedded objects, slide layouts, masters, themes, footnotes, and everything
else it cannot draw. All of it is in the saved file unchanged.

The panel on the right names them in as few lines as it honestly can. Eighteen embedded fonts is one line saying
*18 embedded fonts, 10.4 MB*, not eighteen rows you cannot tell apart. Across a set of real business documents that
takes the panel from 847 rows to 86.

## Honest limits

- **No page layout.** Matching Word's pagination needs Word's own fonts and line breaking. Plain shows a document
  as one scrolling column and says so. Printing and PDF give you that view, not Word's pages.
- **It does not draw shapes, charts or pictures in place.** They are kept, listed, and can be opened.
- **Retyping a paragraph flattens mixed formatting inside it.** Plain says so when it happens.
- **The window is English.** The help page is also in Spanish; the buttons are not.
- **Left to right only.** There is no right-to-left support in the editing surface.
- **It has never been opened in Microsoft Office.** Everything is checked against LibreOffice, against macOS's PDF
  engine, and against Plain's own byte-for-byte round trip. Those have caught real bugs, but Office is stricter in
  places and nobody has confirmed this on a machine that has it.
- **Not a replacement for Office.** It is what to reach for when you need to change three words in a contract
  without a four gigabyte install.

## Privacy

No account, no telemetry, no analytics. Nothing about you or your files leaves the machine.

The one network request Plain makes is a daily check with GitHub for a newer version. It sends no identifier and
nothing about your files, downloads nothing and installs nothing. Turn it off in Reading and settings and Plain
makes no network request at all. The console twin never checks. See [PRIVACY.md](PRIVACY.md) and
[SECURITY.md](SECURITY.md).

**Portable:** put an empty file called `plain-portable` beside the exe and everything Plain remembers lives in a
folder beside it. It then leaves nothing at all on the machine.

## How it is checked

| | |
|---|---|
| Checks in the engine | 1,008 |
| End to end through the console | 52 |
| Driving the real window with real keystrokes | 29 |
| Real business documents round tripped byte for byte | 39 of 39 |
| Real documents edited, saved, and reopened elsewhere | 39 of 39 |
| Formula results agreeing with LibreOffice | 52 of 53 |
| Damaged-file mutations survived | 90,000 |

That is what `dotnet run --project src/Plain.Selftest` gives you on a fresh clone. Pointing it at the demo files
adds twenty-nine more that need real content to run:

```
PLAIN_CORPUS=docs/demo dotnet run --project src/Plain.Selftest      # 1,024
```

All of it runs on Windows on ARM and on x64. `PLAIN_CORPUS=<folder>` also works on a folder of your own documents,
and demands a byte-identical round trip on every one.

## The console twin

`plain.exe` is the same program without a window, for scripts and scheduled jobs.
[x64](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.4.4-x64.exe) ·
[arm64](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.4.4-arm64.exe)

**Reading**

```
plain info <file>                 what it is, and what Plain keeps untouched
plain parts <file> [--json]       every part, and whether Plain shows it or preserves it
plain text <file> [--numbered]    the text, as plain text
plain cells <file> [sheet]        every filled cell, as reference<tab>value
plain get <file> <ref>            one cell
plain count <file>                words, characters, paragraphs
plain reads <file> <cell>         what a formula reads, and what reads the cell
plain choices <file>              the choices cells will accept
plain links <file>                every link, and where it really goes
plain hidden <file>               what this file would carry with it
plain find <folder> <words>       which files in a folder hold those words
plain compare <before> <after>    what changed between two versions
plain images <file> [--save <dir>]
plain describe <file>             the words that describe each picture
plain notes <file>                what the presenter was going to say
```

**Changing**

```
plain set <file> <ref> <value>    change one cell or block
plain replace <file> <find> <with> [--case] [--whole] [--dry-run]
plain row|column <file> insert|delete <n>
plain sort <file> <range> <column> [--down]
plain tidy <file> dedupe|split <range> [--on ","]
plain sheet <file> add|rename|remove|move <n> [name|to]
plain width|height <file> <n> <size|fit|auto>
plain freeze <file> <rows> [columns]
plain align|colour|border <file> <range> ...
plain slide <file> add|remove|move
plain tablerow <file> add|remove <table> <row>
plain paragraph <file> add|remove <number>
plain link|picture <file> ...
plain band <file> header|footer <text>
plain props <file> [--set Name=value] [--strip]
plain changes <file> [--accept N|--reject N|--accept-all|--reject-all]
plain comments <file> [--remove N|--remove-all]
plain apply <file> --script <s>   many changes in one pass
```

**Making and proving**

```
plain new <file.xlsx|.docx|.pptx>
plain pdf <file> [out.pdf] [--paper A4|Letter] [--landscape] [--footer "..."]
plain csv <file> [sheet] [--formatted]
plain import <file> <csv> [at]
plain explorer on|off|status      the right-click menu
plain roundtrip <file>...         prove a save changes nothing
plain selftest
```

Exit codes: 0 fine, 1 something to look at, 2 problem, 64 usage.

## Building it yourself

```
dotnet run --project src/Plain.Selftest      # the checks
scripts/publish.sh all                       # the four exes into dist/
```

`Plain.Core` is a dependency-free .NET 9 library and builds anywhere. The window is WPF and needs Windows.

## Help

Press **F1** in the app for every keystroke, or open [docs/Help.html](docs/Help.html)
([español](docs/Help.es.html)).

Free, MIT, built by Keith Adler. More at [keithadler.github.io](https://keithadler.github.io/).
