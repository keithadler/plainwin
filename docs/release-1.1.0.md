**[Download Plain-for-Windows-1.1.0-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.1.0/Plain-for-Windows-1.1.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.1.0-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.1.0/Plain-for-Windows-1.1.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

## Sort rows

By any column, up or down. Numbers sort by value rather than as text, so a hundred comes after twenty. Blank rows go to the bottom whichever way you sort.

Sorting is the one common spreadsheet job that can quietly ruin a file, because moving a row moves what every formula was pointing at. Plain will not guess its way through that. It refuses, and says why, whenever a formula would be made to mean something else: one inside the block being sorted, one reading only part of it, or a lookup whose answer depends on where a value sits.

A total underneath a table is the ordinary case, and it goes ahead. `SUM(D2:D8)` reads every row of the table, and shuffling those rows does not change the total, so there is nothing to protect you from.

## Column widths, and a header that stays

Drag the edge of a column heading to resize it. Double-click the edge to fit the column to its longest value. Widths are written where the file keeps them, so Excel opens the sheet the way you left it.

Right-click and choose to keep the rows above the selected one on screen, and they stay put while the rest scrolls. This is read from and written to the sheet's own pane settings, so a sheet that already had a frozen header keeps it here, and one you freeze here is frozen when it goes back to Excel.

## Page setup for printing and PDF

Paper (A4, Letter, Legal, A3, A5), sideways or not, how much white at the edges, and a line at the top and bottom where `{page}` and `{pages}` are filled in.

Plain still lays the text out on its own page, and still says so: it does not know where Word would break one, and a PDF that pretended to would be a lie in a format people trust.

## Open from Explorer

An optional **Edit in Plain** line on the right-click menu for Word, Excel and PowerPoint files, and Plain offered under Open with.

It writes under your own account only, so it never asks to be an administrator and never touches anyone else's. It does not make itself the default for anything: Windows decides defaults, and an app that seizes them is the kind of app this one exists as an alternative to. Turning it off takes back exactly what it added, which is checked on every build by counting what is in the registry before and after: thirteen entries on, none left off.

## A daily check for a new version

Once a day Plain asks GitHub what the latest release is, and if there is a newer one it puts a line in the status bar with a button to see what changed. That is all it does. No identifier is sent, nothing about your files is sent, nothing is downloaded and nothing installs itself. The switch is in Reading and settings, and turned off Plain makes no network request at all. The console twin never checks, because a scheduled job should not reach the network because of something you did not ask for.

Plain 1.0 said it had no network code at all. That is no longer true, so the README, the Help, PRIVACY.md, the About box and the announcement page now say what the one request is and how to stop it, rather than leaving a promise standing that the app no longer keeps.

## Fixed

**Bullets and curly quotes came out of a PDF as nonsense.** A bullet appeared as `€42`. Characters above 126 were written as an octal escape of their Unicode number, and a PDF reader takes the first three digits as one character and prints the rest as text. Bullets, dashes, curly quotes and the ellipsis are now written where the built-in fonts actually keep them, and anything genuinely outside those fonts becomes a question mark rather than nonsense. This had been wrong since 1.0.0, and affected any document with a bulleted list.

## Console twin

`plain sort <file> <range> <col> [down] [sheet]`, `plain width <file> <col> <chars|fit>`, `plain freeze <file> <rows>`, `plain explorer on|off|status`, and `--paper`, `--landscape`, `--margin`, `--header`, `--footer` for `plain pdf`.

Nothing about the file format changed. Files written by 1.0 and by 1.1 are the same bytes.
