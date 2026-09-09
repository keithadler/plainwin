**[Download Plain-for-Windows-1.3.0-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.3.0/Plain-for-Windows-1.3.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.3.0-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.3.0/Plain-for-Windows-1.3.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

## Sheets

Right-click a sheet tab to add one, rename it, move it or take it out. Slides could already be managed this way and sheets could not.

Renaming is the one that needs care beyond the tab. A formula saying `Detail!B2` means the sheet called Detail, so renaming without rewriting the formulas that name it would leave them pointing at a sheet that is not there. Every formula in the workbook is rewritten to follow the new name, and Plain says how many it changed.

Taking a sheet out is refused while any formula still reads it, and says which. Nothing is silently broken. What was on it stays in the file, unused, rather than being deleted.

## How a sheet looks

Rows have a height, as columns already had a width. Columns can be held on screen as well as rows; both live in the same pane, so they are set together and neither quietly undoes the other.

Right-click a selection for **How these cells look**: left, centred, right, wrapped, and a colour behind the cells or for their words. Each change builds on the format the cell is already wearing, so giving a cell a colour does not throw away its number format.

## Documents

**Headers and footers.** The lines that run along the top and bottom of every page: what they say, and a way to change the words. They are the part of a file people forget is there. A header naming the client a template was written for travels with every copy of it, and nobody sees it on screen.

**Where a paragraph sits.** Ctrl+L, Ctrl+E, Ctrl+R and Ctrl+J, the shortcuts every word processor uses, and the document shows the result rather than only recording it.

## Links, and where they actually go

A link's words and its destination are two different things, and a file is one of the easier places to hide that they disagree. Plain lists both side by side.

It will open ordinary web and mail addresses and nothing else. Anything that would run a program is shown in red and left alone. A document should not be able to make something happen because you clicked it.

## Compare with another file

What changed between two versions: which parts of the file differ at all, and which lines of text arrived or went. Plain holds a file as the parts it is made of, so it can answer the first question exactly rather than by guessing.

It compares what is there. It does not try to work out how one became the other, so a line that moved reads as one gone and one arrived. Saying that plainly is better than guessing at a move and being wrong.

## See the pages first

Print preview, on the paper and margins chosen in settings, using the same laying out that printing does. A preview built a different way would be a picture of something else.

## Console twin

`plain sheet <file> add|rename|remove|move`, `plain height`, `plain align`, `plain colour`, `plain band`, `plain links`, `plain compare <before> <after>`.

939 checks, up from 843. Nothing about the file format changed.
