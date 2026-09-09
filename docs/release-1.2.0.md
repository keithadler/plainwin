**[Download Plain-for-Windows-1.2.0-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.2.0/Plain-for-Windows-1.2.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.2.0-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.2.0/Plain-for-Windows-1.2.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

## Before you send it

One list of what travels with a file that you may not have meant to send: who wrote it and who saved it last, comments arguing about a number, tracked changes showing what a clause used to say, sheets and rows someone hid rather than deleted, speaker notes meant for the presenter.

It takes out only what you tick. Hidden sheets and rows are somebody's working rather than an accident, and speaker notes live in parts of their own, so Plain lists those and leaves them where they are rather than tearing them out. And it says plainly that it can only find what it knows to look for: this is a list, not a promise that a file is safe.

## Working in a sheet

**Filter.** Show only the rows whose cell in a column contains what you type. Nothing is written to the file: it is a way of looking at the sheet, hiding a row never deletes anything, and clearing it puts every row back.

**What the selection adds up to**, in the status bar: how many numbers, the sum, the average, the lowest, the highest. It is the question a spreadsheet is usually opened to answer, and having it there means not typing a formula into an empty cell and deleting it afterwards.

**Getting about.** Ctrl with an arrow goes to the end of the run of filled cells, or across a gap to the next thing there is. Ctrl+Home and Ctrl+End go to the corners. Ctrl+G asks which cell and goes there.

**Joined cells** are drawn as one cell instead of several empty ones with lines through them, and clicking any part of a block selects the corner that holds the value.

## Redo

Ctrl+Y, and Ctrl+Shift+Z. Every step now carries how to undo it and, where it can be worked out, how to do it again. A step that cannot say how to repeat itself makes everything after it unrepeatable, so the pile waiting to be redone is thrown away rather than left to put things back in the wrong order. Half a redo is worse than none.

## Changing the shape of a file

**Slides** can be added, taken out and moved. **Rows** can be added to and taken out of a table in a document; a new row is a copy of a neighbour with the words removed, so it keeps the borders, shading and widths of the table it joins.

A change of shape empties the undo stack and builds the view again from the file, and says so. The numbering of slides and paragraphs is different afterwards, and carrying on with the old numbering is how a tool ends up editing the wrong thing.

Removing a slide takes it out of the running order and leaves its part in the file, unused. That costs a few kilobytes; deleting a part that something unnoticed still points at is how a deck gets broken.

## Find in a folder

Which Office files in a folder hold the words, and where in each one. Ctrl+Shift+F.

It only ever reads. It tells you which files matter and then you open those and change them yourself, one at a time, watching what happens. A tool that offered to change forty files at once would be asking for a great deal of trust in exchange for very little work saved.

## A file with a password says so

Opening one used to report "no ZIP end record found", which sends someone looking for damage that is not there. A file with a password on it is not a damaged package, it is a different kind of container altogether, and Plain now says that, says it cannot remove the password, and says what would work instead. An older `.doc`, `.xls` or `.ppt` is named as an older file too.

## Under all of it

The package can now add a part, which is what a deck needs in order to gain a slide. It is deliberately separate from writing an existing part, so that adding one can never happen by accident, and the checks prove that after a part is added every part that was already there still comes back byte for byte. The promise the whole app rests on is unchanged.

843 checks, up from 744.

## Console twin

`plain hidden <file> [--remove names,comments,tracked]`, `plain find <folder> <words> [--deep]`, `plain slide <file> add|remove|move`, `plain tablerow <file> add|remove <table> <row>`.

Nothing about the file format changed. Files written by 1.0, 1.1 and 1.2 are the same bytes.
