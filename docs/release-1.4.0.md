**[Download Plain-for-Windows-1.4.0-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.0/Plain-for-Windows-1.4.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.4.0-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.0/Plain-for-Windows-1.4.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

## What a cell reads, and what reads it

This is the question behind most spreadsheet mistakes. A total looks wrong and you cannot see which cells it is adding up; you change a number and cannot see what else moved because of it. Right-click a cell and Plain says, both ways:

```
D5 reads:
    Summary!C5   26,400
    Summary!B5   24,000

D5 is read by:
    Summary!D11   =SUM(D5:D10)
    Summary!B13   =INDEX(A5:A10,MATCH(MAX(D5:D10),D5:D10,0))
    Summary!B15   =SUMIF(D5:D10,">0")
```

The answer was in the file already. Plain only reads it out.

## Tidying a list

**Take out rows that say the same thing.** The first of each set stays where it is, so what is left is in the order it was in.

**Split a column** at a comma or a space, into the columns to its right. It refuses rather than writing over the column beside it, and says which column stopped it.

Both refuse, with a reason, wherever a formula would be made to mean something else. That is the same care sorting takes.

## The choices a cell will accept

A sheet someone else built often has cells that only take certain values: a status that must be one of four words. Plain does not draw the little arrow, but it reads the rule and says what the cell wants. A cell that silently refuses what you type is worse than one that tells you.

## Lines round cells

Thin, medium, thick, dotted, dashed or double. All four sides, or one at a time.

## Speaker notes, and the words that describe a picture

**Speaker notes** travel with every copy of a deck and nobody sees them on the slide. Plain shows what each says and lets it be changed.

**Alt text** is the thing everyone is asked for and skips. Plain lists every picture, says which have nothing describing them, and writes the description in both the places Word and PowerPoint look, because a picture described in only one of them is described to only half the readers.

## Links you can make, and a picture you can put in

A link goes on a whole paragraph. Only ordinary web and mail addresses: Plain will not write into a document a link it would refuse to open itself. The words take the document's own link styling where it has one, and where it has none Plain says so rather than inventing a blue underline.

A picture goes at the end of a document at a size you give. Nothing is scaled or re-encoded, so what goes in is the file you chose.

## What is already in the column, offered as you type

A list of clients typed slightly differently each time is the most common way a spreadsheet quietly goes wrong. Type a few letters and Plain offers what is above, with the rest selected, so carrying on typing replaces it and pressing Enter takes it.

## Console twin

`plain border`, `plain tidy dedupe|split`, `plain reads`, `plain choices`, `plain notes`, `plain describe`, `plain link`, `plain picture`.

1,014 checks, up from 939. Nothing about the file format changed.
