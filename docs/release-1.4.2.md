**[Download Plain-for-Windows-1.4.2-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.2/Plain-for-Windows-1.4.2-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.4.2-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.2/Plain-for-Windows-1.4.2-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

Ten things that are not about editing documents.

## You can check what you downloaded

Every release before this one was built on one laptop and uploaded by hand. Plain is not signed, because a certificate costs money this does not make, so the only thing anyone could check was nothing at all.

From this release the executables are built by a public job on a clean runner, from a commit anyone can read, and a `SHA256SUMS.txt` is published beside them by the same job. Compare what you have:

```powershell
Get-FileHash "Plain-for-Windows-1.4.2-x64.exe" -Algorithm SHA256
```

That does not make the download trusted. It makes it traceable, which is the honest version of the same thing.

## Something going wrong no longer takes your work with it

An unexpected error used to end in the .NET crash box, which says nothing you can act on and takes whatever was unsaved with it. Now a copy of anything unsaved is kept first, what happened is written to a file beside the settings, and you are told in a sentence what happened, where the note is, and that the work was kept. Nothing is sent anywhere.

## A document cannot make Plain fetch anything

A .docx is XML, and XML has a way of saying "go and fetch this and paste it in here". .NET happened not to fetch it, but it accepted the file and quietly dropped what the entity stood for, so a document could lose text with nobody told.

Now such a file is refused, and the parser says its settings out loud rather than leaving them to a default that could change. There are checks that a document naming a file on disk, a web address, or an outside definition is turned away, and that entities which multiply are refused quickly rather than chewed on.

## It can be read by a screen reader

The grid is drawn rather than built out of controls, so a screen reader found one blank surface and could say nothing about it. It now says where you are and what is in the cell as you move, including whether the cell is joined with its neighbours. The buttons and boxes have names.

## The rest

**An icon**, instead of the generic Windows one in Explorer, the taskbar and Alt+Tab.

**F1** lists every keystroke, and there is a **help page inside the exe** that works with no network, in English or Spanish.

**Portable mode.** Put an empty file called `plain-portable` beside the exe and everything lives in a folder beside it: settings, recent files, copies of unsaved work. Plain then leaves nothing at all on the machine, which is what you want on a USB stick, on a locked-down PC, or on somebody else's computer.

**A SECURITY.md** saying what Plain is exposed to, what it refuses to do, that it is not signed and why, and what would worry me.

**Startup was measured** rather than assumed: the window appears in 87 ms and the console twin answers in 122 ms, so nothing needed doing.

1,020 checks, up from 1,014. Nothing about the file format changed.
