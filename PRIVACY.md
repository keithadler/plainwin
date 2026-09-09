# Privacy

Plain for Windows does not have any network code in it. Not an update check, not telemetry, not crash reporting, not an account. There is nothing to turn off.

It reads the file you open and writes the file you save.

It does keep one small settings file, because people asked it to: a list of the last ten files you opened, how big you set the text, and whether you want the preserved panel showing. It is plain JSON at `%LocalAppData%\Plain for Windows\settings.json`, it holds nothing about you beyond those file paths, and deleting it costs you nothing but the list. Earlier versions kept nothing at all; if you would rather it stayed that way, delete the file and it will not be missed.

The one thing it writes, other than the file you asked it to save, is a temporary file next to your document while saving, named after it with `.plain-tmp` on the end. That file is renamed into place the moment it is complete, so a save that is interrupted leaves your original untouched rather than half-written.

`plain.exe`, the console twin, behaves the same way.
