# Privacy

Plain for Windows does not have any network code in it. Not an update check, not telemetry, not crash reporting, not an account. There is nothing to turn off.

It reads the file you open and writes the file you save. It keeps no settings file, no recent-files list and no cache. Close it and there is no trace of what you looked at.

The one thing it writes, other than the file you asked it to save, is a temporary file next to your document while saving, named after it with `.plain-tmp` on the end. That file is renamed into place the moment it is complete, so a save that is interrupted leaves your original untouched rather than half-written.

`plain.exe`, the console twin, behaves the same way.
