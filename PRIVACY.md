# Privacy

Plain for Windows makes exactly one kind of network request, and only if you leave it switched on: once a day it asks GitHub what the latest released version is, so it can tell you when there is a newer one.

That request sends no identifier, nothing about you, and nothing about your files. It says only which version you are running, because GitHub asks every caller to name itself. Nothing is downloaded and nothing installs itself: if there is a new version, Plain puts a line in the status bar with a button that opens the release page in your browser, and you decide. Turn it off in Reading and settings and Plain makes no network request at all, ever.

There is no telemetry, no crash reporting, no analytics and no account.

Otherwise it reads the file you open and writes the file you save.

It does keep one small settings file, because people asked it to: a list of the last ten files you opened, how big you set the text, and whether you want the preserved panel showing. It is plain JSON at `%LocalAppData%\Plain for Windows\settings.json`, it holds nothing about you beyond those file paths, and deleting it costs you nothing but the list. Earlier versions kept nothing at all; if you would rather it stayed that way, delete the file and it will not be missed.

If keeping unsaved work is left on, a copy of anything you have changed but not saved is written into a `recovery` folder beside that settings file, every half minute, and offered back the next time Plain starts. Those copies are your documents, so they are as private as the originals and they live on your machine only. Saving the real file deletes its copy, and the setting can be turned off in Reading.

The one thing it writes, other than the file you asked it to save, is a temporary file next to your document while saving, named after it with `.plain-tmp` on the end. That file is renamed into place the moment it is complete, so a save that is interrupted leaves your original untouched rather than half-written.

`plain.exe`, the console twin, behaves the same way, and never checks for updates at all: it is meant for scripts and scheduled jobs, which should not reach the network because of something you did not ask for.
