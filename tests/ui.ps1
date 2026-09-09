# Drives the real window with real keystrokes and checks the file on disk afterwards.
# Must run in an interactive session (a scheduled task with /it), because a service session has no desktop.
# Usage:  powershell -File tests\ui.ps1 C:\Plain
# $Home is a PowerShell built-in, so the folder is $Root.
param([string]$Root = "C:\Plain")

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Fore {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@
$ErrorActionPreference = "Continue"
$exe   = Join-Path $Root "Plain for Windows.exe"
$cli   = Join-Path $Root "plain.exe"
$work  = Join-Path $Root "ui-work.xlsx"
$doc   = Join-Path $Root "ui-work.docx"
$log   = Join-Path $Root "ui-result.txt"
$fail  = 0
$lines = @()

function Note($ok, $what, $detail = "") {
  if ($ok) { $script:lines += "ok    $what" }
  else     { $script:lines += "FAIL  $what $detail"; $script:fail = 1 }
}

function Window($process) {
  $root = [System.Windows.Automation.AutomationElement]::RootElement
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $process.Id)
  for ($i = 0; $i -lt 40; $i++) {
    $w = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
    if ($w) { return $w }
    Start-Sleep -Milliseconds 500
  }
  return $null
}

# The window appears before the file is in it: opening is done when the Loaded handler runs, and on a slow
# machine that is seconds after the frame is on screen. Until the title names the file nothing is open, so
# keystrokes land nowhere and every check after them is measuring an empty window.
function Opened($window, $name) {
  for ($i = 0; $i -lt 60; $i++) {
    if ($window.Current.Name -like "*$name*") { return $true }
    Start-Sleep -Milliseconds 500
  }
  return $false
}

function ById($window, $id) {
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $id)
  return $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
}

function SetText($element, $text) {
  $pattern = $element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
  $pattern.SetValue($text)
  Start-Sleep -Milliseconds 400
}

# Keystrokes go to whatever is in front, so nothing may be typed until the window really is.
function Front($process) {
  for ($i = 0; $i -lt 30; $i++) {
    $process.Refresh()
    $h = $process.MainWindowHandle
    if ($h -ne 0) {
      [Fore]::ShowWindow($h, 9) | Out-Null
      [Fore]::SetForegroundWindow($h) | Out-Null
      Start-Sleep -Milliseconds 400
      if ([Fore]::GetForegroundWindow() -eq $h) { Start-Sleep -Milliseconds 500; return $true }
    }
    Start-Sleep -Milliseconds 400
  }
  return $false
}

function Button($window, $name) {
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::NameProperty, $name)
  return $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
}

function Press($element) {
  $pattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
  $pattern.Invoke()
  Start-Sleep -Milliseconds 700
}

function Texts($window) {
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Text)
  $found = @()
  foreach ($e in $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)) { $found += $e.Current.Name }
  return $found
}

# ---------- a spreadsheet: type into it, save it, check the file ----------

Copy-Item (Join-Path $Root "fixtures\sheet.xlsx") $work -Force
Stop-Process -Name "Plain for Windows" -Force -ErrorAction SilentlyContinue
Start-Sleep 1
$p = Start-Process $exe -ArgumentList "`"$work`"" -PassThru
$w = Window $p
Note ($null -ne $w) "the window opens"
Note ($w -and (Opened $w "ui-work.xlsx")) "and has the file in it before anything is typed"
if (-not $w) { $lines | Set-Content $log; exit 1 }
Note (Front $p) "and comes to the front, so it can be typed into" 

Note ($w.Current.Name -like "*ui-work.xlsx*") "the title names the file" $w.Current.Name

# The grid starts on A1 with the keyboard in it. Type a value and commit it.
[System.Windows.Forms.SendKeys]::SendWait("Northern Europe{ENTER}")
Start-Sleep -Milliseconds 600
[System.Windows.Forms.SendKeys]::SendWait("^s")
Start-Sleep 2

$a1 = (& $cli get $work A1)
Note ($a1 -eq "Northern Europe") "typing into a cell and Ctrl+S writes it to the file" "got '$a1'"

# The status bar must report the promise after a save.
$after = Texts $w
Note (($after | Where-Object { $_ -match "kept byte for byte" }).Count -gt 0) "the status bar states what was kept"

# Undo must put it back.
[System.Windows.Forms.SendKeys]::SendWait("^z")
Start-Sleep -Milliseconds 700
[System.Windows.Forms.SendKeys]::SendWait("^s")
Start-Sleep 2
$a1 = (& $cli get $work A1)
Note ($a1 -eq "Region") "Ctrl+Z puts the old value back and it saves" "got '$a1'"

# Find and replace through the bar.
[System.Windows.Forms.SendKeys]::SendWait("^h")
Start-Sleep -Milliseconds 900
$replaceAll = Button $w "Replace all"
$findBox    = ById $w "FindBox"
$replaceBox = ById $w "ReplaceBox"
Note ($null -ne $replaceAll) "the replace bar appears on Ctrl+H"
Note (($null -ne $findBox) -and ($null -ne $replaceBox)) "both boxes are reachable"
if ($replaceAll -and $findBox -and $replaceBox) {
  SetText $findBox "EMEA"
  SetText $replaceBox "Europe"
  $script:lines += "      find box now: '" + $findBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value + "'"
  $script:lines += "      replace box now: '" + $replaceBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value + "'"
  $before = (Texts $w) -join " | "
  Press $replaceAll
  $said = (Texts $w) | Where-Object { $_ -match "Changed|not in this file|occurrence" }
  $script:lines += "      the app said: '" + ($said -join " ") + "'" 
  [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
  Start-Sleep -Milliseconds 400
  [System.Windows.Forms.SendKeys]::SendWait("^s")
  Start-Sleep 2
  $a3 = (& $cli get $work A3)
  Note ($a3 -eq "Europe") "Replace all, then Ctrl+S, changes the file" "got '$a3'"

  # Was it the keyboard that failed, or the save? Press the button instead and see.
  if ($a3 -ne "Europe") {
    $saveButton = ById $w "SaveBtn"
    if ($saveButton) {
      Press $saveButton
      Start-Sleep 2
      $a3 = (& $cli get $work A3)
      Note ($a3 -eq "Europe") "pressing Save writes it" "got '$a3'"
    } else { $script:lines += "      no save button found" }
  }
}

# The preserved panel must be reachable and must name something.
$keep = Button $w "Preserved  6"
if (-not $keep) { $keep = Button $w "Preserved  5" }
Note ($null -ne $keep) "the preserved button shows a count"

Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
Start-Sleep 1

# ---------- a document: bold a line, check the file ----------

Copy-Item (Join-Path $Root "fixtures\doc.docx") $doc -Force
$p2 = Start-Process $exe -ArgumentList "`"$doc`"" -PassThru
$w2 = Window $p2
Note ($w2 -and (Opened $w2 "ui-work.docx")) "the document is open before anything is typed"
Note ($null -ne $w2) "a document opens too"
if ($w2) {
  $words = (Texts $w2 | Where-Object { $_ -match "^\d[\d,]* words" })
  Note ($words.Count -gt 0) "the status bar shows a word count" ($words -join "")

  # Put the caret in the first line, then bold it.
  [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
  Start-Sleep -Milliseconds 700
  [System.Windows.Forms.SendKeys]::SendWait("^b")
  Start-Sleep -Milliseconds 700
  [System.Windows.Forms.SendKeys]::SendWait("^s")
  Start-Sleep 2

  $props = (& $cli parts $doc | Out-String)
  Note ($props -match "word/document.xml") "the document still reads as a document after a save in the app"
  Stop-Process -Id $p2.Id -Force -ErrorAction SilentlyContinue
  Start-Sleep 1
}

# ---------- tracked changes: settle one from the panel and check the file ----------

$review = Join-Path $Root "ui-review.docx"
Copy-Item (Join-Path $Root "fixtures\review.docx") $review -Force
$p3 = Start-Process $exe -ArgumentList "`"$review`"" -PassThru
$w3 = Window $p3
Note ($w3 -and (Opened $w3 "ui-review.docx")) "the marked up document is open before anything is typed"
Note ($null -ne $w3) "a document with mark-up opens"
if ($w3) {
  $before = (& $cli changes $review | Measure-Object -Line).Lines
  Note ($before -eq 2) "it starts with two tracked changes" "got $before"

  $comments = Button $w3 "Comments  4"
  Note ($null -ne $comments) "the comments button shows the count"
  if ($comments) {
    Press $comments
    $accept = Button $w3 "Accept all tracked changes"
    Note ($null -ne $accept) "the panel offers to settle them"
    if ($accept) {
      Press $accept
      Start-Sleep 1
      $saveButton = ById $w3 "SaveBtn"
      if ($saveButton) { Press $saveButton }
      Start-Sleep 2
      $after = (& $cli changes $review | Out-String)
      Note ($after -match "no tracked changes") "accepting them all, then saving, changes the file" "got '$after'"
      $text = (& $cli text $review | Out-String)
      Note ($text -match "thirty \(30\) days") "what was added is still there"
      Note (-not ($text -match "sixty \(60\) days")) "what was struck out is gone"
    }
  }
  Stop-Process -Id $p3.Id -Force -ErrorAction SilentlyContinue
}


# ---------- a brand new blank document: type, save, and the words must be in the file ----------
# The editors hand their text over when they lose the caret. A new document has exactly one box, so the caret
# never leaves it on its own, and before this was fixed you could type a page, press Ctrl+S, and save nothing.

$blank = Join-Path $Root "ui-blank.docx"
Remove-Item $blank -Force -ErrorAction SilentlyContinue
& $cli new $blank | Out-Null
Note (Test-Path $blank) "the console twin makes a blank document"

if (Test-Path $blank) {
  Stop-Process -Name "Plain for Windows" -Force -ErrorAction SilentlyContinue
  Start-Sleep 1
  $p4 = Start-Process $exe -ArgumentList "`"$blank`"" -PassThru
  $w4 = Window $p4
  Note ($w4 -and (Opened $w4 "ui-blank.docx")) "the blank document is open before anything is typed"
  Note ($null -ne $w4) "a blank document opens"
  if ($w4) {
    Note (Front $p4) "and comes to the front"
    # Click into the one paragraph there is, the way a person would, then type and save without leaving it.
    $box = $null
    $cond = New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
      [System.Windows.Automation.ControlType]::Edit)
    $edits = $w4.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
    foreach ($e in $edits) { if ($e.Current.IsKeyboardFocusable -and -not $box) { $box = $e } }
    if ($box) { $box.SetFocus() }
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("The roof survey is booked for September.")
    Start-Sleep -Milliseconds 600
    [System.Windows.Forms.SendKeys]::SendWait("^s")
    Start-Sleep 2
    $typed = (& $cli text $blank | Out-String)
    Note ($typed -match "roof survey is booked") "typing into a new document and Ctrl+S writes it to the file" "got '$typed'"
    Stop-Process -Id $p4.Id -Force -ErrorAction SilentlyContinue
  }
}

if ($fail -eq 0) { $lines += "ui: all passed" } else { $lines += "ui: FAILURES" }
$lines | Set-Content $log
