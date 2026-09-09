# The More menu, driven the way a person drives it: open it, click the item, then look at the file.
#
# A script of its own rather than another block in ui.ps1, because a menu popup needs the window the desktop is
# pointing at, and after a long run of window tests it is not that window any more. Run it interactively (a
# scheduled task with /it): a service session has no desktop and so has no menus.
#
# It exists because the engine could add a paragraph a whole release before anything in the window could ask it
# to, and every test was green. What a menu offers is not something a unit test can see.
#   powershell -File tests\ui-shape.ps1
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
$ErrorActionPreference = "Continue"
$log = "C:\Plain\ui-shape-result.txt"
$lines = @(); $fail = 0
function Note($ok, $what, $detail = "") {
  if ($ok) { $script:lines += "ok    $what" } else { $script:lines += "FAIL  $what $detail"; $script:fail = 1 }
}
Stop-Process -Name "Plain for Windows" -Force -ErrorAction SilentlyContinue
Start-Sleep 2
Copy-Item C:\Plain\fixtures\review.docx C:\Plain\shape.docx -Force
$p = Start-Process "C:\Plain\Plain for Windows.exe" -ArgumentList C:\Plain\shape.docx -PassThru
Start-Sleep 8
$root = [System.Windows.Automation.AutomationElement]::RootElement
$cond = New-Object System.Windows.Automation.PropertyCondition(
  [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
$w = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
Note ($null -ne $w) "the document opens"
if ($w) {
  $byId = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::AutomationIdProperty, "MoreBtn")
  $more = $w.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $byId)
  Note ($null -ne $more) "the More button is in the window"
  if ($more) {
    $more.SetFocus(); Start-Sleep 1
    [System.Windows.Forms.SendKeys]::SendWait(" ")
    Start-Sleep 3
    $byName = New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::NameProperty, "Add a paragraph after this one")
    $item = $root.FindFirst([System.Windows.Automation.TreeScope]::Subtree, $byName)
    Note ($null -ne $item) "the menu offers adding a paragraph"
    if ($item) {
      $item.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
      Start-Sleep 2
      [System.Windows.Forms.SendKeys]::SendWait("^s")
      Start-Sleep 4
      $round = (& C:\Plain\plain.exe roundtrip C:\Plain\shape.docx | Out-String)
      Note ($round -match "identical") "the file is whole after the change" $round.Trim()
    }
  }
  Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
}
if ($fail -eq 0) { $lines += "ui-shape: all passed" } else { $lines += "ui-shape: FAILURES" }
$lines | Set-Content $log
