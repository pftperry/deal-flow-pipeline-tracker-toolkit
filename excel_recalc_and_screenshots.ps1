$ErrorActionPreference = "Stop"
$d = $PSScriptRoot; $p = Join-Path $d "Deal_Flow_Pipeline_Tracker_Toolkit.xlsx"; $ev = Join-Path $d "evidence"
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($p); $xl.CalculateFull()
  $total = 0; $ftotal = 0
  foreach ($ws in $wb.Worksheets) {
    $f = 0; $e = 0
    try { $f = $ws.UsedRange.SpecialCells(-4123).Count } catch {}
    try { $e = $ws.UsedRange.SpecialCells(-4123, 16).Count } catch {}
    "{0}: formula cells={1} error cells={2}" -f $ws.Name, $f, $e; $total += $e; $ftotal += $f
  }
  "TOTAL FORMULA CELLS=$ftotal TOTAL ERROR CELLS=$total"
  $pl = $wb.Worksheets.Item("Deal Pipeline Log"); $db = $wb.Worksheets.Item("Ranking Dashboard"); $ft = $wb.Worksheets.Item("Follow-Up Cadence Tracker"); $st = $wb.Worksheets.Item("Settings and Scoring Rules"); $rm = $wb.Worksheets.Item("ReadMe - Validator Instructions")
  foreach ($r in 4..15) { "PL {0} {1}: score={2} rank={3} flag={4}" -f $pl.Cells.Item($r,1).Text, $pl.Cells.Item($r,2).Text, $pl.Cells.Item($r,18).Text, $pl.Cells.Item($r,20).Text, $pl.Cells.Item($r,23).Text }
  foreach ($r in 5..14) { "FT {0} {1}: due={2} flag={3}" -f $ft.Cells.Item($r,1).Text, $ft.Cells.Item($r,2).Text, $ft.Cells.Item($r,6).Text, $ft.Cells.Item($r,9).Text }
  foreach ($r in 4..15) { "DB {0}: {1}" -f $db.Cells.Item($r,1).Text, $db.Cells.Item($r,2).Text }
  foreach ($r in 19..23) { "DB rank {0}: {1} | {2} | {3} | {4} | {5}" -f $db.Cells.Item($r,1).Text, $db.Cells.Item($r,2).Text, $db.Cells.Item($r,3).Text, $db.Cells.Item($r,4).Text, $db.Cells.Item($r,5).Text, $db.Cells.Item($r,6).Text }
  foreach ($r in 41..42) { "DB {0}: {1} ({2}d) {3}" -f $db.Cells.Item($r,1).Text, $db.Cells.Item($r,2).Text, $db.Cells.Item($r,3).Text, $db.Cells.Item($r,4).Text }
  "DB37 total deals=" + $db.Range("B37").Text + " weighted=" + $db.Range("E37").Text
  "DASHBOARD OUTPUT: " + $db.Range("B44").Value2
  "Settings C16: " + $st.Range("C16").Text
  foreach ($r in 27..36) { "RM chk {0}: {1}" -f $rm.Cells.Item($r,2).Text, $rm.Cells.Item($r,5).Text }
  "PL R4 formula: " + $pl.Range("R4").Formula; "PL W4 formula: " + $pl.Range("W4").Formula; "DB B15 formula: " + $db.Range("B15").Formula
  $wb.Save()
  function Shot($ws, $rng, $file) {
    $ws.Activate(); $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
    $r = $ws.Range($rng); $r.CopyPicture(1, 2); Start-Sleep -Milliseconds 400
    $co = $ws.ChartObjects().Add(0, 0, $r.Width, $r.Height); $co.Activate() | Out-Null; $co.Chart.Paste(); Start-Sleep -Milliseconds 400
    $co.Chart.Export((Join-Path $ev $file)) | Out-Null; $co.Delete(); "exported $file"
  }
  Shot $db "A3:F45" "ranking_dashboard.png"
  Shot $pl "A3:L15" "pipeline_inputs.png"
  Shot $pl "M3:X15" "pipeline_formulas_outputs.png"
  Shot $ft "A3:K14" "followup_cadence_tracker.png"
  Shot $wb.Worksheets.Item("Anonymized Worked Example") "A3:G31" "worked_example.png"
  Shot $rm "A25:E43" "readme_validator_checklist.png"
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
