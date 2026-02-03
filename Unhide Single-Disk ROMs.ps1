<#
PURPOSE: Review and optionally unhide single-disk hidden entries in Batocera gamelist.xml files
VERSION: 1.0
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process multiple platforms, OR into a platform's individual folder to process just that one.
- This script DOES modify gamelist.xml when you choose to unhide entries.
- It searches for entries marked hidden via: <hidden>true</hidden> (also accepts 1/yes, case-insensitive).
- For each hidden entry found, it prompts you to:
    - Unhide (remove the <hidden> node entirely)
    - Skip (leave hidden)
    - Cancel (abort the script safely; any changes already saved remain saved)

IMPORTANT MULTI-DISK SAFETY (robust bypass):
- Hidden entries that appear to be "Disk 2+" (or equivalent) members of a multi-disk set are automatically BYPASSED.
- This is done via stable grouping logic consistent with the Export Game List methodology:
    - GroupKey primary: <name> when present
    - GroupKey fallback: <path> when <name> is missing/blank
- Within each group, a single "primary candidate" is chosen (Disk 1 / representative).
  Only that primary candidate is eligible for prompting if it is hidden.
  All other hidden siblings in the same group are treated as Disk 2+ and remain hidden (no prompt).

ROBUSTNESS:
- If a gamelist.xml is malformed and cannot be parsed as XML, the script will skip it (safe behavior; no edits).
- Before saving any edits for a given gamelist.xml, the script creates a timestamped .bak backup next to the file.

DISCOVERY / SCOPE:
- Determines runtime mode based on where the script is located:
    - If a gamelist.xml exists in the script directory, treat it as a single-platform run.
    - Otherwise, treat the script directory as ROMS root and scan up to 2 folder levels deep for gamelist.xml files.
- In ROMS root mode, prints per-platform start + finished lines.
- Always prints a final summary and runtime.

UI:
- Uses a top-most WinForms MessageBox (Yes/No/Cancel) when running in STA + WinForms available.
- Falls back to console prompts otherwise.

#>

# ==================================================================================================
# SCRIPT STARTUP: STRICT MODE, PATHS, RUNTIME
# ==================================================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$startDir  = $scriptDir

# --------------------------------------------------------------------------------------------------
# Runtime tracking
# PURPOSE:
# - Provide a consistent runtime report:
#     - <60s  => "X seconds"
#     - <60m  => "M:SS"
#     - >=60m => "H:MM:SS"
# --------------------------------------------------------------------------------------------------
$__runtimeStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function Format-ElapsedRuntime {
  param([Parameter(Mandatory=$true)][TimeSpan]$Elapsed)

  if ($Elapsed.TotalSeconds -lt 60) {
    $sec = [int][math]::Floor($Elapsed.TotalSeconds)
    return ("{0} seconds" -f $sec)
  }

  if ($Elapsed.TotalMinutes -lt 60) {
    $min = [int][math]::Floor($Elapsed.TotalMinutes)
    $sec = [int]$Elapsed.Seconds
    return ("{0}:{1:00}" -f $min, $sec)
  }

  $hrs = [int][math]::Floor($Elapsed.TotalHours)
  $min = [int]$Elapsed.Minutes
  $sec = [int]$Elapsed.Seconds
  return ("{0}:{1:00}:{2:00}" -f $hrs, $min, $sec)
}

function Write-RuntimeReport {
  param([switch]$Stop)

  if ($null -eq $__runtimeStopwatch) { return }

  try {
    if ($Stop -and $__runtimeStopwatch.IsRunning) { $__runtimeStopwatch.Stop() }
    $rt = Format-ElapsedRuntime -Elapsed $__runtimeStopwatch.Elapsed
    Write-Host ("Runtime: {0}" -f $rt) -ForegroundColor Gray
  } catch {
    # Intentionally ignore runtime reporting failures
  }
}

# --------------------------------------------------------------------------------------------------
# Phase output helper
# --------------------------------------------------------------------------------------------------
function Write-Phase {
  param([string]$Message)
  Write-Host ""
  Write-Host $Message -ForegroundColor Cyan
}

# ==================================================================================================
# UI HELPERS (TOPMOST MESSAGEBOX WHEN AVAILABLE; CONSOLE FALLBACK)
# ==================================================================================================

$script:WinFormsOk = $false

function Test-IsSTA {
  try { return [System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA' }
  catch { return $false }
}

function Ensure-WinForms {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing       | Out-Null
  try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

function Can-UseGui {
  return ((Test-IsSTA) -and $script:WinFormsOk)
}

function Show-TopMostMessageBox {
  <#
    Shows a message box with a hidden top-most owner form so it stays in front of editors/ISE.
  #>
  param(
    [Parameter(Mandatory=$true)][string]$Text,
    [Parameter(Mandatory=$true)][string]$Title,
    [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
    [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Question
  )

  $owner = New-Object System.Windows.Forms.Form
  $owner.TopMost = $true
  $owner.StartPosition = 'Manual'
  $owner.Size = New-Object System.Drawing.Size(1,1)
  $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
  $owner.ShowInTaskbar = $false
  $owner.Show() | Out-Null

  try {
    return [System.Windows.Forms.MessageBox]::Show($owner, $Text, $Title, $Buttons, $Icon)
  } finally {
    $owner.Close()
    $owner.Dispose()
  }
}

# Load WinForms when possible (do NOT swallow and then attempt GUI types)
try {
  Ensure-WinForms
  $script:WinFormsOk = $true
} catch {
  $script:WinFormsOk = $false
}

# ==================================================================================================
# DISCOVERY: DETERMINE MODE AND FIND TARGET gamelist.xml FILES (UP TO 2 LEVELS DEEP)
# ==================================================================================================

$localGamelistPath     = Join-Path $startDir 'gamelist.xml'
$isSinglePlatformMode  = (Test-Path -LiteralPath $localGamelistPath)
$isRomsRootMode        = (-not $isSinglePlatformMode)

Write-Phase "Starting hidden-entry review..."

function Get-PlatformTargets {
  <#
    PURPOSE:
    - Determine which gamelist.xml file(s) to process:
        - Single-platform mode: script folder contains gamelist.xml
        - ROMS root mode: scan up to 2 levels deep for gamelist.xml
  #>
  param([string]$StartDir)

  $start = (Resolve-Path -LiteralPath $StartDir).Path
  $local = Join-Path $start 'gamelist.xml'

  if (Test-Path -LiteralPath $local) {
    return @([pscustomobject]@{
      PlatformFolder = (Split-Path -Leaf $start)
      PlatformRoot   = $start
      GamelistPath   = $local
      Depth          = 0
    })
  }

  # Folder ignore list (ROMS root mode only)
  $ignoredFolders = @(
    'windows_installers'
  )

  $targets = @()

  # Scan depth 1 + depth 2:
  $dirsDepth1 = @(Get-ChildItem -LiteralPath $start -Directory -ErrorAction Stop)
  foreach ($d1 in $dirsDepth1) {

    if ($ignoredFolders -contains $d1.Name.ToLowerInvariant()) { continue }

    $g1 = Join-Path $d1.FullName 'gamelist.xml'
    if (Test-Path -LiteralPath $g1) {
      $targets += [pscustomobject]@{
        PlatformFolder = $d1.Name
        PlatformRoot   = $d1.FullName
        GamelistPath   = $g1
        Depth          = 1
      }
      continue
    }

    # Depth 2 scan (subfolders under a platform-ish folder)
    $dirsDepth2 = @()
    try { $dirsDepth2 = @(Get-ChildItem -LiteralPath $d1.FullName -Directory -ErrorAction Stop) } catch { $dirsDepth2 = @() }

    foreach ($d2 in $dirsDepth2) {
      $g2 = Join-Path $d2.FullName 'gamelist.xml'
      if (Test-Path -LiteralPath $g2) {
        $targets += [pscustomobject]@{
          PlatformFolder = ($d1.Name + "\" + $d2.Name)
          PlatformRoot   = $d2.FullName
          GamelistPath   = $g2
          Depth          = 2
        }
      }
    }
  }

  return $targets
}

if ($isSinglePlatformMode) {
  Write-Host ("MODE: Single-platform ({0})" -f (Split-Path -Leaf $startDir)) -ForegroundColor Green
} else {
  Write-Host "MODE: ROMS root (discovering gamelist.xml up to 2 levels deep...)" -ForegroundColor Green
}

Write-Phase "Discovering gamelist.xml files..."

$targets = @(Get-PlatformTargets -StartDir $startDir)

if (@($targets).Count -eq 0) {
  Write-Warning "No gamelist.xml found. Run from /roms or a folder that contains gamelist.xml."
  Write-RuntimeReport -Stop
  return
}

Write-Host "Found $($targets.Count) gamelist.xml file(s) to process." -ForegroundColor Green

# ==================================================================================================
# XML + GROUPING HELPERS (EXPORT-LIST CONSISTENT GROUP KEY + MULTI-DISK PRIMARY SELECTION)
# ==================================================================================================

function Get-XmlNodeText {
  <#
    Safely read an XML child node's InnerText without throwing when missing.
    Uses local-name() so namespaces don't break child selection.
  #>
  param(
    [Parameter(Mandatory=$true)][System.Xml.XmlNode]$Node,
    [Parameter(Mandatory=$true)][string]$ChildName
  )

  if ($null -eq $Node) { return '' }
  if ([string]::IsNullOrWhiteSpace($ChildName)) { return '' }

  $child = $null
  try { $child = $Node.SelectSingleNode("*[local-name()='$ChildName']") } catch { $child = $null }
  if ($null -eq $child) { return '' }

  $text = ''
  try { $text = [string]$child.InnerText } catch { $text = '' }
  return $text.Trim()
}

function Is-HiddenTrue {
  <#
    Determines if a <hidden> value should be treated as true.
    Accepts: true / 1 / yes (case-insensitive)
  #>
  param([AllowNull()][AllowEmptyString()][string]$HiddenText)

  $h = ([string]$HiddenText).Trim()
  if ([string]::IsNullOrWhiteSpace($h)) { return $false }
  return ($h -match '^(?i)(true|1|yes)$')
}

function Backup-FileOnce {
  <#
    Creates a timestamped .bak backup next to a file before edits.
  #>
  param([Parameter(Mandatory=$true)][string]$FilePath)

  $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $bak = $FilePath + "." + $ts + ".bak"
  Copy-Item -LiteralPath $FilePath -Destination $bak -Force
  return $bak
}

function Get-GroupKey {
  <#
    PURPOSE:
    - Use the same stable grouping key approach as the export list script:
        - Primary: raw <name> when present
        - Fallback: <path> when name missing/blank (prevents unrelated entries collapsing)
  #>
  param(
    [AllowNull()][AllowEmptyString()][string]$NameRaw,
    [AllowNull()][AllowEmptyString()][string]$PathRaw
  )

  $n = [string]$NameRaw
  if (-not [string]::IsNullOrWhiteSpace($n)) { return $n.Trim() }

  return [string]$PathRaw
}

function Get-DiskIndexFromPath {
  <#
    PURPOSE:
    - Try to extract a disk/disc/side index from a filename/path string.
    NOTES:
    - Returns $null when no clear index is found.
    - Handles patterns like:
        (Disk 1), (Disk 2 of 4), Disk 3, Disc 2, CD1/CD2, etc.
        Side A / Side B (A=1, B=2, etc.)
  #>
  param([AllowNull()][AllowEmptyString()][string]$PathRaw)

  $p = [string]$PathRaw
  if ([string]::IsNullOrWhiteSpace($p)) { return $null }

  $file = $p
  try { $file = [System.IO.Path]::GetFileName($p) } catch {}

  # Disk/Disc with number
  $m = [regex]::Match($file, '(?i)\b(?:disc|disk)\s*0*([1-9]\d*)\b')
  if ($m.Success) { return [int]$m.Groups[1].Value }

  # "Disk 1 of 3" / "Disc 2 of 4"
  $m = [regex]::Match($file, '(?i)\b(?:disc|disk)\s*0*([1-9]\d*)\s*of\s*([1-9]\d*)\b')
  if ($m.Success) { return [int]$m.Groups[1].Value }

  # CD1 / CD2
  $m = [regex]::Match($file, '(?i)\bcd\s*0*([1-9]\d*)\b')
  if ($m.Success) { return [int]$m.Groups[1].Value }

  # "Side A/B/C" (map A=1, B=2, C=3...)
  $m = [regex]::Match($file, '(?i)\bside\s*([A-Z])\b')
  if ($m.Success) {
    $ch = [char]$m.Groups[1].Value.ToUpperInvariant()
    $idx = ([int]$ch) - ([int][char]'A') + 1
    if ($idx -ge 1 -and $idx -le 26) { return [int]$idx }
  }

  return $null
}

function Select-PrimaryCandidate {
  <#
    PURPOSE:
    - For a group of items representing one "set", pick a single "primary candidate" that represents Disk 1 / main entry.
    - Only this candidate is eligible to be unhidden when the group has multiple members.
  #>
  param([Parameter(Mandatory=$true)]$Items)

  $items = @($Items)
  if ($items.Count -lt 1) { return $null }

  # If an .m3u exists in this group, treat it as primary (prefer visible one).
  $m3u = @($items | Where-Object { ([string]$_.PathRaw) -match '(?i)\.m3u$' })
  if ($m3u.Count -gt 0) {
    $m3uVisible = @($m3u | Where-Object { -not $_.Hidden })
    if ($m3uVisible.Count -gt 0) {
      return @($m3uVisible | Sort-Object PathRaw | Select-Object -First 1)[0]
    }
    return @($m3u | Sort-Object PathRaw | Select-Object -First 1)[0]
  }

  # Prefer an item explicitly marked Disk/Disc/Side 1
  $withDisk = foreach ($it in $items) {
    $di = Get-DiskIndexFromPath -PathRaw ([string]$it.PathRaw)
    [pscustomobject]@{ Item = $it; DiskIndex = $di }
  }

  $disk1 = @($withDisk | Where-Object { $_.DiskIndex -eq 1 })
  if ($disk1.Count -gt 0) {
    return @($disk1 | Sort-Object { $_.Item.PathRaw } | Select-Object -First 1).Item
  }

  # Next best: any visible item in the group (common case: Disk 1 is visible)
  $visible = @($items | Where-Object { -not $_.Hidden })
  if ($visible.Count -gt 0) {
    return @($visible | Sort-Object PathRaw | Select-Object -First 1)[0]
  }

  # All hidden: fall back to stable representative (same pattern as export list: first by path)
  return @($items | Sort-Object PathRaw | Select-Object -First 1)[0]
}

# ==================================================================================================
# BYPASS REASON TRACKING (PER-PLATFORM BREAKOUT)
# ==================================================================================================

function Add-BypassReason {
  <#
    PURPOSE:
    - Track why hidden entries were bypassed (never prompted) in a per-platform breakdown.
    NOTES:
    - Does not affect prompt behavior; it only counts reporting.
  #>
  param(
    [Parameter(Mandatory=$true)][hashtable]$BypassCounts,
    [Parameter(Mandatory=$true)][string]$Reason
  )

  $k = [string]$Reason
  if ([string]::IsNullOrWhiteSpace($k)) { $k = '(unspecified)' }

  if ($BypassCounts.ContainsKey($k)) { $BypassCounts[$k] = [int]$BypassCounts[$k] + 1 }
  else { $BypassCounts[$k] = 1 }
}

function Format-BypassBreakdown {
  <#
    PURPOSE:
    - Render bypass reason counts into a compact string for the per-platform summary line.
    OUTPUT EXAMPLE:
    - "Disk2+=12, bios/firmware=3"
  #>
  param([Parameter(Mandatory=$true)][hashtable]$BypassCounts)

  if ($null -eq $BypassCounts -or $BypassCounts.Count -lt 1) { return '' }

  $parts = foreach ($k in ($BypassCounts.Keys | Sort-Object)) {
    ("{0}={1}" -f $k, [int]$BypassCounts[$k])
  }

  return ($parts -join ', ')
}

# ==================================================================================================
# MAIN: PROCESS EACH gamelist.xml, PROMPT PER ELIGIBLE HIDDEN ENTRY, SAVE IF MODIFIED
# ==================================================================================================

Write-Phase "Scanning for hidden entries..."

$totalHiddenFound    = 0
$totalBypassed       = 0
$totalUnhidden       = 0
$totalSkipped        = 0
$totalMalformed      = 0
$totalPlatformsModified = 0

foreach ($t in $targets) {

  # Per-platform counters (mini-summary)
  $platHiddenFound = 0
  $platBypassed    = 0
  $platUnhidden    = 0
  $platSkipped     = 0

  # Per-platform bypass reason tracking (items we will NOT prompt to unhide)
  $bypassCounts = @{}

  Write-Host ""
  Write-Host ("Platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Green
  Write-Host ("gamelist.xml: {0}" -f [string]$t.GamelistPath) -ForegroundColor Gray

  $gamelistPath = [string]$t.GamelistPath

  # Parse XML safely. If malformed, skip (do not attempt edits).
  $doc = $null
  $parsedOk = $false
  try {
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.PreserveWhitespace = $true
    $doc.Load($gamelistPath)
    $parsedOk = $true
  } catch {
    $parsedOk = $false
  }

  if (-not $parsedOk -or $null -eq $doc) {
    Write-Host "Malformed XML (skipping safely; no edits)." -ForegroundColor Yellow
    $totalMalformed++

    Write-Host ("Summary: Malformed (skipped)") -ForegroundColor Yellow
    Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan
    continue
  }

  # Find all <game> nodes
  $gameNodes = $null
  try { $gameNodes = $doc.SelectNodes("//*[local-name()='game']") } catch { $gameNodes = $null }

  if ($null -eq $gameNodes -or $gameNodes.Count -eq 0) {
    Write-Host "No <game> nodes found (skipping)." -ForegroundColor Yellow
    Write-Host ("Summary: Found 0 hidden, bypassed 0, unhid 0, skipped 0") -ForegroundColor Gray
    Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan
    continue
  }

  # Flatten entries with export-consistent grouping fields
  $entries = @()
  foreach ($gn in $gameNodes) {
    $name   = Get-XmlNodeText -Node $gn -ChildName 'name'
    $path   = Get-XmlNodeText -Node $gn -ChildName 'path'
    $hidden = Get-XmlNodeText -Node $gn -ChildName 'hidden'

    $entries += [pscustomobject]@{
      Node     = $gn
      NameRaw  = [string]$name
      PathRaw  = [string]$path
      Hidden   = (Is-HiddenTrue $hidden)
      GroupKey = (Get-GroupKey -NameRaw $name -PathRaw $path)
    }
  }

  # Count total hidden entries found for platform (raw)
  $platHiddenFound = @($entries | Where-Object { $_.Hidden }).Count
  $totalHiddenFound += $platHiddenFound

  if ($platHiddenFound -eq 0) {
    Write-Host "No hidden entries found." -ForegroundColor Cyan
    Write-Host ("Summary: Found 0 hidden, bypassed 0, unhid 0, skipped 0") -ForegroundColor Gray
    Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan
    continue
  }

  $entryLabel = if ($platHiddenFound -eq 1) { "entry" } else { "entries" }
  Write-Host ("Found {0} hidden {1} (evaluating multi-disk bypass rules)..." -f $platHiddenFound, $entryLabel) -ForegroundColor Cyan

  # Build list of hidden entries eligible for prompting (bypass disk2+ in multi-disk groups)
  $promptList = New-Object System.Collections.Generic.List[object]

  foreach ($group in @($entries | Group-Object GroupKey)) {

    $items = @($group.Group)
    if ($items.Count -lt 1) { continue }

    $primary = Select-PrimaryCandidate -Items $items

    # In a multi-item group, only the primary candidate is eligible (if hidden).
    if ($items.Count -gt 1) {
      foreach ($it in $items) {
        if (-not $it.Hidden) { continue }

        if ($null -ne $primary -and [object]::ReferenceEquals($it, $primary)) {
          $promptList.Add($it) | Out-Null
        } else {
          $platBypassed++
          $totalBypassed++
          Add-BypassReason -BypassCounts $bypassCounts -Reason 'Disk2+'
        }
      }
      continue
    }

    # Single-item group: if hidden, it is eligible.
    if ($items.Count -eq 1 -and $items[0].Hidden) {
      $promptList.Add($items[0]) | Out-Null
    }
  }

  # If everything hidden was bypassed, we won't prompt.
  if ($promptList.Count -eq 0) {
    Write-Host "All hidden entries were bypassed as Disk 2+ members of multi-disk sets." -ForegroundColor Gray

    $bypassText = Format-BypassBreakdown -BypassCounts $bypassCounts
    if (-not [string]::IsNullOrWhiteSpace($bypassText)) {
      Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid 0, skipped 0, bypassed ({2})" -f $platHiddenFound, $platBypassed, $bypassText) -ForegroundColor Gray
    } else {
      Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid 0, skipped 0" -f $platHiddenFound, $platBypassed) -ForegroundColor Gray
    }

    Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan
    continue
  }

  $modified = $false
  $bakPath = $null

  for ($i = 0; $i -lt $promptList.Count; $i++) {

    $hg = $promptList[$i]
    $idx = $i + 1
    $max = $promptList.Count

    $displayName = [string]$hg.NameRaw
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = "(no name)" }

    $displayPath = [string]$hg.PathRaw
    if ([string]::IsNullOrWhiteSpace($displayPath)) { $displayPath = "(no path)" }

    $promptTitle = "Unhide entry? ($idx of $max)"
    $promptText  = "Name: $displayName`r`nPath: $displayPath`r`n`r`nChoose:`r`nYes = Unhide`r`nNo = Skip`r`nCancel = Abort script"

    $decision = $null

    if (Can-UseGui) {
      $res = Show-TopMostMessageBox -Text $promptText -Title $promptTitle `
        -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNoCancel) `
        -Icon ([System.Windows.Forms.MessageBoxIcon]::Question)

      if ($res -eq [System.Windows.Forms.DialogResult]::Yes)    { $decision = 'Yes' }
      elseif ($res -eq [System.Windows.Forms.DialogResult]::No) { $decision = 'No' }
      else { $decision = 'Cancel' }
    }
    else {
      Write-Host ""
      Write-Host "Hidden entry ($idx of $max):" -ForegroundColor Cyan
      Write-Host ("  Name: {0}" -f $displayName) -ForegroundColor White
      Write-Host ("  Path: {0}" -f $displayPath) -ForegroundColor White
      $ans = Read-Host "Unhide? (Y=Yes, N=No, Q=Cancel)"
      if ($ans -match '^(?i)y') { $decision = 'Yes' }
      elseif ($ans -match '^(?i)n') { $decision = 'No' }
      else { $decision = 'Cancel' }
    }

    if ($decision -eq 'Cancel') {
      Write-Host ""
      Write-Host "Cancelled by user." -ForegroundColor Yellow

      # If we already modified in-memory but not saved yet, save safely before exit.
      if ($modified) {
        try {
          if ($null -eq $bakPath) { $bakPath = Backup-FileOnce -FilePath $gamelistPath }
          $doc.Save($gamelistPath)
          Write-Host ("Saved changes before exit. Backup: {0}" -f $bakPath) -ForegroundColor Green
          $totalPlatformsModified++
        } catch {
          Write-Host "Failed to save changes before exit." -ForegroundColor Red
        }
      }

      $bypassText = Format-BypassBreakdown -BypassCounts $bypassCounts
      if (-not [string]::IsNullOrWhiteSpace($bypassText)) {
        Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid {2}, skipped {3}, bypassed ({4})" -f $platHiddenFound, $platBypassed, $platUnhidden, $platSkipped, $bypassText) -ForegroundColor Gray
      } else {
        Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid {2}, skipped {3}" -f $platHiddenFound, $platBypassed, $platUnhidden, $platSkipped) -ForegroundColor Gray
      }

      Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan

      Write-RuntimeReport -Stop
      return
    }

    if ($decision -eq 'No') {
      $platSkipped++
      $totalSkipped++
      continue
    }

    # YES: remove the <hidden> node entirely
    try {
      $hn = $null
      try { $hn = $hg.Node.SelectSingleNode("*[local-name()='hidden']") } catch { $hn = $null }

      if ($null -ne $hn) {
        $modified = $true
        $platUnhidden++
        $totalUnhidden++

        [void]$hg.Node.RemoveChild($hn)
      } else {
        $platSkipped++
        $totalSkipped++
      }
    } catch {
      $platSkipped++
      $totalSkipped++
    }
  }

  if ($modified) {
    try {
      if ($null -eq $bakPath) { $bakPath = Backup-FileOnce -FilePath $gamelistPath }
      $doc.Save($gamelistPath)
      Write-Host ("Saved updated gamelist.xml. Backup: {0}" -f $bakPath) -ForegroundColor Green
      $totalPlatformsModified++
    } catch {
      Write-Host "Failed to save updated gamelist.xml (no changes written)." -ForegroundColor Red
    }
  } else {
    Write-Host "No changes selected for this platform." -ForegroundColor Gray
  }

  # Per-platform mini-summary + finished line
  $bypassText = Format-BypassBreakdown -BypassCounts $bypassCounts
  if (-not [string]::IsNullOrWhiteSpace($bypassText)) {
    Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid {2}, skipped {3}, bypassed ({4})" -f $platHiddenFound, $platBypassed, $platUnhidden, $platSkipped, $bypassText) -ForegroundColor Gray
  } else {
    Write-Host ("Summary: Found {0} hidden, bypassed {1}, unhid {2}, skipped {3}" -f $platHiddenFound, $platBypassed, $platUnhidden, $platSkipped) -ForegroundColor Gray
  }

  Write-Host ("Finished platform: {0}" -f [string]$t.PlatformFolder) -ForegroundColor Cyan
}

# ==================================================================================================
# FINAL SUMMARY
# ==================================================================================================

Write-Phase "Finished."
Write-Host ("Hidden entries found:  {0}" -f $totalHiddenFound) -ForegroundColor Green
Write-Host ("Bypassed (Disk 2+):    {0}" -f $totalBypassed)    -ForegroundColor Green
Write-Host ("Unhidden:              {0}" -f $totalUnhidden)     -ForegroundColor Green
Write-Host ("Skipped:               {0}" -f $totalSkipped)      -ForegroundColor Green

if ($totalMalformed -gt 0) {
  Write-Host ("Malformed skipped:     {0}" -f $totalMalformed) -ForegroundColor Yellow
}

Write-Host ("Platforms modified:    {0}" -f $totalPlatformsModified) -ForegroundColor Green
Write-RuntimeReport -Stop