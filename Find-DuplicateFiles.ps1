<# 
  Find-DuplicateFiles_v1.3.4.ps1
  PS 5.1-compatible duplicate finder (memory-first results + audit CSVs)
  Key fix: robust enumeration (no -File), emits ScanStats, and always writes CSV headers.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$Path,

  [switch]$Recurse,

  [switch]$IncludeHidden,

  [int]$MinSizeMB = 0,

  [switch]$ConfirmContent,   # full SHA256 confirmation across same-size groups

  [switch]$ConfirmedOnly,    # output only confirmed duplicates (SHA256)

  [ValidateSet('ShortestPath','LongestPath','NewestWriteTime','OldestWriteTime','NewestCreationTime','OldestCreationTime')]
  [string]$KeepRule = 'ShortestPath',

  [string]$ReportPath,

  [string]$SummaryReportPath,

  [int]$SampleRemovePaths = 5
)

Set-StrictMode -Version 2

function New-DefaultOutPath {
  param([string]$Prefix,[string]$Ext)
  $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
  return (Join-Path -Path (Get-Location) -ChildPath ("out\{0}_{1}.{2}" -f $Prefix,$ts,$Ext))
}

function Ensure-DirForFile {
  param([string]$FilePath)
  $dir = Split-Path -Parent $FilePath
  if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
  }
}

function Write-CsvWithHeaders {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string[]]$Headers,
    [Parameter(Mandatory=$true)][object[]]$Rows
  )
  Ensure-DirForFile -FilePath $FilePath
  if ($Rows -and $Rows.Count -gt 0) {
    $Rows | Select-Object $Headers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $FilePath
  } else {
    ($Headers -join ",") | Out-File -Encoding UTF8 -FilePath $FilePath -Force
  }
}

function Get-QuickHash {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [int]$QuickHashBytes = 1048576
  )
  try {
    $fi = Get-Item -LiteralPath $FilePath -ErrorAction Stop
    $len = [int64]$fi.Length
    if ($len -le 0) { return "EMPTY" }

    $firstN = [Math]::Min($QuickHashBytes, $len)
    $lastN  = [Math]::Min($QuickHashBytes, $len)

    $fs = [System.IO.File]::Open($FilePath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
    try {
      $buf1 = New-Object byte[] $firstN
      [void]$fs.Read($buf1,0,$firstN)

      $buf2 = New-Object byte[] $lastN
      $fs.Seek([Math]::Max(0, $len - $lastN), [System.IO.SeekOrigin]::Begin) | Out-Null
      [void]$fs.Read($buf2,0,$lastN)

      $sha = [System.Security.Cryptography.SHA256]::Create()
      try {
        $all = New-Object byte[] ($buf1.Length + 8 + $buf2.Length)
        [Array]::Copy($buf1,0,$all,0,$buf1.Length)
        [Array]::Copy([BitConverter]::GetBytes($len),0,$all,$buf1.Length,8)
        [Array]::Copy($buf2,0,$all,$buf1.Length+8,$buf2.Length)
        $hash = $sha.ComputeHash($all)
        return ([BitConverter]::ToString($hash)).Replace("-","")
      } finally { $sha.Dispose() }
    } finally { $fs.Dispose() }
  } catch {
    return $null
  }
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$errorsSuppressed = 0

# Resolve output paths
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
  $ReportPath = New-DefaultOutPath -Prefix "duplicates_confirmed_keepremove" -Ext "csv"
}
if ([string]::IsNullOrWhiteSpace($SummaryReportPath)) {
  $SummaryReportPath = New-DefaultOutPath -Prefix "duplicates_summary" -Ext "csv"
}

# Robust enumeration (PS5.1 safe): avoid -File; optionally -Force
$gciParams = @{
  Path = $Path
  ErrorAction = 'SilentlyContinue'
}
if ($Recurse) { $gciParams.Recurse = $true }
if ($IncludeHidden) { $gciParams.Force = $true }

$all = @()
try {
  $all = Get-ChildItem @gciParams | Where-Object { -not $_.PSIsContainer }
} catch {
  # If enumeration itself fails unexpectedly, continue with empty set but flag error
  $errorsSuppressed++
  $all = @()
}

$filesEnumerated = if ($all) { $all.Count } else { 0 }

# Apply size filter
$minBytes = [int64]$MinSizeMB * 1024 * 1024
$files = @()
if ($all -and $all.Count -gt 0) {
  $files = $all | Where-Object { $_.Length -ge $minBytes }
}
$filesScanned = if ($files) { $files.Count } else { 0 }

# Prepare outputs
$detailHeaders = @(
  "GroupId","Confidence","Action","KeepRule","KeepPath","KeepReason",
  "SizeBytes","SizeMB","Name","FullPath","QuickHash","FullHash",
  "CreationTime","LastWriteTime"
)

$summaryHeaders = @(
  "GroupId","Confidence","Count","SizeBytes","SizeMB","KeepPath","KeepReason",
  "RemoveCount","PotentialWastedMB","SampleRemovePaths"
)

# If no files at all, emit stats and write empty CSVs with headers.
if ($filesScanned -eq 0) {
  $sw.Stop()
  $stats = [pscustomobject]@{
    _Type = 'ScanStats'
    Path = $Path
    FilesEnumerated = $filesEnumerated
    FilesScanned = $filesScanned
    SizeGroupsGE2 = 0
    ConfirmedGroups = 0
    ResultRows = 0
    ErrorsSuppressed = $errorsSuppressed
    DurationSec = [Math]::Round($sw.Elapsed.TotalSeconds,2)
    ReportPath = $ReportPath
    SummaryReportPath = $SummaryReportPath
  }
  Write-CsvWithHeaders -FilePath $ReportPath -Headers $detailHeaders -Rows @()
  Write-CsvWithHeaders -FilePath $SummaryReportPath -Headers $summaryHeaders -Rows @()
  $stats
  return
}

# Group by size first
$sizeGroups = $files | Group-Object Length | Where-Object { $_.Count -gt 1 }
$sizeGroupsGE2 = if ($sizeGroups) { $sizeGroups.Count } else { 0 }

# Precompute quickhash for probable grouping (optional but helps performance)
$probableRows = @()
$confirmedRows = @()
$allRows = @()

# Build a flat list with quickhash
$flat = New-Object System.Collections.Generic.List[object]
foreach ($g in $sizeGroups) {
  foreach ($f in $g.Group) {
    $qh = Get-QuickHash -FilePath $f.FullName
    $flat.Add([pscustomobject]@{
      Size = [int64]$f.Length
      Name = $f.Name
      FullPath = $f.FullName
      CreationTime = $f.CreationTime
      LastWriteTime = $f.LastWriteTime
      QuickHash = $qh
    })
  }
}

# Probable by Size + QuickHash (only if quickhash available)
$probableGroups = $flat | Where-Object { $_.QuickHash } | Group-Object Size,QuickHash | Where-Object { $_.Count -gt 1 }

foreach ($pg in $probableGroups) {
  $gid = "P-" + ([guid]::NewGuid().ToString("N").Substring(0,10))
  foreach ($item in $pg.Group) {
    $row = [pscustomobject]@{
      GroupId = $gid
      Confidence = "Probable"
      Action = ""
      KeepRule = ""
      KeepPath = ""
      KeepReason = ""
      SizeBytes = $item.Size
      SizeMB = [Math]::Round($item.Size/1MB,2)
      Name = $item.Name
      FullPath = $item.FullPath
      QuickHash = $item.QuickHash
      FullHash = ""
      CreationTime = $item.CreationTime
      LastWriteTime = $item.LastWriteTime
    }
    $allRows += $row
  }
}

# Potential by Size only (showing candidates that share size but no quickhash match)
$potentialGroups = $flat | Group-Object Size | Where-Object { $_.Count -gt 1 }
foreach ($sg in $potentialGroups) {
  # Skip if already covered by probable grouping with same members (keep noise down)
  # We'll still emit Potential for groups that don't have any probable duplicates.
  $hasProbable = $false
  foreach ($m in $sg.Group) {
    if ($m.QuickHash -and ($probableGroups | Where-Object { $_.Name -like "*$($m.Size),$($m.QuickHash)*" })) { $hasProbable = $true; break }
  }
  if ($hasProbable) { continue }

  $gid = "S-" + ([guid]::NewGuid().ToString("N").Substring(0,10))
  foreach ($item in $sg.Group) {
    $row = [pscustomobject]@{
      GroupId = $gid
      Confidence = "Potential"
      Action = ""
      KeepRule = ""
      KeepPath = ""
      KeepReason = ""
      SizeBytes = $item.Size
      SizeMB = [Math]::Round($item.Size/1MB,2)
      Name = $item.Name
      FullPath = $item.FullPath
      QuickHash = $item.QuickHash
      FullHash = ""
      CreationTime = $item.CreationTime
      LastWriteTime = $item.LastWriteTime
    }
    $allRows += $row
  }
}

# Confirmed by Size + SHA256 across ALL same-size groups (critical fix)
$confirmedGroupsCount = 0
if ($ConfirmContent) {
  # For each size-group >=2: hash all files (SHA256) and group by hash.
  foreach ($sg in $potentialGroups) {
    $items = $sg.Group
    if (-not $items -or $items.Count -lt 2) { continue }

    $hashed = New-Object System.Collections.Generic.List[object]
    foreach ($it in $items) {
      try {
        $h = (Get-FileHash -Algorithm SHA256 -LiteralPath $it.FullPath -ErrorAction Stop).Hash
        $hashed.Add([pscustomobject]@{
          Size = $it.Size
          Name = $it.Name
          FullPath = $it.FullPath
          CreationTime = $it.CreationTime
          LastWriteTime = $it.LastWriteTime
          QuickHash = $it.QuickHash
          FullHash = $h
        })
      } catch {
        $errorsSuppressed++
      }
    }

    $dupeByHash = $hashed | Group-Object FullHash | Where-Object { $_.Count -gt 1 }
    foreach ($hg in $dupeByHash) {
      $confirmedGroupsCount++
      $gid = "C-" + ([guid]::NewGuid().ToString("N").Substring(0,10))

      # Choose KEEP
      $keep = $null
      switch ($KeepRule) {
        'ShortestPath'       { $keep = $hg.Group | Sort-Object { $_.FullPath.Length }, FullPath | Select-Object -First 1; $keepReason = "ShortestPath" }
        'LongestPath'        { $keep = $hg.Group | Sort-Object @{Expression={ $_.FullPath.Length }; Descending=$true}, FullPath | Select-Object -First 1; $keepReason = "LongestPath" }
        'NewestWriteTime'    { $keep = $hg.Group | Sort-Object @{Expression='LastWriteTime'; Descending=$true}, FullPath | Select-Object -First 1; $keepReason = "NewestWriteTime" }
        'OldestWriteTime'    { $keep = $hg.Group | Sort-Object LastWriteTime, FullPath | Select-Object -First 1; $keepReason = "OldestWriteTime" }
        'NewestCreationTime' { $keep = $hg.Group | Sort-Object @{Expression='CreationTime'; Descending=$true}, FullPath | Select-Object -First 1; $keepReason = "NewestCreationTime" }
        'OldestCreationTime' { $keep = $hg.Group | Sort-Object CreationTime, FullPath | Select-Object -First 1; $keepReason = "OldestCreationTime" }
      }

      foreach ($item in $hg.Group) {
        $action = if ($keep -and ($item.FullPath -eq $keep.FullPath)) { "KEEP" } else { "REMOVE" }
        $row = [pscustomobject]@{
          GroupId = $gid
          Confidence = "Confirmed"
          Action = $action
          KeepRule = $KeepRule
          KeepPath = if ($keep) { $keep.FullPath } else { "" }
          KeepReason = $keepReason
          SizeBytes = [int64]$item.Size
          SizeMB = [Math]::Round([int64]$item.Size/1MB,2)
          Name = $item.Name
          FullPath = $item.FullPath
          QuickHash = $item.QuickHash
          FullHash = $item.FullHash
          CreationTime = $item.CreationTime
          LastWriteTime = $item.LastWriteTime
        }
        $confirmedRows += $row
      }
    }
  }
}

# Combine rows
if ($ConfirmedOnly) {
  $finalRows = $confirmedRows
} else {
  # Include confirmed too (dedupe by FullPath keeping confirmed if exists)
  $map = @{}
  foreach ($r in $allRows) { $map[$r.FullPath] = $r }
  foreach ($r in $confirmedRows) { $map[$r.FullPath] = $r }
  $finalRows = $map.Values
}

# Build summary rows for confirmed groups
$summaryRows = @()
if ($confirmedRows -and $confirmedRows.Count -gt 0) {
  $confirmedRows | Group-Object GroupId | ForEach-Object {
    $gid = $_.Name
    $grp = $_.Group
    $keepRow = $grp | Where-Object { $_.Action -eq "KEEP" } | Select-Object -First 1
    $remove = $grp | Where-Object { $_.Action -eq "REMOVE" }
    $removeCount = if ($remove) { $remove.Count } else { 0 }
    $sizeBytes = [int64]($grp | Select-Object -First 1).SizeBytes
    $wastedMB = [Math]::Round(($removeCount * $sizeBytes)/1MB,2)
    $samplePaths = ""
    if ($removeCount -gt 0) {
      $samplePaths = ($remove | Select-Object -First $SampleRemovePaths -ExpandProperty FullPath) -join " | "
    }
    $summaryRows += [pscustomobject]@{
      GroupId = $gid
      Confidence = "Confirmed"
      Count = $grp.Count
      SizeBytes = $sizeBytes
      SizeMB = [Math]::Round($sizeBytes/1MB,2)
      KeepPath = if ($keepRow) { $keepRow.KeepPath } else { "" }
      KeepReason = if ($keepRow) { $keepRow.KeepReason } else { "" }
      RemoveCount = $removeCount
      PotentialWastedMB = $wastedMB
      SampleRemovePaths = $samplePaths
    }
  }
}

# Always write audit CSVs (with headers even if empty)
Write-CsvWithHeaders -FilePath $ReportPath -Headers $detailHeaders -Rows ($confirmedRows | Sort-Object GroupId, Action, FullPath)
Write-CsvWithHeaders -FilePath $SummaryReportPath -Headers $summaryHeaders -Rows ($summaryRows | Sort-Object GroupId)

$sw.Stop()

# Emit stats first
[pscustomobject]@{
  _Type = 'ScanStats'
  Path = $Path
  FilesEnumerated = $filesEnumerated
  FilesScanned = $filesScanned
  SizeGroupsGE2 = $sizeGroupsGE2
  ConfirmedGroups = $confirmedGroupsCount
  ResultRows = if ($finalRows) { $finalRows.Count } else { 0 }
  ErrorsSuppressed = $errorsSuppressed
  DurationSec = [Math]::Round($sw.Elapsed.TotalSeconds,2)
  ReportPath = $ReportPath
  SummaryReportPath = $SummaryReportPath
}

# Emit rows for UI binding
$finalRows | Sort-Object Confidence, GroupId, Action, FullPath
