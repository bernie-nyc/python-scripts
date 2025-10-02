<# 
Reorder-LegacyId.ps1
--------------------
Goal: Rename folders like "12302729 - Darr, Courtney" to "Darr, Courtney - 12302729"

What this does (step-by-step):
1) Finds folders in a target directory (optionally, also in its subfolders).
2) Reads the folder name. It must look like: 8 digits, space, dash, space, then a name.
   Example: "12302729 - Darr, Courtney"
3) Moves the first 10 characters ("12302729 -") to the END of the name,
   but also puts the 9th and 10th characters (" -") IN FRONT of the 8 digits.
   So the tail becomes " - 12302729".
4) Final result: "Darr, Courtney - 12302729"
5) By default it's a DRY RUN (it only prints what it *would* do). 
   Add -DoIt to actually rename.

Usage examples:
  .\Reorder-LegacyId.ps1 -Path "U:\General Portfolio\No_Match"
  .\Reorder-LegacyId.ps1 -Path "U:\General Portfolio\No_Match" -Recurse
  .\Reorder-LegacyId.ps1 -Path "U:\General Portfolio\No_Match" -DoIt
#>

param(
  # The folder that contains all the student folders to rename
  [Parameter(Mandatory=$true)]
  [string]$Path,

  # Also process subfolders (off by default)
  [switch]$Recurse,

  # Actually rename things (off by default = dry run)
  [switch]$DoIt
)

# Helper: build the new folder name from the old one.
function Get-NewName {
  param([string]$OldName)

  # Match: 8 digits, then " - ", then the rest of the text (the person's name, etc.)
  # (?<ID>) and (?<Rest>) give us easy-to-read variable names for the matches.
  if ($OldName -match '^(?<ID>\d{8})\s*-\s*(?<Rest>.+)$') {
    $id   = $matches['ID']           # "12302729"
    $rest = $matches['Rest'].Trim()  # "Darr, Courtney" (trim removes extra spaces)

    # The instructions say:
    # - Move the first 10 chars to the end,
    # - Put the 9th and 10th chars (" -") in front of the 8-digit ID.
    # That yields: " - 12302729"
    $tail = " - $id"

    # New format: "<Last, First> - <ID>"
    return "$rest$tail"
  }

  # If it doesn't match the expected pattern, tell the caller by returning $null.
  return $null
}

# Get the target folders to process.
$dirs = Get-ChildItem -Path $Path -Directory -ErrorAction Stop -Recurse:$Recurse

foreach ($d in $dirs) {
  $oldName = $d.Name
  $newName = Get-NewName -OldName $oldName

  if (-not $newName) {
    Write-Host "[SKIP] $oldName   (name not in expected pattern)" -ForegroundColor Yellow
    continue
  }

  if ($newName -eq $oldName) {
    Write-Host "[SKIP] $oldName   (already correct)" -ForegroundColor DarkYellow
    continue
  }

  # Where the renamed folder will live (same parent, new name)
  $targetPath = Join-Path -Path $d.Parent.FullName -ChildPath $newName

  # If another folder already has that name, add a small suffix to avoid collisions.
  $suffix = 2
  while (Test-Path -LiteralPath $targetPath) {
    $altName    = "$newName ($suffix)"
    $targetPath = Join-Path -Path $d.Parent.FullName -ChildPath $altName
    $suffix++
  }

  if ($DoIt) {
    try {
      Rename-Item -LiteralPath $d.FullName -NewName ([IO.Path]::GetFileName($targetPath)) -ErrorAction Stop
      Write-Host "[OK]   $oldName -> $([IO.Path]::GetFileName($targetPath))" -ForegroundColor Green
    }
    catch {
      Write-Host "[ERR]  $oldName -> $([IO.Path]::GetFileName($targetPath)) : $($_.Exception.Message)" -ForegroundColor Red
    }
  } else {
    # Dry run (preview only)
    Write-Host "[DRY]  $oldName -> $([IO.Path]::GetFileName($targetPath))"
  }
}
