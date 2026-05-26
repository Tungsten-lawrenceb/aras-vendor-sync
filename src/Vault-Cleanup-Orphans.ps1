<#
.SYNOPSIS
    Quarantine + delete orphan File items.

.DESCRIPTION
    Identifies File rows with no current reference from any expected
    holder (Manufacturer Part File, Document File, CAD's file slots,
    DatasheetFetchRequest.file_id). For each one:

      1. Moves the blob directory to a dated quarantine path under
         <vault_root>.quarantine\<run-timestamp>\<sharded-id>.
      2. Deletes the FILE row and the FILECONTAINERLOCATOR row.

    NO hard-delete of bytes — the quarantine move is reversible. After
    a verification period, that folder can be removed manually.
#>

[CmdletBinding()]
param(
    [string]$Server      = 'localhost\SQLEXPRESS',
    [string]$Database    = 'InnovatorSolutions',
    [string]$DbUser      = 'innovator',
    [string]$DbPassword  = 'ArasDB-2025!',
    [string]$VaultRoot   = 'C:\Aras\Vault\InnovatorSolutions',
    [string]$QuarantineRoot = 'C:\Aras\Vault\InnovatorSolutions.quarantine',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$ts = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$runDir = Join-Path $QuarantineRoot $ts

if (-not (Test-Path $VaultRoot)) { throw "VaultRoot not found: $VaultRoot" }

# --- Pull orphan ids from SQL ---
$cs = "Server=$Server;Database=$Database;User ID=$DbUser;Password=$DbPassword;TrustServerCertificate=True"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()
try {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = @"
SELECT f.id, f.filename, f.file_size
FROM innovator.[FILE] f
WHERE f.is_current = '1'
  AND NOT EXISTS (SELECT 1 FROM innovator.[MANUFACTURER_PART_FILE] r WHERE r.related_id = f.id AND r.is_current='1')
  AND NOT EXISTS (SELECT 1 FROM innovator.[DOCUMENT_FILE] r WHERE r.related_id = f.id AND r.is_current='1')
  AND NOT EXISTS (
      SELECT 1 FROM innovator.[CAD] c
      WHERE c.is_current='1'
        AND (c.native_file = f.id OR c.viewable_file = f.id
          OR c.monolithic_model_file = f.id OR c.view_file = f.id)
  )
  AND NOT EXISTS (SELECT 1 FROM innovator.[CADFILES] cf WHERE cf.is_current='1' AND cf.attached_file = f.id)
  AND NOT EXISTS (SELECT 1 FROM innovator.[DATASHEETFETCHREQUEST] d WHERE d.is_current='1' AND d.file_id = f.id);
"@
    $reader = $cmd.ExecuteReader()
    $orphans = New-Object 'System.Collections.Generic.List[object]'
    while ($reader.Read()) {
        $orphans.Add([pscustomobject]@{
            id        = [string]$reader['id']
            filename  = [string]$reader['filename']
            file_size = [int64]$reader['file_size']
        })
    }
    $reader.Close()
}
finally {
    $conn.Close(); $conn.Dispose()
}

Write-Host "Orphan File rows: $($orphans.Count)"
$totalBytes = ($orphans | Measure-Object -Property file_size -Sum).Sum
Write-Host ("Total bytes: {0:N0} ({1:N1} MB)" -f $totalBytes, ($totalBytes / 1MB))

if ($orphans.Count -eq 0) { Write-Host "Nothing to do."; exit 0 }

if (-not $DryRun) {
    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
    Write-Host "Quarantine dir: $runDir"
}

$moved = 0; $missing = 0; $errors = 0
foreach ($o in $orphans) {
    $id = $o.id
    if ($id.Length -ne 32) { Write-Host "  SKIP non-32 id: $id"; $errors++; continue }
    $c0 = $id.Substring(0, 1)
    $c12 = $id.Substring(1, 2)
    $rest = $id.Substring(3)
    $src = Join-Path -Path (Join-Path -Path (Join-Path -Path $VaultRoot -ChildPath $c0) -ChildPath $c12) -ChildPath $rest
    if (-not (Test-Path $src)) {
        Write-Host "  MISSING blob dir: $src (id=$id filename=$($o.filename))"
        $missing++
        continue
    }
    if ($DryRun) {
        Write-Host "  would move $src"
        continue
    }
    $dst = Join-Path $runDir "$c0\$c12\$rest"
    $dstParent = Split-Path $dst -Parent
    if (-not (Test-Path $dstParent)) { New-Item -ItemType Directory -Path $dstParent -Force | Out-Null }
    try {
        Move-Item -Path $src -Destination $dst -Force
        $moved++
    } catch {
        Write-Host "  ERROR moving $src -> $dst : $($_.Exception.Message)"
        $errors++
    }
}

Write-Host ""
Write-Host "Moved blobs: $moved"
Write-Host "Missing blobs (DB row had no on-disk match): $missing"
Write-Host "Errors: $errors"

if ($DryRun) {
    Write-Host "DryRun: skipping DB delete."
    exit 0
}

# --- Delete the DB rows in one shot ---
$idsCsv = ($orphans | ForEach-Object { "'$($_.id)'" }) -join ','
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()
try {
    $delCmd = $conn.CreateCommand()
    $delCmd.CommandTimeout = 120
    $delCmd.CommandText = @"
BEGIN TRANSACTION;
DELETE FROM innovator.[FILECONTAINERLOCATOR] WHERE file_id IN ($idsCsv);
DECLARE @fcl INT = @@ROWCOUNT;
DELETE FROM innovator.[FILE] WHERE id IN ($idsCsv);
DECLARE @f INT = @@ROWCOUNT;
COMMIT TRANSACTION;
SELECT @fcl AS fcl_deleted, @f AS file_deleted;
"@
    $r = $delCmd.ExecuteReader()
    $r.Read() | Out-Null
    Write-Host "FILECONTAINERLOCATOR rows deleted: $($r['fcl_deleted'])"
    Write-Host "FILE rows deleted: $($r['file_deleted'])"
    $r.Close()
}
finally {
    $conn.Close(); $conn.Dispose()
}

Write-Host ""
Write-Host "Done. To rollback: move folder back from $runDir to $VaultRoot."
