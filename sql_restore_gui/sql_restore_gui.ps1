<#
.SYNOPSIS
    SQL Server Parallel Striped Database Restore Utility - Native Windows GUI (No Python Required)
.DESCRIPTION
    Self-contained Windows Forms GUI for restoring a batch of single-file or multi-stream
    striped SQL Server databases (.bak sets and .zip archives) as quickly as possible.
    - Hybrid Discovery: Scans source folder for single/striped .bak sets (_stripe01of0N.bak)
      and .zip archives produced by sql_backup_gui, with latest-per-DB filtering and
      RESTORE HEADERONLY / RESTORE FILELISTONLY metadata verification.
    - High-Speed Execution:
      * Largest-Database-First (LPT) queue scheduling so multi-stripe DBs start immediately.
      * Parallel database restore worker pool via PowerShell RunspacePool.
      * Fast multi-threaded .zip extraction (7z.exe -mmt=on or .NET System.IO.Compression)
        with automatic staging cleanup.
      * Tuned MAXTRANSFERSIZE = 4194304 (4 MB) and auto BUFFERCOUNT (with automatic fallback
        on SQL Server virtual memory pressure).
      * Automatic WITH MOVE file relocation, WITH REPLACE, and SINGLE_USER connection drop.
#>

param(
    [switch]$NoGui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:AppTitle = "SQL Server Parallel Striped Database Restore Utility (Server / Native Edition)"
$Script:ConfigFile = Join-Path $PSScriptRoot ".sql_restore_config.json"
$Script:StagingFolderName = "_RestoreStaging"

# ============================================================================
# Pure Helper Functions (Discovery, Stripe Grouping, WITH MOVE & T-SQL Builders)
# ============================================================================

function Get-SafeFileName {
    param([string]$Name)
    $cleaned = [regex]::Replace($Name.Trim(), '[<>:"/\\|?*\x00-\x1F]', '_')
    $cleaned = $cleaned.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($cleaned)) { return "database" }
    return $cleaned
}

function Split-BackupFileName {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileName($FileName)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
    $pattern = '^(?<stem>.+?)(?:_(?<ts>\d{8}_\d{6}|\d{8}|\d{14}))?(?:_(?:stripe|part|stream|file|s)?(?<idx>\d{1,3})of(?<total>\d{1,3}))?$'
    $m = [regex]::Match($stem, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) {
        return [PSCustomObject]@{
            DbName      = $stem
            Timestamp   = ""
            StripeIndex = 1
            StripeTotal = 1
        }
    }
    $db = $m.Groups['stem'].Value.Trim('_', '-', ' ')
    if ([string]::IsNullOrWhiteSpace($db)) { $db = $stem }
    $ts = $m.Groups['ts'].Value
    $idx = 1
    $tot = 1
    if ($m.Groups['idx'].Success -and $m.Groups['total'].Success) {
        $idx = [Math]::Max(1, [int]$m.Groups['idx'].Value)
        $tot = [Math]::Max($idx, [int]$m.Groups['total'].Value)
    }
    return [PSCustomObject]@{
        DbName      = $db
        Timestamp   = $ts
        StripeIndex = $idx
        StripeTotal = $tot
    }
}

function Get-OptimalBufferCount {
    param(
        [int]$StreamCount,
        [int]$CustomBufferCount = 0
    )
    if ($CustomBufferCount -gt 0) { return $CustomBufferCount }
    $s = [Math]::Max(1, $StreamCount)
    return [Math]::Min(128, [Math]::Max(24, $s * 8))
}

function Get-TargetDatabaseName {
    param(
        [string]$SourceDbName,
        [string]$Prefix = "",
        [string]$Suffix = "",
        [string]$CustomTargetName = ""
    )
    if (-not [string]::IsNullOrWhiteSpace($CustomTargetName)) {
        return $CustomTargetName.Trim()
    }
    return "$($Prefix.Trim())$SourceDbName$($Suffix.Trim())"
}

function Get-DefaultLogicalFiles {
    param([string]$DbName)
    return @(
        [PSCustomObject]@{
            LogicalName  = $DbName
            PhysicalName = "$DbName.mdf"
            FileType     = "D"
            FileId       = 1
        },
        [PSCustomObject]@{
            LogicalName  = "${DbName}_log"
            PhysicalName = "${DbName}_log.ldf"
            FileType     = "L"
            FileId       = 2
        }
    )
}

function Find-BackupRestoreSets {
    param(
        [string]$SourceFolder,
        [bool]$LatestOnlyPerDb = $true
    )
    if ([string]::IsNullOrWhiteSpace($SourceFolder) -or (-not (Test-Path -LiteralPath $SourceFolder))) {
        return @()
    }

    $bakGroups = @{}
    $discovered = New-Object System.Collections.ArrayList

    $allFiles = @(Get-ChildItem -LiteralPath $SourceFolder -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
        $_.FullName -notmatch [regex]::Escape("\$Script:StagingFolderName\")
    })

    foreach ($f in $allFiles) {
        $ext = $f.Extension.ToLowerInvariant()
        if ($ext -eq ".bak") {
            $parsed = Split-BackupFileName -FileName $f.Name
            $dirKey = $f.DirectoryName.ToLowerInvariant()
            $key = "$dirKey|$($parsed.DbName)|$($parsed.Timestamp)|$($parsed.StripeTotal)"
            if (-not $bakGroups.ContainsKey($key)) {
                $bakGroups[$key] = New-Object System.Collections.ArrayList
            }
            [void]$bakGroups[$key].Add([PSCustomObject]@{
                StripeIndex = $parsed.StripeIndex
                FullPath    = $f.FullName
                Length      = $f.Length
                LastWrite   = $f.LastWriteTimeUtc
                DbName      = $parsed.DbName
                Timestamp   = $parsed.Timestamp
                StripeTotal = $parsed.StripeTotal
            })
        }
        elseif ($ext -eq ".zip") {
            try {
                $za = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
                try {
                    $bakEntries = @($za.Entries | Where-Object { $_.Name -like "*.bak" })
                    if ($bakEntries.Count -gt 0) {
                        $zipParsed = Split-BackupFileName -FileName $f.Name
                        $dbName = $zipParsed.DbName
                        $ts = $zipParsed.Timestamp
                        $entryObjs = New-Object System.Collections.ArrayList
                        $totalBytes = [int64]0
                        $expected = 1
                        foreach ($e in $bakEntries) {
                            $ep = Split-BackupFileName -FileName $e.Name
                            if ($ep.DbName) { $dbName = $ep.DbName }
                            if ($ep.Timestamp) { $ts = $ep.Timestamp }
                            if ($ep.StripeTotal -gt $expected) { $expected = $ep.StripeTotal }
                            $totalBytes += $e.Length
                            [void]$entryObjs.Add([PSCustomObject]@{
                                StripeIndex = $ep.StripeIndex
                                EntryName   = $e.FullName
                            })
                        }
                        $sortedEntries = @($entryObjs | Sort-Object StripeIndex, EntryName | ForEach-Object { $_.EntryName })
                        if ([string]::IsNullOrWhiteSpace($ts)) {
                            $ts = $f.LastWriteTime.ToString("yyyyMMdd_HHmmss")
                        }
                        [void]$discovered.Add([PSCustomObject]@{
                            SetId            = "zip::$($f.FullName)"
                            SourceDbName     = $dbName
                            TimestampLabel   = $ts
                            SourceType       = "zip"
                            StripePaths      = [string[]]$sortedEntries
                            ZipPath          = $f.FullName
                            ExpectedStripes  = $expected
                            FoundStripes     = $sortedEntries.Count
                            TotalSizeBytes   = $totalBytes
                            SizeGB           = [Math]::Round($totalBytes / 1GB, 3)
                            LastWriteUtc     = $f.LastWriteTimeUtc
                            LogicalFiles     = (Get-DefaultLogicalFiles -DbName $dbName)
                            HeaderVerified   = $false
                            CustomTargetName = ""
                        })
                    }
                }
                finally {
                    $za.Dispose()
                }
            }
            catch {}
        }
    }

    foreach ($k in $bakGroups.Keys) {
        $items = @($bakGroups[$k] | Sort-Object StripeIndex, FullPath)
        $first = $items[0]
        $paths = [string[]]@($items | ForEach-Object { $_.FullPath })
        $totalBytes = [int64]0
        $latestWrite = [datetime]::MinValue
        foreach ($it in $items) {
            $totalBytes += $it.Length
            if ($it.LastWrite -gt $latestWrite) { $latestWrite = $it.LastWrite }
        }
        $tsLabel = if ($first.Timestamp) { $first.Timestamp } else { $latestWrite.ToString("yyyyMMdd_HHmmss") }
        [void]$discovered.Add([PSCustomObject]@{
            SetId            = "bak::$($paths[0])"
            SourceDbName     = $first.DbName
            TimestampLabel   = $tsLabel
            SourceType       = "bak"
            StripePaths      = $paths
            ZipPath          = ""
            ExpectedStripes  = $first.StripeTotal
            FoundStripes     = $paths.Count
            TotalSizeBytes   = $totalBytes
            SizeGB           = [Math]::Round($totalBytes / 1GB, 3)
            LastWriteUtc     = $latestWrite
            LogicalFiles     = (Get-DefaultLogicalFiles -DbName $first.DbName)
            HeaderVerified   = $false
            CustomTargetName = ""
        })
    }

    $sortedAll = @($discovered | Sort-Object SourceDbName, @{Expression = { $_.TimestampLabel }; Descending = $true })
    if (-not $LatestOnlyPerDb) {
        return $sortedAll
    }

    $byDb = @{}
    foreach ($s in $sortedAll) {
        $dbKey = $s.SourceDbName.ToLowerInvariant()
        if (-not $byDb.ContainsKey($dbKey)) {
            $byDb[$dbKey] = $s
        }
        else {
            $curr = $byDb[$dbKey]
            $sComplete = ($s.FoundStripes -ge $s.ExpectedStripes)
            $cComplete = ($curr.FoundStripes -ge $curr.ExpectedStripes)
            if (($sComplete -and -not $cComplete) -or ($sComplete -eq $cComplete -and $s.TimestampLabel -gt $curr.TimestampLabel)) {
                $byDb[$dbKey] = $s
            }
        }
    }
    return @($byDb.Values | Sort-Object SourceDbName)
}

function ConvertTo-TsqlIdentifier {
    param([string]$Identifier)
    return "[" + $Identifier.Replace("]", "]]") + "]"
}

function ConvertTo-TsqlString {
    param([string]$Value)
    return "N'" + $Value.Replace("'", "''") + "'"
}

function Get-CreateSubdirTsql {
    param([string]$SubfolderPath)
    return "BEGIN TRY EXEC master.dbo.xp_create_subdir $(ConvertTo-TsqlString -Value $SubfolderPath); END TRY BEGIN CATCH END CATCH;"
}

function Get-RestoreMoveClauses {
    param(
        [string]$SourceDbName,
        [string]$TargetDbName,
        [array]$LogicalFiles,
        [string]$TargetDataFolder,
        [string]$TargetLogFolder
    )
    $files = if ($null -ne $LogicalFiles -and $LogicalFiles.Count -gt 0) { $LogicalFiles } else { Get-DefaultLogicalFiles -DbName $SourceDbName }
    $safeTarget = Get-SafeFileName -Name $TargetDbName
    $safeSource = Get-SafeFileName -Name $SourceDbName
    $isRenamed = ($safeTarget.ToLowerInvariant() -ne $safeSource.ToLowerInvariant())

    $dataBase = $TargetDataFolder.Trim().TrimEnd('\', '/')
    $logBase = $TargetLogFolder.Trim().TrimEnd('\', '/')
    $usedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $moves = New-Object System.Collections.ArrayList

    $idx = 1
    foreach ($lf in $files) {
        $isLog = ($lf.FileType -eq "L")
        $destDir = if ($isLog) { $logBase } else { $dataBase }
        $origFile = if ($lf.PhysicalName) { [System.IO.Path]::GetFileName($lf.PhysicalName) } else { "" }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($origFile)
        $ext = [System.IO.Path]::GetExtension($origFile)
        if ([string]::IsNullOrWhiteSpace($ext)) {
            $ext = if ($isLog) { ".ldf" } elseif ($idx -eq 1) { ".mdf" } else { ".ndf" }
        }
        if ([string]::IsNullOrWhiteSpace($stem)) {
            $stem = if ($isLog) { "${safeTarget}_log" } else { $safeTarget }
        }
        elseif ($isRenamed) {
            $pat = "^" + [regex]::Escape($safeSource)
            if ($stem -match $pat) {
                $stem = [regex]::Replace($stem, $pat, $safeTarget, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            }
            else {
                $stem = "${safeTarget}_$stem"
            }
        }
        $cand = "$stem$ext"
        if ($usedNames.Contains($cand)) {
            $cand = "${stem}_$idx$ext"
        }
        [void]$usedNames.Add($cand)
        [void]$moves.Add([PSCustomObject]@{
            LogicalName = $lf.LogicalName
            TargetPath  = "$destDir\$cand"
        })
        $idx++
    }
    return @($moves)
}

function Get-RestoreDatabaseTsql {
    param(
        [string]$SourceDbName,
        [string]$TargetDbName,
        [string[]]$StripeFilePaths,
        [array]$LogicalFiles = $null,
        [bool]$RelocateFiles = $true,
        [string]$TargetDataFolder = "C:\SqlData",
        [string]$TargetLogFolder = "C:\SqlLogs",
        [bool]$WithReplace = $true,
        [bool]$KillConnections = $true,
        [string]$RecoveryMode = "RECOVERY",
        [bool]$HighSpeedTuning = $true,
        [int]$MaxTransferSize = 4194304,
        [int]$CustomBufferCount = 0,
        [bool]$KeepCdc = $false,
        [int]$StatsInterval = 5
    )
    $escTarget = ConvertTo-TsqlIdentifier -Identifier $TargetDbName
    $litTarget = ConvertTo-TsqlString -Value $TargetDbName
    $diskLines = foreach ($p in $StripeFilePaths) {
        "    DISK = $(ConvertTo-TsqlString -Value $p)"
    }
    $fromClause = [string]::Join(",`r`n", $diskLines)

    $withOpts = New-Object System.Collections.Generic.List[string]
    if ($RelocateFiles -and $TargetDataFolder -and $TargetLogFolder) {
        $moves = Get-RestoreMoveClauses -SourceDbName $SourceDbName -TargetDbName $TargetDbName -LogicalFiles $LogicalFiles -TargetDataFolder $TargetDataFolder -TargetLogFolder $TargetLogFolder
        foreach ($m in $moves) {
            $withOpts.Add("MOVE $(ConvertTo-TsqlString -Value $m.LogicalName) TO $(ConvertTo-TsqlString -Value $m.TargetPath)")
        }
    }
    if ($WithReplace) { $withOpts.Add("REPLACE") }
    if ($KeepCdc) { $withOpts.Add("KEEP_CDC") }

    $rec = $RecoveryMode.Trim().ToUpperInvariant()
    if ($rec -notin @("RECOVERY", "NORECOVERY")) { $rec = "RECOVERY" }
    $withOpts.Add($rec)

    if ($HighSpeedTuning) {
        $bufCnt = Get-OptimalBufferCount -StreamCount $StripeFilePaths.Count -CustomBufferCount $CustomBufferCount
        $withOpts.Add("MAXTRANSFERSIZE = $MaxTransferSize")
        $withOpts.Add("BUFFERCOUNT = $bufCnt")
    }
    $statsVal = [Math]::Max(1, [Math]::Min(100, $StatsInterval))
    $withOpts.Add("STATS = $statsVal")

    $withStr = [string]::Join(",`r`n    ", $withOpts)
    $restoreStmt = "RESTORE DATABASE $escTarget`r`nFROM`r`n$fromClause`r`nWITH`r`n    $withStr;"

    $blocks = New-Object System.Collections.Generic.List[string]
    if ($KillConnections) {
        $blocks.Add("IF DB_ID($litTarget) IS NOT NULL AND DATABASEPROPERTYEX($litTarget, 'Status') = 'ONLINE'`r`nBEGIN`r`n    ALTER DATABASE $escTarget SET SINGLE_USER WITH ROLLBACK IMMEDIATE;`r`nEND;")
    }
    $blocks.Add($restoreStmt)
    if ($KillConnections -and $rec -eq "RECOVERY") {
        $blocks.Add("IF DB_ID($litTarget) IS NOT NULL AND DATABASEPROPERTYEX($litTarget, 'Status') = 'ONLINE'`r`nBEGIN`r`n    ALTER DATABASE $escTarget SET MULTI_USER;`r`nEND;")
    }
    return [string]::Join("`r`n`r`n", $blocks)
}

function New-SqlConnectionString {
    param(
        [string]$Server,
        [string]$AuthMode,
        [string]$Username,
        [string]$Password,
        [bool]$TrustServerCertificate = $true,
        [string]$Database = "master"
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Server=$($Server.Trim())")
    $parts.Add("Database=$Database")
    $parts.Add("Connection Timeout=15")
    $parts.Add("Application Name=SqlParallelStripedRestorePsGui")
    if ($TrustServerCertificate) {
        $parts.Add("TrustServerCertificate=True")
    }
    if ($AuthMode -eq "windows") {
        $parts.Add("Integrated Security=True")
    }
    else {
        $parts.Add("Integrated Security=False")
        $parts.Add("User ID=$($Username.Trim())")
        $safePwd = $Password.Replace(";", ";;")
        $parts.Add("Password=$safePwd")
    }
    return ([string]::Join(";", $parts) + ";")
}

if ($NoGui) {
    return
}

# ============================================================================
# WinForms GUI Construction
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = $Script:AppTitle
$form.Size = New-Object System.Drawing.Size(1240, 880)
$form.MinimumSize = New-Object System.Drawing.Size(1040, 740)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# State containers
$Script:BackupSets = New-Object System.Collections.ArrayList
$Script:ExistingDbs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$Script:UiEventQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$Script:ActiveJobs = New-Object System.Collections.ArrayList
$Script:RunspacePool = $null
$Script:IsRunning = $false
$Script:TotalBatchCount = 0
$Script:FinishedBatchCount = 0
$Script:ProgressMap = @{}

$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$mainPanel.ColumnCount = 1
$mainPanel.RowCount = 5
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 88)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 126)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 56)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 44)))
$form.Controls.Add($mainPanel)

# 0. Header Banner
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Fill"
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "SQL Server Parallel Striped Database Restore Utility"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(2, 2)
$headerPanel.Controls.Add($lblTitle)

$lblSubTitle = New-Object System.Windows.Forms.Label
$lblSubTitle.Text = "Largest-First Queue  |  Multi-Stripe .bak & Fast-Unzip  |  4 MB Transfer + Auto BUFFERCOUNT"
$lblSubTitle.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblSubTitle.AutoSize = $true
$lblSubTitle.Anchor = "Top, Right"
$lblSubTitle.Location = New-Object System.Drawing.Point(640, 6)
$headerPanel.Controls.Add($lblSubTitle)
$mainPanel.Controls.Add($headerPanel, 0, 0)

# 1. Target SQL Server Connection GroupBox
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = " 1. Target SQL Server Connection "
$grpConn.Dock = "Fill"
$mainPanel.Controls.Add($grpConn, 0, 1)

$lblSrv = New-Object System.Windows.Forms.Label
$lblSrv.Text = "Target Server:"
$lblSrv.Location = New-Object System.Drawing.Point(12, 24)
$lblSrv.AutoSize = $true
$grpConn.Controls.Add($lblSrv)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Text = "localhost"
$txtServer.Location = New-Object System.Drawing.Point(100, 21)
$txtServer.Width = 200
$grpConn.Controls.Add($txtServer)

$lblAuth = New-Object System.Windows.Forms.Label
$lblAuth.Text = "Auth:"
$lblAuth.Location = New-Object System.Drawing.Point(312, 24)
$lblAuth.AutoSize = $true
$grpConn.Controls.Add($lblAuth)

$cmbAuth = New-Object System.Windows.Forms.ComboBox
$cmbAuth.DropDownStyle = "DropDownList"
[void]$cmbAuth.Items.AddRange(@("windows", "sql"))
$cmbAuth.SelectedIndex = 0
$cmbAuth.Location = New-Object System.Drawing.Point(350, 21)
$cmbAuth.Width = 95
$grpConn.Controls.Add($cmbAuth)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Login:"
$lblUser.Location = New-Object System.Drawing.Point(456, 24)
$lblUser.AutoSize = $true
$grpConn.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = "sa"
$txtUser.Enabled = $false
$txtUser.Location = New-Object System.Drawing.Point(498, 21)
$txtUser.Width = 110
$grpConn.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.Location = New-Object System.Drawing.Point(616, 24)
$lblPass.AutoSize = $true
$grpConn.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.UseSystemPasswordChar = $true
$txtPass.Enabled = $false
$txtPass.Location = New-Object System.Drawing.Point(680, 21)
$txtPass.Width = 115
$grpConn.Controls.Add($txtPass)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Show"
$chkShowPass.Location = New-Object System.Drawing.Point(802, 23)
$chkShowPass.AutoSize = $true
$grpConn.Controls.Add($chkShowPass)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect & Detect Paths"
$btnConnect.Location = New-Object System.Drawing.Point(862, 19)
$btnConnect.Size = New-Object System.Drawing.Size(160, 27)
$grpConn.Controls.Add($btnConnect)

$btnSaveCfg = New-Object System.Windows.Forms.Button
$btnSaveCfg.Text = "Save Cfg"
$btnSaveCfg.Location = New-Object System.Drawing.Point(1030, 19)
$btnSaveCfg.Size = New-Object System.Drawing.Size(75, 27)
$grpConn.Controls.Add($btnSaveCfg)

$chkTrustCert = New-Object System.Windows.Forms.CheckBox
$chkTrustCert.Text = "Trust Server Certificate"
$chkTrustCert.Checked = $true
$chkTrustCert.Location = New-Object System.Drawing.Point(15, 54)
$chkTrustCert.AutoSize = $true
$grpConn.Controls.Add($chkTrustCert)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = "Not connected -- enter target SQL Server instance and click 'Connect & Detect Paths'"
$lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 86, 160)
$lblConnStatus.Location = New-Object System.Drawing.Point(200, 55)
$lblConnStatus.AutoSize = $true
$grpConn.Controls.Add($lblConnStatus)

# 2. Source Backup Folder, WITH MOVE Relocation & High-Speed Tuning GroupBox
$grpSettings = New-Object System.Windows.Forms.GroupBox
$grpSettings.Text = " 2. Source Backup Folder (.bak / .zip), File Relocation (WITH MOVE) & High-Speed Tuning "
$grpSettings.Dock = "Fill"
$mainPanel.Controls.Add($grpSettings, 0, 2)

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Backup Folder (.bak/.zip):"
$lblSource.Location = New-Object System.Drawing.Point(12, 24)
$lblSource.AutoSize = $true
$grpSettings.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Text = "C:\SqlBackups"
$txtSource.Location = New-Object System.Drawing.Point(195, 21)
$txtSource.Width = 420
$grpSettings.Controls.Add($txtSource)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "Browse..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(622, 19)
$btnBrowseSource.Size = New-Object System.Drawing.Size(80, 26)
$grpSettings.Controls.Add($btnBrowseSource)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Backups"
$btnScan.Location = New-Object System.Drawing.Point(708, 19)
$btnScan.Size = New-Object System.Drawing.Size(100, 26)
$grpSettings.Controls.Add($btnScan)

$chkLatestOnly = New-Object System.Windows.Forms.CheckBox
$chkLatestOnly.Text = "Latest per DB only"
$chkLatestOnly.Checked = $true
$chkLatestOnly.Location = New-Object System.Drawing.Point(818, 23)
$chkLatestOnly.AutoSize = $true
$grpSettings.Controls.Add($chkLatestOnly)

$chkCleanupStaging = New-Object System.Windows.Forms.CheckBox
$chkCleanupStaging.Text = "Delete staged .bak after .zip restore"
$chkCleanupStaging.Checked = $true
$chkCleanupStaging.Location = New-Object System.Drawing.Point(950, 23)
$chkCleanupStaging.AutoSize = $true
$grpSettings.Controls.Add($chkCleanupStaging)

# Row B: WITH MOVE Data & Log Folders + Prefix/Suffix
$chkRelocate = New-Object System.Windows.Forms.CheckBox
$chkRelocate.Text = "Relocate Files (WITH MOVE) Data:"
$chkRelocate.Checked = $true
$chkRelocate.Location = New-Object System.Drawing.Point(14, 54)
$chkRelocate.AutoSize = $true
$grpSettings.Controls.Add($chkRelocate)

$txtDataFolder = New-Object System.Windows.Forms.TextBox
$txtDataFolder.Text = "C:\SqlData"
$txtDataFolder.Location = New-Object System.Drawing.Point(225, 52)
$txtDataFolder.Width = 240
$grpSettings.Controls.Add($txtDataFolder)

$btnBrowseData = New-Object System.Windows.Forms.Button
$btnBrowseData.Text = "..."
$btnBrowseData.Location = New-Object System.Drawing.Point(470, 50)
$btnBrowseData.Size = New-Object System.Drawing.Size(34, 26)
$grpSettings.Controls.Add($btnBrowseData)

$lblLogFolder = New-Object System.Windows.Forms.Label
$lblLogFolder.Text = "Log (.ldf):"
$lblLogFolder.Location = New-Object System.Drawing.Point(514, 55)
$lblLogFolder.AutoSize = $true
$grpSettings.Controls.Add($lblLogFolder)

$txtLogFolder = New-Object System.Windows.Forms.TextBox
$txtLogFolder.Text = "C:\SqlLogs"
$txtLogFolder.Location = New-Object System.Drawing.Point(576, 52)
$txtLogFolder.Width = 240
$grpSettings.Controls.Add($txtLogFolder)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Text = "..."
$btnBrowseLog.Location = New-Object System.Drawing.Point(820, 50)
$btnBrowseLog.Size = New-Object System.Drawing.Size(34, 26)
$grpSettings.Controls.Add($btnBrowseLog)

$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "DB Prefix:"
$lblPrefix.Location = New-Object System.Drawing.Point(866, 55)
$lblPrefix.AutoSize = $true
$grpSettings.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Location = New-Object System.Drawing.Point(926, 52)
$txtPrefix.Width = 75
$grpSettings.Controls.Add($txtPrefix)

$lblSuffix = New-Object System.Windows.Forms.Label
$lblSuffix.Text = "Suffix:"
$lblSuffix.Location = New-Object System.Drawing.Point(1010, 55)
$lblSuffix.AutoSize = $true
$grpSettings.Controls.Add($lblSuffix)

$txtSuffix = New-Object System.Windows.Forms.TextBox
$txtSuffix.Location = New-Object System.Drawing.Point(1052, 52)
$txtSuffix.Width = 75
$grpSettings.Controls.Add($txtSuffix)

# Row C: Parallelism & High-Speed Tuning
$lblParallel = New-Object System.Windows.Forms.Label
$lblParallel.Text = "Max Parallel DBs:"
$lblParallel.Location = New-Object System.Drawing.Point(12, 88)
$lblParallel.AutoSize = $true
$grpSettings.Controls.Add($lblParallel)

$numParallelDbs = New-Object System.Windows.Forms.NumericUpDown
$numParallelDbs.Minimum = 1
$numParallelDbs.Maximum = 16
$numParallelDbs.Value = 3
$numParallelDbs.Location = New-Object System.Drawing.Point(116, 85)
$numParallelDbs.Width = 50
$grpSettings.Controls.Add($numParallelDbs)

$chkReplace = New-Object System.Windows.Forms.CheckBox
$chkReplace.Text = "WITH REPLACE"
$chkReplace.Checked = $true
$chkReplace.Location = New-Object System.Drawing.Point(180, 87)
$chkReplace.AutoSize = $true
$grpSettings.Controls.Add($chkReplace)

$chkKillConn = New-Object System.Windows.Forms.CheckBox
$chkKillConn.Text = "Kill Connections (SINGLE_USER)"
$chkKillConn.Checked = $true
$chkKillConn.Location = New-Object System.Drawing.Point(295, 87)
$chkKillConn.AutoSize = $true
$grpSettings.Controls.Add($chkKillConn)

$lblRecMode = New-Object System.Windows.Forms.Label
$lblRecMode.Text = "State:"
$lblRecMode.Location = New-Object System.Drawing.Point(502, 88)
$lblRecMode.AutoSize = $true
$grpSettings.Controls.Add($lblRecMode)

$cmbRecMode = New-Object System.Windows.Forms.ComboBox
$cmbRecMode.DropDownStyle = "DropDownList"
[void]$cmbRecMode.Items.AddRange(@("RECOVERY", "NORECOVERY"))
$cmbRecMode.SelectedIndex = 0
$cmbRecMode.Location = New-Object System.Drawing.Point(542, 85)
$cmbRecMode.Width = 105
$grpSettings.Controls.Add($cmbRecMode)

$chkHighSpeed = New-Object System.Windows.Forms.CheckBox
$chkHighSpeed.Text = "High-Speed I/O (4MB Transfer + BUFFERCOUNT)"
$chkHighSpeed.Checked = $true
$chkHighSpeed.Location = New-Object System.Drawing.Point(662, 87)
$chkHighSpeed.AutoSize = $true
$grpSettings.Controls.Add($chkHighSpeed)

$lblBufCnt = New-Object System.Windows.Forms.Label
$lblBufCnt.Text = "Buf (0=Auto):"
$lblBufCnt.Location = New-Object System.Drawing.Point(958, 88)
$lblBufCnt.AutoSize = $true
$grpSettings.Controls.Add($lblBufCnt)

$numBufCnt = New-Object System.Windows.Forms.NumericUpDown
$numBufCnt.Minimum = 0
$numBufCnt.Maximum = 512
$numBufCnt.Increment = 8
$numBufCnt.Value = 0
$numBufCnt.Location = New-Object System.Drawing.Point(1042, 85)
$numBufCnt.Width = 54
$grpSettings.Controls.Add($numBufCnt)

$chkKeepCdc = New-Object System.Windows.Forms.CheckBox
$chkKeepCdc.Text = "KEEP_CDC"
$chkKeepCdc.Checked = $false
$chkKeepCdc.Location = New-Object System.Drawing.Point(1105, 87)
$chkKeepCdc.AutoSize = $true
$grpSettings.Controls.Add($chkKeepCdc)

# 3. Backup Set Selection GroupBox
$grpDbs = New-Object System.Windows.Forms.GroupBox
$grpDbs.Text = " 3. Select Striped Backup Sets to Restore "
$grpDbs.Dock = "Fill"
$mainPanel.Controls.Add($grpDbs, 0, 3)

$dbLayout = New-Object System.Windows.Forms.TableLayoutPanel
$dbLayout.Dock = "Fill"
$dbLayout.ColumnCount = 1
$dbLayout.RowCount = 2
[void]$dbLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
[void]$dbLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$grpDbs.Controls.Add($dbLayout)

$dbToolbar = New-Object System.Windows.Forms.Panel
$dbToolbar.Dock = "Fill"
$dbLayout.Controls.Add($dbToolbar, 0, 0)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Text = "Select All"
$btnSelAll.Location = New-Object System.Drawing.Point(4, 3)
$btnSelAll.Size = New-Object System.Drawing.Size(82, 26)
$dbToolbar.Controls.Add($btnSelAll)

$btnDeselAll = New-Object System.Windows.Forms.Button
$btnDeselAll.Text = "Deselect All"
$btnDeselAll.Location = New-Object System.Drawing.Point(92, 3)
$btnDeselAll.Size = New-Object System.Drawing.Size(90, 26)
$dbToolbar.Controls.Add($btnDeselAll)

$btnSelStriped = New-Object System.Windows.Forms.Button
$btnSelStriped.Text = "Select Multi-Stripe Only"
$btnSelStriped.Location = New-Object System.Drawing.Point(188, 3)
$btnSelStriped.Size = New-Object System.Drawing.Size(155, 26)
$dbToolbar.Controls.Add($btnSelStriped)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(358, 7)
$lblFilter.AutoSize = $true
$dbToolbar.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(400, 4)
$txtFilter.Width = 160
$dbToolbar.Controls.Add($txtFilter)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "Selected: 0 DBs (0.00 GB | 0 total streams)"
$lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSummary.Location = New-Object System.Drawing.Point(580, 7)
$lblSummary.AutoSize = $true
$dbToolbar.Controls.Add($lblSummary)

$lvSets = New-Object System.Windows.Forms.ListView
$lvSets.Dock = "Fill"
$lvSets.View = "Details"
$lvSets.CheckBoxes = $true
$lvSets.FullRowSelect = $true
$lvSets.GridLines = $true
[void]$lvSets.Columns.Add("Backup DB Name", 180)
[void]$lvSets.Columns.Add("Target DB Name", 185)
[void]$lvSets.Columns.Add("Size (GB)", 95)
[void]$lvSets.Columns.Add("Stripes / Buffers", 150)
[void]$lvSets.Columns.Add("Format & Timestamp", 160)
[void]$lvSets.Columns.Add("Status / Progress", 220)
[void]$lvSets.Columns.Add("Source / Target Relocation", 290)
$dbLayout.Controls.Add($lvSets, 0, 1)

# 4. Bottom Panel (Split T-SQL Preview + Execution Log + Action Bar)
$bottomLayout = New-Object System.Windows.Forms.TableLayoutPanel
$bottomLayout.Dock = "Fill"
$bottomLayout.ColumnCount = 1
$bottomLayout.RowCount = 2
[void]$bottomLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$bottomLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
$mainPanel.Controls.Add($bottomLayout, 0, 4)

$splitBottom = New-Object System.Windows.Forms.SplitContainer
$splitBottom.Dock = "Fill"
$splitBottom.Orientation = "Vertical"
$splitBottom.SplitterDistance = 580
$bottomLayout.Controls.Add($splitBottom, 0, 0)

$grpPreview = New-Object System.Windows.Forms.GroupBox
$grpPreview.Text = " Live High-Speed RESTORE DATABASE T-SQL Preview "
$grpPreview.Dock = "Fill"
$splitBottom.Panel1.Controls.Add($grpPreview)

$txtPreview = New-Object System.Windows.Forms.TextBox
$txtPreview.Multiline = $true
$txtPreview.ReadOnly = $true
$txtPreview.ScrollBars = "Both"
$txtPreview.WordWrap = $false
$txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtPreview.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$txtPreview.Dock = "Fill"
$grpPreview.Controls.Add($txtPreview)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = " Parallel Restore & Fast-Unzip Execution Log "
$grpLog.Dock = "Fill"
$splitBottom.Panel2.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$txtLog.Dock = "Fill"
$grpLog.Controls.Add($txtLog)

# Bottom Action Bar
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Dock = "Fill"
$bottomLayout.Controls.Add($actionPanel, 0, 1)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Parallel Striped Restores"
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnStart.Location = New-Object System.Drawing.Point(0, 4)
$btnStart.Size = New-Object System.Drawing.Size(230, 28)
$actionPanel.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Restores"
$btnCancel.Enabled = $false
$btnCancel.Location = New-Object System.Drawing.Point(238, 4)
$btnCancel.Size = New-Object System.Drawing.Size(120, 28)
$actionPanel.Controls.Add($btnCancel)

$btnCopyTsql = New-Object System.Windows.Forms.Button
$btnCopyTsql.Text = "Copy T-SQL"
$btnCopyTsql.Location = New-Object System.Drawing.Point(366, 4)
$btnCopyTsql.Size = New-Object System.Drawing.Size(95, 28)
$actionPanel.Controls.Add($btnCopyTsql)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Location = New-Object System.Drawing.Point(472, 6)
$progressBar.Size = New-Object System.Drawing.Size(725, 24)
$progressBar.Anchor = "Top, Left, Right"
$actionPanel.Controls.Add($progressBar)

# ============================================================================
# UI Logic & Event Wiring
# ============================================================================

function Write-UiLog {
    param([string]$Message)
    $ts = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("[$ts] $Message`r`n")
}

function Update-SummaryAndPreview {
    $maxParallel = [int]$numParallelDbs.Value
    $customBuf = [int]$numBufCnt.Value
    $prefix = $txtPrefix.Text.Trim()
    $suffix = $txtSuffix.Text.Trim()
    $relocate = $chkRelocate.Checked
    $dataDir = $txtDataFolder.Text.Trim()
    $logDir = $txtLogFolder.Text.Trim()
    $sourceDir = $txtSource.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($sourceDir)) { $sourceDir = "C:\SqlBackups" }
    $stagingDir = "$($sourceDir.TrimEnd('\','/'))\$Script:StagingFolderName"

    $checkedSets = New-Object System.Collections.ArrayList
    $totalGb = 0.0
    $totalStreams = 0
    $stripedCount = 0

    foreach ($item in $lvSets.Items) {
        $bset = $item.Tag
        $targetDb = Get-TargetDatabaseName -SourceDbName $bset.SourceDbName -Prefix $prefix -Suffix $suffix -CustomTargetName $bset.CustomTargetName
        $streamCount = [Math]::Max(1, $bset.StripePaths.Count)
        $bufCnt = Get-OptimalBufferCount -StreamCount $streamCount -CustomBufferCount $customBuf
        $stripeText = if ($streamCount -eq 1) { "1 stripe (Buf=$bufCnt)" } else { "$($bset.FoundStripes)/$($bset.ExpectedStripes) stripes (Buf=$bufCnt)" }
        $existsTag = if ($Script:ExistingDbs.Contains($targetDb)) { " (Overwrite)" } else { "" }

        $item.SubItems[1].Text = "$targetDb$existsTag"
        $item.SubItems[3].Text = $stripeText

        if ($item.Checked) {
            [void]$checkedSets.Add($bset)
            $totalGb += $bset.SizeGB
            $totalStreams += $streamCount
            if ($streamCount -gt 1) { $stripedCount++ }
        }
    }

    $lblSummary.Text = ("Selected: {0} DBs ({1:N2} GB  |  {2} multi-stripe  |  {3} total streams)" -f $checkedSets.Count, $totalGb, $stripedCount, $totalStreams)

    # Order preview Largest-First
    $orderedChecked = @($checkedSets | Sort-Object TotalSizeBytes -Descending)
    $previewSets = New-Object System.Collections.ArrayList
    if ($lvSets.SelectedItems.Count -gt 0) {
        [void]$previewSets.Add($lvSets.SelectedItems[0].Tag)
    }
    foreach ($s in $orderedChecked) {
        if (-not $previewSets.Contains($s)) {
            [void]$previewSets.Add($s)
        }
    }

    if ($previewSets.Count -eq 0) {
        [void]$previewSets.Add([PSCustomObject]@{
            SourceDbName     = "EnterpriseWarehouseDB"
            TimestampLabel   = "20261021_100000"
            SourceType       = "bak"
            StripePaths      = @(
                "$sourceDir\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe01of04.bak",
                "$sourceDir\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe02of04.bak",
                "$sourceDir\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe03of04.bak",
                "$sourceDir\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe04of04.bak"
            )
            ZipPath          = ""
            SizeGB           = 28.0
            LogicalFiles     = (Get-DefaultLogicalFiles -DbName "EnterpriseWarehouseDB")
            CustomTargetName = ""
        })
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("-- Parallel Restore Workers: $maxParallel concurrent DBs | Queue Order: Largest-Database-First (LPT)")
    [void]$sb.AppendLine("-- High-Speed I/O Tuning: $(if ($chkHighSpeed.Checked) { 'ENABLED (MAXTRANSFERSIZE = 4194304 + Auto BUFFERCOUNT)' } else { 'DISABLED' })")
    [void]$sb.AppendLine("")

    if ($relocate -and $dataDir -and $logDir) {
        [void]$sb.AppendLine((Get-CreateSubdirTsql -SubfolderPath $dataDir))
        [void]$sb.AppendLine((Get-CreateSubdirTsql -SubfolderPath $logDir))
        [void]$sb.AppendLine("")
    }

    $limit = [Math]::Min(5, $previewSets.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $bset = $previewSets[$i]
        $targetDb = Get-TargetDatabaseName -SourceDbName $bset.SourceDbName -Prefix $prefix -Suffix $suffix -CustomTargetName $bset.CustomTargetName
        $effPaths = $bset.StripePaths
        if ($bset.SourceType -eq "zip") {
            $safeDb = Get-SafeFileName -Name $bset.SourceDbName
            $subStage = "$stagingDir\${safeDb}_$($bset.TimestampLabel)"
            $effPaths = @($bset.StripePaths | ForEach-Object { "$subStage\$([System.IO.Path]::GetFileName($_))" })
            [void]$sb.AppendLine("-- Fast-Unzip Archive: $($bset.ZipPath) -> $subStage")
        }
        $tsql = Get-RestoreDatabaseTsql `
            -SourceDbName $bset.SourceDbName `
            -TargetDbName $targetDb `
            -StripeFilePaths $effPaths `
            -LogicalFiles $bset.LogicalFiles `
            -RelocateFiles $relocate `
            -TargetDataFolder $dataDir `
            -TargetLogFolder $logDir `
            -WithReplace $chkReplace.Checked `
            -KillConnections $chkKillConn.Checked `
            -RecoveryMode ([string]$cmbRecMode.SelectedItem) `
            -HighSpeedTuning $chkHighSpeed.Checked `
            -CustomBufferCount $customBuf `
            -KeepCdc $chkKeepCdc.Checked
        [void]$sb.AppendLine(("-- Database: [{0}] -> [{1}] ({2:N2} GB | {3} stripe(s))" -f $bset.SourceDbName, $targetDb, $bset.SizeGB, $effPaths.Count))
        [void]$sb.AppendLine($tsql)
        [void]$sb.AppendLine("")
    }
    $txtPreview.Text = $sb.ToString()
}

function Populate-BackupSetsListView {
    $lvSets.BeginUpdate()
    try {
        $lvSets.Items.Clear()
        $filter = $txtFilter.Text.Trim().ToLowerInvariant()
        $prefix = $txtPrefix.Text.Trim()
        $suffix = $txtSuffix.Text.Trim()
        $customBuf = [int]$numBufCnt.Value
        $dataDir = $txtDataFolder.Text.Trim()

        foreach ($bset in $Script:BackupSets) {
            $targetDb = Get-TargetDatabaseName -SourceDbName $bset.SourceDbName -Prefix $prefix -Suffix $suffix -CustomTargetName $bset.CustomTargetName
            if ($filter -and (-not $bset.SourceDbName.ToLowerInvariant().Contains($filter)) -and (-not $targetDb.ToLowerInvariant().Contains($filter))) {
                continue
            }
            $streamCount = [Math]::Max(1, $bset.StripePaths.Count)
            $bufCnt = Get-OptimalBufferCount -StreamCount $streamCount -CustomBufferCount $customBuf
            $stripeText = if ($streamCount -eq 1) { "1 stripe (Buf=$bufCnt)" } else { "$($bset.FoundStripes)/$($bset.ExpectedStripes) stripes (Buf=$bufCnt)" }
            $fmtText = "$(if ($bset.SourceType -eq 'zip') { 'ZIP' } else { 'RAW .bak' }) | $($bset.TimestampLabel)"
            $isComplete = ($bset.FoundStripes -ge $bset.ExpectedStripes)
            $statusInit = if ($isComplete) { "Ready" } else { "Missing stripes ($($bset.FoundStripes)/$($bset.ExpectedStripes))" }
            $srcFile = if ($bset.SourceType -eq "zip") { [System.IO.Path]::GetFileName($bset.ZipPath) } else { [System.IO.Path]::GetFileName($bset.StripePaths[0]) }

            $item = New-Object System.Windows.Forms.ListViewItem($bset.SourceDbName)
            $item.Checked = $isComplete
            [void]$item.SubItems.Add($targetDb)
            [void]$item.SubItems.Add(("{0:N2} GB" -f $bset.SizeGB))
            [void]$item.SubItems.Add($stripeText)
            [void]$item.SubItems.Add($fmtText)
            [void]$item.SubItems.Add($statusInit)
            [void]$item.SubItems.Add("$srcFile -> $dataDir")
            $item.Tag = $bset
            [void]$lvSets.Items.Add($item)
        }
    }
    finally {
        $lvSets.EndUpdate()
    }
    Update-SummaryAndPreview
}

function Invoke-ScanBackups {
    $src = $txtSource.Text.Trim()
    if (-not (Test-Path -LiteralPath $src)) { return }
    $sets = Find-BackupRestoreSets -SourceFolder $src -LatestOnlyPerDb $chkLatestOnly.Checked
    $Script:BackupSets.Clear()
    foreach ($s in $sets) { [void]$Script:BackupSets.Add($s) }
    Write-UiLog "Scanned '$src': discovered $($Script:BackupSets.Count) backup set(s)."
    Populate-BackupSetsListView
}

$cmbAuth.Add_SelectedIndexChanged({
    $isSql = ($cmbAuth.SelectedItem -eq "sql")
    $txtUser.Enabled = $isSql
    $txtPass.Enabled = $isSql
})

$chkShowPass.Add_CheckedChanged({
    $txtPass.UseSystemPasswordChar = (-not $chkShowPass.Checked)
})

$chkRelocate.Add_CheckedChanged({
    $en = $chkRelocate.Checked
    $txtDataFolder.Enabled = $en
    $btnBrowseData.Enabled = $en
    $txtLogFolder.Enabled = $en
    $btnBrowseLog.Enabled = $en
    Update-SummaryAndPreview
})

$btnBrowseSource.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Source Backup Folder (.bak / .zip)"
    $dlg.SelectedPath = $txtSource.Text
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSource.Text = $dlg.SelectedPath
        Invoke-ScanBackups
    }
})

$btnBrowseData.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Target SQL Data Folder (.mdf / .ndf)"
    $dlg.SelectedPath = $txtDataFolder.Text
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDataFolder.Text = $dlg.SelectedPath
        Update-SummaryAndPreview
    }
})

$btnBrowseLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Target SQL Log Folder (.ldf)"
    $dlg.SelectedPath = $txtLogFolder.Text
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogFolder.Text = $dlg.SelectedPath
        Update-SummaryAndPreview
    }
})

$btnScan.Add_Click({ Invoke-ScanBackups })
$chkLatestOnly.Add_CheckedChanged({ Invoke-ScanBackups })
$chkCleanupStaging.Add_CheckedChanged({ Update-SummaryAndPreview })
$txtDataFolder.Add_TextChanged({ Update-SummaryAndPreview })
$txtLogFolder.Add_TextChanged({ Update-SummaryAndPreview })
$txtPrefix.Add_TextChanged({ Update-SummaryAndPreview })
$txtSuffix.Add_TextChanged({ Update-SummaryAndPreview })
$numParallelDbs.Add_ValueChanged({ Update-SummaryAndPreview })
$chkReplace.Add_CheckedChanged({ Update-SummaryAndPreview })
$chkKillConn.Add_CheckedChanged({ Update-SummaryAndPreview })
$cmbRecMode.Add_SelectedIndexChanged({ Update-SummaryAndPreview })
$chkHighSpeed.Add_CheckedChanged({ Update-SummaryAndPreview })
$numBufCnt.Add_ValueChanged({ Update-SummaryAndPreview })
$chkKeepCdc.Add_CheckedChanged({ Update-SummaryAndPreview })
$txtFilter.Add_TextChanged({ Populate-BackupSetsListView })
$lvSets.Add_ItemChecked({ Update-SummaryAndPreview })
$lvSets.Add_SelectedIndexChanged({ Update-SummaryAndPreview })

$btnSelAll.Add_Click({
    foreach ($item in $lvSets.Items) { $item.Checked = $true }
    Update-SummaryAndPreview
})
$btnDeselAll.Add_Click({
    foreach ($item in $lvSets.Items) { $item.Checked = $false }
    Update-SummaryAndPreview
})
$btnSelStriped.Add_Click({
    foreach ($item in $lvSets.Items) {
        $item.Checked = ($item.Tag.StripePaths.Count -gt 1)
    }
    Update-SummaryAndPreview
})

$btnCopyTsql.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($txtPreview.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($txtPreview.Text)
        Write-UiLog "Copied RESTORE DATABASE T-SQL preview to clipboard."
    }
})

$btnSaveCfg.Add_Click({
    $cfg = @{
        server = $txtServer.Text.Trim()
        auth_mode = [string]$cmbAuth.SelectedItem
        username = $txtUser.Text.Trim()
        trust_server_certificate = [bool]$chkTrustCert.Checked
        source_folder = $txtSource.Text.Trim()
        latest_only_per_db = [bool]$chkLatestOnly.Checked
        cleanup_staging_bak = [bool]$chkCleanupStaging.Checked
        relocate_files = [bool]$chkRelocate.Checked
        target_data_folder = $txtDataFolder.Text.Trim()
        target_log_folder = $txtLogFolder.Text.Trim()
        target_db_prefix = $txtPrefix.Text.Trim()
        target_db_suffix = $txtSuffix.Text.Trim()
        max_parallel_dbs = [int]$numParallelDbs.Value
        with_replace = [bool]$chkReplace.Checked
        kill_connections = [bool]$chkKillConn.Checked
        recovery_mode = [string]$cmbRecMode.SelectedItem
        high_speed_tuning = [bool]$chkHighSpeed.Checked
        custom_buffer_count = [int]$numBufCnt.Value
        keep_cdc = [bool]$chkKeepCdc.Checked
    }
    $cfg | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Script:ConfigFile -Encoding UTF8
    Write-UiLog "Saved configuration to .sql_restore_config.json"
})

# Connect & Detect Instance Paths + IFI
$btnConnect.Add_Click({
    if ($Script:IsRunning) { return }
    $connStr = New-SqlConnectionString -Server $txtServer.Text -AuthMode ([string]$cmbAuth.SelectedItem) -Username $txtUser.Text -Password $txtPass.Text -TrustServerCertificate $chkTrustCert.Checked
    Write-UiLog "Connecting to target SQL Server '$($txtServer.Text)'..."

    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
DECLARE @ifi NVARCHAR(32) = N'Unknown';
BEGIN TRY
    SELECT TOP (1) @ifi = CASE WHEN instant_file_initialization_enabled = 'Y' THEN N'Enabled' ELSE N'Disabled' END
    FROM sys.dm_server_services WHERE servicename LIKE N'SQL Server (%';
END TRY
BEGIN CATCH
    SET @ifi = N'Unknown';
END CATCH;
SELECT
    CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(256)) AS srv,
    CAST(SERVERPROPERTY('Edition') AS NVARCHAR(256)) AS ed,
    CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS ver,
    CAST(ISNULL(SERVERPROPERTY('InstanceDefaultDataPath'), N'C:\SqlData') AS NVARCHAR(512)) AS data_path,
    CAST(ISNULL(SERVERPROPERTY('InstanceDefaultLogPath'), N'C:\SqlLogs') AS NVARCHAR(512)) AS log_path,
    @ifi AS ifi_status;
"@
        $r = $cmd.ExecuteReader()
        $srvSummary = $txtServer.Text
        if ($r.Read()) {
            $srvSummary = "$($r['srv']) -- $($r['ed']) (v$($r['ver'])) | IFI: $($r['ifi_status'])"
            $dp = ([string]$r['data_path']).TrimEnd('\', '/')
            $lp = ([string]$r['log_path']).TrimEnd('\', '/')
            if ($dp) { $txtDataFolder.Text = $dp }
            if ($lp) { $txtLogFolder.Text = $lp }
        }
        $r.Close()

        $dbCmd = $conn.CreateCommand()
        $dbCmd.CommandText = "SELECT name FROM sys.databases WHERE database_id > 4;"
        $dr = $dbCmd.ExecuteReader()
        $Script:ExistingDbs.Clear()
        while ($dr.Read()) {
            [void]$Script:ExistingDbs.Add([string]$dr['name'])
        }
        $dr.Close()

        $lblConnStatus.Text = "Connected: $srvSummary ($($Script:ExistingDbs.Count) existing DBs)"
        Write-UiLog "Connected to $srvSummary. Default Data: '$($txtDataFolder.Text)', Log: '$($txtLogFolder.Text)'."
        Update-SummaryAndPreview
    }
    catch {
        $lblConnStatus.Text = "Connection failed"
        Write-UiLog "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "SQL Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    finally {
        if ($conn.State -ne 'Closed') { $conn.Close() }
        $conn.Dispose()
    }
})

# Parallel Restore & Fast-Unzip Worker ScriptBlock
$restoreWorkerScriptBlock = {
    param(
        [string]$ConnStr,
        [string]$SetId,
        [string]$SourceDbName,
        [string]$TargetDbName,
        [string]$SourceType,
        [string[]]$StripePaths,
        [string]$ZipPath,
        [string]$StagingRoot,
        [string]$TimestampLabel,
        [bool]$RelocateFiles,
        [string]$TargetDataFolder,
        [string]$TargetLogFolder,
        [bool]$WithReplace,
        [bool]$KillConnections,
        [string]$RecoveryMode,
        [bool]$HighSpeedTuning,
        [int]$CustomBufferCount,
        [bool]$KeepCdc,
        [bool]$CleanupStaging,
        $EventQueue
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $extractDir = ""
    $activePaths = $StripePaths
    $unzipElapsed = 0.0

    try {
        if ($SourceType -eq "zip") {
            $EventQueue.Enqueue(@{
                Type = "PROGRESS"
                SetId = $SetId
                DbName = $TargetDbName
                Percent = 5
                Status = "Fast-Unzipping archive to staging..."
            })
            $uSw = [System.Diagnostics.Stopwatch]::StartNew()
            $safeDb = [regex]::Replace($SourceDbName.Trim(), '[<>:"/\\|?*\x00-\x1F]', '_')
            $extractDir = "$StagingRoot\${safeDb}_$TimestampLabel"
            if (-not (Test-Path -LiteralPath $extractDir)) {
                New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
            }
            & icacls $StagingRoot /grant "*S-1-5-32-545:(OI)(CI)M" /grant "*S-1-5-11:(OI)(CI)M" /Q 2>$null | Out-Null
            & icacls $extractDir /grant "*S-1-5-32-545:(OI)(CI)M" /grant "*S-1-5-11:(OI)(CI)M" /Q 2>$null | Out-Null

            $sevenZip = $null
            foreach ($cand in @("C:\Program Files\7-Zip\7z.exe", "C:\Program Files (x86)\7-Zip\7z.exe")) {
                if (Test-Path -LiteralPath $cand) { $sevenZip = $cand; break }
            }
            if ($null -ne $sevenZip) {
                $argList = @("e", "-y", "-mmt=on", "`"-o$extractDir`"", "`"$ZipPath`"", "*.bak")
                $p = Start-Process -FilePath $sevenZip -ArgumentList $argList -NoNewWindow -Wait -PassThru
                if ($p.ExitCode -ne 0) { throw "7-Zip extraction exited with code $($p.ExitCode)" }
            }
            else {
                $za = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
                try {
                    foreach ($entry in $za.Entries) {
                        if ($entry.Name -like "*.bak") {
                            $destPath = Join-Path $extractDir $entry.Name
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
                        }
                    }
                }
                finally {
                    $za.Dispose()
                }
            }
            $activePaths = [string[]]@(Get-ChildItem -LiteralPath $extractDir -Filter "*.bak" | Sort-Object Name | ForEach-Object { $_.FullName })
            if ($activePaths.Count -eq 0) { throw "No .bak files extracted into $extractDir" }
            $uSw.Stop()
            $unzipElapsed = $uSw.Elapsed.TotalSeconds
        }

        $streamCount = $activePaths.Count
        $EventQueue.Enqueue(@{
            Type = "PROGRESS"
            SetId = $SetId
            DbName = $TargetDbName
            Percent = 15
            Status = "Starting SQL Restore ($streamCount stripe(s))..."
        })

        $conn = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
        $conn.FireInfoMessageEventOnUserErrors = $false

        $msgHandler = [System.Data.SqlClient.SqlInfoMessageEventHandler]{
            param($sender, $args)
            foreach ($err in $args.Errors) {
                if ($null -ne $err -and $err.Message -match '(\d+)\s+percent\s+processed') {
                    $pct = [int]$Matches[1]
                    $scaled = if ($SourceType -eq "zip") { [int](15 + $pct * 0.83) } else { $pct }
                    $EventQueue.Enqueue(@{
                        Type = "PROGRESS"
                        SetId = $SetId
                        DbName = $TargetDbName
                        Percent = $scaled
                        Status = "SQL Restore ($pct% | $streamCount stripe(s))"
                    })
                }
            }
        }
        $conn.add_InfoMessage($msgHandler)

        try {
            $conn.Open()

            # Query RESTORE FILELISTONLY on stripe 1 for accurate WITH MOVE mapping
            $moveList = New-Object System.Collections.Generic.List[string]
            if ($RelocateFiles -and $TargetDataFolder -and $TargetLogFolder) {
                try {
                    $dirCmd = $conn.CreateCommand()
                    $dirCmd.CommandText = "BEGIN TRY EXEC master.dbo.xp_create_subdir N'$($TargetDataFolder.Replace("'","''"))'; END TRY BEGIN CATCH END CATCH; BEGIN TRY EXEC master.dbo.xp_create_subdir N'$($TargetLogFolder.Replace("'","''"))'; END TRY BEGIN CATCH END CATCH;"
                    [void]$dirCmd.ExecuteNonQuery()
                } catch {}

                $flCmd = $conn.CreateCommand()
                $flCmd.CommandTimeout = 60
                $flCmd.CommandText = "RESTORE FILELISTONLY FROM DISK = N'$($activePaths[0].Replace("'","''"))';"
                $fr = $flCmd.ExecuteReader()
                $safeTarget = [regex]::Replace($TargetDbName.Trim(), '[<>:"/\\|?*\x00-\x1F]', '_')
                $safeSource = [regex]::Replace($SourceDbName.Trim(), '[<>:"/\\|?*\x00-\x1F]', '_')
                $isRenamed = ($safeTarget.ToLowerInvariant() -ne $safeSource.ToLowerInvariant())
                $usedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                $fIdx = 1
                while ($fr.Read()) {
                    $lName = [string]$fr['LogicalName']
                    $pName = [string]$fr['PhysicalName']
                    $fType = [string]$fr['Type']
                    $isLog = ($fType -eq 'L')
                    $destBase = if ($isLog) { $TargetLogFolder.TrimEnd('\','/') } else { $TargetDataFolder.TrimEnd('\','/') }
                    $origFile = [System.IO.Path]::GetFileName($pName)
                    $stem = [System.IO.Path]::GetFileNameWithoutExtension($origFile)
                    $ext = [System.IO.Path]::GetExtension($origFile)
                    if (-not $ext) { $ext = if ($isLog) { ".ldf" } elseif ($fIdx -eq 1) { ".mdf" } else { ".ndf" } }
                    if ($isRenamed) {
                        $pat = "^" + [regex]::Escape($safeSource)
                        if ($stem -match $pat) {
                            $stem = [regex]::Replace($stem, $pat, $safeTarget, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        } else {
                            $stem = "${safeTarget}_$stem"
                        }
                    }
                    $cand = "$stem$ext"
                    if ($usedNames.Contains($cand)) { $cand = "${stem}_$fIdx$ext" }
                    [void]$usedNames.Add($cand)
                    $fullDest = "$destBase\$cand"
                    $moveList.Add("MOVE N'$($lName.Replace("'","''"))' TO N'$($fullDest.Replace("'","''"))'")
                    $fIdx++
                }
                $fr.Close()
            }

            $escTarget = "[" + $TargetDbName.Replace("]", "]]") + "]"
            $litTarget = "N'" + $TargetDbName.Replace("'", "''") + "'"
            $diskClauses = foreach ($p in $activePaths) { "    DISK = N'$($p.Replace("'","''"))'" }
            $fromClause = [string]::Join(",`r`n", $diskClauses)

            $baseOpts = New-Object System.Collections.Generic.List[string]
            foreach ($mv in $moveList) { $baseOpts.Add($mv) }
            if ($WithReplace) { $baseOpts.Add("REPLACE") }
            if ($KeepCdc) { $baseOpts.Add("KEEP_CDC") }
            $baseOpts.Add($RecoveryMode)

            $tunedOpts = New-Object System.Collections.Generic.List[string]($baseOpts)
            if ($HighSpeedTuning) {
                $bufCnt = if ($CustomBufferCount -gt 0) { $CustomBufferCount } else { [Math]::Min(128, [Math]::Max(24, $streamCount * 8)) }
                $tunedOpts.Add("MAXTRANSFERSIZE = 4194304")
                $tunedOpts.Add("BUFFERCOUNT = $bufCnt")
            }
            $tunedOpts.Add("STATS = 5")
            $baseOpts.Add("STATS = 5")

            $preSql = if ($KillConnections) { "IF DB_ID($litTarget) IS NOT NULL AND DATABASEPROPERTYEX($litTarget, 'Status') = 'ONLINE' ALTER DATABASE $escTarget SET SINGLE_USER WITH ROLLBACK IMMEDIATE;`r`n" } else { "" }
            $postSql = if ($KillConnections -and $RecoveryMode -eq "RECOVERY") { "`r`nIF DB_ID($litTarget) IS NOT NULL AND DATABASEPROPERTYEX($litTarget, 'Status') = 'ONLINE' ALTER DATABASE $escTarget SET MULTI_USER;" } else { "" }

            $sqlTuned = "${preSql}RESTORE DATABASE $escTarget`r`nFROM`r`n$fromClause`r`nWITH`r`n    $([string]::Join(",`r`n    ", $tunedOpts));${postSql}"
            $sqlFallback = "${preSql}RESTORE DATABASE $escTarget`r`nFROM`r`n$fromClause`r`nWITH`r`n    $([string]::Join(",`r`n    ", $baseOpts));${postSql}"

            try {
                $resCmd = $conn.CreateCommand()
                $resCmd.CommandTimeout = 0
                $resCmd.CommandText = $sqlTuned
                [void]$resCmd.ExecuteNonQuery()
            }
            catch {
                $exMsg = $_.Exception.GetBaseException().Message
                if ($HighSpeedTuning -and ($exMsg -match 'MAXTRANSFERSIZE' -or $exMsg -match 'BUFFERCOUNT' -or $exMsg -match 'memory' -or $exMsg -match '3204')) {
                    $EventQueue.Enqueue(@{
                        Type = "LOG"
                        SetId = $SetId
                        DbName = $TargetDbName
                        Message = "Buffer memory pressure detected; retrying with default SQL Server buffers..."
                    })
                    $resCmd2 = $conn.CreateCommand()
                    $resCmd2.CommandTimeout = 0
                    $resCmd2.CommandText = $sqlFallback
                    [void]$resCmd2.ExecuteNonQuery()
                }
                else {
                    throw
                }
            }
        }
        finally {
            if ($conn.State -ne 'Closed') { $conn.Close() }
            $conn.Dispose()
        }

        if ($SourceType -eq "zip" -and $CleanupStaging -and $extractDir -and (Test-Path -LiteralPath $extractDir)) {
            Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        $sw.Stop()
        $EventQueue.Enqueue(@{
            Type = "FINISHED"
            SetId = $SetId
            DbName = $TargetDbName
            Success = $true
            Elapsed = $sw.Elapsed.TotalSeconds
            UnzipElapsed = $unzipElapsed
            StreamCount = $streamCount
        })
    }
    catch {
        $sw.Stop()
        if ($SourceType -eq "zip" -and $CleanupStaging -and $extractDir -and (Test-Path -LiteralPath $extractDir)) {
            Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        $EventQueue.Enqueue(@{
            Type = "FINISHED"
            SetId = $SetId
            DbName = $TargetDbName
            Success = $false
            Elapsed = $sw.Elapsed.TotalSeconds
            StreamCount = $StripePaths.Count
            Error = ($_.Exception.GetBaseException().Message -replace "(\r|\n)+", " ")
        })
    }
}

$uiTimer = New-Object System.Windows.Forms.Timer
$uiTimer.Interval = 150
$uiTimer.Add_Tick({
    $evt = $null
    while ($Script:UiEventQueue.TryDequeue([ref]$evt)) {
        $setId = $evt.SetId
        $lvItem = $null
        foreach ($it in $lvSets.Items) {
            if ($it.Tag.SetId -eq $setId) { $lvItem = $it; break }
        }

        if ($evt.Type -eq "PROGRESS") {
            if ($evt.Percent -ge 0) {
                $Script:ProgressMap[$setId] = $evt.Percent
            }
            if ($null -ne $lvItem) {
                $lvItem.SubItems[5].Text = $evt.Status
            }
            if ($evt.Percent -in @(5, 15)) {
                Write-UiLog "[$($evt.DbName)] $($evt.Status)"
            }
        }
        elseif ($evt.Type -eq "LOG") {
            Write-UiLog "[$($evt.DbName)] $($evt.Message)"
        }
        elseif ($evt.Type -eq "FINISHED") {
            $Script:FinishedBatchCount++
            $Script:ProgressMap[$setId] = 100
            if ($evt.Success) {
                $uNote = if ($evt.UnzipElapsed -gt 0) { (", unzipped {0:N1}s" -f $evt.UnzipElapsed) } else { "" }
                $stText = ("Restored in {0:N1}s ({1} stripe(s){2})" -f $evt.Elapsed, $evt.StreamCount, $uNote)
                if ($null -ne $lvItem) { $lvItem.SubItems[5].Text = $stText }
                Write-UiLog ("SUCCESS [{0}] Restored in {1:N1}s ({2} stripe(s){3})" -f $evt.DbName, $evt.Elapsed, $evt.StreamCount, $uNote)
            }
            else {
                $stText = "Failed: $($evt.Error)"
                if ($null -ne $lvItem) { $lvItem.SubItems[5].Text = $stText }
                Write-UiLog "FAILED [$($evt.DbName)]: $($evt.Error)"
            }
        }

        if ($Script:TotalBatchCount -gt 0) {
            $sumPct = 0
            foreach ($v in $Script:ProgressMap.Values) { $sumPct += $v }
            $avg = [int][Math]::Min(100, [Math]::Round($sumPct / $Script:TotalBatchCount))
            $progressBar.Value = $avg
        }
    }

    if ($Script:IsRunning -and $Script:FinishedBatchCount -ge $Script:TotalBatchCount -and $Script:TotalBatchCount -gt 0) {
        $Script:IsRunning = $false
        $uiTimer.Stop()
        $btnStart.Enabled = $true
        $btnConnect.Enabled = $true
        $btnScan.Enabled = $true
        $btnCancel.Enabled = $false
        if ($null -ne $Script:RunspacePool) {
            $Script:RunspacePool.Close()
            $Script:RunspacePool.Dispose()
            $Script:RunspacePool = $null
        }
        Write-UiLog "================================================================"
        Write-UiLog "Parallel striped restore batch finished ($($Script:FinishedBatchCount) database(s) processed)."
    }
})

$btnStart.Add_Click({
    if ($Script:IsRunning) { return }
    $selectedItems = @($lvSets.CheckedItems)
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Please select at least one backup set to restore.", "No Backup Sets Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $connStr = New-SqlConnectionString -Server $txtServer.Text -AuthMode ([string]$cmbAuth.SelectedItem) -Username $txtUser.Text -Password $txtPass.Text -TrustServerCertificate $chkTrustCert.Checked
    $maxParallel = [int]$numParallelDbs.Value
    $customBuf = [int]$numBufCnt.Value
    $prefix = $txtPrefix.Text.Trim()
    $suffix = $txtSuffix.Text.Trim()
    $sourceDir = $txtSource.Text.Trim()
    $stagingRoot = "$($sourceDir.TrimEnd('\','/'))\$Script:StagingFolderName"

    # Sort selected items Largest-First (LPT)
    $orderedItems = @($selectedItems | Sort-Object { $_.Tag.TotalSizeBytes } -Descending)

    $Script:IsRunning = $true
    $Script:TotalBatchCount = $orderedItems.Count
    $Script:FinishedBatchCount = 0
    $Script:ProgressMap = @{}
    $Script:ActiveJobs.Clear()
    $progressBar.Value = 0

    $btnStart.Enabled = $false
    $btnConnect.Enabled = $false
    $btnScan.Enabled = $false
    $btnCancel.Enabled = $true

    Write-UiLog "Starting parallel striped restore batch for $($orderedItems.Count) database(s) (Max Parallel DBs: $maxParallel, Largest-First Queue)..."

    $Script:RunspacePool = [runspacefactory]::CreateRunspacePool(1, $maxParallel)
    $Script:RunspacePool.Open()
    $uiTimer.Start()

    foreach ($item in $orderedItems) {
        $bset = $item.Tag
        $targetDb = Get-TargetDatabaseName -SourceDbName $bset.SourceDbName -Prefix $prefix -Suffix $suffix -CustomTargetName $bset.CustomTargetName
        $item.SubItems[5].Text = "Queued (LPT)"
        $Script:ProgressMap[$bset.SetId] = 0

        $ps = [powershell]::Create()
        $ps.RunspacePool = $Script:RunspacePool
        [void]$ps.AddScript($restoreWorkerScriptBlock).AddArgument($connStr).AddArgument($bset.SetId).AddArgument($bset.SourceDbName).AddArgument($targetDb).AddArgument($bset.SourceType).AddArgument($bset.StripePaths).AddArgument($bset.ZipPath).AddArgument($stagingRoot).AddArgument($bset.TimestampLabel).AddArgument([bool]$chkRelocate.Checked).AddArgument($txtDataFolder.Text.Trim()).AddArgument($txtLogFolder.Text.Trim()).AddArgument([bool]$chkReplace.Checked).AddArgument([bool]$chkKillConn.Checked).AddArgument([string]$cmbRecMode.SelectedItem).AddArgument([bool]$chkHighSpeed.Checked).AddArgument($customBuf).AddArgument([bool]$chkKeepCdc.Checked).AddArgument([bool]$chkCleanupStaging.Checked).AddArgument($Script:UiEventQueue)
        $handle = $ps.BeginInvoke()
        [void]$Script:ActiveJobs.Add(@{ PowerShell = $ps; Handle = $handle })
    }
})

$btnCancel.Add_Click({
    if (-not $Script:IsRunning) { return }
    Write-UiLog "Cancelling active restore workers..."
    foreach ($job in $Script:ActiveJobs) {
        try { $job.PowerShell.Stop() } catch {}
        try { $job.PowerShell.Dispose() } catch {}
    }
    if ($null -ne $Script:RunspacePool) {
        try { $Script:RunspacePool.Close() } catch {}
        try { $Script:RunspacePool.Dispose() } catch {}
        $Script:RunspacePool = $null
    }
    $Script:IsRunning = $false
    $uiTimer.Stop()
    $btnStart.Enabled = $true
    $btnConnect.Enabled = $true
    $btnScan.Enabled = $true
    $btnCancel.Enabled = $false
})

# Load saved config on startup if available
if (Test-Path -LiteralPath $Script:ConfigFile) {
    try {
        $saved = Get-Content -LiteralPath $Script:ConfigFile -Raw | ConvertFrom-Json
        if ($saved.server) { $txtServer.Text = $saved.server }
        if ($saved.auth_mode) { $cmbAuth.SelectedItem = $saved.auth_mode }
        if ($saved.username) { $txtUser.Text = $saved.username }
        if ($null -ne $saved.trust_server_certificate) { $chkTrustCert.Checked = [bool]$saved.trust_server_certificate }
        if ($saved.source_folder) { $txtSource.Text = $saved.source_folder }
        if ($null -ne $saved.latest_only_per_db) { $chkLatestOnly.Checked = [bool]$saved.latest_only_per_db }
        if ($null -ne $saved.cleanup_staging_bak) { $chkCleanupStaging.Checked = [bool]$saved.cleanup_staging_bak }
        if ($null -ne $saved.relocate_files) { $chkRelocate.Checked = [bool]$saved.relocate_files }
        if ($saved.target_data_folder) { $txtDataFolder.Text = $saved.target_data_folder }
        if ($saved.target_log_folder) { $txtLogFolder.Text = $saved.target_log_folder }
        if ($saved.target_db_prefix) { $txtPrefix.Text = $saved.target_db_prefix }
        if ($saved.target_db_suffix) { $txtSuffix.Text = $saved.target_db_suffix }
        if ($saved.max_parallel_dbs) { $numParallelDbs.Value = [decimal]$saved.max_parallel_dbs }
        if ($null -ne $saved.with_replace) { $chkReplace.Checked = [bool]$saved.with_replace }
        if ($null -ne $saved.kill_connections) { $chkKillConn.Checked = [bool]$saved.kill_connections }
        if ($saved.recovery_mode) { $cmbRecMode.SelectedItem = $saved.recovery_mode }
        if ($null -ne $saved.high_speed_tuning) { $chkHighSpeed.Checked = [bool]$saved.high_speed_tuning }
        if ($null -ne $saved.custom_buffer_count) { $numBufCnt.Value = [decimal]$saved.custom_buffer_count }
        if ($null -ne $saved.keep_cdc) { $chkKeepCdc.Checked = [bool]$saved.keep_cdc }
    }
    catch {}
}

Invoke-ScanBackups
Update-SummaryAndPreview
[void]$form.ShowDialog()
