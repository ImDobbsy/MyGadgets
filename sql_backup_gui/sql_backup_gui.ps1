<#
.SYNOPSIS
    SQL Server Parallel Copy-Only Backup & Fast-Zip Utility - Native Windows GUI (No Python Required)
.DESCRIPTION
    Self-contained Windows Forms GUI for backing up user-selected SQL Server databases
    in COPY_ONLY mode to a target folder, with fast post-backup ZIP archiving.
    - Runs multiple database backups in parallel via PowerShell RunspacePool.
    - Automatically splits large databases (> 8 GB default) into parallel striped backup streams
      (multiple `TO DISK = N'..._stripe01of0N.bak'` clauses).
    - Places each database backup into its own subfolder (<TargetFolder>\<DbName>\) and creates
      the directory on the SQL Server host via `EXEC master.dbo.xp_create_subdir`.
    - Archives the resulting .bak file(s) into a target Zip folder using Fastest compression mode
      (`CompressionLevel::Fastest` or `7z.exe -tzip -mx=1 -mmt=on` when installed), prioritizing
      speed of zip over compression ratio.
    - Uses built-in .NET `System.Data.SqlClient` & `System.IO.Compression.FileSystem` (zero external dependencies).
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

$Script:AppTitle = "SQL Server Parallel Copy-Only Backup & Fast-Zip Utility (Server / Native Edition)"
$Script:ConfigFile = Join-Path $PSScriptRoot ".sql_backup_config.json"

# ============================================================================
# Pure Helper Functions (Stream Calculation, T-SQL & Fast-Zip Builders)
# ============================================================================

function Get-BackupStreamCount {
    param(
        [double]$SizeGB,
        [double]$ThresholdGB = 8.0,
        [int]$MaxStreams = 8
    )
    if ($ThresholdGB -le 0) { $ThresholdGB = 8.0 }
    if ($MaxStreams -lt 1) { $MaxStreams = 1 }
    if ($SizeGB -le $ThresholdGB -or $MaxStreams -eq 1) {
        return 1
    }
    $raw = [int][Math]::Ceiling($SizeGB / $ThresholdGB)
    if ($raw -lt 2) { $raw = 2 }
    if ($raw -gt $MaxStreams) { $raw = $MaxStreams }
    return $raw
}

function Get-SafeFileName {
    param([string]$Name)
    $cleaned = [regex]::Replace($Name.Trim(), '[<>:"/\\|?*\x00-\x1F]', '_')
    $cleaned = $cleaned.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($cleaned)) { return "database" }
    return $cleaned
}

function Get-BackupFilePaths {
    param(
        [string]$TargetFolder,
        [string]$DbName,
        [int]$StreamCount,
        [string]$Timestamp = ""
    )
    if ([string]::IsNullOrWhiteSpace($Timestamp)) {
        $Timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    }
    $base = $TargetFolder.Trim().TrimEnd('\', '/')
    $safeDb = Get-SafeFileName -Name $DbName
    $dbSubfolder = "$base\$safeDb"

    if ($StreamCount -lt 1) { $StreamCount = 1 }
    $paths = New-Object System.Collections.Generic.List[string]

    if ($StreamCount -eq 1) {
        $paths.Add("$dbSubfolder\${safeDb}_${Timestamp}.bak")
    }
    else {
        $width = [Math]::Max(2, $StreamCount.ToString().Length)
        $fmt = "D$width"
        $totalStr = $StreamCount.ToString($fmt)
        for ($i = 1; $i -le $StreamCount; $i++) {
            $idxStr = $i.ToString($fmt)
            $paths.Add("$dbSubfolder\${safeDb}_${Timestamp}_stripe${idxStr}of${totalStr}.bak")
        }
    }

    return [PSCustomObject]@{
        Subfolder = $dbSubfolder
        FilePaths = $paths.ToArray()
    }
}

function Get-ZipFilePath {
    param(
        [string]$TargetFolder,
        [string]$DbName,
        [string]$Timestamp = "",
        [string]$CustomZipFolder = ""
    )
    if ([string]::IsNullOrWhiteSpace($Timestamp)) {
        $Timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    }
    $safeDb = Get-SafeFileName -Name $DbName
    $zipFolder = if (-not [string]::IsNullOrWhiteSpace($CustomZipFolder)) {
        $CustomZipFolder.Trim().TrimEnd('\', '/')
    }
    else {
        $base = $TargetFolder.Trim().TrimEnd('\', '/')
        "$base\Zipped"
    }
    return [PSCustomObject]@{
        ZipFolder = $zipFolder
        ZipPath   = "$zipFolder\${safeDb}_${Timestamp}.zip"
    }
}

function Invoke-FastZipBackupFiles {
    param(
        [string[]]$FilePaths,
        [string]$ZipPath,
        [bool]$DeleteSourceAfterZip = $false,
        [string]$SourceSubfolder = ""
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $existing = @($FilePaths | Where-Object { Test-Path -LiteralPath $_ })
    if ($existing.Count -eq 0) {
        throw "Cannot create Zip archive: .bak file(s) not found at $($FilePaths[0])"
    }

    $zipDir = Split-Path -Parent $ZipPath
    if ($zipDir -and (-not (Test-Path -LiteralPath $zipDir))) {
        New-Item -ItemType Directory -Force -Path $zipDir | Out-Null
    }
    if (Test-Path -LiteralPath $ZipPath) {
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
    }

    $archive = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($src in $existing) {
            $entryName = [System.IO.Path]::GetFileName($src)
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $src,
                $entryName,
                [System.IO.Compression.CompressionLevel]::Fastest
            )
        }
    }
    finally {
        $archive.Dispose()
    }

    if (-not (Test-Path -LiteralPath $ZipPath)) {
        throw "Zip verification failed: $ZipPath was not created."
    }

    if ($DeleteSourceAfterZip) {
        foreach ($src in $existing) {
            Remove-Item -LiteralPath $src -Force -ErrorAction SilentlyContinue
        }
        if ($SourceSubfolder -and (Test-Path -LiteralPath $SourceSubfolder)) {
            $remaining = @(Get-ChildItem -LiteralPath $SourceSubfolder -Force -ErrorAction SilentlyContinue)
            if ($remaining.Count -eq 0) {
                Remove-Item -LiteralPath $SourceSubfolder -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $sw.Stop()
    return [PSCustomObject]@{
        ZipPath        = $ZipPath
        ElapsedSeconds = $sw.Elapsed.TotalSeconds
    }
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

function Get-BackupDatabaseTsql {
    param(
        [string]$DbName,
        [string[]]$FilePaths,
        [bool]$CopyOnly = $true,
        [bool]$Compression = $true,
        [bool]$Checksum = $true,
        [int]$StatsInterval = 5
    )
    $escDb = ConvertTo-TsqlIdentifier -Identifier $DbName
    $diskLines = foreach ($p in $FilePaths) {
        "    DISK = $(ConvertTo-TsqlString -Value $p)"
    }
    $toClause = [string]::Join(",`r`n", $diskLines)

    $opts = New-Object System.Collections.Generic.List[string]
    if ($CopyOnly) { $opts.Add("COPY_ONLY") }
    if ($Compression) { $opts.Add("COMPRESSION") }
    if ($Checksum) { $opts.Add("CHECKSUM") }
    $opts.Add("INIT")
    $statsVal = [Math]::Max(1, [Math]::Min(100, $StatsInterval))
    $opts.Add("STATS = $statsVal")

    $withClause = [string]::Join(",`r`n    ", $opts)
    return "BACKUP DATABASE $escDb`r`nTO`r`n$toClause`r`nWITH`r`n    $withClause;"
}

function Get-VerifyBackupTsql {
    param(
        [string[]]$FilePaths,
        [bool]$Checksum = $true
    )
    $diskLines = foreach ($p in $FilePaths) {
        "    DISK = $(ConvertTo-TsqlString -Value $p)"
    }
    $fromClause = [string]::Join(",`r`n", $diskLines)
    $withSuffix = if ($Checksum) { "`r`nWITH CHECKSUM;" } else { ";" }
    return "RESTORE VERIFYONLY`r`nFROM`r`n$fromClause$withSuffix"
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
    $parts.Add("Application Name=SqlParallelCopyOnlyBackupPsGui")
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
$form.Size = New-Object System.Drawing.Size(1200, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# State containers
$Script:DatabaseList = New-Object System.Collections.ArrayList
$Script:UiEventQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$Script:ActiveJobs = New-Object System.Collections.ArrayList
$Script:RunspacePool = $null
$Script:IsRunning = $false
$Script:TotalBatchCount = 0
$Script:FinishedBatchCount = 0
$Script:ProgressMap = @{}
$Script:UserEditedZipFolder = $false

# Top Layout Container
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$mainPanel.ColumnCount = 1
$mainPanel.RowCount = 5
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 122)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 56)))
[void]$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 44)))
$form.Controls.Add($mainPanel)

# 0. Header Banner
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Fill"
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "SQL Server Parallel Copy-Only Backup & Fast-Zip Utility"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(2, 2)
$headerPanel.Controls.Add($lblTitle)

$lblSubTitle = New-Object System.Windows.Forms.Label
$lblSubTitle.Text = "Parallel DB Workers  |  Auto Multi-Stream Striping (> 8 GB)  |  Fastest-Speed Zip Archive"
$lblSubTitle.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblSubTitle.AutoSize = $true
$lblSubTitle.Anchor = "Top, Right"
$lblSubTitle.Location = New-Object System.Drawing.Point(650, 6)
$headerPanel.Controls.Add($lblSubTitle)
$mainPanel.Controls.Add($headerPanel, 0, 0)

# 1. SQL Server Connection GroupBox
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = " 1. SQL Server Connection "
$grpConn.Dock = "Fill"
$mainPanel.Controls.Add($grpConn, 0, 1)

$lblSrv = New-Object System.Windows.Forms.Label
$lblSrv.Text = "Server / Instance:"
$lblSrv.Location = New-Object System.Drawing.Point(12, 24)
$lblSrv.AutoSize = $true
$grpConn.Controls.Add($lblSrv)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Text = "localhost"
$txtServer.Location = New-Object System.Drawing.Point(116, 21)
$txtServer.Width = 195
$grpConn.Controls.Add($txtServer)

$lblAuth = New-Object System.Windows.Forms.Label
$lblAuth.Text = "Auth:"
$lblAuth.Location = New-Object System.Drawing.Point(322, 24)
$lblAuth.AutoSize = $true
$grpConn.Controls.Add($lblAuth)

$cmbAuth = New-Object System.Windows.Forms.ComboBox
$cmbAuth.DropDownStyle = "DropDownList"
[void]$cmbAuth.Items.AddRange(@("windows", "sql"))
$cmbAuth.SelectedIndex = 0
$cmbAuth.Location = New-Object System.Drawing.Point(360, 21)
$cmbAuth.Width = 95
$grpConn.Controls.Add($cmbAuth)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Login:"
$lblUser.Location = New-Object System.Drawing.Point(466, 24)
$lblUser.AutoSize = $true
$grpConn.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = "sa"
$txtUser.Enabled = $false
$txtUser.Location = New-Object System.Drawing.Point(508, 21)
$txtUser.Width = 110
$grpConn.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.Location = New-Object System.Drawing.Point(626, 24)
$lblPass.AutoSize = $true
$grpConn.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.UseSystemPasswordChar = $true
$txtPass.Enabled = $false
$txtPass.Location = New-Object System.Drawing.Point(690, 21)
$txtPass.Width = 115
$grpConn.Controls.Add($txtPass)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Show"
$chkShowPass.Location = New-Object System.Drawing.Point(812, 23)
$chkShowPass.AutoSize = $true
$grpConn.Controls.Add($chkShowPass)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect & Load DBs"
$btnConnect.Location = New-Object System.Drawing.Point(875, 19)
$btnConnect.Size = New-Object System.Drawing.Size(145, 27)
$grpConn.Controls.Add($btnConnect)

$btnSaveCfg = New-Object System.Windows.Forms.Button
$btnSaveCfg.Text = "Save Cfg"
$btnSaveCfg.Location = New-Object System.Drawing.Point(1028, 19)
$btnSaveCfg.Size = New-Object System.Drawing.Size(70, 27)
$grpConn.Controls.Add($btnSaveCfg)

$chkTrustCert = New-Object System.Windows.Forms.CheckBox
$chkTrustCert.Text = "Trust Server Certificate"
$chkTrustCert.Checked = $true
$chkTrustCert.Location = New-Object System.Drawing.Point(15, 55)
$chkTrustCert.AutoSize = $true
$grpConn.Controls.Add($chkTrustCert)

$chkIncludeSys = New-Object System.Windows.Forms.CheckBox
$chkIncludeSys.Text = "Include System DBs (master, model, msdb)"
$chkIncludeSys.Checked = $false
$chkIncludeSys.Location = New-Object System.Drawing.Point(190, 55)
$chkIncludeSys.AutoSize = $true
$grpConn.Controls.Add($chkIncludeSys)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = "Not connected -- enter SQL Server instance and click 'Connect & Load DBs'"
$lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 86, 160)
$lblConnStatus.Location = New-Object System.Drawing.Point(475, 56)
$lblConnStatus.AutoSize = $true
$grpConn.Controls.Add($lblConnStatus)

# 2. Target Folders, Parallelism & Fast-Zip GroupBox
$grpSettings = New-Object System.Windows.Forms.GroupBox
$grpSettings.Text = " 2. Target Folders, Parallelism, Multi-Stream Striping (> 8 GB) & Fast-Zip Options "
$grpSettings.Dock = "Fill"
$mainPanel.Controls.Add($grpSettings, 0, 2)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target Backup Folder (.bak):"
$lblTarget.Location = New-Object System.Drawing.Point(12, 24)
$lblTarget.AutoSize = $true
$grpSettings.Controls.Add($lblTarget)

$defaultTargetDir = "C:\SqlBackups"
$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Text = $defaultTargetDir
$txtTarget.Location = New-Object System.Drawing.Point(180, 21)
$txtTarget.Width = 760
$txtTarget.Anchor = "Top, Left, Right"
$grpSettings.Controls.Add($txtTarget)

$btnBrowseTarget = New-Object System.Windows.Forms.Button
$btnBrowseTarget.Text = "Browse Folder..."
$btnBrowseTarget.Location = New-Object System.Drawing.Point(950, 19)
$btnBrowseTarget.Size = New-Object System.Drawing.Size(135, 26)
$btnBrowseTarget.Anchor = "Top, Right"
$grpSettings.Controls.Add($btnBrowseTarget)

# Fast-Zip Row
$chkZipAfter = New-Object System.Windows.Forms.CheckBox
$chkZipAfter.Text = "Zip Backups into Folder (Fastest Speed):"
$chkZipAfter.Checked = $true
$chkZipAfter.Location = New-Object System.Drawing.Point(14, 54)
$chkZipAfter.AutoSize = $true
$grpSettings.Controls.Add($chkZipAfter)

$txtZipFolder = New-Object System.Windows.Forms.TextBox
$txtZipFolder.Text = (Join-Path $defaultTargetDir "Zipped")
$txtZipFolder.Location = New-Object System.Drawing.Point(255, 52)
$txtZipFolder.Width = 490
$txtZipFolder.Anchor = "Top, Left, Right"
$grpSettings.Controls.Add($txtZipFolder)

$btnBrowseZip = New-Object System.Windows.Forms.Button
$btnBrowseZip.Text = "Browse Zip Folder..."
$btnBrowseZip.Location = New-Object System.Drawing.Point(755, 50)
$btnBrowseZip.Size = New-Object System.Drawing.Size(135, 26)
$btnBrowseZip.Anchor = "Top, Right"
$grpSettings.Controls.Add($btnBrowseZip)

$chkDelBakAfterZip = New-Object System.Windows.Forms.CheckBox
$chkDelBakAfterZip.Text = "Delete raw .bak after Zip"
$chkDelBakAfterZip.Checked = $false
$chkDelBakAfterZip.Location = New-Object System.Drawing.Point(905, 54)
$chkDelBakAfterZip.Anchor = "Top, Right"
$chkDelBakAfterZip.AutoSize = $true
$grpSettings.Controls.Add($chkDelBakAfterZip)

# Parallelism & SQL Options Row
$lblParallel = New-Object System.Windows.Forms.Label
$lblParallel.Text = "Max Parallel DBs:"
$lblParallel.Location = New-Object System.Drawing.Point(12, 87)
$lblParallel.AutoSize = $true
$grpSettings.Controls.Add($lblParallel)

$numParallelDbs = New-Object System.Windows.Forms.NumericUpDown
$numParallelDbs.Minimum = 1
$numParallelDbs.Maximum = 16
$numParallelDbs.Value = 3
$numParallelDbs.Location = New-Object System.Drawing.Point(120, 84)
$numParallelDbs.Width = 52
$grpSettings.Controls.Add($numParallelDbs)

$lblThresh = New-Object System.Windows.Forms.Label
$lblThresh.Text = "Split Threshold (GB):"
$lblThresh.Location = New-Object System.Drawing.Point(188, 87)
$lblThresh.AutoSize = $true
$grpSettings.Controls.Add($lblThresh)

$numThresholdGb = New-Object System.Windows.Forms.NumericUpDown
$numThresholdGb.Minimum = 1
$numThresholdGb.Maximum = 512
$numThresholdGb.DecimalPlaces = 1
$numThresholdGb.Increment = 1.0
$numThresholdGb.Value = 8.0
$numThresholdGb.Location = New-Object System.Drawing.Point(312, 84)
$numThresholdGb.Width = 60
$grpSettings.Controls.Add($numThresholdGb)

$lblMaxStreams = New-Object System.Windows.Forms.Label
$lblMaxStreams.Text = "Max Streams (>8GB):"
$lblMaxStreams.Location = New-Object System.Drawing.Point(388, 87)
$lblMaxStreams.AutoSize = $true
$grpSettings.Controls.Add($lblMaxStreams)

$numMaxStreams = New-Object System.Windows.Forms.NumericUpDown
$numMaxStreams.Minimum = 1
$numMaxStreams.Maximum = 64
$numMaxStreams.Value = 8
$numMaxStreams.Location = New-Object System.Drawing.Point(516, 84)
$numMaxStreams.Width = 52
$grpSettings.Controls.Add($numMaxStreams)

$chkCopyOnly = New-Object System.Windows.Forms.CheckBox
$chkCopyOnly.Text = "COPY_ONLY (Always On)"
$chkCopyOnly.Checked = $true
$chkCopyOnly.Enabled = $false
$chkCopyOnly.Location = New-Object System.Drawing.Point(585, 86)
$chkCopyOnly.AutoSize = $true
$grpSettings.Controls.Add($chkCopyOnly)

$chkCompression = New-Object System.Windows.Forms.CheckBox
$chkCompression.Text = "COMPRESSION"
$chkCompression.Checked = $true
$chkCompression.Location = New-Object System.Drawing.Point(755, 86)
$chkCompression.AutoSize = $true
$grpSettings.Controls.Add($chkCompression)

$chkChecksum = New-Object System.Windows.Forms.CheckBox
$chkChecksum.Text = "CHECKSUM"
$chkChecksum.Checked = $true
$chkChecksum.Location = New-Object System.Drawing.Point(875, 86)
$chkChecksum.AutoSize = $true
$grpSettings.Controls.Add($chkChecksum)

$chkVerify = New-Object System.Windows.Forms.CheckBox
$chkVerify.Text = "VERIFYONLY"
$chkVerify.Checked = $false
$chkVerify.Location = New-Object System.Drawing.Point(975, 86)
$chkVerify.AutoSize = $true
$grpSettings.Controls.Add($chkVerify)

# 3. Database Selection GroupBox
$grpDbs = New-Object System.Windows.Forms.GroupBox
$grpDbs.Text = " 3. Select Databases to Backup "
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

$btnSelLarge = New-Object System.Windows.Forms.Button
$btnSelLarge.Text = "Select Large (> Threshold)"
$btnSelLarge.Location = New-Object System.Drawing.Point(188, 3)
$btnSelLarge.Size = New-Object System.Drawing.Size(165, 26)
$dbToolbar.Controls.Add($btnSelLarge)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(368, 7)
$lblFilter.AutoSize = $true
$dbToolbar.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(410, 4)
$txtFilter.Width = 160
$dbToolbar.Controls.Add($txtFilter)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "Selected: 0 DBs (0.00 GB | 0 total streams)"
$lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSummary.Location = New-Object System.Drawing.Point(590, 7)
$lblSummary.AutoSize = $true
$dbToolbar.Controls.Add($lblSummary)

$lvDbs = New-Object System.Windows.Forms.ListView
$lvDbs.Dock = "Fill"
$lvDbs.View = "Details"
$lvDbs.CheckBoxes = $true
$lvDbs.FullRowSelect = $true
$lvDbs.GridLines = $true
[void]$lvDbs.Columns.Add("Database Name", 220)
[void]$lvDbs.Columns.Add("Size (GB)", 100)
[void]$lvDbs.Columns.Add("Streams (Stripes)", 140)
[void]$lvDbs.Columns.Add("Recovery Model", 110)
[void]$lvDbs.Columns.Add("Status / Progress", 225)
[void]$lvDbs.Columns.Add("Target .bak / .zip Output", 310)
$dbLayout.Controls.Add($lvDbs, 0, 1)

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
$splitBottom.SplitterDistance = 550
$bottomLayout.Controls.Add($splitBottom, 0, 0)

$grpPreview = New-Object System.Windows.Forms.GroupBox
$grpPreview.Text = " Live T-SQL & Fast-Zip Pipeline Preview "
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
$grpLog.Text = " Parallel Backup & Fast-Zip Execution Log "
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
$btnStart.Text = "Start Parallel Copy-Only Backups"
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnStart.Location = New-Object System.Drawing.Point(0, 4)
$btnStart.Size = New-Object System.Drawing.Size(230, 28)
$actionPanel.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Backups"
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
$progressBar.Size = New-Object System.Drawing.Size(685, 24)
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
    $thresh = [double]$numThresholdGb.Value
    $maxStreams = [int]$numMaxStreams.Value
    $maxParallel = [int]$numParallelDbs.Value
    $targetFolder = $txtTarget.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetFolder)) { $targetFolder = "C:\SqlBackups" }
    $zipEnabled = $chkZipAfter.Checked
    $customZip = $txtZipFolder.Text.Trim()

    $checkedItems = New-Object System.Collections.ArrayList
    $totalGb = 0.0
    $totalStreams = 0
    $largeCount = 0

    foreach ($item in $lvDbs.Items) {
        $dbObj = $item.Tag
        $streams = Get-BackupStreamCount -SizeGB $dbObj.SizeGB -ThresholdGB $thresh -MaxStreams $maxStreams
        $streamLabel = if ($streams -eq 1) { "1 stream" } else { "$streams striped streams" }
        $pathInfo = Get-BackupFilePaths -TargetFolder $targetFolder -DbName $dbObj.Name -StreamCount $streams -Timestamp "YYYYMMDD"
        $item.SubItems[2].Text = $streamLabel
        if ($zipEnabled) {
            $zInfo = Get-ZipFilePath -TargetFolder $targetFolder -DbName $dbObj.Name -Timestamp "YYYYMMDD" -CustomZipFolder $customZip
            $item.SubItems[5].Text = "$($pathInfo.Subfolder) -> $($zInfo.ZipPath)"
        }
        else {
            $item.SubItems[5].Text = $pathInfo.Subfolder
        }

        if ($item.Checked) {
            [void]$checkedItems.Add($dbObj)
            $totalGb += $dbObj.SizeGB
            $totalStreams += $streams
            if ($dbObj.SizeGB -gt $thresh) { $largeCount++ }
        }
    }

    $zipBadge = if ($zipEnabled) { " + Fast-Zip" } else { "" }
    $lblSummary.Text = ("Selected: {0} DBs ({1:N2} GB  |  {2} > {3:G} GB  |  {4} streams{5})" -f $checkedItems.Count, $totalGb, $largeCount, $thresh, $totalStreams, $zipBadge)

    # Build T-SQL preview
    $tsNow = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $previewDbs = New-Object System.Collections.ArrayList
    if ($lvDbs.SelectedItems.Count -gt 0) {
        [void]$previewDbs.Add($lvDbs.SelectedItems[0].Tag)
    }
    foreach ($d in $checkedItems) {
        if (-not $previewDbs.Contains($d)) {
            [void]$previewDbs.Add($d)
        }
    }

    if ($previewDbs.Count -eq 0) {
        $previewDbs.Add([PSCustomObject]@{ Name = "AppConfigDB"; SizeGB = 2.0 }) | Out-Null
        $previewDbs.Add([PSCustomObject]@{ Name = "EnterpriseWarehouseDB"; SizeGB = 28.0 }) | Out-Null
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("-- Parallel Workers: $maxParallel concurrent DBs | Split Threshold: > $thresh GB (up to $maxStreams streams)")
    if ($zipEnabled) {
        [void]$sb.AppendLine("-- Post-Backup Fast-Zip: ENABLED (CompressionLevel::Fastest -> $customZip)")
    }
    else {
        [void]$sb.AppendLine("-- Post-Backup Fast-Zip: DISABLED")
    }
    [void]$sb.AppendLine("")

    $limit = [Math]::Min(5, $previewDbs.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $db = $previewDbs[$i]
        $st = Get-BackupStreamCount -SizeGB $db.SizeGB -ThresholdGB $thresh -MaxStreams $maxStreams
        $pInfo = Get-BackupFilePaths -TargetFolder $targetFolder -DbName $db.Name -StreamCount $st -Timestamp $tsNow
        $mkdirSql = Get-CreateSubdirTsql -SubfolderPath $pInfo.Subfolder
        $bakSql = Get-BackupDatabaseTsql -DbName $db.Name -FilePaths $pInfo.FilePaths -CopyOnly $true -Compression $chkCompression.Checked -Checksum $chkChecksum.Checked -StatsInterval 5
        [void]$sb.AppendLine(("-- Database: [{0}] ({1:N2} GB) -> {2} stream(s)" -f $db.Name, $db.SizeGB, $st))
        [void]$sb.AppendLine($mkdirSql)
        [void]$sb.AppendLine($bakSql)
        if ($chkVerify.Checked) {
            [void]$sb.AppendLine((Get-VerifyBackupTsql -FilePaths $pInfo.FilePaths -Checksum $chkChecksum.Checked))
        }
        if ($zipEnabled) {
            $zInfo = Get-ZipFilePath -TargetFolder $targetFolder -DbName $db.Name -Timestamp $tsNow -CustomZipFolder $customZip
            [void]$sb.AppendLine("-- Fast-Zip Output (Fastest Mode): $($zInfo.ZipPath)")
        }
        [void]$sb.AppendLine("")
    }
    $txtPreview.Text = $sb.ToString()
}

function Populate-DatabaseListView {
    $lvDbs.BeginUpdate()
    try {
        $lvDbs.Items.Clear()
        $filter = $txtFilter.Text.Trim().ToLowerInvariant()
        $thresh = [double]$numThresholdGb.Value
        $maxStreams = [int]$numMaxStreams.Value
        $targetFolder = $txtTarget.Text.Trim()

        foreach ($db in $Script:DatabaseList) {
            if ($filter -and (-not $db.Name.ToLowerInvariant().Contains($filter))) {
                continue
            }
            $streams = Get-BackupStreamCount -SizeGB $db.SizeGB -ThresholdGB $thresh -MaxStreams $maxStreams
            $streamLabel = if ($streams -eq 1) { "1 stream" } else { "$streams striped streams" }
            $pathInfo = Get-BackupFilePaths -TargetFolder $targetFolder -DbName $db.Name -StreamCount $streams -Timestamp "YYYYMMDD"

            $item = New-Object System.Windows.Forms.ListViewItem($db.Name)
            $item.Checked = (-not $db.IsSystemDb)
            [void]$item.SubItems.Add(("{0:N2} GB" -f $db.SizeGB))
            [void]$item.SubItems.Add($streamLabel)
            [void]$item.SubItems.Add($db.RecoveryModel)
            [void]$item.SubItems.Add("Ready")
            [void]$item.SubItems.Add($pathInfo.Subfolder)
            $item.Tag = $db
            [void]$lvDbs.Items.Add($item)
        }
    }
    finally {
        $lvDbs.EndUpdate()
    }
    Update-SummaryAndPreview
}

$cmbAuth.Add_SelectedIndexChanged({
    $isSql = ($cmbAuth.SelectedItem -eq "sql")
    $txtUser.Enabled = $isSql
    $txtPass.Enabled = $isSql
})

$chkShowPass.Add_CheckedChanged({
    $txtPass.UseSystemPasswordChar = (-not $chkShowPass.Checked)
})

$chkZipAfter.Add_CheckedChanged({
    $en = $chkZipAfter.Checked
    $txtZipFolder.Enabled = $en
    $btnBrowseZip.Enabled = $en
    $chkDelBakAfterZip.Enabled = $en
    Update-SummaryAndPreview
})

$btnBrowseTarget.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Target Backup Folder (.bak)"
    $dlg.SelectedPath = $txtTarget.Text
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTarget.Text = $dlg.SelectedPath
        Update-SummaryAndPreview
    }
})

$btnBrowseZip.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Target Zip Output Folder"
    $dlg.SelectedPath = $txtZipFolder.Text
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:UserEditedZipFolder = $true
        $txtZipFolder.Text = $dlg.SelectedPath
        Update-SummaryAndPreview
    }
})

$txtTarget.Add_TextChanged({
    if (-not $Script:UserEditedZipFolder) {
        $base = $txtTarget.Text.Trim().TrimEnd('\', '/')
        if ($base) { $txtZipFolder.Text = "$base\Zipped" }
    }
    Update-SummaryAndPreview
})
$txtZipFolder.Add_KeyUp({
    $Script:UserEditedZipFolder = $true
    Update-SummaryAndPreview
})
$chkDelBakAfterZip.Add_CheckedChanged({ Update-SummaryAndPreview })
$numThresholdGb.Add_ValueChanged({ Update-SummaryAndPreview })
$numMaxStreams.Add_ValueChanged({ Update-SummaryAndPreview })
$numParallelDbs.Add_ValueChanged({ Update-SummaryAndPreview })
$chkCompression.Add_CheckedChanged({ Update-SummaryAndPreview })
$chkChecksum.Add_CheckedChanged({ Update-SummaryAndPreview })
$chkVerify.Add_CheckedChanged({ Update-SummaryAndPreview })
$txtFilter.Add_TextChanged({ Populate-DatabaseListView })
$lvDbs.Add_ItemChecked({ Update-SummaryAndPreview })
$lvDbs.Add_SelectedIndexChanged({ Update-SummaryAndPreview })

$btnSelAll.Add_Click({
    foreach ($item in $lvDbs.Items) { $item.Checked = $true }
    Update-SummaryAndPreview
})

$btnDeselAll.Add_Click({
    foreach ($item in $lvDbs.Items) { $item.Checked = $false }
    Update-SummaryAndPreview
})

$btnSelLarge.Add_Click({
    $thresh = [double]$numThresholdGb.Value
    foreach ($item in $lvDbs.Items) {
        $item.Checked = ($item.Tag.SizeGB -gt $thresh)
    }
    Update-SummaryAndPreview
})

$btnCopyTsql.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($txtPreview.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($txtPreview.Text)
        Write-UiLog "Copied T-SQL preview to clipboard."
    }
})

$btnSaveCfg.Add_Click({
    $cfg = @{
        server = $txtServer.Text.Trim()
        auth_mode = [string]$cmbAuth.SelectedItem
        username = $txtUser.Text.Trim()
        trust_server_certificate = [bool]$chkTrustCert.Checked
        include_system_dbs = [bool]$chkIncludeSys.Checked
        target_folder = $txtTarget.Text.Trim()
        zip_after_backup = [bool]$chkZipAfter.Checked
        zip_output_folder = $txtZipFolder.Text.Trim()
        delete_bak_after_zip = [bool]$chkDelBakAfterZip.Checked
        max_parallel_dbs = [int]$numParallelDbs.Value
        split_threshold_gb = [double]$numThresholdGb.Value
        max_streams_per_db = [int]$numMaxStreams.Value
        compression = [bool]$chkCompression.Checked
        checksum = [bool]$chkChecksum.Checked
        verify_only_after = [bool]$chkVerify.Checked
    }
    $cfg | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Script:ConfigFile -Encoding UTF8
    Write-UiLog "Saved configuration to .sql_backup_config.json"
})

# Connect & Fetch Databases
$btnConnect.Add_Click({
    if ($Script:IsRunning) { return }
    $connStr = New-SqlConnectionString -Server $txtServer.Text -AuthMode ([string]$cmbAuth.SelectedItem) -Username $txtUser.Text -Password $txtPass.Text -TrustServerCertificate $chkTrustCert.Checked
    $includeSys = $chkIncludeSys.Checked
    Write-UiLog "Connecting to SQL Server '$($txtServer.Text)'..."

    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    try {
        $conn.Open()
        $metaCmd = $conn.CreateCommand()
        $metaCmd.CommandText = "SELECT CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(256)) AS srv, CAST(SERVERPROPERTY('Edition') AS NVARCHAR(256)) AS ed, CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS ver;"
        $mr = $metaCmd.ExecuteReader()
        $srvSummary = $txtServer.Text
        if ($mr.Read()) {
            $srvSummary = "$($mr['srv']) -- $($mr['ed']) (v$($mr['ver']))"
        }
        $mr.Close()

        $dbCmd = $conn.CreateCommand()
        $dbCmd.CommandText = @"
SELECT
    d.name AS database_name,
    CAST( ROUND(SUM(CAST(mf.size AS BIGINT)) * 8.0 / 1024.0, 2) AS FLOAT ) AS size_mb,
    d.recovery_model_desc,
    d.state_desc,
    CASE WHEN d.database_id <= 4 THEN 1 ELSE 0 END AS is_system_db
FROM sys.databases AS d
LEFT JOIN sys.master_files AS mf ON d.database_id = mf.database_id
WHERE d.state_desc = 'ONLINE' AND d.name <> 'tempdb' AND d.source_database_id IS NULL
GROUP BY d.name, d.recovery_model_desc, d.state_desc, d.database_id
ORDER BY d.name ASC;
"@
        $dr = $dbCmd.ExecuteReader()
        $Script:DatabaseList.Clear()
        while ($dr.Read()) {
            $isSys = ([int]$dr['is_system_db'] -eq 1)
            if (-not $includeSys -and $isSys) { continue }
            $sizeMb = [double]$dr['size_mb']
            [void]$Script:DatabaseList.Add([PSCustomObject]@{
                Name = [string]$dr['database_name']
                SizeMB = $sizeMb
                SizeGB = [Math]::Round($sizeMb / 1024.0, 3)
                RecoveryModel = [string]$dr['recovery_model_desc']
                IsSystemDb = $isSys
            })
        }
        $dr.Close()

        $lblConnStatus.Text = "Connected: $srvSummary ($($Script:DatabaseList.Count) online DBs)"
        Write-UiLog "Connected to $srvSummary. Loaded $($Script:DatabaseList.Count) online database(s)."
        Populate-DatabaseListView
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

# Parallel Backup & Fast-Zip Worker ScriptBlock
$workerScriptBlock = {
    param(
        [string]$ConnStr,
        [string]$DbName,
        [string]$Subfolder,
        [string[]]$FilePaths,
        [string]$CreateDirSql,
        [string]$BackupSql,
        [string]$BackupFallbackSql,
        [string]$VerifySql,
        [bool]$AllowFallback,
        [int]$StreamCount,
        [bool]$ZipAfterBackup,
        [string]$ZipPath,
        [bool]$DeleteBakAfterZip,
        $EventQueue
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $EventQueue.Enqueue(@{
        Type = "PROGRESS"
        DbName = $DbName
        Percent = 0
        Status = "Starting ($StreamCount stream(s))..."
    })

    try {
        if (-not (Test-Path -LiteralPath $Subfolder)) {
            New-Item -ItemType Directory -Force -Path $Subfolder -ErrorAction SilentlyContinue | Out-Null
        }
        if (Test-Path -LiteralPath $Subfolder) {
            $parentDir = Split-Path -Parent $Subfolder
            if ($parentDir -and (Test-Path -LiteralPath $parentDir)) {
                & icacls $parentDir /grant "*S-1-5-32-545:(OI)(CI)M" /grant "*S-1-5-11:(OI)(CI)M" /Q 2>$null | Out-Null
            }
            & icacls $Subfolder /grant "*S-1-5-32-545:(OI)(CI)M" /grant "*S-1-5-11:(OI)(CI)M" /Q 2>$null | Out-Null
        }
    }
    catch {}

    $conn = New-Object System.Data.SqlClient.SqlConnection($ConnStr)
    $conn.FireInfoMessageEventOnUserErrors = $false

    $msgHandler = [System.Data.SqlClient.SqlInfoMessageEventHandler]{
        param($sender, $args)
        foreach ($err in $args.Errors) {
            if ($null -ne $err -and $err.Message -match '(\d+)\s+percent\s+processed') {
                $pct = [int]$Matches[1]
                $scaled = if ($ZipAfterBackup) { [int]($pct * 0.85) } else { $pct }
                $EventQueue.Enqueue(@{
                    Type = "PROGRESS"
                    DbName = $DbName
                    Percent = $scaled
                    Status = "SQL Backup ($pct% | $StreamCount stream(s))"
                })
            }
        }
    }
    $conn.add_InfoMessage($msgHandler)

    try {
        $conn.Open()

        if (-not (Test-Path -LiteralPath $Subfolder)) {
            try {
                $dirCmd = $conn.CreateCommand()
                $dirCmd.CommandTimeout = 60
                $dirCmd.CommandText = $CreateDirSql
                [void]$dirCmd.ExecuteNonQuery()
            }
            catch {
                # Ignore xp_create_subdir Error 183 ('Cannot create a file when that file already exists')
            }
        }

        try {
            $bakCmd = $conn.CreateCommand()
            $bakCmd.CommandTimeout = 0
            $bakCmd.CommandText = $BackupSql
            [void]$bakCmd.ExecuteNonQuery()
        }
        catch {
            $exMsg = $_.Exception.GetBaseException().Message
            if ($AllowFallback -and ($exMsg -match 'COMPRESSION is not supported' -or $exMsg -match '1844')) {
                $EventQueue.Enqueue(@{
                    Type = "LOG"
                    DbName = $DbName
                    Message = "Compression not supported by SQL edition; retrying without COMPRESSION..."
                })
                $bakCmd2 = $conn.CreateCommand()
                $bakCmd2.CommandTimeout = 0
                $bakCmd2.CommandText = $BackupFallbackSql
                [void]$bakCmd2.ExecuteNonQuery()
            }
            else {
                throw
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($VerifySql)) {
            $verCmd = $conn.CreateCommand()
            $verCmd.CommandTimeout = 0
            $verCmd.CommandText = $VerifySql
            [void]$verCmd.ExecuteNonQuery()
        }

        # Step 4: Fast-Zip (CompressionLevel::Fastest or 7z -mx=1 -mmt=on)
        $zipElapsed = 0.0
        $finalOutput = $FilePaths[0]
        if ($ZipAfterBackup) {
            $EventQueue.Enqueue(@{
                Type = "PROGRESS"
                DbName = $DbName
                Percent = 90
                Status = "Fast-Zipping .bak file(s)..."
            })
            $zSw = [System.Diagnostics.Stopwatch]::StartNew()
            $zipDir = Split-Path -Parent $ZipPath
            if ($zipDir -and (-not (Test-Path -LiteralPath $zipDir))) {
                New-Item -ItemType Directory -Force -Path $zipDir | Out-Null
            }
            if (Test-Path -LiteralPath $ZipPath) {
                Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            }

            $sevenZip = $null
            foreach ($cand in @("C:\Program Files\7-Zip\7z.exe", "C:\Program Files (x86)\7-Zip\7z.exe")) {
                if (Test-Path -LiteralPath $cand) { $sevenZip = $cand; break }
            }

            if ($null -ne $sevenZip) {
                $argList = @("a", "-tzip", "-mx=1", "-mmt=on", "`"$ZipPath`"") + ($FilePaths | ForEach-Object { "`"$_`"" })
                $p = Start-Process -FilePath $sevenZip -ArgumentList $argList -NoNewWindow -Wait -PassThru
                if ($p.ExitCode -ne 0) {
                    throw "7-Zip exited with code $($p.ExitCode)"
                }
            }
            else {
                $archive = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
                try {
                    foreach ($src in $FilePaths) {
                        $entryName = [System.IO.Path]::GetFileName($src)
                        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                            $archive,
                            $src,
                            $entryName,
                            [System.IO.Compression.CompressionLevel]::Fastest
                        )
                    }
                }
                finally {
                    $archive.Dispose()
                }
            }

            if ($DeleteBakAfterZip) {
                foreach ($src in $FilePaths) {
                    Remove-Item -LiteralPath $src -Force -ErrorAction SilentlyContinue
                }
                if ($Subfolder -and (Test-Path -LiteralPath $Subfolder)) {
                    $rem = @(Get-ChildItem -LiteralPath $Subfolder -Force -ErrorAction SilentlyContinue)
                    if ($rem.Count -eq 0) {
                        Remove-Item -LiteralPath $Subfolder -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            $zSw.Stop()
            $zipElapsed = $zSw.Elapsed.TotalSeconds
            $finalOutput = $ZipPath
        }

        $sw.Stop()
        $EventQueue.Enqueue(@{
            Type = "FINISHED"
            DbName = $DbName
            Success = $true
            Elapsed = $sw.Elapsed.TotalSeconds
            ZipElapsed = $zipElapsed
            StreamCount = $StreamCount
            OutputPath = $finalOutput
        })
    }
    catch {
        $sw.Stop()
        $EventQueue.Enqueue(@{
            Type = "FINISHED"
            DbName = $DbName
            Success = $false
            Elapsed = $sw.Elapsed.TotalSeconds
            StreamCount = $StreamCount
            Error = ($_.Exception.GetBaseException().Message -replace "(\r|\n)+", " ")
        })
    }
    finally {
        if ($conn.State -ne 'Closed') { $conn.Close() }
        $conn.Dispose()
    }
}

$uiTimer = New-Object System.Windows.Forms.Timer
$uiTimer.Interval = 150
$uiTimer.Add_Tick({
    $evt = $null
    while ($Script:UiEventQueue.TryDequeue([ref]$evt)) {
        $dbName = $evt.DbName
        $lvItem = $null
        foreach ($it in $lvDbs.Items) {
            if ($it.Text -eq $dbName) { $lvItem = $it; break }
        }

        if ($evt.Type -eq "PROGRESS") {
            if ($evt.Percent -ge 0) {
                $Script:ProgressMap[$dbName] = $evt.Percent
            }
            if ($null -ne $lvItem) {
                $lvItem.SubItems[4].Text = $evt.Status
            }
            if ($evt.Percent -in @(0, 90)) {
                Write-UiLog "[$dbName] $($evt.Status)"
            }
        }
        elseif ($evt.Type -eq "LOG") {
            Write-UiLog "[$dbName] $($evt.Message)"
        }
        elseif ($evt.Type -eq "FINISHED") {
            $Script:FinishedBatchCount++
            $Script:ProgressMap[$dbName] = 100
            if ($evt.Success) {
                $zipNote = if ($evt.ZipElapsed -gt 0) { (", zipped in {0:N1}s" -f $evt.ZipElapsed) } else { "" }
                $stText = ("Done in {0:N1}s ({1} stream(s){2})" -f $evt.Elapsed, $evt.StreamCount, $zipNote)
                if ($null -ne $lvItem) { $lvItem.SubItems[4].Text = $stText }
                Write-UiLog ("SUCCESS [{0}] Completed in {1:N1}s ({2} stream(s){3}) -> {4}" -f $dbName, $evt.Elapsed, $evt.StreamCount, $zipNote, $evt.OutputPath)
            }
            else {
                $stText = "Failed: $($evt.Error)"
                if ($null -ne $lvItem) { $lvItem.SubItems[4].Text = $stText }
                Write-UiLog "FAILED [$dbName]: $($evt.Error)"
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
        $btnCancel.Enabled = $false
        if ($null -ne $Script:RunspacePool) {
            $Script:RunspacePool.Close()
            $Script:RunspacePool.Dispose()
            $Script:RunspacePool = $null
        }
        Write-UiLog "================================================================"
        Write-UiLog "Parallel backup & Fast-Zip batch finished ($($Script:FinishedBatchCount) database(s) processed)."
    }
})

$btnStart.Add_Click({
    if ($Script:IsRunning) { return }
    $selectedItems = @($lvDbs.CheckedItems)
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Please select at least one database to back up.", "No Databases Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $targetFolder = $txtTarget.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($targetFolder)) {
        [System.Windows.Forms.MessageBox]::Show($form, "Please specify a target backup folder.", "Target Folder Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $connStr = New-SqlConnectionString -Server $txtServer.Text -AuthMode ([string]$cmbAuth.SelectedItem) -Username $txtUser.Text -Password $txtPass.Text -TrustServerCertificate $chkTrustCert.Checked
    $maxParallel = [int]$numParallelDbs.Value
    $thresh = [double]$numThresholdGb.Value
    $maxStreams = [int]$numMaxStreams.Value
    $zipEnabled = [bool]$chkZipAfter.Checked
    $customZip = $txtZipFolder.Text.Trim()
    $delBak = [bool]$chkDelBakAfterZip.Checked
    $batchTs = (Get-Date).ToString("yyyyMMdd_HHmmss")

    $Script:IsRunning = $true
    $Script:TotalBatchCount = $selectedItems.Count
    $Script:FinishedBatchCount = 0
    $Script:ProgressMap = @{}
    $Script:ActiveJobs.Clear()
    $progressBar.Value = 0

    $btnStart.Enabled = $false
    $btnConnect.Enabled = $false
    $btnCancel.Enabled = $true

    Write-UiLog "Starting parallel COPY_ONLY backup batch for $($selectedItems.Count) database(s) (Max Parallel DBs: $maxParallel, Split Threshold: > $thresh GB, Fast-Zip: $zipEnabled)..."

    $Script:RunspacePool = [runspacefactory]::CreateRunspacePool(1, $maxParallel)
    $Script:RunspacePool.Open()
    $uiTimer.Start()

    foreach ($item in $selectedItems) {
        $db = $item.Tag
        $item.SubItems[4].Text = "Queued"
        $Script:ProgressMap[$db.Name] = 0

        $stCount = Get-BackupStreamCount -SizeGB $db.SizeGB -ThresholdGB $thresh -MaxStreams $maxStreams
        $pInfo = Get-BackupFilePaths -TargetFolder $targetFolder -DbName $db.Name -StreamCount $stCount -Timestamp $batchTs
        $zInfo = Get-ZipFilePath -TargetFolder $targetFolder -DbName $db.Name -Timestamp $batchTs -CustomZipFolder $customZip
        $createDirSql = Get-CreateSubdirTsql -SubfolderPath $pInfo.Subfolder
        $bakSql = Get-BackupDatabaseTsql -DbName $db.Name -FilePaths $pInfo.FilePaths -CopyOnly $true -Compression $chkCompression.Checked -Checksum $chkChecksum.Checked -StatsInterval 5
        $bakFallbackSql = Get-BackupDatabaseTsql -DbName $db.Name -FilePaths $pInfo.FilePaths -CopyOnly $true -Compression $false -Checksum $chkChecksum.Checked -StatsInterval 5
        $verSql = if ($chkVerify.Checked) { Get-VerifyBackupTsql -FilePaths $pInfo.FilePaths -Checksum $chkChecksum.Checked } else { "" }

        $ps = [powershell]::Create()
        $ps.RunspacePool = $Script:RunspacePool
        [void]$ps.AddScript($workerScriptBlock).AddArgument($connStr).AddArgument($db.Name).AddArgument($pInfo.Subfolder).AddArgument($pInfo.FilePaths).AddArgument($createDirSql).AddArgument($bakSql).AddArgument($bakFallbackSql).AddArgument($verSql).AddArgument([bool]$chkCompression.Checked).AddArgument($stCount).AddArgument($zipEnabled).AddArgument($zInfo.ZipPath).AddArgument($delBak).AddArgument($Script:UiEventQueue)
        $handle = $ps.BeginInvoke()
        [void]$Script:ActiveJobs.Add(@{ PowerShell = $ps; Handle = $handle })
    }
})

$btnCancel.Add_Click({
    if (-not $Script:IsRunning) { return }
    Write-UiLog "Cancelling active backup workers..."
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
        if ($null -ne $saved.include_system_dbs) { $chkIncludeSys.Checked = [bool]$saved.include_system_dbs }
        if ($saved.target_folder) { $txtTarget.Text = $saved.target_folder }
        if ($null -ne $saved.zip_after_backup) { $chkZipAfter.Checked = [bool]$saved.zip_after_backup }
        if ($saved.zip_output_folder) {
            $Script:UserEditedZipFolder = $true
            $txtZipFolder.Text = $saved.zip_output_folder
        }
        if ($null -ne $saved.delete_bak_after_zip) { $chkDelBakAfterZip.Checked = [bool]$saved.delete_bak_after_zip }
        if ($saved.max_parallel_dbs) { $numParallelDbs.Value = [decimal]$saved.max_parallel_dbs }
        if ($saved.split_threshold_gb) { $numThresholdGb.Value = [decimal]$saved.split_threshold_gb }
        if ($saved.max_streams_per_db) { $numMaxStreams.Value = [decimal]$saved.max_streams_per_db }
        if ($null -ne $saved.compression) { $chkCompression.Checked = [bool]$saved.compression }
        if ($null -ne $saved.checksum) { $chkChecksum.Checked = [bool]$saved.checksum }
        if ($null -ne $saved.verify_only_after) { $chkVerify.Checked = [bool]$saved.verify_only_after }
    }
    catch {}
}

Update-SummaryAndPreview
[void]$form.ShowDialog()
