<#
.SYNOPSIS
    AWS S3 CP Transfer Utility - Native Windows GUI (No Python Required)
.DESCRIPTION
    Self-contained Windows Forms GUI for copying files and folders to/from Amazon S3 (`aws s3 cp`).
    Designed for Windows Servers and workstations where Python is not installed.
    Supports both:
      1. Executing `aws s3 cp` via AWS CLI (if installed)
      2. Built-in .NET AWS Signature Version 4 (SigV4) S3 REST engine (zero external dependencies)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Ensure TLS 1.2+ is enabled on Windows Server .NET Framework defaults
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:AppTitle = "AWS S3 CP Transfer Utility (Server / Native Edition)"
$Script:LocalConfigFile = Join-Path $PSScriptRoot ".aws_cp_config.json"
$Script:EmptyPayloadHash = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
$Script:BucketRegionCache = @{}

# ============================================================================
# Helper Functions & AWS SigV4 Implementation (.NET Pure Stdlib)
# ============================================================================

function Find-AwsCli {
    $cmd = Get-Command "aws" -ErrorAction SilentlyContinue
    if ($null -ne $cmd) { return $cmd.Source }
    $candidates = @(
        "C:\Program Files\Amazon\AWSCLIV2\aws.exe",
        "C:\Program Files (x86)\Amazon\AWSCLIV2\aws.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Amazon\AWSCLIV2\aws.exe")
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }
    return $null
}

function Format-ByteSize([long]$Bytes) {
    if ($Bytes -lt 1024) { return "$Bytes B" }
    $units = @("KB", "MB", "GB", "TB")
    $val = [double]$Bytes
    foreach ($u in $units) {
        $val = $val / 1024.0
        if ($val -lt 1024.0 -or $u -eq "TB") {
            return ("{0:N1} {1}" -f $val, $u)
        }
    }
    return "$Bytes B"
}

function ConvertTo-AwsUriEncoded([string]$Text, [bool]$EncodeSlash = $true) {
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $sb = New-Object System.Text.StringBuilder
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    foreach ($b in $bytes) {
        $ch = [char]$b
        $isUnreserved = (
            ($b -ge 0x41 -and $b -le 0x5A) -or
            ($b -ge 0x61 -and $b -le 0x7A) -or
            ($b -ge 0x30 -and $b -le 0x39) -or
            $ch -eq '-' -or $ch -eq '_' -or $ch -eq '.' -or $ch -eq '~'
        )
        if ($isUnreserved -or ((-not $EncodeSlash) -and $ch -eq '/')) {
            [void]$sb.Append($ch)
        } else {
            [void]$sb.AppendFormat("%{0:X2}", $b)
        }
    }
    return $sb.ToString()
}

function Get-Sha256Hex([string]$Text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Invoke-HmacSha256([byte[]]$Key, [string]$Message) {
    $hmac = New-Object System.Security.Cryptography.HMACSHA256 -ArgumentList @(,$Key)
    try {
        return $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Message))
    } finally {
        $hmac.Dispose()
    }
}

function Get-AwsSigningKey([string]$SecretKey, [string]$DateStamp, [string]$Region, [string]$Service = "s3") {
    $kSecret = [System.Text.Encoding]::UTF8.GetBytes("AWS4" + $SecretKey)
    $kDate = Invoke-HmacSha256 -Key $kSecret -Message $DateStamp
    $kRegion = Invoke-HmacSha256 -Key $kDate -Message $Region
    $kService = Invoke-HmacSha256 -Key $kRegion -Message $Service
    return (Invoke-HmacSha256 -Key $kService -Message "aws4_request")
}

function Split-S3Uri([string]$Uri) {
    $clean = $Uri.Trim()
    if ($clean.StartsWith("s3://")) { $clean = $clean.Substring(5) }
    $clean = $clean.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($clean)) {
        throw "S3 URI must include a bucket name (e.g. s3://my-bucket/)."
    }
    $idx = $clean.IndexOf('/')
    if ($idx -lt 0) {
        return @{ Bucket = $clean; Key = "" }
    }
    return @{
        Bucket = $clean.Substring(0, $idx)
        Key    = $clean.Substring($idx + 1)
    }
}

function Invoke-S3SignedRequest {
    param(
        [string]$Method,
        [string]$AccessKey,
        [string]$SecretKey,
        [string]$SessionToken,
        [string]$Region,
        [string]$Bucket = "",
        [string]$Key = "",
        [hashtable]$QueryParams = @{},
        [string]$PayloadHash = $Script:EmptyPayloadHash,
        [string]$StorageClass = "",
        [string]$InFile = "",
        [string]$OutFile = ""
    )

    if ([string]::IsNullOrWhiteSpace($AccessKey) -or [string]::IsNullOrWhiteSpace($SecretKey)) {
        throw "AWS Access Key ID and Secret Access Key are required."
    }

    $effRegion = $Region.Trim()
    if ([string]::IsNullOrWhiteSpace($effRegion)) { $effRegion = "us-east-1" }
    if ($Bucket -and $Script:BucketRegionCache.ContainsKey($Bucket)) {
        $effRegion = $Script:BucketRegionCache[$Bucket]
    }

    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        $cleanKey = $Key.TrimStart('/')
        if ([string]::IsNullOrWhiteSpace($Bucket)) {
            $hostName = "s3.$effRegion.amazonaws.com"
            $canonicalUri = "/"
        } elseif ($Bucket.Contains(".")) {
            $hostName = "s3.$effRegion.amazonaws.com"
            $encKey = ConvertTo-AwsUriEncoded -Text $cleanKey -EncodeSlash $false
            $canonicalUri = if ($encKey) { "/$Bucket/$encKey" } else { "/$Bucket" }
        } else {
            $hostName = "$Bucket.s3.$effRegion.amazonaws.com"
            $encKey = ConvertTo-AwsUriEncoded -Text $cleanKey -EncodeSlash $false
            $canonicalUri = "/$encKey"
        }

        $utcNow = [DateTime]::UtcNow
        $amzDate = $utcNow.ToString("yyyyMMddTHHmmssZ")
        $dateStamp = $utcNow.ToString("yyyyMMdd")

        $sortedQueryKeys = @($QueryParams.Keys | Sort-Object)
        $queryParts = New-Object System.Collections.Generic.List[string]
        foreach ($qk in $sortedQueryKeys) {
            $qv = [string]$QueryParams[$qk]
            $queryParts.Add((ConvertTo-AwsUriEncoded $qk $true) + "=" + (ConvertTo-AwsUriEncoded $qv $true))
        }
        $canonicalQuery = [string]::Join("&", $queryParts)

        $headersToSign = @{
            "host"                 = $hostName
            "x-amz-content-sha256" = $PayloadHash
            "x-amz-date"           = $amzDate
        }
        if (-not [string]::IsNullOrWhiteSpace($SessionToken)) {
            $headersToSign["x-amz-security-token"] = $SessionToken.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($StorageClass) -and $StorageClass -ne "(Default)") {
            $headersToSign["x-amz-storage-class"] = $StorageClass.Trim()
        }

        $sortedHeaderKeys = @($headersToSign.Keys | Sort-Object)
        $canonicalHeadersSb = New-Object System.Text.StringBuilder
        foreach ($hk in $sortedHeaderKeys) {
            [void]$canonicalHeadersSb.AppendFormat("{0}:{1}`n", $hk, $headersToSign[$hk])
        }
        $signedHeaders = [string]::Join(";", $sortedHeaderKeys)

        $canonicalRequest = @(
            $Method.ToUpperInvariant(),
            $canonicalUri,
            $canonicalQuery,
            $canonicalHeadersSb.ToString(),
            $signedHeaders,
            $PayloadHash
        ) -join "`n"

        $credScope = "$dateStamp/$effRegion/s3/aws4_request"
        $stringToSign = @(
            "AWS4-HMAC-SHA256",
            $amzDate,
            $credScope,
            (Get-Sha256Hex $canonicalRequest)
        ) -join "`n"

        $signingKey = Get-AwsSigningKey -SecretKey $SecretKey.Trim() -DateStamp $dateStamp -Region $effRegion -Service "s3"
        $sigBytes = Invoke-HmacSha256 -Key $signingKey -Message $stringToSign
        $signature = ([System.BitConverter]::ToString($sigBytes)).Replace("-", "").ToLowerInvariant()

        $authHeader = "AWS4-HMAC-SHA256 Credential=$($AccessKey.Trim())/$credScope, SignedHeaders=$signedHeaders, Signature=$signature"

        $url = "https://$hostName$canonicalUri"
        if ($canonicalQuery) { $url = "$url?$canonicalQuery" }

        $req = [System.Net.HttpWebRequest]::Create($url)
        $req.Method = $Method.ToUpperInvariant()
        $req.Host = $hostName
        $req.Headers.Add("x-amz-content-sha256", $PayloadHash)
        $req.Headers.Add("x-amz-date", $amzDate)
        if (-not [string]::IsNullOrWhiteSpace($SessionToken)) {
            $req.Headers.Add("x-amz-security-token", $SessionToken.Trim())
        }
        if (-not [string]::IsNullOrWhiteSpace($StorageClass) -and $StorageClass -ne "(Default)") {
            $req.Headers.Add("x-amz-storage-class", $StorageClass.Trim())
        }
        $req.Headers.Add("Authorization", $authHeader)

        if ($InFile) {
            $fileInfo = Get-Item -LiteralPath $InFile
            $req.ContentLength = $fileInfo.Length
            $req.AllowWriteStreamBuffering = $false
            if ($fileInfo.Length -gt 0) {
                $inStream = [System.IO.File]::OpenRead($InFile)
                $reqStream = $req.GetRequestStream()
                try {
                    $inStream.CopyTo($reqStream, 262144)
                } finally {
                    $reqStream.Close()
                    $inStream.Close()
                }
            }
        } elseif ($Method.ToUpperInvariant() -eq "PUT") {
            $req.ContentLength = 0
        }

        try {
            $resp = [System.Net.HttpWebResponse]$req.GetResponse()
            try {
                if ($OutFile) {
                    $outDir = Split-Path -Parent $OutFile
                    if ($outDir -and (-not (Test-Path -LiteralPath $outDir))) {
                        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
                    }
                    $respStream = $resp.GetResponseStream()
                    $fs = [System.IO.File]::Create($OutFile)
                    try {
                        $respStream.CopyTo($fs, 262144)
                    } finally {
                        $fs.Close()
                        $respStream.Close()
                    }
                    return ""
                } else {
                    $reader = New-Object System.IO.StreamReader($resp.GetResponseStream(), [System.Text.Encoding]::UTF8)
                    try {
                        return $reader.ReadToEnd()
                    } finally {
                        $reader.Close()
                    }
                }
            } finally {
                $resp.Close()
            }
        } catch [System.Net.WebException] {
            $webEx = $_.Exception
            $errResp = [System.Net.HttpWebResponse]$webEx.Response
            if ($null -ne $errResp) {
                $discRegion = $errResp.Headers["x-amz-bucket-region"]
                $errBody = ""
                try {
                    $sr = New-Object System.IO.StreamReader($errResp.GetResponseStream())
                    $errBody = $sr.ReadToEnd()
                    $sr.Close()
                } catch {}
                $errResp.Close()

                if (-not $discRegion -and $errBody -match "<Region>([^<]+)</Region>") {
                    $discRegion = $Matches[1].Trim()
                }
                if ($attempt -eq 0 -and $Bucket -and $discRegion -and ($discRegion -ne $effRegion)) {
                    $Script:BucketRegionCache[$Bucket] = $discRegion
                    $effRegion = $discRegion
                    continue
                }
                if ($errBody -match "<Code>([^<]+)</Code>" -and $errBody -match "<Message>([^<]+)</Message>") {
                    throw "S3 Error ($($Matches[1])): $($Matches[2])"
                }
            }
            throw
        }
    }
}

function Build-AwsCpCommandPreview {
    param(
        [string]$Source,
        [string]$Target,
        [string]$Region,
        [string]$OutputFormat,
        [bool]$Recursive,
        [bool]$DryRun,
        [string]$StorageClass
    )
    $src = if ($Source) { $Source } else { "<source>" }
    $dst = if ($Target) { $Target } else { "<target>" }
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("aws s3 cp")
    $parts.Add(('"{0}"' -f $src))
    $parts.Add(('"{0}"' -f $dst))
    if ($Recursive) { $parts.Add("--recursive") }
    if ($DryRun) { $parts.Add("--dryrun") }
    if ($StorageClass -and $StorageClass -ne "(Default)") {
        $parts.Add("--storage-class $StorageClass")
    }
    if ($Region) { $parts.Add("--region $($Region.Trim())") }
    if ($OutputFormat) { $parts.Add("--output $($OutputFormat.Trim())") }
    return ($parts -join " ")
}

# ============================================================================
# Build Windows Forms GUI
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = $Script:AppTitle
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.MinimumSize = New-Object System.Drawing.Size(960, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- 1. AWS Configuration GroupBox ---
$grpCreds = New-Object System.Windows.Forms.GroupBox
$grpCreds.Text = "1. AWS Configuration Attributes (Access Key ID, Secret Key, Region, Default Format)"
$grpCreds.Location = New-Object System.Drawing.Point(12, 10)
$grpCreds.Size = New-Object System.Drawing.Size(1140, 95)
$grpCreds.Anchor = "Top, Left, Right"
$form.Controls.Add($grpCreds)

$lblAk = New-Object System.Windows.Forms.Label
$lblAk.Text = "AWS Access Key ID:"
$lblAk.Location = New-Object System.Drawing.Point(12, 26)
$lblAk.AutoSize = $true
$grpCreds.Controls.Add($lblAk)

$txtAk = New-Object System.Windows.Forms.TextBox
$txtAk.Location = New-Object System.Drawing.Point(135, 23)
$txtAk.Size = New-Object System.Drawing.Size(230, 23)
$grpCreds.Controls.Add($txtAk)

$lblSk = New-Object System.Windows.Forms.Label
$lblSk.Text = "AWS Secret Key:"
$lblSk.Location = New-Object System.Drawing.Point(380, 26)
$lblSk.AutoSize = $true
$grpCreds.Controls.Add($lblSk)

$txtSk = New-Object System.Windows.Forms.TextBox
$txtSk.Location = New-Object System.Drawing.Point(485, 23)
$txtSk.Size = New-Object System.Drawing.Size(250, 23)
$txtSk.UseSystemPasswordChar = $true
$grpCreds.Controls.Add($txtSk)

$chkShowSk = New-Object System.Windows.Forms.CheckBox
$chkShowSk.Text = "Show"
$chkShowSk.Location = New-Object System.Drawing.Point(742, 24)
$chkShowSk.AutoSize = $true
$chkShowSk.Add_CheckedChanged({
    $txtSk.UseSystemPasswordChar = (-not $chkShowSk.Checked)
})
$grpCreds.Controls.Add($chkShowSk)

$lblTok = New-Object System.Windows.Forms.Label
$lblTok.Text = "Token (Opt):"
$lblTok.Location = New-Object System.Drawing.Point(810, 26)
$lblTok.AutoSize = $true
$grpCreds.Controls.Add($lblTok)

$txtTok = New-Object System.Windows.Forms.TextBox
$txtTok.Location = New-Object System.Drawing.Point(890, 23)
$txtTok.Size = New-Object System.Drawing.Size(235, 23)
$txtTok.UseSystemPasswordChar = $true
$grpCreds.Controls.Add($txtTok)

$lblReg = New-Object System.Windows.Forms.Label
$lblReg.Text = "Default Region:"
$lblReg.Location = New-Object System.Drawing.Point(12, 58)
$lblReg.AutoSize = $true
$grpCreds.Controls.Add($lblReg)

$cmbRegion = New-Object System.Windows.Forms.ComboBox
$cmbRegion.Location = New-Object System.Drawing.Point(135, 55)
$cmbRegion.Size = New-Object System.Drawing.Size(160, 23)
[void]$cmbRegion.Items.AddRange(@(
    "us-east-1","us-east-2","us-west-1","us-west-2",
    "eu-west-1","eu-west-2","eu-west-3","eu-central-1","eu-north-1",
    "ap-south-1","ap-southeast-1","ap-southeast-2","ap-northeast-1","ap-northeast-2",
    "ca-central-1","sa-east-1","me-south-1","af-south-1"
))
$cmbRegion.Text = "us-east-1"
$grpCreds.Controls.Add($cmbRegion)

$lblFmt = New-Object System.Windows.Forms.Label
$lblFmt.Text = "Default Format:"
$lblFmt.Location = New-Object System.Drawing.Point(380, 58)
$lblFmt.AutoSize = $true
$grpCreds.Controls.Add($lblFmt)

$cmbFormat = New-Object System.Windows.Forms.ComboBox
$cmbFormat.Location = New-Object System.Drawing.Point(485, 55)
$cmbFormat.Size = New-Object System.Drawing.Size(120, 23)
$cmbFormat.DropDownStyle = "DropDownList"
[void]$cmbFormat.Items.AddRange(@("json", "text", "table", "yaml", "yaml-stream"))
$cmbFormat.SelectedIndex = 0
$grpCreds.Controls.Add($cmbFormat)

$btnSaveLocal = New-Object System.Windows.Forms.Button
$btnSaveLocal.Text = "Save Local Config"
$btnSaveLocal.Location = New-Object System.Drawing.Point(780, 54)
$btnSaveLocal.Size = New-Object System.Drawing.Size(140, 26)
$grpCreds.Controls.Add($btnSaveLocal)

$btnListBuckets = New-Object System.Windows.Forms.Button
$btnListBuckets.Text = "Connect & List Buckets"
$btnListBuckets.Location = New-Object System.Drawing.Point(930, 54)
$btnListBuckets.Size = New-Object System.Drawing.Size(195, 26)
$grpCreds.Controls.Add($btnListBuckets)

# --- 2. Transfer Direction GroupBox ---
$grpDir = New-Object System.Windows.Forms.GroupBox
$grpDir.Text = "2. Transfer Direction"
$grpDir.Location = New-Object System.Drawing.Point(12, 110)
$grpDir.Size = New-Object System.Drawing.Size(1140, 52)
$grpDir.Anchor = "Top, Left, Right"
$form.Controls.Add($grpDir)

$rbUpload = New-Object System.Windows.Forms.RadioButton
$rbUpload.Text = "Copy TO S3 Bucket (Upload: Local File/Folder -> S3 Target Folder)"
$rbUpload.Location = New-Object System.Drawing.Point(16, 20)
$rbUpload.AutoSize = $true
$rbUpload.Checked = $true
$grpDir.Controls.Add($rbUpload)

$rbDownload = New-Object System.Windows.Forms.RadioButton
$rbDownload.Text = "Copy FROM S3 Bucket (Download: S3 Object/Folder -> Local Target Folder)"
$rbDownload.Location = New-Object System.Drawing.Point(480, 20)
$rbDownload.AutoSize = $true
$grpDir.Controls.Add($rbDownload)

# --- 3. Dual Explorers (Local & S3) ---
$split = New-Object System.Windows.Forms.SplitContainer
$split.Location = New-Object System.Drawing.Point(12, 168)
$split.Size = New-Object System.Drawing.Size(1140, 320)
$split.SplitterDistance = 560
$split.Anchor = "Top, Bottom, Left, Right"
$form.Controls.Add($split)

# Left Panel: Local Explorer
$grpLocal = New-Object System.Windows.Forms.GroupBox
$grpLocal.Text = "Local File & Folder Explorer"
$grpLocal.Dock = "Fill"
$split.Panel1.Controls.Add($grpLocal)

$txtLocalDir = New-Object System.Windows.Forms.TextBox
$txtLocalDir.Location = New-Object System.Drawing.Point(10, 22)
$txtLocalDir.Size = New-Object System.Drawing.Size(360, 23)
$txtLocalDir.Text = (Get-Location).Path
$grpLocal.Controls.Add($txtLocalDir)

$btnLocalGo = New-Object System.Windows.Forms.Button
$btnLocalGo.Text = "Go"
$btnLocalGo.Location = New-Object System.Drawing.Point(376, 21)
$btnLocalGo.Size = New-Object System.Drawing.Size(45, 25)
$grpLocal.Controls.Add($btnLocalGo)

$btnLocalUp = New-Object System.Windows.Forms.Button
$btnLocalUp.Text = "Up"
$btnLocalUp.Location = New-Object System.Drawing.Point(425, 21)
$btnLocalUp.Size = New-Object System.Drawing.Size(45, 25)
$grpLocal.Controls.Add($btnLocalUp)

$btnPickFile = New-Object System.Windows.Forms.Button
$btnPickFile.Text = "Pick File Dialog..."
$btnPickFile.Location = New-Object System.Drawing.Point(10, 52)
$btnPickFile.Size = New-Object System.Drawing.Size(150, 25)
$grpLocal.Controls.Add($btnPickFile)

$btnPickFolder = New-Object System.Windows.Forms.Button
$btnPickFolder.Text = "Pick Folder Dialog..."
$btnPickFolder.Location = New-Object System.Drawing.Point(166, 52)
$btnPickFolder.Size = New-Object System.Drawing.Size(150, 25)
$grpLocal.Controls.Add($btnPickFolder)

$lvLocal = New-Object System.Windows.Forms.ListView
$lvLocal.Location = New-Object System.Drawing.Point(10, 84)
$lvLocal.Size = New-Object System.Drawing.Size(535, 225)
$lvLocal.Anchor = "Top, Bottom, Left, Right"
$lvLocal.View = "Details"
$lvLocal.FullRowSelect = $true
[void]$lvLocal.Columns.Add("Name", 240)
[void]$lvLocal.Columns.Add("Type", 70)
[void]$lvLocal.Columns.Add("Size", 85)
[void]$lvLocal.Columns.Add("Modified", 125)
$grpLocal.Controls.Add($lvLocal)

# Right Panel: S3 Explorer
$grpS3 = New-Object System.Windows.Forms.GroupBox
$grpS3.Text = "S3 Bucket & Folder Explorer"
$grpS3.Dock = "Fill"
$split.Panel2.Controls.Add($grpS3)

$lblBucket = New-Object System.Windows.Forms.Label
$lblBucket.Text = "Bucket:"
$lblBucket.Location = New-Object System.Drawing.Point(10, 25)
$lblBucket.AutoSize = $true
$grpS3.Controls.Add($lblBucket)

$cmbBucket = New-Object System.Windows.Forms.ComboBox
$cmbBucket.Location = New-Object System.Drawing.Point(60, 22)
$cmbBucket.Size = New-Object System.Drawing.Size(170, 23)
$grpS3.Controls.Add($cmbBucket)

$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "Prefix:"
$lblPrefix.Location = New-Object System.Drawing.Point(238, 25)
$lblPrefix.AutoSize = $true
$grpS3.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Location = New-Object System.Drawing.Point(282, 22)
$txtPrefix.Size = New-Object System.Drawing.Size(160, 23)
$grpS3.Controls.Add($txtPrefix)

$btnS3Browse = New-Object System.Windows.Forms.Button
$btnS3Browse.Text = "Browse"
$btnS3Browse.Location = New-Object System.Drawing.Point(448, 21)
$btnS3Browse.Size = New-Object System.Drawing.Size(65, 25)
$grpS3.Controls.Add($btnS3Browse)

$btnS3Up = New-Object System.Windows.Forms.Button
$btnS3Up.Text = "Up"
$btnS3Up.Location = New-Object System.Drawing.Point(516, 21)
$btnS3Up.Size = New-Object System.Drawing.Size(45, 25)
$grpS3.Controls.Add($btnS3Up)

$btnUseCurrentS3 = New-Object System.Windows.Forms.Button
$btnUseCurrentS3.Text = "Use Current S3 Folder"
$btnUseCurrentS3.Location = New-Object System.Drawing.Point(10, 52)
$btnUseCurrentS3.Size = New-Object System.Drawing.Size(170, 25)
$grpS3.Controls.Add($btnUseCurrentS3)

$lvS3 = New-Object System.Windows.Forms.ListView
$lvS3.Location = New-Object System.Drawing.Point(10, 84)
$lvS3.Size = New-Object System.Drawing.Size(550, 225)
$lvS3.Anchor = "Top, Bottom, Left, Right"
$lvS3.View = "Details"
$lvS3.FullRowSelect = $true
[void]$lvS3.Columns.Add("S3 Object / Prefix", 250)
[void]$lvS3.Columns.Add("Type", 70)
[void]$lvS3.Columns.Add("Size", 85)
[void]$lvS3.Columns.Add("Last Modified", 130)
$grpS3.Controls.Add($lvS3)

# --- 4. Transfer Paths, Preview & Action Controls ---
$grpTransfer = New-Object System.Windows.Forms.GroupBox
$grpTransfer.Text = "3. Selected Paths, Options & AWS CLI Command Preview"
$grpTransfer.Location = New-Object System.Drawing.Point(12, 495)
$grpTransfer.Size = New-Object System.Drawing.Size(1140, 155)
$grpTransfer.Anchor = "Bottom, Left, Right"
$form.Controls.Add($grpTransfer)

$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Source Path:"
$lblSrc.Location = New-Object System.Drawing.Point(12, 26)
$lblSrc.AutoSize = $true
$grpTransfer.Controls.Add($lblSrc)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(120, 23)
$txtSource.Size = New-Object System.Drawing.Size(1005, 23)
$txtSource.Anchor = "Top, Left, Right"
$grpTransfer.Controls.Add($txtSource)

$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = "Target Path:"
$lblDst.Location = New-Object System.Drawing.Point(12, 56)
$lblDst.AutoSize = $true
$grpTransfer.Controls.Add($lblDst)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(120, 53)
$txtTarget.Size = New-Object System.Drawing.Size(1005, 23)
$txtTarget.Anchor = "Top, Left, Right"
$grpTransfer.Controls.Add($txtTarget)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Recursive (--recursive)"
$chkRecursive.Location = New-Object System.Drawing.Point(120, 84)
$chkRecursive.AutoSize = $true
$grpTransfer.Controls.Add($chkRecursive)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run (--dryrun)"
$chkDryRun.Location = New-Object System.Drawing.Point(300, 84)
$chkDryRun.AutoSize = $true
$grpTransfer.Controls.Add($chkDryRun)

$lblCmd = New-Object System.Windows.Forms.Label
$lblCmd.Text = "CLI Preview:"
$lblCmd.Location = New-Object System.Drawing.Point(12, 118)
$lblCmd.AutoSize = $true
$grpTransfer.Controls.Add($lblCmd)

$txtCmdPreview = New-Object System.Windows.Forms.TextBox
$txtCmdPreview.Location = New-Object System.Drawing.Point(120, 115)
$txtCmdPreview.Size = New-Object System.Drawing.Size(810, 23)
$txtCmdPreview.ReadOnly = $true
$txtCmdPreview.Anchor = "Top, Left, Right"
$grpTransfer.Controls.Add($txtCmdPreview)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start S3 CP Transfer"
$btnStart.Location = New-Object System.Drawing.Point(940, 112)
$btnStart.Size = New-Object System.Drawing.Size(185, 28)
$btnStart.Anchor = "Top, Right"
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTransfer.Controls.Add($btnStart)

# --- 5. Log Output ---
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(12, 658)
$txtLog.Size = New-Object System.Drawing.Size(1140, 110)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.Anchor = "Bottom, Left, Right"
$form.Controls.Add($txtLog)

# ============================================================================
# UI Logic & Event Wiring
# ============================================================================

function Write-UiLog([string]$Message) {
    $ts = (Get-Date).ToString("HH:mm:ss")
    $txtLog.AppendText("[$ts] $Message`r`n")
}

function Update-Preview {
    $txtCmdPreview.Text = Build-AwsCpCommandPreview `
        -Source $txtSource.Text `
        -Target $txtTarget.Text `
        -Region $cmbRegion.Text `
        -OutputFormat $cmbFormat.Text `
        -Recursive $chkRecursive.Checked `
        -DryRun $chkDryRun.Checked `
        -StorageClass "(Default)"
}

function Refresh-LocalExplorer {
    $dir = $txtLocalDir.Text.Trim()
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("Directory not found: $dir", "Invalid Directory") | Out-Null
        return
    }
    $resolved = (Resolve-Path -LiteralPath $dir).Path
    $txtLocalDir.Text = $resolved
    $lvLocal.Items.Clear()

    $items = Get-ChildItem -LiteralPath $resolved -Force -ErrorAction SilentlyContinue |
        Sort-Object @{Expression = { -not $_.PSIsContainer }}, Name
    foreach ($it in $items) {
        $lvi = New-Object System.Windows.Forms.ListViewItem($it.Name)
        $lvi.Tag = $it.FullName
        if ($it.PSIsContainer) {
            [void]$lvi.SubItems.Add("Folder")
            [void]$lvi.SubItems.Add("")
        } else {
            [void]$lvi.SubItems.Add("File")
            [void]$lvi.SubItems.Add((Format-ByteSize $it.Length))
        }
        [void]$lvi.SubItems.Add($it.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
        [void]$lvLocal.Items.Add($lvi)
    }
}

$btnLocalGo.Add_Click({ Refresh-LocalExplorer })
$btnLocalUp.Add_Click({
    $parent = Split-Path -Parent $txtLocalDir.Text
    if ($parent) {
        $txtLocalDir.Text = $parent
        Refresh-LocalExplorer
    }
})

$lvLocal.Add_SelectedIndexChanged({
    if ($lvLocal.SelectedItems.Count -eq 0) { return }
    $selItem = $lvLocal.SelectedItems[0]
    $path = [string]$selItem.Tag
    $isDir = ($selItem.SubItems[1].Text -eq "Folder")
    if ($rbUpload.Checked) {
        $txtSource.Text = $path
        $chkRecursive.Checked = $isDir
    } else {
        $txtTarget.Text = if ($isDir) { $path } else { (Split-Path -Parent $path) }
    }
    Update-Preview
})

$lvLocal.Add_DoubleClick({
    if ($lvLocal.SelectedItems.Count -eq 0) { return }
    $selItem = $lvLocal.SelectedItems[0]
    if ($selItem.SubItems[1].Text -eq "Folder") {
        $txtLocalDir.Text = [string]$selItem.Tag
        Refresh-LocalExplorer
    }
})

$btnPickFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.InitialDirectory = $txtLocalDir.Text
    if ($dlg.ShowDialog() -eq "OK") {
        if ($rbUpload.Checked) {
            $txtSource.Text = $dlg.FileName
            $chkRecursive.Checked = $false
        } else {
            $txtTarget.Text = $dlg.FileName
        }
        Update-Preview
    }
})

$btnPickFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtLocalDir.Text
    if ($dlg.ShowDialog() -eq "OK") {
        $txtLocalDir.Text = $dlg.SelectedPath
        Refresh-LocalExplorer
        if ($rbUpload.Checked) {
            $txtSource.Text = $dlg.SelectedPath
            $chkRecursive.Checked = $true
        } else {
            $txtTarget.Text = $dlg.SelectedPath
        }
        Update-Preview
    }
})

$btnListBuckets.Add_Click({
    try {
        Write-UiLog "Connecting to S3 to list buckets..."
        $xmlText = Invoke-S3SignedRequest `
            -Method "GET" `
            -AccessKey $txtAk.Text `
            -SecretKey $txtSk.Text `
            -SessionToken $txtTok.Text `
            -Region $cmbRegion.Text
        [xml]$doc = $xmlText
        $cmbBucket.Items.Clear()
        $buckets = @($doc.ListAllMyBucketsResult.Buckets.Bucket)
        foreach ($b in $buckets) {
            if ($null -ne $b -and $b.Name) {
                [void]$cmbBucket.Items.Add($b.Name)
            }
        }
        if ($cmbBucket.Items.Count -gt 0 -and [string]::IsNullOrWhiteSpace($cmbBucket.Text)) {
            $cmbBucket.SelectedIndex = 0
        }
        Write-UiLog "Listed $($cmbBucket.Items.Count) bucket(s)."
    } catch {
        Write-UiLog "ListBuckets error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "S3 Error") | Out-Null
    }
})

function Refresh-S3Objects {
    $bucket = $cmbBucket.Text.Trim()
    if (-not $bucket) {
        [System.Windows.Forms.MessageBox]::Show("Please enter or select an S3 bucket name.", "Bucket Required") | Out-Null
        return
    }
    $prefix = $txtPrefix.Text.Trim().TrimStart('/')
    if ($prefix -and (-not $prefix.EndsWith("/"))) { $prefix = "$prefix/" }
    $txtPrefix.Text = $prefix

    try {
        $q = @{ "list-type" = "2"; "delimiter" = "/"; "max-keys" = "1000" }
        if ($prefix) { $q["prefix"] = $prefix }
        $xmlText = Invoke-S3SignedRequest `
            -Method "GET" `
            -AccessKey $txtAk.Text `
            -SecretKey $txtSk.Text `
            -SessionToken $txtTok.Text `
            -Region $cmbRegion.Text `
            -Bucket $bucket `
            -QueryParams $q
        [xml]$doc = $xmlText
        $lvS3.Items.Clear()

        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace("s3", $doc.DocumentElement.NamespaceURI)

        foreach ($cp in $doc.SelectNodes("//s3:CommonPrefixes/s3:Prefix", $ns)) {
            $pVal = $cp.InnerText
            $rel = if ($pVal.StartsWith($prefix)) { $pVal.Substring($prefix.Length).TrimEnd('/') } else { $pVal }
            $lvi = New-Object System.Windows.Forms.ListViewItem("[DIR] $rel/")
            $lvi.Tag = $pVal
            [void]$lvi.SubItems.Add("Folder")
            [void]$lvi.SubItems.Add("")
            [void]$lvi.SubItems.Add("")
            [void]$lvS3.Items.Add($lvi)
        }
        foreach ($c in $doc.SelectNodes("//s3:Contents", $ns)) {
            $kVal = $c.Key
            if ($kVal -eq $prefix) { continue }
            $rel = if ($kVal.StartsWith($prefix)) { $kVal.Substring($prefix.Length) } else { $kVal }
            $lvi = New-Object System.Windows.Forms.ListViewItem($rel)
            $lvi.Tag = $kVal
            [void]$lvi.SubItems.Add("Object")
            [void]$lvi.SubItems.Add((Format-ByteSize ([long]$c.Size)))
            [void]$lvi.SubItems.Add($c.LastModified.Substring(0, [Math]::Min(19, $c.LastModified.Length)).Replace("T", " "))
            [void]$lvS3.Items.Add($lvi)
        }
        Write-UiLog "Loaded $($lvS3.Items.Count) item(s) in s3://$bucket/$prefix"
    } catch {
        Write-UiLog "Error listing s3://$bucket/$prefix : $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "S3 Browse Error") | Out-Null
    }
}

$btnS3Browse.Add_Click({ Refresh-S3Objects })
$btnS3Up.Add_Click({
    $p = $txtPrefix.Text.Trim().Trim('/')
    if (-not $p) { return }
    $idx = $p.LastIndexOf('/')
    $txtPrefix.Text = if ($idx -ge 0) { $p.Substring(0, $idx + 1) } else { "" }
    Refresh-S3Objects
})

$btnUseCurrentS3.Add_Click({
    $b = $cmbBucket.Text.Trim()
    if (-not $b) { return }
    $p = $txtPrefix.Text.Trim().TrimStart('/')
    $uri = if ($p) { "s3://$b/$p" } else { "s3://$b/" }
    if ($rbUpload.Checked) {
        $txtTarget.Text = $uri
    } else {
        $txtSource.Text = $uri
        $chkRecursive.Checked = $true
    }
    Update-Preview
})

$lvS3.Add_SelectedIndexChanged({
    if ($lvS3.SelectedItems.Count -eq 0) { return }
    $b = $cmbBucket.Text.Trim()
    $sel = $lvS3.SelectedItems[0]
    $key = [string]$sel.Tag
    $isDir = ($sel.SubItems[1].Text -eq "Folder")
    if ($rbUpload.Checked) {
        if ($isDir) {
            $txtTarget.Text = "s3://$b/$key"
        } else {
            $txtTarget.Text = "s3://$b/$($txtPrefix.Text)"
        }
    } else {
        $txtSource.Text = "s3://$b/$key"
        $chkRecursive.Checked = $isDir
    }
    Update-Preview
})

$lvS3.Add_DoubleClick({
    if ($lvS3.SelectedItems.Count -eq 0) { return }
    $sel = $lvS3.SelectedItems[0]
    if ($sel.SubItems[1].Text -eq "Folder") {
        $txtPrefix.Text = [string]$sel.Tag
        Refresh-S3Objects
    }
})

$btnSaveLocal.Add_Click({
    $cfg = @{
        aws_access_key_id     = $txtAk.Text.Trim()
        aws_secret_access_key = $txtSk.Text.Trim()
        region                = $cmbRegion.Text.Trim()
        output_format         = $cmbFormat.Text.Trim()
        session_token         = $txtTok.Text.Trim()
        last_bucket           = $cmbBucket.Text.Trim()
    }
    $cfg | ConvertTo-Json | Set-Content -LiteralPath $Script:LocalConfigFile -Encoding UTF8
    Write-UiLog "Saved local configuration to $Script:LocalConfigFile"
})

foreach ($ctrl in @($txtSource, $txtTarget, $cmbRegion, $cmbFormat)) {
    $ctrl.Add_TextChanged({ Update-Preview })
}
$chkRecursive.Add_CheckedChanged({ Update-Preview })
$chkDryRun.Add_CheckedChanged({ Update-Preview })

$btnStart.Add_Click({
    $src = $txtSource.Text.Trim()
    $dst = $txtTarget.Text.Trim()
    if (-not $src -or -not $dst) {
        [System.Windows.Forms.MessageBox]::Show("Please specify both Source and Target paths.", "Missing Paths") | Out-Null
        return
    }
    $awsBin = Find-AwsCli
    if ($null -ne $awsBin) {
        Write-UiLog "Executing via AWS CLI ($awsBin)..."
        $env:AWS_ACCESS_KEY_ID = $txtAk.Text.Trim()
        $env:AWS_SECRET_ACCESS_KEY = $txtSk.Text.Trim()
        $env:AWS_DEFAULT_REGION = $cmbRegion.Text.Trim()
        $env:AWS_DEFAULT_OUTPUT = $cmbFormat.Text.Trim()
        if ($txtTok.Text.Trim()) { $env:AWS_SESSION_TOKEN = $txtTok.Text.Trim() }
        $cliArgs = @("s3", "cp", $src, $dst, "--region", $cmbRegion.Text.Trim(), "--output", $cmbFormat.Text.Trim())
        if ($chkRecursive.Checked) { $cliArgs += "--recursive" }
        if ($chkDryRun.Checked) { $cliArgs += "--dryrun" }
        try {
            $output = & $awsBin @cliArgs 2>&1
            foreach ($line in $output) { Write-UiLog ([string]$line) }
            Write-UiLog "Transfer completed."
        } catch {
            Write-UiLog "AWS CLI error: $($_.Exception.Message)"
        }
    } else {
        try {
            if ($rbUpload.Checked) {
                $parsed = Split-S3Uri $dst
                if (Test-Path -LiteralPath $src -PathType Container) {
                    if (-not $chkRecursive.Checked) {
                        throw "Source is a folder. Check 'Recursive (--recursive)' to upload all files inside it."
                    }
                    $basePath = (Resolve-Path -LiteralPath $src).Path.TrimEnd('\')
                    $basePrefix = $parsed.Key.Trim('/')
                    if ($basePrefix) { $basePrefix = "$basePrefix/" }
                    $files = @(Get-ChildItem -LiteralPath $basePath -Recurse -File)
                    Write-UiLog "Uploading $($files.Count) file(s) from folder $basePath to s3://$($parsed.Bucket)/$basePrefix ..."
                    foreach ($f in $files) {
                        $rel = $f.FullName.Substring($basePath.Length).TrimStart('\').Replace('\', '/')
                        $destKey = "$basePrefix$rel"
                        Write-UiLog "upload: $($f.FullName) -> s3://$($parsed.Bucket)/$destKey"
                        if (-not $chkDryRun.Checked) {
                            [void](Invoke-S3SignedRequest `
                                -Method "PUT" `
                                -AccessKey $txtAk.Text `
                                -SecretKey $txtSk.Text `
                                -SessionToken $txtTok.Text `
                                -Region $cmbRegion.Text `
                                -Bucket $parsed.Bucket `
                                -Key $destKey `
                                -PayloadHash "UNSIGNED-PAYLOAD" `
                                -InFile $f.FullName)
                        }
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                } else {
                    $key = $parsed.Key
                    if (-not $key -or $key.EndsWith("/")) {
                        $key = "$key$(Split-Path -Leaf $src)"
                    }
                    Write-UiLog "Uploading $src -> s3://$($parsed.Bucket)/$key ..."
                    if (-not $chkDryRun.Checked) {
                        [void](Invoke-S3SignedRequest `
                            -Method "PUT" `
                            -AccessKey $txtAk.Text `
                            -SecretKey $txtSk.Text `
                            -SessionToken $txtTok.Text `
                            -Region $cmbRegion.Text `
                            -Bucket $parsed.Bucket `
                            -Key $key `
                            -PayloadHash "UNSIGNED-PAYLOAD" `
                            -InFile $src)
                    }
                }
                Write-UiLog "Upload completed successfully."
                if ($cmbBucket.Text.Trim()) { Refresh-S3Objects }
            } else {
                $parsed = Split-S3Uri $src
                if ($chkRecursive.Checked -or $parsed.Key.EndsWith("/") -or [string]::IsNullOrEmpty($parsed.Key)) {
                    $q = @{ "list-type" = "2"; "max-keys" = "1000" }
                    if ($parsed.Key) { $q["prefix"] = $parsed.Key }
                    $xmlText = Invoke-S3SignedRequest `
                        -Method "GET" `
                        -AccessKey $txtAk.Text `
                        -SecretKey $txtSk.Text `
                        -SessionToken $txtTok.Text `
                        -Region $cmbRegion.Text `
                        -Bucket $parsed.Bucket `
                        -QueryParams $q
                    [xml]$doc = $xmlText
                    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
                    $ns.AddNamespace("s3", $doc.DocumentElement.NamespaceURI)
                    $nodes = @($doc.SelectNodes("//s3:Contents", $ns))
                    Write-UiLog "Downloading $($nodes.Count) object(s) from s3://$($parsed.Bucket)/$($parsed.Key) ..."
                    foreach ($c in $nodes) {
                        $kVal = [string]$c.Key
                        if (-not $kVal -or $kVal.EndsWith("/")) { continue }
                        $rel = if ($parsed.Key -and $kVal.StartsWith($parsed.Key)) {
                            $kVal.Substring($parsed.Key.Length).TrimStart('/')
                        } else {
                            $kVal
                        }
                        if (-not $rel) { $rel = Split-Path -Leaf $kVal }
                        $outFile = Join-Path $dst ($rel.Replace('/', '\'))
                        Write-UiLog "download: s3://$($parsed.Bucket)/$kVal -> $outFile"
                        if (-not $chkDryRun.Checked) {
                            [void](Invoke-S3SignedRequest `
                                -Method "GET" `
                                -AccessKey $txtAk.Text `
                                -SecretKey $txtSk.Text `
                                -SessionToken $txtTok.Text `
                                -Region $cmbRegion.Text `
                                -Bucket $parsed.Bucket `
                                -Key $kVal `
                                -OutFile $outFile)
                        }
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                } else {
                    $outFile = if (Test-Path -LiteralPath $dst -PathType Container) {
                        Join-Path $dst (Split-Path -Leaf $parsed.Key)
                    } else {
                        $dst
                    }
                    Write-UiLog "Downloading s3://$($parsed.Bucket)/$($parsed.Key) -> $outFile ..."
                    if (-not $chkDryRun.Checked) {
                        [void](Invoke-S3SignedRequest `
                            -Method "GET" `
                            -AccessKey $txtAk.Text `
                            -SecretKey $txtSk.Text `
                            -SessionToken $txtTok.Text `
                            -Region $cmbRegion.Text `
                            -Bucket $parsed.Bucket `
                            -Key $parsed.Key `
                            -OutFile $outFile)
                    }
                }
                Write-UiLog "Download completed successfully."
                Refresh-LocalExplorer
            }
        } catch {
            Write-UiLog "Transfer error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Transfer Failed") | Out-Null
        }
    }
})

# Load local config if present
if (Test-Path -LiteralPath $Script:LocalConfigFile) {
    try {
        $loaded = Get-Content -LiteralPath $Script:LocalConfigFile -Raw | ConvertFrom-Json
        if ($loaded.aws_access_key_id) { $txtAk.Text = $loaded.aws_access_key_id }
        if ($loaded.aws_secret_access_key) { $txtSk.Text = $loaded.aws_secret_access_key }
        if ($loaded.region) { $cmbRegion.Text = $loaded.region }
        if ($loaded.output_format) { $cmbFormat.Text = $loaded.output_format }
        if ($loaded.session_token) { $txtTok.Text = $loaded.session_token }
        if ($loaded.last_bucket) { $cmbBucket.Text = $loaded.last_bucket }
    } catch {}
}

Refresh-LocalExplorer
Update-Preview
Write-UiLog "Ready. Configure AWS attributes above, pick Source and Target paths, and click 'Start S3 CP Transfer'."

[void]$form.ShowDialog()
