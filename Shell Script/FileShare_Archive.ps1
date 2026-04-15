param(
    # Auth
    [Parameter(Mandatory=$true)]  [string]$TenantId,
    [Parameter(Mandatory=$true)]  [string]$ClientId,
    [Parameter(Mandatory=$true)]  [string]$ClientSecret,
    [Parameter(Mandatory=$true)]  [string]$SiteHostname,
    [Parameter(Mandatory=$true)]  [string]$SitePath,

    # Rule
    [Parameter(Mandatory=$true)]  [int]   $RuleId,
    [Parameter(Mandatory=$true)]  [string]$RuleName,
    [Parameter(Mandatory=$true)]  [string]$FromPath,
    [Parameter(Mandatory=$true)]  [string]$ToPath,
    [Parameter(Mandatory=$true)]  [string]$ActionType,
    [Parameter(Mandatory=$false)] [string]$FilePattern     = "*.*",
    [Parameter(Mandatory=$true)]  [int]   $FileAgeDays,
    [Parameter(Mandatory=$false)] [string]$Recursive       = "Y",
    [Parameter(Mandatory=$false)] [string]$DeleteEmptyDirs = "N",
    [Parameter(Mandatory=$true)]  [int]   $BatchSize,
    [Parameter(Mandatory=$true)]  [int]   $NoOfBatchRun,

    # Partition (optional - populated only for parallel runs)
    [Parameter(Mandatory=$false)] [string]$DateFrom   = "",
    [Parameter(Mandatory=$false)] [string]$DateTo     = "",
    [Parameter(Mandatory=$false)] [string]$PartLabel  = "",

    # Infra
    [Parameter(Mandatory=$true)]  [string]$LogBasePath,
    [Parameter(Mandatory=$false)] [string]$ProxyHost     = "",
    [Parameter(Mandatory=$false)] [string]$ProxyPort     = "",
    [Parameter(Mandatory=$false)] [string]$SevenZipPath  = "C:\Program Files\7-Zip\7z.exe"
)

# Redirect information stream (Write-Host) to stdout so JVS can capture it
$InformationPreference = 'Continue'

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------- CONFIG ----------
$MaxRetries     = 3
$RetryDelayBase = 2   # seconds
$ChunkSize      = 100 * 320KB

# ---------- PROXY ----------
$proxyArgs = @{}
if (-not [string]::IsNullOrWhiteSpace($ProxyHost) -and -not [string]::IsNullOrWhiteSpace($ProxyPort)) {
    $proxyArgs["Proxy"]                      = "http://${ProxyHost}:${ProxyPort}"
    $proxyArgs["ProxyUseDefaultCredentials"] = $true
}
else {
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.GlobalProxySelection]::GetEmptyWebProxy()
}

# ---------- LOGGING ----------
$logSuffix = if ($PartLabel -ne "") { $PartLabel } else { $RuleName }

if (-not (Test-Path $LogBasePath)) {
    New-Item -ItemType Directory -Path $LogBasePath -Force | Out-Null
}

$LogFile = Join-Path $LogBasePath ("FileArchive_" + $logSuffix + "_" + (Get-Date -Format "yyyyMMdd") + ".log")

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "[{0}] [{1,-5}] [{2}] {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $logSuffix, $Message
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
}

function Write-Section {
    param([string]$Title)
    $line  = "=" * 70
    $entry = "`n$line`n  $Title`n$line"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
}

function Write-Progress-Log {
    param([string]$Message)
    Write-Log ">> $Message"
}

function Get-ErrorMessage {
    param($err)
    try {
        if ($err.ErrorDetails -and $err.ErrorDetails.Message) {
            return $err.ErrorDetails.Message
        }
        elseif ($err.Exception -and $err.Exception.Message) {
            return $err.Exception.Message
        }
        else {
            return $err.ToString()
        }
    } catch {
        return "Unknown error"
    }
}

function Write-ResponseBody {
    param($err, [string]$Prefix)

    try {
        if ($err.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($err.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                Write-Log "$Prefix$responseBody" "ERROR"
            }
        }
    } catch {
        # best effort only
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName
    )

    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $msg = Get-ErrorMessage $_
            Write-Log "$OperationName failed (attempt $i): $msg" "WARN"
            Write-ResponseBody $_ "$OperationName response body: "

            if ($i -eq $MaxRetries) {
                Write-Log "$OperationName failed after $MaxRetries attempts" "ERROR"
                throw
            }

            $delay = [math]::Pow($RetryDelayBase, $i)
            Write-Log "Retrying in $delay sec..."
            Start-Sleep -Seconds $delay
        }
    }
}

# ---------- GRAPH AUTH ----------
function Get-AccessToken {
    Write-Progress-Log "Acquiring OAuth token from Azure AD..."

    $token = Invoke-WithRetry {
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }

        $resp = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded" `
            @proxyArgs

        $resp.access_token
    } "Token Request"

    Write-Log "OAuth token acquired successfully"
    return $token
}

function Get-DriveId {
    param([string]$Token)

    $headers = @{ Authorization = "Bearer $Token" }

    Write-Progress-Log "Resolving SharePoint Drive ID for: $SiteHostname$SitePath"

    $site = Invoke-WithRetry {
        $siteUrl = "https://graph.microsoft.com/v1.0/sites/${SiteHostname}:${SitePath}"
        Write-Log "Get Site URL: $siteUrl"
        Invoke-RestMethod -Uri $siteUrl -Headers $headers @proxyArgs
    } "Get Site"

    $siteId = $site.id
    Write-Log "Site resolved: $siteId"

    $drive = Invoke-WithRetry {
        $driveUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drive"
        Write-Log "Get Drive URL: $driveUrl"
        Invoke-RestMethod -Uri $driveUrl -Headers $headers @proxyArgs
    } "Get Drive"

    Write-Log "Drive ID resolved: $($drive.id)"
    return $drive.id
}

# ---------- SHAREPOINT HELPERS ----------

function Ensure-SpFolder {
    param([string]$Token, [string]$DriveId, [string]$FolderPath)

    $headers   = @{ Authorization = "Bearer $Token" }
    $parts     = $FolderPath.TrimStart('/') -split '/'
    $builtPath = ""

    foreach ($part in $parts) {
        $builtPath = if ($builtPath -eq "") { $part } else { "$builtPath/$part" }
        $checkUrl  = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$builtPath"

        try {
            Invoke-RestMethod -Uri $checkUrl -Headers $headers @proxyArgs | Out-Null
        } catch {
            $parentPath = if ($builtPath -match "/") {
                $builtPath.Substring(0, $builtPath.LastIndexOf('/'))
            } else { "" }

            $parentUrl = if ($parentPath -eq "") {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } else {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${parentPath}:/children"
            }

            $statusCode = $null
            try {
                $statusCode = $_.Exception.Response.StatusCode.value__
            } catch {}

            if ($statusCode -eq 409) {
                Write-Log "SP folder already exists (409 ignored): $builtPath"
            } else {
                Invoke-WithRetry {
                    $folderBody = @{
                        name   = $part
                        folder = @{}
                    } | ConvertTo-Json -Compress

                    Invoke-RestMethod -Method Post -Uri $parentUrl `
                        -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } `
                        -Body $folderBody `
                        @proxyArgs | Out-Null
                } "Create SP Folder [$builtPath]"

                Write-Log "Created SP folder: $builtPath"
            }
        }
    }
}

function Get-UniqueArchiveName {
    param([string]$Token, [string]$DriveId, [string]$FolderPath, [string]$ArchiveBaseName)

    $headers = @{ Authorization = "Bearer $Token" }

    for ($seq = 0; $seq -lt 10000; $seq++) {
        $candidate = if ($seq -eq 0) {
            "${ArchiveBaseName}.zip"
        } else {
            "${ArchiveBaseName}_${seq}.zip"
        }

        $checkUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${FolderPath}/${candidate}"

        try {
            Invoke-RestMethod -Uri $checkUrl -Headers $headers @proxyArgs | Out-Null
        } catch {
            $statusCode = $null
            try {
                $statusCode = $_.Exception.Response.StatusCode.value__
            } catch {}

            if ($statusCode -eq 404) {
                Write-Log "Archive name available: $candidate"
                return $candidate
            }

            throw
        }
    }

    throw "Unable to determine unique archive name for base [$ArchiveBaseName]"
}

function Upload-File {
    param([string]$Token, [string]$DriveId, [string]$LocalPath, [string]$SpFolderPath)

    $headers    = @{ Authorization = "Bearer $Token" }
    $fileName   = Split-Path $LocalPath -Leaf
    $uploadPath = "$SpFolderPath/$fileName"
    $fileSize   = (Get-Item $LocalPath).Length

    if ($fileSize -gt 4MB) {
        Write-Log "Large file ($([math]::Round($fileSize/1MB,2)) MB) - using upload session"

        $sessionUrl  = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${uploadPath}:/createUploadSession"
        $sessionBody = @{
            item = @{ "@microsoft.graph.conflictBehavior" = "replace" }
        } | ConvertTo-Json

        $session = Invoke-WithRetry {
            Invoke-RestMethod -Method Post -Uri $sessionUrl `
                -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } `
                -Body $sessionBody `
                @proxyArgs
        } "Create Upload Session"

       $offset = 0
$stream = [System.IO.File]::OpenRead($LocalPath)
$buffer = New-Object byte[] $ChunkSize

try {
    while ($offset -lt $fileSize) {
        $read = $stream.Read($buffer, 0, $ChunkSize)
        $end  = $offset + $read - 1

        Invoke-WithRetry {
            $webReq                = [System.Net.HttpWebRequest]::Create($session.uploadUrl)
            $webReq.Method         = "PUT"
            $webReq.ContentLength  = $read
            $webReq.Headers.Add("Content-Range", "bytes $offset-$end/$fileSize")

            if ($proxyArgs.ContainsKey("Proxy")) {
                $webReq.Proxy = New-Object System.Net.WebProxy($proxyArgs["Proxy"], $true)
                $webReq.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            }

            $reqStream = $webReq.GetRequestStream()
            $reqStream.Write($buffer, 0, $read)
            $reqStream.Close()

            $resp = $webReq.GetResponse()
            $resp.Close()
        } "Chunk Upload"

        $offset += $read
    }
} finally {
    $stream.Close()
}

    } else {
        $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${uploadPath}:/content"
        $fileBytes = [System.IO.File]::ReadAllBytes($LocalPath)

        Invoke-WithRetry {
            Invoke-RestMethod -Method Put -Uri $uploadUrl `
                -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/octet-stream" } `
                -Body $fileBytes `
                @proxyArgs | Out-Null
        } "Small Upload"
    }
}

# ---------- FILE HELPERS ----------
function Compress-Files {
    param([string]$ZipPath, [string[]]$FilePaths)

    if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

    if (Test-Path $SevenZipPath) {
        $listFile = [System.IO.Path]::GetTempFileName()
        $FilePaths | Set-Content -Path $listFile -Encoding UTF8
        $proc = Start-Process -FilePath $SevenZipPath `
                    -ArgumentList "a", "-tzip", "-mx=3", "-spf", $ZipPath, "@$listFile" `
                    -Wait -PassThru -NoNewWindow
        Remove-Item $listFile -Force
        if ($proc.ExitCode -ne 0) { throw "7-Zip failed with exit code $($proc.ExitCode)" }
        Write-Log "Compression via 7-Zip completed: $(Split-Path $ZipPath -Leaf)"
    } else {
        Write-Log "7-Zip not found at $SevenZipPath - falling back to Compress-Archive" "WARN"
        foreach ($file in $FilePaths) {
            Compress-Archive -Path $file -DestinationPath $ZipPath -Update
        }
        Write-Log "Compression via Compress-Archive completed: $(Split-Path $ZipPath -Leaf)"
    }
}

function Remove-EmptyDirs {
    param([string]$Path)
    @(Get-ChildItem -Path $Path -Recurse -Directory |
        Sort-Object FullName -Descending |
        Where-Object { @((Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue)).Count -eq 0 }) |
        ForEach-Object {
            Remove-Item $_.FullName -Force
            Write-Log "Deleted empty dir: $($_.FullName)"
        }
}

function Remove-FilesParallel {
    param([string[]]$FilePaths)

    $pool = [RunspaceFactory]::CreateRunspacePool(1, 8)
    $pool.Open()
    $jobs = @()

    foreach ($path in $FilePaths) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript({ param($p) Remove-Item $p -Force -ErrorAction SilentlyContinue }).AddArgument($path)
        $jobs += [PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Path = $path }
    }

    $failed = 0
    foreach ($job in $jobs) {
        try {
            $job.PS.EndInvoke($job.Handle)
        } catch {
            Write-Log "DELETE FAILED: $($job.Path) - $_" "WARN"
            $failed++
        } finally {
            $job.PS.Dispose()
        }
    }

    $pool.Close()
    return $failed
}

function Process-ArchiveBucket {
    param(
        [Parameter(Mandatory=$true)] [string]$BucketLabel,
        [Parameter(Mandatory=$true)] [string]$SpFolderPath,
        [Parameter(Mandatory=$true)] [string]$ArchiveBaseName,
        [Parameter(Mandatory=$true)] [System.Object[]]$Files,
        [Parameter(Mandatory=$true)] [string]$Token,
        [Parameter(Mandatory=$true)] [string]$DriveId
    )

    $fileCount   = $Files.Count
    $bucketStart = Get-Date

    Write-Progress-Log "Processing $BucketLabel - $fileCount files"
    Write-Progress-Log "Checking existing zips in SP: $SpFolderPath for $BucketLabel..."

    Ensure-SpFolder -Token $Token -DriveId $DriveId -FolderPath $SpFolderPath
    Write-Log "SP folder confirmed: $SpFolderPath"

    $zipName = Get-UniqueArchiveName -Token $Token -DriveId $DriveId `
                  -FolderPath $SpFolderPath -ArchiveBaseName $ArchiveBaseName
    $zipTemp = Join-Path $env:TEMP $zipName

    Write-Log "Zip name  : $zipName"
    Write-Log "Temp path : $zipTemp"

    try {
        Write-Progress-Log "Compressing $fileCount files -> $zipName"
        Compress-Files -ZipPath $zipTemp -FilePaths ($Files | ForEach-Object { $_.FullName })
    } catch {
        Write-Log "COMPRESSION FAILED for $zipName - $_" "ERROR"
        $script:totalFailed += $fileCount
        return
    }

    if (-not (Test-Path $zipTemp)) {
        Write-Log "ZIP OUTPUT MISSING after compression: $zipTemp" "ERROR"
        $script:totalFailed += $fileCount
        return
    }

    $zipSizeMB = [math]::Round((Get-Item $zipTemp).Length / 1MB, 2)
    Write-Log "Zip size  : ${zipSizeMB} MB"

    try {
        Write-Progress-Log "Uploading $zipName (${zipSizeMB} MB) -> SP: $SpFolderPath"
        Upload-File -Token $Token -DriveId $DriveId `
                    -LocalPath $zipTemp -SpFolderPath $SpFolderPath
        Write-Log "Upload successful: $zipName"

        Write-Progress-Log "Deleting $fileCount source files (parallel)..."
        $delFailed   = Remove-FilesParallel -FilePaths ($Files | ForEach-Object { $_.FullName })
        $delSuccess  = $fileCount - $delFailed
        $script:totalMoved += $delSuccess
        $script:totalFailed += $delFailed
        Write-Log "Source deletion: $delSuccess succeeded | $delFailed failed"
    } catch {
        Write-Log "UPLOAD FAILED: $zipName - $_" "ERROR"
        Write-ResponseBody $_ "Upload response body: "
        $script:totalFailed += $fileCount
    } finally {
        if (Test-Path $zipTemp) {
            Remove-Item $zipTemp -Force
            Write-Log "Temp zip cleaned: $zipTemp"
        }
    }

    $bucketDuration = [int](New-TimeSpan -Start $bucketStart -End (Get-Date)).TotalSeconds
    Write-Progress-Log "$BucketLabel done in ${bucketDuration}s | Moved: $script:totalMoved | Failed: $script:totalFailed"
}

function Invoke-SwiftProcessedByDay {
    param(
        [Parameter(Mandatory=$true)] [System.Object[]]$Files,
        [Parameter(Mandatory=$true)] [string]$Token,
        [Parameter(Mandatory=$true)] [string]$DriveId
    )

    Write-Section "PROCESSING - YEAR / MONTH BUCKETING (SWIFT_PROCESSED)"

    $groupedByYear = $Files | Group-Object { $_.LastWriteTime.ToString("yyyy") }
    Write-Log "Years to process: $(($groupedByYear | ForEach-Object { $_.Name }) -join ', ')"

    foreach ($yearGroup in $groupedByYear) {
        $year = $yearGroup.Name
        Write-Progress-Log ("--- YEAR: {0} ({1} files) ---" -f $year, $yearGroup.Group.Count)

        $groupedByMonth = $yearGroup.Group | Group-Object { $_.LastWriteTime.ToString("MMM") }
        Write-Log ("Months in {0}: {1}" -f $year, (($groupedByMonth | ForEach-Object { $_.Name }) -join ', '))

        foreach ($monthGroup in $groupedByMonth) {
            $monthAbbr = $monthGroup.Name
            $monthSpPath = "$ToPath/$year/$monthAbbr"
            $uniqueSuffix = if ($PartLabel -ne "") { "_" + $PartLabel } else { "_" + $RuleName }
            $oldestInMonth = @($monthGroup.Group | Sort-Object LastWriteTime | Select-Object -First 1)[0]
            $archiveDayLabel = $oldestInMonth.LastWriteTime.ToString("dd-MMM")
            $archiveBaseName = "${archiveDayLabel}$uniqueSuffix"

            Write-Progress-Log "Refreshing OAuth token for bucket: $year/$monthAbbr"
            $token = Get-AccessToken

            Process-ArchiveBucket -BucketLabel "$year/$monthAbbr" `
                                  -SpFolderPath $monthSpPath `
                                  -ArchiveBaseName $archiveBaseName `
                                  -Files @($monthGroup.Group | Sort-Object LastWriteTime) `
                                  -Token $token `
                                  -DriveId $DriveId
        }

        Write-Progress-Log "Year $year complete."
    }
}

# ---------- MAIN ----------
Write-Section "FILE ARCHIVE JOB START"
Write-Log "Rule ID        : $RuleId"
Write-Log "Rule Name      : $RuleName"
Write-Log "Part Label     : $(if ($PartLabel -ne '') { $PartLabel } else { 'N/A (single)' })"
Write-Log "Action         : $ActionType"
Write-Log "From Path      : $FromPath"
Write-Log "To Path        : $ToPath"
Write-Log "File Pattern   : $FilePattern"
Write-Log "File Age Days  : $FileAgeDays"
Write-Log "Recursive      : $Recursive"
Write-Log "Del Empty Dirs : $DeleteEmptyDirs"
Write-Log "Batch Size     : $BatchSize"
Write-Log "Batch Runs     : $NoOfBatchRun"
Write-Log "Max Files/Run  : $($BatchSize * $NoOfBatchRun)"
Write-Log "Date From      : $(if ($DateFrom -ne '') { $DateFrom } else { 'N/A' })"
Write-Log "Date To        : $(if ($DateTo   -ne '') { $DateTo   } else { 'N/A' })"
Write-Log "Log File       : $LogFile"
Write-Log "7-Zip Path     : $SevenZipPath (available: $(Test-Path $SevenZipPath))"

$cutoffDate  = (Get-Date).AddDays(-$FileAgeDays)
$dateFromBound = if ($DateFrom -ne "") { ([datetime]$DateFrom).Date } else { $null }
$dateToBound   = if ($DateTo   -ne "") { ([datetime]$DateTo).Date.AddDays(1) } else { $null }
$isRecursive = ($Recursive -eq "Y")
$totalMoved  = 0
$totalFailed = 0
$jobStart    = Get-Date

Write-Log "Cutoff date    : $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))"

# One-time startup temp sweep across all zip files in shared TEMP.
try {
    $tempZipPattern = "*.zip"
    $staleZips = @(Get-ChildItem -Path $env:TEMP -Filter $tempZipPattern -File -ErrorAction SilentlyContinue)
    if ($staleZips.Count -gt 0) {
        Write-Log "Startup temp cleanup: removing $($staleZips.Count) zip file(s) from [$env:TEMP] matching [$tempZipPattern]"
        foreach ($staleZip in $staleZips) {
            try {
                Remove-Item $staleZip.FullName -Force -ErrorAction Stop
                Write-Log "Deleted stale temp zip: $($staleZip.FullName)"
            } catch {
                Write-Log "Failed to delete stale temp zip: $($staleZip.FullName) - $(Get-ErrorMessage $_)" "WARN"
            }
        }
    } else {
        Write-Log "Startup temp cleanup: no stale zips found in [$env:TEMP] for pattern [$tempZipPattern]"
    }
} catch {
    Write-Log "Startup temp cleanup skipped due to error: $(Get-ErrorMessage $_)" "WARN"
}

try {
    if ($ActionType -eq "DELETE_EMPTY_DIRS") {
        if ($DeleteEmptyDirs -ne "Y") {
            Write-Log "DELETE_EMPTY_DIRS action skipped because DeleteEmptyDirs flag is [$DeleteEmptyDirs]" "WARN"
            Write-Section "JOB COMPLETE - DELETE SKIPPED"
            Write-Log "Duration: $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
            exit 0
        }

        Write-Section "MODE: DELETE EMPTY DIRECTORIES"
        Write-Progress-Log "Scanning for empty directories under: $FromPath"
        Remove-EmptyDirs -Path $FromPath
        Write-Section "JOB COMPLETE"
        Write-Log "Duration: $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
        exit 0
    }

    Write-Section "AUTHENTICATION"
    $token   = Get-AccessToken
    $driveId = Get-DriveId -Token $token

    Write-Section "FILE ENUMERATION"

    # Validate source path exists before attempting scan
    if (-not (Test-Path $FromPath)) {
        Write-Log "FROM_PATH does not exist or is not accessible: $FromPath" "ERROR"
        Write-Log "Skipping rule - path not found." "WARN"
        Write-Section "JOB COMPLETE - PATH NOT FOUND"
        Write-Log "Duration: $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
        exit 0
    }

    $maxFiles = $BatchSize * $NoOfBatchRun
    Write-Progress-Log "Scanning: $FromPath | pattern=$FilePattern | age>${FileAgeDays}d | recursive=$Recursive"
    if ($DateFrom -ne "") {
        Write-Progress-Log "Date partition: $($dateFromBound.ToString('yyyy-MM-dd'))  ->  $($dateToBound.AddDays(-1).ToString('yyyy-MM-dd'))"
    }

    $getParams = @{
        Path    = $FromPath
        Filter  = $FilePattern
        File    = $true
        Recurse = $isRecursive
    }

    $allFiles = @(Get-ChildItem @getParams |
              Where-Object {
                  $_.LastWriteTime -lt $cutoffDate `
                  -and ($null -eq $dateFromBound -or $_.LastWriteTime -ge $dateFromBound) `
                  -and ($null -eq $dateToBound   -or $_.LastWriteTime -lt $dateToBound)
              } |
              Sort-Object LastWriteTime |
              Select-Object -First $maxFiles)

    Write-Log "Eligible files (capped at $maxFiles): $($allFiles.Count)"

    if ($allFiles.Count -eq 0) {
        Write-Log "No eligible files found. Nothing to process."
        Write-Section "JOB COMPLETE - NO FILES"
        Write-Log "Duration: $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
        exit 0
    }

    $oldest = ($allFiles | Select-Object -First 1).LastWriteTime.ToString("yyyy-MM-dd")
    $newest = ($allFiles | Select-Object -Last  1).LastWriteTime.ToString("yyyy-MM-dd")
    Write-Log "File date range: $oldest  ->  $newest"

    if ($RuleName -eq "SWIFT_PROCESSED") {
        Invoke-SwiftProcessedByDay -Files $allFiles -Token $token -DriveId $driveId
    } else {
        Write-Section "PROCESSING - YEAR / MONTH BUCKETING"
        $groupedByYear = $allFiles | Group-Object { $_.LastWriteTime.ToString("yyyy") }
        Write-Log "Years to process: $(($groupedByYear | ForEach-Object { $_.Name }) -join ', ')"

        foreach ($yearGroup in $groupedByYear) {
            $year       = $yearGroup.Name
            $yearSpPath = "$ToPath/$year"

            Write-Progress-Log ("--- YEAR: {0} ({1} files) ---" -f $year, $yearGroup.Group.Count)
            Write-Progress-Log "Refreshing OAuth token for year: $year"
            $token = Get-AccessToken

            Write-Progress-Log "Ensuring SP folder exists: $yearSpPath"
            Ensure-SpFolder -Token $token -DriveId $driveId -FolderPath $yearSpPath
            Write-Log "SP folder confirmed: $yearSpPath"

            $groupedByMonth = $yearGroup.Group | Group-Object { $_.LastWriteTime.ToString("MMM") }
            Write-Log ("Months in {0}: {1}" -f $year, (($groupedByMonth | ForEach-Object { $_.Name }) -join ', '))

            foreach ($monthGroup in $groupedByMonth) {
                $monthAbbr  = $monthGroup.Name
                $fileCount  = $monthGroup.Group.Count
                $monthStart = Get-Date

                Write-Progress-Log "Processing $year/$monthAbbr - $fileCount files"

                Write-Progress-Log "Checking existing zips in SP: $yearSpPath for $monthAbbr..."
                $uniqueSuffix = if ($PartLabel -ne "") { "_" + $PartLabel } else { "_" + $RuleName }
                $archiveBaseName = "${monthAbbr}_${year}$uniqueSuffix"
                $zipName = Get-UniqueArchiveName -Token $token -DriveId $driveId `
                              -FolderPath $yearSpPath -ArchiveBaseName $archiveBaseName
			    $zipTemp      = Join-Path $env:TEMP $zipName

                Write-Log "Zip name  : $zipName"
                Write-Log "Temp path : $zipTemp"

                Write-Progress-Log "Compressing $fileCount files -> $zipName"
                try {
                    Compress-Files -ZipPath $zipTemp -FilePaths ($monthGroup.Group | ForEach-Object { $_.FullName })
                } catch {
                    Write-Log "COMPRESSION FAILED for $zipName - $_" "ERROR"
                    $totalFailed += $fileCount
                    continue
                }

                if (-not (Test-Path $zipTemp)) {
                    Write-Log "ZIP OUTPUT MISSING after compression: $zipTemp" "ERROR"
                    $totalFailed += $fileCount
                    continue
                }

                $zipSizeMB = [math]::Round((Get-Item $zipTemp).Length / 1MB, 2)
                Write-Log "Zip size  : ${zipSizeMB} MB"

                try {
                    Write-Progress-Log "Uploading $zipName (${zipSizeMB} MB) -> SP: $yearSpPath"
                    Upload-File -Token $token -DriveId $driveId `
                                -LocalPath $zipTemp -SpFolderPath $yearSpPath
                    Write-Log "Upload successful: $zipName"

                    Write-Progress-Log "Deleting $fileCount source files (parallel)..."
                    $delFailed   = Remove-FilesParallel -FilePaths ($monthGroup.Group | ForEach-Object { $_.FullName })
                    $delSuccess  = $fileCount - $delFailed
                    $totalMoved += $delSuccess
                    $totalFailed += $delFailed
                    Write-Log "Source deletion: $delSuccess succeeded | $delFailed failed"

                } catch {
                    Write-Log "UPLOAD FAILED: $zipName - $_" "ERROR"
                    Write-ResponseBody $_ "Upload response body: "
                    $totalFailed += $fileCount
                } finally {
                    if (Test-Path $zipTemp) {
                        Remove-Item $zipTemp -Force
                        Write-Log "Temp zip cleaned: $zipTemp"
                    }
                }

                $monthDuration = [int](New-TimeSpan -Start $monthStart -End (Get-Date)).TotalSeconds
                Write-Progress-Log "$year/$monthAbbr done in ${monthDuration}s | Moved: $totalMoved | Failed: $totalFailed"
            }

            Write-Progress-Log "Year $year complete."
        }
    }

    if ($DeleteEmptyDirs -eq "Y") {
        Write-Section "EMPTY DIRECTORY CLEANUP"
        Write-Progress-Log "Scanning for empty subdirectories under: $FromPath"
        Remove-EmptyDirs -Path $FromPath
        Write-Log "Empty directory cleanup complete"
    }

    Write-Section "JOB SUMMARY"
    Write-Log "Rule           : [$RuleId] $RuleName"
    Write-Log "Part Label     : $(if ($PartLabel -ne '') { $PartLabel } else { 'N/A' })"
    Write-Log "Files moved    : $totalMoved"
    Write-Log "Files failed   : $totalFailed"
    Write-Log "Duration       : $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
    Write-Log "Log file       : $LogFile"

    if ($totalFailed -gt 0) {
        Write-Log "Completed WITH FAILURES - review log for details" "WARN"
        exit 1
    }

    Write-Log "Completed SUCCESSFULLY"
    exit 0

} catch {
    Write-Log "FATAL: $(Get-ErrorMessage $_)" "ERROR"
    Write-ResponseBody $_ "Fatal response body: "
    Write-Log "Stack: $($_.ScriptStackTrace)" "ERROR"
    Write-Section "JOB FAILED"
    Write-Log "Duration: $([int](New-TimeSpan -Start $jobStart -End (Get-Date)).TotalSeconds) seconds"
    exit 1
}
