# =====================================================================
# Runbook : Generate-MonthlyPatchReport
# Auth    : System-Assigned Managed Identity
# Runtime : PowerShell 7.2
# Strategy: Pull raw 'properties' from each Resource Graph table with
#           minimal KQL transformation. Do ALL vmId parsing/matching in
#           PowerShell using case-insensitive regex - this avoids KQL's
#           case-sensitive split() and arg_max(*) column-overwrite bugs
#           that caused Unknown/zero values in earlier attempts.
#
# PARAMETERS:
#   StorageAccount  - Name of the Storage Account to upload the report to.
#   StorageRG       - Resource Group containing that Storage Account.
#   ContainerName   - Blob container name (default: 'monthly-reports').
#
# These are intentionally left blank below (no environment-specific
# defaults) so the script has no hardcoded reference to any particular
# tenant. Supply your own values either as Runbook parameters when
# starting the job, or as the "Default value" on each parameter in the
# Automation Account's Runbook > Parameters pane.
# =====================================================================
param(
    [Parameter(Mandatory = $true)]
    [string]$StorageAccount,

    [Parameter(Mandatory = $true)]
    [string]$StorageRG,

    [string]$ContainerName = 'monthly-reports'
)

# ── 1. Connect ────────────────────────────────────────────────────────
Connect-AzAccount -Identity -ErrorAction Stop -WarningAction SilentlyContinue *>$null
Import-Module Az.Storage  -Force -ErrorAction Stop
Import-Module ImportExcel -Force -ErrorAction Stop

# ── 2. Auth token ─────────────────────────────────────────────────────
$token    = (Get-AzAccessToken -ResourceUrl "https://management.azure.com/").Token
$authHdr  = @{ 'Authorization' = "Bearer $token"; 'Content-Type' = 'application/json' }
$graphUrl = "https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01"

# ── 3. Discover subscriptions ─────────────────────────────────────────
$subsResp        = Invoke-RestMethod -Uri "https://management.azure.com/subscriptions?api-version=2022-12-01" -Headers $authHdr -Method Get
$subscriptionIds = @($subsResp.value | Where-Object { $_.state -eq 'Enabled' } | Select-Object -ExpandProperty subscriptionId)
$subNames        = @{}
$subsResp.value  | ForEach-Object { $subNames[$_.subscriptionId] = $_.displayName }
Write-Verbose "Subscriptions: $($subscriptionIds.Count)" -Verbose

# ── 4. Date range ─────────────────────────────────────────────────────
$now       = Get-Date
$startDate = Get-Date -Year $now.Year -Month $now.Month -Day 1   # Change to AddMonths(-1) for prod
$endDate   = $now
$monthStr  = $now.ToString('yyyy-MM')
$startIso  = $startDate.ToString('o')
$endIso    = $endDate.ToString('o')
Write-Verbose "Period: $startIso to $endIso" -Verbose

# ── 5. Paginated Resource Graph helper ────────────────────────────────
function Invoke-GraphQuery {
    param([string]$Kql, [string[]]$SubIds, [hashtable]$Hdr, [string]$Url)
    $all = @(); $skip = $null
    do {
        $b = @{ subscriptions = $SubIds; query = $Kql
                options = @{ resultFormat = "ObjectArray"; '$top' = 1000 } }
        if ($skip) { $b.options.'$skipToken' = $skip }
        $r    = Invoke-RestMethod -Uri $Url -Method Post -Headers $Hdr -Body ($b | ConvertTo-Json -Depth 6)
        $all += $r.data
        $skip = $r.'$skipToken'
    } while ($skip)
    return $all
}

# ── 6. Helper: derive base VM resource ID from a child result's id ────
# Strips any "/xxxResults/yyy" suffix using case-insensitive regex,
# avoiding KQL's case-sensitive split() which caused earlier failures.
function Get-BaseVmId {
    param([string]$RawId)
    if (-not $RawId) { return "" }
    return ($RawId -replace '(?i)/(patchassessmentresults|patchinstallationresults)(/.*)?$', '').ToLower().TrimEnd('/')
}

# ── 7. Query: maintenanceresources (patch RUN status per VM) ──────────
$runKql = "maintenanceresources" +
    " | where type =~ 'microsoft.maintenance/applyupdates'" +
    " | extend startTime = todatetime(properties.startDateTime)" +
    " | where startTime >= datetime('$startIso') and startTime <= datetime('$endIso')" +
    " | project id, resourceId=tostring(properties.resourceId)," +
    "     status=tostring(properties.status)," +
    "     correlationId=tostring(properties.correlationId)," +
    "     maintenanceConfigId=tostring(properties.maintenanceConfigurationId)," +
    "     startTime, endTime=todatetime(properties.endDateTime)"

Write-Verbose "Querying maintenanceresources..." -Verbose
$rawRun = Invoke-GraphQuery -Kql $runKql -SubIds $subscriptionIds -Hdr $authHdr -Url $graphUrl
Write-Verbose "Raw run records: $($rawRun.Count)" -Verbose

# Deduplicate: keep the latest run per VM (in case of multiple maintenance windows)
$runRows = $rawRun | Group-Object resourceId | ForEach-Object {
    $_.Group | Sort-Object { [datetime]$_.startTime } -Descending | Select-Object -First 1
}
Write-Verbose "Deduped run records: $($runRows.Count)" -Verbose

# ── 8. Query: patchinstallationresources (actual install counts) ─────
$installKql = "patchinstallationresources" +
    " | where type =~ 'microsoft.compute/virtualmachines/patchinstallationresults'" +
    "    or  type =~ 'microsoft.hybridcompute/machines/patchinstallationresults'" +
    " | extend startTime = todatetime(properties.startDateTime)" +
    " | where startTime >= datetime('$startIso') and startTime <= datetime('$endIso')" +
    " | project id, startTime," +
    "     installed=coalesce(toint(properties.installedPatchCount),0)," +
    "     failed=coalesce(toint(properties.failedPatchCount),0)," +
    "     notSelected=coalesce(toint(properties.notSelectedPatchCount),0)," +
    "     excluded=coalesce(toint(properties.excludedPatchCount),0)," +
    "     pending=coalesce(toint(properties.pendingPatchCount),0)," +
    "     rebootStatus=tostring(properties.rebootStatus)," +
    "     osType=tostring(properties.osType)"

Write-Verbose "Querying patchinstallationresources..." -Verbose
$rawInstall = Invoke-GraphQuery -Kql $installKql -SubIds $subscriptionIds -Hdr $authHdr -Url $graphUrl
Write-Verbose "Raw install records: $($rawInstall.Count)" -Verbose

# Compute base vmId per record in PowerShell, then dedupe keeping latest
$rawInstall | ForEach-Object { $_ | Add-Member -NotePropertyName vmId -NotePropertyValue (Get-BaseVmId $_.id) -Force }
$installRows = $rawInstall | Group-Object vmId | ForEach-Object {
    $_.Group | Sort-Object { [datetime]$_.startTime } -Descending | Select-Object -First 1
}
Write-Verbose "Deduped install records: $($installRows.Count)" -Verbose
if ($installRows.Count -gt 0) {
    Write-Verbose "SAMPLE_INSTALL: vmId=$($installRows[0].vmId) installed=$($installRows[0].installed) failed=$($installRows[0].failed) osType=$($installRows[0].osType)" -Verbose
}

# ── 9. Query: patchassessmentresources (critical/security pending) ───
$assessKql = "patchassessmentresources" +
    " | where type =~ 'microsoft.compute/virtualmachines/patchassessmentresults'" +
    "    or  type =~ 'microsoft.hybridcompute/machines/patchassessmentresults'" +
    " | project id," +
    "     critical=coalesce(toint(properties.availablePatchCountByClassification.critical),0)," +
    "     security=coalesce(toint(properties.availablePatchCountByClassification.security),0)," +
    "     lastAssessed=tostring(properties.lastModifiedDateTime)," +
    "     osType=tostring(properties.osType)"

Write-Verbose "Querying patchassessmentresources..." -Verbose
$rawAssess = Invoke-GraphQuery -Kql $assessKql -SubIds $subscriptionIds -Hdr $authHdr -Url $graphUrl
Write-Verbose "Raw assess records: $($rawAssess.Count)" -Verbose

$rawAssess | ForEach-Object { $_ | Add-Member -NotePropertyName vmId -NotePropertyValue (Get-BaseVmId $_.id) -Force }
$assessRows = $rawAssess | Group-Object vmId | ForEach-Object {
    $_.Group | Sort-Object { [datetime]$_.lastAssessed } -Descending | Select-Object -First 1
}
Write-Verbose "Deduped assess records: $($assessRows.Count)" -Verbose
if ($assessRows.Count -gt 0) {
    Write-Verbose "SAMPLE_ASSESS: vmId=$($assessRows[0].vmId) osType=$($assessRows[0].osType)" -Verbose
}

# ── 9b. Query: Microsoft.Compute/virtualMachines (VM Size + OS image detail) ──
# Patch APIs only return generic "Windows"/"Linux". To get the specific OS
# version (e.g. "Windows Server 2022 Datacenter") and VM Size, we query the
# actual compute resource. The VM's own 'id' matches maintenanceresources'
# resourceId exactly (no suffix stripping needed).
$computeKql = "resources" +
    " | where type =~ 'microsoft.compute/virtualmachines'" +
    " | project id," +
    "     vmSize=tostring(properties.hardwareProfile.vmSize)," +
    "     osDiskType=tostring(properties.storageProfile.osDisk.osType)," +
    "     imgPublisher=tostring(properties.storageProfile.imageReference.publisher)," +
    "     imgOffer=tostring(properties.storageProfile.imageReference.offer)," +
    "     imgSku=tostring(properties.storageProfile.imageReference.sku)"

Write-Verbose "Querying compute resources for VM Size + OS detail..." -Verbose
$rawCompute = Invoke-GraphQuery -Kql $computeKql -SubIds $subscriptionIds -Hdr $authHdr -Url $graphUrl
Write-Verbose "Raw compute records: $($rawCompute.Count)" -Verbose

$computeLookup = @{}
foreach ($r in $rawCompute) { $computeLookup[$r.id.ToLower()] = $r }
if ($rawCompute.Count -gt 0) {
    Write-Verbose "SAMPLE_COMPUTE: vmSize=$($rawCompute[0].vmSize) imgOffer=$($rawCompute[0].imgOffer) imgSku=$($rawCompute[0].imgSku)" -Verbose
}

# ── 9c. Helper: build a friendly OS version label from image reference ──
function Get-OsVersionLabel {
    param($ComputeRecord, [string]$FallbackOsType)

    if (-not $ComputeRecord -or -not $ComputeRecord.imgSku) {
        return $FallbackOsType
    }

    $sku = $ComputeRecord.imgSku
    $offer = $ComputeRecord.imgOffer

    # Windows Server pattern: e.g. "2022-datacenter-azure-edition", "2019-datacenter"
    if ($offer -match '(?i)WindowsServer') {
        if ($sku -match '(?<year>19|20|22|25)(19|20|22|25)?[-_]?(?<edition>datacenter|standard)?') {
            $yearMatch = [regex]::Match($sku, '(20\d{2})')
            $year = if ($yearMatch.Success) { $yearMatch.Value } else { "" }
            $edition = if ($sku -match '(?i)datacenter') { "Datacenter" }
                       elseif ($sku -match '(?i)standard') { "Standard" }
                       else { "" }
            $core = if ($sku -match '(?i)core') { " Core" } else { "" }
            return "Windows Server $year $edition$core".Trim() -replace '\s+',' '
        }
        return "Windows Server ($sku)"
    }

    # Linux distros: e.g. offer=0001-com-ubuntu-server-jammy, sku=22_04-lts
    if ($offer -match '(?i)ubuntu') {
        $verMatch = [regex]::Match($sku, '(\d{2})[-_](\d{2})')
        if ($verMatch.Success) { return "Ubuntu $($verMatch.Groups[1].Value).$($verMatch.Groups[2].Value) LTS" }
        return "Ubuntu ($sku)"
    }
    if ($offer -match '(?i)rhel') {
        return "RHEL $sku"
    }
    if ($offer -match '(?i)centos') {
        return "CentOS $sku"
    }
    if ($offer -match '(?i)sles|suse') {
        return "SUSE $sku"
    }

    # Fallback: publisher/offer/sku raw, or generic Windows/Linux
    if ($offer -and $sku) { return "$offer $sku" }
    return $FallbackOsType
}

# ── 10. Build lookup hashtables for O(1) joins (much faster than Where-Object loops) ──
$installLookup = @{}
foreach ($r in $installRows) { $installLookup[$r.vmId] = $r }
$assessLookup = @{}
foreach ($r in $assessRows) { $assessLookup[$r.vmId] = $r }

# ── 11. Build consolidated report ─────────────────────────────────────
$report = foreach ($row in $runRows) {
    $vmIdL = (Get-BaseVmId $row.resourceId)
    if (-not $vmIdL) { $vmIdL = $row.resourceId.ToLower().TrimEnd('/') }

    $install = $installLookup[$vmIdL]
    $assess  = $assessLookup[$vmIdL]
    $compute = $computeLookup[$vmIdL]

    $inst  = if ($install) { [int]$install.installed }   else { 0 }
    $fail  = if ($install) { [int]$install.failed }      else { 0 }
    $notS  = if ($install) { [int]$install.notSelected } else { 0 }
    $excl  = if ($install) { [int]$install.excluded }    else { 0 }
    $pend  = if ($install) { [int]$install.pending }     else { 0 }
    $total = $inst + $fail + $notS + $excl + $pend

    $genericOsType = if ($install -and $install.osType) { $install.osType }
              elseif ($assess -and $assess.osType) { $assess.osType }
              elseif ($compute -and $compute.osDiskType) { $compute.osDiskType }
              else { "Unknown" }

    $osType = Get-OsVersionLabel -ComputeRecord $compute -FallbackOsType $genericOsType
    $vmSize = if ($compute -and $compute.vmSize) { $compute.vmSize } else { "Unknown" }

    $reboot = if ($install -and $install.rebootStatus) { $install.rebootStatus } else { "N/A" }

    $maintWindow = if ($row.maintenanceConfigId) { ($row.maintenanceConfigId -split '/')[-1] } else { "N/A" }

    $vmName        = ($row.resourceId -split '/')[-1]
    $resourceGroup = ($row.resourceId -split '/')[4]
    $subId         = ($row.resourceId -split '/')[2]
    $subName       = if ($subNames.ContainsKey($subId)) { $subNames[$subId] } else { $subId }

    $friendlyStatus = switch ($row.status) {
        "UpdateSuccessfullyApplied" { "Succeeded"  }
        "Timedout"                  { "Timed Out"  }
        "CompletedWithWarnings"     { "Warnings"   }
        "Failed"                    { "Failed"     }
        "Canceled"                  { "Cancelled"  }
        default                     { $row.status  }
    }

    $duration = "N/A"
    if ($row.startTime -and $row.endTime) {
        try {
            $diff = ([datetime]$row.endTime) - ([datetime]$row.startTime)
            $duration = if ($diff.TotalHours -ge 1) {
                "{0}h {1}m {2}s" -f [int]$diff.TotalHours, $diff.Minutes, $diff.Seconds
            } else { "{0}m {1}s" -f $diff.Minutes, $diff.Seconds }
        } catch {}
    }

    [PSCustomObject]@{
        'VM Name'           = if ($vmName)        { $vmName }        else { "Unknown" }
        'Resource Group'    = if ($resourceGroup) { $resourceGroup } else { "Unknown" }
        'Subscription'      = $subName
        'Subscription ID'   = $subId
        'OS Type'           = $osType
        'VM Size'           = $vmSize
        'Patch Status'      = $friendlyStatus
        'Raw Status'        = $row.status
        'Total Patches'     = $total
        'Installed'         = $inst
        'Failed Patches'    = $fail
        'Not Selected'      = $notS
        'Excluded'          = $excl
        'Pending'           = $pend
        'Reboot Status'     = $reboot
        'Critical Pend.'    = if ($assess) { [int]$assess.critical } else { 0 }
        'Security Pend.'    = if ($assess) { [int]$assess.security } else { 0 }
        'Last Assessment'   = if ($assess) { $assess.lastAssessed }  else { "N/A" }
        'Duration'          = $duration
        'Maintenance Window'= $maintWindow
        'Patch Run (Start)' = $row.startTime
        'Patch Run (End)'   = $row.endTime
        'Run ID'            = $row.correlationId
    }
}

Write-Verbose "Report rows built: $($report.Count)" -Verbose
$matchedInstall = @($report | Where-Object { $_.'Total Patches' -gt 0 }).Count
$matchedOs      = @($report | Where-Object { $_.'OS Type' -ne 'Unknown' }).Count
Write-Verbose "Rows with patch counts matched: $matchedInstall / $($report.Count)" -Verbose
Write-Verbose "Rows with OS Type matched: $matchedOs / $($report.Count)" -Verbose

# ── 12. Export to Excel ───────────────────────────────────────────────
$tmpFile = Join-Path $env:TEMP "PatchReport_$monthStr.xlsx"

if ($report) {
    $pkg = $report | Export-Excel -Path $tmpFile `
        -WorksheetName 'Detailed Report' -TableName 'PatchData' `
        -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow -PassThru

    # ── Color-code the Patch Status column ──
    # Green = Succeeded, Red = Failed, Orange = Timed Out/Warnings/Cancelled
    $wsDetail = $pkg.Workbook.Worksheets['Detailed Report']
    $statusColLetter = ($wsDetail.Cells["1:1"] | Where-Object { $_.Value -eq 'Patch Status' }).Start.Address -replace '\d',''
    $lastRow = $wsDetail.Dimension.End.Row
    if ($statusColLetter) {
        $rangeAddr = "$statusColLetter`2:$statusColLetter$lastRow"
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rangeAddr `
            -RuleType ContainsText -ConditionValue 'Succeeded' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xC6,0xEF,0xCE)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0x00,0x61,0x00))
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rangeAddr `
            -RuleType ContainsText -ConditionValue 'Failed' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xC7,0xCE)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0x9C,0x00,0x06))
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rangeAddr `
            -RuleType ContainsText -ConditionValue 'Timed Out' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rangeAddr `
            -RuleType ContainsText -ConditionValue 'Warnings' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
    }

    # ── Color-code the Reboot Status column (orange for pending) ──
    $rebootColLetter = ($wsDetail.Cells["1:1"] | Where-Object { $_.Value -eq 'Reboot Status' }).Start.Address -replace '\d',''
    if ($rebootColLetter) {
        $rebootRange = "$rebootColLetter`2:$rebootColLetter$lastRow"
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rebootRange `
            -RuleType ContainsText -ConditionValue 'Required' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
        Add-ConditionalFormatting -Worksheet $wsDetail -Range $rebootRange `
            -RuleType ContainsText -ConditionValue 'Pending' `
            -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) `
            -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
    }

    $subSummary = $report | Group-Object 'Subscription' | ForEach-Object {
        $grp  = $_.Group
        $succ = ($grp | Where-Object { $_.'Patch Status' -eq 'Succeeded' }).Count
        [PSCustomObject]@{
            'Subscription'   = $_.Name
            'Total VMs'      = $grp.Count
            'Succeeded'      = $succ
            'Timed Out'      = ($grp | Where-Object { $_.'Patch Status' -eq 'Timed Out' }).Count
            'Failed'         = ($grp | Where-Object { $_.'Failed Patches' -gt 0 }).Count
            'Pending Reboot' = ($grp | Where-Object { $_.'Reboot Status' -like '*Reboot*' }).Count
            'Success %'      = "$([math]::Round($succ / $grp.Count * 100, 1))%"
        }
    }
    $subSummary | Export-Excel -ExcelPackage $pkg `
        -WorksheetName 'Summary by Subscription' -TableName 'SubSummary' `
        -TableStyle Medium6 -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $needsAction = @($report | Where-Object { $_.'Patch Status' -in @('Failed','Timed Out','Warnings') -or $_.'Reboot Status' -like '*Reboot*' -or $_.'Reboot Status' -like '*Required*' -or $_.'Reboot Status' -like '*Pending*' })
    if ($needsAction.Count -gt 0) {
        $needsAction | Select-Object 'VM Name','Resource Group','Subscription','OS Type','VM Size',
            'Patch Status','Total Patches','Installed','Failed Patches',
            'Reboot Status','Critical Pend.','Last Assessment','Maintenance Window','Patch Run (Start)','Run ID' |
        Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'Needs Attention' -TableName 'NeedsAction' `
            -TableStyle Medium3 -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

        $wsNeeds = $pkg.Workbook.Worksheets['Needs Attention']
        $needsStatusCol = ($wsNeeds.Cells["1:1"] | Where-Object { $_.Value -eq 'Patch Status' }).Start.Address -replace '\d',''
        $needsLastRow = $wsNeeds.Dimension.End.Row
        if ($needsStatusCol) {
            $r = "$needsStatusCol`2:$needsStatusCol$needsLastRow"
            Add-ConditionalFormatting -Worksheet $wsNeeds -Range $r -RuleType ContainsText -ConditionValue 'Failed' `
                -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xC7,0xCE)) -ForegroundColor ([System.Drawing.Color]::FromArgb(0x9C,0x00,0x06))
            Add-ConditionalFormatting -Worksheet $wsNeeds -Range $r -RuleType ContainsText -ConditionValue 'Timed Out' `
                -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
        }
    }

    $pendingReboot = @($report | Where-Object { $_.'Reboot Status' -like '*Reboot*' -or $_.'Reboot Status' -like '*Required*' -or $_.'Reboot Status' -like '*Pending*' })
    if ($pendingReboot.Count -gt 0) {
        $pendingReboot | Select-Object 'VM Name','Resource Group','Subscription',
            'OS Type','VM Size','Patch Status','Reboot Status',
            'Last Assessment','Maintenance Window','Patch Run (Start)','Run ID' |
        Export-Excel -ExcelPackage $pkg `
            -WorksheetName 'Pending Reboot' -TableName 'PendingReboot' `
            -TableStyle Medium7 -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

        $wsReboot = $pkg.Workbook.Worksheets['Pending Reboot']
        $rebootLastRow = $wsReboot.Dimension.End.Row
        $rebootStatusColLtr = ($wsReboot.Cells["1:1"] | Where-Object { $_.Value -eq 'Reboot Status' }).Start.Address -replace '\d',''
        if ($rebootStatusColLtr) {
            $r2 = "$rebootStatusColLtr`2:$rebootStatusColLtr$rebootLastRow"
            Add-ConditionalFormatting -Worksheet $wsReboot -Range $r2 -RuleType ContainsText -ConditionValue 'Required' `
                -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
            Add-ConditionalFormatting -Worksheet $wsReboot -Range $r2 -RuleType ContainsText -ConditionValue 'Pending' `
                -BackgroundColor ([System.Drawing.Color]::FromArgb(0xFF,0xE0,0xB2)) -ForegroundColor ([System.Drawing.Color]::FromArgb(0xB2,0x5C,0x00))
        }
    }

    Close-ExcelPackage $pkg

} else {
    [PSCustomObject]@{
        'VM Name'='No data'; 'Resource Group'=''; 'Subscription'=''; 'Subscription ID'='';
        'OS Type'=''; 'VM Size'=''; 'Patch Status'="No patch runs found for $monthStr";
        'Raw Status'=''; 'Total Patches'=0; 'Installed'=0; 'Failed Patches'=0;
        'Not Selected'=0; 'Excluded'=0; 'Pending'=0; 'Reboot Status'='';
        'Critical Pend.'=0; 'Security Pend.'=0; 'Last Assessment'='';
        'Duration'=''; 'Maintenance Window'=''; 'Patch Run (Start)'=''; 'Patch Run (End)'=''; 'Run ID'=''
    } | Export-Excel -Path $tmpFile -WorksheetName 'Detailed Report' -AutoSize
}

Write-Verbose "Excel created: $tmpFile" -Verbose

# ── 13. Upload to Blob ────────────────────────────────────────────────
$ctx      = New-AzStorageContext -StorageAccountName $StorageAccount -UseConnectedAccount
$blobName = "AzureUpdateManager_PatchReport_$monthStr.xlsx"
Set-AzStorageBlobContent -Context $ctx -Container $ContainerName `
    -File $tmpFile -Blob $blobName -Force | Out-Null
Write-Verbose "Uploaded: $blobName" -Verbose

# ── 14. Output for Logic App (ONLY Write-Output in script) ───────────
[PSCustomObject]@{
    BlobUrl       = "https://$StorageAccount.blob.core.windows.net/$ContainerName/$blobName"
    BlobPath      = "$ContainerName/$blobName"
    BlobName      = $blobName
    Month         = $monthStr
    Subscriptions = $subscriptionIds.Count
    Total         = if ($report) { @($report).Count } else { 0 }
    Failed        = if ($report) { @($report | Where-Object { $_.'Failed Patches' -gt 0 }).Count } else { 0 }
} | ConvertTo-Json -Compress | Write-Output
