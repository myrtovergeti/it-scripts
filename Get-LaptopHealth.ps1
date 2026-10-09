<#
.SYNOPSIS
    Έλεγχος υγείας laptop: μνήμη, δίσκος, επανεκκίνηση, μπαταρία, Windows Update, Dell BIOS/updates.

.DESCRIPTION
    Μόνο ανάγνωση. Δεν εγκαθιστά και δεν αλλάζει τίποτα στο σύστημα.
    Με -DellScan ελέγχει μέσω Dell Command | Update αν υπάρχουν διαθέσιμα updates
    (χρειάζεται PowerShell ως Administrator και εγκατεστημένο Dell Command | Update).
    Με -CsvPath προσθέτει το αποτέλεσμα ως γραμμή σε CSV (ανοίγει σε Excel).

.EXAMPLE
    .\Get-LaptopHealth.ps1
    .\Get-LaptopHealth.ps1 -DellScan
    .\Get-LaptopHealth.ps1 -CsvPath "$env:USERPROFILE\Desktop\laptop-health.csv"
#>

[CmdletBinding()]
param(
    [int]$MinFreeDiskPercent = 20,
    [int]$MinFreeRamPercent  = 10,
    [int]$MaxUptimeDays      = 14,
    [int]$MaxUpdateAgeDays   = 45,
    [switch]$DellScan,
    [string]$CsvPath
)

$warnings = @()
$now = Get-Date

# --- Λειτουργικό & επανεκκίνηση ---
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$lastBoot   = $os.LastBootUpTime
$uptimeDays = [math]::Round(($now - $lastBoot).TotalDays, 1)
if ($uptimeDays -gt $MaxUptimeDays) {
    $warnings += "Έχει να κάνει επανεκκίνηση $uptimeDays ημέρες (από $($lastBoot.ToString('yyyy-MM-dd HH:mm')))."
}

# --- Μνήμη ---
$totalRamGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
$freeRamGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
$freeRamPct = if ($os.TotalVisibleMemorySize -gt 0) {
    [math]::Round(($os.FreePhysicalMemory / $os.TotalVisibleMemorySize) * 100, 0)
} else { 0 }
if ($freeRamPct -lt $MinFreeRamPercent) {
    $warnings += "Ελεύθερη μνήμη $freeRamPct% (< $MinFreeRamPercent%)."
}

# --- Δίσκος συστήματος ---
$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'"
$diskTotalGB = [math]::Round($disk.Size / 1GB, 1)
$diskFreeGB  = [math]::Round($disk.FreeSpace / 1GB, 1)
$diskFreePct = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0) } else { 0 }
if ($diskFreePct -lt $MinFreeDiskPercent) {
    $warnings += "Ελεύθερος χώρος δίσκου $diskFreePct% (< $MinFreeDiskPercent%)."
}

# --- Μπαταρία ---
$battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
$batteryText = if ($battery) { "$($battery.EstimatedChargeRemaining)%" } else { "Δεν βρέθηκε" }

# --- Τελευταίο Windows Update ---
$lastHotfix = Get-HotFix -ErrorAction SilentlyContinue |
    Where-Object { $_.InstalledOn } |
    Sort-Object InstalledOn -Descending |
    Select-Object -First 1
$lastWinUpdate = if ($lastHotfix) { $lastHotfix.InstalledOn.ToString('yyyy-MM-dd') } else { "Άγνωστο" }
if ($lastHotfix -and (($now - $lastHotfix.InstalledOn).TotalDays -gt $MaxUpdateAgeDays)) {
    $warnings += "Τελευταίο Windows Update πριν από $MaxUpdateAgeDays+ ημέρες ($lastWinUpdate)."
}

# --- Dell: μοντέλο, BIOS, Dell Command | Update ---
$cs    = Get-CimInstance -ClassName Win32_ComputerSystem
$bios  = Get-CimInstance -ClassName Win32_BIOS
$isDell = $cs.Manufacturer -like 'Dell*'

$biosText = "$($bios.SMBIOSBIOSVersion)"
if ($bios.ReleaseDate) {
    $biosText += " (έκδοση της $($bios.ReleaseDate.ToString('yyyy-MM-dd')))"
}

$dcuPath = @(
    (Join-Path $env:ProgramFiles 'Dell\CommandUpdate\dcu-cli.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'Dell\CommandUpdate\dcu-cli.exe')
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

$dellUpdatesText = if (-not $isDell) {
    "Δεν είναι Dell"
} elseif (-not $dcuPath) {
    "Το Dell Command | Update δεν είναι εγκατεστημένο"
} elseif (-not $DellScan) {
    "Δεν ελέγχθηκε (τρέξε με -DellScan)"
} else {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        "Χρειάζεται PowerShell ως Administrator"
    } else {
        $reportDir = Join-Path $env:TEMP ("dcu-report-" + $now.ToString('yyyyMMddHHmmss'))
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        $logFile = Join-Path $reportDir 'dcu-output.txt'

        Write-Host "Έλεγχος Dell updates (μπορεί να πάρει 1-3 λεπτά)..." -ForegroundColor Cyan
        $proc = Start-Process -FilePath $dcuPath `
            -ArgumentList '/scan', "-report=`"$reportDir`"" `
            -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $logFile

        switch ($proc.ExitCode) {
            500 { "Κανένα διαθέσιμο update" }
            0 {
                $xmlFile = Get-ChildItem -Path $reportDir -Filter '*.xml' -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                if ($xmlFile) {
                    [xml]$report = Get-Content -Path $xmlFile.FullName -Raw
                    $updates = @($report.SelectNodes('//update'))
                    if ($updates.Count -gt 0) {
                        $names = ($updates | ForEach-Object { $_.name } | Where-Object { $_ }) -join '; '
                        $warnings += "Υπάρχουν $($updates.Count) διαθέσιμα Dell updates."
                        "$($updates.Count) διαθέσιμα: $names"
                    } else {
                        "Κανένα διαθέσιμο update"
                    }
                } else {
                    "Ο έλεγχος ολοκληρώθηκε, αλλά δεν βρέθηκε αναφορά (δες $logFile)"
                }
            }
            default { "Σφάλμα ελέγχου, κωδικός $($proc.ExitCode) (δες $logFile)" }
        }
    }
}

# --- Αποτέλεσμα ---
$result = [PSCustomObject][ordered]@{
    'Ημ/νία ελέγχου'         = $now.ToString('yyyy-MM-dd HH:mm')
    'Υπολογιστής'            = $env:COMPUTERNAME
    'Μοντέλο'                = "$($cs.Manufacturer) $($cs.Model)"
    'Windows'                = "$($os.Caption) ($($os.Version))"
    'Τελ. επανεκκίνηση'      = $lastBoot.ToString('yyyy-MM-dd HH:mm')
    'Μέρες χωρίς restart'    = $uptimeDays
    'RAM ελεύθερη (GB)'      = "$freeRamGB / $totalRamGB ($freeRamPct%)"
    'Δίσκος ελεύθερος (GB)'  = "$diskFreeGB / $diskTotalGB ($diskFreePct%)"
    'Μπαταρία'               = $batteryText
    'Τελ. Windows Update'    = $lastWinUpdate
    'BIOS'                   = $biosText
    'Dell updates'           = $dellUpdatesText
    'Προειδοποιήσεις'        = ($warnings -join ' | ')
}

$result | Format-List

if ($warnings.Count -gt 0) {
    Write-Host "Προειδοποιήσεις:" -ForegroundColor Yellow
    $warnings | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
} else {
    Write-Host "Όλα εντάξει." -ForegroundColor Green
}

if ($CsvPath) {
    $result | Export-Csv -Path $CsvPath -Append -NoTypeInformation -Encoding UTF8
    Write-Host "Αποθηκεύτηκε στο: $CsvPath" -ForegroundColor Cyan
}
