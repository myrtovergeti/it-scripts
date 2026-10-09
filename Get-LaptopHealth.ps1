<#
.SYNOPSIS
    Γρήγορος έλεγχος υγείας laptop (δίσκος, μνήμη, uptime, μπαταρία, Windows Update).

.DESCRIPTION
    Εκτελείται τοπικά, μόνο ανάγνωση. Δεν αλλάζει τίποτα στο σύστημα
    και δεν συλλέγει προσωπικά δεδομένα χρήστη.

.EXAMPLE
    .\Get-LaptopHealth.ps1
    .\Get-LaptopHealth.ps1 -MinFreeDiskPercent 20
#>

[CmdletBinding()]
param(
    [int]$MinFreeDiskPercent = 15,
    [int]$MaxUptimeDays = 14
)

$warnings = @()

# --- Λειτουργικό ---
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$uptimeDays = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
if ($uptimeDays -gt $MaxUptimeDays) {
    $warnings += "Uptime $uptimeDays ημέρες - προτείνεται επανεκκίνηση."
}

# --- Μνήμη ---
$totalRamGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
$freeRamGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)

# --- Δίσκος συστήματος ---
$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'"
$diskTotalGB = [math]::Round($disk.Size / 1GB, 1)
$diskFreeGB  = [math]::Round($disk.FreeSpace / 1GB, 1)
$diskFreePct = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0) } else { 0 }
if ($diskFreePct -lt $MinFreeDiskPercent) {
    $warnings += "Ελεύθερος χώρος δίσκου $diskFreePct% (< $MinFreeDiskPercent%)."
}

# --- Μπαταρία (αν υπάρχει) ---
$battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
$batteryText = if ($battery) { "$($battery.EstimatedChargeRemaining)%" } else { "Δεν βρέθηκε" }

# --- Τελευταίο Windows Update ---
$lastHotfix = Get-HotFix -ErrorAction SilentlyContinue |
    Where-Object { $_.InstalledOn } |
    Sort-Object InstalledOn -Descending |
    Select-Object -First 1
$lastUpdate = if ($lastHotfix) { $lastHotfix.InstalledOn.ToString('yyyy-MM-dd') } else { "Άγνωστο" }
if ($lastHotfix -and ((Get-Date) - $lastHotfix.InstalledOn).TotalDays -gt 45) {
    $warnings += "Τελευταίο update πριν από 45+ ημέρες ($lastUpdate)."
}

# --- Αποτέλεσμα ---
[PSCustomObject]@{
    Υπολογιστής     = $env:COMPUTERNAME
    Windows         = "$($os.Caption) ($($os.Version))"
    'Uptime (ημ.)'  = $uptimeDays
    'RAM (GB)'      = "$freeRamGB ελεύθερα / $totalRamGB"
    'Δίσκος (GB)'   = "$diskFreeGB ελεύθερα / $diskTotalGB ($diskFreePct%)"
    Μπαταρία        = $batteryText
    'Τελ. Update'   = $lastUpdate
} | Format-List

if ($warnings.Count -gt 0) {
    Write-Host "Προειδοποιήσεις:" -ForegroundColor Yellow
    $warnings | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
} else {
    Write-Host "Όλα εντάξει." -ForegroundColor Green
}
