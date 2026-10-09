# IT Scripts

Συλλογή βοηθητικών scripts για την ομάδα IT.

| Script | Τι κάνει |
|---|---|
| `Get-LaptopHealth.ps1` | Έλεγχος RAM, δίσκου, τελευταίας επανεκκίνησης, μπαταρίας, Windows Update, Dell BIOS και Dell updates |

## Χρήση

```powershell
.\Get-LaptopHealth.ps1                                   # βασικός έλεγχος
.\Get-LaptopHealth.ps1 -DellScan                          # + έλεγχος Dell updates (ως Administrator)
.\Get-LaptopHealth.ps1 -CsvPath "$env:USERPROFILE\Desktop\laptop-health.csv"   # + αποθήκευση σε CSV
.\Get-LaptopHealth.ps1 -MaxUptimeDays 7 -MinFreeDiskPercent 25                 # δικά σου όρια
```

Για το `-DellScan` χρειάζεται εγκατεστημένο **Dell Command | Update**.

## Κανόνες

- Μην ανεβάζετε κωδικούς, API keys, tokens ή προσωπικά δεδομένα.
- Κάθε αλλαγή μπαίνει μέσω Pull Request, με έλεγχο από δεύτερο άτομο.
