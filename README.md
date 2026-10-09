# IT Scripts

Συλλογή βοηθητικών scripts για την ομάδα IT.

| Script | Τι κάνει |
|---|---|
| `Get-LaptopHealth.ps1` | Έλεγχος δίσκου, RAM, uptime, μπαταρίας και Windows Update σε laptop |

## Χρήση

```powershell
.\Get-LaptopHealth.ps1
.\Get-LaptopHealth.ps1 -MinFreeDiskPercent 20 -MaxUptimeDays 7
```

## Κανόνες

- Μην ανεβάζετε κωδικούς, API keys, tokens ή προσωπικά δεδομένα.
- Κάθε αλλαγή μπαίνει μέσω Pull Request, με έλεγχο από δεύτερο άτομο.
