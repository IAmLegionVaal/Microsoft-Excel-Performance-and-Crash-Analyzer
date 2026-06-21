# Microsoft Excel Performance and Crash Analyzer

Created by **Dewald Pretorius**.

The repository contains the original read-only analyzer plus a guarded `Repair.ps1` helper.

Supported actions:

- `Diagnose`
- `ResetPerformanceCaches`
- `RepairOffice`

```powershell
.\Microsoft_Excel_Performance_and_Crash_Analyzer.ps1
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetPerformanceCaches -WhatIf
.\Repair.ps1 -Action RepairOffice -Confirm
```

Close Excel before cache repair. Existing Office cache and telemetry folders are preserved as timestamped backups. Microsoft 365 Apps Quick Repair may require elevation and closes Office applications. Source-reviewed for Windows PowerShell 5.1; not runtime-tested against every Excel or Office build.
