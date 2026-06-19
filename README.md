# Microsoft Excel Performance and Crash Analyzer

Created by **Dewald Pretorius**.

A read-only PowerShell toolkit for investigating slow Excel performance, excessive memory usage, startup crashes, add-in conflicts, XLSTART files, temporary files, Office builds, and recent application faults.

## Checks

- Running Excel processes and total memory usage
- COM and application add-ins
- XLSTART startup folders
- Recent Excel-related Application log events
- Stale or excessive temporary files
- Microsoft 365 Click-to-Run version and platform

## Run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Microsoft_Excel_Performance_and_Crash_Analyzer.ps1"
```

Reports are saved to `Desktop\Excel_Performance_Reports` in TXT and CSV format.

## Useful scenarios

- Excel opens slowly
- Workbooks freeze or stop responding
- Excel crashes during startup
- High CPU or memory use
- Add-in conflicts
- Problems that disappear in Safe Mode
- Crashes tied to a particular Office build or workbook

## Safety

The analyzer does not disable add-ins, delete files, modify Office, or repair workbooks. It gathers evidence and provides next-step guidance.
