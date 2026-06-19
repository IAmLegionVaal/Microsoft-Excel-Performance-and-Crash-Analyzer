#requires -Version 5.1
<# Created by Dewald Pretorius #>
[CmdletBinding()]
param([string]$OutputPath)
$ErrorActionPreference='SilentlyContinue'
if(-not $OutputPath){$OutputPath=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Excel_Performance_Reports'}
New-Item -ItemType Directory -Path $OutputPath -Force|Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$report=Join-Path $OutputPath "Excel_Analysis_$stamp.txt"
$findings=@()
function Add-Finding{param($Area,$Status,$Detail,$Recommendation);[pscustomobject]@{Area=$Area;Status=$Status;Detail=$Detail;Recommendation=$Recommendation}}
$p=Get-Process EXCEL -ErrorAction SilentlyContinue
$findings+=Add-Finding 'Excel process' ($(if($p){'Detected'}else{'Not running'})) ($(if($p){"Instances=$($p.Count); WorkingSetMB=$([math]::Round((($p|Measure-Object WorkingSet -Sum).Sum/1MB),1))"}else{'No active Excel process'})) 'Close duplicate instances and reproduce the issue with one workbook.'
$addinRoots='HKCU:\Software\Microsoft\Office\Excel\Addins','HKLM:\Software\Microsoft\Office\Excel\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\Excel\Addins'
$addins=foreach($root in $addinRoots){Get-ChildItem $root -ErrorAction SilentlyContinue|ForEach-Object{[pscustomobject]@{Root=$root;Name=$_.PSChildName;LoadBehavior=(Get-ItemProperty $_.PSPath).LoadBehavior}}}
$findings+=Add-Finding 'Add-ins' ($(if($addins){'Review'}else{'None found'})) "Count=$($addins.Count)" 'Test Excel Safe Mode and disable nonessential add-ins one at a time.'
$xlstart=@("$env:APPDATA\Microsoft\Excel\XLSTART","$env:ProgramFiles\Microsoft Office\root\Office16\XLSTART")
$startupFiles=foreach($folder in $xlstart){Get-ChildItem $folder -File -ErrorAction SilentlyContinue}
$findings+=Add-Finding 'Startup files' ($(if($startupFiles){'Review'}else{'Pass'})) "Count=$($startupFiles.Count)" 'Temporarily move unexpected XLSTART files when troubleshooting startup crashes.'
$crashes=Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddDays(-7)}|Where-Object{$_.Message -match 'EXCEL.EXE|Excel'}|Select-Object -First 40 TimeCreated,Id,ProviderName,LevelDisplayName,Message
$findings+=Add-Finding 'Recent events' ($(if($crashes){'Review'}else{'Pass'})) "Count=$($crashes.Count)" 'Correlate faulting modules with add-ins, graphics drivers, Office builds, or specific workbooks.'
$temp=Get-ChildItem $env:TEMP -File -ErrorAction SilentlyContinue|Where-Object{$_.Name -match '^~\$|Excel'}
$findings+=Add-Finding 'Temporary files' ($(if($temp.Count -gt 25){'Review'}else{'Pass'})) "Count=$($temp.Count)" 'Close Office apps before removing stale temporary files.'
$office=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$findings+=Add-Finding 'Office build' ($(if($office){'Detected'}else{'Review'})) "Version=$($office.VersionToReport); Platform=$($office.Platform)" 'Compare the installed build with your approved enterprise update channel.'
$findings|Export-Csv (Join-Path $OutputPath "Excel_Findings_$stamp.csv") -NoTypeInformation -Encoding UTF8
@('MICROSOFT EXCEL PERFORMANCE AND CRASH ANALYZER','Created by Dewald Pretorius',"Generated: $(Get-Date)",'',($findings|Format-Table -AutoSize|Out-String -Width 240),'ADD-INS',($addins|Format-Table -AutoSize|Out-String -Width 240),'RECENT EVENTS',($crashes|Format-List|Out-String -Width 240))|Set-Content $report -Encoding UTF8
Write-Host "Report created: $report" -ForegroundColor Green
