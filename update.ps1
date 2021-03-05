#Sets "Active Hours" so servers dont reboot automatically
$ahStart = 17 #5PM
$ahEnd = 23   #11PM
Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name ActiveHoursStart -Value $ahStart -PassThru
Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name ActiveHoursEnd -Value $ahEnd -PassThru 

#Define update criteria
$Criteria = "IsInstalled=0"

#Search for relevant updates
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$SearchResult = $Searcher.Search($Criteria).Updates
$SearchResult | Select -ExpandProperty Title

#Download updates
$Session = New-Object -ComObject Microsoft.Update.Session
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $SearchResult
$Downloader.Download()

#Install updates
$Installer = New-Object -ComObject Microsoft.Update.Installer
$Installer.Updates = $SearchResult
$SearchResult = $Installer.Install()

Write-Output "Update Complete"
