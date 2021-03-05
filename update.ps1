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
