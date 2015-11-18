$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()
$HistoryCount = $Searcher.GetTotalHistoryCount()
$Searcher.QueryHistory(1,$HistoryCount) | ft title

#Get-WMIObject Win32_QuickFixEngineering
