Write-Host "--- Running Windows Update ---"
Write-Host "Searching for updates..."
$updateSession = new-object -com "Microsoft.Update.Session"
$updateSearcher = $updateSession.CreateupdateSearcher()
$searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and AutoSelectOnWebSites=1")
Write-Host "List of applicable items on the machine:"
if ($searchResult.Updates.Count -eq 0) {
 Write-Host "There are no applicable updates."
}
else
{
 $downloadReq = $False
 $i = 0
 foreach ($update in $searchResult.Updates){
  $i++
  if ( $update.IsDownloaded ) {
   Write-Host $i">" $update.Title "(downloaded)"
  }
  else
  {
   $downloadReq = $true
   Write-Host $i">" $update.Title "(not downloaded)"
  }
 }
}
