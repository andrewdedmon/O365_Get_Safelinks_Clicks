#Andrew Dedmon - Powershell to capture all clicks in organization

$username = "service_account@email.com"
$pwdTxt = Get-Content "wherever your encrypted service account password is" 
$securePwd = $pwdTxt | ConvertTo-SecureString
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd

$StartDate = (Get-Date).ToString("MM/dd/yyy")
$EndDate = (Get-Date).AddDays(1).ToString("MM/dd/yyy")
$Date = Get-Date -UFormat "%m_%d_%Y"
$File_Name = "all_clicks_" + $Date + ".csv"
$Location = "where_to_save_all_clicks" + $File_Name 
$File_Name2 = "blocked_clicks_" + $Date + ".csv"
$Location2 = "where_to_save_blocked_clicks" + $File_Name2

$eop_sessions=Get-PSSession |Where-Object ComputerName -Match "outlook.office365.com"|Measure-Object|Select-Object Count

 if($eop_sessions -Match "0"){
   $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credObject -Authentication "Basic" -AllowRedirection
   Import-PSSession $exchangeSession -CommandName Get-URLTrace, Get-MessageTrace -AllowClobber
     }
      

$Page1 = Get-UrlTrace -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page 1 | Export-CSV -Path $Location -Append
$Page2 = Get-UrlTrace -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page 2 | Export-CSV -Path $Location -Append
$Page3 = Get-UrlTrace -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page 3 | Export-CSV -Path $Location -Append
#add more pages if you have more results per day on average

$inputCsv = Import-Csv $Location | Sort-Object "Clicked" -Unique
$inputCsv | Export-Csv -NoTypeInformation $Location

$ClickCsv = Import-Csv $Location
$ClickCsv = $ClickCsv | Where-Object UrlBlocked -eq 'TRUE'
Remove-Item $Location2
$ClickCsv | Export-Csv -NoTypeInformation $Location2

$MessageIds = $ClickCsv | Select "MessageId"
$MessageIds = $MessageIds -replace ('@{MessageId=<','')
$MessageIds = $MessageIds -replace ('>}','')

Get-PSSession | Remove-PSSession