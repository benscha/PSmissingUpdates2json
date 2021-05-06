#Global
$hostname = [System.Net.Dns]::GetHostByName($env:computerName) | Select -ExpandProperty Hostname | Out-String
$LastUpdate = get-Hotfix| select InstalledON  | sort { [datetime]$_.InstalledON  } -desc| Select-Object -First 1 
$LogDate = (Get-Date -Format "yyyy.MM.dd_HH_mm")
$LastBootTime = Get-WmiObject win32_operatingsystem | select @{LABEL=’LastBootUpTime’;EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
$hostinfo = [pscustomobject]@{'hostname'=$hostname;'LastUpdated'= $LastUpdate.InstalledOn;'LogDate'= (Get-Date -Format "dd.MM.yyyy HH:mm");'LastBoot'=$LastBootTime.lastbootuptime}
$FileName = $env:computerName +"_" + $LogDate.ToString() + ".json" 




#Collect RAW Data
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateupdateSearcher()
$Updates1 = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)
$Updates1 | Select-Object *


#Prepare Data for Arra<y
$UpdTypes = $Updates1.Categories.Name.Split(“`n”) 
$UpdTitle = $Updates1.Title.Split(“`n”) 
$UpdKB = $Updates1.KBArticleIDs
$UpdRelDate = $Updates1.LastDeploymentChangeTime


# Combine Arrays
[int]$max = $UpdTitle.Count
if ([int]$UpdTypes.count -gt [int]$UpdTitle.count) { $max = $UpdTypes.Count; }
 
$ResultsAll = for ( $i = 0; $i -lt $max; $i++)
{
    Write-Verbose "$($UpdTypes[$i]),$($UpdTitle[$i]),$($UpdKB[$i]),$($UpdRelDate[$i])"
    [PSCustomObject]@{
        UpdateType = $UpdTypes[$i]
        Title = $UpdTitle[$i]
        KBNo = $UpdKB[$i]
        ReleaseDate = $UpdRelDate[$i]      
    }
}
$Results = $ResultsAll | Where-Object {$_.UpdateType -ne "Drivers"} 


$jsonBase = @{}
$jsonBase.Add("Data",$Results)
$jsonBase.Add("Hostinfo",$hostinfo)
$jsonBase | ConvertTo-Json -Depth 10 | Out-File $FileName
