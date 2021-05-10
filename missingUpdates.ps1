###Script###
#Global
$hostname = [System.Net.Dns]::GetHostByName($env:computerName) | Select -ExpandProperty Hostname
$LastUpdate = Get-Date -Format "yyyy.MM.dd HH:mm:ss" (get-Hotfix | select InstalledON  | sort { [datetime]$_.InstalledON  } -desc| Select-Object -ExpandProperty InstalledOn -First 1)
$LogDate = (Get-Date -Format "yyyy.MM.dd HH:mm:ss")
$LastBootTime = Get-Date -Format "yyyy.MM.dd HH:mm:ss" (Get-WmiObject win32_operatingsystem | select @{LABEL=’LastBootUpTime’;EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}} | Select-Object -ExpandProperty LastBootUpTime )
$os = (Get-WmiObject -class Win32_OperatingSystem).Caption
$hostinfo = [pscustomobject]@{'hostname'=$hostname;'LastUpdated'= $LastUpdate;'LogDate'= (Get-Date -Format "yyyy.MM.dd HH:mm:ss");'LastBoot'=$LastBootTime;'os'=$os}

#Collect RAW Data
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateupdateSearcher()
$Updates1 = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)

#Prepare Data for Arra<y
$UpdTypes = $Updates1.Categories.Name.Split(“`n”) 
$UpdTitle = $Updates1.Title.Split(“`n”) 
$UpdKB = $Updates1.KBArticleIDs
$UpdRelDate = $Updates1.LastDeploymentChangeTime

# Combine Arrays
[int]$max = $UpdTitle.Count

$ResultsAll = @()

foreach ($update in $Updates1) {
    $ResultsAll += (
        [PSCustomObject]@{
            UpdateType = $update.Categories[0].Name
            Title = $update.Title
            KBNo = $update.KBArticleIDs[0]
            ReleaseDate = Get-Date -Format "yyyy.MM.dd HH:mm:ss" $update.LastDeploymentChangeTime
            }
        )
}

$Results = $ResultsAll | Where-Object {$_.UpdateType -ne "Drivers"} 

$jsonBase = @{}
$jsonBase.Add("Data",$Results)
$jsonBase.Add("Hostinfo",$hostinfo)
$jsonBase = $jsonBase | ConvertTo-Json



