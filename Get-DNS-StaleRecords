 #set parameters  
 $dnsServer = "dc01"  
 $domain = "contoso.com"  
 $agetreshold = 14  
   
 # calculate how many hours is the age which will be the threshold  
 $minimumTimeStamp = [int] (New-TimeSpan -Start $(Get-Date ("01/01/1601 00:00")) -End $((Get-Date).AddDays(-$agetreshold))).TotalHours  
   
 # get all records from the zone whose age is more than our threshold   
 $records = Get-WmiObject -ComputerName $dnsServer -Namespace "root\MicrosoftDNS" -Query "select * from MicrosoftDNS_AType where Containername='$domain' AND TimeStamp<$minimumTimeStamp AND TimeStamp<>0 "  
   
 # list the name and the calculated last update time stamp  
 $stale = $records | Select Ownername, @{n="timestamp";e={([datetime]"1.1.1601").AddHours($_.Timestamp)}} , IPAddress
 $result = @()

 foreach ($record in $stale){
    if (Test-Connection $record.IpAddress -Count 1 -ThrottleLimit 2 -ErrorAction SilentlyContinue){
        write-host $record.Ownername "is Online" -ForegroundColor Green
        $result += $record.Ownername + " - " + $record.IPAddress + " - Online"
    } else {
        write-host $record.Ownername "is Offline" -ForegroundColor Red
        $result += $record.Ownername + " - " + $record.IPAddress + " - Offline"
    }
 }
   
$result | Out-File -FilePath C:\Temp\DNS-Stale.csv
