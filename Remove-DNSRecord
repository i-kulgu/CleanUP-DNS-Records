Function Remove-DNSRecord
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true,Position=0)]
        # RecordName is the name of host or CNAMe to Delete

        [string]$RecordName,

        # DNSServer or domain name
        [Parameter(Mandatory=$true,Position=1)]
        [string]$DNSServer
    )

    Begin   {$NodeARecord=$null}
    Process  {
    if ($pscmdlet.ShouldProcess($RecordName)){ $bTest = $False}Else{$bTest = $true}

    Write-Host "Getting Zones for $DNSServer" -ForegroundColor "Green"
    $Zones = @(Get-DnsServerZone -ComputerName $DNSServer)
    $NotLookup = $Zones | where {$_.ZoneName -notmatch 'arpa'}

    ForEach ($ZoneName in ($NotLookup).ZoneName) {
        #$RecordName = "automat004"
	    Write-Host "Searching $ZoneName" -ForegroundColor "Green"
	    $Zone | Foreach{
            $NodeARecord = Get-DnsServerResourceRecord -ZoneName $ZoneName -ComputerName $DNSServer -Name $RecordName -ErrorAction SilentlyContinue
            if($NodeARecord){
                Remove-DnsServerResourceRecord -ZoneName $ZoneName -ComputerName $DNSServer -InputObject $NodeARecord -Force -whatif:$bTest
                Write-Host ("A record deleted: "+$NodeARecord.HostName)
                Break
            }
        }
    }

    if ($NodeARecord){
        $IPAddress = $NodeARecord.RecordData.IPv4Address.IPAddressToString
        $IPAddressArray = $IPAddress.Split(".")
        $DNSName = $IPAddressArray[3]
        $IPAddressFormatted = ($IPAddressArray[3]+"."+$IPAddressArray[2])
        $ZonePrefix = ($IPAddressArray[2]+"."+$IPAddressArray[1]+"."+$IPAddressArray[0])
        $ReverseZoneName = "$ZonePrefix`.in-addr.arpa"

        $NodePTRRecord = Get-DnsServerResourceRecord -ZoneName $ReverseZoneName -ComputerName $DNSServer -Name $DNSName -RRType Ptr -ErrorAction SilentlyContinue
        if($NodePTRRecord -eq $null){
            Write-Host "No PTR record found"
        } else {
            Remove-DnsServerResourceRecord -ZoneName $ReverseZoneName -ComputerName $DNSServer -InputObject $NodePTRRecord -Force -WhatIf:$bTest
            Write-Host ("PTR Record Deleted: "+$IPAddressFormatted)
        }

    }ELSE{
        Write-warning "No record for $RecordName found at $DNSServer"
  }
  }#End Process


  End{  Write "Done"  }

}

$records = Import-XLSX -Path D:\DNS-Stale-Records.xlsx
$Hosts = $records.Kolom1
$HostArray = new-object System.Collections.ArrayList

foreach ($hostlist in $Hosts){
    $HostArray += $hostlist
}

$HostArray | foreach {Remove-DNSRecord -RecordName $_ -DNSServer "DC01" }
