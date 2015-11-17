$Computer = "bmobccsccmc01"
$Namespace = "root\sms\site_B01"
$Class = "SMS_Advertisement"

$SCCMAdvertisements = Get-WmiObject -class $Class -computer $Computer -namespace $Namespace
$SCCMAdvertisements.psbase.Get()
foreach ($Advert in $SCCMAdvertisements)
	{
	$Advert.psbase.Get()
	Write-Host
	Write-Host $Advert.AdvertisementName
    Write-Host $Advert.ExpirationTime.Substring(0,4)'/'$Advert.ExpirationTime.Substring(4,2)'/'$Advert.ExpirationTime.Substring(6,2) $Advert.ExpirationTime.Substring(8,2)':'$Advert.ExpirationTime.Substring(10,2)':'$Advert.ExpirationTime.Substring(12,2)
	$AdvProperties = $Advert.AssignedSchedule
	foreach ($Adv in $AdvProperties)
		{
		Write-Host $Adv.__CLASS
        Write-host $Adv.StartTime.Substring(0,4)'/'$Adv.StartTime.Substring(4,2)'/'$Adv.StartTime.Substring(6,2) $Adv.StartTime.Substring(8,2)':'$Adv.StartTime.Substring(10,2)':'$Adv.StartTime.Substring(12,2)
		}
	}
