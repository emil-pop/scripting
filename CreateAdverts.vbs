intAnswer = _
    Msgbox("Do you want to create patching advertisements?", _
        vbYesNo, "Create Advertisements")

If intAnswer = vbNo Then
    wscript.quit
End If

'TimeFlags
Const ENABLE_AVAILABLE = &H00000004
Const ENABLE_MANDATORY = &H00000010

'AdvertFlags
Const WAKE_ON_LAN_ENABLED = &H00400000       
Const REBOOT_OUTSIDE_OF_MAINTENANCE_WINDOWS = &H00200000  
Const OVERRIDE_MAINTENANCE_WINDOWS = &H00100000     
Const ENABLE_TS_FROM_CD_AND_PXE = &H00040000     
Const NO_DISPLAY = &H02000000         
Const DONOT_FALLBACK = &H00020000

'RemoteClientFlags
Const RERUN_ALWAYS = &H00000800         
Const RERUN_IF_FAILED = &H00002000         
Const RUN_FROM_LOCAL_DISPPOINT = &H00000008      
Const DOWNLOAD_FROM_LOCAL_DISPPOINT = &H00000010
Const DOWNLOAD_FROM_REMOTE_DISPPOINT = &H00000040    
Const DONT_RUN_NO_LOCAL_DISPPOINT = &H00000020

Dim Coll(42), Name(42), Program(42)

TitleMonth = "Nov"
PackageID = "B01000E6" 'Microsoft Security Patches - November 2015

Coll(1) = "B01002A3"	'IOC SC 21Fr
Name(1) = "SC-Fr-2100"
Program(1) = "__Install Patches - REBOOT"

Coll(2) = "B01002A4"	'IOC SC 22Fr
Name(2) = "SC-Fr-2200"
Program(2) = "__Install Patches - REBOOT"

Coll(3) = "B01002A5"	'IOC SC 03Sa
Name(3) = "SC-Fr-0300"
Program(3) = "__Install Patches - REBOOT"

Coll(4) = "B01002B8"	'IOC DIR Man
Name(4) = "DIR-8-Manual"
Program(4) = "Install Patches - NO REBOOT"

Coll(5) = "B01002B6"	'IOC DIR W1 21Sa
Name(5) = "DIR-0-BCC"
Program(5) = "__Install Patches - REBOOT"

Coll(6) = "B01002B7"	'IOC DIR W2 21Sa
Name(6) = "DIR-7-SCC"
Program(6) = "__Install Patches - REBOOT"

Coll(7) = "B01002B5"	'IOC FP Man
Name(7) = "FP-Manual"
Program(7) = "__Install Patches - REBOOT"

Coll(8) = "B01002AC"	'IOC FP W1 21Sa
Name(8) = "FP-W1-Sat-2100"
Program(8) = "__Install Patches - REBOOT"

Coll(9) = "B01002AD"	'IOC FP W2 21Sa
Name(9) = "FP-W2-Sat-2100"
Program(9) = "__Install Patches - REBOOT"

Coll(10) = "B01002A9"	'IOC Proxy 21Sa
Name(10) = "Proxy-Sat-2100"
Program(10) = "__Install Patches - REBOOT"

Coll(11) = "B01002AA"	'IOC Proxy 22Sa
Name(11) = "Proxy-Sat-2200"
Program(11) = "__Install Patches - REBOOT"

Coll(12) = "B01002AB"	'IOC Proxy 23Sa
Name(12) = "Proxy-Sat-2300"
Program(12) = "__Install Patches - REBOOT"

Coll(13) = "B0100305"	'IOC-VMH-W1_We
Name(13) = "1-VMH-W1_We"
Program(13) = "__Install Patches - REBOOT"

Coll(14) = "B0100308"	'IOC-VMH-W1_Th
Name(14) = "2-VMH-W1_Th"
Program(14) = "__Install Patches - REBOOT"

Coll(15) = "B010030C"	'IOC-VMH-W2_Mo
Name(15) = "3-VMH-W2_Mo"
Program(15) = "__Install Patches - REBOOT"

Coll(16) = "B0100306"	'IOC-VMH-W2_Tu
Name(16) = "4-VMH-W2_Tu"
Program(16) = "__Install Patches - REBOOT"

Coll(17) = "B0100307"	'IOC-VMH-W2_We
Name(17) = "5-VMH-W2_We"
Program(17) = "__Install Patches - REBOOT"

Coll(18) = "B010030B"	'IOC-VMH-W2_Th
Name(18) = "6-VMH-W2_Th"
Program(18) = "__Install Patches - REBOOT"

Coll(19) = "B0100309"	'IOC-VMH-W3_Mo
Name(19) = "7-VMH-W3_Mo"
Program(19) = "__Install Patches - REBOOT"

Coll(20) = "B010030A"	'IOC-VMH-W3_Tu
Name(20) = "8-VMH-W3_Tu"
Program(20) = "__Install Patches - REBOOT"

Coll(21) = "B01002C2"	'IOC MSG Man
Name(21) = "MSG-Manual"
Program(21) = "__Install Patches - REBOOT"

Coll(22) = "B01002B9"	'IOC MSG 21Sa
Name(22) = "MSG-Sat-2100"
Program(22) = "__Install Patches - REBOOT"

Coll(23) = "B01002BC"	'IOC MSG 23Sa
Name(23) = "MSG-Sat-2300"
Program(23) = "__Install Patches - REBOOT"

Coll(24) = "B01002BA"	'IOC MSG LN 21Sa
Name(24) = "MSG-Notes-Sat-2100"
Program(24) = "Install Patches on Domino Servers"

Coll(25) = "B01002BB"	'IOC MSG LN 23Sa
Name(25) = "MSG-Notes-Sat-2300"
Program(25) = "Install Patches on Domino Servers"

Coll(26) = "B01002E8"	'IOC MSG LN v 8.5.3 21Sa
Name(26) = "MSG-Notes853-Sat-2100"
Program(26) = "Install Patches on Domino v8.5.3"

Coll(27) = "B01002DB"	'IOC MSG Autonomy 21Fr
Name(27) = "MSG-Autonomy-Fr-2100"
Program(27) = "Install Patches on Autonomy Servers"

Coll(28) = "B01002DC"	'IOC MSG Autonomy 23Fr
Name(28) = "MSG-Autonomy-Fr-2300"
Program(28) = "Install Patches on Autonomy Servers"

Coll(29) = "B01002DD"	'IOC MSG Autonomy 21Sa
Name(29) = "MSG-Autonomy-Sat-2100"
Program(29) = "Install Patches on Autonomy Servers"

Coll(30) = "B01002DE"	'IOC MSG Autonomy 23Sa
Name(30) = "MSG-Autonomy-Sat-2300"
Program(30) = "Install Patches on Autonomy Servers"

Coll(31) = "B01002EA"	'IOC MSG Lync 21Fr
Name(31) = "MSG-Lync-Fr-2100"
Program(31) = "Install Patches on LYNC Servers"

Coll(32) = "B01002EB"	'IOC MSG Lync 23Fr
Name(32) = "MSG-Lync-Fr-2300"
Program(32) = "Install Patches on LYNC Servers"

Coll(33) = "B01002EC"	'IOC MSG Lync 01Sa
Name(33) = "MSG-Lync-Sat-0100"
Program(33) = "Install Patches on LYNC Servers"

Coll(34) = "B01002ED"	'IOC MSG Lync 21Sa
Name(34) = "MSG-Lync-Sat-2100"
Program(34) = "Install Patches on LYNC Servers"

Coll(35) = "B01002EE"	'IOC MSG Lync 23Sa
Name(35) = "MSG-Lync-Sat-2300"
Program(35) = "Install Patches on LYNC Servers"

Coll(36) = "B01002EF"	'IOC MSG Lync 01Su
Name(36) = "MSG-Lync-Sun-0100"
Program(36) = "Install Patches on LYNC Servers"

Coll(37) = "B010033B"	'IOC DIR 1 21Su
Name(37) = "DIR-1-Sun-2100-UTC"
Program(37) = "__Install Patches - REBOOT"

Coll(38) = "B0100340"	'IOC DIR 2 21Mo
Name(38) = "DIR-2-Mon-2100-UTC"
Program(38) = "__Install Patches - REBOOT"

Coll(39) = "B0100341"	'IOC DIR 3 21Tu
Name(39) = "DIR-3-Tue-2100-UTC"
Program(39) = "__Install Patches - REBOOT"

Coll(40) = "B0100342"	'IOC DIR 4 21We
Name(40) = "DIR-4-Wed-2100-UTC"
Program(40) = "__Install Patches - REBOOT"

Coll(41) = "B0100343"	'IOC DIR 5 21Th
Name(41) = "DIR-5-Thu-2100-UTC"
Program(41) = "__Install Patches - REBOOT"

Coll(42) = "B0100344"	'IOC DIR 6 21Fr
Name(42) = "DIR-6-Fri-2100-UTC"
Program(42) = "__Install Patches - REBOOT"

'Connect to provider namespace for local computer.
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices= objSWbemLocator.ConnectServer("bmobccsccmc01", "root\sms")
Set ProviderLoc = objSWbemServices.InstancesOf("SMS_ProviderLocation")

For Each Location In ProviderLoc
        If Location.ProviderForLocalSite = True Then     
            Set objSWbemServices = objSWbemLocator.ConnectServer _
                 (Location.Machine, "root\sms\site_" + Location.SiteCode)
            Set Site = objSWbemServices.Get("SMS_Site='" & Location.SiteCode  & "'")
         End If
Next

sFormattedMonth = Month(Now)
If Len(sFormattedMonth) = 1 Then
	sFormattedMonth = "0" & sFormattedMonth
End If
sFormattedDay = Day(Now)
If Len(sFormattedDay) = 1 Then
	sFormattedDay = "0" & sFormattedDay
End If
strDateTime = Year(Now) & sFormattedMonth & sFormattedDay & Left(formatdatetime(Now, 4),2) & Right(formatdatetime(Now, 4),2) & "00.000000+***"

REM ExpireDate = DateAdd("h",12,Now)
REM sFormattedMonth = Month(ExpireDate)
REM If Len(sFormattedMonth) = 1 Then
	REM sFormattedMonth = "0" & sFormattedMonth
REM End If
REM sFormattedDay = Day(ExpireDate)
REM If Len(sFormattedDay) = 1 Then
	REM sFormattedDay = "0" & sFormattedDay
REM End If
REM strExpireDate = Year(Now) & sFormattedMonth & sFormattedDay & Left(formatdatetime(ExpireDate, 4),2) & Right(formatdatetime(ExpireDate, 4),2) & "00.000000+***"

'Create new advertisements, and configure some properties

For x = 1 to uBound(Coll)
    AdName="IOC-" + TitleMonth + "Patches-" + name(x)
    Set newAdvertisement = objSWbemServices.Get("SMS_Advertisement").SpawnInstance_()
    newAdvertisement.AdvertisementName = AdName
    newAdvertisement.CollectionID = Coll(x)
    newAdvertisement.PackageID = PackageID
    newAdvertisement.ProgramName = Program(x)
    newAdvertisement.PresentTime = strDateTime
    'newAdvertisement.ExpirationTime = strExpireDate
    'newAdvertisement.ExpirationTimeEnabled = True
    newAdvertisement.AssignedScheduleEnabled = True
    newAdvertisement.IncludeSubCollection = False
    'newAdvertisement.TimeFlags = ENABLE_AVAILABLE OR ENABLE_MANDATORY
    newAdvertisement.RemoteClientFlags = RERUN_ALWAYS OR DONT_RUN_NO_LOCAL_DISPPOINT OR DOWNLOAD_FROM_LOCAL_DISPPOINT
    'newAdvertisement.AdvertFlags = NO_DISPLAY or DONOT_FALLBACK
    newAdvertisement.Priority = 1
    'newAdvertisement.AssignedSchedule = Array
  
    'Save advertisement
    Handle = newAdvertisement.Put_
Next
